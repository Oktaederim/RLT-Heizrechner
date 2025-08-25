# app.py – Heizenergie-Rechner Lüftungsanlagen (TRY → Monats/Jahreswerte)
# Deutsch, stabiler Import (24:00), klare UI (Formulare), nachvollziehbare Rechnung, Excel/PDF
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st

# Optional: PDF
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

# ---------- Helpers ----------
def parse_dt_iso(s: str) -> Optional[datetime]:
    s = str(s).strip().replace(" ", "T")
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def minutes(hhmm: str) -> int:
    h, m = hhmm.split(":")
    return int(h)*60 + int(m)

def overlap_mins(a0, a1, b0, b1) -> int:
    return max(0, min(a1, b1) - max(a0, b0))

def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))

# ---------- Datenklassen ----------
@dataclass
class Zeitfenster:
    start: str     # "HH:MM"
    ende: str      # "HH:MM"
    aktiv: bool
    T_soll_C: float
    V_m3h: Optional[float] = None  # None = Standardwert aus NORMAL/ABSENK-Defaults

@dataclass
class Tagesplan:
    tag: int                 # 0=Mo..6=So
    normal: List[Zeitfenster]
    absenk: List[Zeitfenster]

@dataclass
class Defaults:
    T_normal_C: float = 20.0
    T_absenk_C: float = 17.0
    V_normal_m3h: float = 5000.0
    V_absenk_m3h: Optional[float] = 2000.0  # None ⇒ wie normal

@dataclass
class Anlage:
    id: str
    name: str
    V_nominal_m3h: float
    anzahl: int
    wrg: bool
    eta_t: float               # 0..1
    fan_kW: Optional[float]    # gesamt bei V_nominal
    SFP_kW_per_m3s: Optional[float]
    wochenplan: List[Tagesplan]

WOCHENTAGE = ["Mo","Di","Mi","Do","Fr","Sa","So"]

def leerer_wochenplan(defs: Defaults) -> List[Tagesplan]:
    plan: List[Tagesplan] = []
    for d in range(7):
        normal = [Zeitfenster("06:30","17:00", True, defs.T_normal_C, None)] if d < 5 else []
        absenk = [Zeitfenster("17:00","06:30", True, defs.T_absenk_C, defs.V_absenk_m3h)] if d < 5 else []
        plan.append(Tagesplan(d, normal, absenk))
    return plan

def normiere_wochenplan(plan: List[Tagesplan]) -> List[List[Tuple[int,int,str,float,Optional[float]]]]:
    """je Tag: (startMin, endMin, 'Normal'|'Absenk', T_soll, V_override) – über Mitternacht gesplittet"""
    out = [[] for _ in range(7)]
    def add(t: int, f: Zeitfenster, modus: str):
        if not f.aktiv: return
        s, e = minutes(f.start), minutes(f.ende)
        if s == e: return
        entry = (None if f.V_m3h is None else float(f.V_m3h))
        if e > s:
            out[t].append((s, e, modus, float(f.T_soll_C), entry))
        else:
            out[t].append((s, 1440, modus, float(f.T_soll_C), entry))
            out[(t+1)%7].append((0, e, modus, float(f.T_soll_C), entry))
    for d in plan:
        for f in d.normal: add(d.tag, f, "Normal")
        for f in d.absenk: add(d.tag, f, "Absenk")
    for i in range(7): out[i].sort(key=lambda x: x[0])
    return out

# ---------- TRY-CSV robust laden ----------
def parse_try_csv(raw: pd.DataFrame) -> Tuple[pd.DataFrame, list, list]:
    """gibt df[datetime,T_out_C], jahre[], problems[] zurück; behandelt 24:00 korrekt & reindiziert stündlich"""
    def _find(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
        low = {c.lower().strip(): c for c in df.columns}
        for a in candidates:
            if a in low: return low[a]
        for c in df.columns:
            cl = c.lower().strip()
            if any(a in cl for a in candidates): return c
        return None

    dt_alias = ["datetime","date_time","date/time","date","timestamp","zeit","zeitstempel","datestamp"]
    t_alias  = ["t_out_c","t_out","tout","temp_out","temperature_out","aussen","außen","ta","t2m"]

    dt_col = _find(raw, dt_alias) or st.selectbox("Datums-/Zeitspalte wählen", raw.columns)
    t_col  = _find(raw, t_alias)  or st.selectbox("Außentemperatur-Spalte wählen", raw.columns)

    df = raw[[dt_col, t_col]].copy()
    # Temperatur robust zu float
    if df[t_col].dtype == object:
        df[t_col] = (df[t_col].astype(str)
                                 .str.replace(",", ".", regex=False)
                                 .str.replace("°C", "", regex=False)
                                 .str.strip())
    df[t_col] = pd.to_numeric(df[t_col], errors="coerce")
    df = df.rename(columns={t_col: "T_out_C"})

    # 24:00 sauber behandeln → 00:00 + 1 Tag
    def _fix_24h(s: str):
        s = str(s).strip().replace(" ", "T")
        has_24 = ("T24:" in s) or s.endswith("24:00")
        s2 = s.replace("T24:", "T00:").replace(" 24:", " 00:")
        dt = pd.to_datetime(s2, errors="coerce")
        if has_24 and pd.notna(dt):
            dt = dt + pd.Timedelta(days=1)
        return dt

    df["datetime"] = raw[dt_col].astype(str).apply(_fix_24h)
    df = df[["datetime","T_out_C"]].dropna().sort_values("datetime")

    # Doppelte Timestamps zusammenfassen (letzter Wert gewinnt), aber erst NACH 24:00-Korrektur
    before = len(df)
    df = df.groupby("datetime", as_index=False)["T_out_C"].last()
    duplicates_removed = before - len(df)

    # Auf durchgängiges Stundenraster reindizieren
    probs = []
    if not df.empty:
        start = df["datetime"].min().replace(minute=0, second=0, microsecond=0)
        end   = df["datetime"].max().replace(minute=0, second=0, microsecond=0)
        full = pd.date_range(start, end, freq="H")
        s = df.set_index("datetime")["T_out_C"].reindex(full)
        missing = int(s.isna().sum())
        if missing > 0:
            probs.append(f"{missing} fehlende Stunde(n) – Quelle prüfen.")
        df = s.reset_index().rename(columns={"index":"datetime", 0:"T_out_C"})

    years = sorted(df["datetime"].dt.year.unique().tolist()) if not df.empty else []
    # Nur informieren, wenn wirklich Doppelte nach der Korrektur existierten
    if duplicates_removed > 0:
        probs.append(f"{duplicates_removed} doppelte Zeitstempel zusammengefasst.")

    return df, years, probs

# ---------- Berechnung ----------
def berechne(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults):
    """liefert Monate, Jahr, Protokoll (Stundenausschnitt)"""
    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = normiere_wochenplan(anlage.wochenplan)

    rows = []
    prot = []

    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])
        d = dt.weekday()
        m0, m1 = dt.hour*60 + dt.minute, dt.hour*60 + dt.minute + 60

        Q_h = 0.0
        E_h = 0.0
        fan_h = 0.0

        for (s, e, modus, T_soll, V_ovr) in plan[d]:
            ol = overlap_mins(m0, m1, s, e)
            if ol <= 0: 
                continue
            anteil = ol/60.0

            # Volumenstrom
            if V_ovr is not None:
                V = max(0.0, float(V_ovr))
            else:
                V = V_norm_total if modus=="Normal" else (V_norm_total if defs.V_absenk_m3h is None else float(defs.V_absenk_m3h))

            # Heizenergie (kWh) – korrekt mit 0.00034
            dT = max(0.0, float(T_soll) - Tout)
            dT_eff = (1.0 - clamp(float(anlage.eta_t), 0.0, 1.0))*dT if anlage.wrg else dT
            Q_kWh = 0.00034 * V * dT_eff * anteil

            # Ventilator
            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = float(anlage.SFP_kW_per_m3s) * (V/3600.0)
                elif anlage.fan_kW is not None:
                    P_kW = float(anlage.fan_kW) * clamp(V/max(1.0, V_norm_total), 0.0, 1.0)
            E_kWh = P_kW * anteil

            Q_h += Q_kWh
            E_h += E_kWh
            fan_h += (anteil if V > 0 else 0.0)

            # Protokoll (nur wenige Zeilen anzeigen wir später)
            prot.append({
                "Zeit": dt, "Modus": modus, "Anteil_h": round(anteil,3),
                "T_out": round(Tout,1), "T_soll": round(T_soll,1),
                "ΔT_eff": round(dT_eff,2), "V_m3h": round(V,0),
                "Q_kWh": round(Q_kWh,4), "P_fan_kW": round(P_kW,3), "E_kWh": round(E_kWh,4)
            })

        rows.append({
            "datetime": dt, "year": dt.year, "month": dt.month,
            "kWh_th": Q_h, "kWh_el": E_h, "Betriebsstunden_Vent": fan_h
        })

    dfh = pd.DataFrame.from_records(rows)
    mon = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    jahr = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    return mon, jahr, pd.DataFrame(prot)

# ---------- Exporte ----------
def xlsx_export(mon: pd.DataFrame, jahr: pd.DataFrame, prot: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        m = mon.copy(); j = jahr.copy(); p = prot.copy()
        for col in ("kWh_th","kWh_el"):
            if col in m: m[col] = m[col].round(0)
            if col in j: j[col] = j[col].round(0)
        if "Betriebsstunden_Vent" in m: m["Betriebsstunden_Vent"] = m["Betriebsstunden_Vent"].round(1)
        if "Betriebsstunden_Vent" in j: j["Betriebsstunden_Vent"] = j["Betriebsstunden_Vent"].round(1)
        if not p.empty:
            p["Q_kWh"] = p["Q_kWh"].round(4); p["E_kWh"] = p["E_kWh"].round(4)

        m.to_excel(w, index=False, sheet_name="Monate")
        j.to_excel(w, index=False, sheet_name="Jahr")
        p.to_excel(w, index=False, sheet_name="Protokoll")
    return out.getvalue()

def pdf_export(info: str, defs: Defaults, anl: Anlage, mon: pd.DataFrame, jahr: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab nicht installiert.")
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]; H1.fontSize=14
    H2 = styles["Heading2"]; H2.fontSize=12
    N  = styles["BodyText"]; N.leading=14

    story = []
    story += [Paragraph("ISO 50001 – Heizenergie Lüftungsanlagen (v1, vereinfachtes Verfahren)", H1), Spacer(1,6)]
    story += [Paragraph(f"Erstellt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1,8)]
    story += [Paragraph("Quelle / TRY", H2), Paragraph(info or "TRY-CSV (stündlich).", N), Spacer(1,6)]
    story += [Paragraph("Annahmen & Parameter", H2),
              Paragraph(f"T_normal: {defs.T_normal_C} °C; T_absenk: {defs.T_absenk_C} °C; "
                        f"V_normal: {defs.V_normal_m3h} m³/h; V_absenk: {defs.V_absenk_m3h if defs.V_absenk_m3h is not None else 'wie normal'} m³/h. "
                        f"WRG: {'ja' if anl.wrg else 'nein'} (η_t={anl.eta_t}). "
                        f"Ventilator: {'SFP '+str(anl.SFP_kW_per_m3s)+' kW/(m³/s)' if anl.SFP_kW_per_m3s is not None else 'fan_kW '+str(anl.fan_kW)+' kW'}.", N),
              Spacer(1,6)]
    story += [Paragraph("Methodik", H2),
              Paragraph("Stündlich aktives Zeitfenster (Normal/Absenk). Heizfall: ΔT = max(0, T_soll − T_out); "
                        "mit WRG: ΔT_eff = (1 − η_t)·ΔT. Wärme: Q_kWh = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h. "
                        "Ventilator: P_fan = SFP·(V/3600) oder fan_kW·(V/V_nominal), Energie E_kWh = P_fan·Anteil_h. "
                        "Aggregation zu Monaten/Jahr.", N), Spacer(1,6)]

    if not jahr.empty:
        jj = jahr.copy()
        jj[["kWh_th","kWh_el"]] = jj[["kWh_th","kWh_el"]].round(0)
        jj["Betriebsstunden_Vent"] = jj["Betriebsstunden_Vent"].round(1)
        t1 = [["Jahr","kWh_th","kWh_el","Betriebsstunden Vent."], *jj.values.tolist()]
        T1 = Table(t1, hAlign="LEFT")
        T1.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("ALIGN",(1,1),(-1,-1),"RIGHT")]))
        story += [Paragraph("Ergebnisse – Jahreswerte", H2), T1, Spacer(1,6)]

    if not mon.empty:
        mm = mon.copy()
        mm[["kWh_th","kWh_el"]] = mm[["kWh_th","kWh_el"]].round(0)
        mm["Betriebsstunden_Vent"] = mm["Betriebsstunden_Vent"].round(1)
        t2 = [["Jahr","Monat","kWh_th","kWh_el","Betriebsstunden Vent."], *mm.astype({"year":int,"month":int}).values.tolist()]
        T2 = Table(t2, hAlign="LEFT", repeatRows=1)
        T2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("ALIGN",(2,1),(-1,-1),"RIGHT")]))
        story += [Paragraph("Ergebnisse – Monate", H2), T2]

    doc.build(story); out.seek(0); return out.read()

# ---------- UI ----------
st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001)", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session
for k, v in [("try_df",None), ("try_info",""), ("mon_df",None), ("jahr_df",None), ("prot_df",None),
             ("defs", None), ("anlage", None), ("wochenplan", None),
             ("auto_calc", False), ("trigger", False)]:
    if k not in st.session_state: st.session_state[k] = v

st.markdown("**Ablauf:** 1) TRY laden → 2) Anlage → 3) Zeiten → 4) Berechnen. "
            "Ergebnisse werden **sichtbar gezeigt**, **Downloads darunter**.")

# 1) TRY laden (Form)
with st.form("form_try"):
    st.subheader("1) TRY‑CSV laden")
    f = st.file_uploader("TRY‑CSV (stündlich: Datum/Zeit + Außentemperatur)", type=["csv"])
    auto = st.checkbox("Automatisch nach Upload berechnen", value=st.session_state["auto_calc"])
    ok1 = st.form_submit_button("TRY übernehmen")
    if ok1:
        if f is None:
            st.error("Bitte CSV auswählen.")
        else:
            raw = pd.read_csv(f)
            df, years, probs = parse_try_csv(raw)
            if df.empty:
                st.error("Kein gültiger Datensatz.")
            else:
                tmin, tmax = df["T_out_C"].min(), df["T_out_C"].max()
                st.session_state["try_df"] = df
                st.session_state["try_info"] = f"Datensätze: {len(df)} | Jahre: {', '.join(map(str, years))} | T_out: {round(tmin,1)}…{round(tmax,1)} °C"
                st.session_state["auto_calc"] = auto
                st.success("TRY übernommen.")
                st.write(st.session_state["try_info"])
                for p in probs: st.info(p)
                if auto: st.session_state["trigger"] = True

# 2) Anlage (Form)
with st.form("form_anlage"):
    st.subheader("2) Anlagendaten")
    c1 = st.columns([1.1,1,1,1])
    a_id   = c1[0].text_input("Anlagen‑ID", value="A01")
    a_name = c1[1].text_input("Bezeichnung", value="Zuluft")
    V_nom  = c1[2].number_input("V_nominal (Einzelanlage) [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    n_eq   = c1[3].number_input("Anzahl gleicher Anlagen", value=1, min_value=1, max_value=100, step=1)

    c2 = st.columns(3)
    wrg   = c2[0].checkbox("WRG vorhanden", value=True)
    eta_t = c2[1].number_input("η_t WRG (0–1)", value=0.7, min_value=0.0, max_value=1.0, step=0.05)
    vm    = c2[2].selectbox("Ventilator-Modell", ("SFP [kW/(m³/s)]","fan_kW gesamt"), index=1)

    fan_kW = None; SFP = None
    if vm.startswith("SFP"):
        SFP = st.number_input("SFP [kW/(m³/s)]", value=1.8, min_value=0.0, max_value=10.0, step=0.1)
    else:
        fan_kW = st.number_input("fan_kW (bei V_nominal gesamt) [kW]", value=5.0, min_value=0.0, max_value=1000.0, step=0.1)

    ok2 = st.form_submit_button("Anladendaten übernehmen")
    if ok2:
        st.session_state["anlage"] = Anlage(a_id, a_name, float(V_nom), int(n_eq), bool(wrg), float(eta_t),
                                            float(fan_kW) if fan_kW is not None else None,
                                            float(SFP) if SFP is not None else None,
                                            st.session_state["wochenplan"] or [])
        st.success("Anlagendaten übernommen.")

# 3) Standard‑Werte & Zeiten (Form)
with st.form("form_zeiten"):
    st.subheader("3) Standardwerte & Betriebs-/Absenkzeiten")
    c = st.columns(4)
    Tn = c[0].number_input("Soll‑Zuluft NORMAL [°C]", value=20.0, step=0.5)
    Ta = c[1].number_input("Soll‑Zuluft ABSENK [°C]", value=17.0, step=0.5)
    Vn = c[2].number_input("Volumenstrom NORMAL [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    Va = c[3].number_input("Volumenstrom ABSENK [m³/h] (0 = wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0)

    defs = Defaults(float(Tn), float(Ta), float(Vn), (None if Va==0.0 else float(Va)))
    if st.session_state["wochenplan"] is None:
        st.session_state["wochenplan"] = leerer_wochenplan(defs)

    wp: List[Tagesplan] = st.session_state["wochenplan"]

    st.caption("Pro Tag beliebig viele Fenster. Über‑Mitternacht (z. B. 17:00–06:30) möglich.")
    for d in range(7):
        st.markdown(f"**{WOCHENTAGE[d]}**")
        day = wp[d]

        st.markdown("_Normalbetrieb_")
        if st.button(f"+ Normal‑Fenster {WOCHENTAGE[d]}", key=f"add_n_{d}"):
            day.normal.append(Zeitfenster("08:00","16:00", True, defs.T_normal_C, None))
        del_n = None
        for i, f in enumerate(day.normal):
            cols = st.columns([1,1,0.8,1,1,0.6])
            f.start = cols[0].text_input("Start", key=f"n_s_{d}_{i}", value=f.start)
            f.ende  = cols[1].text_input("Ende",  key=f"n_e_{d}_{i}", value=f.ende)
            f.aktiv = cols[2].checkbox("aktiv", key=f"n_a_{d}_{i}", value=f.aktiv)
            f.T_soll_C = cols[3].number_input("T_soll [°C]", key=f"n_t_{d}_{i}", value=float(f.T_soll_C), step=0.5)
            vval = 0.0 if f.V_m3h is None else float(f.V_m3h)
            newv = cols[4].number_input("V [m³/h] (0 = Standard)", key=f"n_v_{d}_{i}", value=vval, min_value=0.0, max_value=500000.0, step=100.0)
            f.V_m3h = None if newv==0.0 else float(newv)
            if cols[5].button("–", key=f"n_del_{d}_{i}"): del_n = i
        if del_n is not None: day.normal.pop(del_n)

        st.markdown("_Absenkbetrieb_")
        if st.button(f"+ Absenk‑Fenster {WOCHENTAGE[d]}", key=f"add_a_{d}"):
            day.absenk.append(Zeitfenster("17:00","06:30", True, defs.T_absenk_C, defs.V_absenk_m3h))
        del_a = None
        for i, f in enumerate(day.absenk):
            cols = st.columns([1,1,0.8,1,1,0.6])
            f.start = cols[0].text_input("Start", key=f"a_s_{d}_{i}", value=f.start)
            f.ende  = cols[1].text_input("Ende",  key=f"a_e_{d}_{i}", value=f.ende)
            f.aktiv = cols[2].checkbox("aktiv", key=f"a_a_{d}_{i}", value=f.aktiv)
            f.T_soll_C = cols[3].number_input("T_soll [°C]", key=f"a_t_{d}_{i}", value=float(f.T_soll_C), step=0.5)
            vval = 0.0 if f.V_m3h is None else float(f.V_m3h)
            newv = cols[4].number_input("V [m³/h] (0 = Standard)", key=f"a_v_{d}_{i}", value=vval, min_value=0.0, max_value=500000.0, step=100.0)
            f.V_m3h = None if newv==0.0 else float(newv)
            if cols[5].button("–", key=f"a_del_{d}_{i}"): del_a = i
        if del_a is not None: day.absenk.pop(del_a)

    ok3 = st.form_submit_button("Zeiten übernehmen")
    if ok3:
        st.session_state["defs"] = defs
        st.session_state["wochenplan"] = wp
        st.success("Standardwerte & Zeiten übernommen.")

# 4) Berechnen (Button außerhalb der Formulare)
def _rechnen():
    df = st.session_state["try_df"]
    defs = st.session_state["defs"]
    wp = st.session_state["wochenplan"]
    anl = st.session_state["anlage"]
    if df is None:
        st.error("Zuerst TRY laden."); return
    if anl is None:
        st.error("Zuerst Anlagendaten übernehmen."); return
    if defs is None or wp is None:
        st.error("Zuerst Standardwerte & Zeiten übernehmen."); return
    anl = Anlage(anl.id, anl.name, anl.V_nominal_m3h, anl.anzahl, anl.wrg, anl.eta_t, anl.fan_kW, anl.SFP_kW_per_m3s, wp)
    mon, jahr, prot = berechne(df, anl, defs)
    st.session_state["mon_df"], st.session_state["jahr_df"], st.session_state["prot_df"] = mon, jahr, prot
    st.success("Berechnung abgeschlossen.")

cols_run = st.columns([1,3])
if cols_run[0].button("Berechnen", type="primary"): _rechnen()
if st.session_state["trigger"]:
    _rechnen(); st.session_state["trigger"] = False

# Ergebnisse sichtbar + Downloads darunter
mon = st.session_state["mon_df"]; jahr = st.session_state["jahr_df"]; prot = st.session_state["prot_df"]
if mon is not None and jahr is not None:
    m_show = mon.copy(); j_show = jahr.copy()
    for c in ("kWh_th","kWh_el"):
        if c in m_show: m_show[c] = m_show[c].round(0)
        if c in j_show: j_show[c] = j_show[c].round(0)
    if "Betriebsstunden_Vent" in m_show: m_show["Betriebsstunden_Vent"] = m_show["Betriebsstunden_Vent"].round(1)
    if "Betriebsstunden_Vent" in j_show: j_show["Betriebsstunden_Vent"] = j_show["Betriebsstunden_Vent"].round(1)

    st.subheader("Ergebnisse – Jahreswerte")
    st.dataframe(j_show, use_container_width=True)

    st.subheader("Ergebnisse – Monate")
    st.dataframe(m_show, use_container_width=True)

    st.subheader("Rechen‑Protokoll (Ausschnitt)")
    if prot is not None and not prot.empty:
        st.dataframe(prot.head(200), use_container_width=True)
    else:
        st.info("Kein Protokoll verfügbar.")

    st.subheader("Downloads")
    c1,c2 = st.columns(2)
    c1.download_button("Excel (Monate + Jahr + Protokoll)", xlsx_export(m_show, j_show, prot if prot is not None else pd.DataFrame()),
                       file_name="Heizenergie_Auswertung.xlsx")
    if REPORTLAB_OK:
        # Dummy für PDF-Kopf
        anl = st.session_state["anlage"]
        defs = st.session_state["defs"]
        if anl and defs:
            c2.download_button("PDF (ISO 50001 Kurzbericht)",
                               pdf_export(st.session_state["try_info"], defs, anl, m_show, j_show),
                               file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
    else:
        st.info("PDF‑Export nicht verfügbar (ReportLab nicht installiert).")
