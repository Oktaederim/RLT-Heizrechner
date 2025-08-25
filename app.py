# app.py
# Heizenergie-Rechner Lüftungsanlagen (TRY → Monats/Jahreswerte), ISO 50001-tauglich
# Deutsch, robuste Importe, klare UI, nachvollziehbare Rechnung, Excel & PDF

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

# =========================
# Hilfsfunktionen
# =========================
def parse_datetime(s: str) -> Optional[datetime]:
    s = str(s).strip().replace(" ", "T")
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def minutes(hhmm: str) -> int:
    h, m = hhmm.split(":")
    return int(h) * 60 + int(m)

def overlap_minutes(a0: int, a1: int, b0: int, b1: int) -> int:
    return max(0, min(a1, b1) - max(a0, b0))

def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))

# =========================
# Datenklassen
# =========================
@dataclass
class Zeitfenster:
    start: str        # "HH:MM"
    ende: str         # "HH:MM"
    aktiv: bool       # an/aus
    T_soll_C: float   # Soll-Zulufttemperatur
    V_m3h: Optional[float] = None  # None = Standard (Normal/Absenk)

@dataclass
class Tagesplan:
    tag: int                 # 0=Mo..6=So
    normal: List[Zeitfenster]
    absenk: List[Zeitfenster]

@dataclass
class Defaults:
    T_normal_C: float = 20.0
    T_absenk_C: float = 17.0
    V_normal_m3h: float = 5000.0     # Standard, wenn kein Override im Fenster
    V_absenk_m3h: Optional[float] = 2000.0  # None = wie normal

@dataclass
class Anlage:
    id: str
    name: str
    V_nominal_m3h: float
    anzahl: int
    wrg: bool
    eta_t: float               # 0..1 (Wärmerückgewinnung)
    fan_kW: Optional[float]    # Gesamtleistung bei V_nominal
    SFP_kW_per_m3s: Optional[float]
    wochenplan: List[Tagesplan]
    notizen: str = ""

WOCHENTAGE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]

def leerer_wochenplan(defaults: Defaults) -> List[Tagesplan]:
    plan: List[Tagesplan] = []
    for d in range(7):
        normal = [Zeitfenster("06:30", "17:00", True, defaults.T_normal_C, None)] if d < 5 else []
        absenk = [Zeitfenster("17:00", "06:30", True, defaults.T_absenk_C, defaults.V_absenk_m3h)] if d < 5 else []
        plan.append(Tagesplan(tag=d, normal=normal, absenk=absenk))
    return plan

def normiere_wochenplan(plan: List[Tagesplan]) -> List[List[Tuple[int,int,str,float,Optional[float]]]]:
    """
    Liefert je Tag eine Liste (startMin, endMin, modus, T_soll, V_override)
    modus = "Normal" | "Absenk"
    Über-Mitternacht wird gesplittet
    """
    norm: List[List[Tuple[int,int,str,float,Optional[float]]]] = [[] for _ in range(7)]
    def add(tag:int, f:Zeitfenster, modus:str):
        if not f.aktiv: return
        s = minutes(f.start); e = minutes(f.ende)
        if s == e: return
        if e > s:
            norm[tag].append((s, e, modus, float(f.T_soll_C), None if f.V_m3h is None else float(f.V_m3h)))
        else:
            norm[tag].append((s, 1440, modus, float(f.T_soll_C), None if f.V_m3h is None else float(f.V_m3h)))
            norm[(tag+1)%7].append((0, e, modus, float(f.T_soll_C), None if f.V_m3h is None else float(f.V_m3h)))
    for d in plan:
        for f in d.normal: add(d.tag, f, "Normal")
        for f in d.absenk: add(d.tag, f, "Absenk")
    for i in range(7): norm[i].sort(key=lambda x: x[0])
    return norm

# =========================
# robuste TRY-CSV Verarbeitung
# =========================
def parse_try_csv(raw: pd.DataFrame) -> Tuple[pd.DataFrame, list, list]:
    """Erkennt Spalten, behandelt 24:00, reindiziert auf stündliches Raster, liefert df[datetime,T_out_C]."""
    def _find_column(df: pd.DataFrame, aliases: list[str]) -> Optional[str]:
        low = {c.lower().strip(): c for c in df.columns}
        for a in aliases:
            if a in low: return low[a]
        for c in df.columns:
            cl = c.lower().strip()
            if any(a in cl for a in aliases): return c
        return None

    dt_aliases = ["datetime","date_time","date/time","date","timestamp","zeit","zeitstempel","datestamp","datetime_local","datetime_utc"]
    t_aliases  = ["t_out_c","t_out","tout","temp_out","temperature_out","aussen","außen","ta","t2m","t_out(°c)"]

    dt_col = _find_column(raw, dt_aliases) or st.selectbox("Datums-/Zeitspalte wählen", raw.columns)
    t_col  = _find_column(raw, t_aliases)  or st.selectbox("Außentemperatur-Spalte wählen", raw.columns)

    df = raw[[dt_col, t_col]].copy()
    # Temperatur zu float (Komma/°C zulassen)
    if df[t_col].dtype == object:
        df[t_col] = (df[t_col].astype(str)
                                 .str.replace(",", ".", regex=False)
                                 .str.replace("°C", "", regex=False)
                                 .str.strip())
    df[t_col] = pd.to_numeric(df[t_col], errors="coerce")
    df = df.rename(columns={t_col: "T_out_C"})

    # 24:00 → 00:00 + 1 Tag
    def _fix_24h(s: str):
        s = str(s).strip().replace(" ", "T")
        had_24 = ("T24:" in s) or s.endswith("24:00")
        s2 = s.replace("T24:", "T00:").replace(" 24:", " 00:")
        dt = pd.to_datetime(s2, errors="coerce")
        if had_24 and pd.notna(dt):
            dt = dt + pd.Timedelta(days=1)
        return dt

    df["datetime"] = raw[dt_col].astype(str).apply(_fix_24h)
    df = df[["datetime","T_out_C"]].dropna().sort_values("datetime")

    # Doppelte Zeitstempel – letzten Wert pro Zeit behalten
    df = df.groupby("datetime", as_index=False)["T_out_C"].last()

    # Reindex auf stündliches Raster
    problems = []
    if not df.empty:
        start = df["datetime"].min().replace(minute=0, second=0, microsecond=0)
        end   = df["datetime"].max().replace(minute=0, second=0, microsecond=0)
        full_index = pd.date_range(start, end, freq="H")
        s = df.set_index("datetime")["T_out_C"].reindex(full_index)
        missing = int(s.isna().sum())
        if missing > 0:
            problems.append(f"{missing} fehlende Stunde(n) – Quelle prüfen.")
            # Optional: Interpolation aktivieren
            # s = s.interpolate(limit_direction="both")
        df = s.reset_index().rename(columns={"index":"datetime", 0:"T_out_C"})
    years = sorted(df["datetime"].dt.year.unique().tolist()) if not df.empty else []
    return df, years, problems

# =========================
# Berechnung
# =========================
def berechne(try_df: pd.DataFrame, anlage: Anlage, defaults: Defaults):
    """Monats-/Jahres-Summen + optionales Stundenprotokoll (für Nachvollziehbarkeit)."""
    V_normal_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_normal_total = float(clamp(V_normal_total, 0.0, 500000.0))

    norm = normiere_wochenplan(anlage.wochenplan)

    protokoll = []  # für Transparenz
    rec = []

    for i in range(len(try_df)):
        dt0 = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])
        tag = dt0.weekday()
        m0 = dt0.hour * 60 + dt0.minute
        m1 = m0 + 60

        # alle Fenster des Tages durchgehen
        stunden_Q_kWh = 0.0
        stunden_E_kWh = 0.0
        stunden_h_vent = 0.0

        for (s, e, modus, T_soll, V_override) in norm[tag]:
            ol = overlap_minutes(m0, m1, s, e)
            if ol <= 0: 
                continue
            anteil_h = ol / 60.0

            # Setpoints
            if V_override is not None:
                V_m3h = max(0.0, float(V_override))
            else:
                V_m3h = V_normal_total if modus == "Normal" else (V_normal_total if defaults.V_absenk_m3h is None else float(defaults.V_absenk_m3h))

            # Wärmebedarf – nur bei Heizfall
            dT = max(0.0, float(T_soll) - Tout)
            dT_eff = (1.0 - clamp(float(anlage.eta_t), 0.0, 1.0)) * dT if anlage.wrg else dT

            # **KORREKT:** kWh für den Minutenanteil
            Q_kWh = 0.00034 * V_m3h * dT_eff * anteil_h

            # Ventilator
            P_fan_kW = 0.0
            if V_m3h > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_fan_kW = float(anlage.SFP_kW_per_m3s) * (V_m3h / 3600.0)
                elif anlage.fan_kW is not None:
                    # v1 linear mit Volumenstrom
                    ref = max(1.0, V_normal_total)
                    P_fan_kW = float(anlage.fan_kW) * clamp(V_m3h / ref, 0.0, 1.0)
            E_kWh = P_fan_kW * anteil_h

            stunden_Q_kWh += Q_kWh
            stunden_E_kWh += E_kWh
            stunden_h_vent += anteil_h if V_m3h > 0 else 0.0

            # Protokolleintrag (nur erste N für Anzeige)
            protokoll.append({
                "Zeit": dt0,
                "Modus": modus,
                "Anteil_h": round(anteil_h, 3),
                "T_out": round(Tout, 1),
                "T_soll": float(T_soll),
                "ΔT_eff": round(dT_eff, 2),
                "V_m3h": round(V_m3h, 0),
                "Q_kWh": round(Q_kWh, 4),
                "P_fan_kW": round(P_fan_kW, 3),
                "E_kWh": round(E_kWh, 4)
            })

        rec.append({
            "datetime": dt0,
            "year": dt0.year,
            "month": dt0.month,
            "kWh_th": stunden_Q_kWh,
            "kWh_el": stunden_E_kWh,
            "Betriebsstunden_Vent": stunden_h_vent,
        })

    dfh = pd.DataFrame.from_records(rec)
    monat = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    jahr  = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    return monat, jahr, pd.DataFrame(protokoll)

# =========================
# Exporte
# =========================
def build_excel(monat: pd.DataFrame, jahr: pd.DataFrame, protokoll: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        m = monat.copy(); j = jahr.copy(); p = protokoll.copy()
        # Rundungen für Lesbarkeit
        for col in ("kWh_th","kWh_el"):
            if col in m: m[col] = m[col].round(0)
            if col in j: j[col] = j[col].round(0)
        if "Betriebsstunden_Vent" in m: m["Betriebsstunden_Vent"] = m["Betriebsstunden_Vent"].round(1)
        if "Betriebsstunden_Vent" in j: j["Betriebsstunden_Vent"] = j["Betriebsstunden_Vent"].round(1)
        if not p.empty:
            p["Q_kWh"] = p["Q_kWh"].round(4)
            p["E_kWh"] = p["E_kWh"].round(4)

        m.to_excel(writer, index=False, sheet_name="Monate")
        j.to_excel(writer, index=False, sheet_name="Jahr")
        p.to_excel(writer, index=False, sheet_name="Protokoll")

    return out.getvalue()

def build_pdf(info: str, defaults: Defaults, anlage: Anlage, monat: pd.DataFrame, jahr: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab ist nicht installiert.")
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]; H1.fontSize = 14
    H2 = styles["Heading2"]; H2.fontSize = 12
    N  = styles["BodyText"]; N.leading = 14

    story = []
    story += [Paragraph("ISO 50001 – Heizenergie Lüftungsanlagen (v1, vereinfachtes Verfahren)", H1), Spacer(1, 6)]
    story += [Paragraph(f"Erstellt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1, 8)]
    story += [Paragraph("Quelle / TRY", H2), Paragraph(info or "TRY-CSV eingelesen (stündlich).", N), Spacer(1, 6)]
    story += [Paragraph("Annahmen & Parameter", H2),
              Paragraph(f"T_normal: {defaults.T_normal_C} °C; T_absenk: {defaults.T_absenk_C} °C; V_normal: {defaults.V_normal_m3h} m³/h; "
                        f"V_absenk: {defaults.V_absenk_m3h if defaults.V_absenk_m3h is not None else 'wie normal'} m³/h. "
                        f"WRG: {'ja' if anlage.wrg else 'nein'} (η_t={anlage.eta_t}). "
                        f"Ventilator: {'SFP='+str(anlage.SFP_kW_per_m3s)+' kW/(m³/s)' if anlage.SFP_kW_per_m3s is not None else 'fan_kW='+str(anlage.fan_kW)+' kW'}.", N),
              Spacer(1, 6)]
    story += [Paragraph("Methodik", H2),
              Paragraph("Für jede Stunde wird das aktive Zeitfenster (Normal/Absenk) ermittelt. "
                        "Heizfall: ΔT = max(0, T_soll − T_out); mit WRG: ΔT_eff = (1 − η_t)·ΔT. "
                        "Wärme je Stunde: Q_kWh = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h. "
                        "Ventilator: P_fan = SFP·(V/3600) oder fan_kW·(V/V_nominal), Energie E_kWh = P_fan·Anteil_h. "
                        "Aggregation zu Monat/Jahr.", N),
              Spacer(1, 6)]

    if not jahr.empty:
        j = jahr.copy()
        j[["kWh_th","kWh_el"]] = j[["kWh_th","kWh_el"]].round(0)
        j["Betriebsstunden_Vent"] = j["Betriebsstunden_Vent"].round(1)
        data = [["Jahr","kWh_th","kWh_el","Betriebsstunden Vent."], *j.values.tolist()]
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(1,1),(-1,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ]))
        story += [Paragraph("Ergebnisse – Jahreswerte", H2), tbl, Spacer(1, 6)]

    if not monat.empty:
        m = monat.copy()
        m[["kWh_th","kWh_el"]] = m[["kWh_th","kWh_el"]].round(0)
        m["Betriebsstunden_Vent"] = m["Betriebsstunden_Vent"].round(1)
        data = [["Jahr","Monat","kWh_th","kWh_el","Betriebsstunden Vent."], *m.astype({"year":int,"month":int}).values.tolist()]
        tbl = Table(data, hAlign="LEFT", repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(2,1),(-1,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ]))
        story += [Paragraph("Ergebnisse – Monate", H2), tbl]

    doc.build(story); buf.seek(0); return buf.read()

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001)", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session initialisieren
for k, v in [
    ("try_df", None), ("try_info",""),
    ("monat_df", None), ("jahr_df", None),
    ("protokoll_df", None),
    ("auto_calc", False), ("trigger_recalc", False),
]:
    if k not in st.session_state:
        st.session_state[k] = v

st.markdown("**Ziel:** Monats-/Jahreswerte (kWh Wärme / kWh elektrisch) aus TRY‑CSV und Anlagenparametern. "
            "Die Tabellen werden **zuerst angezeigt**, **darunter** gibt es die Download‑Links.")

# ---------- 1) TRY-CSV laden (robust) ----------
with st.expander("1) TRY‑CSV laden", expanded=True):
    f = st.file_uploader("TRY‑CSV (stündlich: Datum/Zeit + Außentemperatur)", type=["csv"])
    st.session_state["auto_calc"] = st.checkbox("Automatisch nach Upload berechnen", value=st.session_state["auto_calc"])

    if f is not None:
        raw = pd.read_csv(f)
        df, years, problems = parse_try_csv(raw)
        if df.empty:
            st.error("Kein gültiger Datensatz (prüfe Spalten & Format).")
        else:
            tmin, tmax = df["T_out_C"].min(), df["T_out_C"].max()
            st.session_state["try_df"]  = df
            st.session_state["try_info"] = f"Datensätze: {len(df)} | Jahre: {', '.join(map(str, years))} | T_out: {round(tmin,1)}…{round(tmax,1)} °C"
            st.success("TRY‑CSV eingelesen.")
            st.text(st.session_state["try_info"])
            if problems:
                st.warning(" / ".join(problems))
            if st.session_state["auto_calc"]:
                st.session_state["trigger_recalc"] = True
    else:
        if st.session_state["try_df"] is not None:
            st.info("TRY‑CSV bereits geladen.")
            if st.session_state["try_info"]:
                st.text(st.session_state["try_info"])

# ---------- 2) Standard‑Parameter ----------
with st.expander("2) Standard‑Parameter (Temperaturen, Volumenströme)", expanded=True):
    c = st.columns(4)
    T_normal = c[0].number_input("Soll-Zuluft NORMAL [°C]", value=20.0, step=0.5)
    T_absenk = c[1].number_input("Soll-Zuluft ABSENK [°C]", value=17.0, step=0.5)
    V_normal = c[2].number_input("Volumenstrom NORMAL [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    V_absenk = c[3].number_input("Volumenstrom ABSENK [m³/h] (leer=wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0)
    defaults = Defaults(T_normal_C=float(T_normal), T_absenk_C=float(T_absenk), V_normal_m3h=float(V_normal), V_absenk_m3h=float(V_absenk))

# ---------- 3) Anlage & Betrieb ----------
with st.expander("3) Anlage & Betrieb / Absenkzeiten", expanded=True):
    r1 = st.columns([1.2,1,1,1])
    anlage_id   = r1[0].text_input("Anlagen-ID", value="A01")
    anlage_name = r1[1].text_input("Bezeichnung", value="Zuluft Beispiel")
    V_nominal   = r1[2].number_input("V_nominal (Einzelanlage) [m³/h]", value=V_normal, min_value=500.0, max_value=500000.0, step=100.0)
    anzahl      = r1[3].number_input("Anzahl gleicher Anlagen", value=1, min_value=1, max_value=100, step=1)

    r2 = st.columns(3)
    wrg    = r2[0].checkbox("WRG vorhanden", value=True)
    eta_t  = r2[1].number_input("η_t WRG (0–1)", value=0.7, min_value=0.0, max_value=1.0, step=0.05)
    vent_m = r2[2].selectbox("Ventilator-Modell", ("SFP (kW/(m³/s))","fan_kW gesamt"), index=1)

    fan_kW = None; SFP = None
    if vent_m.startswith("SFP"):
        SFP = st.number_input("SFP [kW/(m³/s)]", value=1.8, min_value=0.0, max_value=10.0, step=0.1)
    else:
        fan_kW = st.number_input("fan_kW (bei V_nominal, gesamt)", value=5.0, min_value=0.0, max_value=1000.0, step=0.1)

    # Wochenplan – einfache, übersichtliche Eingabe:
    if "wochenplan" not in st.session_state:
        st.session_state["wochenplan"] = leerer_wochenplan(defaults)
    wp: List[Tagesplan] = st.session_state["wochenplan"]

    st.caption("Für jeden Tag können mehrere Normal- oder Absenk-Fenster definiert werden. "
               "Über‑Mitternacht (z. B. 17:00–06:30) ist erlaubt.")
    for d in range(7):
        st.write(f"**{WOCHENTAGE[d]}**")
        day = wp[d]
        # NORMAl
        st.markdown("_Normalbetrieb (T & V standardmäßig aus Abschnitt 2)_")
        ncols = st.columns([1,1,1,1,0.7])
        if st.button("+ Normal-Fenster", key=f"add_n_{d}"):
            day.normal.append(Zeitfenster("08:00","16:00",True,defaults.T_normal_C,None))
            st.experimental_rerun()
        del_n = None
        for i, f in enumerate(day.normal):
            f.start = ncols[0].text_input("Start", key=f"n_s_{d}_{i}", value=f.start)
            f.ende  = ncols[1].text_input("Ende",  key=f"n_e_{d}_{i}", value=f.ende)
            f.aktiv = ncols[2].checkbox("aktiv", key=f"n_a_{d}_{i}", value=f.aktiv)
            f.T_soll_C = ncols[3].number_input("T_soll [°C]", key=f"n_t_{d}_{i}", value=float(f.T_soll_C), step=0.5)
            if ncols[4].button("–", key=f"n_del_{d}_{i}"): del_n = i
        if del_n is not None:
            day.normal.pop(del_n); st.experimental_rerun()

        # ABSENK
        st.markdown("_Absenkbetrieb (eigene T & ggf. V)_")
        acols = st.columns([1,1,1,1,1,0.7])
        if st.button("+ Absenk-Fenster", key=f"add_a_{d}"):
            day.absenk.append(Zeitfenster("17:00","06:30",True,defaults.T_absenk_C,defaults.V_absenk_m3h))
            st.experimental_rerun()
        del_a = None
        for i, f in enumerate(day.absenk):
            f.start = acols[0].text_input("Start", key=f"a_s_{d}_{i}", value=f.start)
            f.ende  = acols[1].text_input("Ende",  key=f"a_e_{d}_{i}", value=f.ende)
            f.aktiv = acols[2].checkbox("aktiv", key=f"a_a_{d}_{i}", value=f.aktiv)
            f.T_soll_C = acols[3].number_input("T_soll [°C]", key=f"a_t_{d}_{i}", value=float(f.T_soll_C), step=0.5)
            vval = 0.0 if f.V_m3h is None else float(f.V_m3h)
            newv = acols[4].number_input("V [m³/h] (0 = Standard)", key=f"a_v_{d}_{i}", value=vval, min_value=0.0, max_value=500000.0, step=100.0)
            f.V_m3h = None if newv == 0.0 else float(newv)
            if acols[5].button("–", key=f"a_del_{d}_{i}"): del_a = i
        if del_a is not None:
            day.absenk.pop(del_a); st.experimental_rerun()

# ---------- 4) Berechnen ----------
def _do_calc():
    df = st.session_state["try_df"]
    if df is None or df.empty:
        st.error("Bitte zuerst eine gültige TRY‑CSV laden."); return
    anlage = Anlage(
        id=anlage_id, name=anlage_name,
        V_nominal_m3h=float(V_nominal), anzahl=int(anzahl),
        wrg=bool(wrg), eta_t=float(eta_t),
        fan_kW=float(fan_kW) if fan_kW is not None else None,
        SFP_kW_per_m3s=float(SFP) if SFP is not None else None,
        wochenplan=st.session_state["wochenplan"]
    )
    monat, jahr, prot = berechne(df, anlage, defaults)
    st.session_state["monat_df"] = monat
    st.session_state["jahr_df"] = jahr
    st.session_state["protokoll_df"] = prot
    st.success("Berechnung abgeschlossen.")

if st.button("Berechnen", type="primary"):
    _do_calc()

# Auto-Calc nach Upload
if st.session_state["trigger_recalc"]:
    _do_calc()
    st.session_state["trigger_recalc"] = False

# ---------- 5) Ergebnisse – sichtbar, dann Download ----------
m = st.session_state["monat_df"]; j = st.session_state["jahr_df"]; p = st.session_state["protokoll_df"]
if m is not None and j is not None:
    # Anzeige gerundet
    m_view = m.copy(); j_view = j.copy()
    for col in ("kWh_th","kWh_el"):
        if col in m_view: m_view[col] = m_view[col].round(0)
        if col in j_view: j_view[col] = j_view[col].round(0)
    if "Betriebsstunden_Vent" in m_view: m_view["Betriebsstunden_Vent"] = m_view["Betriebsstunden_Vent"].round(1)
    if "Betriebsstunden_Vent" in j_view: j_view["Betriebsstunden_Vent"] = j_view["Betriebsstunden_Vent"].round(1)

    st.subheader("Ergebnisse – Jahreswerte (sichtbar)")
    st.dataframe(j_view, use_container_width=True)

    st.subheader("Ergebnisse – Monate (sichtbar)")
    st.dataframe(m_view, use_container_width=True)

    st.subheader("Stunden‑Protokoll (Ausschnitt zur Kontrolle)")
    if p is not None and not p.empty:
        st.dataframe(p.head(100), use_container_width=True)
    else:
        st.info("Kein Protokoll verfügbar.")

    # Plausibilitäts-Checks
    st.subheader("Plausibilitäts‑Check")
    if not j_view.empty:
        th = float(j_view["kWh_th"].sum())
        el = float(j_view["kWh_el"].sum())
        hints = []
        if th <= 0: hints.append("Wärmebedarf = 0 kWh → Prüfe Soll‑Temperaturen und Zeitfenster.")
        if th > 10_000_000: hints.append("Sehr hoher Wärmebedarf (>10 GWh) → Prüfe Volumenstrom/Zeiten/Formel.")
        if el > th*0.5: hints.append("Ventilatorstrom sehr hoch im Verhältnis zur Wärme → SFP/Fan‑Werte prüfen.")
        if hints:
            for h in hints: st.warning(h)
        else:
            st.success("Werte im erwartbaren Bereich (grober Heuristik‑Check).")

    # Downloads UNTER den Tabellen
    st.subheader("Downloads")
    c1, c2, c3 = st.columns(3)
    c1.download_button("Excel (Monate + Jahr + Protokoll)", build_excel(m_view, j_view, p if p is not None else pd.DataFrame()),
                       file_name="Heizenergie_Auswertung.xlsx")
    c2.download_button("CSV – Monate (sichtbar)", m_view.to_csv(index=False).encode("utf-8"),
                       file_name="Heizenergie_Monate.csv", mime="text/csv")
    if REPORTLAB_OK:
        # Dummy-Anlage nur für PDF-Kopf
        anlage_dummy = Anlage(anlage_id, anlage_name, float(V_nominal), int(anzahl), bool(wrg), float(eta_t),
                              float(fan_kW) if fan_kW is not None else None,
                              float(SFP) if SFP is not None else None,
                              st.session_state["wochenplan"])
        st.download_button("PDF – ISO 50001 Kurzbericht",
                           build_pdf(st.session_state["try_info"], defaults, anlage_dummy, m_view, j_view),
                           file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
    else:
        st.info("PDF‑Export: Paket 'reportlab' nicht installiert.")
