# app.py – Heizenergie-Rechner Lüftungsanlagen (TRY → Monats/Jahreswerte)
# Deutsch · robuste TRY-Prüfung · klare UI · nachvollziehbare Formeln · Detail- & Überschlagsrechnung · Excel/PDF

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
import sys

# ---------------- Sidebar: Build-Info ----------------
APP_VERSION = "2025-08-25_final_v1"
st.sidebar.caption(f"Build: {APP_VERSION} · Python {sys.version.split()[0]} · Streamlit {st.__version__}")
if st.sidebar.button("Cache leeren & neu laden"):
    st.cache_data.clear(); st.cache_resource.clear(); st.experimental_rerun()

# ---------------- Optional: PDF ----------------
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm  # WICHTIG: nicht als Variablenname überschreiben
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

# ---------------- Helfer ----------------
def minutes(hhmm: str) -> int:
    h, m = hhmm.split(":")
    return int(h)*60 + int(m)

def overlap_mins(a0: int, a1: int, b0: int, b1: int) -> int:
    return max(0, min(a1, b1) - max(a0, b0))

def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))

def add_sum_row(df: pd.DataFrame, label_col: str = None, label: str = "Summe") -> pd.DataFrame:
    """Fügt eine Summenzeile ans Ende an (numerische Spalten werden summiert)."""
    if df is None or df.empty:
        return df
    sums = df.select_dtypes("number").sum(numeric_only=True)
    row = {c: sums.get(c, "") for c in df.columns}
    if label_col and label_col in df.columns:
        row[label_col] = label
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

# ---------------- Datenklassen ----------------
@dataclass
class Zeitfenster:
    start: str
    ende: str
    aktiv: bool
    T_soll_C: float
    V_m3h: Optional[float] = None  # None = Standard (aus Abschnitt "Standardwerte")

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
    V_absenk_m3h: Optional[float] = 2000.0  # None = wie normal

@dataclass
class Anlage:
    id: str
    name: str
    V_nominal_m3h: float
    anzahl: int
    wrg: bool
    eta_t: float
    fan_kW: Optional[float]
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
    """ je Tag: (startMin, endMin, 'Normal'|'Absenk', T_soll, V_override) – über Mitternacht gesplittet """
    out = [[] for _ in range(7)]
    def add(tag: int, f: Zeitfenster, modus: str):
        if not f.aktiv: return
        s, e = minutes(f.start), minutes(f.ende)
        if s == e: return
        Vovr = None if f.V_m3h is None else float(f.V_m3h)
        if e > s:
            out[tag].append((s,e,modus,float(f.T_soll_C),Vovr))
        else:
            out[tag].append((s,1440,modus,float(f.T_soll_C),Vovr))
            out[(tag+1)%7].append((0,e,modus,float(f.T_soll_C),Vovr))
    for d in plan:
        for f in d.normal: add(d.tag, f, "Normal")
        for f in d.absenk: add(d.tag, f, "Absenk")
    for i in range(7): out[i].sort(key=lambda x: x[0])
    return out

# ---------------- TRY-CSV Parser (robust) ----------------
def parse_try_csv(raw: pd.DataFrame, interpolate_missing: bool) -> Tuple[pd.DataFrame, list, list]:
    """
    Liefert df[datetime,T_out_C], Jahre[], Hinweise[].
    - 24:00 → 00:00 + 1 Tag (vor GroupBy!)
    - GroupBy(datetime).last() (echte Dubletten zusammenfassen)
    - Reindex je Jahr auf Jan-01 00:00 … Dec-31 23:00 (8760/8784)
    - Optionale Interpolation fehlender Stunden
    """
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

    # Temperatur -> float
    if df[t_col].dtype == object:
        df[t_col] = (df[t_col].astype(str)
                               .str.replace(",", ".", regex=False)
                               .str.replace("°C", "", regex=False)
                               .str.strip())
    df[t_col] = pd.to_numeric(df[t_col], errors="coerce")
    df = df.rename(columns={t_col: "T_out_C"})

    # 24:00 sauber -> 00:00 + 1 Tag
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

    # Dubletten zusammenfassen (nach 24:00-Korrektur!)
    before = len(df)
    df = df.groupby("datetime", as_index=False)["T_out_C"].last()
    dups = before - len(df)

    years = sorted(df["datetime"].dt.year.unique().tolist())
    hints = []

    # Für jedes Jahr eigenes volles Stundenraster (Jan-01 00:00 bis Dec-31 23:00)
    parts = []
    missing_total = 0
    for y in years:
        start = pd.Timestamp(y,1,1,0)
        end   = pd.Timestamp(y,12,31,23)
        idx = pd.date_range(start, end, freq="H")
        s = df[df["datetime"].dt.year==y].set_index("datetime")["T_out_C"].reindex(idx)
        miss = int(s.isna().sum())
        missing_total += miss
        if miss > 0 and interpolate_missing:
            s = s.interpolate(limit_direction="both")
        parts.append(s.reset_index().rename(columns={"index":"datetime", 0:"T_out_C"}))

    if parts:
        out = pd.concat(parts, ignore_index=True)
    else:
        out = pd.DataFrame(columns=["datetime","T_out_C"])

    if missing_total > 0:
        if interpolate_missing:
            hints.append(f"{missing_total} fehlende Stunde(n) wurden interpoliert.")
        else:
            hints.append(f"{missing_total} fehlende Stunde(n) – Quelle prüfen (oder Interpolation aktivieren).")
    if dups > 0:
        hints.append(f"{dups} doppelte Zeitstempel zusammengefasst.")

    return out, years, hints

# ---------------- Detail-Berechnung ----------------
def berechne_detail(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults):
    """Stündliche Detailrechnung (Monate, Jahr, Protokoll (Ausschnitt), stundenweises Vollprotokoll)."""
    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = normiere_wochenplan(anlage.wochenplan)

    rows = []; prot = []; prot_full = []

    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])
        d = dt.weekday()
        m0, m1 = dt.hour*60 + dt.minute, dt.hour*60 + dt.minute + 60

        Q_h = 0.0; E_h = 0.0; fan_h = 0.0

        for (s,e,modus,T_soll,V_ovr) in plan[d]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            # Volumenstrom
            if V_ovr is not None:
                V = max(0.0, float(V_ovr))
            else:
                V = V_norm_total if modus=="Normal" else (V_norm_total if defs.V_absenk_m3h is None else float(defs.V_absenk_m3h))

            # Heizenergie (kWh)
            dT = max(0.0, float(T_soll) - Tout)
            dT_eff = (1.0 - clamp(float(anlage.eta_t),0.0,1.0))*dT if anlage.wrg else dT
            Q_kWh = 0.00034 * V * dT_eff * anteil

            # Ventilator
            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = float(anlage.SFP_kW_per_m3s) * (V/3600.0)
                elif anlage.fan_kW is not None:
                    P_kW = float(anlage.fan_kW) * (V / max(1.0, V_norm_total))
            E_kWh = P_kW * anteil

            Q_h += Q_kWh; E_h += E_kWh; fan_h += (anteil if V>0 else 0.0)

            rec_common = {
                "Zeit":dt,"Modus":modus,"Anteil [h]":round(anteil,3),
                "T_out [°C]":round(Tout,1),"T_soll [°C]":round(T_soll,1),
                "ΔT_eff [K]":round(dT_eff,2),"V [m³/h]":round(V,0),
                "Wärme [kWh]":round(Q_kWh,6),"P_fan [kW]":round(P_kW,4),"Strom [kWh]":round(E_kWh,6)
            }
            prot_full.append(rec_common)
            if i < 500:  # UI-Ausschnitt
                prot.append(rec_common)

        rows.append({"datetime":dt,"year":dt.year,"month":dt.month,
                     "kWh_th":Q_h,"kWh_el":E_h,"Betriebsstunden_Vent":fan_h})

    dfh = pd.DataFrame.from_records(rows)
    mon = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    jahr = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent"]].sum()
    return mon, jahr, pd.DataFrame(prot), pd.DataFrame(prot_full)

# ---------------- Überschlagsrechnung (Kontrollrechner) ----------------
def berechne_ueberschlag(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults):
    """
    Vereinfachung: statt stündlicher T_out nutzen wir den Monats-Mittelwert T_out_m.
    Das Zeitfenster/Volumenstrom je Stunde bleibt wie in der Detailrechnung (nur ΔT wird durch ΔT_m ersetzt).
    Ergebnis: Monats- und Jahreswerte (kWh_th/kWh_el) als Kontrollgröße.
    """
    if try_df is None or try_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Monatsmittel T_out
    try_df = try_df.copy()
    try_df["year"] = try_df["datetime"].dt.year
    try_df["month"] = try_df["datetime"].dt.month
    t_mean = try_df.groupby(["year","month"], as_index=False)["T_out_C"].mean().rename(columns={"T_out_C":"T_out_mean"})

    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = normiere_wochenplan(anlage.wochenplan)

    rec = []
    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        y, m = dt.year, dt.month
        Tout_m = float(t_mean[(t_mean["year"]==y)&(t_mean["month"]==m)]["T_out_mean"].values[0])

        day = dt.weekday()
        m0, m1 = dt.hour*60 + dt.minute, dt.hour*60 + dt.minute + 60

        Q_h = 0.0; E_h = 0.0
        for (s,e,modus,T_soll,V_ovr) in plan[day]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            if V_ovr is not None:
                V = max(0.0, float(V_ovr))
            else:
                V = V_norm_total if modus=="Normal" else (V_norm_total if defs.V_absenk_m3h is None else float(defs.V_absenk_m3h))

            # ΔT_eff mit Monatsmittel
            dT_m = max(0.0, float(T_soll) - Tout_m)
            dT_eff_m = (1.0 - clamp(float(anlage.eta_t),0.0,1.0))*dT_m if anlage.wrg else dT_m
            Q_kWh = 0.00034 * V * dT_eff_m * anteil

            # Ventilator wie Detail (abhängig von V, nicht von T)
            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = float(anlage.SFP_kW_per_m3s) * (V/3600.0)
                elif anlage.fan_kW is not None:
                    P_kW = float(anlage.fan_kW) * (V / max(1.0, V_norm_total))
            E_kWh = P_kW * anteil

            Q_h += Q_kWh; E_h += E_kWh

        rec.append({"year":y,"month":m,"kWh_th":Q_h,"kWh_el":E_h})

    dfh = pd.DataFrame.from_records(rec)
    mon = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el"]].sum()
    jahr = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el"]].sum()
    return mon, jahr

# ---------------- Exporte ----------------
def xlsx_export(mon: pd.DataFrame, jahr: pd.DataFrame, prot: pd.DataFrame, mon_ue: pd.DataFrame, jahr_ue: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        m = mon.copy(); j = jahr.copy(); p = prot.copy()
        mu = mon_ue.copy(); ju = jahr_ue.copy()

        # Runden & Deutsch
        for col in ("kWh_th","kWh_el"):
            if col in m: m[col] = m[col].round(0)
            if col in j: j[col] = j[col].round(0)
            if col in mu: mu[col] = mu[col].round(0)
            if col in ju: ju[col] = ju[col].round(0)
        if "Betriebsstunden_Vent" in m: m["Betriebsstunden_Vent"] = m["Betriebsstunden_Vent"].round(1)
        if "Betriebsstunden_Vent" in j: j["Betriebsstunden_Vent"] = j["Betriebsstunden_Vent"].round(1)

        m_de = m.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]","Betriebsstunden_Vent":"Betriebsstd. Vent."})
        j_de = j.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]","Betriebsstunden_Vent":"Betriebsstd. Vent."})
        mu_de = mu.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        ju_de = ju.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})

        add_sum = lambda df, lab: add_sum_row(df, label_col="Monat", label=lab) if "Monat" in df.columns else add_sum_row(df)

        m_de = add_sum(m_de, "Summe")
        j_de = add_sum(j_de)
        mu_de = add_sum(mu_de, "Summe")
        ju_de = add_sum(ju_de)

        m_de.to_excel(w, index=False, sheet_name="Monate (Detail)")
        j_de.to_excel(w, index=False, sheet_name="Jahr (Detail)")
        mu_de.to_excel(w, index=False, sheet_name="Monate (Überschlag)")
        ju_de.to_excel(w, index=False, sheet_name="Jahr (Überschlag)")
        if not p.empty:
            p.to_excel(w, index=False, sheet_name="Protokoll (Ausschnitt)")
    return out.getvalue()

def pdf_export(info: str, defs: Defaults, anl: Anlage, mon: pd.DataFrame, jahr: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab nicht installiert.")
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet(); H1=styles["Heading1"]; H1.fontSize=14; H2=styles["Heading2"]; H2.fontSize=12; N=styles["BodyText"]; N.leading=14
    story=[]
    story+=[Paragraph("ISO 50001 – Heizenergie Lüftungsanlagen (v1, vereinfachtes Verfahren)", H1), Spacer(1,6)]
    story+=[Paragraph(f"Erstellt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1,8)]
    story+=[Paragraph("TRY / Quelle", H2), Paragraph(info or "TRY-CSV (stündlich).", N), Spacer(1,6)]
    story+=[Paragraph("Annahmen & Parameter", H2),
            Paragraph(f"T_normal: {defs.T_normal_C} °C; T_absenk: {defs.T_absenk_C} °C; "
                      f"V_normal: {defs.V_normal_m3h} m³/h; V_absenk: {defs.V_absenk_m3h if defs.V_absenk_m3h is not None else 'wie normal'} m³/h. "
                      f"WRG: {'ja' if anl.wrg else 'nein'} (η_t={anl.eta_t}). "
                      f"Ventilator: "
                      f"{'SFP ' + str(anl.SFP_kW_per_m3s) + ' kW/(m³/s)' if anl.SFP_kW_per_m3s is not None else 'fan_kW ' + str(anl.fan_kW) + ' kW'}.", N),
            Spacer(1,6)]
    story+=[Paragraph("Methodik (Rechengang)", H2),
            Paragraph("Für jede Stunde wird das aktive Zeitfenster (Normal/Absenk) ermittelt. "
                      "Heizfall: ΔT = max(0, T_soll − T_out); mit WRG: ΔT_eff = (1 − η_t)·ΔT. "
                      "Wärme: Q_kWh = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h. "
                      "Ventilator: P_fan = SFP·(V/3600) oder fan_kW·(V/V_nominal), Energie E_kWh = P_fan·Anteil_h. "
                      "Aggregation zu Monaten/Jahr.", N), Spacer(1,6)]

    if not jahr.empty:
        j = jahr.copy(); j[["kWh_th","kWh_el"]] = j[["kWh_th","kWh_el"]].round(0); j["Betriebsstunden_Vent"]=j["Betriebsstunden_Vent"].round(1)
        data=[["Jahr","Wärme [kWh]","Strom Vent. [kWh]","Betriebsstd. Vent."], *j.values.tolist()]
        T=Table(data,hAlign="LEFT"); T.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("ALIGN",(1,1),(-1,-1),"RIGHT")]))
        story+=[Paragraph("Ergebnisse – Jahreswerte", H2), T, Spacer(1,6)]
    if not mon.empty:
        m = mon.copy(); m[["kWh_th","kWh_el"]]=m[["kWh_th","kWh_el"]].round(0); m["Betriebsstunden_Vent"]=m["Betriebsstunden_Vent"].round(1)
        data=[["Jahr","Monat","Wärme [kWh]","Strom Vent. [kWh]","Betriebsstd. Vent."], *m.astype({"year":int,"month":int}).values.tolist()]
        T=Table(data,hAlign="LEFT",repeatRows=1); T.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("ALIGN",(2,1),(-1,-1),"RIGHT")]))
        story+=[Paragraph("Ergebnisse – Monate", H2), T]

    doc.build(story); out.seek(0); return out.read()

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001)", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session-Init
for k, v in [
    ("try_df", None), ("try_info",""), ("years", []), ("interp", False), ("try_hints", []),
    ("anlage", None), ("defs", None), ("wochenplan", None),
    ("mon_df", None), ("jahr_df", None), ("prot_df", None), ("prot_full_df", None),
    ("mon_ue_df", None), ("jahr_ue_df", None),
    ("auto_calc", False), ("trigger", False),
]:
    if k not in st.session_state: st.session_state[k] = v

# --- Infofenster: Methodik & Kontrolle ---
with st.expander("ℹ️ Erläuterung – Rechenschritte & Formeln", expanded=False):
    st.markdown("""
**Datenquelle:** TRY‑CSV (stündlich) mit Spalten `datetime` und `T_out_C`.

**Zeitlogik:** Für jede Stunde wird geprüft, welche **Zeitfenster** (Normal/Absenk) aktiv sind. Über‑Mitternacht (z. B. 17:00–06:30) wird korrekt aufgeteilt.

**Heizfall & WRG:**  
ΔT = max(0, T_soll − T_out) → ΔT_eff = (1 − η_t) · ΔT (η_t = Wärmerückgewinnungs‑Wirkungsgrad)

**Wärme je Stunde:**  
Q_kWh = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h

**Ventilatorleistung:**  
P_fan = SFP · (V/3600) **oder** P_fan = fan_kW · (V/V_nominal)

**Ventilator‑Energie je Stunde:**  
E_kWh = P_fan · Anteil_h

**Aggregation:** Summen je Monat / Jahr.

**Warum 0,00034?** 0,34 Wh/(m³·K) = 0,00034 kWh/(m³·K).
""")

with st.expander("🧪 Datenprüfung & Qualität", expanded=False):
    st.markdown("""
- **24:00‑Zeitstempel** werden zu **00:00 + 1 Tag** korrigiert.  
- **Doppelte Zeitstempel** werden zusammengefasst (letzter Wert gewinnt).  
- Je Jahr wird auf **vollständiges Stundenraster** reindiziert (8760/8784 h).  
- **Fehlende Stunden** werden **optional interpoliert** (Checkbox im Upload).  

**Tipp:** Für Audits Interpolation deaktivieren, dann Lücken explizit ausweisen.
""")

with st.expander("🔍 Kontrollrechner (Überschlag) – Idee", expanded=False):
    st.markdown("""
Überschlag nutzt **Monats‑Mittelwerte** der Außentemperatur:  
ΔT_eff,Monat = (1 − η_t) · max(0, T_soll − ⟨T_out⟩_Monat)

Volumenströme & Zeitanteile bleiben wie im Zeitfensterplan, nur ΔT wird durch Monats‑Mittel ersetzt.  
Abweichungen zur Detailrechnung zeigen, ob die **Stunden‑Variabilität** stark wirkt.
""")

st.markdown("**Ablauf:** 1) TRY laden → 2) Anlagendaten → 3) Standardwerte & Zeiten → 4) Berechnen. Ergebnisse werden **sichtbar** gezeigt, **Downloads** stehen **darunter**.")

# ----- 1) TRY laden (FORM) -----
with st.form("form_try"):
    st.subheader("1) TRY‑CSV laden")
    f = st.file_uploader("TRY‑CSV (stündlich: Datum/Zeit + Außentemperatur)", type=["csv"])
    st.session_state["auto_calc"] = st.checkbox("Automatisch nach Upload berechnen", value=st.session_state["auto_calc"])
    st.session_state["interp"] = st.checkbox("Fehlende Stunden automatisch interpolieren", value=st.session_state["interp"])
    ok1 = st.form_submit_button("TRY übernehmen")
    if ok1:
        if f is None:
            st.error("Bitte CSV auswählen.")
        else:
            raw = pd.read_csv(f)
            df, years, hints = parse_try_csv(raw, interpolate_missing=st.session_state["interp"])
            if df.empty:
                st.error("Kein gültiger Datensatz.")
            else:
                tmin, tmax = df["T_out_C"].min(), df["T_out_C"].max()
                st.session_state["try_df"] = df
                st.session_state["years"] = years
                st.session_state["try_hints"] = hints
                st.session_state["try_info"] = f"Datensätze: {len(df)} | Jahre: {', '.join(map(str, years))} | T_out: {round(tmin,1)}…{round(tmax,1)} °C"
                st.success("TRY übernommen.")
                st.text(st.session_state["try_info"])
                for h in hints: st.info(h)
                if st.session_state["auto_calc"]:
                    st.session_state["trigger"] = True

# ----- 2) Anlagendaten (FORM) -----
with st.form("form_anlage"):
    st.subheader("2) Anlagendaten")
    c1 = st.columns([1.2,1,1,1])
    a_id   = c1[0].text_input("Anlagen‑ID", value="A01")
    a_name = c1[1].text_input("Bezeichnung", value="Zuluft")
    V_nom  = c1[2].number_input("V_nominal je Anlage [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    n_eq   = c1[3].number_input("Anzahl gleicher Anlagen", value=1, min_value=1, max_value=100, step=1)
    c2 = st.columns(3)
    wrg   = c2[0].checkbox("WRG vorhanden", value=True)
    eta_t = c2[1].number_input("η_t WRG (0–1)", value=0.7, min_value=0.0, max_value=1.0, step=0.05)
    vm    = c2[2].selectbox("Ventilator-Modell", ("SFP [kW/(m³/s)]","fan_kW gesamt"), index=1)
    fan_kW = None; SFP = None
    if vm.startswith("SFP"):
        SFP = st.number_input("SFP [kW/(m³/s)]", value=1.8, min_value=0.0, max_value=10.0, step=0.1)
    else:
        fan_kW = st.number_input("fan_kW bei V_nominal (gesamt) [kW]", value=5.0, min_value=0.0, max_value=1000.0, step=0.1)
    ok2 = st.form_submit_button("Anlagendaten übernehmen")
    if ok2:
        st.session_state["anlage"] = Anlage(a_id, a_name, float(V_nom), int(n_eq), bool(wrg), float(eta_t),
                                            float(fan_kW) if fan_kW is not None else None,
                                            float(SFP) if SFP is not None else None,
                                            st.session_state["wochenplan"] or [])
        st.success("Anlagendaten übernommen.")

# ----- 3) Standardwerte & Zeiten (KEIN FORM – Buttons erlaubt) -----
st.subheader("3) Standardwerte & Betriebs-/Absenkzeiten")
c = st.columns(4)
Tn = c[0].number_input("Soll‑Zuluft NORMAL [°C]", value=20.0, step=0.5, key="Tn")
Ta = c[1].number_input("Soll‑Zuluft ABSENK [°C]", value=17.0, step=0.5, key="Ta")
Vn = c[2].number_input("Volumenstrom NORMAL [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0, key="Vn")
Va = c[3].number_input("Volumenstrom ABSENK [m³/h] (0 = wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0, key="Va")
defs = Defaults(float(Tn), float(Ta), float(Vn), (None if Va==0.0 else float(Va)))

if st.session_state["wochenplan"] is None:
    st.session_state["wochenplan"] = leerer_wochenplan(defs)
wp: List[Tagesplan] = st.session_state["wochenplan"]

st.caption("Pro Fenster: **Standard** (aus Abschnitt 3), **Eigener Wert**, oder **Anlage AUS (0 m³/h)**. Über‑Mitternacht wird korrekt geteilt.")

def render_fenster(tag_index: int, liste: List[Zeitfenster], modus_label: str):
    st.markdown(modus_label)
    add_key = f"add_{modus_label}_{tag_index}"
    if st.button(f"+ Fenster {modus_label} {WOCHENTAGE[tag_index]}", key=add_key):
        if "Normal" in modus_label:
            liste.append(Zeitfenster("08:00","16:00", True, defs.T_normal_C, None))
        else:
            liste.append(Zeitfenster("17:00","06:30", True, defs.T_absenk_C, defs.V_absenk_m3h))
        st.experimental_rerun()

    delete_index = None
    for i, f in enumerate(liste):
        cols = st.columns([1,1,0.8,1,1,1,0.6])  # Start | Ende | aktiv | T_soll | Auswahl | (Wert) | Löschen
        f.start = cols[0].text_input("Start", key=f"{modus_label}_start_{tag_index}_{i}", value=f.start)
        f.ende  = cols[1].text_input("Ende",  key=f"{modus_label}_ende_{tag_index}_{i}", value=f.ende)
        f.aktiv = cols[2].checkbox("aktiv",  key=f"{modus_label}_aktiv_{tag_index}_{i}", value=f.aktiv)
        f.T_soll_C = cols[3].number_input("T_soll [°C]", key=f"{modus_label}_Tsoll_{tag_index}_{i}", value=float(f.T_soll_C), step=0.5)

        # Volumenstrom-Optionen
        if f.V_m3h is None:
            default_mode = "Standard (Abschnitt 3)"
        elif f.V_m3h == 0.0:
            default_mode = "Anlage AUS (0 m³/h)"
        else:
            default_mode = "Eigener Wert"

        mode = cols[4].selectbox("Volumenstrom",
                                 ["Standard (Abschnitt 3)", "Eigener Wert", "Anlage AUS (0 m³/h)"],
                                 index=["Standard (Abschnitt 3)", "Eigener Wert", "Anlage AUS (0 m³/h)"].index(default_mode),
                                 key=f"{modus_label}_Vmode_{tag_index}_{i}")

        if mode == "Eigener Wert":
            v_default = 0.0 if (f.V_m3h is None or f.V_m3h == 0.0) else float(f.V_m3h)
            v_input = cols[5].number_input("V [m³/h]", key=f"{modus_label}_Vval_{tag_index}_{i}",
                                           value=v_default, min_value=0.0, max_value=500000.0, step=100.0)
            f.V_m3h = float(v_input)
        elif mode == "Anlage AUS (0 m³/h)":
            cols[5].markdown("—")
            f.V_m3h = 0.0
        else:
            cols[5].markdown("—")
            f.V_m3h = None

        if cols[6].button("–", key=f"{modus_label}_del_{tag_index}_{i}"): delete_index = i

    if delete_index is not None:
        liste.pop(delete_index); st.experimental_rerun()

for d in range(7):
    st.markdown(f"**{WOCHENTAGE[d]}**")
    day = wp[d]
    render_fenster(d, day.normal, "Normalbetrieb")
    render_fenster(d, day.absenk, "Absenkbetrieb")

# ----- 4) Berechnen -----
def _rechnen():
    df = st.session_state["try_df"]; anl = st.session_state["anlage"]
    if df is None or df.empty: st.error("Bitte zuerst TRY übernehmen."); return
    if anl is None: st.error("Bitte zuerst Anlagendaten übernehmen."); return
    anl = Anlage(anl.id, anl.name, anl.V_nominal_m3h, anl.anzahl, anl.wrg, anl.eta_t, anl.fan_kW, anl.SFP_kW_per_m3s, st.session_state["wochenplan"])
    mon, jahr, prot, prot_full = berechne_detail(df, anl, defs)
    mon_u, jahr_u = berechne_ueberschlag(df, anl, defs)

    st.session_state["mon_df"], st.session_state["jahr_df"] = mon, jahr
    st.session_state["prot_df"], st.session_state["prot_full_df"] = prot, prot_full
    st.session_state["mon_ue_df"], st.session_state["jahr_ue_df"] = mon_u, jahr_u
    st.success("Berechnung abgeschlossen.")

cols_run = st.columns([1,3])
if cols_run[0].button("Berechnen", type="primary"): _rechnen()
if st.session_state["trigger"]:
    _rechnen(); st.session_state["trigger"] = False

# ----- Ergebnisse sichtbar + Downloads darunter -----
mon = st.session_state["mon_df"]; jahr = st.session_state["jahr_df"]
prot = st.session_state["prot_df"]; prot_full = st.session_state["prot_full_df"]
mon_u = st.session_state["mon_ue_df"]; jahr_u = st.session_state["jahr_ue_df"]

if mon is not None and jahr is not None:
    # Anzeige runden und Spalten DE benennen
    m_show = mon.copy(); j_show = jahr.copy()
    for c in ("kWh_th","kWh_el"):
        if c in m_show: m_show[c] = m_show[c].round(0)
        if c in j_show: j_show[c] = j_show[c].round(0)
    if "Betriebsstunden_Vent" in m_show: m_show["Betriebsstunden_Vent"] = m_show["Betriebsstunden_Vent"].round(1)
    if "Betriebsstunden_Vent" in j_show: j_show["Betriebsstunden_Vent"] = j_show["Betriebsstunden_Vent"].round(1)

    m_de = m_show.rename(columns={
        "year":"Jahr", "month":"Monat",
        "kWh_th":"Wärme [kWh]", "kWh_el":"Strom Vent. [kWh]",
        "Betriebsstunden_Vent":"Betriebsstd. Vent."
    })
    j_de = j_show.rename(columns={
        "year":"Jahr",
        "kWh_th":"Wärme [kWh]", "kWh_el":"Strom Vent. [kWh]",
        "Betriebsstunden_Vent":"Betriebsstd. Vent."
    })

    # Summenzeilen
    m_de = add_sum_row(m_de, label_col="Monat", label="Summe")
    j_de = add_sum_row(j_de)

    st.subheader("Ergebnisse – Jahreswerte (Detail)")
    st.dataframe(j_de, use_container_width=True)

    st.subheader("Ergebnisse – Monate (Detail)")
    st.dataframe(m_de, use_container_width=True)

    # Überschlag & Vergleich
    if mon_u is not None and not mon_u.empty:
        mu = mon_u.copy(); ju = jahr_u.copy()
        mu[["kWh_th","kWh_el"]] = mu[["kWh_th","kWh_el"]].round(0)
        ju[["kWh_th","kWh_el"]] = ju[["kWh_th","kWh_el"]].round(0)

        mu_de = mu.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        ju_de = ju.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        mu_de = add_sum_row(mu_de, label_col="Monat", label="Summe")
        ju_de = add_sum_row(ju_de)

        st.subheader("Kontrollrechner – Monate (Überschlag)")
        st.dataframe(mu_de, use_container_width=True)

        # Abweichung in %
        cmp = m_show.merge(mu, on=["year","month"], suffixes=("_det","_ue"))
        cmp["Abw. Wärme [%]"] = (cmp["kWh_th_det"]-cmp["kWh_th_ue"]) / cmp["kWh_th_det"].replace(0, pd.NA) * 100
        cmp["Abw. Strom [%]"] = (cmp["kWh_el_det"]-cmp["kWh_el_ue"]) / cmp["kWh_el_det"].replace(0, pd.NA) * 100
        cmp_out = cmp[["year","month","Abw. Wärme [%]","Abw. Strom [%]"]].round(1).rename(columns={"year":"Jahr","month":"Monat"})
        st.caption("Abweichung Überschlag vs. Detail (positiv = Überschlag kleiner).")
        st.dataframe(cmp_out, use_container_width=True)

    # Rechen‑Protokoll
    st.subheader("Rechen‑Protokoll (Ausschnitt)")
    if prot is not None and not prot.empty:
        st.dataframe(prot.head(200), use_container_width=True)
        # Vollprotokoll-Download vorbereiten
        if prot_full is not None and not prot_full.empty:
            csv_full = prot_full.to_csv(index=False).encode("utf-8")
        else:
            csv_full = b""
    else:
        st.info("Kein Protokoll verfügbar.")
        csv_full = b""

    # Plausibilitäts-Check (grobe Heuristik)
    if not j_show.empty:
        th = float(j_show["kWh_th"].sum()); el = float(j_show["kWh_el"].sum())
        hints=[]
        if th <= 0: hints.append("Wärmebedarf = 0 kWh → Prüfe Soll‑Temperaturen und Zeitfenster.")
        if th > 10_000_000: hints.append("Sehr hoher Wärmebedarf (>10 GWh) → Prüfe Volumenstrom/Zeiten/Formel.")
        if el > th*0.5: hints.append("Ventilatorstrom sehr hoch im Verhältnis zur Wärme → SFP/Fan‑Werte prüfen.")
        for h in hints: st.warning(h)

    # Downloads
    st.subheader("Downloads")
    c1, c2, c3 = st.columns(3)
    c1.download_button("Excel (Detail + Überschlag + Protokoll)", xlsx_export(m_show, j_show, prot if prot is not None else pd.DataFrame(), mon_u if mon_u is not None else pd.DataFrame(), jahr_u if jahr_u is not None else pd.DataFrame()),
                       file_name="Heizenergie_Auswertung.xlsx")
    c2.download_button("CSV – Monate (Detail, DE)", m_de.to_csv(index=False).encode("utf-8"),
                       file_name="Heizenergie_Monate_DE.csv", mime="text/csv")
    if REPORTLAB_OK:
        anl_obj = st.session_state["anlage"]
        if anl_obj is not None:
            c3.download_button("PDF (ISO 50001 Kurzbericht)",
                               pdf_export(st.session_state.get("try_info",""), defs, anl_obj, m_show, j_show),
                               file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
    else:
        st.info("PDF‑Export nicht verfügbar (ReportLab nicht installiert).")

    if csv_full:
        st.download_button("CSV – Rechen‑Protokoll (vollständig)", csv_full,
                           file_name="Heizenergie_Protokoll_Stunden.csv", mime="text/csv")
