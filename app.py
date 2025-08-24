# app.py
# Streamlit-App: Heizenergieabschätzung Lüftungsanlagen (ohne Zähler)
# Eingabe: TRY-CSV (datetime, T_out_C). Ausgabe: Monats-/Jahressummen (kWh_th, kWh_el), CSV/PDF (ISO 50001-tauglich)
# Autor: v1

import io
import math
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict

import pandas as pd
import streamlit as st

# Optional: PDF-Erzeugung
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.units import mm
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

# ------------------------------
# Hilfsfunktionen
# ------------------------------

def parse_datetime(s: str) -> Optional[datetime]:
    s = s.strip().replace(" ", "T")
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def minutes(hhmm: str) -> int:
    h, m = hhmm.split(":")
    return int(h) * 60 + int(m)


def overlap_minutes(a0: int, a1: int, b0: int, b1: int) -> int:
    # alle in [0,1440]
    return max(0, min(a1, b1) - max(a0, b0))


def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))

# ------------------------------
# Datenklassen
# ------------------------------

@dataclass
class Window:
    start: str  # HH:MM
    end: str    # HH:MM
    mode: str   # "Normal" | "Absenk"
    T_override_C: Optional[float] = None
    V_override_m3h: Optional[float] = None  # 0 => aus


@dataclass
class DayPlan:
    day: int  # 0=Mo..6=So
    windows: List[Window]


@dataclass
class Defaults:
    T_normal_C: float = 20.0
    T_absenk_C: float = 17.0
    V_absenk_m3h: Optional[float] = 2000.0  # None => wie V_normal


@dataclass
class Plant:
    id: str
    name: str
    V_nominal_m3h: float  # Einzelanlage
    units_count: int
    has_HRV: bool
    eta_t: float  # 0..1
    fan_power_kW: Optional[float]  # alternativ zu SFP
    SFP_kW_per_m3s: Optional[float]
    plan: List[DayPlan]
    notes: str = ""

# ------------------------------
# Plan-Editor-Helfer
# ------------------------------

DAYS = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


def empty_week_plan(normal_start: str = "06:30", normal_end: str = "17:00",
                     abs_start: str = "17:00", abs_end: str = "06:30",
                     weekend_off: bool = True) -> List[DayPlan]:
    week: List[DayPlan] = []
    for d in range(7):
        wins: List[Window] = []
        if d < 5:  # Werktage
            wins.append(Window(normal_start, normal_end, "Normal"))
            wins.append(Window(abs_start, abs_end, "Absenk"))
        else:
            if not weekend_off:
                wins.append(Window(normal_start, normal_end, "Normal"))
                wins.append(Window(abs_start, abs_end, "Absenk"))
            else:
                if d == 5:
                    wins.append(Window("08:00", "14:00", "Absenk", V_override_m3h=3000.0))
                # So: aus (keine Fenster)
        week.append(DayPlan(day=d, windows=wins))
    return week


# Normalisiert Fenster auf per-Tag Intervalle [start,end) in Minuten
# gesplittet bei Über-Mitternacht

def normalize_week(plan: List[DayPlan]) -> List[List[Tuple[int, int, Window]]]:
    norm: List[List[Tuple[int, int, Window]]] = [[] for _ in range(7)]
    for d in plan:
        for w in d.windows:
            s = minutes(w.start)
            e = minutes(w.end)
            if s == e:
                continue
            if e > s:
                norm[d.day].append((s, e, w))
            else:
                norm[d.day].append((s, 1440, w))
                norm[(d.day + 1) % 7].append((0, e, w))
    for i in range(7):
        norm[i].sort(key=lambda t: t[0])
    return norm

# ------------------------------
# Kernberechnung
# ------------------------------

def compute(try_df: pd.DataFrame, plant: Plant, defaults: Defaults) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # Erwartet try_df mit Spalten: datetime (datetime64), T_out_C (float)
    # Liefert: monthly_df, yearly_df
    V_nom_total = plant.V_nominal_m3h * plant.units_count
    V_nom_total = float(clamp(V_nom_total, 0, 500000.0))

    norm = normalize_week(plant.plan)

    records: List[Dict] = []

    for i in range(len(try_df)):
        dt0: datetime = try_df.iloc[i]["datetime"]
        Tout: float = float(try_df.iloc[i]["T_out_C"])
        dt1 = dt0 + timedelta(hours=1)
        day_js = dt0.weekday()  # 0=Mo..6=So
        m0 = dt0.hour * 60 + dt0.minute
        m1 = m0 + 60

        for (s, e, w) in norm[day_js]:
            ol = overlap_minutes(m0, m1, s, e)
            if ol <= 0:
                continue
            frac_h = ol / 60.0
            # Setpoints
            T_soll = w.T_override_C if (w.T_override_C is not None) else (defaults.T_normal_C if w.mode == "Normal" else defaults.T_absenk_C)
            if w.V_override_m3h is not None:
                V_m3h = max(0.0, float(w.V_override_m3h))
            else:
                if w.mode == "Normal":
                    V_m3h = V_nom_total
                else:
                    V_m3h = V_nom_total if defaults.V_absenk_m3h is None else float(defaults.V_absenk_m3h)
                    V_m3h = max(0.0, V_m3h)

            dT = max(0.0, T_soll - Tout)
            dT_eff = (1.0 - clamp(plant.eta_t, 0.0, 1.0)) * dT if plant.has_HRV else dT
            Qdot_kW = 0.34 * V_m3h * dT_eff
            Q_kWh = Qdot_kW * frac_h

            # Ventilator
            P_fan_kW = 0.0
            if V_m3h > 0.0:
                if plant.SFP_kW_per_m3s is not None:
                    P_fan_kW = float(plant.SFP_kW_per_m3s) * (V_m3h / 3600.0)
                elif plant.fan_power_kW is not None:
                    ref = max(1.0, V_nom_total)
                    factor = clamp(V_m3h / ref, 0.0, 1.0)
                    P_fan_kW = float(plant.fan_power_kW) * factor
            E_kWh = P_fan_kW * frac_h

            records.append({
                "datetime": dt0,
                "year": dt0.year,
                "month": dt0.month,
                "kWh_th": Q_kWh,
                "kWh_el": E_kWh,
                "fan_hours": frac_h if V_m3h > 0 else 0.0,
            })

    if not records:
        monthly = pd.DataFrame(columns=["year", "month", "kWh_th", "kWh_el", "fan_hours"])  # empty
        yearly = pd.DataFrame(columns=["year", "kWh_th", "kWh_el", "fan_hours"])  # empty
        return monthly, yearly

    df = pd.DataFrame.from_records(records)
    monthly = df.groupby(["year", "month"], as_index=False)[["kWh_th", "kWh_el", "fan_hours"]].sum()
    yearly = df.groupby(["year"], as_index=False)[["kWh_th", "kWh_el", "fan_hours"]].sum()
    return monthly, yearly

# ------------------------------
# Excel-Export
# ------------------------------
from io import BytesIO

def build_excel(plant: Plant, monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    """Erzeugt eine formatierte XLSX mit getrennten Blättern und sauberen Zahlformaten."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Kopien mit gerundeten Werten (Lesbarkeit)
        m = monthly.copy()
        y = yearly.copy()
        for col in ("kWh_th", "kWh_el"):
            if col in m: m[col] = m[col].round(0)
            if col in y: y[col] = y[col].round(0)
        if "fan_hours" in m: m["fan_hours"] = m["fan_hours"].round(1)
        if "fan_hours" in y: y["fan_hours"] = y["fan_hours"].round(1)

        m.to_excel(writer, index=False, sheet_name="Monate")
        y.to_excel(writer, index=False, sheet_name="Jahreswerte")

        wb  = writer.book
        fmt_int  = wb.add_format({"num_format": "#,##0"})
        fmt_1dec = wb.add_format({"num_format": "#,##0.0"})
        fmt_hdr  = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})

        for sh in ("Monate", "Jahreswerte"):
            ws = writer.sheets[sh]
            # Kopfzeile formatieren
            ws.set_row(0, 20, fmt_hdr)
            # Spaltenbreiten & Formate setzen
            headers = [c for c in (m.columns if sh=="Monate" else y.columns)]
            for i, col in enumerate(headers):
                if col in ("kWh_th", "kWh_el"):
                    ws.set_column(i, i, 14, fmt_int)
                elif col == "fan_hours":
                    ws.set_column(i, i, 14, fmt_1dec)
                else:
                    ws.set_column(i, i, 12)
            # Auto-Filter an
            ws.autofilter(0, 0, (len(m) if sh=="Monate" else len(y)), len(headers)-1)
            # Kopf fixieren
            ws.freeze_panes(1, 0)

    return output.getvalue()

# ------------------------------
# PDF-Export (übersichtlich mit Platypus)
# ------------------------------
def build_pdf(try_info: str, defaults: Defaults, plant: Plant,
              monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab ist nicht installiert.")

    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from datetime import datetime

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=18*mm, bottomMargin=18*mm
    )
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]; H1.fontSize = 14
    H2 = styles["Heading2"]; H2.fontSize = 12
    N  = styles["BodyText"]; N.leading = 14

    story = []
    story += [Paragraph("ISO 50001 – Heizenergieabschätzung Lüftungsanlagen (v1)", H1), Spacer(1, 6)]
    story += [Paragraph(f"Erzeugt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1, 8)]

    story += [
        Paragraph("1. Zweck & Systemgrenzen", H2),
        Paragraph(
            "Lüftungsanlagen ohne direkte Wärmemengen-/Stromzähler. Abschätzung auf Basis TRY-Außenluft (stündlich) "
            "und Anlagenparametern (Volumenstrom, Soll-Zuluft, Betriebsfenster, WRG, SFP/fan_kW). "
            "Systemgrenze: Heizbedarf der Zuluft und Ventilatorarbeit.", N),
        Spacer(1, 6),
    ]

    story += [
        Paragraph("2. Datenquelle (TRY)", H2),
        Paragraph(try_info or "TRY-CSV eingelesen.", N),
        Spacer(1, 6),
    ]

    story += [
        Paragraph("3. Annahmen & Parameter (Defaults)", H2),
        Paragraph(
            f"T_normal: {defaults.T_normal_C} °C, "
            f"T_absenk: {defaults.T_absenk_C} °C, "
            f"V_absenk: {defaults.V_absenk_m3h if defaults.V_absenk_m3h is not None else '=V_normal'} m³/h.", N),
        Spacer(1, 6),
    ]

    story += [
        Paragraph("4. Methodik (v1)", H2),
        Paragraph(
            "Stündlich: ΔT = max(0, T_soll − T_out). Mit WRG: ΔT_eff = (1 − η_t)·ΔT. "
            "Heizleistung: Q̇ = 0,34·V·ΔT_eff (kW), pro Stunde Q_kWh = Q̇. "
            "Ventilator: SFP·V/3600 oder fan_kW·(V/V_normal). "
            "Überlappung minütlich berücksichtigt (z. B. 06:30–17:00). Aggregation je Monat/Jahr.", N),
        Spacer(1, 6),
    ]

    # Jahreswerte (gerundet für Lesbarkeit)
    if not yearly.empty:
        ysum = yearly.iloc[0].copy()
        ysum["kWh_th"] = round(ysum["kWh_th"], 0)
        ysum["kWh_el"] = round(ysum["kWh_el"], 0)
        ysum["fan_hours"] = round(ysum["fan_hours"], 1)
        story += [Paragraph("5. Ergebnisse – Jahreswerte", H2)]
        data = [
            ["Jahr", "kWh_th", "kWh_el", "Betriebsstunden"],
            [int(ysum["year"]), int(ysum["kWh_th"]), int(ysum["kWh_el"]), f"{ysum['fan_hours']:.1f}"]
        ]
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"),
            ("BOTTOMPADDING", (0,0), (-1,0), 6),
            ("TOPPADDING", (0,0), (-1,0), 6),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ]))
        story += [tbl, Spacer(1, 8)]

    # Monatstabelle (gerundet)
    if not monthly.empty:
        m = monthly.copy()
        m["kWh_th"] = m["kWh_th"].round(0)
        m["kWh_el"] = m["kWh_el"].round(0)
        m["fan_hours"] = m["fan_hours"].round(1)
        data = [["Jahr", "Monat", "kWh_th", "kWh_el", "Betriebsstunden"]]
        data += m.astype({"year": int, "month": int}).values.tolist()
        tbl = Table(data, hAlign="LEFT", repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (2,1), (-1,-1), "RIGHT"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ]))
        story += [Paragraph("6. Ergebnisse – Monate", H2), tbl]

    story += [
        Spacer(1, 8),
        Paragraph("7. Limitierungen & Hinweise", H2),
        Paragraph(
            "Vereinfachtes Modell: konstante Luftdichte/c_p (Faktor 0,34), keine Feuchte-/Bypass‑Logik, "
            "Ventilator linear mit Volumenstrom. Genauigkeit abhängig von Parametrierung und Lastprofil.", N),
    ]

    doc.build(story)
    buf.seek(0)
    return buf.read()
# ------------------------------
# Streamlit UI
# ------------------------------

st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001) v1", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001) v1")

st.markdown("""
**Ziel:** Monats-/Jahreswerte (kWh_th, optional kWh_el) aus TRY‑CSV und Anlagenparametern.

**CSV‑Schnittstelle:** Spalten `datetime` (ISO), `T_out_C` (°C), stündlich.
""")

# 1) TRY-Upload
with st.expander("1) TRY‑CSV laden", expanded=True):
    f = st.file_uploader("TRY‑CSV auswählen (Spalten: datetime, T_out_C)", type=["csv"]) 
    try_df: Optional[pd.DataFrame] = None
    try_info = ""
    if f is not None:
        raw = pd.read_csv(f)
        # Spalten erkennen
        dt_col = None
        for c in raw.columns:
            if c.lower() in ("datetime", "date", "timestamp", "zeit", "zeitstempel"):
                dt_col = c
                break
        t_col = None
        for c in raw.columns:
            if c.lower() in ("t_out_c", "tout", "t_out", "aussen", "außen", "ta"):
                t_col = c
                break
        if dt_col is None:
            st.error("Spalte 'datetime' nicht gefunden.")
        elif t_col is None:
            st.error("Spalte 'T_out_C' nicht gefunden.")
        else:
            # Parse
            # Kommas zulassen
            if raw[t_col].dtype == object:
                raw[t_col] = raw[t_col].astype(str).str.replace(",", ".").astype(float)
            # datetime
            raw["datetime"] = raw[dt_col].astype(str).apply(parse_datetime)
            raw = raw.rename(columns={t_col: "T_out_C"})
            ok = raw["datetime"].notna() & raw["T_out_C"].notna()
            parsed = raw.loc[ok, ["datetime", "T_out_C"]].copy()
            parsed.sort_values("datetime", inplace=True)
            try_df = parsed.reset_index(drop=True)
            years = sorted(try_df["datetime"].dt.year.unique().tolist()) if not try_df.empty else []
            try_info = f"Datensätze: {len(try_df)} | Jahre: {', '.join(map(str, years))} | T_out: {try_df['T_out_C'].min()}…{try_df['T_out_C'].max()} °C"
            st.success("TRY‑CSV eingelesen.")
            st.text(try_info)

# 2) Defaults
with st.expander("2) Defaults – Normal/Absenk", expanded=True):
    cols = st.columns(3)
    T_normal = cols[0].number_input("T_normal [°C]", value=20.0, step=0.5)
    T_absenk = cols[1].number_input("T_absenk [°C]", value=17.0, step=0.5)
    V_absenk = cols[2].number_input("V_absenk [m³/h] (leer = wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0)
    defaults = Defaults(T_normal_C=T_normal, T_absenk_C=T_absenk, V_absenk_m3h=V_absenk)
    st.caption("Hinweis: V_absenk = 0 ⇒ Ventilator AUS im Absenk‑Fenster. Leer lassen (=wie normal) in einer späteren Version via Optionen.")

# 3) Anlage(n)
with st.expander("3) Anlage konfigurieren", expanded=True):
    c1, c2, c3, c4 = st.columns([1.1, 1, 1, 1])
    plant_id = c1.text_input("ID", value="A01")
    plant_name = c2.text_input("Name", value="Zuluft – Beispiel")
    V_nominal = c3.number_input("V_nominal [m³/h] (Einzelanlage)", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    units = c4.number_input("Anzahl gleicher Anlagen", value=1, min_value=1, max_value=100, step=1)

    c5, c6, c7 = st.columns(3)
    has_hrv = c5.checkbox("WRG vorhanden", value=True)
    eta_t = c6.number_input("η_t (0–1)", value=0.7, min_value=0.0, max_value=1.0, step=0.05)
    mode = c7.selectbox("Ventilator-Modell", ("SFP", "fan_kW"), index=1)

    fan_kW: Optional[float] = None
    SFP: Optional[float] = None
    if mode == "SFP":
        SFP = st.number_input("SFP [kW/(m³/s)]", value=1.8, min_value=0.0, max_value=10.0, step=0.1)
    else:
        fan_kW = st.number_input("fan_kW (gesamt bei V_nominal)", value=5.0, min_value=0.0, max_value=1000.0, step=0.1)

    st.markdown("**Betrieb & Absenkung (pro Tag, minütlich):**")

    # Einfache Editorform: je Tag eine Tabelle der Fenster
    week_plan: List[DayPlan] = empty_week_plan()

    plan_state_key = "week_plan_state"
    if plan_state_key not in st.session_state:
        st.session_state[plan_state_key] = week_plan
    week_plan = st.session_state[plan_state_key]

    for d in range(7):
        st.write(f"**{DAYS[d]}**")
        day = week_plan[d]
        # vorhandene Fenster darstellen
        for w_idx, w in enumerate(list(day.windows)):
            cc = st.columns([1,1,1,1,1,0.5])
            w.start = cc[0].text_input("Start", key=f"start_{d}_{w_idx}", value=w.start)
            w.end = cc[1].text_input("Ende", key=f"end_{d}_{w_idx}", value=w.end)
            w.mode = cc[2].selectbox("Modus", ["Normal", "Absenk"], key=f"mode_{d}_{w_idx}", index=(0 if w.mode=="Normal" else 1))
            w.T_override_C = cc[3].number_input("T_override [°C]", key=f"To_{d}_{w_idx}", value=float(w.T_override_C) if w.T_override_C is not None else 0.0, step=0.5)
            w.V_override_m3h = cc[4].number_input("V_override [m³/h]", key=f"Vo_{d}_{w_idx}", value=float(w.V_override_m3h) if w.V_override_m3h is not None else 0.0, min_value=0.0, max_value=500000.0, step=100.0)
            if cc[5].button("–", key=f"del_{d}_{w_idx}"):
                day.windows.pop(w_idx)
                st.experimental_rerun()
        if st.button("+ Fenster hinzufügen", key=f"add_{d}"):
            day.windows.append(Window("06:30", "17:00", "Normal"))
            st.experimental_rerun()
        st.divider()

    st.caption("Regeln: Keine Überlappung am selben Tag. Über‑Mitternacht wird automatisch gesplittet (z. B. 17:00–06:30). V_override=0 ⇒ Ventilator aus in diesem Fenster.")

# 4) Berechnen
monthly_df = None
yearly_df = None

if st.button("Berechnen", type="primary"):
    if try_df is None or try_df.empty:
        st.error("Bitte zuerst eine gültige TRY‑CSV laden.")
    else:
        plant = Plant(
            id=plant_id,
            name=plant_name,
            V_nominal_m3h=V_nominal,
            units_count=int(units),
            has_HRV=has_hrv,
            eta_t=float(eta_t),
            fan_power_kW=float(fan_kW) if fan_kW is not None else None,
            SFP_kW_per_m3s=float(SFP) if SFP is not None else None,
            plan=week_plan,
        )
        monthly_df, yearly_df = compute(try_df, plant, defaults)
        st.success("Berechnung abgeschlossen.")

# 5) Ergebnisse & Exporte
if monthly_df is not None and yearly_df is not None:
    # Kopien mit Rundung für Anzeige
    m_round = monthly_df.copy()
    y_round = yearly_df.copy()
    if not m_round.empty:
        m_round["kWh_th"] = m_round["kWh_th"].round(0)
        m_round["kWh_el"] = m_round["kWh_el"].round(0)
        m_round["fan_hours"] = m_round["fan_hours"].round(1)
    if not y_round.empty:
        y_round["kWh_th"] = y_round["kWh_th"].round(0)
        y_round["kWh_el"] = y_round["kWh_el"].round(0)
        y_round["fan_hours"] = y_round["fan_hours"].round(1)

    st.subheader("Ergebnisse – Monate (eine Anlage)")
    st.dataframe(m_round, use_container_width=True)

    st.subheader("Jahressumme (eine Anlage)")
    st.dataframe(y_round, use_container_width=True)

    cexp1, cexp2, cexp3 = st.columns(3)

    # Excel-Export (bevorzugt)
    try:
      xlsx_bytes = build_excel(Plant(plant_id, plant_name, V_nominal, int(units), has_hrv, float(eta_t), fan_kW, SFP, week_plan), m_round, y_round)
      cexp1.download_button("Excel (.xlsx) herunterladen", xlsx_bytes, file_name="Heizenergie_Monate_Jahr.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
      st.warning(f"Excel-Export nicht möglich: {e}")

    # CSV-Export (optional)
    csv_month = m_round.to_csv(index=False).encode("utf-8")
    cexp2.download_button("CSV (Monate) herunterladen", csv_month, file_name="per_anlage_monat.csv", mime="text/csv")

    # PDF-Export (verbessert)
    if REPORTLAB_OK:
        try:
            pdf_bytes = build_pdf(try_info, defaults, Plant(plant_id, plant_name, V_nominal, int(units), has_hrv, float(eta_t), fan_kW, SFP, week_plan), m_round, y_round)
            cexp3.download_button("PDF (ISO 50001) herunterladen", pdf_bytes, file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
        except Exception as e:
            st.warning(f"PDF konnte nicht erzeugt werden: {e}")
    else:
        st.info("PDF‑Export: Paket 'reportlab' ist nicht installiert. In requirements.txt hinzufügen: reportlab")

st.markdown(
        """
        **Rechenweg:**
        - Stunde für Stunde wird geprüft, ob ein Fenster (Normal/Absenk) aktiv ist.
        - Setpoints: `T_soll` und `V_m3h` kommen aus Fenster‑Overrides oder den Defaults.
        - `ΔT = max(0, T_soll − T_out)`. Mit WRG: `ΔT_eff = (1 − η_t) · ΔT`.
        - Heizleistung: `Q̇ = 0,34 · V · ΔT_eff` (kW) ⇒ pro Stunde `Q_kWh = Q̇`.
        - Ventilator: `SFP · (V/3600)` oder `fan_kW · (V/V_nominal)` (lineare Näherung in v1).
        - Aggregation: Summen je Monat/Jahr.

        **Grenzen v1:** konstante Luftdichte/c_p (Faktor 0,34), keine Feuchte-/Bypass‑Logik, kein Würfelgesetz für Ventilator (optional v2), konstante Volumenströme innerhalb eines Fensters.
        """
    )

st.markdown("""
---
**ISO‑50001‑Hinweis:** Der PDF‑Export enthält Zweck/Systemgrenzen, Datenquelle, Annahmen (Parameter), Methodik, Ergebnisse und Limitierungen. Für Audits sollten zusätzlich Versionsstand und Verantwortlichkeiten ergänzt werden.
""")
