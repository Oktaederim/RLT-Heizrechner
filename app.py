# app.py
# Heizenergieabschätzung Lüftungsanlagen (TRY → Monats/Jahreswerte), ISO 50001-tauglich
# Robust: Session-State, sicherer CSV-Import, Validierungen, XLSX/PDF-Exporte

from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st

# Optional: PDF-Erzeugung (Platypus)
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
class Window:
    start: str
    end: str
    mode: str   # "Normal" | "Absenk"
    T_override_C: Optional[float] = None
    V_override_m3h: Optional[float] = None

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
    V_nominal_m3h: float
    units_count: int
    has_HRV: bool
    eta_t: float
    fan_power_kW: Optional[float]
    SFP_kW_per_m3s: Optional[float]
    plan: List[DayPlan]
    notes: str = ""


# =========================
# Wochenplan
# =========================
DAYS = ["Mo","Di","Mi","Do","Fr","Sa","So"]

def empty_week_plan() -> List[DayPlan]:
    week: List[DayPlan] = []
    for d in range(7):
        wins: List[Window] = []
        if d < 5:
            wins.append(Window("06:30","17:00","Normal"))
            wins.append(Window("17:00","06:30","Absenk"))
        # Sa/So standardmäßig aus (keine Fenster)
        week.append(DayPlan(day=d, windows=wins))
    return week

def normalize_week(plan: List[DayPlan]) -> List[List[Tuple[int,int,Window]]]:
    norm: List[List[Tuple[int,int,Window]]] = [[] for _ in range(7)]
    for d in plan:
        for w in d.windows:
            s = minutes(w.start); e = minutes(w.end)
            if s == e:  # 0-Länge ignorieren
                continue
            if e > s:
                norm[d.day].append((s,e,w))
            else:
                norm[d.day].append((s,1440,w))
                norm[(d.day+1)%7].append((0,e,w))
    for i in range(7): norm[i].sort(key=lambda t: t[0])
    return norm


# =========================
# Kernberechnung
# =========================
def compute(try_df: pd.DataFrame, plant: Plant, defaults: Defaults):
    V_nom_total = float(plant.V_nominal_m3h) * int(plant.units_count)
    V_nom_total = float(clamp(V_nom_total, 0.0, 500000.0))
    norm = normalize_week(plant.plan)

    rec = []
    for i in range(len(try_df)):
        dt0 = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])
        day_js = dt0.weekday()
        m0 = dt0.hour * 60 + dt0.minute
        m1 = m0 + 60

        for (s,e,w) in norm[day_js]:
            ol = overlap_minutes(m0,m1,s,e)
            if ol <= 0: continue
            frac_h = ol/60.0

            T_soll = float(w.T_override_C) if w.T_override_C is not None else (defaults.T_normal_C if w.mode=="Normal" else defaults.T_absenk_C)
            if w.V_override_m3h is not None:
                V_m3h = max(0.0, float(w.V_override_m3h))
            else:
                V_m3h = V_nom_total if w.mode=="Normal" else (V_nom_total if defaults.V_absenk_m3h is None else float(defaults.V_absenk_m3h))

            dT = max(0.0, T_soll - Tout)
            dT_eff = (1.0 - clamp(float(plant.eta_t),0.0,1.0))*dT if plant.has_HRV else dT
            Qdot_kW = 0.34 * V_m3h * dT_eff
            Q_kWh = Qdot_kW * frac_h

            P_fan_kW = 0.0
            if V_m3h > 0.0:
                if plant.SFP_kW_per_m3s is not None:
                    P_fan_kW = float(plant.SFP_kW_per_m3s) * (V_m3h/3600.0)
                elif plant.fan_power_kW is not None:
                    ref = max(1.0, V_nom_total)
                    factor = clamp(V_m3h/ref, 0.0, 1.0)  # v1 linear
                    P_fan_kW = float(plant.fan_power_kW) * factor

            E_kWh = P_fan_kW * frac_h

            rec.append({
                "datetime": dt0,
                "year": dt0.year,
                "month": dt0.month,
                "kWh_th": Q_kWh,
                "kWh_el": E_kWh,
                "fan_hours": frac_h if V_m3h > 0 else 0.0,
            })

    df = pd.DataFrame.from_records(rec)
    monthly = df.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","fan_hours"]].sum()
    yearly  = df.groupby(["year"], as_index=False)[["kWh_th","kWh_el","fan_hours"]].sum()
    return monthly, yearly


# =========================
# Excel-Exporte
# =========================
def build_excel(plant: Plant, monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        m = monthly.copy(); y = yearly.copy()
        for col in ("kWh_th","kWh_el"):
            if col in m: m[col] = m[col].round(0)
            if col in y: y[col] = y[col].round(0)
        if "fan_hours" in m: m["fan_hours"] = m["fan_hours"].round(1)
        if "fan_hours" in y: y["fan_hours"] = y["fan_hours"].round(1)

        m.to_excel(writer, index=False, sheet_name="Monate")
        y.to_excel(writer, index=False, sheet_name="Jahreswerte")

        wb = writer.book
        fmt_int  = wb.add_format({"num_format": "#,##0"})
        fmt_1dec = wb.add_format({"num_format": "#,##0.0"})
        fmt_hdr  = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})

        for sh, df in (("Monate", m), ("Jahreswerte", y)):
            ws = writer.sheets[sh]
            ws.set_row(0, 20, fmt_hdr)
            headers = list(df.columns)
            for i, col in enumerate(headers):
                if col in ("kWh_th","kWh_el"):
                    ws.set_column(i, i, 14, fmt_int)
                elif col == "fan_hours":
                    ws.set_column(i, i, 14, fmt_1dec)
                else:
                    ws.set_column(i, i, 12)
            ws.autofilter(0, 0, len(df), len(headers)-1)
            ws.freeze_panes(1, 0)
    return out.getvalue()

def build_excel_months(monthly: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        m = monthly.copy()
        if not m.empty:
            m["kWh_th"] = m["kWh_th"].round(0)
            m["kWh_el"] = m["kWh_el"].round(0)
            m["fan_hours"] = m["fan_hours"].round(1)
        m.to_excel(writer, index=False, sheet_name="Monate")
    return out.getvalue()


# =========================
# PDF-Export (Platypus)
# =========================
def build_pdf(try_info: str, defaults: Defaults, plant: Plant, monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab ist nicht installiert.")
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]; H1.fontSize = 14
    H2 = styles["Heading2"]; H2.fontSize = 12
    N  = styles["BodyText"]; N.leading = 14

    story = []
    story += [Paragraph("ISO 50001 – Heizenergieabschätzung Lüftungsanlagen (v1)", H1), Spacer(1, 6)]
    story += [Paragraph(f"Erzeugt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1, 8)]

    story += [Paragraph("1. Zweck & Systemgrenzen", H2),
              Paragraph("Lüftungsanlagen ohne direkte Wärmemengen-/Stromzähler. Abschätzung auf Basis TRY-Außenluft (stündlich) und Anlagenparametern (Volumenstrom, Soll-Zuluft, Betriebsfenster, WRG, SFP/fan_kW). Systemgrenze: Heizbedarf der Zuluft und Ventilatorarbeit.", N),
              Spacer(1, 6)]
    story += [Paragraph("2. Datenquelle (TRY)", H2), Paragraph(try_info or "TRY-CSV eingelesen.", N), Spacer(1, 6)]
    story += [Paragraph("3. Annahmen & Parameter (Defaults)", H2),
              Paragraph(f"T_normal: {defaults.T_normal_C} °C, T_absenk: {defaults.T_absenk_C} °C, V_absenk: {defaults.V_absenk_m3h if defaults.V_absenk_m3h is not None else '=V_normal'} m³/h.", N),
              Spacer(1, 6)]
    story += [Paragraph("4. Methodik (v1)", H2),
              Paragraph("Stündlich: ΔT = max(0, T_soll − T_out). Mit WRG: ΔT_eff = (1 − η_t)·ΔT. Heizleistung: Q̇ = 0,34·V·ΔT_eff (kW), pro Stunde Q_kWh = Q̇. Ventilator: SFP·V/3600 oder fan_kW·(V/V_normal). Überlappung minütlich berücksichtigt (z. B. 06:30–17:00). Aggregation je Monat/Jahr.", N),
              Spacer(1, 6)]

    if not yearly.empty:
        ysum = yearly.iloc[0].copy()
        ysum["kWh_th"] = round(ysum["kWh_th"], 0)
        ysum["kWh_el"] = round(ysum["kWh_el"], 0)
        ysum["fan_hours"] = round(ysum["fan_hours"], 1)
        story += [Paragraph("5. Ergebnisse – Jahreswerte", H2)]
        data = [["Jahr","kWh_th","kWh_el","Betriebsstunden"],
                [int(ysum["year"]), int(ysum["kWh_th"]), int(ysum["kWh_el"]), f"{ysum['fan_hours']:.1f}"]]
        tbl = Table(data, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(1,1),(-1,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ]))
        story += [tbl, Spacer(1, 8)]

    if not monthly.empty:
        m = monthly.copy()
        m["kWh_th"] = m["kWh_th"].round(0)
        m["kWh_el"] = m["kWh_el"].round(0)
        m["fan_hours"] = m["fan_hours"].round(1)
        data = [["Jahr","Monat","kWh_th","kWh_el","Betriebsstunden"], *m.astype({"year":int,"month":int}).values.tolist()]
        tbl = Table(data, hAlign="LEFT", repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(2,1),(-1,-1),"RIGHT"),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ]))
        story += [Paragraph("6. Ergebnisse – Monate", H2), tbl]

    story += [Spacer(1, 8), Paragraph("7. Limitierungen & Hinweise", H2),
              Paragraph("Vereinfachtes Modell: konstante Luftdichte/c_p (Faktor 0,34), keine Feuchte-/Bypass‑Logik, Ventilator linear mit Volumenstrom. Genauigkeit abhängig von Parametrierung und Lastprofil.", N)]

    doc.build(story); buf.seek(0); return buf.read()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Heizenergie – ISO 50001", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session: Daten & Ergebnisse
for k, v in [("try_df", None), ("try_info",""), ("monthly_df", None), ("yearly_df", None)]:
    if k not in st.session_state: st.session_state[k] = v

st.markdown("**Ziel:** Monats-/Jahreswerte (kWh_th, kWh_el) aus TRY‑CSV und Anlagenparametern. CSV: `datetime` (ISO), `T_out_C` (°C).")

# ---------- 1) TRY-CSV laden (robust) ----------
with st.expander("1) TRY‑CSV laden", expanded=True):
    f = st.file_uploader("TRY‑CSV auswählen (stündlich, Datum/Zeit + Außentemperatur)", type=["csv"])

    def _find_column(df: pd.DataFrame, aliases: list[str]) -> Optional[str]:
        low = {c.lower().strip(): c for c in df.columns}
        for a in aliases:
            if a in low: return low[a]
        for c in df.columns:
            cl = c.lower().strip()
            if any(a in cl for a in aliases):
                return c
        return None

    if f is not None:
        raw = pd.read_csv(f)

        dt_aliases = ["datetime","date_time","date/time","date","timestamp","zeit","zeitstempel","datestamp","datetime_local","datetime_utc"]
        t_aliases  = ["t_out_c","t_out","tout","temp_out","temperature_out","aussen","außen","ta","t2m","t_out(°c)"]

        dt_col = _find_column(raw, dt_aliases)
        t_col  = _find_column(raw, t_aliases)

        c1, c2 = st.columns(2)
        dt_col = c1.selectbox("Datums-/Zeitspalte", options=["— bitte wählen —"]+raw.columns.tolist(),
                              index=(raw.columns.tolist().index(dt_col)+1 if dt_col in raw.columns else 0))
        t_col  = c2.selectbox("Außentemperatur-Spalte", options=["— bitte wählen —"]+raw.columns.tolist(),
                              index=(raw.columns.tolist().index(t_col)+1 if t_col in raw.columns else 0))

        if dt_col == "— bitte wählen —" or t_col == "— bitte wählen —":
            st.info("Bitte beide Spalten auswählen.")
        else:
            df = raw[[dt_col, t_col]].copy()

            # Temp zu float (Komma zulassen)
            if df[t_col].dtype == object:
                df[t_col] = (df[t_col].astype(str)
                                       .str.replace(",", ".", regex=False)
                                       .str.replace("°C", "", regex=False)
                                       .str.strip())
            df[t_col] = pd.to_numeric(df[t_col], errors="coerce")

            # Datum/Zeit
            df["datetime"] = df[dt_col].astype(str).apply(parse_datetime)
            df = df.rename(columns={t_col: "T_out_C"})
            df = df[["datetime","T_out_C"]].dropna().sort_values("datetime").reset_index(drop=True)

            # --- Validierungen ---
            problems = []

            # Dubletten
            dup = df["datetime"].duplicated().sum()
            if dup > 0:
                problems.append(f"{dup} doppelte Zeitstempel entfernt")
                df = df[~df["datetime"].duplicated(keep="first")]

            # Lücken prüfen (grobe Heuristik)
            if len(df) >= 2:
                diffs = df["datetime"].diff().dropna()
                # Anteil 1h-Schritte
                one_h = (diffs == pd.Timedelta(hours=1)).mean()
                if one_h < 0.95:
                    problems.append("Unregelmäßiges Raster (weniger als 95 % 1‑Stunden‑Schritte).")

            # Jahreslänge (Info)
            years = sorted(df["datetime"].dt.year.unique().tolist())
            tmin, tmax = df["T_out_C"].min(), df["T_out_C"].max()

            st.session_state["try_df"] = df
            st.session_state["try_info"] = f"Datensätze: {len(df)} | Jahre: {', '.join(map(str, years))} | T_out: {tmin}…{tmax} °C"
            st.success("TRY‑CSV eingelesen.")
            st.text(st.session_state["try_info"])
            if problems:
                st.warning(" / ".join(problems))
    else:
        if st.session_state["try_df"] is not None:
            st.info("TRY‑CSV bereits geladen.")
            if st.session_state["try_info"]:
                st.text(st.session_state["try_info"])

# ---------- 2) Defaults ----------
with st.expander("2) Defaults – Normal/Absenk", expanded=True):
    c = st.columns(3)
    T_normal = c[0].number_input("T_normal [°C]", value=20.0, step=0.5)
    T_absenk = c[1].number_input("T_absenk [°C]", value=17.0, step=0.5)
    V_absenk = c[2].number_input("V_absenk [m³/h] (leer = wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0)
    defaults = Defaults(T_normal_C=float(T_normal), T_absenk_C=float(T_absenk), V_absenk_m3h=float(V_absenk))
    st.caption("Hinweis: V_absenk = 0 ⇒ Ventilator AUS im Absenk‑Fenster. None (=wie normal) ist v2‑Option.")

# ---------- 3) Anlage ----------
with st.expander("3) Anlage", expanded=True):
    r1 = st.columns([1.2,1,1,1])
    plant_id   = r1[0].text_input("ID", value="A01")
    plant_name = r1[1].text_input("Name", value="Zuluft – Beispiel")
    V_nominal  = r1[2].number_input("V_nominal [m³/h] (Einzelanlage)", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0)
    units      = r1[3].number_input("Anzahl gleicher Anlagen", value=1, min_value=1, max_value=100, step=1)

    r2 = st.columns(3)
    has_hrv = r2[0].checkbox("WRG vorhanden", value=True)
    eta_t   = r2[1].number_input("η_t (0–1)", value=0.7, min_value=0.0, max_value=1.0, step=0.05)
    vent_model = r2[2].selectbox("Ventilator-Modell", ("SFP","fan_kW"), index=1)

    fan_kW = None; SFP = None
    if vent_model == "SFP":
        SFP = st.number_input("SFP [kW/(m³/s)]", value=1.8, min_value=0.0, max_value=10.0, step=0.1)
    else:
        fan_kW = st.number_input("fan_kW (bei V_nominal gesamt)", value=5.0, min_value=0.0, max_value=1000.0, step=0.1)

    # Wochenplan-Editor (einfach): je Tag Fensterliste
    if "week_plan" not in st.session_state:
        st.session_state["week_plan"] = empty_week_plan()
    wp: List[DayPlan] = st.session_state["week_plan"]

    st.markdown("**Betrieb & Absenkung (minütliche Zeitfenster, Über‑Mitternacht erlaubt)**")
    for d in range(7):
        st.write(f"**{DAYS[d]}**")
        day = wp[d]
        # vorhandene Fenster
        to_del = None
        for w_idx, w in enumerate(list(day.windows)):
            cols = st.columns([1,1,1,1,1,0.5])
            w.start = cols[0].text_input("Start", key=f"start_{d}_{w_idx}", value=w.start)
            w.end   = cols[1].text_input("Ende", key=f"end_{d}_{w_idx}", value=w.end)
            w.mode  = cols[2].selectbox("Modus", ["Normal","Absenk"], key=f"mode_{d}_{w_idx}", index=(0 if w.mode=="Normal" else 1))
            # Konsistenter float-Input
            tval = float(w.T_override_C) if w.T_override_C is not None else 0.0
            vval = float(w.V_override_m3h) if w.V_override_m3h is not None else 0.0
            new_t = cols[3].number_input("T_override [°C]", key=f"To_{d}_{w_idx}", value=tval, step=0.5)
            new_v = cols[4].number_input("V_override [m³/h]", key=f"Vo_{d}_{w_idx}", value=vval, min_value=0.0, max_value=500000.0, step=100.0)
            # Leereingabe als 0 interpretieren → 0 = aus / nur Temp
            day.windows[w_idx].T_override_C = float(new_t) if new_t != 0.0 else None
            day.windows[w_idx].V_override_m3h = float(new_v) if new_v != 0.0 else None
            if cols[5].button("–", key=f"del_{d}_{w_idx}"):
                to_del = w_idx
        if to_del is not None:
            day.windows.pop(to_del)
            st.experimental_rerun()
        if st.button("+ Fenster", key=f"add_{d}"):
            day.windows.append(Window("06:30","17:00","Normal"))
            st.experimental_rerun()

# ---------- 4) Berechnen ----------
if st.button("Berechnen", type="primary"):
    try_df = st.session_state["try_df"]
    if try_df is None or try_df.empty:
        st.error("Bitte zuerst eine gültige TRY‑CSV laden.")
    else:
        plant = Plant(
            id=plant_id, name=plant_name,
            V_nominal_m3h=float(V_nominal), units_count=int(units),
            has_HRV=bool(has_hrv), eta_t=float(eta_t),
            fan_power_kW=float(fan_kW) if fan_kW is not None else None,
            SFP_kW_per_m3s=float(SFP) if SFP is not None else None,
            plan=st.session_state["week_plan"],
        )
        m, y = compute(try_df, plant, defaults)
        st.session_state["monthly_df"] = m
        st.session_state["yearly_df"]  = y
        st.success("Berechnung abgeschlossen.")

# ---------- 5) Ergebnisse & Exporte ----------
m = st.session_state["monthly_df"]; y = st.session_state["yearly_df"]
if m is not None and y is not None:
    # Anzeige gerundet
    m_view = m.copy(); y_view = y.copy()
    for col in ("kWh_th","kWh_el"):
        if col in m_view: m_view[col] = m_view[col].round(0)
        if col in y_view: y_view[col] = y_view[col].round(0)
    if "fan_hours" in m_view: m_view["fan_hours"] = m_view["fan_hours"].round(1)
    if "fan_hours" in y_view: y_view["fan_hours"] = y_view["fan_hours"].round(1)

    st.subheader("Ergebnisse – Monate")
    st.dataframe(m_view, use_container_width=True)
    st.subheader("Jahressumme")
    st.dataframe(y_view, use_container_width=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.download_button("Excel (gesamt) .xlsx", build_excel(
        Plant(plant_id, plant_name, float(V_nominal), int(units), bool(has_hrv), float(eta_t),
              float(fan_kW) if fan_kW is not None else None,
              float(SFP) if SFP is not None else None,
              st.session_state["week_plan"]),
        m_view, y_view
    ), file_name="Heizenergie_Monate_und_Jahr.xlsx")

    c2.download_button("Monate nur .xlsx", build_excel_months(m_view),
                       file_name="Heizenergie_Monate.xlsx")

    c3.download_button("Monate .csv", m_view.to_csv(index=False).encode("utf-8"),
                       file_name="Heizenergie_Monate.csv", mime="text/csv")

    if REPORTLAB_OK:
        c4.download_button("PDF (ISO 50001)", build_pdf(
            st.session_state.get("try_info",""), defaults,
            Plant(plant_id, plant_name, float(V_nominal), int(units), bool(has_hrv), float(eta_t),
                  float(fan_kW) if fan_kW is not None else None,
                  float(SFP) if SFP is not None else None,
                  st.session_state["week_plan"]),
            m_view, y_view
        ), file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
    else:
        st.info("PDF‑Export: Paket 'reportlab' nicht installiert (requirements.txt).")
