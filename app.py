# app.py
# Streamlit-App: Heizenergieabschätzung Lüftungsanlagen (ohne Zähler)
# Eingabe: TRY-CSV (datetime, T_out_C). Ausgabe: Monats-/Jahressummen (kWh_th, kWh_el), CSV/PDF/XLSX
# Autor: v1

import io
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict

import pandas as pd
import streamlit as st

# Optional: PDF-Erzeugung
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
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
    return max(0, min(a1, b1) - max(a0, b0))

def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))

# ------------------------------
# Datenklassen
# ------------------------------
@dataclass
class Window:
    start: str
    end: str
    mode: str   # "Normal" | "Absenk"
    T_override_C: Optional[float] = None
    V_override_m3h: Optional[float] = None

@dataclass
class DayPlan:
    day: int
    windows: List[Window]

@dataclass
class Defaults:
    T_normal_C: float = 20.0
    T_absenk_C: float = 17.0
    V_absenk_m3h: Optional[float] = 2000.0

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

# ------------------------------
# Wochenplan
# ------------------------------
DAYS = ["Mo","Di","Mi","Do","Fr","Sa","So"]

def empty_week_plan():
    week: List[DayPlan] = []
    for d in range(7):
        wins: List[Window] = []
        if d < 5:
            wins.append(Window("06:30","17:00","Normal"))
            wins.append(Window("17:00","06:30","Absenk"))
        else:
            # Wochenende aus (Beispiel)
            pass
        week.append(DayPlan(day=d, windows=wins))
    return week

def normalize_week(plan: List[DayPlan]) -> List[List[Tuple[int,int,Window]]]:
    norm: List[List[Tuple[int,int,Window]]] = [[] for _ in range(7)]
    for d in plan:
        for w in d.windows:
            s = minutes(w.start)
            e = minutes(w.end)
            if s == e: continue
            if e > s:
                norm[d.day].append((s,e,w))
            else:
                norm[d.day].append((s,1440,w))
                norm[(d.day+1)%7].append((0,e,w))
    for i in range(7): norm[i].sort(key=lambda t: t[0])
    return norm

# ------------------------------
# Kernberechnung
# ------------------------------
def compute(try_df: pd.DataFrame, plant: Plant, defaults: Defaults):
    V_nom_total = plant.V_nominal_m3h * plant.units_count
    V_nom_total = float(clamp(V_nom_total, 0, 500000.0))
    norm = normalize_week(plant.plan)

    records = []
    for i in range(len(try_df)):
        dt0 = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])
        dt1 = dt0 + timedelta(hours=1)
        day_js = dt0.weekday()
        m0 = dt0.hour*60 + dt0.minute
        m1 = m0 + 60

        for (s,e,w) in norm[day_js]:
            ol = overlap_minutes(m0,m1,s,e)
            if ol <= 0: continue
            frac_h = ol/60.0

            T_soll = w.T_override_C if w.T_override_C is not None else (defaults.T_normal_C if w.mode=="Normal" else defaults.T_absenk_C)
            if w.V_override_m3h is not None:
                V_m3h = max(0.0, float(w.V_override_m3h))
            else:
                V_m3h = V_nom_total if w.mode=="Normal" else (defaults.V_absenk_m3h or V_nom_total)

            dT = max(0.0, T_soll - Tout)
            dT_eff = (1.0 - clamp(plant.eta_t,0.0,1.0))*dT if plant.has_HRV else dT
            Qdot_kW = 0.34 * V_m3h * dT_eff
            Q_kWh = Qdot_kW * frac_h

            P_fan_kW = 0.0
            if V_m3h>0:
                if plant.SFP_kW_per_m3s is not None:
                    P_fan_kW = float(plant.SFP_kW_per_m3s) * (V_m3h/3600.0)
                elif plant.fan_power_kW is not None:
                    factor = clamp(V_m3h/V_nom_total,0.0,1.0)
                    P_fan_kW = plant.fan_power_kW * factor
            E_kWh = P_fan_kW*frac_h

            records.append({
                "datetime": dt0,
                "year": dt0.year,
                "month": dt0.month,
                "kWh_th": Q_kWh,
                "kWh_el": E_kWh,
                "fan_hours": frac_h if V_m3h>0 else 0.0,
            })
    df = pd.DataFrame.from_records(records)
    monthly = df.groupby(["year","month"],as_index=False)[["kWh_th","kWh_el","fan_hours"]].sum()
    yearly  = df.groupby(["year"],as_index=False)[["kWh_th","kWh_el","fan_hours"]].sum()
    return monthly, yearly

# ------------------------------
# Excel-Export
# ------------------------------
from io import BytesIO
def build_excel(plant: Plant, monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        m = monthly.copy(); y = yearly.copy()
        for col in ("kWh_th","kWh_el"):
            if col in m: m[col]=m[col].round(0)
            if col in y: y[col]=y[col].round(0)
        if "fan_hours" in m: m["fan_hours"]=m["fan_hours"].round(1)
        if "fan_hours" in y: y["fan_hours"]=y["fan_hours"].round(1)
        m.to_excel(writer,index=False,sheet_name="Monate")
        y.to_excel(writer,index=False,sheet_name="Jahreswerte")
    return out.getvalue()

def build_excel_months(monthly: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        m = monthly.copy()
        if not m.empty:
            m["kWh_th"]=m["kWh_th"].round(0)
            m["kWh_el"]=m["kWh_el"].round(0)
            m["fan_hours"]=m["fan_hours"].round(1)
        m.to_excel(writer,index=False,sheet_name="Monate")
    return out.getvalue()

# ------------------------------
# PDF-Export (übersichtlich mit Platypus)
# ------------------------------
def build_pdf(try_info: str, defaults: Defaults, plant: Plant, monthly: pd.DataFrame, yearly: pd.DataFrame) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab ist nicht installiert.")
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]; H1.fontSize=14
    H2 = styles["Heading2"]; H2.fontSize=12
    N = styles["BodyText"]; N.leading=14

    story=[]
    story += [Paragraph("ISO 50001 – Heizenergieabschätzung Lüftungsanlagen (v1)",H1),Spacer(1,6)]
    story += [Paragraph(f"Erzeugt: {datetime.now():%d.%m.%Y %H:%M}",N),Spacer(1,8)]
    story += [Paragraph("Ergebnisse Monats-/Jahressummen",H2)]

    if not yearly.empty:
        ysum = yearly.iloc[0].copy()
        ysum["kWh_th"]=round(ysum["kWh_th"],0)
        ysum["kWh_el"]=round(ysum["kWh_el"],0)
        ysum["fan_hours"]=round(ysum["fan_hours"],1)
        data=[["Jahr","kWh_th","kWh_el","Betriebsstunden"],
              [int(ysum["year"]),int(ysum["kWh_th"]),int(ysum["kWh_el"]),f"{ysum['fan_hours']:.1f}"]]
        tbl=Table(data,hAlign="LEFT"); tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2"))]))
        story+=[Paragraph("Jahressumme",H2),tbl]

    if not monthly.empty:
        m=monthly.copy()
        m["kWh_th"]=m["kWh_th"].round(0)
        m["kWh_el"]=m["kWh_el"].round(0)
        m["fan_hours"]=m["fan_hours"].round(1)
        data=[["Jahr","Monat","kWh_th","kWh_el","Betriebsstunden"],*m.astype({"year":int,"month":int}).values.tolist()]
        tbl=Table(data,hAlign="LEFT",repeatRows=1)
        story+=[Paragraph("Monatswerte",H2),tbl]

    doc.build(story); buf.seek(0); return buf.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(page_title="Heizenergie – ISO 50001", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# --- Session-State ---
if "try_df" not in st.session_state: st.session_state["try_df"]=None
if "try_info" not in st.session_state: st.session_state["try_info"]=""
if "monthly_df" not in st.session_state: st.session_state["monthly_df"]=None
if "yearly_df" not in st.session_state: st.session_state["yearly_df"]=None

# 1) TRY-Upload
with st.expander("1) TRY-CSV laden", expanded=True):
    f = st.file_uploader("TRY-CSV auswählen", type=["csv"]) 
    if f is not None:
        raw=pd.read_csv(f)
        dt_col=[c for c in raw.columns if "date" in c.lower() or "time" in c.lower()][0]
        t_col=[c for c in raw.columns if "t" in c.lower()][0]
        raw["datetime"]=raw[dt_col].astype(str).apply(parse_datetime)
        raw=raw.rename(columns={t_col:"T_out_C"})
        try_df=raw[["datetime","T_out_C"]].dropna()
        st.session_state["try_df"]=try_df
        st.session_state["try_info"]=f"Datensätze: {len(try_df)}"
        st.success("TRY-CSV eingelesen.")
    elif st.session_state["try_df"] is not None:
        st.info("TRY-CSV bereits geladen."); st.text(st.session_state["try_info"])

# 2) Defaults
defaults=Defaults()

# 3) Anlage
plant=Plant("A01","Beispiel",5000,1,True,0.7,5.0,None,empty_week_plan())

# 4) Berechnen
if st.button("Berechnen",type="primary"):
    if st.session_state["try_df"] is None: st.error("Bitte CSV laden")
    else:
        m,y=compute(st.session_state["try_df"],plant,defaults)
        st.session_state["monthly_df"]=m; st.session_state["yearly_df"]=y
        st.success("Berechnung abgeschlossen.")

# 5) Ergebnisse
m=st.session_state["monthly_df"]; y=st.session_state["yearly_df"]
if m is not None and y is not None:
    st.subheader("Monate")
    st.dataframe(m)
    st.subheader("Jahr")
    st.dataframe(y)

    col1,col2,col3,col4=st.columns(4)
    col1.download_button("Excel Gesamt",build_excel(plant,m,y),"Heizenergie.xlsx")
    col2.download_button("Excel Monate",build_excel_months(m),"Heizenergie_Monate.xlsx")
    col3.download_button("CSV Monate",m.to_csv(index=False).encode(),"Heizenergie_Monate.csv")
    if REPORTLAB_OK:
        col4.download_button("PDF Bericht",build_pdf(st.session_state["try_info"],defaults,plant,m,y),"Bericht.pdf")
