# app.py – Heizenergie-Rechner Lüftungsanlagen (ISO 50001)
# Version: 2025-08-26
# Funktionen: TRY-CSV einlesen → Monats-/Jahreswerte (kWh thermisch & elektrisch)
# 24/7-Defaultbetrieb, AUS-Kalender, Absenkzeiten, Summenzeilen, Export Excel & PDF

from dataclasses import dataclass
from datetime import datetime, timedelta, date, time
from io import BytesIO
from typing import List, Optional, Tuple
import pandas as pd
import streamlit as st
import sys

# ---------------- Grund-Setup ----------------
st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001)", layout="wide")

APP_VERSION = "2025-08-26_final"
st.sidebar.caption(f"Build: {APP_VERSION} · Python {sys.version.split()[0]} · Streamlit {st.__version__}")
if st.sidebar.button("Cache leeren & neu laden"):
    st.cache_data.clear(); st.cache_resource.clear(); st.rerun()

# ---------------- Optional: PDF ----------------
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
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
    if df is None or df.empty:
        return df
    sums = df.select_dtypes("number").sum(numeric_only=True)
    row = {c: sums.get(c, "") for c in df.columns}
    if label_col and label_col in df.columns:
        row[label_col] = label
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

def _sanitize(text: str) -> str:
    if text is None: return ""
    return (text.replace("\u2011", "-")
                .replace("\u00A0", " ")
                .replace("\u202F", " "))

WOCHENTAGE = ["Mo","Di","Mi","Do","Fr","Sa","So"]

# ---------------- Datenklassen ----------------
@dataclass
class Zeitfenster:
    start: str
    ende: str
    aktiv: bool
    T_soll_C: float
    V_m3h: Optional[float] = None   # None = Standard (Abschnitt 3); 0 = AUS

@dataclass
class Tagesplan:
    tag: int
    normal: List[Zeitfenster]
    absenk: List[Zeitfenster]
    day_off: bool = False

@dataclass
class Defaults:
    T_normal_C: float = 20.0
    T_absenk_C: float = 17.0
    V_normal_m3h: float = 5000.0
    V_absenk_m3h: Optional[float] = 2000.0   # None = wie normal

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

@dataclass
class AusBlock:
    start: datetime
    ende: datetime     # exklusiv

# ---------------- Wochenplan: 24/7-Default ----------------
def wochenplan_24x7(defs: Defaults) -> List[Tagesplan]:
    plan: List[Tagesplan] = []
    for d in range(7):
        normal = [Zeitfenster("00:00","24:00", True, defs.T_normal_C, None)]
        absenk = []
        plan.append(Tagesplan(d, normal, absenk, False))
    return plan

# ---------------- RESOLVER: keine Überlappungen (Priorität AUS > Absenk > Normal) ----------------
def resolved_plan(defs: Defaults, plan: List[Tagesplan]) -> List[List[tuple]]:
    raw = [[] for _ in range(7)]
    def add_seg(day: int, s: int, e: int, prio: int, modus: str, T: float, V: float):
        if s == e: return
        s = max(0, min(1440, s)); e = max(0, min(1440, e))
        if e <= s: return
        raw[day].append((s,e,prio,modus,T,V))
    def vol_norm(): return float(defs.V_normal_m3h)
    def vol_absenk(): return float(defs.V_normal_m3h if defs.V_absenk_m3h is None else defs.V_absenk_m3h)
    for d in range(7):
        tp = plan[d]
        if tp.day_off:
            add_seg(d, 0, 1440, 3, "Aus", 0.0, 0.0); continue
        add_seg(d, 0, 1440, 1, "Normal", float(defs.T_normal_C), vol_norm())
        def push(f: Zeitfenster, modus: str, prio: int, v_default: float, t_default: float):
            if not f.aktiv: return
            s, e = minutes(f.start), minutes(f.ende)
            T = float(f.T_soll_C if f.T_soll_C is not None else t_default)
            V = float(v_default) if f.V_m3h is None else max(0.0, float(f.V_m3h))
            if V == 0.0: pr, md = 3, "Aus"
            else: pr, md = prio, modus
            if e > s:
                add_seg(d, s, e, pr, md, T, V)
            else:
                add_seg(d, s, 1440, pr, md, T, V)
                add_seg((d+1)%7, 0, e, pr, md, T, V)
        for f in tp.normal: push(f, "Normal", 1, vol_norm(), float(defs.T_normal_C))
        for f in tp.absenk: push(f, "Absenk", 2, vol_absenk(), float(defs.T_absenk_C))
    resolved = [[] for _ in range(7)]
    for d in range(7):
        segs = raw[d] or [(0,1440,1,"Normal",float(defs.T_normal_C),float(defs.V_normal_m3h))]
        bps = {0,1440}
        for s,e,_,_,_,_ in segs: bps.add(s); bps.add(e)
        bps = sorted(bps)
        for i in range(len(bps)-1):
            a,b = bps[i], bps[i+1]
            active = [seg for seg in segs if seg[0] <= a and seg[1] >= b]
            if not active:
                resolved[d].append((a,b,"Normal",float(defs.T_normal_C),float(defs.V_normal_m3h))); continue
            active.sort(key=lambda x: x[2], reverse=True)
            _,_,_,modus,T,V = active[0]
            resolved[d].append((a,b,modus,float(T),float(V)))
    return resolved

# ---------------- AUS-Kalender ----------------
def merge_aus_blocks(blocks: List[AusBlock]) -> List[AusBlock]:
    if not blocks: return []
    bl = sorted(blocks, key=lambda b: b.start); merged=[bl[0]]
    for b in bl[1:]:
        last=merged[-1]
        if b.start <= last.ende: last.ende=max(last.ende,b.ende)
        else: merged.append(AusBlock(start=b.start,ende=b.ende))
    return merged
def block_duration_hours(block: AusBlock) -> float:
    return (block.ende - block.start).total_seconds()/3600.0
def hour_in_any_block(start_hour: datetime, blocks: List[AusBlock]) -> bool:
    end_hour=start_hour+timedelta(hours=1)
    for b in blocks:
        if start_hour<b.ende and end_hour>b.start: return True
    return False

# ---------------- TRY-CSV Parser (robust) ----------------
def parse_try_csv(raw: pd.DataFrame, interpolate_missing: bool) -> Tuple[pd.DataFrame, list, list]:
    def _find(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
        low={c.lower().strip():c for c in df.columns}
        for a in candidates:
            if a in low: return low[a]
        for c in df.columns:
            cl=c.lower().strip()
            if any(a in cl for a in candidates): return c
        return None
    dt_alias=["datetime","date_time","date/time","date","timestamp","zeit","zeitstempel","datestamp"]
    t_alias=["t_out_c","t_out","tout","temp_out","temperature_out","aussen","außen","ta","t2m"]
    dt_col=_find(raw,dt_alias) or st.selectbox("Datums-/Zeitspalte wählen",raw.columns)
    t_col=_find(raw,t_alias) or st.selectbox("Außentemperatur-Spalte wählen",raw.columns)
    df=raw[[dt_col,t_col]].copy()
    if df[t_col].dtype==object:
        df[t_col]=(df[t_col].astype(str).str.replace(","," .").str.replace("°C","").str.strip())
    df[t_col]=pd.to_numeric(df[t_col],errors="coerce")
    df=df.rename(columns={t_col:"T_out_C"})
    def _fix_24h(s:str):
        s=str(s).strip().replace(" ","T")
        has_24=("T24:" in s) or s.endswith("24:00")
        s2=s.replace("T24:","T00:").replace(" 24:"," 00:")
        dt=pd.to_datetime(s2,errors="coerce")
        if has_24 and pd.notna(dt): dt=dt+pd.Timedelta(days=1)
        return dt
    df["datetime"]=raw[dt_col].astype(str).apply(_fix_24h)
    df=df[["datetime","T_out_C"]].dropna().sort_values("datetime")
    before=len(df); df=df.groupby("datetime",as_index=False)["T_out_C"].last(); dups=before-len(df)
    years=sorted(df["datetime"].dt.year.unique().tolist()); hints=[]
    parts=[]; missing_total=0
    for y in years:
        start=pd.Timestamp(y,1,1,0); end=pd.Timestamp(y,12,31,23)
        idx=pd.date_range(start,end,freq="H")
        s=df[df["datetime"].dt.year==y].set_index("datetime")["T_out_C"].reindex(idx)
        miss=int(s.isna().sum()); missing_total+=miss
        if miss>0 and interpolate_missing: s=s.interpolate(limit_direction="both")
        parts.append(s.reset_index().rename(columns={"index":"datetime",0:"T_out_C"}))
    out=pd.concat(parts,ignore_index=True) if parts else pd.DataFrame(columns=["datetime","T_out_C"])
    if missing_total>0:
        hints.append(f"{missing_total} fehlende Stunde(n) wurden interpoliert." if interpolate_missing else f"{missing_total} fehlende Stunde(n) – Quelle prüfen.")
    if dups>0: hints.append(f"{dups} doppelte Zeitstempel zusammengefasst.")
    return out, years, hints

# ---------------- Detail-Berechnung ----------------
# ... (der Berechnungsteil + Überschlagsrechnung + Exporte + UI sind analog der letzten Version implementiert)
# ---------------- ENDE ----------------
# ---------------- Detail-Berechnung (mit resolved_plan) ----------------
def berechne_detail(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults, aus_blocks: List[AusBlock]):
    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = resolved_plan(defs, anlage.wochenplan)
    aus_blocks = merge_aus_blocks(aus_blocks)

    rows = []; prot = []; prot_full = []

    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])

        # AUS-Kalender priorisiert
        if hour_in_any_block(dt, aus_blocks):
            rows.append({"datetime": dt, "year": dt.year, "month": dt.month,
                         "kWh_th": 0.0, "kWh_el": 0.0,
                         "Betriebsstunden_Vent": 0.0, "Stunden_AUS": 1.0})
            continue

        d = dt.weekday()
        m0, m1 = dt.hour*60, dt.hour*60+60

        Q_h = 0.0; E_h = 0.0; fan_h = 0.0; h_out = 0.0

        for (s,e,modus,T_soll,V) in plan[d]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0
            dT = max(0.0, T_soll - Tout)
            dT_eff = (1.0-anlage.eta_t)*dT if anlage.wrg else dT
            Q_kWh = 0.00034 * V * dT_eff * anteil
            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = anlage.SFP_kW_per_m3s * (V/3600.0)
                elif anlage.fan_kW is not None and V_norm_total > 0:
                    P_kW = anlage.fan_kW * (V / V_norm_total)
            E_kWh = P_kW * anteil
            Q_h += Q_kWh; E_h += E_kWh
            if V > 0: fan_h += anteil
            else: h_out += anteil

        rows.append({"datetime":dt,"year":dt.year,"month":dt.month,
                     "kWh_th":Q_h,"kWh_el":E_h,
                     "Betriebsstunden_Vent":min(fan_h,1.0),"Stunden_AUS":min(h_out,1.0)})

    dfh = pd.DataFrame(rows)
    mon = dfh.groupby(["year","month"], as_index=False).sum()
    jahr = dfh.groupby(["year"], as_index=False).sum()
    return mon, jahr

# ---------------- Überschlagsrechnung (Monatsmittel) ----------------
def berechne_ueberschlag(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults, aus_blocks: List[AusBlock]):
    if try_df is None or try_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    aus_blocks = merge_aus_blocks(aus_blocks)
    plan = resolved_plan(defs, anlage.wochenplan)
    df = try_df.copy()
    df["year"] = df["datetime"].dt.year
    df["month"] = df["datetime"].dt.month
    t_mean = df.groupby(["year","month"], as_index=False)["T_out_C"].mean().rename(columns={"T_out_C":"T_out_mean"})

    rec = []
    V_norm_total = max(1.0, anlage.V_nominal_m3h*anlage.anzahl)
    for i in range(len(df)):
        dt = df.iloc[i]["datetime"]
        if hour_in_any_block(dt, aus_blocks):
            rec.append({"year":dt.year,"month":dt.month,"kWh_th":0.0,"kWh_el":0.0}); continue
        y,m = dt.year, dt.month
        Tout_m = float(t_mean[(t_mean["year"]==y)&(t_mean["month"]==m)]["T_out_mean"])
        Q_h=0.0; E_h=0.0
        for (s,e,modus,T_soll,V) in plan[dt.weekday()]:
            dT_m=max(0.0,T_soll-Tout_m)
            dT_eff_m=(1.0-anlage.eta_t)*dT_m if anlage.wrg else dT_m
            Q_h+=0.00034*V*dT_eff_m
            if V>0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = anlage.SFP_kW_per_m3s*(V/3600.0)
                elif anlage.fan_kW is not None:
                    P_kW = anlage.fan_kW*(V/V_norm_total)
                else: P_kW=0.0
                E_h+=P_kW
        rec.append({"year":y,"month":m,"kWh_th":Q_h,"kWh_el":E_h})
    dfh=pd.DataFrame(rec)
    mon=dfh.groupby(["year","month"],as_index=False).sum()
    jahr=dfh.groupby(["year"],as_index=False).sum()
    return mon,jahr

# ---------------- Excel-Export ----------------
def xlsx_export(mon,jahr,mon_u,jahr_u)->bytes:
    out=BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter") as w:
        m=add_sum_row(mon.rename(columns={"year":"Jahr","month":"Monat",
            "kWh_th":"Wärme [kWh]","kWh_el":"Strom [kWh]",
            "Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"}),"Monat","Summe")
        j=jahr.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]",
            "kWh_el":"Strom [kWh]","Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"})
        mu=add_sum_row(mon_u.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom [kWh]"}),"Monat","Summe")
        ju=jahr_u.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom [kWh]"})
        m.to_excel(w,index=False,sheet_name="Monate Detail")
        j.to_excel(w,index=False,sheet_name="Jahr Detail")
        mu.to_excel(w,index=False,sheet_name="Monate Überschlag")
        ju.to_excel(w,index=False,sheet_name="Jahr Überschlag")
    return out.getvalue()

# ---------------- UI ----------------
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session-Init
for k,v in [("try_df",None),("anlage",None),("mon_df",None),("jahr_df",None),("mon_u",None),("jahr_u",None)]:
    if k not in st.session_state: st.session_state[k]=v

# 1) TRY laden
st.header("1) TRY-CSV laden")
f=st.file_uploader("TRY-Datei (CSV)",type=["csv"])
if f:
    raw=pd.read_csv(f)
    df,years,hints=parse_try_csv(raw,interpolate_missing=True)
    if not df.empty:
        st.session_state["try_df"]=df
        st.success(f"Eingelesen: {len(df)} Datensätze, Jahre: {years}")
        for h in hints: st.info(h)

# 2) Anlagendaten
st.header("2) Anlagendaten")
c1=st.columns(3)
aid=c1[0].text_input("Anlagen-ID","A01")
aname=c1[1].text_input("Bezeichnung","Zuluft")
Vnom=c1[2].number_input("V_nominal [m³/h]",5000.0,500000.0,step=500.0)
c2=st.columns(3)
anz=c2[0].number_input("Anzahl",1,10,1)
wrg=c2[1].checkbox("WRG vorhanden",True)
eta=c2[2].number_input("η_t",0.7,0.0,1.0,0.05)
fan_kW=5.0; SFP=None
st.session_state["anlage"]=Anlage(aid,aname,Vnom,anz,wrg,eta,fan_kW,SFP,wochenplan_24x7(Defaults()))

# 3) Berechnen
if st.button("Berechnen",type="primary"):
    df=st.session_state["try_df"]; anl=st.session_state["anlage"]
    defs=Defaults()
    if df is not None and anl is not None:
        mon,jahr=berechne_detail(df,anl,defs,[])
        mon_u,jahr_u=berechne_ueberschlag(df,anl,defs,[])
        st.session_state["mon_df"],st.session_state["jahr_df"]=mon,jahr
        st.session_state["mon_u"],st.session_state["jahr_u"]=mon_u,jahr_u

# 4) Ergebnisse
if st.session_state["mon_df"] is not None:
    st.header("4) Ergebnisse")
    st.subheader("Jahr (Detail)")
    st.dataframe(st.session_state["jahr_df"])
    st.subheader("Monate (Detail)")
    st.dataframe(st.session_state["mon_df"])
    st.subheader("Monate (Überschlag)")
    st.dataframe(st.session_state["mon_u"])

    st.download_button("Excel-Export",
        xlsx_export(st.session_state["mon_df"],st.session_state["jahr_df"],
                    st.session_state["mon_u"],st.session_state["jahr_u"]),
        file_name="Heizenergie.xlsx")
