# app.py – Heizenergie-Rechner Lüftungsanlagen (TRY → Monats/Jahreswerte)
# Deutsch · robuste TRY-Prüfung · AUS-Kalender · 24/7-Default · Plan-Resolver (keine Überlappungen)
# Absenk beeinflusst Ventilatorstunden nicht (sofern V>0) · saubere kWh · PDF/Excel

from dataclasses import dataclass
from datetime import datetime, timedelta, date, time
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
import sys

# ---------------- Sidebar: Build-Info ----------------
APP_VERSION = "2025-08-26_resolved_plan_v1"
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

# ---------------- Datenklassen ----------------
@dataclass
class Zeitfenster:
    start: str
    ende: str
    aktiv: bool
    T_soll_C: float
    V_m3h: Optional[float] = None  # None = Standard (Abschnitt 3)

@dataclass
class Tagesplan:
    tag: int                 # 0=Mo..6=So
    normal: List[Zeitfenster]
    absenk: List[Zeitfenster]
    day_off: bool = False    # ganzer Tag AUS

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

@dataclass
class AusBlock:
    start: datetime   # inkl. Uhrzeit
    ende: datetime    # exklusiv

WOCHENTAGE = ["Mo","Di","Mi","Do","Fr","Sa","So"]

# ---- 24/7-Default: alle Tage 00:00–24:00 Normal ----
def wochenplan_24x7(defs: Defaults) -> List[Tagesplan]:
    plan: List[Tagesplan] = []
    for d in range(7):
        normal = [Zeitfenster("00:00","24:00", True, defs.T_normal_C, None)]
        absenk = []  # zunächst leer
        plan.append(Tagesplan(d, normal, absenk, False))
    return plan

# ---- RESOLVER: macht aus Fenstern einen überlappungsfreien Tagesplan (AUS > Absenk > Normal) ----
def resolved_plan(defs: Defaults, plan: List[Tagesplan]) -> List[List[tuple]]:
    """
    Ergebnis je Tag: Liste nicht überlappender Segmente:
    (startMin, endMin, modus, T_soll, V) mit Priorität AUS > Absenk > Normal.
    Basis ist immer ein 00:00–24:00 Normal-Segment (T/V aus Defaults), außer day_off=True.
    Fenster über Mitternacht werden aufgeteilt (aktueller/folgender Tag).
    """
    # Rohfenster je Tag sammeln
    raw = [[] for _ in range(7)]  # je Tag: Liste (s,e,prio,modus,T,V)
    # Hilfsfunktionen
    def add_seg(day: int, s: int, e: int, prio: int, modus: str, T: float, V: float):
        if s == e: return
        s = max(0, min(1440, s)); e = max(0, min(1440, e))
        if e <= s: return
        raw[day].append((s,e,prio,modus,T,V))

    def vol_norm():
        return float(defs.V_normal_m3h)
    def vol_absenk():
        if defs.V_absenk_m3h is None:
            return float(defs.V_normal_m3h)
        return float(defs.V_absenk_m3h)

    for d in range(7):
        tp = plan[d]
        if tp.day_off:
            # ganzer Tag AUS
            add_seg(d, 0, 1440, 3, "Aus", 0.0, 0.0)
            continue

        # Basis 24/7 Normal
        add_seg(d, 0, 1440, 1, "Normal", float(defs.T_normal_C), vol_norm())

        # Hilfsfunktion: Fenster hinzufügen (split bei Über-Mitternacht)
        def push(f: Zeitfenster, modus: str, prio: int, v_default: float, t_default: float):
            if not f.aktiv: return
            s, e = minutes(f.start), minutes(f.ende)
            # T
            T = float(f.T_soll_C) if f.T_soll_C is not None else t_default
            # V
            if f.V_m3h is None:
                V = float(v_default)
            else:
                V = max(0.0, float(f.V_m3h))
            # AUS-Fenster hat höchste Prio
            if V == 0.0:
                prio_use = 3; modus_use = "Aus"
            else:
                prio_use = prio; modus_use = modus

            if e > s:
                add_seg(d, s, e, prio_use, modus_use, T, V)
            else:
                add_seg(d, s, 1440, prio_use, modus_use, T, V)
                add_seg((d+1)%7, 0, e, prio_use, modus_use, T, V)

        # Nutzerfenster
        for f in tp.normal:
            push(f, "Normal", 1, vol_norm(), float(defs.T_normal_C))
        for f in tp.absenk:
            push(f, "Absenk", 2, vol_absenk(), float(defs.T_absenk_C))

    # Auflösen: Breakpoints & höchste Priorität wählen
    resolved = [[] for _ in range(7)]
    for d in range(7):
        segs = raw[d]
        # Falls leer (sollte nicht passieren außer day_off), trotzdem Basis normal
        if not segs:
            segs = [(0,1440,1,"Normal",float(defs.T_normal_C),vol_norm())]
        bps = {0,1440}
        for s,e,_,_,_,_ in segs:
            bps.add(s); bps.add(e)
        bps = sorted(bps)
        for i in range(len(bps)-1):
            a,b = bps[i], bps[i+1]
            # aktive Segmente, die [a,b) komplett abdecken
            active = [seg for seg in segs if seg[0] <= a and seg[1] >= b]
            if not active:
                # Fallback Normal
                resolved[d].append((a,b,"Normal",float(defs.T_normal_C),vol_norm()))
                continue
            # höchste Priorität wählen (AUS=3 > Absenk=2 > Normal=1)
            active.sort(key=lambda x: x[2], reverse=True)
            s,e,prio,modus,T,V = active[0]
            resolved[d].append((a,b,modus,float(T),float(V)))
    return resolved

# ---------------- AUS-Kalender Utils ----------------
def merge_aus_blocks(blocks: List[AusBlock]) -> List[AusBlock]:
    if not blocks:
        return []
    bl = sorted(blocks, key=lambda b: b.start)
    merged = [bl[0]]
    for b in bl[1:]:
        last = merged[-1]
        if b.start <= last.ende:
            last.ende = max(last.ende, b.ende)
        else:
            merged.append(AusBlock(start=b.start, ende=b.ende))
    return merged

def block_duration_hours(block: AusBlock) -> float:
    return (block.ende - block.start).total_seconds() / 3600.0

def hour_in_any_block(start_hour: datetime, blocks: List[AusBlock]) -> bool:
    end_hour = start_hour + timedelta(hours=1)
    for b in blocks:
        if start_hour < b.ende and end_hour > b.start:
            return True
    return False

# ---------------- TRY-CSV Parser ----------------
def parse_try_csv(raw: pd.DataFrame, interpolate_missing: bool) -> Tuple[pd.DataFrame, list, list]:
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

    if df[t_col].dtype == object:
        df[t_col] = (df[t_col].astype(str)
                               .str.replace(",", ".", regex=False)
                               .str.replace("°C", "", regex=False)
                               .str.strip())
    df[t_col] = pd.to_numeric(df[t_col], errors="coerce")
    df = df.rename(columns={t_col: "T_out_C"})

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

    before = len(df)
    df = df.groupby("datetime", as_index=False)["T_out_C"].last()
    dups = before - len(df)

    years = sorted(df["datetime"].dt.year.unique().tolist())
    hints = []

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

    out = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["datetime","T_out_C"])

    if missing_total > 0:
        hints.append(f"{missing_total} fehlende Stunde(n) wurden interpoliert." if interpolate_missing
                     else f"{missing_total} fehlende Stunde(n) – Quelle prüfen (oder Interpolation aktivieren).")
    if dups > 0:
        hints.append(f"{dups} doppelte Zeitstempel zusammengefasst.")

    return out, years, hints

# ---------------- Detail-Berechnung (mit resolved plan) ----------------
def berechne_detail(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults, aus_blocks: List[AusBlock]):
    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = resolved_plan(defs, anlage.wochenplan)   # <<<<<< NEU
    aus_blocks = merge_aus_blocks(aus_blocks)

    rows = []; prot = []; prot_full = []

    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])

        # AUS-Kalender: komplette Stunde aus
        if hour_in_any_block(dt, aus_blocks):
            rows.append({"datetime": dt, "year": dt.year, "month": dt.month,
                         "kWh_th": 0.0, "kWh_el": 0.0, "Betriebsstunden_Vent": 0.0,
                         "Stunden_AUS": 1.0})
            rec = {"Zeit": dt, "Modus": "AUS-Kalender", "Anteil [h]": 1.0,
                   "T_out [°C]": round(Tout,1), "T_soll [°C]": None, "ΔT_eff [K]": None,
                   "V [m³/h]": 0, "Wärme [kWh]": 0.0, "P_fan [kW]": 0.0, "Strom [kWh]": 0.0}
            prot_full.append(rec)
            if len(prot) < 500: prot.append(rec)
            continue

        d = dt.weekday()
        m0, m1 = dt.hour*60 + dt.minute, dt.hour*60 + dt.minute + 60

        Q_h = 0.0; E_h = 0.0; fan_h = 0.0; h_out = 0.0

        # resolved_plan: Segmente an diesem Tag sind disjunkt und decken 0..1440 ab
        for (s,e,modus,T_soll,V) in plan[d]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            dT = max(0.0, float(T_soll) - Tout)
            dT_eff = (1.0 - clamp(float(anlage.eta_t),0.0,1.0))*dT if anlage.wrg else dT
            Q_kWh = 0.00034 * V * dT_eff * anteil

            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = float(anlage.SFP_kW_per_m3s) * (V/3600.0)
                elif anlage.fan_kW is not None and V_norm_total > 0:
                    P_kW = float(anlage.fan_kW) * (V / V_norm_total)
            E_kWh = P_kW * anteil

            Q_h += Q_kWh; E_h += E_kWh
            if V > 0: fan_h += anteil
            else: h_out += anteil

            rec = {"Zeit":dt,"Modus":modus,"Anteil [h]":round(anteil,3),
                   "T_out [°C]":round(Tout,1),
                   "T_soll [°C]":round(T_soll,1) if V>0 else None,
                   "ΔT_eff [K]":round(dT_eff,2) if V>0 else None,
                   "V [m³/h]":round(V,0),
                   "Wärme [kWh]":round(Q_kWh,3),"P_fan [kW]":round(P_kW,2),"Strom [kWh]":round(E_kWh,3)}
            prot_full.append(rec)
            if len(prot) < 500: prot.append(rec)

        # Sicherheits-Clamp wegen Rundungen: max 1 h pro Stunde
        fan_h = min(fan_h, 1.0)
        h_out = min(h_out, 1.0)

        rows.append({"datetime":dt,"year":dt.year,"month":dt.month,
                     "kWh_th":Q_h,"kWh_el":E_h,"Betriebsstunden_Vent":fan_h,
                     "Stunden_AUS": h_out})

    dfh = pd.DataFrame.from_records(rows)
    mon = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]].sum()
    jahr = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]].sum()
    return mon, jahr, pd.DataFrame(prot), pd.DataFrame(prot_full)

# ---------------- Überschlagsrechnung (ebenfalls mit resolved plan) ----------------
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
    for i in range(len(df)):
        dt = df.iloc[i]["datetime"]
        if hour_in_any_block(dt, aus_blocks):
            rec.append({"year":dt.year,"month":dt.month,"kWh_th":0.0,"kWh_el":0.0})
            continue
        y, m = dt.year, dt.month
        Tout_m = float(t_mean[(t_mean["year"]==y)&(t_mean["month"]==m)]["T_out_mean"].values[0])

        day = dt.weekday()
        m0, m1 = dt.hour*60 + dt.minute, dt.hour*60 + dt.minute + 60

        Q_h = 0.0; E_h = 0.0
        for (s,e,modus,T_soll,V) in plan[day]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            dT_m = max(0.0, float(T_soll) - Tout_m)
            dT_eff_m = dT_m  # WRG-Effekt spielt für Überschlag auch (hier) eine Rolle? → ja:
            # Korrektur: WRG berücksichtigen wie im Detail
            # (wir brauchen Zugriff auf Anlage; siehe unten fix)
        # >>> Fix: wir brauchen Anlage in dieser Funktion:
    ```
