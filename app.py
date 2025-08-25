# app.py – Heizenergie-Rechner Lüftungsanlagen (TRY → Monats/Jahreswerte)
# Deutsch · robuste TRY-Prüfung · AUS-Kalender · 24/7-Default · Kopfwerte→Kalender · Detail & Überschlag · saubere kWh · PDF/Excel

from dataclasses import dataclass
from datetime import datetime, timedelta, date, time
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
import sys

# ---------------- Sidebar: Build-Info ----------------
APP_VERSION = "2025-08-25_final_24x7_v1"
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

def human_hours(delta_hours: float) -> str:
    h = float(delta_hours)
    d = int(h // 24)
    r = h - d*24
    return f"{d} d {r:.0f} h" if d else f"{r:.0f} h"

# PDF: Sonderzeichen sanieren
def _sanitize(text: str) -> str:
    if text is None: return ""
    return (text
            .replace("\u2011", "-")   # non-breaking hyphen
            .replace("\u00A0", " ")   # non-breaking space
            .replace("\u202F", " ")   # narrow no-break space
            )

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

def normiere_wochenplan(plan: List[Tagesplan]) -> List[List[tuple]]:
    """
    je Tag: (startMin, endMin, 'Normal'|'Absenk'|'Aus', T_soll, V_override)
    'Aus' wird mit V_override=0.0 erzeugt (T_soll egal, V=0 ⇒ 0 kWh).
    """
    out = [[] for _ in range(7)]

    def add(tag: int, f: Zeitfenster, modus: str):
        if not f.aktiv:
            return
        s, e = minutes(f.start), minutes(f.ende)
        if s == e:
            return
        Vovr = None if f.V_m3h is None else float(f.V_m3h)
        if e > s:
            out[tag].append((s, e, modus, float(f.T_soll_C), Vovr))
        else:
            out[tag].append((s, 1440, modus, float(f.T_soll_C), Vovr))
            out[(tag + 1) % 7].append((0, e, modus, float(f.T_soll_C), Vovr))

    for d in plan:
        if d.day_off:
            out[d.tag] = [(0, 1440, "Aus", 0.0, 0.0)]
            continue
        for f in d.normal: add(d.tag, f, "Normal")
        for f in d.absenk: add(d.tag, f, "Absenk")

    for i in range(7):
        out[i].sort(key=lambda x: x[0])
    return out

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

# ---------------- TRY-CSV Parser (robust) ----------------
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

# ---------------- Detail-Berechnung ----------------
def berechne_detail(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults, aus_blocks: List[AusBlock]):
    """Stündliche Detailrechnung (Monate, Jahr, Protokoll (Ausschnitt), Vollprotokoll)."""
    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = normiere_wochenplan(anlage.wochenplan)
    aus_blocks = merge_aus_blocks(aus_blocks)

    rows = []; prot = []; prot_full = []

    for i in range(len(try_df)):
        dt = try_df.iloc[i]["datetime"]
        Tout = float(try_df.iloc[i]["T_out_C"])

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

        for (s,e,modus,T_soll,V_ovr) in plan[d]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            if V_ovr is not None:
                V = max(0.0, float(V_ovr))
            else:
                V = V_norm_total if modus=="Normal" else (V_norm_total if defs.V_absenk_m3h is None else float(defs.V_absenk_m3h))

            dT = max(0.0, float(T_soll) - Tout)
            dT_eff = (1.0 - clamp(float(anlage.eta_t),0.0,1.0))*dT if anlage.wrg else dT
            Q_kWh = 0.00034 * V * dT_eff * anteil

            P_kW = 0.0
            if V > 0:
                if anlage.SFP_kW_per_m3s is not None:
                    P_kW = float(anlage.SFP_kW_per_m3s) * (V/3600.0)
                elif anlage.fan_kW is not None:
                    P_kW = float(anlage.fan_kW) * (V / max(1.0, V_norm_total))
            E_kWh = P_kW * anteil

            Q_h += Q_kWh; E_h += E_kWh; fan_h += (anteil if V>0 else 0.0)
            if V == 0.0: h_out += anteil

            rec = {"Zeit":dt,"Modus":modus,"Anteil [h]":round(anteil,3),
                   "T_out [°C]":round(Tout,1),"T_soll [°C]":round(T_soll,1) if V>0 else None,
                   "ΔT_eff [K]":round(dT_eff,2) if V>0 else None,"V [m³/h]":round(V,0),
                   "Wärme [kWh]":round(Q_kWh,3),"P_fan [kW]":round(P_kW,2),"Strom [kWh]":round(E_kWh,3)}
            prot_full.append(rec)
            if len(prot) < 500: prot.append(rec)

        rows.append({"datetime":dt,"year":dt.year,"month":dt.month,
                     "kWh_th":Q_h,"kWh_el":E_h,"Betriebsstunden_Vent":fan_h,
                     "Stunden_AUS": h_out})

    dfh = pd.DataFrame.from_records(rows)
    mon = dfh.groupby(["year","month"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]].sum()
    jahr = dfh.groupby(["year"], as_index=False)[["kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]].sum()
    return mon, jahr, pd.DataFrame(prot), pd.DataFrame(prot_full)

# ---------------- Überschlagsrechnung ----------------
def berechne_ueberschlag(try_df: pd.DataFrame, anlage: Anlage, defs: Defaults, aus_blocks: List[AusBlock]):
    if try_df is None or try_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    aus_blocks = merge_aus_blocks(aus_blocks)

    df = try_df.copy()
    df["year"] = df["datetime"].dt.year
    df["month"] = df["datetime"].dt.month
    t_mean = df.groupby(["year","month"], as_index=False)["T_out_C"].mean().rename(columns={"T_out_C":"T_out_mean"})

    V_norm_total = float(anlage.V_nominal_m3h) * int(anlage.anzahl)
    V_norm_total = float(clamp(V_norm_total, 0.0, 500000.0))
    plan = normiere_wochenplan(anlage.wochenplan)

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
        for (s,e,modus,T_soll,V_ovr) in plan[day]:
            ol = overlap_mins(m0,m1,s,e)
            if ol <= 0: continue
            anteil = ol/60.0

            if V_ovr is not None:
                V = max(0.0, float(V_ovr))
            else:
                V = V_norm_total if modus=="Normal" else (V_norm_total if defs.V_absenk_m3h is None else float(defs.V_absenk_m3h))

            dT_m = max(0.0, float(T_soll) - Tout_m)
            dT_eff_m = (1.0 - clamp(float(anlage.eta_t),0.0,1.0))*dT_m if anlage.wrg else dT_m
            Q_kWh = 0.00034 * V * dT_eff_m * anteil

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

# ---------------- Exporte (Excel/PDF) ----------------
def xlsx_export(mon: pd.DataFrame, jahr: pd.DataFrame, prot: pd.DataFrame,
                mon_ue: pd.DataFrame, jahr_ue: pd.DataFrame, aus_blocks: List[AusBlock]) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        m = mon.copy(); j = jahr.copy(); p = prot.copy()
        mu = mon_ue.copy(); ju = jahr_ue.copy()

        # Runden & Deutsch
        for col in ("kWh_th","kWh_el","Stunden_AUS","Betriebsstunden_Vent"):
            if col in m: m[col] = m[col].astype(float).round(0 if col in ("kWh_th","kWh_el") else 1)
            if col in j: j[col] = j[col].astype(float).round(0 if col in ("kWh_th","kWh_el") else 1)
        for col in ("kWh_th","kWh_el"):
            if col in mu: mu[col] = mu[col].round(0)
            if col in ju: ju[col] = ju[col].round(0)
        if not p.empty:
            # Protokoll angenehmer runden
            for c in ("Wärme [kWh]","Strom [kWh]"):
                if c in p: p[c] = p[c].astype(float).round(3)
            if "P_fan [kW]" in p: p["P_fan [kW]"] = p["P_fan [kW]"].astype(float).round(2)

        m_de = m.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]",
                                 "Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"})
        j_de = j.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]",
                                 "Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"})
        mu_de = mu.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        ju_de = ju.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})

        # Summenzeile: nur bei MONATEN
        m_de = add_sum_row(m_de, label_col="Monat", label="Summe")
        mu_de = add_sum_row(mu_de, label_col="Monat", label="Summe")
        # Keine Summenzeile bei Jahres-Tabellen (j_de, ju_de)

        m_de.to_excel(w, index=False, sheet_name="Monate (Detail)")
        j_de.to_excel(w, index=False, sheet_name="Jahr (Detail)")
        mu_de.to_excel(w, index=False, sheet_name="Monate (Überschlag)")
        ju_de.to_excel(w, index=False, sheet_name="Jahr (Überschlag)")
        if not p.empty:
            p.to_excel(w, index=False, sheet_name="Protokoll (Ausschnitt)")

        # AUS-Kalender
        aus_list = merge_aus_blocks(aus_blocks)
        if aus_list:
            df_aus = pd.DataFrame([{
                "Start": b.start, "Ende (exkl.)": b.ende,
                "Dauer [h]": round(block_duration_hours(b), 1)
            } for b in aus_list])
            df_aus.to_excel(w, index=False, sheet_name="AUS-Kalender")

    return out.getvalue()

def xlsx_single_sheet(df: pd.DataFrame, sheet_name: str = "Daten") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    return out.getvalue()

def _pdf_header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(A4[0]-18*mm, 10*mm, f"Seite {doc.page}")
    canvas.restoreState()

def pdf_export(info: str, defs: Defaults, anl: Anlage,
               mon: pd.DataFrame, jahr: pd.DataFrame,
               aus_blocks: List[AusBlock], try_hints: List[str]) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("ReportLab nicht installiert.")
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4,
                            leftMargin=18*mm, rightMargin=18*mm,
                            topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    H1=styles["Heading1"]; H1.fontSize=14
    H2=styles["Heading2"]; H2.fontSize=12
    N=styles["BodyText"]; N.leading=14
    S=ParagraphStyle("small", parent=N, fontSize=9, leading=12, textColor=colors.grey)

    def P(t, style=N):  # Paragraph mit Sanitizer
        return Paragraph(_sanitize(t), style)

    story=[]
    story+=[P("ISO 50001 – Heizenergie Lüftungsanlagen (vereinfachtes Verfahren v1)", H1), Spacer(1,6)]
    story+=[P(f"Erstellt: {datetime.now():%d.%m.%Y %H:%M}", N), Spacer(1,6)]
    story+=[P("Quelle / TRY", H2), P(info or "TRY-CSV (stündlich).", N)]

    if try_hints:
        story+=[Spacer(1,4), P("Datenqualität (Import-Hinweise):", S)]
        for h in try_hints:
            story+=[P(f"- {h}", S)]
    story+=[Spacer(1,8)]

    story+=[P("Annahmen & Parameter", H2)]
    ann = [
        ["Soll-Zuluft NORMAL [°C]", f"{defs.T_normal_C}"],
        ["Soll-Zuluft ABSENK [°C]", f"{defs.T_absenk_C}"],
        ["V NORMAL [m³/h]", f"{defs.V_normal_m3h:,.0f}".replace(",", ".")],
        ["V ABSENK [m³/h]", "wie normal" if defs.V_absenk_m3h is None else f"{defs.V_absenk_m3h:,.0f}".replace(",", ".")],
        ["WRG / η_t [–]", ("ja" if anl.wrg else "nein") + f" / {anl.eta_t}"],
        ["Ventilator-Modell", f"{'SFP '+str(anl.SFP_kW_per_m3s)+' kW/(m³/s)' if anl.SFP_kW_per_m3s is not None else 'fan_kW '+str(anl.fan_kW)+' kW'}"],
        ["V_nominal gesamt [m³/h]", f"{(anl.V_nominal_m3h*anl.anzahl):,.0f}".replace(",", ".")],
        ["Anlagenanzahl [–]", f"{anl.anzahl}"]
    ]
    ann = [[_sanitize(a), _sanitize(b)] for a,b in ann]
    T = Table([["Parameter","Wert"], *ann], hAlign="LEFT", colWidths=[70*mm, None])
    T.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE")
    ]))
    story+=[T, Spacer(1,8)]

    aus_list = merge_aus_blocks(aus_blocks)
    story+=[P("AUS-Kalender (Betriebsferien/Wartung)", H2)]
    if aus_list:
        rows = [["Start","Ende (exkl.)","Dauer [h]"]]
        rows += [[b.start.strftime("%d.%m.%Y %H:%M"), b.ende.strftime("%d.%m.%Y %H:%M"), f"{block_duration_hours(b):.1f}"] for b in aus_list]
        T2 = Table(rows, hAlign="LEFT", repeatRows=1, colWidths=[45*mm,45*mm,25*mm])
        T2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("ALIGN",(2,1),(2,-1),"RIGHT")
        ]))
        story += [T2]
    else:
        story += [P("Keine AUS-Blöcke hinterlegt.", N)]
    story+=[Spacer(1,8)]

    story+=[P("Methodik (Rechengang)", H2),
            P("Stündlich aktives Zeitfenster (Normal/Absenk/AUS) je Wochentag. "
              "AUS-Kalender hat Vorrang (Komplettabschaltung). "
              "Heizfall: ΔT = max(0, T_soll − T_out); mit WRG: ΔT_eff = (1 − η_t)·ΔT. "
              "Wärme: Q = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h. "
              "Ventilator: P_fan = SFP·(V/3600) oder fan_kW·(V/V_nominal); Energie E = P_fan · Anteil_h. "
              "Aggregation zu Monaten/Jahr.", N),
            Spacer(1,6),
            P("Hinweis: Vereinfachung (konst. Luftdichte/Wärmekapazität), keine Feuchteeinflüsse/Bypass-Logik.", S),
            Spacer(1,8)]

    if not jahr.empty:
        j = jahr.copy()
        for c in ("kWh_th","kWh_el"): j[c]=j[c].round(0)
        if "Betriebsstunden_Vent" in j: j["Betriebsstunden_Vent"]=j["Betriebsstunden_Vent"].round(1)
        if "Stunden_AUS" in j: j["Stunden_AUS"]=j["Stunden_AUS"].round(1)
        data=[["Jahr","Wärme [kWh]","Strom Vent. [kWh]","Betriebsstd. Vent.","Stunden AUS"],
              *j.reindex(columns=["year","kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]).values.tolist()]
        T3=Table(data,hAlign="LEFT")
        T3.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
                                ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                ("ALIGN",(1,1),(-1,-1),"RIGHT")]))
        story+=[P("Ergebnisse – Jahreswerte (Detail)", H2), T3, Spacer(1,8)]

    if not mon.empty:
        m = mon.copy()
        for c in ("kWh_th","kWh_el"): m[c]=m[c].round(0)
        if "Betriebsstunden_Vent" in m: m["Betriebsstunden_Vent"]=m["Betriebsstunden_Vent"].round(1)
        if "Stunden_AUS" in m: m["Stunden_AUS"]=m["Stunden_AUS"].round(1)
        data=[["Jahr","Monat","Wärme [kWh]","Strom Vent. [kWh]","Betriebsstd. Vent.","Stunden AUS"],
              *m.reindex(columns=["year","month","kWh_th","kWh_el","Betriebsstunden_Vent","Stunden_AUS"]).astype({"year":int,"month":int}).values.tolist()]
        T4=Table(data,hAlign="LEFT", repeatRows=1)
        T4.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
                                ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                ("ALIGN",(2,1),(-1,-1),"RIGHT")]))
        story+=[P("Ergebnisse – Monate (Detail)", H2), T4]

    doc.build(story, onFirstPage=_pdf_header_footer, onLaterPages=_pdf_header_footer)
    out.seek(0); return out.read()

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Heizenergie – Lüftungsanlagen (ISO 50001)", layout="wide")
st.title("Heizenergie – Lüftungsanlagen (ISO 50001)")

# Session-Init
defaults = [
    ("try_df", None), ("try_info",""), ("years", []), ("interp", False), ("try_hints", []),
    ("anlage", None), ("defs", None), ("wochenplan", None),
    ("aus_bloecke", []),
    ("mon_df", None), ("jahr_df", None), ("prot_df", None), ("prot_full_df", None),
    ("mon_ue_df", None), ("jahr_ue_df", None),
    ("auto_calc", False), ("trigger", False),
]
for k, v in defaults:
    if k not in st.session_state: st.session_state[k] = v

# --- Infofenster ---
with st.expander("ℹ️ Erläuterung – Rechenschritte & Formeln", expanded=False):
    st.markdown("""
**Zeitlogik (Priorität):**
1) **AUS-Kalender** (Betriebsferien/Wartung) → ganze Stunde AUS  
2) **Tages-Schalter** „Anlage AUS“ (00:00–24:00)  
3) **Fenster** Normal/Absenk/AUS (pro Zeitraum)

**Heizfall & WRG:** ΔT = max(0, T_soll − T_out) → ΔT_eff = (1 − η_t) · ΔT  
**Wärme je Stunde:** Q = 0,00034 · V(m³/h) · ΔT_eff · Anteil_h  
**Ventilator:** P_fan = SFP · (V/3600) **oder** fan_kW · (V/V_nominal) → E = P_fan · Anteil_h
""")

# ----- 1) TRY laden -----
with st.form("form_try"):
    st.subheader("1) TRY-CSV laden")
    f = st.file_uploader("TRY-CSV (stündlich: Datum/Zeit + Außentemperatur)", type=["csv"])
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
                st.success("TRY übernommen."); st.text(st.session_state["try_info"])
                for h in hints: st.info(h)
                if st.session_state["auto_calc"]:
                    st.session_state["trigger"] = True

# ----- 2) Anlagendaten -----
with st.form("form_anlage"):
    st.subheader("2) Anlagendaten")
    c1 = st.columns([1.2,1,1,1])
    a_id   = c1[0].text_input("Anlagen-ID", value="A01")
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

# ----- 3) Standardwerte & Zeiten -----
st.subheader("3) Standardwerte & Betriebs-/Absenkzeiten")
c = st.columns(4)
Tn = c[0].number_input("Soll-Zuluft NORMAL [°C]", value=20.0, step=0.5, key="Tn")
Ta = c[1].number_input("Soll-Zuluft ABSENK [°C]", value=17.0, step=0.5, key="Ta")
Vn = c[2].number_input("Volumenstrom NORMAL [m³/h]", value=5000.0, min_value=500.0, max_value=500000.0, step=100.0, key="Vn")
Va = c[3].number_input("Volumenstrom ABSENK [m³/h] (0 = wie normal)", value=2000.0, min_value=0.0, max_value=500000.0, step=100.0, key="Va")
defs = Defaults(float(Tn), float(Ta), float(Vn), (None if Va==0.0 else float(Va)))

# Wochenplan initialisieren (24/7 Default) und Schaltfläche zum Übernehmen der Kopfwerte
if st.session_state["wochenplan"] is None:
    st.session_state["wochenplan"] = wochenplan_24x7(defs)

if st.button("Kopfwerte in Kalender übernehmen (24/7 Normal für alle Tage)"):
    st.session_state["wochenplan"] = wochenplan_24x7(defs)
    st.success("Kalender aktualisiert (24/7 Normalbetrieb mit aktuellen Kopfwerten).")

wp: List[Tagesplan] = st.session_state["wochenplan"]

st.caption("Voreinstellung: **24/7 Normalbetrieb** auf allen Tagen. "
           "Fenster können ergänzt werden (Absenk/AUS). Über-Mitternacht wird korrekt geteilt. "
           "Tages-Schalter setzt einen Tag komplett AUS.")

def render_fenster(tag_index: int, liste: List[Zeitfenster], modus_label: str):
    st.markdown(modus_label)
    add_key = f"add_{modus_label}_{tag_index}"
    if st.button(f"+ Fenster {modus_label} {WOCHENTAGE[tag_index]}", key=add_key):
        # Absenk-Defaults nutzen, wenn Absenk; sonst Normal
        if "Absenk" in modus_label:
            liste.append(Zeitfenster("17:00","06:30", True, defs.T_absenk_C, defs.V_absenk_m3h))
        else:
            liste.append(Zeitfenster("08:00","16:00", True, defs.T_normal_C, None))
        st.rerun()

    delete_index = None
    for i, f in enumerate(liste):
        cols = st.columns([1,1,1,1,1,1,0.6])  # Start|Ende|Aktiv|T_soll|Vol-Mode|(Wert)|Löschen
        f.start = cols[0].text_input("Start", key=f"{modus_label}_start_{tag_index}_{i}", value=f.start)
        f.ende  = cols[1].text_input("Ende",  key=f"{modus_label}_ende_{tag_index}_{i}", value=f.ende)
        f.aktiv = cols[2].checkbox("Fenster verwenden", key=f"{modus_label}_aktiv_{tag_index}_{i}", value=f.aktiv)
        f.T_soll_C = cols[3].number_input("T_soll [°C]", key=f"{modus_label}_Tsoll_{tag_index}_{i}", value=float(f.T_soll_C), step=0.5)

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
            cols[5].markdown("—"); f.V_m3h = 0.0
        else:
            cols[5].markdown("—"); f.V_m3h = None

        if cols[6].button("–", key=f"{modus_label}_del_{tag_index}_{i}"):
            delete_index = i

    if delete_index is not None:
        liste.pop(delete_index); st.rerun()

# --- Tagweise Editor + Tages-AUS ---
for d in range(7):
    st.markdown(f"### {WOCHENTAGE[d]}")
    day = wp[d]
    day.day_off = st.checkbox("Anlage an diesem Tag AUS (00:00–24:00)", key=f"day_off_{d}", value=day.day_off)
    if day.day_off:
        st.info("Dieser Tag ist komplett AUS. Einzelne Fenster werden ignoriert.")
        if st.button(f"Alle Fenster von {WOCHENTAGE[d]} löschen", key=f"clear_day_{d}"):
            day.normal.clear(); day.absenk.clear(); st.rerun()
    else:
        render_fenster(d, day.normal, "Normalbetrieb")
        render_fenster(d, day.absenk, "Absenkbetrieb")

# ----- 3b) AUS-Kalender -----
st.subheader("3b) Betriebsferien / Wartungsblöcke (AUS-Kalender)")
with st.expander("AUS-Blöcke verwalten", expanded=False):
    c0 = st.columns(4)
    start_d = c0[0].date_input("Start-Datum", value=date.today())
    start_t = c0[1].time_input("Start-Uhrzeit", value=time(0,0))
    end_d   = c0[2].date_input("Ende-Datum (exkl.)", value=date.today())
    end_t   = c0[3].time_input("Ende-Uhrzeit (exkl.)", value=time(0,0))
    c1 = st.columns([1,1,2])
    if c1[0].button("AUS-Block hinzufügen"):
        start_dt = datetime.combine(start_d, start_t)
        end_dt   = datetime.combine(end_d, end_t)
        if end_dt <= start_dt:
            st.error("Ende muss nach Start liegen (Ende ist exklusiv).")
        else:
            st.session_state["aus_bloecke"].append(AusBlock(start=start_dt, ende=end_dt))
            st.session_state["aus_bloecke"] = merge_aus_blocks(st.session_state["aus_bloecke"])
            st.success("AUS-Block hinzugefügt."); st.rerun()
    if c1[1].button("Alle AUS-Blöcke löschen"):
        st.session_state["aus_bloecke"].clear(); st.rerun()

    aus_list = merge_aus_blocks(st.session_state["aus_bloecke"])
    if aus_list:
        rows = []
        for idx, b in enumerate(aus_list):
            rows.append({"#": idx+1, "Start": b.start, "Ende (exkl.)": b.ende, "Dauer [h]": round(block_duration_hours(b),1)})
        df_aus = pd.DataFrame(rows)
        st.dataframe(df_aus, use_container_width=True)
        del_idx = st.number_input("AUS-Block Nr. löschen (0 = keiner)", min_value=0, max_value=len(aus_list), value=0, step=1)
        if del_idx and st.button("Ausgewählten AUS-Block löschen"):
            del st.session_state["aus_bloecke"][del_idx-1]
            st.session_state["aus_bloecke"] = merge_aus_blocks(st.session_state["aus_bloecke"]); st.rerun()
    else:
        st.caption("Keine AUS-Blöcke hinterlegt.")

# ----- 4) Berechnen -----
def _rechnen():
    df = st.session_state["try_df"]; anl = st.session_state["anlage"]
    if df is None or df.empty: st.error("Bitte zuerst TRY übernehmen."); return
    if anl is None: st.error("Bitte zuerst Anlagendaten übernehmen."); return
    anl = Anlage(anl.id, anl.name, anl.V_nominal_m3h, anl.anzahl, anl.wrg, anl.eta_t, anl.fan_kW, anl.SFP_kW_per_m3s, st.session_state["wochenplan"])
    mon, jahr, prot, prot_full = berechne_detail(df, anl, defs, st.session_state["aus_bloecke"])
    mon_u, jahr_u = berechne_ueberschlag(df, anl, defs, st.session_state["aus_bloecke"])
    st.session_state["mon_df"], st.session_state["jahr_df"] = mon, jahr
    st.session_state["prot_df"], st.session_state["prot_full_df"] = prot, prot_full
    st.session_state["mon_ue_df"], st.session_state["jahr_ue_df"] = mon_u, jahr_u
    st.success("Berechnung abgeschlossen.")

cols_run = st.columns([1,3])
if cols_run[0].button("Berechnen", type="primary"): _rechnen()
if st.session_state["trigger"]:
    _rechnen(); st.session_state["trigger"] = False

# ----- Ergebnisse + Downloads -----
mon = st.session_state["mon_df"]; jahr = st.session_state["jahr_df"]
prot = st.session_state["prot_df"]; prot_full = st.session_state["prot_full_df"]
mon_u = st.session_state["mon_ue_df"]; jahr_u = st.session_state["jahr_ue_df"]

if mon is not None and jahr is not None:
    # Anzeige runden (kWh ganzzahlig; Stunden 1 Dezimal)
    m_show = mon.copy(); j_show = jahr.copy()
    for c in ("kWh_th","kWh_el"):
        if c in m_show: m_show[c] = m_show[c].round(0)
        if c in j_show: j_show[c] = j_show[c].round(0)
    for c in ("Betriebsstunden_Vent","Stunden_AUS"):
        if c in m_show: m_show[c] = m_show[c].round(1)
        if c in j_show: j_show[c] = j_show[c].round(1)

    m_de = m_show.rename(columns={
        "year":"Jahr", "month":"Monat",
        "kWh_th":"Wärme [kWh]", "kWh_el":"Strom Vent. [kWh]",
        "Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"
    })
    j_de = j_show.rename(columns={
        "year":"Jahr",
        "kWh_th":"Wärme [kWh]", "kWh_el":"Strom Vent. [kWh]",
        "Betriebsstunden_Vent":"Betriebsstd. Vent.","Stunden_AUS":"Stunden AUS"
    })

    # Summenzeile nur bei Monaten
    m_de = add_sum_row(m_de, label_col="Monat", label="Summe")

    st.subheader("Ergebnisse – Jahreswerte (Detail)")
    st.dataframe(j_de, use_container_width=True)

    st.subheader("Ergebnisse – Monate (Detail)")
    st.dataframe(m_de, use_container_width=True)

    # Überschlag
    if mon_u is not None and not mon_u.empty:
        mu = mon_u.copy(); ju = jahr_u.copy()
        mu[["kWh_th","kWh_el"]] = mu[["kWh_th","kWh_el"]].round(0)
        ju[["kWh_th","kWh_el"]] = ju[["kWh_th","kWh_el"]].round(0)

        mu_de = mu.rename(columns={"year":"Jahr","month":"Monat","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        ju_de = ju.rename(columns={"year":"Jahr","kWh_th":"Wärme [kWh]","kWh_el":"Strom Vent. [kWh]"})
        mu_de = add_sum_row(mu_de, label_col="Monat", label="Summe")  # nur bei Monaten

        st.subheader("Kontrollrechner – Monate (Überschlag)")
        st.dataframe(mu_de, use_container_width=True)

        # Abweichung in %
        cmp = m_show.merge(mu, on=["year","month"], suffixes=("_det","_ue"))
        cmp["Abw. Wärme [%]"] = (cmp["kWh_th_det"]-cmp["kWh_th_ue"]) / cmp["kWh_th_det"].replace(0, pd.NA) * 100
        cmp["Abw. Strom [%]"] = (cmp["kWh_el_det"]-cmp["kWh_el_ue"]) / cmp["kWh_el_det"].replace(0, pd.NA) * 100
        cmp_out = cmp[["year","month","Abw. Wärme [%]","Abw. Strom [%]"]].round(1).rename(columns={"year":"Jahr","month":"Monat"})
        st.caption("Abweichung Überschlag vs. Detail (positiv = Überschlag kleiner).")
        st.dataframe(cmp_out, use_container_width=True)

    # Rechen-Protokoll (Ausschnitt)
    st.subheader("Rechen-Protokoll (Ausschnitt)")
    if prot is not None and not prot.empty:
        st.dataframe(prot.head(200), use_container_width=True)
        prot_full_xlsx = xlsx_single_sheet(prot_full, "Protokoll") if prot_full is not None and not prot_full.empty else b""
    else:
        st.info("Kein Protokoll verfügbar.")
        prot_full_xlsx = b""

    # Plausibilitäts-Checks (grobe Heuristik)
    if not j_show.empty:
        th = float(j_show["kWh_th"].sum()); el = float(j_show["kWh_el"].sum())
        hints=[]
        if th < 0: hints.append("Wärmebedarf < 0 kWh – Eingaben prüfen.")
        if el < 0: hints.append("Strom < 0 kWh – Eingaben prüfen.")
        if el > th*0.5 and th > 0: hints.append("Ventilatorstrom sehr hoch im Verhältnis zur Wärme → SFP/Fan-Werte prüfen.")
        for h in hints: st.warning(h)

    # Downloads (Excel/PDF)
    st.subheader("Downloads")
    c1, c2, c3, c4 = st.columns(4)
    c1.download_button(
        "Excel – Auswertung (Detail + Überschlag + Protokoll + AUS-Kalender)",
        xlsx_export(m_show, j_show, prot if prot is not None else pd.DataFrame(),
                    mon_u if mon_u is not None else pd.DataFrame(),
                    jahr_u if jahr_u is not None else pd.DataFrame(),
                    st.session_state["aus_bloecke"]),
        file_name="Heizenergie_Auswertung.xlsx"
    )
    c2.download_button(
        "Excel – Monate (Detail, DE)",
        xlsx_single_sheet(m_de, "Monate (Detail, DE)"),
        file_name="Heizenergie_Monate_DE.xlsx"
    )
    if prot_full_xlsx:
        c3.download_button(
            "Excel – Protokoll (vollständig)",
            prot_full_xlsx,
            file_name="Heizenergie_Protokoll_Stunden.xlsx"
        )
    if REPORTLAB_OK:
        anl_obj = st.session_state["anlage"]
        if anl_obj is not None:
            c4.download_button("PDF – ISO 50001 Kurzbericht",
                               pdf_export(st.session_state.get("try_info",""), defs, anl_obj,
                                          m_show, j_show, st.session_state["aus_bloecke"], st.session_state["try_hints"]),
                               file_name="ISO50001_Heizenergiebericht.pdf", mime="application/pdf")
    else:
        st.info("PDF-Export nicht verfügbar (ReportLab nicht installiert).")
