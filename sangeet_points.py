#!/usr/bin/env python3
"""
Sangeet Points Calculator — revised logic (GBM is part of D, not extra)

Usage:
  python sangeet_points.py "/path/to/2025-2026 Attendance Sheet.xlsx"

Scoring logic:
- Let D = total number of standing meetings a member is expected to attend weekly:
      D = (# of distinct subgroups the member is in) + (1 if a GBM sheet exists)
  (GBM is INCLUDED in D and is NOT an extra point source on its own.)
- Each ATTENDED meeting (GBM or any subgroup) earns 1/D points.
  ATTENDED = Present, Tardy, Tardy (Excused), Excessive Tardy. Absences = 0.
- Tabling: +1 point per tabling count.
- Gigs:    +2 points per gig/performance count.
- Output: final sheet **Points Standing** with only two columns: Member | Total_Points
"""

import argparse
import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# ------------------------- Normalization helpers ------------------------- #

def _is_nan(x) -> bool:
    try:
        return pd.isna(x)
    except Exception:
        return False

def normalize_status(val) -> str:
    if _is_nan(val):
        return ""
    s = str(val).strip().lower()
    mapping = {
        "present": "present", "p": "present", "attended": "present",
        "y": "present", "yes": "present", "here": "present",
        "a": "absent", "absent": "absent", "n": "absent", "no": "absent",
        "excused absence": "excused_absence", "excused": "excused_absence",
        "unexcused absence": "unexcused_absence",
        "tardy (excused)": "tardy_excused", "tardy(excused)": "tardy_excused", "tardy - excused": "tardy_excused",
        "tardy": "tardy", "late": "tardy",
        "excessive tardy": "excessive_tardy", "excessive tardiness": "excessive_tardy",
    }
    if s in mapping:
        return mapping[s]
    if "excessive" in s and "tard" in s: return "excessive_tardy"
    if "tard" in s and "excuse" in s:    return "tardy_excused"
    if "tard" in s or "late" in s:       return "tardy"
    if "unexcused" in s and "abs" in s:  return "unexcused_absence"
    if "excused" in s and "abs" in s:    return "excused_absence"
    if "abs" in s:                       return "absent"
    if "pres" in s or "attend" in s:     return "present"
    return s


def classify_sheet(name: str) -> str:
    """
    Returns one of: 'ignore', 'gbm', 'tabling', 'gigs', 'subgroup'
    """
    s = (name or "").strip().lower()
    if "points standing" in s:  return "ignore"
    if "gbm" in s or "general body" in s: return "gbm"
    if "tabling" in s:          return "tabling"
    if "gig" in s or "performance" in s:  return "gigs"
    return "subgroup"


# ------------------------- DataFrame helpers ------------------------- #

def _stringify_headers(df: pd.DataFrame) -> pd.DataFrame:
    new_cols, seen = [], set()
    for i, c in enumerate(df.columns):
        col = f"col_{i}" if _is_nan(c) or str(c).strip() == "" else str(c)
        base, k = col, 1
        while col in seen:
            col = f"{base}.{k}"; k += 1
        seen.add(col); new_cols.append(col)
    df.columns = new_cols
    return df


def detect_wide_format(df: pd.DataFrame) -> bool:
    if df.empty or df.shape[1] < 2:
        return False
    df = _stringify_headers(df.copy())
    first = df.columns[0].strip().lower()
    looks_like_name = any(k in first for k in ["name", "member", "person", "attendee"])
    date_like, status_like = 0, 0
    for c in df.columns[1:]:
        cstr = c.strip().lower()
        if re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", cstr) or re.search(r"\d{4}-\d{1,2}-\d{1,2}", cstr) \
           or any(k in cstr for k in ["week", "mtg", "rehearsal", "meeting"]):
            date_like += 1
        vals = df[c].dropna().astype(str).str.lower().head(20).tolist()
        if vals:
            hits = sum(normalize_status(v) in {
                "present","absent","excused_absence","unexcused_absence","tardy","tardy_excused","excessive_tardy"
            } for v in vals)
            if hits / max(1, len(vals)) >= 0.3:
                status_like += 1
    return looks_like_name and (date_like >= 1 or status_like >= 2)


def melt_wide(df: pd.DataFrame) -> pd.DataFrame:
    df = _stringify_headers(df.copy())
    id_col = df.columns[0]
    out = df.melt(id_vars=[id_col], var_name="Meeting", value_name="Status")
    return out.rename(columns={id_col: "Member"})


def _guess_member_and_status_cols(df: pd.DataFrame) -> Tuple[str, str]:
    df = _stringify_headers(df.copy())
    cols_lower = {c: c.strip().lower() for c in df.columns}
    member_col = next((c for c, lc in cols_lower.items() if any(k in lc for k in ["member","name","person","attendee"])), None)
    status_col = next((c for c, lc in cols_lower.items() if any(k in lc for k in ["status","attendance"])), None)
    if member_col is None: member_col = df.columns[0]
    if status_col is None:
        best_col, best_score = None, -1.0
        for c in df.columns:
            if c == member_col: continue
            vals = df[c].dropna().astype(str).str.lower().tolist()
            if not vals: continue
            hits = sum(normalize_status(v) in {
                "present","absent","excused_absence","unexcused_absence","tardy","tardy_excused","excessive_tardy"
            } for v in vals)
            score = hits / max(1, len(vals))
            if score > best_score: best_score, best_col = score, c
        status_col = best_col if best_col is not None else df.columns[-1]
    return member_col, status_col


def normalize_attendance_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Member","Status"])
    df = df.dropna(how="all").dropna(axis=1, how="all")
    if df.empty:
        return pd.DataFrame(columns=["Member","Status"])
    if detect_wide_format(df):
        out = melt_wide(df)
    else:
        mcol, scol = _guess_member_and_status_cols(df)
        df = _stringify_headers(df)
        if mcol not in df.columns: mcol = df.columns[0]
        if scol not in df.columns: scol = df.columns[-1]
        out = df.rename(columns={mcol:"Member", scol:"Status"})[["Member","Status"]]
    out["Member"] = out["Member"].astype(str).str.strip()
    out["Status"] = out["Status"].apply(normalize_status)
    return out[ out["Member"].str.len() > 0 ]


def status_counts(df: pd.DataFrame, member: str) -> Dict[str, int]:
    out = {k:0 for k in ["present","absent","excused_absence","unexcused_absence","tardy","tardy_excused","excessive_tardy"]}
    if df.empty: return out
    sub = df[df["Member"] == member]
    for s in sub["Status"].tolist():
        ns = normalize_status(s)
        if ns in out: out[ns] += 1
        elif ns and ns not in {"absent","excused_absence","unexcused_absence",""}:
            out["present"] += 1
    return out


def sum_numeric_row_values(df: pd.DataFrame) -> pd.DataFrame:
    """
    For Tabling/Gigs sheets: sum all numeric columns per row -> Count.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["Member","Count"])
    df = _stringify_headers(df.dropna(how="all").dropna(axis=1, how="all"))
    if df.empty:
        return pd.DataFrame(columns=["Member","Count"])
    # Guess a member-like column
    mcol, _ = _guess_member_and_status_cols(df)
    if mcol not in df.columns: mcol = df.columns[0]
    num_df = df.drop(columns=[mcol], errors="ignore").apply(pd.to_numeric, errors="coerce")
    counts = num_df.sum(axis=1, skipna=True).fillna(0.0)
    out = pd.DataFrame({"Member": df[mcol].astype(str).str.strip(), "Count": counts})
    return out[out["Member"].str.len() > 0]


# ------------------------- Main compute ------------------------- #

def compute_points(workbook_path: str) -> None:
    wb = load_workbook(workbook_path)
    sheet_names = wb.sheetnames

    gbm_frames: List[pd.DataFrame] = []
    subgroup_frames_by_name: Dict[str, pd.DataFrame] = {}
    tabling_frames: List[pd.DataFrame] = []
    gigs_frames: List[pd.DataFrame] = []

    for sname in sheet_names:
        kind = classify_sheet(sname)
        if kind == "ignore":
            continue

        ws = wb[sname]
        values = list(ws.values)
        if not values:
            continue
        header = list(values[0]) if values else []
        data = values[1:] if len(values) > 1 else []
        df = pd.DataFrame(data, columns=header).dropna(how="all")
        if df.empty:
            continue

        if kind == "tabling":
            tabling_frames.append(sum_numeric_row_values(df))
            continue
        if kind == "gigs":
            gigs_frames.append(sum_numeric_row_values(df))
            continue

        att = normalize_attendance_frame(df)
        att["Sheet"] = sname
        if kind == "gbm":
            gbm_frames.append(att)
        else:
            subgroup_frames_by_name[sname] = att

    gbm_att = pd.concat(gbm_frames, ignore_index=True) if gbm_frames else pd.DataFrame(columns=["Member","Status","Sheet"])
    subgroup_att = pd.concat(list(subgroup_frames_by_name.values()), ignore_index=True) if subgroup_frames_by_name else pd.DataFrame(columns=["Member","Status","Sheet"])
    tabling_counts = pd.concat(tabling_frames, ignore_index=True) if tabling_frames else pd.DataFrame(columns=["Member","Count"])
    gigs_counts = pd.concat(gigs_frames, ignore_index=True) if gigs_frames else pd.DataFrame(columns=["Member","Count"])

    # Member universe
    members = set()
    for frame in (gbm_att, subgroup_att, tabling_counts, gigs_counts):
        if not frame.empty and "Member" in frame.columns:
            members.update([m for m in frame["Member"].tolist() if isinstance(m, str) and m.strip()])

    # Infer #subgroups per member = count of distinct subgroup SHEETS they appear on
    member_to_subgroups: Dict[str,int] = {m:0 for m in members}
    for sname, att in subgroup_frames_by_name.items():
        if att.empty: continue
        for m in set(att["Member"]):
            member_to_subgroups[m] = member_to_subgroups.get(m, 0) + 1

    gbm_exists = not gbm_att.empty  # GBM contributes to D but has no special extra points

    # Compute totals
    rows = []
    for m in sorted(members):
        n_subgroups = member_to_subgroups.get(m, 0)
        # D = #subgroups + 1 if GBM sheet exists (GBM is included in D, not extra)
        D = n_subgroups + (1 if gbm_exists else 0)
        per_meeting = 0.0 if D == 0 else 1.0 / float(D)

        total_attendance_points = 0.0

        # Count ATTENDED across GBM + Subgroups; each attended meeting is worth 1/D
        if gbm_exists:
            g = status_counts(gbm_att, m)
            gbm_attended = g["present"] + g["tardy"] + g["tardy_excused"] + g["excessive_tardy"]
            total_attendance_points += gbm_attended * per_meeting

        if not subgroup_att.empty:
            s = status_counts(subgroup_att, m)
            sg_attended = s["present"] + s["tardy"] + s["tardy_excused"] + s["excessive_tardy"]
            total_attendance_points += sg_attended * per_meeting

        # Tabling: +1 each
        tab_points = float(tabling_counts.loc[tabling_counts["Member"] == m, "Count"].sum()) if not tabling_counts.empty else 0.0
        # Gigs: +2 each
        gig_points = 2.0 * float(gigs_counts.loc[gigs_counts["Member"] == m, "Count"].sum()) if not gigs_counts.empty else 0.0

        total_points = total_attendance_points + tab_points + gig_points
        rows.append({"Member": m, "Total_Points": total_points})

    summary = pd.DataFrame(rows).sort_values(["Total_Points","Member"], ascending=[False, True])

    # Write final sheet
    for s in [s for s in wb.sheetnames if s.strip().lower() == "points standing"]:
        del wb[s]
    ws = wb.create_sheet("Points Standing")
    for r in dataframe_to_rows(summary[["Member","Total_Points"]], index=False, header=True):
        ws.append(r)
    wb.save(workbook_path)


# ------------------------- CLI ------------------------- #

def main():
    ap = argparse.ArgumentParser(description="Compute total points (attendance with GBM included in D, tabling, gigs) into 'Points Standing'.")
    ap.add_argument("xlsx_path", help="Path to the attendance .xlsx workbook")
    args = ap.parse_args()
    compute_points(args.xlsx_path)

if __name__ == "__main__":
    main()
