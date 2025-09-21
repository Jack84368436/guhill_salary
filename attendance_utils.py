# -*- coding: utf-8 -*-
import re
import pandas as pd

TOTAL_MIN_COL = "總分鐘數"

def to_hhmm(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isnull(val):
            return ""
    except Exception:
        pass
    try:
        hour = getattr(val, "hour", None)
        minute = getattr(val, "minute", None)
        if hour is not None and minute is not None:
            if hour == 0 and minute == 0:
                return ""
            return f"{hour:02d}:{minute:02d}"
    except Exception:
        pass
    s = str(val).strip()
    if not s:
        return ""
    if " " in s:
        s = s.split()[-1]
    if len(s) >= 5 and s[2] == ":":
        s = s[:5]
    return "" if s in ("00:00", "0:00") else s

def fmt_date_label(val) -> str:
    s = "" if val is None else str(val).strip()
    if not s:
        return ""
    try:
        d = pd.to_datetime(s)
        return f"{d.month}/{d.day}"
    except Exception:
        return s

def _iter_day_groups(df_raw, df, row_idx: int, start_col: int, date_row_index: int, group_stride: int):
    ncols = len(df.columns)
    col = start_col
    while col < ncols:
        raw_date = df_raw.iloc[date_row_index, col]
        date_str = fmt_date_label(raw_date)
        in_val  = df.iloc[row_idx, col]
        out_val = df.iloc[row_idx, col + 1] if col + 1 < ncols else None
        min_val = df.iloc[row_idx, col + 2] if col + 2 < ncols else None
        yield date_str, in_val, out_val, min_val
        col += group_stride

def list_employee_days(df_raw, df, target_name: str, scan_rows: int = 30, start_col: int = 1, date_row_index: int = 1, group_stride: int = 3):
    a_col = df.iloc[:scan_rows, 0].astype(str).str.strip()
    match_idx = a_col[a_col == target_name].index
    if len(match_idx) == 0:
        return None, []
    row_idx = match_idx[0]

    total_min = None
    if TOTAL_MIN_COL in df.columns:
        try:
            total_min = int(float(df.loc[row_idx, TOTAL_MIN_COL]))
        except Exception:
            pass

    outputs = []
    for date_str, in_val, out_val, min_val in _iter_day_groups(df_raw, df, row_idx, start_col, date_row_index, group_stride):
        try:
            has_min = (min_val is not None) and (str(min_val).strip() not in ("", "0")) and (float(min_val) > 0)
        except Exception:
            has_min = False
        if has_min:
            in_txt  = to_hhmm(in_val)
            out_txt = to_hhmm(out_val)
            if in_txt or out_txt:
                try:
                    m = int(float(min_val))
                except Exception:
                    m = str(min_val)
                outputs.append(f"{date_str} {in_txt}--{out_txt} 共{m}分鐘")
    return total_min, outputs

def extract_employee_records(df_raw, df, target_name: str, scan_rows: int = 30, start_col: int = 1, date_row_index: int = 1, group_stride: int = 3):
    a_col = df.iloc[:scan_rows, 0].astype(str).str.strip()
    match_idx = a_col[a_col == target_name].index
    if len(match_idx) == 0:
        return None, []
    row_idx = match_idx[0]

    total_min = None
    if TOTAL_MIN_COL in df.columns:
        try:
            total_min = int(float(df.loc[row_idx, TOTAL_MIN_COL]))
        except Exception:
            pass

    records = []
    for date_str, in_val, out_val, min_val in _iter_day_groups(df_raw, df, row_idx, start_col, date_row_index, group_stride):
        try:
            has_min = (min_val is not None) and (str(min_val).strip() not in ("", "0")) and (float(min_val) > 0)
        except Exception:
            has_min = False
        if has_min:
            try:
                m = int(float(min_val))
            except Exception:
                m = None
            rec = {"date": date_str, "in": to_hhmm(in_val), "out": to_hhmm(out_val), "minutes": m if m is not None else str(min_val)}
            if rec["in"] or rec["out"]:
                records.append(rec)
    return total_min, records

# ── 新增：把 "NT$1,234"、"1,234"、數字或 NaN 轉成整數 ──
def parse_ntd(val) -> int:
    if val is None:
        return 0
    try:
        if pd.isnull(val):
            return 0
    except Exception:
        pass
    if isinstance(val, (int, float)):
        try:
            return int(round(val))
        except Exception:
            return 0
    s = str(val)
    # 去掉所有非數字/負號字元（含 NT$、逗號、空白、引號等）
    s = re.sub(r"[^0-9\-]", "", s)
    if s in ("", "-", "--"):
        return 0
    try:
        return int(s)
    except Exception:
        return 0

def fmt_ntd(n: int) -> str:
    return f"NT${n:,}"

# ── 新增：依姓名抓薪資九項，回傳 {欄名:int,...}（會過濾 0） ──
def extract_pay_items(df_summary: pd.DataFrame, target_name: str, field_names: list[str]) -> dict:
    """df_summary：工作表3（header=0）；第一欄為姓名。"""
    if df_summary is None or df_summary.empty:
        return {}
    a_col = df_summary.iloc[:, 0].astype(str).str.strip()
    idx = a_col[a_col == target_name].index
    if len(idx) == 0:
        return {}
    r = idx[0]

    result = {}
    for col in field_names:
        if col not in df_summary.columns:
            continue
        val = parse_ntd(df_summary.loc[r, col])
        if val != 0:
            result[col] = val
    return result

