# -*- coding: utf-8 -*-
import math
import re
import unicodedata
import difflib
import pandas as pd

from attendance_utils import (
    list_employee_days,
    extract_pay_items,
    fmt_ntd,
    parse_ntd,
)

from config import (
    # 檔案/分頁名稱
    FILE_PATH, SHEET_ATTEND, SHEET_BONUS, SHEET_SUMMARY,
    # 出勤解析規則
    START_COL, DATE_ROW_INDEX, GROUP_STRIDE,
    # CLI 互動
    SCAN_ROWS_DEFAULT,
    # 摘要九項
    SUMMARY_FIELDS, SUMMARY_NAME_COL_INDEX,
    # 獎金欄位設定
    BONUS_FIELD, BONUS_COL_INDEX,
    # << 新增：是否使用索引與三分頁索引 >>
    USE_SHEET_INDEX,
    SHEET_ATTEND_INDEX, SHEET_BONUS_INDEX, SHEET_SUMMARY_INDEX,
)

# 想關閉偵錯輸出改為 False
VERBOSE = False


# ====== 工具：姓名正規化（解決全形空白/零寬字元/Unicode 變體） ======
def _norm_name(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u00A0", "").replace("\u3000", "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("\u200b", "").replace("\ufeff", "")
    return s.strip()

def _series_norm(sr: pd.Series) -> pd.Series:
    return sr.astype(str).map(_norm_name)


# ====== 工具：依設定解析要用的工作表名稱（索引優先） ======
def resolve_sheet_name(xls: pd.ExcelFile, prefer_name: str, index: int | None,
                       use_index: bool, fallback_index: int | None = None) -> str:
    """
    - use_index=True：直接用 index -> xls.sheet_names[index]
    - use_index=False：先精確比對名稱，不行做去空白匹配，最後可回退 fallback_index
    """
    if use_index:
        if index is None:
            raise ValueError("use_index=True 但 index 未設定")
        if index < 0 or index >= len(xls.sheet_names):
            raise IndexError(f"索引 {index} 超出範圍（共有 {len(xls.sheet_names)} 張）")
        return xls.sheet_names[index]

    # 名稱模式
    if prefer_name in xls.sheet_names:
        return prefer_name
    norm = lambda s: str(s).strip().replace(" ", "")
    cand = [s for s in xls.sheet_names if norm(s) == norm(prefer_name)]
    if cand:
        return cand[0]
    if fallback_index is not None and 0 <= fallback_index < len(xls.sheet_names):
        return xls.sheet_names[fallback_index]
    raise ValueError(f"找不到工作表：{prefer_name}（且無可用回退）")


# ====== 讀取各分頁（摘要、獎金） ======
def _read_summary_sheet(verbose: bool = False) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(FILE_PATH)
        sheet_name = resolve_sheet_name(
            xls, SHEET_SUMMARY, SHEET_SUMMARY_INDEX, USE_SHEET_INDEX, fallback_index=2
        )

        def _load_with_header(h):
            df = pd.read_excel(xls, sheet_name=sheet_name, header=h)
            df.columns = df.columns.map(lambda s: str(s).strip())
            try:
                df.iloc[:, SUMMARY_NAME_COL_INDEX] = (
                    df.iloc[:, SUMMARY_NAME_COL_INDEX].astype(str).str.strip()
                )
            except Exception:
                pass
            if False in df.columns:  # 偶發布林欄
                df = df.drop(columns=[False])
            return df

        df_sum = _load_with_header(1)
        if not any(col in df_sum.columns for col in SUMMARY_FIELDS):
            df_sum = _load_with_header(0)

        if verbose:
            print("[debug] 使用摘要分頁：", sheet_name)
            print("[debug] 摘要表欄名：", list(df_sum.columns))
        return df_sum

    except Exception as e:
        if verbose:
            print(f"[debug] 讀取摘要表失敗：{e}")
        return None


def _read_bonus_sheet(verbose: bool = False) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(FILE_PATH)
        sheet_name = resolve_sheet_name(
            xls, SHEET_BONUS, SHEET_BONUS_INDEX, USE_SHEET_INDEX, fallback_index=1
        )

        def _load(h):
            df = pd.read_excel(xls, sheet_name=sheet_name, header=h)
            df.columns = df.columns.map(lambda s: str(s).strip())
            if False in df.columns:
                df = df.drop(columns=[False])
            return df

        df_bonus = _load(1)
        if df_bonus.shape[1] <= BONUS_COL_INDEX:
            df_bonus = _load(0)

        if verbose:
            print("[debug] 使用獎金分頁：", sheet_name, "；shape=", df_bonus.shape)
        return df_bonus

    except Exception as e:
        if verbose:
            print(f"[debug] 讀取獎金表失敗：{e}")
        return None


# ====== 尋找姓名欄 / 獎金欄 ======
def _guess_name_col_in_bonus(df_bonus: pd.DataFrame, att_names: set[str]) -> int | None:
    att_norm = {_norm_name(x) for x in att_names if x is not None}
    best_i, best_hit = None, -1
    for i in range(len(df_bonus.columns)):
        col_vals = _series_norm(df_bonus.iloc[:, i])
        hit = col_vals.isin(att_norm).sum()
        if hit > best_hit:
            best_i, best_hit = i, hit
    if VERBOSE:
        print(f"[debug] bonus name col -> {best_i}, hits={best_hit}")
    return best_i if best_hit > 0 else None

def _find_bonus_col(df: pd.DataFrame) -> int | None:
    keywords = ["獎金總和", "獎金", "bonus", "緊急"]

    # 1) 欄名直接找
    cols_norm = [str(c).strip().lower() for c in df.columns]
    for k in keywords:
        k = k.lower()
        for idx, cname in enumerate(cols_norm):
            if k in cname and cname != "false":
                if VERBOSE:
                    print(f"[debug] bonus col found in columns -> {idx} ({df.columns[idx]!r})")
                return idx

    # 2) 前 40 列內掃字
    R = min(40, len(df))
    C = len(df.columns)
    for r in range(R):
        row = df.iloc[r, :]
        for c in range(C):
            s = str(row.iloc[c]).strip().lower()
            if not s:
                continue
            for k in keywords:
                if k in s:
                    if VERBOSE:
                        print(f"[debug] bonus col found in cells -> row={r}, col={c}, text={row.iloc[c]!r}")
                    return c
    return None


# ====== 全表搜尋（加強 fallback） ======
def _find_row_anywhere_by_name(df_bonus: pd.DataFrame, target_norm: str, min_ratio: float = 0.70):
    best = None  # (ratio, r, c, raw)
    R, C = df_bonus.shape

    # 1) 完全相等（正規化）
    for c in range(C):
        col_norm = _series_norm(df_bonus.iloc[:, c].astype(str))
        idx = col_norm[col_norm == target_norm].index
        if len(idx) > 0:
            r = int(idx[0])
            raw = df_bonus.iloc[r, c]
            return (r, c, str(raw), 1.0)

    # 2) 模糊比對
    for r in range(R):
        row = df_bonus.iloc[r, :]
        for c in range(C):
            raw = str(row.iloc[c])
            nm = _norm_name(raw)
            if not nm:
                continue
            ratio = difflib.SequenceMatcher(None, target_norm, nm).ratio()
            if best is None or ratio > best[0]:
                best = (ratio, r, c, raw)

    if best and best[0] >= min_ratio:
        ratio, r, c, raw = best
        return (r, c, raw, ratio)
    return None


# ====== 取得獎金（含終極 fallback） ======
def _get_bonus_by_name(df_bonus: pd.DataFrame, target: str, att_names: set[str]) -> int | None:
    if df_bonus is None:
        return None

    # 1) 找姓名欄（用出勤名字集合推測）
    name_col = _guess_name_col_in_bonus(df_bonus, att_names)

    # 2) 直接用「輸入的名字」查找（不做別名）
    target_norm = _norm_name(target)
    r = None  # 目標列

    # 3) 先用「推測到的姓名欄」找列
    if name_col is not None:
        names_norm = _series_norm(df_bonus.iloc[:, name_col].astype(str))
        idx = names_norm[names_norm == target_norm].index
        if len(idx) > 0:
            r = int(idx[0])
            if VERBOSE:
                print(f"[debug] name found in guessed name_col={name_col}, row={r}")

    # 4) 推測欄找不到 → 逐欄掃描
    if r is None:
        C = df_bonus.shape[1]
        found = False
        for c_try in range(C):
            col_norm = _series_norm(df_bonus.iloc[:, c_try].astype(str))
            idx = col_norm[col_norm == target_norm].index
            if len(idx) > 0:
                r = int(idx[0])
                name_col = c_try
                found = True
                if VERBOSE:
                    print(f"[debug] name found by full-scan at col={c_try}, row={r}")
                break

        # 5) 還是找不到 → 整張表模糊比對（最終 fallback）
        if not found:
            hit = _find_row_anywhere_by_name(df_bonus, target_norm, min_ratio=0.70)
            if hit is not None:
                r, name_col, raw_cell, ratio = hit
                if VERBOSE:
                    print(f"[debug] name found by fuzzy-anywhere at col={name_col}, row={r}, "
                          f"text={raw_cell!r}, ratio={ratio:.3f}")
            else:
                if VERBOSE:
                    print(f"[debug] target '{target}' not found in bonus sheet.")
                return None

    # 6) 找獎金欄（欄名→內容掃字→fallback 到 config）
    c = _find_bonus_col(df_bonus)
    if c is None:
        c = BONUS_COL_INDEX
        if VERBOSE:
            print(f"[debug] bonus col fallback to config index -> {c}")

    if r < 0 or r >= len(df_bonus) or c < 0 or c >= len(df_bonus.columns):
        if VERBOSE:
            print(f"[debug] out of range: r={r}, c={c}, shape={df_bonus.shape}")
        return None

    val = df_bonus.iloc[r, c]
    if VERBOSE:
        print(f"[debug] bonus raw value at (r={r}, c={c}) -> {val!r}")

    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None

    n = parse_ntd(val)
    return n if n != 0 else None


# ====== 主流程 ======
def main():
    # 先用索引/名稱解析出勤分頁，再讀 header=None / header=1
    xls_att = pd.ExcelFile(FILE_PATH)
    att_sheet_name = resolve_sheet_name(
        xls_att, SHEET_ATTEND, SHEET_ATTEND_INDEX, USE_SHEET_INDEX, fallback_index=None
    )
    df_raw = pd.read_excel(xls_att, sheet_name=att_sheet_name, header=None)
    df_att = pd.read_excel(xls_att, sheet_name=att_sheet_name, header=1)

    # 讀薪資摘要與獎金表（已支援索引）
    df_sum = _read_summary_sheet(verbose=VERBOSE)
    df_bonus = _read_bonus_sheet(verbose=VERBOSE)

    # 只詢問一次要掃描的 A 欄人數
    try:
        x = int(input(f"輸入A欄要檢查的員工人數x（預設 {SCAN_ROWS_DEFAULT}）："))
    except Exception:
        x = SCAN_ROWS_DEFAULT
        print(f"輸入非數字，預設使用 x={SCAN_ROWS_DEFAULT}")

    # 顯示一次 A 欄供參考
    a_values = df_att.iloc[:x, 0]
    print("A欄內容：")
    for idx, val in enumerate(a_values, start=1):
        print(f"A{idx}: {val}")

    # **一次算好出勤姓名集合，迴圈重用**
    att_names = set(_series_norm(df_att.iloc[:, 0]))

    # 進入查詢迴圈
    while True:
        target = input("\n輸入員工姓名：").strip()

        total_min, lines = list_employee_days(
            df_raw, df_att, target,
            scan_rows=x,
            start_col=START_COL,
            date_row_index=DATE_ROW_INDEX,
            group_stride=GROUP_STRIDE,
        )

        if total_min is None and not lines:
            print("找不到對應資料！")
        else:
            if total_min is not None:
                print(f"{target} 的 總分鐘數為 :{total_min}分鐘")
            print("上班天數")
            for s in lines:
                print(s)

            # 出勤統計
            try:
                days = len(lines)
                hours = round(total_min / 60, 2) if total_min is not None else None
                if days:
                    print(f"\n統計：出勤天數 {days} 天", end="")
                    if hours is not None:
                        print(f"；總時數 {hours} 小時")
                    else:
                        print("")
            except Exception:
                pass

            # 薪資摘要（九項，為 0 自動略過）
            if df_sum is None:
                print("\n[提示] 未載入到薪資摘要表。請檢查 config 的設定。")
                pay_items = {}
            else:
                pay_items = extract_pay_items(df_sum, target, SUMMARY_FIELDS)

            # 獎金（名字定位 + 欄位偵測）
            bonus_amt = None
            if df_bonus is not None:
                bonus_amt = _get_bonus_by_name(df_bonus, target, att_names)

            # 輸出
            if pay_items or (bonus_amt is not None and bonus_amt > 0):
                print("\n薪資明細：")
                if bonus_amt is not None and bonus_amt > 0:
                    print(f"- {BONUS_FIELD}: {fmt_ntd(bonus_amt)}")
                for key in SUMMARY_FIELDS:
                    if key in pay_items:
                        print(f"- {key}: {fmt_ntd(pay_items[key])}")

        # 詢問是否繼續
        while True:
            cont = input("\n是否繼續查詢下一位？(Y/N)：").strip().upper()
            if cont in ("Y", "N"):
                break
            print("請輸入 Y 或 N。")

        if cont == "N":
            print("結束查詢。")
            break
        else:
            print("-" * 40)


if __name__ == "__main__":
    main()
