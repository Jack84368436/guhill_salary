# -*- coding: utf-8 -*-
import io
import math
import re
import unicodedata
import difflib
import hashlib
from typing import Tuple, List, Dict

import pandas as pd
import streamlit as st

from attendance_utils import (
    extract_employee_records,  # 取得每日出勤紀錄（date, in, out, minutes）
    extract_pay_items,         # 依姓名從工作表3取九項（自動略過 0）
    fmt_ntd,                   # NT$ 千分位格式化
    parse_ntd,                 # 解析 NT$ 或數字為 int
)

from config import (
    # 檔案/分頁名稱
    FILE_PATH, SHEET_ATTEND, SHEET_BONUS, SHEET_SUMMARY,
    # 出勤解析規則
    START_COL, DATE_ROW_INDEX, GROUP_STRIDE,
    # 摘要九項
    SUMMARY_FIELDS, SUMMARY_NAME_COL_INDEX,
    # 獎金欄名/預設索引
    BONUS_FIELD, BONUS_COL_INDEX,
    # PDF 字型候選
    PDF_FONT_CANDIDATES,
    # << 新增：是否採用索引與三分頁索引 >>
    USE_SHEET_INDEX,
    SHEET_ATTEND_INDEX, SHEET_BONUS_INDEX, SHEET_SUMMARY_INDEX,
)

# =============== PDF 相關（reportlab）==============
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ---- 資料來源抽象（上傳/切換）----
if "excel_bytes" not in st.session_state:
    st.session_state["excel_bytes"] = None   # 放上傳檔 bytes；None 代表用預設檔
if "source_label" not in st.session_state:
    st.session_state["source_label"] = f"預設檔：{FILE_PATH}"

def get_active_excel_file() -> io.BytesIO | str:
    """回傳目前生效的 Excel 來源（BytesIO 或 檔案路徑）。"""
    if st.session_state.get("excel_bytes") is not None:
        return io.BytesIO(st.session_state["excel_bytes"])
    return FILE_PATH

def get_active_source_label() -> str:
    """顯示目前資料來源標籤。"""
    return st.session_state.get("source_label", f"預設檔：{FILE_PATH}")

def get_excel_cache_key() -> str:
    """給 cache 用的 key：上傳檔用其 md5；預設檔用固定字串。"""
    if st.session_state.get("excel_bytes") is not None:
        md5 = hashlib.md5(st.session_state["excel_bytes"]).hexdigest()
        return f"upload:{md5}"
    return f"default:{FILE_PATH}"


# ------------------ 共用工具 ------------------
# app.py
from pathlib import Path
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def _register_cjk_font() -> str:
    """
    嘗試註冊 CJK 字型，成功回傳字型名稱 'CJK'；失敗回傳 'Helvetica'。
    會嘗試 config.PDF_FONT_CANDIDATES 中的路徑。
    """
    from config import PDF_FONT_CANDIDATES  # 避免循環引用

    for p in PDF_FONT_CANDIDATES:
        try:
            path = Path(p)
            if path.exists():
                pdfmetrics.registerFont(TTFont("CJK", str(path)))
                # 讓粗體/斜體也不會回退成 Helvetica
                pdfmetrics.registerFontFamily("CJK", normal="CJK", bold="CJK",
                                              italic="CJK", boldItalic="CJK")
                return "CJK"
        except Exception as e:
            # 你可以在 Streamlit 裡顯示 debug：st.caption(f"[debug] font register fail: {p} -> {e}")
            pass
    return "Helvetica"


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

def resolve_sheet_name(xls: pd.ExcelFile, prefer_name: str, index: int | None,
                       use_index: bool, fallback_index: int | None = None) -> str:
    """
    依設定回傳要用的工作表名稱：
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


# ------------------ Schema 檢核（索引優先） ------------------
def _find_bonus_col(df: pd.DataFrame) -> int | None:
    """在欄名或前 40 列內找『獎金總和/獎金/bonus/緊急』。"""
    keywords = ["獎金總和", "獎金", "bonus", "緊急"]
    cols_norm = [str(c).strip().lower() for c in df.columns]
    for k in keywords:
        k = k.lower()
        for idx, cname in enumerate(cols_norm):
            if k in cname and cname != "false":
                return idx
    R = min(40, len(df))
    C = len(df.columns)
    for r in range(R):
        for c in range(C):
            s = str(df.iloc[r, c]).strip().lower()
            if not s:
                continue
            for k in keywords:
                if k in s:
                    return c
    return None

def _validate_attend_sheet(xls: pd.ExcelFile) -> Tuple[bool, str]:
    try:
        att_name = resolve_sheet_name(
            xls, SHEET_ATTEND, SHEET_ATTEND_INDEX, USE_SHEET_INDEX, fallback_index=None
        )
        df = pd.read_excel(xls, sheet_name=att_name, header=1)
        if df.shape[1] < 2:
            return False, "出勤分頁欄數異常。"
        a = df.iloc[:, 0].dropna().astype(str).str.strip()
        if a.empty:
            return False, "出勤分頁 A 欄（姓名）為空。"
        return True, ""
    except Exception as e:
        return False, f"讀取出勤分頁失敗：{e}"

def _validate_bonus_sheet(xls: pd.ExcelFile) -> Tuple[bool, str]:
    try:
        b_name = resolve_sheet_name(
            xls, SHEET_BONUS, SHEET_BONUS_INDEX, USE_SHEET_INDEX, fallback_index=1
        )
        # header=1 / 0 都試，只要能找出獎金欄即可
        for h in (1, 0):
            df = pd.read_excel(xls, sheet_name=b_name, header=h)
            c = _find_bonus_col(df)
            if c is not None:
                return True, ""
        return False, "獎金分頁未偵測到『獎金總和』欄位（或關鍵字）。"
    except Exception as e:
        return False, f"讀取獎金分頁失敗：{e}"

def _validate_summary_sheet(xls: pd.ExcelFile) -> Tuple[bool, str, List[str], List[str]]:
    """
    回傳 (ok, msg, expected_cols, actual_cols)
    """
    last_cols: List[str] = []
    try:
        s_name = resolve_sheet_name(
            xls, SHEET_SUMMARY, SHEET_SUMMARY_INDEX, USE_SHEET_INDEX, fallback_index=2
        )
        for h in (1, 0):
            df = pd.read_excel(xls, sheet_name=s_name, header=h)
            cols = [str(c).strip() for c in df.columns]
            last_cols = cols
            if all(field in cols for field in SUMMARY_FIELDS):
                return True, "", SUMMARY_FIELDS, cols
        missing = [f for f in SUMMARY_FIELDS if f not in last_cols]
        return False, f"摘要分頁缺少欄位：{missing}", SUMMARY_FIELDS, last_cols
    except Exception as e:
        return False, f"讀取摘要分頁失敗：{e}", SUMMARY_FIELDS, last_cols

def validate_schema(xls: pd.ExcelFile) -> Tuple[bool, List[str], Dict[str, List[str]]]:
    problems: List[str] = []
    extras: Dict[str, List[str]] = {}

    ok1, msg1 = _validate_attend_sheet(xls)
    ok2, msg2 = _validate_bonus_sheet(xls)
    ok3, msg3, expected, actual = _validate_summary_sheet(xls)

    if not ok1: problems.append(f"[出勤] {msg1}")
    if not ok2: problems.append(f"[獎金] {msg2}")
    if not ok3:
        problems.append(f"[摘要] {msg3}")
        extras["摘要-實際欄位"] = actual
        extras["摘要-期望欄位"] = expected

    return (len(problems) == 0, problems, extras)


# ------------------ 讀取工作表3（摘要/九項） ------------------
@st.cache_data(show_spinner=False)
def _read_summary_sheet(excel_key: str, verbose: bool = False) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(get_active_excel_file())
        s_name = resolve_sheet_name(
            xls, SHEET_SUMMARY, SHEET_SUMMARY_INDEX, USE_SHEET_INDEX, fallback_index=2
        )

        def _load_with_header(h):
            df = pd.read_excel(xls, sheet_name=s_name, header=h)
            df.columns = df.columns.map(lambda s: str(s).strip())
            try:
                col = df.iloc[:, SUMMARY_NAME_COL_INDEX].astype("object")
                df.iloc[:, SUMMARY_NAME_COL_INDEX] = col.astype(str).str.strip()
            except Exception:
                pass
            if False in df.columns:
                df = df.drop(columns=[False])
            return df

        df_sum = _load_with_header(1)
        if not any(col in df_sum.columns for col in SUMMARY_FIELDS):
            df_sum = _load_with_header(0)
        return df_sum
    except Exception:
        return None


# ------------------ 讀取工作表2（獎金） ------------------
@st.cache_data(show_spinner=False)
def _read_bonus_sheet(excel_key: str, verbose: bool = False) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(get_active_excel_file())
        b_name = resolve_sheet_name(
            xls, SHEET_BONUS, SHEET_BONUS_INDEX, USE_SHEET_INDEX, fallback_index=1
        )

        def _load(h):
            df = pd.read_excel(xls, sheet_name=b_name, header=h)
            df.columns = df.columns.map(lambda s: str(s).strip())
            if False in df.columns:
                df = df.drop(columns=[False])
            return df

        df_bonus = _load(1)
        if df_bonus.shape[1] <= BONUS_COL_INDEX:
            df_bonus = _load(0)
        return df_bonus
    except Exception:
        return None


# ------------------ 找姓名欄/獎金欄 & 取獎金 ------------------
def _guess_name_col_in_bonus(df_bonus: pd.DataFrame, att_names: set[str], verbose: bool=False) -> int | None:
    att_norm = {_norm_name(x) for x in att_names if x is not None}
    best_i, best_hit = None, -1
    for i in range(len(df_bonus.columns)):
        col_vals = _series_norm(df_bonus.iloc[:, i])
        hit = col_vals.isin(att_norm).sum()
        if hit > best_hit:
            best_i, best_hit = i, hit
    return best_i if best_hit > 0 else None

def _find_bonus_col_for_runtime(df: pd.DataFrame, verbose: bool=False) -> int | None:
    return _find_bonus_col(df)

def _find_row_anywhere_by_name(df_bonus: pd.DataFrame, target_norm: str, min_ratio: float = 0.70):
    best = None
    R, C = df_bonus.shape
    for c in range(C):
        col_norm = _series_norm(df_bonus.iloc[:, c].astype(str))
        idx = col_norm[col_norm == target_norm].index
        if len(idx) > 0:
            r = int(idx[0])
            raw = df_bonus.iloc[r, c]
            return (r, c, str(raw), 1.0)
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

def _get_bonus_by_name(df_bonus: pd.DataFrame, target: str, att_names: set[str],
                       verbose: bool = False) -> int | None:
    if df_bonus is None:
        return None
    name_col = _guess_name_col_in_bonus(df_bonus, att_names, verbose)
    target_norm = _norm_name(target)
    r = None

    if name_col is not None:
        names_norm = _series_norm(df_bonus.iloc[:, name_col].astype(str))
        idx = names_norm[names_norm == target_norm].index
        if len(idx) > 0:
            r = int(idx[0])

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
                break

        if not found:
            hit = _find_row_anywhere_by_name(df_bonus, target_norm, min_ratio=0.70)
            if hit is not None:
                r, name_col, raw_cell, ratio = hit
            else:
                return None

    c = _find_bonus_col_for_runtime(df_bonus, verbose)
    if c is None:
        c = BONUS_COL_INDEX

    if r < 0 or r >= len(df_bonus) or c < 0 or c >= len(df_bonus.columns):
        return None

    val = df_bonus.iloc[r, c]
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    n = parse_ntd(val)
    return n if n != 0 else None


# ------------------ PDF 產生 ------------------
def build_pdf(employee: str, total_min: int | None, records: list[dict],
              bonus_amt: int | None, pay_items: dict) -> bytes:
    buf = io.BytesIO()
    font_name = _register_cjk_font()

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleCJK", fontName=font_name, fontSize=18, leading=22, spaceAfter=10))
    styles.add(ParagraphStyle(name="H2CJK",   fontName=font_name, fontSize=14, leading=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="BodyCJK", fontName=font_name, fontSize=11, leading=15, spaceAfter=6))
    styles.add(ParagraphStyle(name="MonoCJK", fontName=font_name, fontSize=10, leading=14))

# 表格要記得也指定字型
    tbl.setStyle(TableStyle([
        ("FONTNAME",  (0, 0), (-1, -1), font_name),
        ("FONTSIZE",  (0, 0), (-1, -1), 10),
        ("GRID",      (0, 0), (-1, -1), 0.25, colors.grey),
        ("BACKGROUND",(0, 0), (-1, 0),  colors.whitesmoke),
        ("ALIGN",     (1, 1), (-1, -1), "CENTER"),
        ("ALIGN",     (3, 1), (3, -1),  "RIGHT"),
        ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
    ]))

    story = []

    story.append(Paragraph(f"{employee} 出勤報表", styles["TitleCJK"]))
    t_min = total_min or 0
    t_days = len(records)
    t_hours = round(t_min / 60, 2) if total_min is not None else 0.0
    story.append(Paragraph(f"出勤天數：{t_days} 天；總分鐘數：{t_min:,} 分；約 {t_hours} 小時", styles["BodyCJK"]))
    story.append(Spacer(1, 6))

    if records:
        table_data = [["日期", "上班", "下班", "分鐘"]]
        for r in records:
            table_data.append([r["date"], r["in"], r["out"], f'{r["minutes"]:,}'])
        tbl = Table(table_data, colWidths=[90, 80, 80, 60])
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ("ALIGN", (3, 1), (3, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    story.append(Paragraph("薪資明細", styles["H2CJK"]))
    printed_any = False
    if bonus_amt is not None and bonus_amt > 0:
        story.append(Paragraph(f"- {BONUS_FIELD}：{fmt_ntd(bonus_amt)}", styles["BodyCJK"]))
        printed_any = True
    for key in SUMMARY_FIELDS:
        if key in pay_items:
            story.append(Paragraph(f"- {key}：{fmt_ntd(pay_items[key])}", styles["BodyCJK"]))
            printed_any = True
    if not printed_any:
        story.append(Paragraph("（此員工沒有可顯示的薪資明細）", styles["BodyCJK"]))

    doc.build(story)
    return buf.getvalue()


# ------------------ 範本 Excel 產生（保留） ------------------
def make_template_excel() -> bytes:
    """
    依 config 產生一份格式正確但空資料的範本：
    - 工作表1：放兩位示範姓名與一組日期欄頭（B欄起每3欄），不填時間
    - 工作表2：獎金；姓名＋『獎金總和』欄位
    - 工作表3：摘要九項；姓名＋SUMMARY_FIELDS
    """
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        # 工作表1（範例欄頭）
        df_att_head = pd.DataFrame([["姓名"] + [""] * 12], columns=None)
        df_att_head.to_excel(writer, sheet_name=SHEET_ATTEND, index=False, header=False)
        # header=1 的表頭列（姓名 + 4 天示意）
        days = ["8/1", "8/2", "8/3", "8/4"]
        cols = ["姓名"]
        for d in days:
            cols.extend([f"{d}-上班", f"{d}-下班", f"{d}-分鐘"])
        df_att = pd.DataFrame(columns=cols)
        df_att.loc[0, "姓名"] = "小明"
        df_att.loc[1, "姓名"] = "小美"
        df_att.to_excel(writer, sheet_name=SHEET_ATTEND, index=False, header=True, startrow=1)

        # 工作表2（獎金）
        df_bonus = pd.DataFrame({
            "姓名": ["小明", "小美"],
            "獎金總和": [0, 0],
        })
        df_bonus.to_excel(writer, sheet_name=SHEET_BONUS, index=False)

        # 工作表3（摘要）
        cols = ["姓名"] + SUMMARY_FIELDS
        df_sum = pd.DataFrame(columns=cols)
        df_sum.loc[0, "姓名"] = "小明"
        df_sum.loc[1, "姓名"] = "小美"
        for f in SUMMARY_FIELDS:
            df_sum.loc[0, f] = 0
            df_sum.loc[1, f] = 0
        df_sum.to_excel(writer, sheet_name=SHEET_SUMMARY, index=False)

    return bio.getvalue()


# ------------------ 主 UI ------------------
def main():
    st.set_page_config(page_title="出勤與薪資總覽", page_icon="🗂️", layout="wide")
    st.title("出勤與薪資總覽")
    st.caption(f"目前資料表：{get_active_source_label()}")

    with st.sidebar:
        st.header("📂 上傳/切換資料來源")
        uploaded = st.file_uploader("上傳薪資 Excel（.xlsx）", type=["xlsx"])
        if uploaded is not None:
            try:
                excel_bytes = uploaded.getvalue()
                xls = pd.ExcelFile(io.BytesIO(excel_bytes))

                # 檢核
                ok, problems, extras = validate_schema(xls)
                if ok:
                    st.success("✅ 格式檢核通過")
                else:
                    st.error("❌ 格式檢核失敗：")
                    for p in problems:
                        st.write("-", p)
                    if "摘要-期望欄位" in extras:
                        st.caption("期望欄位：")
                        st.code(", ".join(extras["摘要-期望欄位"]))
                    if "摘要-實際欄位" in extras:
                        st.caption("實際欄位：")
                        st.code(", ".join(extras["摘要-實際欄位"]))

                # 預覽（前 5×8）
                st.divider()
                st.subheader("👀 上傳檔預覽（前 5×8）")
                try:
                    att_name = resolve_sheet_name(xls, SHEET_ATTEND, SHEET_ATTEND_INDEX, USE_SHEET_INDEX, None)
                    st.caption(f"【{att_name}】")
                    st.dataframe(pd.read_excel(xls, sheet_name=att_name, header=1).iloc[:5, :8])
                except Exception as e:
                    st.warning(f"出勤表預覽失敗：{e}")

                try:
                    b_name = resolve_sheet_name(xls, SHEET_BONUS, SHEET_BONUS_INDEX, USE_SHEET_INDEX, 1)
                    st.caption(f"【{b_name}】")
                    # 優先 header=1；不行退 0
                    try:
                        st.dataframe(pd.read_excel(xls, sheet_name=b_name, header=1).iloc[:5, :8])
                    except Exception:
                        st.dataframe(pd.read_excel(xls, sheet_name=b_name, header=0).iloc[:5, :8])
                except Exception as e:
                    st.warning(f"獎金表預覽失敗：{e}")

                try:
                    s_name = resolve_sheet_name(xls, SHEET_SUMMARY, SHEET_SUMMARY_INDEX, USE_SHEET_INDEX, 2)
                    st.caption(f"【{s_name}】")
                    st.dataframe(pd.read_excel(xls, sheet_name=s_name, header=1).iloc[:5, :8])
                except Exception as e:
                    st.warning(f"摘要表預覽失敗：{e}")

                # 乾跑：拿出勤第一個姓名測試九項/獎金
                st.divider()
                st.subheader("🧪 乾跑測試")
                try:
                    df_att_preview = pd.read_excel(xls, sheet_name=att_name, header=1)
                    first_name = str(df_att_preview.iloc[0, 0]).strip()
                    st.caption(f"測試員工：{first_name}")

                    # 九項
                    df_sum_preview = pd.read_excel(xls, sheet_name=s_name, header=0)
                    pay_items_preview = extract_pay_items(df_sum_preview, first_name, SUMMARY_FIELDS)
                    if pay_items_preview:
                        st.markdown("**九項（非 0 才顯示）**")
                        for k, v in pay_items_preview.items():
                            st.markdown(f"- {k}：{fmt_ntd(v)}")
                    else:
                        st.info("九項全部為 0 或未找到。")

                    # 獎金
                    try:
                        df_bonus_preview = pd.read_excel(xls, sheet_name=b_name, header=1)
                    except Exception:
                        df_bonus_preview = pd.read_excel(xls, sheet_name=b_name, header=0)

                    def _guess_name_col_in_bonus_preview(df_bonus: pd.DataFrame, att_names: set[str]) -> int | None:
                        att_norm = {_norm_name(x) for x in att_names if x is not None}
                        best_i, best_hit = None, -1
                        for i in range(len(df_bonus.columns)):
                            col_vals = _series_norm(df_bonus.iloc[:, i])
                            hit = col_vals.isin(att_norm).sum()
                            if hit > best_hit:
                                best_i, best_hit = i, hit
                        return best_i if best_hit > 0 else None

                    att_names_preview = set(_series_norm(df_att_preview.iloc[:, 0]))
                    name_col_guess = _guess_name_col_in_bonus_preview(df_bonus_preview, att_names_preview)
                    bonus_col_guess = _find_bonus_col(df_bonus_preview)
                    st.caption(f"姓名欄猜測：{name_col_guess}；獎金欄猜測：{bonus_col_guess}")

                    def _get_bonus_preview(df_bonus, name):
                        target_norm = _norm_name(name)
                        # 找列
                        row_idx = None
                        if name_col_guess is not None:
                            col_norm = _series_norm(df_bonus.iloc[:, name_col_guess].astype(str))
                            idx = col_norm[col_norm == target_norm].index
                            if len(idx) > 0:
                                row_idx = int(idx[0])
                        if row_idx is None:
                            C = df_bonus.shape[1]
                            for c in range(C):
                                col_norm = _series_norm(df_bonus.iloc[:, c].astype(str))
                                idx = col_norm[col_norm == target_norm].index
                                if len(idx) > 0:
                                    row_idx = int(idx[0])
                                    break
                        if row_idx is None:
                            return None
                        c = bonus_col_guess if bonus_col_guess is not None else BONUS_COL_INDEX
                        if c >= df_bonus.shape[1]:
                            return None
                        val = df_bonus.iloc[row_idx, c]
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            return None
                        n = parse_ntd(val)
                        return n if n != 0 else None

                    bonus_preview = _get_bonus_preview(df_bonus_preview, first_name)
                    if bonus_preview is not None and bonus_preview > 0:
                        st.markdown(f"**獎金總和**：{fmt_ntd(bonus_preview)}")
                    else:
                        st.info("未取得有效的獎金總和（可能為 0 或未找到）。")

                except Exception as e:
                    st.warning(f"乾跑測試時發生例外：{e}")

                st.divider()
                if ok and st.button("✅ 採用此檔案", type="primary"):
                    st.session_state["excel_bytes"] = excel_bytes
                    st.session_state["source_label"] = f"上傳檔：{uploaded.name}"
                    st.rerun()

            except Exception as e:
                st.error(f"讀取上傳檔失敗：{e}")

        if st.button("↩️ 改回預設檔"):
            st.session_state["excel_bytes"] = None
            st.session_state["source_label"] = f"預設檔：{FILE_PATH}"
            st.rerun()

        st.divider()
        st.header("⬇️ 下載範本")
        tpl_bytes = make_template_excel()
        st.download_button("下載範本 Excel", data=tpl_bytes, file_name="薪資出勤範本.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===== 主內容區 =====

    # 讀工作表1（出勤）——透過 active source + 索引解析
    xls_att = pd.ExcelFile(get_active_excel_file())
    att_name = resolve_sheet_name(xls_att, SHEET_ATTEND, SHEET_ATTEND_INDEX, USE_SHEET_INDEX, None)
    df_raw = pd.read_excel(xls_att, sheet_name=att_name, header=None)
    df_att = pd.read_excel(xls_att, sheet_name=att_name, header=1)

    # 提供「掃描 A 欄前幾列」的控制
    max_rows = len(df_att)
    scan_rows = st.number_input("掃描 A 欄前幾列（人名）", min_value=1, max_value=max_rows, value=min(30, max_rows), step=1)

    # 員工清單（A 欄）
    names = df_att.iloc[:scan_rows, 0].dropna().astype(str).str.strip().tolist()
    target = st.selectbox("選擇員工姓名", names, index=0)

    # 出勤紀錄
    total_min, records = extract_employee_records(
        df_raw, df_att, target,
        scan_rows=scan_rows,
        start_col=START_COL,
        date_row_index=DATE_ROW_INDEX,
        group_stride=GROUP_STRIDE,
    )

    if total_min is None and not records:
        st.warning("找不到對應資料！")
        return

    # 概覽 metrics
    col_top1, col_top2, col_top3 = st.columns([1,1,1])
    with col_top1:
        st.metric("出勤天數", f"{len(records)} 天")
    with col_top2:
        st.metric("總分鐘數", f"{(total_min or 0):,} 分")
    with col_top3:
        st.metric("總時數(約)", f"{round((total_min or 0)/60, 2)} 小時")

    # 出勤明細表
    if records:
        df_show = pd.DataFrame(records, columns=["date", "in", "out", "minutes"]).rename(
            columns={"date":"日期", "in":"上班", "out":"下班", "minutes":"分鐘"}
        )
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    st.divider()

    # ── 薪資明細（工作表2 + 工作表3）----
    st.subheader("薪資明細")

    # cache key（切換來源時刷新 cache）
    excel_key = get_excel_cache_key()

    # 讀工作表3（九項）— 容錯讀表 + cache
    df_sum = _read_summary_sheet(excel_key, False)
    if df_sum is None:
        pay_items = {}
    else:
        try:
            pay_items = extract_pay_items(df_sum, target, SUMMARY_FIELDS)
        except Exception:
            pay_items = {}

    # 讀工作表2（獎金）— 容錯讀表 + cache
    df_bonus = _read_bonus_sheet(excel_key, False)
    att_names = set(_series_norm(df_att.iloc[:, 0]))
    bonus_amt = _get_bonus_by_name(df_bonus, target, att_names, False) if df_bonus is not None else None

    printed_any = False
    if bonus_amt is not None and bonus_amt > 0:
        st.markdown(f"- **{BONUS_FIELD}**：{fmt_ntd(bonus_amt)}")
        printed_any = True
    for key in SUMMARY_FIELDS:
        if key in pay_items:
            st.markdown(f"- **{key}**：{fmt_ntd(pay_items[key])}")
            printed_any = True
    if not printed_any:
        st.info("此員工沒有可顯示的薪資明細（九項皆為 0，且未找到有效獎金）。")

    # ===== 下載 PDF 按鈕（包含薪資明細）=====
    pdf_bytes = build_pdf(
        employee=target,
        total_min=total_min,
        records=records,
        bonus_amt=bonus_amt,
        pay_items=pay_items
    )
    st.download_button(
        label="⬇️ 下載 PDF 報表",
        data=pdf_bytes,
        file_name=f"{target}_出勤報表.pdf",
        mime="application/pdf"
    )


if __name__ == "__main__":
    main()
