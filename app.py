# -*- coding: utf-8 -*-
import io
import math
import re
import unicodedata
import difflib
from pathlib import Path

import pandas as pd
import streamlit as st

from attendance_utils import (
    extract_employee_records,  # 每日出勤紀錄（date, in, out, minutes）
    extract_pay_items,         # 依姓名從工作表3取九項（自動略過 0）
    fmt_ntd,                   # NT$ 千分位格式化
    parse_ntd,                 # 解析 NT$ 或數字為 int
)

from config import (
    FILE_PATH, SHEET_ATTEND,
    START_COL, DATE_ROW_INDEX, GROUP_STRIDE,
    SHEET_SUMMARY, SUMMARY_FIELDS, SUMMARY_NAME_COL_INDEX,
    SHEET_BONUS, BONUS_FIELD, BONUS_COL_INDEX,
    PDF_FONT_CANDIDATES,  # ← 在 config.py 註冊的候選字型
    PDF_FONT_NAME,        # ← 嵌入字型名稱（例如 CJK）
)

# =============== PDF 相關（reportlab）==============
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ---------- PDF 字型註冊 ----------
def _register_cjk_font() -> str:
    """
    依 PDF_FONT_CANDIDATES 順序嘗試註冊字型，註冊成功回傳 PDF_FONT_NAME，
    全失敗則回傳 'Helvetica'（英文內建字型）。
    """
    for p in PDF_FONT_CANDIDATES:
        try:
            # 使用絕對路徑避免相對路徑在雲端失效
            tt_path = str(Path(p).expanduser().resolve())
            pdfmetrics.registerFont(TTFont(PDF_FONT_NAME, tt_path))
            return PDF_FONT_NAME
        except Exception:
            continue
    return "Helvetica"


# ---------- 產生 PDF ----------
def build_pdf(employee: str, total_min: int | None, records: list[dict],
              bonus_amt: int | None, pay_items: dict) -> bytes:
    """
    產生 PDF：包含概覽、每日出勤表、薪資明細（含獎金總和）。
    """
    buf = io.BytesIO()
    font_name = _register_cjk_font()  # ← 先註冊字型，取得 font name

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28
    )

    styles = getSampleStyleSheet()
    # 這裡全部使用 CJK 字型（或 fallback Helvetica）
    styles.add(ParagraphStyle(name="TitleCJK", fontName=font_name, fontSize=18, leading=22, spaceAfter=10))
    styles.add(ParagraphStyle(name="H2CJK", fontName=font_name, fontSize=14, leading=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="BodyCJK", fontName=font_name, fontSize=11, leading=15, spaceAfter=6))

    story = []

    # 標題 + 概覽
    t_min = total_min or 0
    t_days = len(records)
    t_hours = round(t_min / 60, 2) if total_min is not None else 0.0

    story.append(Paragraph(f"員工出勤報表 - {employee}", styles["TitleCJK"]))
    story.append(Paragraph(f"出勤天數：{t_days} 天・總分鐘數：{t_min:,} 分・約 {t_hours} 小時", styles["BodyCJK"]))
    story.append(Spacer(1, 6))

    # 出勤表
    if records:
        story.append(Paragraph("每日出勤", styles["H2CJK"]))
        table_data = [["日期", "上班", "下班", "分鐘"]]
        for r in records:
            table_data.append([r["date"], r["in"], r["out"], f'{r["minutes"]:,}'])

        tbl = Table(table_data, colWidths=[90, 80, 80, 60])
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font_name),       # ← 關鍵：表格套字型
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ("ALIGN", (3, 1), (3, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 8))

    # 薪資明細
    story.append(Paragraph("薪資明細", styles["H2CJK"]))
    had_any = False

    if bonus_amt is not None and bonus_amt > 0:
        story.append(Paragraph(f"- {BONUS_FIELD}：{fmt_ntd(bonus_amt)}", styles["BodyCJK"]))
        had_any = True

    for key in SUMMARY_FIELDS:
        if key in pay_items:
            story.append(Paragraph(f"- {key}：{fmt_ntd(pay_items[key])}", styles["BodyCJK"]))
            had_any = True

    if not had_any:
        story.append(Paragraph("（此員工沒有可顯示的薪資明細）", styles["BodyCJK"]))

    doc.build(story)
    return buf.getvalue()
# =============== /PDF 相關（reportlab）==============


# ------------------ 共用工具 ------------------
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


# ------------------ 讀取工作表2（獎金） ------------------
@st.cache_data(show_spinner=False)
def _read_bonus_sheet(verbose: bool = False) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(FILE_PATH)
        sheet_name = SHEET_BONUS
        if isinstance(sheet_name, int):
            sheet_name = xls.sheet_names[sheet_name]  # 允許用 index 指定

        def _load(h):
            df = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=h)
            df.columns = df.columns.map(lambda s: str(s).strip())
            if False in df.columns:
                df = df.drop(columns=[False])
            return df

        df_bonus = _load(1)
        if df_bonus.shape[1] <= BONUS_COL_INDEX:
            df_bonus = _load(0)

        if verbose:
            st.caption(f"[debug] 使用獎金分頁：{sheet_name}；shape={df_bonus.shape}")
        return df_bonus

    except Exception as e:
        if verbose:
            st.warning(f"[debug] 讀取獎金表失敗：{e}")
        return None


def _guess_name_col_in_bonus(df_bonus: pd.DataFrame, att_names: set[str], verbose: bool=False) -> int | None:
    att_norm = {_norm_name(x) for x in att_names if x is not None}
    best_i, best_hit = None, -1
    for i in range(len(df_bonus.columns)):
        col_vals = _series_norm(df_bonus.iloc[:, i])
        hit = col_vals.isin(att_norm).sum()
        if hit > best_hit:
            best_i, best_hit = i, hit
    if verbose:
        st.caption(f"[debug] bonus name col -> {best_i}, hits={best_hit}")
    return best_i if best_hit > 0 else None

def _find_bonus_col(df: pd.DataFrame, verbose: bool=False) -> int | None:
    keywords = ["獎金總和", "獎金", "bonus", "緊急"]
    cols_norm = [str(c).strip().lower() for c in df.columns]
    for k in keywords:
        k = k.lower()
        for idx, cname in enumerate(cols_norm):
            if k in cname and cname != "false":
                if verbose:
                    st.caption(f"[debug] bonus col in columns -> {idx} ({df.columns[idx]!r})")
                return idx

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
                    if verbose:
                        st.caption(f"[debug] bonus col in cells -> row={r}, col={c}, text={row.iloc[c]!r}")
                    return c
    return None

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

def _get_bonus_by_name(df_bonus: pd.DataFrame, target: str, att_names: set[str], verbose: bool=False) -> int | None:
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
            if verbose:
                st.caption(f"[debug] name in guessed name_col={name_col}, row={r}")

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
                if verbose:
                    st.caption(f"[debug] name by full-scan at col={c_try}, row={r}")
                break

        if not found:
            hit = _find_row_anywhere_by_name(df_bonus, target_norm, min_ratio=0.70)
            if hit is not None:
                r, name_col, raw_cell, ratio = hit
                if verbose:
                    st.caption(f"[debug] name by fuzzy-anywhere at col={name_col}, row={r}, text={raw_cell!r}, ratio={ratio:.3f}")
            else:
                if verbose:
                    st.caption(f"[debug] target '{target}' not found in bonus sheet.")
                return None

    c = _find_bonus_col(df_bonus, verbose)
    if c is None:
        c = BONUS_COL_INDEX
        if verbose:
            st.caption(f"[debug] bonus col fallback to config index -> {c}")

    if r < 0 or r >= len(df_bonus) or c < 0 or c >= len(df_bonus.columns):
        if verbose:
            st.caption(f"[debug] out of range: r={r}, c={c}, shape={df_bonus.shape}")
        return None

    val = df_bonus.iloc[r, c]
    if verbose:
        st.caption(f"[debug] bonus raw value at (r={r}, c={c}) -> {val!r}")
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    n = parse_ntd(val)
    return n if n != 0 else None


# ------------------ 主 UI ------------------
def main():
    st.set_page_config(page_title="出勤與薪資總覽", page_icon="🗂️", layout="wide")
    st.title("出勤與薪資總覽")

    with st.sidebar:
        st.header("⚙️ 進階 / 除錯")
        verbose = st.toggle("顯示偵錯資訊 (VERBOSE)", value=False)
        st.caption(f"目前資料表：**{FILE_PATH} / {SHEET_ATTEND}**")

    # 讀工作表1（出勤）
    xls = pd.ExcelFile(FILE_PATH)
    attend_sheet = SHEET_ATTEND
    if isinstance(attend_sheet, int):
        attend_sheet = xls.sheet_names[attend_sheet]

    df_raw = pd.read_excel(FILE_PATH, sheet_name=attend_sheet, header=None)
    df_att = pd.read_excel(FILE_PATH, sheet_name=attend_sheet, header=1)

    # 員工清單（A 欄）
    names = df_att.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    # 我加了一個「要掃描 A 欄前幾名」輸入
    scan_n = st.number_input("掃描 A 欄前幾列（人名）", min_value=1, max_value=len(names), value=min(30, len(names)), step=1)
    target = st.selectbox("選擇員工姓名", names, index=0)

    # 出勤紀錄
    total_min, records = extract_employee_records(
        df_raw, df_att, target,
        scan_rows=scan_n,
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
        st.metric("總時數（約）", f"{round((total_min or 0)/60, 2)} 小時")

    # 出勤明細表
    if records:
        df_show = pd.DataFrame(records, columns=["date", "in", "out", "minutes"]).rename(
            columns={"date":"日期", "in":"上班", "out":"下班", "minutes":"分鐘"}
        )
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("薪資明細")

    # 讀工作表3（九項）
    try:
        sum_sheet = SHEET_SUMMARY
        if isinstance(sum_sheet, int):
            sum_sheet = xls.sheet_names[sum_sheet]
        df_sum = pd.read_excel(FILE_PATH, sheet_name=sum_sheet, header=0)
        pay_items = extract_pay_items(df_sum, target, SUMMARY_FIELDS)
    except Exception:
        pay_items = {}

    # 讀工作表2（獎金）
    df_bonus = _read_bonus_sheet(verbose)
    att_names = set(_series_norm(df_att.iloc[:, 0]))
    bonus_amt = _get_bonus_by_name(df_bonus, target, att_names, verbose) if df_bonus is not None else None

    # 顯示薪資明細
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

    # 下載 PDF 按鈕（包含薪資明細，且字型正確）
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

    # 除錯資訊
    if verbose:
        with st.expander("除錯資訊（展開查看）", expanded=False):
            st.write("A欄姓名（正規化後）前 30 筆：", list(att_names)[:30])
            if df_bonus is not None:
                st.write("工作表2（前 5x5）：")
                st.dataframe(df_bonus.iloc[:5, :5], use_container_width=True)


if __name__ == "__main__":
    main()
