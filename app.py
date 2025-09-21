# -*- coding: utf-8 -*-
import io, os, math, re, unicodedata, difflib
import pandas as pd
import streamlit as st

from attendance_utils import (
    extract_employee_records,
    extract_pay_items,
    fmt_ntd,
    parse_ntd,
)
from config import (
    # 檔案/分頁設定
    FILE_PATH,
    USE_SHEET_INDEX,
    SHEET_ATTEND_INDEX, SHEET_BONUS_INDEX, SHEET_SUMMARY_INDEX,
    SHEET_ATTEND, SHEET_BONUS, SHEET_SUMMARY,
    # 出勤解析規則
    START_COL, DATE_ROW_INDEX, GROUP_STRIDE,
    # 摘要/獎金
    SUMMARY_FIELDS, SUMMARY_NAME_COL_INDEX,
    BONUS_FIELD, BONUS_COL_INDEX,
    # PDF 字型候選
    PDF_FONT_CANDIDATES,
)

# =============== PDF（reportlab） ===============
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _register_cjk_font(debug: bool=False) -> str:
    """
    依 config.PDF_FONT_CANDIDATES 依序嘗試註冊 CJK 字型。
    成功回傳字型名稱 'CJK'，失敗回 'Helvetica'，並在 debug 顯示失敗清單。
    """
    tried = []
    for path in PDF_FONT_CANDIDATES:
        if not path:
            continue
        # 允許相對路徑
        try_path = os.path.abspath(path)
        if os.path.exists(try_path):
            try:
                pdfmetrics.registerFont(TTFont("CJK", try_path))
                if debug:
                    st.caption(f"[debug] PDF font registered: {try_path}")
                return "CJK"
            except Exception as e:
                tried.append(f"{try_path} ({e})")
        else:
            tried.append(f"{try_path} (not found)")
    if debug and tried:
        st.warning("找不到可用的 PDF 字型，將退回 Helvetica。嘗試過：\n" + "\n".join(tried))
    return "Helvetica"


def build_pdf(employee: str, total_min: int | None, records: list[dict],
              bonus_amt: int | None, pay_items: dict, debug: bool=False) -> bytes:
    buf = io.BytesIO()
    font_name = _register_cjk_font(debug=debug)

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleCJK", fontName=font_name, fontSize=18, leading=22, spaceAfter=10))
    styles.add(ParagraphStyle(name="H2CJK", fontName=font_name, fontSize=14, leading=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="BodyCJK", fontName=font_name, fontSize=11, leading=15, spaceAfter=6))

    story = []
    # 標題 + 概覽
    t_min = total_min or 0
    t_days = len(records)
    t_hours = round(t_min/60, 2) if total_min is not None else 0.0
    story.append(Paragraph(f"員工出勤報表 - {employee}", styles["TitleCJK"]))
    story.append(Paragraph(f"出勤天數：{t_days} 天；總分鐘數：{t_min:,} 分；約 {t_hours} 小時", styles["BodyCJK"]))
    story.append(Spacer(1, 6))

    # 出勤表
    if records:
        table_data = [["日期","上班","下班","分鐘"]]
        for r in records:
            table_data.append([r["date"], r["in"], r["out"], f'{r["minutes"]:,}'])
        tbl = Table(table_data, colWidths=[90, 80, 80, 60])
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0,0), (-1,-1), font_name),
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("ALIGN", (1,1), (-2,-1), "CENTER"),
            ("ALIGN", (-1,1), (-1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    # 薪資明細
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
# =============== /PDF（reportlab） ===============


# ------------------ 共用工具 ------------------
def _norm_name(s: str) -> str:
    if s is None: return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\u00A0","").replace("\u3000","")
    s = re.sub(r"\s+","", s)
    s = s.replace("\u200b","").replace("\ufeff","")
    return s.strip()

def _series_norm(sr: pd.Series) -> pd.Series:
    return sr.astype(str).map(_norm_name)


def load_sheet(file_path: str, prefer_index: int | None, prefer_name: str | None,
               header=None, verbose: bool=False) -> pd.DataFrame | None:
    """
    統一的分頁載入：優先用索引（若 USE_SHEET_INDEX=True 且給了 index），
    否則用明確名稱；名稱不在時做去空白比對；最後回退第一個可用分頁。
    """
    try:
        xls = pd.ExcelFile(file_path)
        # 先依 index
        if USE_SHEET_INDEX and prefer_index is not None:
            try:
                return pd.read_excel(file_path, sheet_name=prefer_index, header=header)
            except Exception as e:
                if verbose: st.caption(f"[debug] 用索引 {prefer_index} 載入失敗：{e}")
        # 再依名稱（含去空白比對）
        if prefer_name:
            name = prefer_name
            if name not in xls.sheet_names:
                norm = lambda s: str(s).strip().replace(" ","")
                cand = [s for s in xls.sheet_names if norm(s)==norm(name)]
                if cand:
                    name = cand[0]
            if name in xls.sheet_names:
                return pd.read_excel(file_path, sheet_name=name, header=header)
        # 最後退回第一個
        return pd.read_excel(file_path, sheet_name=0, header=header)
    except Exception as e:
        if verbose: st.warning(f"[debug] 載入分頁失敗：{e}")
        return None


def find_bonus_col(df: pd.DataFrame, verbose: bool=False) -> int | None:
    keywords = ["獎金總和", "獎金", "bonus", "緊急"]
    cols_norm = [str(c).strip().lower() for c in df.columns]
    for k in keywords:
        k = k.lower()
        for idx, cname in enumerate(cols_norm):
            if k in cname and cname != "false":
                if verbose: st.caption(f"[debug] bonus col in columns -> {idx} ({df.columns[idx]!r})")
                return idx

    R, C = min(40, len(df)), len(df.columns)
    for r in range(R):
        row = df.iloc[r, :]
        for c in range(C):
            s = str(row.iloc[c]).strip().lower()
            if not s: continue
            for k in keywords:
                if k in s:
                    if verbose: st.caption(f"[debug] bonus col in cells -> row={r}, col={c}, text={row.iloc[c]!r}")
                    return c
    return None


def guess_name_col_in_bonus(df_bonus: pd.DataFrame, att_names: set[str], verbose: bool=False) -> int | None:
    att_norm = {_norm_name(x) for x in att_names if x is not None}
    best_i, best_hit = None, -1
    for i in range(len(df_bonus.columns)):
        col_vals = _series_norm(df_bonus.iloc[:, i])
        hit = col_vals.isin(att_norm).sum()
        if hit > best_hit:
            best_i, best_hit = i, hit
    if verbose: st.caption(f"[debug] bonus name col -> {best_i}, hits={best_hit}")
    return best_i if best_hit > 0 else None


def find_row_anywhere_by_name(df_bonus: pd.DataFrame, target_norm: str, min_ratio: float=0.70):
    best = None
    R, C = df_bonus.shape
    # 完全相等
    for c in range(C):
        col_norm = _series_norm(df_bonus.iloc[:, c].astype(str))
        idx = col_norm[col_norm == target_norm].index
        if len(idx) > 0:
            r = int(idx[0])
            raw = df_bonus.iloc[r, c]
            return (r, c, str(raw), 1.0)
    # 模糊比對
    for r in range(R):
        row = df_bonus.iloc[r, :]
        for c in range(C):
            raw = str(row.iloc[c])
            nm = _norm_name(raw)
            if not nm: continue
            ratio = difflib.SequenceMatcher(None, target_norm, nm).ratio()
            if best is None or ratio > best[0]:
                best = (ratio, r, c, raw)
    if best and best[0] >= min_ratio:
        ratio, r, c, raw = best
        return (r, c, raw, ratio)
    return None


def get_bonus_by_name(df_bonus: pd.DataFrame, target: str, att_names: set[str], verbose: bool=False) -> int | None:
    if df_bonus is None:
        return None
    name_col = guess_name_col_in_bonus(df_bonus, att_names, verbose)
    target_norm = _norm_name(target)
    r = None
    if name_col is not None:
        names_norm = _series_norm(df_bonus.iloc[:, name_col].astype(str))
        idx = names_norm[names_norm == target_norm].index
        if len(idx)>0:
            r = int(idx[0])
            if verbose: st.caption(f"[debug] name in guessed name_col={name_col}, row={r}")

    if r is None:
        C = df_bonus.shape[1]
        found = False
        for c_try in range(C):
            col_norm = _series_norm(df_bonus.iloc[:, c_try].astype(str))
            idx = col_norm[col_norm == target_norm].index
            if len(idx)>0:
                r = int(idx[0]); name_col = c_try; found = True
                if verbose: st.caption(f"[debug] name by full-scan at col={c_try}, row={r}")
                break
        if not found:
            hit = find_row_anywhere_by_name(df_bonus, target_norm, min_ratio=0.70)
            if hit is not None:
                r, name_col, raw_cell, ratio = hit
                if verbose: st.caption(f"[debug] name by fuzzy-anywhere at col={name_col}, row={r}, text={raw_cell!r}, ratio={ratio:.3f}")
            else:
                if verbose: st.caption(f"[debug] target '{target}' not found in bonus sheet.")
                return None

    c = find_bonus_col(df_bonus, verbose)
    if c is None:
        c = BONUS_COL_INDEX
        if verbose: st.caption(f"[debug] bonus col fallback to config index -> {c}")

    if r<0 or r>=len(df_bonus) or c<0 or c>=len(df_bonus.columns):
        if verbose: st.caption(f"[debug] out of range: r={r}, c={c}, shape={df_bonus.shape}")
        return None

    val = df_bonus.iloc[r, c]
    if verbose: st.caption(f"[debug] bonus raw value at (r={r}, c={c}) -> {val!r}")
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    n = parse_ntd(val)
    return n if n!=0 else None


# ------------------ 主 UI ------------------
def main():
    st.set_page_config(page_title="出勤與薪資總覽", page_icon="🗂️", layout="wide")
    st.title("出勤與薪資總覽")

    with st.sidebar:
        st.header("資料來源")
        uploaded_file = st.file_uploader("上傳 Excel（.xlsx）", type=["xlsx"])
        use_debug = st.toggle("顯示偵錯資訊", value=False)

    # 決定使用的檔案
    if uploaded_file is not None:
        tmp = os.path.join(st.session_state.get("_tmpdir", "."), "_uploaded.xlsx")
        with open(tmp, "wb") as f:
            f.write(uploaded_file.getbuffer())
        file_path = tmp
    else:
        file_path = FILE_PATH

    # 讀出勤（兩種 header）
    df_raw = load_sheet(file_path, SHEET_ATTEND_INDEX, SHEET_ATTEND, header=None, verbose=use_debug)
    df_att = load_sheet(file_path, SHEET_ATTEND_INDEX, SHEET_ATTEND, header=1, verbose=use_debug)
    if df_raw is None or df_att is None:
        st.error("讀取出勤分頁失敗，請檢查檔案或 config。")
        return

    # 員工清單（A 欄）
    names = df_att.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    if not names:
        st.warning("A 欄沒有找到員工姓名。")
        return

    # 讓你選人
    target = st.selectbox("選擇員工姓名", names, index=0)

    # 出勤紀錄
    total_min, records = extract_employee_records(
        df_raw, df_att, target,
        scan_rows=len(names),
        start_col=START_COL,
        date_row_index=DATE_ROW_INDEX,
        group_stride=GROUP_STRIDE,
    )
    if total_min is None and not records:
        st.warning("找不到對應資料！")
        return

    # 概覽
    col1, col2, col3 = st.columns([1,1,1])
    with col1: st.metric("出勤天數", f"{len(records)} 天")
    with col2: st.metric("總分鐘數", f"{(total_min or 0):,} 分")
    with col3: st.metric("總時數(約)", f"{round((total_min or 0)/60, 2)} 小時")

    # 出勤表
    if records:
        df_show = pd.DataFrame(records, columns=["date","in","out","minutes"]).rename(
            columns={"date":"日期", "in":"上班", "out":"下班", "minutes":"分鐘"}
        )
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("薪資明細")

    # 讀摘要（九項）
    df_sum = load_sheet(file_path, SHEET_SUMMARY_INDEX, SHEET_SUMMARY, header=0, verbose=use_debug)
    pay_items = {}
    if df_sum is not None:
        try:
            pay_items = extract_pay_items(df_sum, target, SUMMARY_FIELDS)
        except Exception as e:
            if use_debug: st.warning(f"[debug] 摘要計算失敗：{e}")
    else:
        if use_debug: st.warning("[debug] 未載入到摘要分頁。")

    # 讀獎金
    df_bonus = load_sheet(file_path, SHEET_BONUS_INDEX, SHEET_BONUS, header=1, verbose=use_debug)
    att_names = set(_series_norm(df_att.iloc[:, 0]))
    bonus_amt = get_bonus_by_name(df_bonus, target, att_names, verbose=use_debug) if df_bonus is not None else None

    # 顯示
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

    # 下載 PDF
    pdf_bytes = build_pdf(
        employee=target,
        total_min=total_min,
        records=records,
        bonus_amt=bonus_amt,
        pay_items=pay_items,
        debug=use_debug
    )
    st.download_button(
        label="⬇️ 下載 PDF 報表",
        data=pdf_bytes,
        file_name=f"{target}_出勤報表.pdf",
        mime="application/pdf"
    )

    if use_debug:
        with st.expander("偵錯資訊"):
            try:
                xls = pd.ExcelFile(file_path)
                st.caption(f"[debug] 檔案：{file_path}")
                st.caption(f"[debug] sheet_names: {xls.sheet_names}")
            except Exception as e:
                st.caption(f"[debug] 讀取 xls 失敗：{e}")

if __name__ == "__main__":
    main()
