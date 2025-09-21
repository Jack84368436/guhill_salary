# -*- coding: utf-8 -*-
"""
集中設定：更換月份/檔名/工作表規則，只改這裡就好。
"""
from pathlib import Path

# ── 專案根目錄（給字型路徑用） ─────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent

# ── 預設 Excel 檔案（當未上傳或按「改回預設檔」時使用） ─────────────────
FILE_PATH = "salary8.xlsx"

# ── 若 USE_SHEET_INDEX=True，將用索引鎖定分頁（0-based） ────────────────
USE_SHEET_INDEX = True            # True=用索引找分頁；False=用名稱/模糊比對/回退
SHEET_ATTEND_INDEX  = 0           # 出勤分頁在第幾張（0-based）
SHEET_BONUS_INDEX   = 1           # 獎金分頁在第幾張（0-based）
SHEET_SUMMARY_INDEX = 2           # 摘要分頁在第幾張（0-based）

# ── 若 USE_SHEET_INDEX=False，則用名稱尋找分頁（仍具去空白比對與回退） ───────
SHEET_ATTEND  = "工作表 1 - 2025_8月"   # 出勤
SHEET_BONUS   = "工作表2"              # 獎金
SHEET_SUMMARY = "工作表3"              # 摘要（九項）

# ── 出勤欄位解析規則（適配你的原始檔格式） ─────────────────────────────
START_COL       = 1   # 從 B 欄開始（0-based），每組3欄：上班、下班、分鐘
DATE_ROW_INDEX  = 1   # 第二列（0-based=1）放日期標籤
GROUP_STRIDE    = 3   # 每組 3 欄

# ── 摘要（工作表3）需要輸出的九項欄位（名稱需與表頭一致；順序即輸出順序） ─────
SUMMARY_FIELDS = [
    "薪水",        # 基本薪水
    "特別薪水",
    "湊整",
    "小結",
    "代買",
    "預支薪水",
    "甜點&肉",
    "車馬費",
    "總計",        # 最終薪水
]
SUMMARY_NAME_COL_INDEX = 0   # 摘要表中姓名所在欄（0-based；通常在第一欄）

# ── 獎金（工作表2）欄位設定 ───────────────────────────────────────────
BONUS_FIELD     = "獎金總和"   # 主要抓取的欄名；程式亦會自動偵測關鍵字
BONUS_COL_INDEX = 19          # 若偵測失敗會回退到此欄索引（0-based；T 欄=19）

# ── PDF 用中文字型候選（reportlab 會依序嘗試註冊第一個存在的字型） ─────────
# 1) 專案內 fonts 目錄的 NotoSansTC-Regular.ttf（強烈建議：確保雲端也找得到）
# 2) Linux 常見位置（Streamlit Cloud）
# 3) Windows 常見位置（本機測試）
PDF_FONT_CANDIDATES = [
    str(BASE_DIR / "font" / "NotoSansTC-Regular.ttf"),                # 專案隨附字型（建議）
    "/usr/share/fonts/truetype/noto/NotoSansTC-Regular.ttf",           # Linux/雲端常見
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",          # 另一個 Linux 可能位置
    "C:/Windows/Fonts/msjh.ttc",                                       # 微軟正黑體（Windows）
    "C:/Windows/Fonts/msjh.ttf",
]

# 在 reportlab 中註冊後使用的字型名稱（對應 app.py 的註冊名稱）
PDF_FONT_NAME = "CJK"

# CLI 互動：預設要掃描 A 欄幾列人名
SCAN_ROWS_DEFAULT = 30
