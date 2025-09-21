#員工薪資匯出+PDF匯出
project/
├─ salary8.xlsx # 出勤資料（可更換）
├─ requirements.txt
├─ config.py # 集中設定（檔名、工作表、欄位步進、日期列…）
├─ attendance_utils.py # 解析與共用邏輯
├─ test.py # 命令列互動查詢（CLI）
└─app.py # Streamlit UI + PDF 下載

命令列
CLI:python test.py
UI:python -m streamlit run app.py
