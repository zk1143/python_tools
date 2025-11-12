import pandas as pd
import os
import shutil

# ================== 設定區 ==================
EXCEL_FILE = "demo_report.xlsx"  # 改成你的 Excel 名稱
OUTPUT_DIR = "data"
PHP_COPY_TO = None  # 設路徑 = 自動複製到 PHP 專案

def format_percentage(df):
    """自動把 0.75 變 75.00%"""
    for col in df.columns:
        if any(k in str(col) for k in ['率', '滲透', '達成', '平均']):
            df[col] = df[col].apply(
                lambda x: f"{float(x)*100:.2f}%" if pd.notna(x) and isinstance(x, (int, float)) else x
            )
    return df

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"找不到 {EXCEL_FILE}！請放進資料夾")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"正在讀取 {EXCEL_FILE}...")

    try:
        all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, engine='openpyxl')
        for sheet_name, df in all_sheets.items():
            df = format_percentage(df)
            safe_name = "".join(c if c.isalnum() or c in " _-()" else "_" for c in sheet_name)
            csv_path = os.path.join(OUTPUT_DIR, f"{safe_name}.csv")
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"已存：{safe_name}.csv")

        print(f"\n成功！CSV 在 {OUTPUT_DIR}/ 資料夾")

        if PHP_COPY_TO and os.path.exists(PHP_COPY_TO):
            shutil.rmtree(PHP_COPY_TO)
            shutil.copytree(OUTPUT_DIR, PHP_COPY_TO)
            print("已複製到 PHP 專案！")

    except Exception as e:
        print(f"錯誤：{e}")

if __name__ == "__main__":
    main()
    input("\n按 Enter 結束...")