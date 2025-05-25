import os
import pandas as pd
from extractor import extract_header_info, parse_product_rows
import tkinter as tk
from tkinter import filedialog

# 隱藏主視窗
root = tk.Tk()
root.withdraw()

def flatten_product_records(header: dict, products: list, filename: str) -> list[dict]:
    return [
        {
            "單號": header["單號"],
            "對象": header["對象"],
            "日期": header["日期"],
            "來源檔名": filename,
            **product
        }
        for product in products
    ]

def process_all_quotation_files(folder_path: str) -> list:
    all_records = []

    for filename in os.listdir(folder_path):
        if (
        not filename.lower().endswith((".xls", ".xlsx"))
        or filename.startswith("~$")
        or "fax" in filename.lower()
        ):  # 跳過非 Excel 檔. fax 檔案和隱藏檔
            continue
            
        else:
            file_path = os.path.join(folder_path, filename)
            print(f"正在處理：{file_path}")
            try:
                df = pd.read_excel(file_path, engine="xlrd" if file_path.endswith(".xls") else None, header=None)
                header = extract_header_info(df)
                products = parse_product_rows(df)
                flat = flatten_product_records(header, products, filename)

                all_records.extend(flat)
            except Exception as e:
                print(f"[錯誤] 無法處理 {filename}：{e}")

    return all_records


folder = filedialog.askdirectory(title="請選擇要合併的資料夾")
records = process_all_quotation_files(folder)

# 匯出成 Excel
df_all = pd.DataFrame(records)
df_all.to_excel("同業報價單.xlsx", index=False)