import os
import pandas as pd
import json
import torch
from concurrent.futures import ThreadPoolExecutor, as_completed
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
import typing
import xlrd  # noqa: F401
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import tkinter.ttk as ttk
import threading
def load_model_once():
    global tokenizer, model, model_ready
    try:
        tokenizer = AutoTokenizer.from_pretrained("google/flan-t5-base")
        model = AutoModelForSeq2SeqLM.from_pretrained("google/flan-t5-base")
        model_ready = True
    except Exception as e:
        print(f"[模型載入失敗] {e}")

def read_excel_to_text(filepath: str) -> str:
    try:
        df = pd.read_excel(filepath, engine='xlrd', header=None)
        
        # 找出項目開始的列 (以 "Item" 開頭的列)
        item_start_idx = df[df.apply(lambda row: row.astype(str).str.contains("Item").any(), axis=1)].index
        if len(item_start_idx) == 0:
            return "[ERROR] 找不到 'Item' 起始行"

        item_start = item_start_idx[0] + 2  # 跳過標題與空行
        text_lines = []
        
        for i in range(item_start, len(df)):
            row = df.iloc[i]
            if isinstance(row.astype(str).str.cat(), str) and "Total:NT" in row.astype(str).str.cat():
                break  # 結束於總價行
            values = [str(cell).strip() for cell in row if pd.notnull(cell)]
            if values:
                text_lines.append(",".join(values))
        
        return "\n".join(text_lines)
    except Exception as e:
        return f"[ERROR] 無法讀取檔案 {filepath}：{e}"


def create_prompt(csv_text: str) -> str:
    return f"請根據以下報價單內容擷取欄位：報價日期、出貨對象、貨號、數量，並以 JSON 格式輸出：\n{csv_text}"

def ask_llm(prompt: str, max_tokens=512) -> str:
    inputs = tokenizer(prompt, return_tensors="pt", truncation=True, max_length=512)
    with torch.no_grad():
        outputs = model.generate(
            **inputs,
            max_new_tokens=max_tokens,
            do_sample=False,
        )
    result = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return result.strip()
file_path = r"C:\Users\Justin\Documents\GitHub\Excel-Quotation-Extractor\報價單測試資料\元誠0905043.xls"
load_model_once()
llm_result = ask_llm(read_excel_to_text(file_path))
print(read_excel_to_text(file_path))
print(llm_result)
# try:
#     # 嘗試將 LLM 輸出轉為 dict
#     data = json.loads(llm_result)
#     # 如果是單一 dict，轉成 DataFrame
#     if isinstance(data, dict):
#         df = pd.DataFrame([data])
#     # 如果是 list of dicts
#     elif isinstance(data, list):
#         df = pd.DataFrame(data)
#     else:
#         raise ValueError("LLM 輸出格式不正確")
#     # 存成 CSV
#     df.to_csv("llm_output.csv", index=False, encoding="utf-8-sig")
#     print("已存成 llm_output.csv")
# except Exception as e:
#     print(f"無法轉換或儲存 LLM 輸出：{e}")