
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

# 全域模型
tokenizer = None
model = None
model_ready = False

def load_model_with_progress(on_finish_callback=None, update_status=None):
    def background_task():
        global tokenizer, model, model_ready
        try:
            if update_status:
                update_status("正在載入 tokenizer...")
            tokenizer = AutoTokenizer.from_pretrained("google/flan-t5-small")

            if update_status:
                update_status("正在載入模型...")
            model = AutoModelForSeq2SeqLM.from_pretrained("google/flan-t5-small")

            model_ready = True
            if on_finish_callback:
                on_finish_callback()
        except Exception as e:
            if update_status:
                update_status(f"[模型載入失敗] {e}")
    threading.Thread(target=background_task).start()

def read_excel_to_text(filepath: str) -> str:
    try:
        if filepath.lower().endswith('.xls'):
            df = pd.read_excel(filepath, engine='xlrd')
        else:
            df = pd.read_excel(filepath)
        return df.to_csv(index=False)
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
            temperature=0,
        )
    result = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return result.strip()

def process_excel_file(filepath: str) -> typing.Tuple[str, typing.Optional[dict], typing.Optional[str]]:
    csv_text = read_excel_to_text(filepath)
    if csv_text.startswith("[ERROR]"):
        return filepath, None, csv_text
    prompt = create_prompt(csv_text)
    llm_output = ask_llm(prompt)
    try:
        data = json.loads(llm_output)
        return filepath, data, None
    except json.JSONDecodeError:
        return filepath, None, f"[PARSE ERROR] 模型輸出非 JSON：{llm_output}"

def set_status(text):
    progress_label.config(text=text)
    root.update_idletasks()

def on_model_ready():
    set_status("✅ 模型已載入完成，可以開始處理")

def run_processing():
    if not model_ready:
        output_box.insert(tk.END, "尚未完成模型載入，請稍候...")
        return

    folder = folder_entry.get()
    if not os.path.isdir(folder):
        messagebox.showerror("錯誤", "請選擇一個有效的資料夾")
        return

    output_box.delete(1.0, tk.END)
    output_box.insert(tk.END, f"開始處理資料夾：{folder}")

    files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
    full_paths = [os.path.join(folder, f) for f in files]
    total = len(full_paths)
    if total == 0:
        output_box.insert(tk.END, "沒有找到任何 Excel 檔案")
        return

    progress_bar["maximum"] = total
    progress_bar["value"] = 0
    root.update_idletasks()

    results = []
    for i, path in enumerate(full_paths, 1):
        filepath, data, error = process_excel_file(path)
        results.append((filepath, data, error))

        output_box.insert(tk.END, f"\n檔案：{os.path.basename(path)}\n")
        if data:
            output_box.insert(tk.END, json.dumps(data, indent=2, ensure_ascii=False) + "\n")
        else:
            output_box.insert(tk.END, f"錯誤：{error}\n")

        progress_bar["value"] = i
        progress_label.config(text=f"{i}/{total} 已完成")
        root.update_idletasks()

def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder)

# GUI
root = tk.Tk()
root.title("報價單 Excel 分析工具")
root.geometry("800x650")

tk.Label(root, text="選擇報價單資料夾：").pack(pady=5)
folder_frame = tk.Frame(root)
folder_frame.pack(pady=5)
folder_entry = tk.Entry(folder_frame, width=70)
folder_entry.pack(side=tk.LEFT, padx=5)
tk.Button(folder_frame, text="瀏覽", command=select_folder).pack(side=tk.LEFT)

tk.Button(root, text="開始處理", command=run_processing).pack(pady=10)

output_box = scrolledtext.ScrolledText(root, width=100, height=25)
output_box.pack(pady=10)

progress_frame = tk.Frame(root)
progress_frame.pack(pady=5)
progress_bar = ttk.Progressbar(progress_frame, length=600, mode='determinate')
progress_bar.pack(side=tk.LEFT, padx=10)
progress_label = tk.Label(progress_frame, text=" 模型初始化中...")
progress_label.pack(anchor='w')

# 啟動模型載入
load_model_with_progress(on_finish_callback=on_model_ready, update_status=set_status)

root.mainloop()
