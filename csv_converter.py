import tkinter as tk
from tkinter import ttk, filedialog
import csv
import os
import sys
import subprocess
import chardet
import zipfile
import windnd

def detect_encoding(file_path):
    # 嘗試使用 chardet 來檢測編碼
    try:
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
        return result['encoding']

    # 如果 chardet 無法檢測，則依次使用 'latin-1' 和 'ANSI' 作為預設
    except Exception as e:
        print(f"Error detecting encoding: {e}")
        for encoding in ['latin-1', 'ANSI']:
            try:
                with open(file_path, 'r', encoding=encoding, errors='replace') as file:
                    lines = file.readlines()
                return encoding
            except Exception as e:
                print(f"Error using {encoding} encoding: {e}")

    # 如果以上嘗試都失敗，返回預設編碼 'utf-8'
    return 'utf-8'

def on_click_compress(*args):
    if compress_zip_var.get():
        remove_duplicates_var.set(True)
    else:
        remove_duplicates_var.set(False)

def on_remove_duplicates_checked(*args):
    if remove_duplicates_var.get():
        open_csv_var.set(False)
        open_csv_checkbox.config(state="disabled")
    else:
        open_csv_checkbox.config(state="normal")
        compress_zip_var.set(False)
        open_csv_var.set(True)
        open_csv_checkbox.config(state="normal")
        
def on_click_csv(*args):
    if open_csv_var.get():
        remove_duplicates_var.set(False)
        compress_zip_var.set(False)
        remove_duplicates_checkbox.config(state="disabled")
        compress_zip_checkbox.config(state="disabled")
    else:
        remove_duplicates_var.set(True)
        compress_zip_var.set(True)
        remove_duplicates_checkbox.config(state="normal")
        compress_zip_checkbox.config(state="normal")
        
def is_valid_phone_number(phone_number):
    # 忽略前後引號，再檢查電話號碼是否以 "09" 開頭且總共有10位數字
    return phone_number.strip("'").startswith("09") and len(phone_number.strip("'")) == 10
        
def convert_to_csv(input_paths, remove_duplicates, open_csv, compress_zip):
    duplicate_data_list = []
    invalid_format_list = []
    merged_output = ""  # 用於合併檔案內容的變數
    merged_csv_outPut = []
    phone_numbers_set = set()
    
    # 生成檔案名稱
    output_file_name = 'sms_distlist'
    output_file_extension = '.csv'
    
    if remove_duplicates:
        output_file_extension = '.txt'

    output_file_path = output_file_name + output_file_extension
    
    for input_path in input_paths:
        if os.path.exists(input_path):
            encoding = detect_encoding(input_path)

            with open(input_path, 'r', encoding=encoding,errors='replace') as file:
                lines = file.readlines()
            
            # # 獲取檔案基本名稱
            # base_name = os.path.splitext(os.path.basename(input_path))[0]

            if remove_duplicates:
                for line in lines:
                    data = line.strip().split(',')
                    phone_number = data[0].strip()
                    if is_valid_phone_number(phone_number) and phone_number not in phone_numbers_set:
                        merged_output += line
                        phone_numbers_set.add(phone_number)
                    elif not is_valid_phone_number(phone_number):
                        invalid_format_list.append(data)
                    else:
                        duplicate_data_list.append(data)
            else:
                for line in lines:
                    merged_csv_outPut.append(line)
        
        else:
            status_label.config(text=f"檔案 {input_path} 不存在。")
            
    
    if open_csv:
        # 寫入排序後的數據到新的 CSV 文件
        with open('sms_distlist.csv', 'w', newline='', encoding='utf-8-sig') as output_file:
            csv_writer = csv.writer(output_file)
            
            # 將數據存入列表
            data = [line.strip().split(',') for line in merged_csv_outPut]
                
            # 按照第一列（假設是電話號碼）進行排序
            sorted_data = sorted(data, key=lambda x: x[0])
            
            for line in sorted_data:
                csv_writer.writerow(line)
            
            subprocess.run(["start", "excel", output_file_name + output_file_extension], shell=True)
            
    else:
        # 將合併的內容寫入新檔案
        with open(output_file_path, 'w', encoding='utf-8') as merged_file:
            merged_file.write(merged_output)

    if compress_zip:
            
        with zipfile.ZipFile(output_file_name + '.zip', 'w') as zip_file:
            zip_file.write(output_file_path, os.path.basename(output_file_path))

            os.remove(output_file_path)  # 刪除原始檔案
            
    deleted_data_lists = {
        "重複的號碼": duplicate_data_list,
        "格式錯誤的號碼": invalid_format_list
    }
    
    if remove_duplicates and not open_csv:
        if duplicate_data_list:
            status_label.config(text="轉換完成，已刪除重複的資料。")
            show_deleted_data(deleted_data_lists)
        elif invalid_format_list:
            status_label.config(text="轉換完成，格式錯誤的資料。")
            show_deleted_data(deleted_data_lists)
        else:
            status_label.config(text="轉換完成，沒有重複的資料。")
    elif open_csv:
            status_label.config(text="轉換完成。")

def show_deleted_data(deleted_data_lists):
    deleted_data_window = tk.Toplevel()
    deleted_data_window.title("已刪除的資料")

    deleted_data_text = tk.Text(deleted_data_window, wrap='word')
    deleted_data_text.grid(row=0, column=0, sticky='nsew')

    # 使 Text 區域可以跟著視窗一起縮放
    deleted_data_window.rowconfigure(0, weight=1)
    deleted_data_window.columnconfigure(0, weight=1)

    # 將已刪除的資料插入 Text 元件中，根據不同的類別顯示
    for category, deleted_data_list in deleted_data_lists.items():
        deleted_data_text.insert(tk.END, f"=== {category} ===\n")
        for data in deleted_data_list:
            deleted_data_text.insert(tk.END, ', '.join(data) + '\n')

    # 使視窗保持開啟狀態
    deleted_data_window.mainloop() 

def on_drop(files):
    # 更新 Entry 中的內容，將所有檔案路徑以分號隔開
    file_paths = ";".join(file.decode('utf-8') if isinstance(file, bytes) else file for file in files)
    path_entry.delete(0, tk.END)
    path_entry.insert(tk.END, file_paths)
    
def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt")])
    path_entry.delete(0, tk.END)

    for file_path in file_paths:
        path_entry.insert(tk.END, file_path + ";")


def convert_button_clicked():
    input_paths = path_entry.get().split(';')
    input_paths = [path.strip() for path in input_paths if path.strip()]
    remove_duplicates = remove_duplicates_var.get()
    open_csv = open_csv_var.get()
    compress_zip = compress_zip_var.get()

     # 如果同時選擇了刪除重複，取消 CSV 的生成並禁用打勾
    if remove_duplicates:
        open_csv = False
        open_csv_var.set(False)
        open_csv_checkbox.config(state="disabled")
    else:
        open_csv = True
        open_csv_var.set(True)
        open_csv_checkbox.config(state="active")

    convert_to_csv(input_paths, remove_duplicates, open_csv, compress_zip)


root = tk.Tk()
root.title("維尼 作弊程式")

# Entry 元件，用於顯示或輸入檔案路徑
windnd.hook_dropfiles(root,func=on_drop)



# 設定窗口大小縮放
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

path_label = tk.Label(root, text="輸入檔案路徑:")
path_label.grid(row=0, column=0, pady=10, padx=10, sticky="w")

path_entry = tk.Entry(root, width=50)
path_entry.grid(row=0, column=1, pady=10, padx=10, sticky="ew")

browse_button = tk.Button(root, text="瀏覽", command=browse_files)
browse_button.grid(row=0, column=2, pady=10, padx=10)

remove_duplicates_var = tk.BooleanVar(value=True)
remove_duplicates_var.trace_add("write", on_remove_duplicates_checked)
remove_duplicates_checkbox = tk.Checkbutton(root, text="刪除重複電話號碼", variable=remove_duplicates_var)
remove_duplicates_checkbox.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="w")

open_csv_var = tk.BooleanVar()
open_csv_var.trace_add("write", on_click_csv)
open_csv_checkbox = tk.Checkbutton(root, text="自動開啟 CSV", variable=open_csv_var)
open_csv_checkbox.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="w")

compress_zip_var = tk.BooleanVar(value=True)
compress_zip_var.trace_add("write",on_click_compress)
compress_zip_checkbox = tk.Checkbutton(root, text="壓縮成 ZIP", variable=compress_zip_var)
compress_zip_checkbox.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="w")

convert_button = tk.Button(root, text="轉換", command=convert_button_clicked)
convert_button.grid(row=4, column=0, columnspan=3, pady=10)

status_label = tk.Label(root, text="")
status_label.grid(row=5, column=0, columnspan=3, pady=10)

# 加入 Sizegrip 元件以支援窗口大小縮放
sizegrip = ttk.Sizegrip(root)
sizegrip.grid(row=999, column=999, sticky="se")

root.mainloop()