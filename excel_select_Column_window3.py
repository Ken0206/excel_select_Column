# 2018/11/1
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import sys

def run_convert():
    f_name = filedialog.askopenfilename(initialdir=os.getcwd(), filetypes=[("Excel", "*.xlsx"), ("All files", "*")])

    try:
        df = pd.read_excel(f_name)
    except FileNotFoundError:
        sys.exit()

    df = df[['客戶公司名稱', '事件服務編號', '處理進度', '報修時間', '客戶姓名', '客戶系統別', '系統別描述', '故障類型', '問題子類別', 'SSR Name', '問題描述',
             '備註/CAG處理過程', '服務之目的或問題徵候之描述', '另約時間訊息', '服務或測試結果', '是否需要追蹤或注意事項', '客戶交辦事項/意見', '客戶Email', '機器SN', '機器類型',
             '機器型號', 'PMR No.', 'LPAR', '相關Lpar', '軟體名稱', '服務型式', 'PMR目前狀態', 'Defect', 'AP Problem', 'Severity Level',
             '案件創建時間', '創建者', '案件修改時間', '修改者', 'OPN操作時間點', 'OPN操作人', 'ACS操作時間點', 'ACS操作人', 'CLS操作時間點', 'CLS操作人',
             'CAN操作時間點', 'CAN操作人']]
    s_f_name = filedialog.asksaveasfilename(initialdir=os.getcwd(), defaultext=".xlsx",
                                            filetypes=[("Excel", "*.xlsx"), ("All files", "*")])

    try:
        df.to_excel(s_f_name, sheet_name='output', index=False)  # .xlsx 使用 xlsxwriter package, 需安裝)

    except ValueError:
        sys.exit()
    sys.exit()

def exit_convert():
    sys.exit()

window = tk.Tk()
window.title('轉換 Excel 內容')
tk.Label(window, text="download.xls 原檔內容是 html 格式，\n"
                      "請用 Excel 開啟後另存為 xlsx 格式，\n"
                      "再使用本程式轉換資料。", font=('', 18)).pack(padx=10, pady=10)
tk.Button(window, text='執行轉換', font=('', 20), command=run_convert).pack()
tk.Button(window, text='結束轉換', font=('', 20), command=exit_convert).pack(pady=10)
window.mainloop()
