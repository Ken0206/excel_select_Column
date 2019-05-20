# 2018/10/29
import pandas as pd
import tkinter as tk
import sys

def hit_me():
    sys.exit()

try:
    df = pd.read_excel("source.xlsx")
except FileNotFoundError:
    window = tk.Tk()
    window.title('找不到 source.xlsx !')
    L0 = tk.Label(window, text="\n找不到 source.xlsx !\n\n"
                               "download.xls 原檔內容是 html 格式，\n"
                               "請用 excel 開啟後另存為 source.xlsx\n"
                               "並與本程式放在同一個目錄，\n"
                               "再重新執行本程式。\n\n"
                               "本程式輸出檔案為 output.xlsx", font=('', 18))
    L0.pack(padx=20, pady=20)
    b1 = tk.Button(window, text='結束', font=('', 18), command=hit_me)
    b1.pack(padx=10, pady=10)
    window.mainloop()

df = df[['客戶公司名稱', '事件服務編號', '處理進度', '報修時間', '客戶姓名', '客戶系統別', '系統別描述', '故障類型', '問題子類別', 'SSR Name', '問題描述', '備註/CAG處理過程', '服務之目的或問題徵候之描述', '另約時間訊息', '服務或測試結果', '是否需要追蹤或注意事項', '客戶交辦事項/意見', '客戶Email', '機器SN', '機器類型', '機器型號', 'PMR No.', 'LPAR', '相關Lpar', '軟體名稱', '服務型式', 'PMR目前狀態', 'Defect', 'AP Problem', 'Severity Level', '案件創建時間', '創建者', '案件修改時間', '修改者', 'OPN操作時間點', 'OPN操作人', 'ACS操作時間點', 'ACS操作人', 'CLS操作時間點', 'CLS操作人', 'CAN操作時間點', 'CAN操作人']]
df.to_excel('output.xlsx', sheet_name='output', index=False)    # .xlsx 使用 xlsxwriter package, 需安裝)
