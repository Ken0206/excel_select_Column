import tkinter as tk

def hit_me():
    exit()

window = tk.Tk()
window.title('找不到 sourcr.xlsx !')
L0 = tk.Label(window, text="找不到 sourcr.xlsx !", font=('', 18))
L0.grid(row=0, column=0, padx=20, pady=25)
L1 = tk.Label(window, text="download.xls 原檔內容是 html 格式，", font=('', 18))
L1.grid(row=1, column=0, padx=20, pady=0, sticky="w")
L2 = tk.Label(window, text="請用 excel 開啟後另存為 source.xlsx", font=('', 18))
L2.grid(row=2, column=0, padx=20, pady=0, sticky="w")
L3 = tk.Label(window, text="並與本程式放在同一個目錄，", font=('', 18))
L3.grid(row=3, column=0, padx=20, pady=0, sticky="w")
L4 = tk.Label(window, text="再重新執行本程式。", font=('', 18))
L4.grid(row=4, column=0, padx=20, pady=0, sticky="w")
L4 = tk.Label(window, text="本程式輸出檔案為 output.xlsx", font=('', 18))
L4.grid(row=5, column=0, padx=20, pady=20)
b1 = tk.Button(window, text='結束', font=('', 18), command=hit_me)
b1.grid(row=6, column=0,padx=20, pady=10)
window.mainloop()
