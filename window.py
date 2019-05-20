import tkinter as tk
#import time
#time.sleep(1)

def hit_me():
    exit()

window = tk.Tk()
window.title('找不到 sourcr.xlsx !')
#window.geometry('200x100')

#Lab1 = tk.Label(window,
#    text=" download.xls 原檔內容是 html 格式 ",     # 标签的文字
#    #bg='green',     # 背景颜色
#    font=('Arial', 18),     # 字体和字体大小
#    #width=15, height=2  # 标签长宽
#    )
#Lab1.pack()    # 固定窗口位置


L0 = tk.Label(window, text="找不到 sourcr.xlsx !", font=('', 18))
L0.pack(padx=20, pady=25)
L1 = tk.Label(window, text="download.xls 原檔內容是 html 格式", font=('', 18))
L1.pack(padx=20, pady=0)
L2 = tk.Label(window, text="請用 excel 開啟後另存為 source.xlsx", font=('', 18))
L2.pack(padx=20, pady=0)
L3 = tk.Label(window, text="並與本程式放在同一個目錄", font=('', 18))
L3.pack(padx=20, pady=0)
L4 = tk.Label(window, text="再重新執行本程式", font=('', 18))
L4.pack(padx=20, pady=0)
L4 = tk.Label(window, text="本程式輸出檔案為 output.xlsx", font=('', 18))
L4.pack(padx=20, pady=20)
b1 = tk.Button(window, text='結束', font=('', 18), command=hit_me)
b1.pack(padx=20, pady=10)
window.mainloop()
