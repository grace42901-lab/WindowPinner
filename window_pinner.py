import tkinter as tk
from tkinter import messagebox
import win32gui
import win32con

class WindowPinnerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("視窗置頂管理器")
        self.root.geometry("400x500")
        
        # 標題
        self.label = tk.Label(root, text="勾選視窗以設為「最上層顯示」", font=("Arial", 12))
        self.label.pack(pady=10)

        # 列表框 (包含捲軸)
        self.frame = tk.Frame(root)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        self.canvas = tk.Canvas(self.frame)
        self.scrollbar = tk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 重新整理按鈕
        self.refresh_btn = tk.Button(root, text="重新整理視窗列表", command=self.refresh_windows, bg="#f0f0f0")
        self.refresh_btn.pack(pady=10, fill=tk.X, padx=10)

        self.checkboxes = {}
        self.refresh_windows()

    def get_windows(self):
        """獲取所有可見視窗的標題與控制碼 (HWND)"""
        windows = []
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # 排除沒有標題的隱藏視窗
                    windows.append((hwnd, title))
        win32gui.EnumWindows(enum_handler, None)
        return windows

    def is_topmost(self, hwnd):
        """檢查視窗目前是否為置頂狀態"""
        # 取得視窗的擴展樣式
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        return (ex_style & win32con.WS_EX_TOPMOST) != 0

    def toggle_topmost(self, hwnd, var):
        """切換置頂狀態"""
        try:
            if var.get():
                # 設為置頂
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            else:
                # 取消置頂
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法設定視窗狀態 (視窗可能已關閉)\n錯誤: {e}")
            self.refresh_windows()

    def refresh_windows(self):
        """重新整理介面"""
        # 清除舊的 checkbox
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkboxes = {}

        windows = self.get_windows()
        
        # 排除自己的視窗，避免混淆
        my_hwnd = self.root.winfo_id() 

        for hwnd, title in windows:
            if str(hwnd) == str(my_hwnd): 
                continue # 跳過自己

            var = tk.BooleanVar()
            # 檢查該視窗目前是否已經置頂，若是則預設勾選
            if self.is_topmost(hwnd):
                var.set(True)

            chk = tk.Checkbutton(self.scrollable_frame, text=title, variable=var, anchor="w",
                                 command=lambda h=hwnd, v=var: self.toggle_topmost(h, v))
            chk.pack(fill=tk.X, padx=5, pady=2)
            self.checkboxes[hwnd] = var

if __name__ == "__main__":
    root = tk.Tk()
    # 讓本程式預設也置頂，方便操作
    root.attributes('-topmost', True) 
    app = WindowPinnerApp(root)
    root.mainloop()
