"""
主程式入口點和主要UI框架
"""
import os
import sys
import tkinter as tk
import traceback
import logging
from tkinter import messagebox
import platform

# --- COM 相關匯入和檢查 (用於 Windows + Word 解析) ---
HAS_PYWIN32 = False
if platform.system() == 'Windows': # COM 只在 Windows 上可用
    try:
        import win32com.client as win32
        import pythoncom # 需要初始化 COM
        HAS_PYWIN32 = True
    except ImportError:
        print("警告：未找到 pywin32 模組。COM 功能需要 Windows、Microsoft Word 和 'pip install pywin32'。")
        pass # 即使沒有 pywin32，程式仍可嘗試啟動 (但功能受限)

# 導入自定義模組
from ui_02_widgets import TextCorrectionTool

def main():
    """程式主入口點"""
    try:
        # 嘗試使用 TkinterDnD2 創建支援拖放的根視窗
        try:
            # 延遲導入 tkinterdnd2，僅在需要時導入
            from tkinterdnd2 import TkinterDnD
            root = TkinterDnD.Tk()
            print("成功使用 TkinterDnD2 初始化根視窗")
        except Exception as e:
            print(f"無法使用 TkinterDnD2: {str(e)}")
            # 退回使用普通的 Tk
            root = tk.Tk()
            print("使用普通 Tk 初始化根視窗")

        app = TextCorrectionTool(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"程式執行時發生嚴重錯誤: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        # 嘗試記錄錯誤
        try:
            logging.error(error_msg)
        except Exception as log_e:
            print(f"記錄錯誤時也發生錯誤: {log_e}")
        # 嘗試顯示完整錯誤訊息框
        try:
            # 創建一個新的Tk窗口來顯示錯誤
            error_root = tk.Tk()
            error_root.title("嚴重錯誤")
            error_root.geometry("800x600")
            
            # 創建一個可滾動的文本區域來顯示完整錯誤
            frame = tk.Frame(error_root)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 添加錯誤圖標和標題
            error_frame = tk.Frame(frame)
            error_frame.pack(fill=tk.X, pady=(0, 10))
            
            # 使用Unicode字符作為錯誤圖標
            tk.Label(error_frame, text="⚠️", font=("Arial", 24), fg="red").pack(side=tk.LEFT, padx=(0, 10))
            tk.Label(error_frame, text="程式執行時發生嚴重錯誤", font=("Arial", 14, "bold")).pack(side=tk.LEFT)
            
            # 創建可滾動文本區域
            text_frame = tk.Frame(frame)
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = tk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            error_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
            error_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=error_text.yview)
            
            # 插入錯誤信息
            error_text.insert(tk.END, f"錯誤信息: {str(e)}\n\n")
            error_text.insert(tk.END, "詳細堆疊追蹤:\n")
            error_text.insert(tk.END, traceback.format_exc())
            
            # 添加關閉按鈕
            tk.Button(frame, text="關閉", command=error_root.destroy, width=15).pack(pady=10)
            
            error_root.mainloop()
        except Exception as msg_e:
            print(f"顯示錯誤訊息框時也發生錯誤: {msg_e}")
            # 最後嘗試使用基本的messagebox
            try:
                messagebox.showerror("嚴重錯誤", f"程式執行時發生嚴重錯誤:\n{str(e)}\n\n{traceback.format_exc()[:500]}...\n(錯誤訊息已截斷，完整訊息請查看控制台)")
            except:
                pass

if __name__ == "__main__":
    main()
