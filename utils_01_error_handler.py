"""
錯誤處理相關功能模組
"""
import os
import sys
import logging
import traceback
import datetime
import tkinter as tk
from tkinter import messagebox

def setup_error_logging():
    """設置錯誤日誌記錄

    回傳:
        logger: 日誌記錄器
    """
    # 確保日誌目錄存在
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # 設置日誌檔案名稱（包含日期）
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"error_log_{current_date}.txt")
    
    # 設置日誌記錄器
    logger = logging.getLogger("TextCorrectionTool")
    logger.setLevel(logging.ERROR)
    
    # 檔案處理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.ERROR)
    
    # 格式化器
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    
    # 添加處理器到記錄器
    logger.addHandler(file_handler)
    
    return logger

def log_error(self, error_type, error_message, error_traceback=None):
    """記錄錯誤到日誌

    參數:
        error_type: 錯誤類型
        error_message: 錯誤訊息
        error_traceback: 錯誤追蹤（可選）
    """
    # 確保logger已設置
    if not hasattr(self, 'logger') or self.logger is None:
        self.logger = setup_error_logging()
    
    # 記錄錯誤
    log_message = f"{error_type}: {error_message}"
    self.logger.error(log_message)
    
    # 如果有錯誤追蹤，也記錄它
    if error_traceback:
        self.logger.error(f"Traceback: {error_traceback}")
    
    # 更新狀態欄
    if hasattr(self, 'status_bar'):
        self.status_bar.config(text=f"錯誤: {error_message}")

def show_error_dialog(self, title, message, error_details=None):
    """顯示錯誤對話框

    參數:
        title: 對話框標題
        message: 錯誤訊息
        error_details: 錯誤詳情（可選）
    """
    # 記錄錯誤
    log_error(self, title, message, error_details)
    
    # 創建錯誤對話框
    error_window = tk.Toplevel(self.root)
    error_window.title(f"錯誤: {title}")
    error_window.geometry("500x300")
    error_window.transient(self.root)  # 設置為主窗口的子窗口
    
    # 錯誤圖標
    try:
        # 使用內置圖標
        error_label = tk.Label(error_window, text="⚠️", font=("Arial", 48))
        error_label.pack(pady=10)
    except:
        # 如果無法顯示Unicode字符，使用文字
        error_label = tk.Label(error_window, text="ERROR", font=("Arial", 24, "bold"))
        error_label.pack(pady=10)
    
    # 錯誤訊息
    message_label = tk.Label(error_window, text=message, wraplength=450)
    message_label.pack(pady=10, padx=20)
    
    # 如果有錯誤詳情，顯示它
    if error_details:
        # 創建框架
        details_frame = tk.Frame(error_window)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 詳情標籤
        details_label = tk.Label(details_frame, text="錯誤詳情:", anchor=tk.W)
        details_label.pack(anchor=tk.W)
        
        # 詳情文字區域
        details_text = tk.Text(details_frame, height=6, width=50, wrap=tk.WORD)
        details_text.pack(fill=tk.BOTH, expand=True)
        details_text.insert(tk.END, error_details)
        details_text.config(state=tk.DISABLED)
        
        # 滾動條
        scrollbar = tk.Scrollbar(details_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        details_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=details_text.yview)
    
    # 按鈕框架
    button_frame = tk.Frame(error_window)
    button_frame.pack(fill=tk.X, pady=10)
    
    # 確定按鈕
    ok_button = tk.Button(button_frame, text="確定", command=error_window.destroy, width=10)
    ok_button.pack(side=tk.RIGHT, padx=20)
    
    # 複製按鈕
    def copy_to_clipboard():
        clipboard_text = f"{title}\n{message}"
        if error_details:
            clipboard_text += f"\n\n{error_details}"
        
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)
        self.root.update()  # 必須調用update
    
    copy_button = tk.Button(button_frame, text="複製錯誤", command=copy_to_clipboard, width=10)
    copy_button.pack(side=tk.RIGHT, padx=5)

def handle_exception(self, exc_type, exc_value, exc_traceback):
    """處理未捕獲的異常

    參數:
        exc_type: 異常類型
        exc_value: 異常值
        exc_traceback: 異常追蹤
    """
    # 獲取錯誤訊息
    error_message = str(exc_value)
    
    # 獲取錯誤追蹤
    error_traceback = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    
    # 記錄錯誤
    log_error(self, "Uncaught Exception", error_message, error_traceback)
    
    # 顯示錯誤對話框
    messagebox.showerror("未處理的異常", f"程式發生未預期的錯誤:\n{error_message}\n\n錯誤已記錄到日誌檔案。")
    
    # 返回True表示異常已處理
    return True

def setup_exception_handler(self):
    """設置全局異常處理器"""
    # 設置未捕獲異常處理器
    sys.excepthook = lambda exc_type, exc_value, exc_traceback: handle_exception(self, exc_type, exc_value, exc_traceback)

def view_error_logs(self):
    """查看錯誤日誌"""
    # 日誌目錄
    log_dir = "logs"
    
    # 檢查日誌目錄是否存在
    if not os.path.exists(log_dir):
        messagebox.showinfo("提示", "尚無錯誤日誌")
        return
    
    # 獲取所有日誌檔案
    log_files = [f for f in os.listdir(log_dir) if f.startswith("error_log_") and f.endswith(".txt")]
    
    # 如果沒有日誌檔案
    if not log_files:
        messagebox.showinfo("提示", "尚無錯誤日誌")
        return
    
    # 按日期排序（最新的在前）
    log_files.sort(reverse=True)
    
    # 創建日誌查看視窗
    log_window = tk.Toplevel(self.root)
    log_window.title("錯誤日誌")
    log_window.geometry("700x500")
    
    # 創建框架
    main_frame = tk.Frame(log_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 分割視窗
    paned_window = tk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=True)
    
    # 左側框架（檔案列表）
    left_frame = tk.Frame(paned_window, width=200)
    paned_window.add(left_frame)
    
    # 右側框架（日誌內容）
    right_frame = tk.Frame(paned_window)
    paned_window.add(right_frame)
    
    # 檔案列表標籤
    tk.Label(left_frame, text="日誌檔案:").pack(anchor=tk.W, pady=(0, 5))
    
    # 檔案列表框
    file_listbox = tk.Listbox(left_frame, width=25)
    file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 滾動條
    file_scrollbar = tk.Scrollbar(left_frame)
    file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # 設置滾動條命令
    file_listbox.config(yscrollcommand=file_scrollbar.set)
    file_scrollbar.config(command=file_listbox.yview)
    
    # 填充檔案列表
    for log_file in log_files:
        # 從檔案名稱中提取日期
        date_str = log_file.replace("error_log_", "").replace(".txt", "")
        file_listbox.insert(tk.END, date_str)
    
    # 日誌內容標籤
    tk.Label(right_frame, text="日誌內容:").pack(anchor=tk.W, pady=(0, 5))
    
    # 日誌內容文字區域
    log_text = tk.Text(right_frame, wrap=tk.WORD)
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 滾動條
    log_scrollbar = tk.Scrollbar(right_frame)
    log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # 設置滾動條命令
    log_text.config(yscrollcommand=log_scrollbar.set)
    log_scrollbar.config(command=log_text.yview)
    
    # 顯示選中檔案的內容
    def show_log_content(event):
        # 獲取選中的索引
        selected = file_listbox.curselection()
        if selected:
            index = selected[0]
            date_str = file_listbox.get(index)
            log_file = f"error_log_{date_str}.txt"
            log_path = os.path.join(log_dir, log_file)
            
            # 讀取日誌內容
            try:
                with open(log_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                
                # 清空文字區域
                log_text.config(state=tk.NORMAL)
                log_text.delete(1.0, tk.END)
                
                # 插入內容
                log_text.insert(tk.END, content)
                
                # 禁用編輯
                log_text.config(state=tk.DISABLED)
            except Exception as e:
                log_text.config(state=tk.NORMAL)
                log_text.delete(1.0, tk.END)
                log_text.insert(tk.END, f"讀取日誌時發生錯誤: {str(e)}")
                log_text.config(state=tk.DISABLED)
    
    # 綁定選擇事件
    file_listbox.bind("<<ListboxSelect>>", show_log_content)
    
    # 按鈕框架
    button_frame = tk.Frame(log_window)
    button_frame.pack(fill=tk.X, pady=10)
    
    # 關閉按鈕
    close_button = tk.Button(button_frame, text="關閉", command=log_window.destroy, width=10)
    close_button.pack(side=tk.RIGHT, padx=10)
    
    # 清除按鈕
    def clear_logs():
        # 確認對話框
        if messagebox.askyesno("確認", "確定要清除所有錯誤日誌嗎？此操作無法撤銷。"):
            try:
                # 清除所有日誌檔案
                for log_file in log_files:
                    os.remove(os.path.join(log_dir, log_file))
                
                # 關閉視窗
                log_window.destroy()
                
                # 顯示成功訊息
                messagebox.showinfo("成功", "所有錯誤日誌已清除")
            except Exception as e:
                messagebox.showerror("錯誤", f"清除日誌時發生錯誤: {str(e)}")
    
    clear_button = tk.Button(button_frame, text="清除所有日誌", command=clear_logs, width=15)
    clear_button.pack(side=tk.LEFT, padx=10)
    
    # 導出按鈕
    def export_log():
        # 獲取選中的索引
        selected = file_listbox.curselection()
        if not selected:
            messagebox.showinfo("提示", "請先選擇一個日誌檔案")
            return
        
        index = selected[0]
        date_str = file_listbox.get(index)
        log_file = f"error_log_{date_str}.txt"
        log_path = os.path.join(log_dir, log_file)
        
        # 選擇保存位置
        save_path = tk.filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文字檔案", "*.txt"), ("所有檔案", "*.*")],
            initialfile=log_file
        )
        
        if save_path:
            try:
                # 複製檔案
                with open(log_path, 'r', encoding='utf-8') as src:
                    content = src.read()
                
                with open(save_path, 'w', encoding='utf-8') as dst:
                    dst.write(content)
                
                # 顯示成功訊息
                messagebox.showinfo("成功", f"日誌已導出到:\n{save_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"導出日誌時發生錯誤: {str(e)}")
    
    export_button = tk.Button(button_frame, text="導出日誌", command=export_log, width=10)
    export_button.pack(side=tk.LEFT, padx=5)
    
    # 自動選擇第一個檔案
    if log_files:
        file_listbox.select_set(0)
        file_listbox.event_generate("<<ListboxSelect>>")
