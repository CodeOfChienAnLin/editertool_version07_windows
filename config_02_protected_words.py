"""
保護詞彙管理相關功能模組
"""
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import traceback

def load_protected_words():
    """載入詞彙保護表

    回傳:
        詞彙保護列表
    """
    # 詞彙保護檔案路徑
    protected_words_path = "protected_words.json"
    
    # 如果檔案存在，載入它
    if os.path.exists(protected_words_path):
        try:
            with open(protected_words_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except Exception as e:
            print(f"載入詞彙保護表時發生錯誤: {str(e)}")
            return []
    
    # 如果檔案不存在，返回空列表
    return []

def save_protected_words(self):
    """儲存詞彙保護表"""
    try:
        # 詞彙保護檔案路徑
        protected_words_path = "protected_words.json"
        
        # 儲存詞彙保護表
        with open(protected_words_path, 'w', encoding='utf-8') as file:
            json.dump(self.protected_words, file, ensure_ascii=False, indent=4)
        
        self.status_bar.config(text="詞彙保護表已儲存")
    except Exception as e:
        error_msg = f"儲存詞彙保護表時發生錯誤: {str(e)}"
        messagebox.showerror("錯誤", error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Protected Words Save Error", error_msg, traceback.format_exc())

def manage_protected_words(self):
    """管理保護詞彙的視窗"""
    # 創建管理視窗
    manage_window = tk.Toplevel(self.root)
    manage_window.title("管理保護詞彙")
    manage_window.geometry("400x500")
    manage_window.transient(self.root)  # 設置為主窗口的子窗口
    
    # 說明標籤
    tk.Label(manage_window, text="保護詞彙不會被自動校正，例如專有名詞或特殊用語", pady=10).pack()
    
    # 建立框架
    list_frame = tk.Frame(manage_window)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    # 詞彙列表
    tk.Label(list_frame, text="目前保護詞彙:").pack(anchor=tk.W)
    
    # 添加滾動條
    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # 列表框
    word_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, height=15)
    word_listbox.pack(fill=tk.BOTH, expand=True)
    
    # 設置滾動條命令
    scrollbar.config(command=word_listbox.yview)
    
    # 填充列表
    for word in self.protected_words:
        word_listbox.insert(tk.END, word)
    
    # 輸入框架
    input_frame = tk.Frame(manage_window)
    input_frame.pack(fill=tk.X, padx=10, pady=5)
    
    # 輸入標籤
    tk.Label(input_frame, text="新增詞彙:").pack(side=tk.LEFT)
    
    # 輸入框
    word_entry = tk.Entry(input_frame)
    word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    # 添加詞彙
    def add_word():
        word = word_entry.get().strip()
        if word:
            if word not in self.protected_words:
                self.protected_words.append(word)
                word_listbox.insert(tk.END, word)
                word_entry.delete(0, tk.END)
            else:
                messagebox.showinfo("提示", f"詞彙 '{word}' 已存在於保護列表中")
    
    # 添加按鈕
    add_button = tk.Button(input_frame, text="新增", command=add_word)
    add_button.pack(side=tk.LEFT, padx=5)
    
    # 移除詞彙
    def remove_word():
        selected = word_listbox.curselection()
        if selected:
            index = selected[0]
            word = word_listbox.get(index)
            self.protected_words.remove(word)
            word_listbox.delete(index)
    
    # 按鈕框架
    button_frame = tk.Frame(manage_window)
    button_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # 移除按鈕
    remove_button = tk.Button(button_frame, text="移除選中詞彙", command=remove_word)
    remove_button.pack(side=tk.LEFT)
    
    # 下載JSON
    def download_json():
        file_path = filedialog.asksaveasfilename(
            title="下載保護詞彙",
            defaultextension=".json",
            filetypes=[("JSON檔案", "*.json"), ("所有檔案", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    json.dump(self.protected_words, file, ensure_ascii=False, indent=4)
                messagebox.showinfo("成功", f"保護詞彙已下載到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"下載保護詞彙時發生錯誤:\n{str(e)}")
    
    # 上傳JSON
    def upload_json():
        file_path = filedialog.askopenfilename(
            title="上傳保護詞彙",
            filetypes=[("JSON檔案", "*.json"), ("所有檔案", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    words = json.load(file)
                
                # 檢查是否為列表
                if not isinstance(words, list):
                    messagebox.showerror("錯誤", "檔案格式不正確，應為詞彙列表")
                    return
                
                # 更新保護詞彙
                self.protected_words = words
                
                # 更新列表框
                word_listbox.delete(0, tk.END)
                for word in self.protected_words:
                    word_listbox.insert(tk.END, word)
                
                messagebox.showinfo("成功", "保護詞彙已上傳")
            except Exception as e:
                messagebox.showerror("錯誤", f"上傳保護詞彙時發生錯誤:\n{str(e)}")
    
    # 下載按鈕
    download_button = tk.Button(button_frame, text="下載為JSON", command=download_json)
    download_button.pack(side=tk.RIGHT, padx=5)
    
    # 上傳按鈕
    upload_button = tk.Button(button_frame, text="從JSON上傳", command=upload_json)
    upload_button.pack(side=tk.RIGHT, padx=5)
    
    # 確定按鈕
    ok_button = tk.Button(manage_window, text="確定", command=lambda: [save_protected_words(self), manage_window.destroy()])
    ok_button.pack(pady=10)
    
    # 綁定回車鍵
    word_entry.bind("<Return>", lambda event: add_word())
    
    # 綁定刪除鍵
    word_listbox.bind("<Delete>", lambda event: remove_word())
    
    # 設置焦點
    word_entry.focus_set()
