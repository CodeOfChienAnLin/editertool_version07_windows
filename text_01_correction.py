"""
文字校正相關功能模組
"""
import tkinter as tk
import threading
import traceback
from tkinter import messagebox

def correct_text(self):
    """校正文字內容"""
    text = self.text_area.get("1.0", tk.END)
    if not text.strip():
        messagebox.showinfo("提示", "沒有文字需要校正")
        return

    # 更新狀態欄
    self.status_bar.config(text="正在校正文字...")
    
    # 禁用校正按鈕，防止重複點擊
    self.correct_button.config(state=tk.DISABLED)
    
    # 在背景執行校正，避免凍結UI
    threading.Thread(target=lambda: correct_text_thread(self, text), daemon=True).start()

def correct_text_thread(self, text):
    """在背景執行文字校正的執行緒

    參數:
        text: 要校正的文字
    """
    try:
        # 檢查是否有轉換器
        if not self.converter:
            # 在主線程中顯示錯誤訊息
            self.root.after(0, lambda: messagebox.showerror("錯誤", "OpenCC轉換器未初始化"))
            self.root.after(0, lambda: self.status_bar.config(text="校正失敗: OpenCC轉換器未初始化"))
            self.root.after(0, lambda: self.correct_button.config(state=tk.NORMAL))
            return

        # 使用OpenCC進行簡繁轉換
        corrected_text = self.converter.convert(text)
        
        # 找出差異
        corrections = []
        find_differences(self, text, corrected_text, corrections)
        
        # 在主線程中更新UI
        self.root.after(0, lambda: _update_text_area(self, corrected_text, corrections))
        self.root.after(0, lambda: self.status_bar.config(text=f"文字校正完成，找到 {len(corrections)} 處差異"))
        self.root.after(0, lambda: self.correct_button.config(state=tk.NORMAL))
        
    except Exception as e:
        error_msg = f"校正文字時發生錯誤: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        
        # 在主線程中顯示錯誤訊息
        self.root.after(0, lambda: messagebox.showerror("錯誤", error_msg))
        self.root.after(0, lambda: self.status_bar.config(text="校正失敗"))
        self.root.after(0, lambda: self.correct_button.config(state=tk.NORMAL))
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Text Correction Error", error_msg, traceback.format_exc())

def find_differences(self, original_text, corrected_text, corrections, offset=0):
    """找出原始文本和校正後文本的差異

    參數:
        original_text: 原始文本
        corrected_text: 校正後的文本
        corrections: 用於存儲差異位置的列表
        offset: 文本在整體文本中的偏移量
    """
    if original_text == corrected_text:
        return

    # 檢查是否有保護詞彙
    for word in self.protected_words:
        if word in original_text:
            # 分割文本在保護詞彙處
            parts = original_text.split(word)
            corrected_parts = []
            
            # 校正每個部分，但保留保護詞彙
            current_offset = offset
            for i, part in enumerate(parts):
                if i > 0:
                    # 保護詞彙的偏移量
                    current_offset += len(parts[i-1])
                    # 添加保護詞彙（不校正）
                    corrected_parts.append(word)
                    current_offset += len(word)
                
                # 校正當前部分
                if part:
                    corrected_part = self.converter.convert(part)
                    corrected_parts.append(corrected_part)
                    # 遞歸查找差異
                    find_differences(self, part, corrected_part, corrections, current_offset)
                    current_offset += len(part)
            
            return
    
    # 如果沒有保護詞彙，則直接比較字符
    for i, (orig_char, corr_char) in enumerate(zip(original_text, corrected_text)):
        if orig_char != corr_char:
            # 添加差異位置
            start = offset + i
            corrections.append((start, start + 1))

def _update_text_area(self, corrected_text, corrections=None):
    """更新文字區域的內容

    參數:
        corrected_text: 校正後的文字
        corrections: 修正的位置列表，每個元素是 (start, end) 元組
    """
    # 清除現有標記
    self.text_area.tag_remove("corrected", "1.0", tk.END)
    
    # 更新文字
    current_text = self.text_area.get("1.0", tk.END)
    if current_text.strip() != corrected_text.strip():
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", corrected_text)
    
    # 標記修正的部分
    if corrections:
        for start, end in corrections:
            # 將字符偏移量轉換為tkinter的行列格式
            start_line = 1
            start_col = 0
            for i in range(start):
                if i < len(corrected_text) and corrected_text[i] == '\n':
                    start_line += 1
                    start_col = 0
                else:
                    start_col += 1
            
            end_line = start_line
            end_col = start_col + (end - start)
            
            # 添加標記
            start_index = f"{start_line}.{start_col}"
            end_index = f"{end_line}.{end_col}"
            self.text_area.tag_add("corrected", start_index, end_index)
