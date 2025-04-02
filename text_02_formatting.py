"""
文字格式化相關功能模組
"""
import tkinter as tk

def adjust_indentation(self, event=None):
    """調整文字縮進，使換行後的文字對齊前一行的第一個字"""
    # 重置 Modified 標記
    self.text_area.edit_modified(False)
    
    # 獲取當前游標位置
    current_pos = self.text_area.index(tk.INSERT)
    
    # 檢查是否剛按下 Enter 鍵
    try:
        # 獲取游標前一個字符
        prev_char = self.text_area.get(f"{current_pos} - 1 chars", current_pos)
        
        # 如果前一個字符是換行符
        if prev_char == '\n':
            adjust_text_formatting(self)
    except:
        pass  # 忽略可能的錯誤
    
    return "break"  # 防止事件進一步傳播

def adjust_text_formatting(self, event=None):
    """調整文字格式，包括縮進和對齊"""
    # 獲取當前游標位置
    current_pos = self.text_area.index(tk.INSERT)
    
    # 獲取當前行號
    current_line = int(current_pos.split('.')[0])
    
    # 如果不是第一行，獲取前一行
    if current_line > 1:
        prev_line = current_line - 1
        prev_line_content = self.text_area.get(f"{prev_line}.0", f"{prev_line}.end")
        
        # 計算前一行的縮進
        indent = 0
        for char in prev_line_content:
            if char.isspace():
                indent += 1
            else:
                break
        
        # 如果前一行有縮進，應用到當前行
        if indent > 0:
            # 獲取當前行內容
            current_line_content = self.text_area.get(f"{current_line}.0", f"{current_line}.end")
            
            # 計算當前行的縮進
            current_indent = 0
            for char in current_line_content:
                if char.isspace():
                    current_indent += 1
                else:
                    break
            
            # 如果當前行縮進不足，添加空格
            if current_indent < indent:
                spaces_to_add = ' ' * (indent - current_indent)
                self.text_area.insert(f"{current_line}.0", spaces_to_add)
                
                # 更新游標位置
                self.text_area.mark_set(tk.INSERT, f"{current_line}.{indent}")
    
    return "break"  # 防止事件進一步傳播
