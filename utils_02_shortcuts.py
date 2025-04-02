"""
快捷鍵管理相關功能模組
"""
import tkinter as tk
from tkinter import ttk, messagebox
import json
import traceback

def setup_shortcuts(self):
    """設置快捷鍵

    為應用程式設置預設和自定義快捷鍵
    """
    # 預設快捷鍵
    self.default_shortcuts = {
        # 檔案操作
        "<Control-o>": self.open_file,
        "<Control-s>": self.save_file,
        "<Control-Shift-S>": self.save_file_as,
        
        # 編輯操作
        "<Control-z>": lambda: self.text_area.edit_undo(),
        "<Control-y>": lambda: self.text_area.edit_redo(),
        "<Control-a>": lambda: self.select_all_text(),
        "<Control-c>": lambda: self.copy_selected_text(),
        "<Control-x>": lambda: self.cut_selected_text(),
        "<Control-v>": lambda: self.paste_text(),
        
        # 文字處理
        "<Control-r>": self.correct_text,
        "<Control-f>": self.find_text,
        "<Control-h>": self.replace_text,
        
        # 格式設定
        "<Control-b>": lambda: self.toggle_text_style("bold"),
        "<Control-i>": lambda: self.toggle_text_style("italic"),
        "<Control-u>": lambda: self.toggle_text_style("underline"),
        
        # 其他功能
        "<F1>": self.show_help,
        "<F5>": self.refresh_view,
        "<Escape>": self.cancel_operation
    }
    
    # 加載自定義快捷鍵
    self.custom_shortcuts = self.settings.get("custom_shortcuts", [])
    
    # 綁定所有快捷鍵
    self.bind_all_shortcuts()

def bind_all_shortcuts(self):
    """綁定所有快捷鍵"""
    # 先解綁所有快捷鍵
    for shortcut in self.default_shortcuts:
        try:
            self.root.unbind(shortcut)
            self.text_area.unbind(shortcut)
        except:
            pass
    
    # 綁定預設快捷鍵
    for shortcut, command in self.default_shortcuts.items():
        self.root.bind(shortcut, lambda event, cmd=command: self.execute_shortcut(event, cmd))
        self.text_area.bind(shortcut, lambda event, cmd=command: self.execute_shortcut(event, cmd))
    
    # 綁定自定義快捷鍵
    for shortcut_info in self.custom_shortcuts:
        shortcut = shortcut_info["key"]
        action = shortcut_info["action"]
        
        # 根據動作名稱獲取對應的函數
        command = self.get_command_by_name(action)
        if command:
            self.root.bind(shortcut, lambda event, cmd=command: self.execute_shortcut(event, cmd))
            self.text_area.bind(shortcut, lambda event, cmd=command: self.execute_shortcut(event, cmd))

def execute_shortcut(self, event, command):
    """執行快捷鍵命令

    參數:
        event: 事件對象
        command: 要執行的命令
    
    回傳:
        'break' 以阻止事件繼續傳播
    """
    try:
        command()
    except Exception as e:
        error_msg = f"執行快捷鍵命令時發生錯誤: {str(e)}"
        self.status_bar.config(text=error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Shortcut Error", error_msg, traceback.format_exc())
    
    return 'break'  # 阻止事件繼續傳播

def get_command_by_name(self, action_name):
    """根據動作名稱獲取對應的函數

    參數:
        action_name: 動作名稱
    
    回傳:
        對應的函數，如果找不到則返回None
    """
    # 動作名稱到函數的映射
    action_map = {
        "open_file": self.open_file,
        "save_file": self.save_file,
        "save_file_as": self.save_file_as,
        "undo": lambda: self.text_area.edit_undo(),
        "redo": lambda: self.text_area.edit_redo(),
        "select_all": self.select_all_text,
        "copy": self.copy_selected_text,
        "cut": self.cut_selected_text,
        "paste": self.paste_text,
        "correct_text": self.correct_text,
        "find_text": self.find_text,
        "replace_text": self.replace_text,
        "toggle_bold": lambda: self.toggle_text_style("bold"),
        "toggle_italic": lambda: self.toggle_text_style("italic"),
        "toggle_underline": lambda: self.toggle_text_style("underline"),
        "show_help": self.show_help,
        "refresh_view": self.refresh_view,
        "cancel_operation": self.cancel_operation
    }
    
    return action_map.get(action_name)

def manage_shortcuts(self):
    """管理快捷鍵視窗"""
    # 創建管理視窗
    shortcut_window = tk.Toplevel(self.root)
    shortcut_window.title("管理快捷鍵")
    shortcut_window.geometry("600x500")
    shortcut_window.transient(self.root)  # 設置為主窗口的子窗口
    
    # 說明標籤
    tk.Label(shortcut_window, text="在此管理應用程式的快捷鍵設定", pady=10).pack()
    
    # 建立框架
    main_frame = tk.Frame(shortcut_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    # 建立筆記本
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    # 預設快捷鍵頁面
    default_frame = tk.Frame(notebook)
    notebook.add(default_frame, text="預設快捷鍵")
    
    # 自定義快捷鍵頁面
    custom_frame = tk.Frame(notebook)
    notebook.add(custom_frame, text="自定義快捷鍵")
    
    # 填充預設快捷鍵頁面
    self.fill_default_shortcuts_tab(default_frame)
    
    # 填充自定義快捷鍵頁面
    self.fill_custom_shortcuts_tab(custom_frame)
    
    # 按鈕框架
    button_frame = tk.Frame(shortcut_window)
    button_frame.pack(fill=tk.X, pady=10)
    
    # 確定按鈕
    ok_button = tk.Button(button_frame, text="確定", command=lambda: [self.save_custom_shortcuts(), shortcut_window.destroy()], width=10)
    ok_button.pack(side=tk.RIGHT, padx=10)
    
    # 取消按鈕
    cancel_button = tk.Button(button_frame, text="取消", command=shortcut_window.destroy, width=10)
    cancel_button.pack(side=tk.RIGHT, padx=5)
    
    # 重置按鈕
    reset_button = tk.Button(button_frame, text="重置所有快捷鍵", command=lambda: self.reset_shortcuts(shortcut_window), width=15)
    reset_button.pack(side=tk.LEFT, padx=10)

def fill_default_shortcuts_tab(self, parent_frame):
    """填充預設快捷鍵頁面

    參數:
        parent_frame: 父框架
    """
    # 說明標籤
    tk.Label(parent_frame, text="以下是應用程式的預設快捷鍵，這些快捷鍵無法修改", pady=5).pack(anchor=tk.W)
    
    # 建立框架
    list_frame = tk.Frame(parent_frame)
    list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    # 建立樹狀視圖
    columns = ("shortcut", "action", "description")
    tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    
    # 設置列標題
    tree.heading("shortcut", text="快捷鍵")
    tree.heading("action", text="動作")
    tree.heading("description", text="描述")
    
    # 設置列寬
    tree.column("shortcut", width=150)
    tree.column("action", width=150)
    tree.column("description", width=250)
    
    # 添加滾動條
    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 填充數據
    shortcut_descriptions = {
        "<Control-o>": ["開啟檔案", "開啟一個檔案"],
        "<Control-s>": ["儲存檔案", "儲存目前的檔案"],
        "<Control-Shift-S>": ["另存新檔", "將檔案儲存為新檔案"],
        "<Control-z>": ["復原", "復原上一個動作"],
        "<Control-y>": ["重做", "重做上一個被復原的動作"],
        "<Control-a>": ["全選", "選取所有文字"],
        "<Control-c>": ["複製", "複製選取的文字"],
        "<Control-x>": ["剪下", "剪下選取的文字"],
        "<Control-v>": ["貼上", "貼上剪貼簿中的文字"],
        "<Control-r>": ["校正文字", "自動校正文字"],
        "<Control-f>": ["尋找", "尋找文字"],
        "<Control-h>": ["取代", "尋找並取代文字"],
        "<Control-b>": ["粗體", "切換粗體"],
        "<Control-i>": ["斜體", "切換斜體"],
        "<Control-u>": ["底線", "切換底線"],
        "<F1>": ["說明", "顯示說明"],
        "<F5>": ["重新整理", "重新整理視圖"],
        "<Escape>": ["取消", "取消目前的操作"]
    }
    
    for shortcut, (action, description) in shortcut_descriptions.items():
        tree.insert("", tk.END, values=(shortcut, action, description))

def fill_custom_shortcuts_tab(self, parent_frame):
    """填充自定義快捷鍵頁面

    參數:
        parent_frame: 父框架
    """
    # 說明標籤
    tk.Label(parent_frame, text="在此添加、修改或刪除自定義快捷鍵", pady=5).pack(anchor=tk.W)
    
    # 建立框架
    list_frame = tk.Frame(parent_frame)
    list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    # 建立樹狀視圖
    columns = ("shortcut", "action")
    self.custom_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    
    # 設置列標題
    self.custom_tree.heading("shortcut", text="快捷鍵")
    self.custom_tree.heading("action", text="動作")
    
    # 設置列寬
    self.custom_tree.column("shortcut", width=150)
    self.custom_tree.column("action", width=150)
    
    # 添加滾動條
    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.custom_tree.yview)
    self.custom_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    self.custom_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 填充數據
    for shortcut_info in self.custom_shortcuts:
        self.custom_tree.insert("", tk.END, values=(shortcut_info["key"], shortcut_info["action"]))
    
    # 按鈕框架
    button_frame = tk.Frame(parent_frame)
    button_frame.pack(fill=tk.X, pady=5)
    
    # 添加按鈕
    add_button = tk.Button(button_frame, text="新增", command=lambda: self.add_custom_shortcut(), width=10)
    add_button.pack(side=tk.LEFT, padx=5)
    
    # 編輯按鈕
    edit_button = tk.Button(button_frame, text="編輯", command=lambda: self.edit_custom_shortcut(), width=10)
    edit_button.pack(side=tk.LEFT, padx=5)
    
    # 刪除按鈕
    delete_button = tk.Button(button_frame, text="刪除", command=lambda: self.delete_custom_shortcut(), width=10)
    delete_button.pack(side=tk.LEFT, padx=5)
    
    # 綁定雙擊事件
    self.custom_tree.bind("<Double-1>", lambda event: self.edit_custom_shortcut())

def add_custom_shortcut(self):
    """添加自定義快捷鍵"""
    self.open_shortcut_dialog()

def edit_custom_shortcut(self):
    """編輯自定義快捷鍵"""
    # 獲取選中的項目
    selected = self.custom_tree.selection()
    if not selected:
        messagebox.showinfo("提示", "請先選擇一個快捷鍵")
        return
    
    # 獲取選中項目的值
    item = self.custom_tree.item(selected[0])
    values = item["values"]
    
    # 開啟對話框
    self.open_shortcut_dialog(values[0], values[1])
    
    # 刪除原項目
    self.custom_tree.delete(selected[0])

def delete_custom_shortcut(self):
    """刪除自定義快捷鍵"""
    # 獲取選中的項目
    selected = self.custom_tree.selection()
    if not selected:
        messagebox.showinfo("提示", "請先選擇一個快捷鍵")
        return
    
    # 獲取選中項目的值
    item = self.custom_tree.item(selected[0])
    values = item["values"]
    
    # 確認刪除
    if messagebox.askyesno("確認", f"確定要刪除快捷鍵 '{values[0]}' 嗎？"):
        # 從列表中刪除
        self.custom_tree.delete(selected[0])

def open_shortcut_dialog(self, shortcut_key=None, action_name=None):
    """開啟快捷鍵對話框

    參數:
        shortcut_key: 快捷鍵（可選，用於編輯）
        action_name: 動作名稱（可選，用於編輯）
    """
    # 創建對話框
    dialog = tk.Toplevel(self.root)
    dialog.title("設定快捷鍵")
    dialog.geometry("400x200")
    dialog.transient(self.root)  # 設置為主窗口的子窗口
    dialog.resizable(False, False)
    
    # 快捷鍵框架
    key_frame = tk.Frame(dialog)
    key_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # 快捷鍵標籤
    tk.Label(key_frame, text="快捷鍵:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    
    # 快捷鍵輸入框
    key_entry = tk.Entry(key_frame, width=20)
    key_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    # 如果有快捷鍵，填入
    if shortcut_key:
        key_entry.insert(0, shortcut_key)
    
    # 快捷鍵說明
    tk.Label(key_frame, text="按下快捷鍵組合或直接輸入，例如: <Control-Shift-A>").grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
    
    # 動作框架
    action_frame = tk.Frame(dialog)
    action_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # 動作標籤
    tk.Label(action_frame, text="動作:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    
    # 動作下拉框
    action_var = tk.StringVar()
    action_combo = ttk.Combobox(action_frame, textvariable=action_var, width=30)
    action_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    # 設置可用的動作
    actions = [
        "open_file", "save_file", "save_file_as",
        "undo", "redo", "select_all", "copy", "cut", "paste",
        "correct_text", "find_text", "replace_text",
        "toggle_bold", "toggle_italic", "toggle_underline",
        "show_help", "refresh_view", "cancel_operation"
    ]
    action_combo["values"] = actions
    
    # 如果有動作，選擇它
    if action_name:
        action_var.set(action_name)
    
    # 按鈕框架
    button_frame = tk.Frame(dialog)
    button_frame.pack(fill=tk.X, pady=10)
    
    # 確定按鈕
    def save_shortcut():
        # 獲取快捷鍵和動作
        key = key_entry.get().strip()
        action = action_var.get()
        
        # 檢查是否為空
        if not key or not action:
            messagebox.showinfo("提示", "快捷鍵和動作不能為空")
            return
        
        # 檢查格式
        if not key.startswith("<") or not key.endswith(">"):
            messagebox.showinfo("提示", "快捷鍵格式不正確，應為 <修飾鍵-按鍵> 格式")
            return
        
        # 檢查是否與預設快捷鍵衝突
        if key in self.default_shortcuts and not shortcut_key:
            messagebox.showinfo("提示", f"快捷鍵 '{key}' 已被預設快捷鍵使用")
            return
        
        # 檢查是否與其他自定義快捷鍵衝突
        for item_id in self.custom_tree.get_children():
            item = self.custom_tree.item(item_id)
            values = item["values"]
            if values[0] == key and key != shortcut_key:
                messagebox.showinfo("提示", f"快捷鍵 '{key}' 已被使用")
                return
        
        # 添加到樹狀視圖
        self.custom_tree.insert("", tk.END, values=(key, action))
        
        # 關閉對話框
        dialog.destroy()
    
    ok_button = tk.Button(button_frame, text="確定", command=save_shortcut, width=10)
    ok_button.pack(side=tk.RIGHT, padx=10)
    
    # 取消按鈕
    cancel_button = tk.Button(button_frame, text="取消", command=dialog.destroy, width=10)
    cancel_button.pack(side=tk.RIGHT, padx=5)
    
    # 捕獲按鍵
    def capture_key(event):
        # 忽略特殊鍵
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return
        
        # 構建快捷鍵字符串
        key_str = "<"
        modifiers = []
        
        if event.state & 0x4:  # Control
            modifiers.append("Control")
        if event.state & 0x1:  # Shift
            modifiers.append("Shift")
        if event.state & 0x8:  # Alt
            modifiers.append("Alt")
        
        # 添加修飾鍵
        if modifiers:
            key_str += "-".join(modifiers) + "-"
        
        # 添加按鍵
        key_str += event.keysym
        key_str += ">"
        
        # 設置輸入框
        key_entry.delete(0, tk.END)
        key_entry.insert(0, key_str)
        
        return "break"  # 阻止事件繼續傳播
    
    # 綁定按鍵事件
    key_entry.bind("<Key>", capture_key)
    
    # 設置焦點
    key_entry.focus_set()

def save_custom_shortcuts(self):
    """儲存自定義快捷鍵"""
    # 清空自定義快捷鍵列表
    self.custom_shortcuts = []
    
    # 獲取所有項目
    for item_id in self.custom_tree.get_children():
        item = self.custom_tree.item(item_id)
        values = item["values"]
        
        # 添加到列表
        self.custom_shortcuts.append({
            "key": values[0],
            "action": values[1]
        })
    
    # 更新設定
    self.settings["custom_shortcuts"] = self.custom_shortcuts
    
    # 儲存設定
    from config_01_settings import save_settings
    save_settings(self)
    
    # 重新綁定快捷鍵
    self.bind_all_shortcuts()
    
    # 更新狀態欄
    self.status_bar.config(text="快捷鍵設定已更新")

def reset_shortcuts(self, parent_window):
    """重置所有快捷鍵

    參數:
        parent_window: 父視窗
    """
    # 確認重置
    if messagebox.askyesno("確認", "確定要重置所有快捷鍵嗎？這將刪除所有自定義快捷鍵。"):
        # 清空自定義快捷鍵
        self.custom_shortcuts = []
        
        # 更新設定
        self.settings["custom_shortcuts"] = []
        
        # 儲存設定
        from config_01_settings import save_settings
        save_settings(self)
        
        # 重新綁定快捷鍵
        self.bind_all_shortcuts()
        
        # 關閉視窗
        parent_window.destroy()
        
        # 顯示成功訊息
        messagebox.showinfo("成功", "所有快捷鍵已重置")

def select_all_text(self):
    """選取所有文字"""
    self.text_area.tag_add(tk.SEL, "1.0", tk.END)
    self.text_area.mark_set(tk.INSERT, "1.0")
    self.text_area.see(tk.INSERT)
    return 'break'

def copy_selected_text(self):
    """複製選取的文字"""
    if self.text_area.tag_ranges(tk.SEL):
        self.text_area.event_generate("<<Copy>>")
    return 'break'

def cut_selected_text(self):
    """剪下選取的文字"""
    if self.text_area.tag_ranges(tk.SEL):
        self.text_area.event_generate("<<Cut>>")
    return 'break'

def paste_text(self):
    """貼上文字"""
    self.text_area.event_generate("<<Paste>>")
    return 'break'

def cancel_operation(self):
    """取消目前的操作"""
    # 取消選取
    self.text_area.tag_remove(tk.SEL, "1.0", tk.END)
    
    # 關閉所有子視窗
    for widget in self.root.winfo_children():
        if isinstance(widget, tk.Toplevel):
            widget.destroy()
    
    # 更新狀態欄
    self.status_bar.config(text="操作已取消")
    
    return 'break'

def create_shortcut_button(self, shortcut_text):
    """創建單個快捷字按鈕並應用主題
    
    參數:
        shortcut_text: 快捷字文字內容
    """
    try:
        new_button = tk.Button(
            self.toolbar_bottom_frame,
            text=shortcut_text,
            command=lambda s=shortcut_text: self.insert_text_at_cursor(s)
        )
        new_button.pack(side=tk.LEFT, padx=2, pady=2)
        
        # 應用當前主題
        from config_01_settings import apply_theme_to_widget
        apply_theme_to_widget(self, new_button)
    except Exception as e:
        print(f"創建快捷按鈕 '{shortcut_text}' 時出錯: {e}")
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Shortcut Button Error", f"創建快捷按鈕 '{shortcut_text}' 時出錯", traceback.format_exc())

def load_custom_shortcut_buttons(self):
    """從設定檔載入並創建自訂快捷字按鈕"""
    # 清除現有的自訂按鈕（如果需要重新載入）
    # for widget in self.toolbar_bottom_frame.winfo_children():
    #     if widget.cget("text") in self.settings.get("custom_shortcuts", []):
    #         widget.destroy()
    # 通常在初始化時調用一次即可，除非有動態重載需求
    
    custom_shortcuts = self.settings.get("custom_shortcuts", [])
    print(f"載入自訂快捷字: {custom_shortcuts}")
    for sc in custom_shortcuts:
        create_shortcut_button(self, sc)

def add_shortcut(self):
    """新增快捷字: 跳出輸入視窗，將輸入文字變成新按鈕加到工具欄下層，並儲存到設定檔"""
    # 創建輸入對話框
    shortcut_window = tk.Toplevel(self.root)
    shortcut_window.title("新增快捷字")
    shortcut_window.geometry("300x150")
    shortcut_window.transient(self.root)  # 設置為主窗口的子窗口
    
    # 說明標籤
    tk.Label(shortcut_window, text="請輸入要新增的快捷字:", pady=10).pack()
    
    # 輸入框
    shortcut_entry = tk.Entry(shortcut_window, width=30)
    shortcut_entry.pack(pady=10)
    shortcut_entry.focus_set()  # 設置焦點
    
    # 按鈕框架
    button_frame = tk.Frame(shortcut_window)
    button_frame.pack(fill=tk.X, pady=10)
    
    # 確定按鈕
    def save_shortcut():
        shortcut_text = shortcut_entry.get().strip()
        if shortcut_text:
            # 檢查是否已存在
            if shortcut_text in self.settings.get("custom_shortcuts", []):
                tk.messagebox.showinfo("提示", f"快捷字 '{shortcut_text}' 已存在")
                return
            
            # 添加到設定
            if "custom_shortcuts" not in self.settings:
                self.settings["custom_shortcuts"] = []
            self.settings["custom_shortcuts"].append(shortcut_text)
            
            # 儲存設定
            from config_01_settings import save_settings
            save_settings(self)
            
            # 創建按鈕
            create_shortcut_button(self, shortcut_text)
            
            # 關閉視窗
            shortcut_window.destroy()
            
            # 更新狀態欄
            self.status_bar.config(text=f"已新增快捷字: {shortcut_text}")
    
    ok_button = tk.Button(button_frame, text="確定", command=save_shortcut, width=10)
    ok_button.pack(side=tk.RIGHT, padx=10)
    
    # 取消按鈕
    cancel_button = tk.Button(button_frame, text="取消", command=shortcut_window.destroy, width=10)
    cancel_button.pack(side=tk.RIGHT, padx=5)
    
    # 綁定回車鍵
    shortcut_entry.bind("<Return>", lambda event: save_shortcut())

def delete_shortcut(self):
    """刪除工具欄下層最右側的自訂快捷字按鈕及其設定"""
    # 獲取工具欄下層的所有按鈕
    buttons = [w for w in self.toolbar_bottom_frame.winfo_children() if isinstance(w, tk.Button)]
    
    # 如果沒有按鈕，直接返回
    if not buttons:
        self.status_bar.config(text="沒有可刪除的快捷字按鈕")
        return
    
    # 獲取最右側的按鈕
    last_button = buttons[-1]
    shortcut_text = last_button.cget("text")
    
    # 從設定中移除
    if shortcut_text in self.settings.get("custom_shortcuts", []):
        self.settings["custom_shortcuts"].remove(shortcut_text)
        
        # 儲存設定
        from config_01_settings import save_settings
        save_settings(self)
    
    # 銷毀按鈕
    last_button.destroy()
    
    # 更新狀態欄
    self.status_bar.config(text=f"已刪除快捷字: {shortcut_text}")
