"""
設定管理相關功能模組
"""
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import traceback

def load_settings():
    """載入設定

    回傳:
        設定字典
    """
    # 預設設定
    default_settings = {
        "font_family": "微軟正黑體",
        "font_size": 12,
        "line_spacing_within": 4,  # 段落內行距
        "dark_mode": False,
        "custom_shortcuts": []
    }
    
    # 設定檔路徑
    settings_path = "settings.json"
    
    # 如果設定檔存在，載入它
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as file:
                settings = json.load(file)
                
                # 確保所有預設設定都存在
                for key, value in default_settings.items():
                    if key not in settings:
                        settings[key] = value
                
                return settings
        except Exception as e:
            print(f"載入設定時發生錯誤: {str(e)}")
            return default_settings
    
    # 如果設定檔不存在，使用預設設定
    return default_settings

def save_settings(self):
    """儲存設定"""
    try:
        # 設定檔路徑
        settings_path = "settings.json"
        
        # 儲存設定
        with open(settings_path, 'w', encoding='utf-8') as file:
            json.dump(self.settings, file, ensure_ascii=False, indent=4)
        
        self.status_bar.config(text="設定已儲存")
    except Exception as e:
        error_msg = f"儲存設定時發生錯誤: {str(e)}"
        messagebox.showerror("錯誤", error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Settings Save Error", error_msg, traceback.format_exc())

def open_text_settings(self):
    """開啟文字格式設定視窗"""
    # 創建設定視窗
    settings_window = tk.Toplevel(self.root)
    settings_window.title("文字格式設定")
    settings_window.geometry("450x500")  # 增加視窗大小，確保有足夠空間
    settings_window.resizable(False, False)
    settings_window.transient(self.root)  # 設置為主窗口的子窗口
    
    # 主框架 - 使用 grid 佈局管理器
    main_frame = tk.Frame(settings_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    main_frame.grid_columnconfigure(0, weight=1)
    
    # 字型設定框架
    font_frame = tk.LabelFrame(main_frame, text="字型設定")
    font_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    
    # 字型家族
    tk.Label(font_frame, text="字型:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    font_family_var = tk.StringVar(value=self.settings["font_family"])
    font_family_combo = ttk.Combobox(font_frame, textvariable=font_family_var)
    font_family_combo['values'] = ('微軟正黑體', '新細明體', '標楷體', 'Arial', 'Times New Roman')
    font_family_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    # 字型大小
    tk.Label(font_frame, text="大小:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    font_size_var = tk.IntVar(value=self.settings["font_size"])
    font_size_combo = ttk.Combobox(font_frame, textvariable=font_size_var, width=5)
    font_size_combo['values'] = (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)
    font_size_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
    
    # 行距設定框架
    spacing_frame = tk.LabelFrame(main_frame, text="行距設定")
    spacing_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    
    # 段落內行距
    tk.Label(spacing_frame, text="段落內行距:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    line_spacing_var = tk.IntVar(value=self.settings["line_spacing_within"])
    line_spacing_combo = ttk.Combobox(spacing_frame, textvariable=line_spacing_var, width=5)
    line_spacing_combo['values'] = (0, 1, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16)
    line_spacing_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    # 預覽框架 - 獨立顯示區域
    preview_frame = tk.LabelFrame(main_frame, text="預覽")
    preview_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
    preview_frame.grid_columnconfigure(0, weight=1)
    preview_frame.grid_rowconfigure(0, weight=1)
    
    # 創建固定高度的預覽容器，避免影響其他元素
    preview_container = tk.Frame(preview_frame, height=150)
    preview_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    preview_container.grid_propagate(False)  # 防止內容改變框架大小
    preview_container.grid_columnconfigure(0, weight=1)
    preview_container.grid_rowconfigure(0, weight=1)
    
    # 創建文字區域和滾動條的框架
    text_scroll_frame = tk.Frame(preview_container)
    text_scroll_frame.grid(row=0, column=0, sticky="nsew")
    text_scroll_frame.grid_columnconfigure(0, weight=1)
    text_scroll_frame.grid_rowconfigure(0, weight=1)
    
    # 添加垂直滾動條
    preview_scrollbar = tk.Scrollbar(text_scroll_frame)
    preview_scrollbar.grid(row=0, column=1, sticky="ns")
    
    # 預覽文字區域
    preview_text = tk.Text(text_scroll_frame, wrap=tk.WORD, 
                          yscrollcommand=preview_scrollbar.set)
    preview_text.grid(row=0, column=0, sticky="nsew")
    preview_scrollbar.config(command=preview_text.yview)
    
    # 設置預覽文字
    preview_text.insert("1.0", "這是預覽文字，用來展示字型和行距設定的效果。\n這是第二行文字。\n\n調整字型和行距後，可以在此區域查看效果。\n如果文字超出顯示範圍，可以使用滾動條查看。")
    
    # 更新預覽的函數
    def update_preview(*args):
        try:
            # 獲取當前設定值
            family = font_family_var.get()
            size = font_size_var.get()
            spacing = line_spacing_var.get()
            
            # 更新預覽文字區域
            preview_text.configure(
                font=(family, size),
                spacing1=spacing
            )
        except Exception as e:
            print(f"更新預覽時發生錯誤: {str(e)}")
    
    # 綁定更新預覽的事件
    font_family_combo.bind("<<ComboboxSelected>>", update_preview)
    font_size_combo.bind("<<ComboboxSelected>>", update_preview)
    line_spacing_combo.bind("<<ComboboxSelected>>", update_preview)
    
    # 初始更新預覽
    update_preview()
    
    # 按鈕框架 - 固定在底部
    button_frame = tk.Frame(main_frame)
    button_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=10)
    button_frame.grid_columnconfigure(1, weight=1)  # 讓確定按鈕靠右
    
    # 確定按鈕
    def apply_settings():
        try:
            # 更新設定
            self.settings["font_family"] = font_family_var.get()
            self.settings["font_size"] = font_size_var.get()
            self.settings["line_spacing_within"] = line_spacing_var.get()
            
            # 儲存設定
            with open("settings.json", 'w', encoding='utf-8') as file:
                json.dump(self.settings, file, ensure_ascii=False, indent=4)
            
            # 更新文字區域
            self.text_area.configure(
                font=(self.settings["font_family"], self.settings["font_size"]),
                spacing1=self.settings["line_spacing_within"]
            )
            
            # 關閉設定視窗
            settings_window.destroy()
            
            # 更新狀態欄
            self.status_bar.config(text="文字格式設定已更新")
        except Exception as e:
            error_msg = f"儲存設定時發生錯誤: {str(e)}"
            messagebox.showerror("錯誤", error_msg)
            
            # 記錄錯誤
            from utils_01_error_handler import log_error
            log_error(self, "Settings Save Error", error_msg, traceback.format_exc())
    
    # 取消按鈕 - 靠左
    cancel_button = tk.Button(button_frame, text="取消", command=settings_window.destroy, width=10)
    cancel_button.grid(row=0, column=0, padx=5, pady=5)
    
    # 確定按鈕 - 靠右
    ok_button = tk.Button(button_frame, text="確定", command=apply_settings, width=10)
    ok_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")

def toggle_dark_mode(self):
    """切換深色模式"""
    # 切換深色模式設定
    self.settings["dark_mode"] = not self.settings["dark_mode"]
    
    # 應用主題
    apply_theme(self)
    
    # 儲存設定
    save_settings(self)
    
    # 更新狀態欄
    mode_text = "深色" if self.settings["dark_mode"] else "淺色"
    self.status_bar.config(text=f"已切換至{mode_text}模式")

def apply_theme_to_widget(self, widget):
    """Apply the current theme to a specific widget."""
    if not hasattr(self, 'settings') or not widget: # Safety check
        return

    # Determine colors based on theme
    if self.settings["dark_mode"]:
        bg_color = "#333333" # 背景色 (深灰)
        fg_color = "white" # 前景色 (文字)
        button_bg = "#555555" # 按鈕背景 (稍淺灰)
        button_fg = "white" # 按鈕文字
        # Add more specific colors if needed
    else:
        # 使用系統預設顏色以獲得更好的外觀整合
        bg_color = "SystemButtonFace" # 主背景色
        fg_color = "black" # 前景色
        button_bg = "SystemButtonFace" # 按鈕背景
        button_fg = "black" # 按鈕文字
        # Add more specific colors if needed

    # Apply theme based on widget type (expand as needed)
    widget_type = widget.winfo_class() # 獲取元件類型 (例如 'Button', 'Label', 'Frame')

    try:
        if widget_type == 'Button':
            widget.configure(bg=button_bg, fg=button_fg)
        elif widget_type in ['Label', 'Frame', 'LabelFrame']:
            widget.configure(bg=bg_color, fg=fg_color)
        elif widget_type == 'Text':
            widget.configure(bg=bg_color, fg=fg_color, insertbackground=fg_color)
        elif widget_type == 'Canvas':
            widget.configure(bg=bg_color)
        # Add more widget types as needed
    except tk.TclError:
        # Some widgets might not support all configurations
        pass

def apply_theme(self):
    """應用主題設定"""
    # 獲取顏色設定
    if self.settings["dark_mode"]:
        bg_color = "#333333" # 背景色 (深灰)
        fg_color = "white" # 前景色 (文字)
        text_bg = "#222222" # 文字區域背景 (更深灰)
        text_fg = "white" # 文字區域前景
        button_bg = "#555555" # 按鈕背景 (稍淺灰)
        button_fg = "white" # 按鈕文字
        canvas_bg = "#444444" # 畫布背景
    else:
        bg_color = "SystemButtonFace" # 主背景色
        fg_color = "black" # 前景色
        text_bg = "white" # 文字區域背景
        text_fg = "black" # 文字區域前景
        button_bg = "SystemButtonFace" # 按鈕背景
        button_fg = "black" # 按鈕文字
        canvas_bg = "white" # 畫布背景
    
    # 應用主題到根視窗
    self.root.configure(bg=bg_color)
    
    # 應用主題到標籤頁
    self.text_correction_tab.configure(bg=bg_color)
    self.notes_tab.configure(bg=bg_color)
    
    # 應用主題到工具欄
    self.toolbar_main_frame.configure(bg=bg_color)
    self.toolbar_top_frame.configure(bg=bg_color)
    self.toolbar_bottom_frame.configure(bg=bg_color)
    
    # 應用主題到按鈕
    for widget in self.toolbar_top_frame.winfo_children():
        if isinstance(widget, tk.Button):
            widget.configure(bg=button_bg, fg=button_fg)
    
    for widget in self.toolbar_bottom_frame.winfo_children():
        if isinstance(widget, tk.Button):
            widget.configure(bg=button_bg, fg=button_fg)
    
    # 應用主題到文字區域
    self.text_area.configure(bg=text_bg, fg=text_fg, insertbackground=text_fg)
    
    # 應用主題到圖片區域
    self.image_frame.configure(bg=bg_color)
    self.image_canvas.configure(bg=canvas_bg)
    self.image_container.configure(bg=canvas_bg)
    
    # 應用主題到圖片按鈕
    for widget in self.image_frame.winfo_children():
        if isinstance(widget, tk.Frame):
            widget.configure(bg=bg_color)
            for child in widget.winfo_children():
                if isinstance(child, tk.Button):
                    child.configure(bg=button_bg, fg=button_fg)
    
    # 應用主題到狀態欄
    self.status_bar.configure(bg=bg_color, fg=fg_color)
