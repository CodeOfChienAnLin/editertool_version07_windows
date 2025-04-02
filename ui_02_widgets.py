"""
主要UI元件和TextCorrectionTool類別
"""
# -*- coding: utf-8 -*-
"""
主要UI元件和TextCorrectionTool類別
"""
import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font # 導入 font
import threading
import datetime
import traceback
import logging
import platform
from pathlib import Path
import msoffcrypto

# 導入自定義模組
from utils_01_error_handler import setup_error_logging, log_error
from config_01_settings import load_settings, save_settings
from config_02_protected_words import load_protected_words, save_protected_words, manage_protected_words
from text_01_correction import correct_text_thread, find_differences
from text_02_formatting import adjust_indentation, adjust_text_formatting
from file_01_word_processor import load_and_display_word_content, parse_word_document_com, handle_password_protected_file
from file_02_image_handler import extract_images_from_docx, display_image, show_full_image, clear_images, download_images, choose_download_path
from utils_02_shortcuts import create_shortcut_button, load_custom_shortcut_buttons

# 導入代辦事項模組
from todo_01_data import load_tasks_from_json, save_tasks_to_json
from todo_02_dialogs import show_archived_tasks_window, show_subtask_dialog, HAS_TKCALENDAR # 導入 HAS_TKCALENDAR
from todo_03_rendering import render_all_tasks, update_todo_scroll_region
from todo_04_handlers import handle_add_main_task_click, handle_add_subtask_click, handle_edit_subtask_click, handle_archive_subtask_click

# 嘗試導入 OpenCC
try:
    import opencc
except ImportError:
    opencc = None
    print("警告：未找到 opencc 模組。中文轉換功能將不可用。")

# 檢查 tkcalendar 是否可用 (從 todo_02_dialogs 導入)
if not HAS_TKCALENDAR:
    print("警告：未找到 tkcalendar 庫，日期選擇將使用普通輸入框。")
    print("請使用 'pip install tkcalendar' 安裝。")

class TextCorrectionTool:
    """文字校正工具主類別"""
    def __init__(self, root):
        """初始化應用程式

        參數:
            root: tkinter的根視窗
        """
        self.root = root
        self.root.title("編審神器")
        self.root.geometry("1000x700")  # 設定視窗大小為1000x700
        self.root.resizable(False, False)  # 禁止調整視窗大小

        # 設定錯誤日誌
        self.setup_error_logging()

        # 載入詞彙保護表
        self.protected_words = load_protected_words()

        # 載入設定 (包含自訂快捷字)
        self.settings = load_settings()
        # 確保 custom_shortcuts 存在且是列表
        if "custom_shortcuts" not in self.settings or not isinstance(self.settings["custom_shortcuts"], list):
             self.settings["custom_shortcuts"] = []

        # 初始化OpenCC轉換器
        try:
            # 使用簡體到繁體的轉換
            self.converter = opencc.OpenCC('s2t') if opencc else None  # 將簡體字轉為繁體字
        except Exception as e:
            messagebox.showerror("錯誤", f"無法初始化OpenCC轉換器: {str(e)}")
            self.converter = None

        # --- 代辦事項相關初始化 (確保所有相關屬性在 create_widgets 前存在) ---
        self.task_groups = []
        self.archived_tasks = []
        self.todo_canvas = None # 初始化為 None
        self.todo_main_task_font = None
        self.todo_sub_task_font = None
        self.todo_sub_task_time_font = None
        # --------------------------

        self.create_widgets()  # 創建UI元件
        self.setup_drag_drop()  # 設置拖放功能

        # 圖片相關變數
        self.images = []  # 存儲原始圖片
        self.image_refs = []  # 存儲 Tkinter PhotoImage 引用
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads")  # 預設下載路徑

        # 應用深色模式設定
        self.apply_theme()

    def setup_error_logging(self):
        """設定錯誤日誌記錄"""
        setup_error_logging()

    def create_widgets(self):
        """創建所有UI元件"""
        # --- 初始化代辦事項字體 (在 create_widgets 中進行) ---
        try: self.todo_main_task_font = font.Font(family="新細明體", size=14, weight="bold")
        except Exception: self.todo_main_task_font = font.Font(size=14, weight="bold") # Fallback
        try: self.todo_sub_task_font = font.Font(family="標楷體", size=12)
        except Exception: self.todo_sub_task_font = font.Font(size=12) # Fallback
        try: self.todo_sub_task_time_font = font.Font(family="標楷體", size=10)
        except Exception: self.todo_sub_task_time_font = font.Font(size=10) # Fallback
        # -------------------------------------------------------

        # 選單列
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 檔案選單
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="開啟", command=self.open_file)
        file_menu.add_command(label="儲存", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label="離開", command=self.root.quit)

        # 編輯選單 (移除文字修正, 加入還原)
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編輯", menu=edit_menu)
        edit_menu.add_command(label="還原上一步", command=self.undo_last_action) # 加入還原
        edit_menu.add_separator()
        edit_menu.add_command(label="管理保護詞彙", command=self.manage_protected_words)
        edit_menu.add_command(label="清除紅色標記", command=self.clear_correction_highlights)

        # 設定選單
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=settings_menu)
        settings_menu.add_command(label="文字格式", command=self.open_text_settings)
        settings_menu.add_command(label="換色模式", command=self.toggle_dark_mode)

        # 檢視選單
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檢視", menu=view_menu)
        view_menu.add_command(label="錯誤日誌", command=self.view_error_logs)

        # 主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 創建標籤頁控件
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 創建文字修正標籤頁
        self.text_correction_tab = tk.Frame(self.notebook)
        self.notebook.add(self.text_correction_tab, text="文字修正")

        # 創建代辦事項標籤頁
        self.notes_tab = tk.Frame(self.notebook) # 使用 self.notes_tab
        self.notebook.add(self.notes_tab, text="代辦事項")

        # --- 代辦事項標籤頁 UI (依照新佈局調整) ---

        # 左側工具欄框架
        todo_sidebar_frame = tk.Frame(self.notes_tab, width=100, relief="solid", borderwidth=1)
        todo_sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5, 0), pady=5)
        todo_sidebar_frame.pack_propagate(False) # 防止框架縮小

        # 任務區按鈕 (預設顯示，無需命令)
        task_area_button = ttk.Button(todo_sidebar_frame, text="任務區", command=None) # 樣式可能需要調整
        task_area_button.pack(fill="x", pady=2, padx=5)

        # 封存區按鈕
        archive_button = ttk.Button(todo_sidebar_frame, text="封存區", command=self.view_archived_tasks)
        archive_button.pack(fill="x", pady=2, padx=5)

        # 下載按鈕 (放置在底部)
        download_button = ttk.Button(todo_sidebar_frame, text="下載", command=self.save_tasks) # 連接到儲存功能
        download_button.pack(side="bottom", fill="x", pady=5, padx=5)

        # 右側主內容框架 (Canvas)
        todo_content_frame = tk.Frame(self.notes_tab)
        todo_content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 代辦事項 Canvas (創建並賦值給 self.todo_canvas)
        self.todo_canvas = tk.Canvas(todo_content_frame, bg='white', highlightthickness=0)

        # 代辦事項捲軸 (父容器: todo_content_frame)
        todo_v_scrollbar = ttk.Scrollbar(todo_content_frame, orient="vertical", command=self.todo_canvas.yview)
        todo_h_scrollbar = ttk.Scrollbar(todo_content_frame, orient="horizontal", command=self.todo_canvas.xview)
        self.todo_canvas.configure(yscrollcommand=todo_v_scrollbar.set, xscrollcommand=todo_h_scrollbar.set)

        # Grid 佈局 Canvas 和捲軸 (在 todo_content_frame 內)
        self.todo_canvas.grid(row=0, column=0, sticky="nsew")
        todo_v_scrollbar.grid(row=0, column=1, sticky="ns")
        todo_h_scrollbar.grid(row=1, column=0, sticky="ew")
        todo_content_frame.grid_rowconfigure(0, weight=1) # 讓 Canvas 行可以擴展
        todo_content_frame.grid_columnconfigure(0, weight=1) # 讓 Canvas 列可以擴展
        # --------------------------

        # --- 文字修正標籤頁 UI (再次確認元件放置順序) ---

        # 1. 工具欄框架 (父容器: self.text_correction_tab)
        self.toolbar_main_frame = tk.Frame(self.text_correction_tab, relief=tk.RAISED, bd=1)
        #   工具欄上層 (父容器: self.toolbar_main_frame)
        self.toolbar_top_frame = tk.Frame(self.toolbar_main_frame)
        self.toolbar_top_frame.pack(side=tk.TOP, fill=tk.X)

        # 工具欄按鈕 (上層)
        self.undo_button = tk.Button(self.toolbar_top_frame, text="還原上一步", command=self.undo_last_action)
        self.undo_button.pack(side=tk.LEFT, padx=2, pady=2)

        self.correct_button = tk.Button(self.toolbar_top_frame, text="文字修正", command=self.correct_text)
        self.correct_button.pack(side=tk.LEFT, padx=2, pady=2)

        self.add_shortcut_button = tk.Button(self.toolbar_top_frame, text="新增快捷字", command=self.add_shortcut)
        self.add_shortcut_button.pack(side=tk.LEFT, padx=2, pady=2)

        # 新增：刪除快捷字按鈕
        self.delete_shortcut_button = tk.Button(self.toolbar_top_frame, text="刪除快捷字", command=self.delete_shortcut)
        self.delete_shortcut_button.pack(side=tk.LEFT, padx=2, pady=2)

        # 工具欄下層
        self.toolbar_bottom_frame = tk.Frame(self.toolbar_main_frame)
        self.toolbar_bottom_frame.pack(side=tk.TOP, fill=tk.X)

        # 工具欄按鈕 (下層 - 預設快捷字/符號)
        default_shortcuts = ["，", "。", "「」", "『』", "民國(下同)", "新臺幣(下同)"]
        for sc in default_shortcuts:
            # Handle quotes needing cursor placement inside
            if sc == "「」" or sc == "『』":
                btn = tk.Button(self.toolbar_bottom_frame, text=sc,
                                command=lambda s=sc: self.insert_text_at_cursor(s, move_cursor=True))
            else:
                btn = tk.Button(self.toolbar_bottom_frame, text=sc,
                                command=lambda s=sc: self.insert_text_at_cursor(s))
            btn.pack(side=tk.LEFT, padx=2, pady=2)

        #   新增：載入並顯示自訂快捷字按鈕
        self.load_custom_shortcut_buttons()

        # 3. 圖片顯示區域框架 (父容器: self.text_correction_tab) - 先定義
        self.image_frame = tk.Frame(self.text_correction_tab, height=120) # 固定高度
        self.image_frame.pack_propagate(False) # 防止縮小

        #   圖片顯示區域的滾動畫布 (父容器: self.image_frame)
        self.image_canvas = tk.Canvas(self.image_frame)
        self.image_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 圖片區域的垂直滾動條
        image_scrollbar = tk.Scrollbar(self.image_frame, orient=tk.VERTICAL, command=self.image_canvas.yview)
        image_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.image_canvas.configure(yscrollcommand=image_scrollbar.set)

        # 創建一個框架來放置圖片
        self.image_container = tk.Frame(self.image_canvas)
        self.image_canvas.create_window((0, 0), window=self.image_container, anchor="nw")

        # 綁定圖片容器的配置事件
        self.image_container.bind("<Configure>", self.on_image_container_configure)

        # 按鈕框架 (圖片下載)
        img_button_frame = tk.Frame(self.image_frame)
        img_button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

        # 下載圖片按鈕
        self.download_button = tk.Button(img_button_frame, text="下載圖片", command=self.download_images)
        self.download_button.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)

        #   選擇路徑按鈕
        self.path_button = tk.Button(img_button_frame, text="選擇路徑", command=self.choose_download_path)
        self.path_button.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)

        # 2. 文字處理區域框架 (父容器: self.text_correction_tab) - 先定義
        text_frame = tk.Frame(self.text_correction_tab)

        #   添加垂直滾動條 (父容器: text_frame)
        y_scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 文字處理區域 - 啟用 undo
        # 修改：使用 spacing1 控制段落內行距
        self.text_area = tk.Text(text_frame,
                               font=(self.settings["font_family"], self.settings["font_size"]),
                               spacing1=self.settings["line_spacing_within"], # 使用 spacing1
                               wrap=tk.WORD,
                               undo=True, # 啟用 Undo/Redo
                               yscrollcommand=y_scrollbar.set)
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 創建紅色底線標籤
        self.text_area.tag_configure("corrected", underline=True, underlinefg="red")

        # 設置縮進
        self.text_area.config(tabs=("1c", "2c", "3c", "4c"), tabstyle="wordprocessor")

        # 綁定事件
        self.text_area.bind("<<Modified>>", self.adjust_indentation)

        #   設置滾動條命令
        y_scrollbar.config(command=self.text_area.yview)

        # --- 按照正確順序 pack 文字修正標籤頁的元件 ---
        # 1. 工具欄置頂
        self.toolbar_main_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))
        # 3. 圖片區置底 (必須在文字區之前 pack side=BOTTOM)
        self.image_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        # 2. 文字區填滿中間剩餘空間
        text_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 狀態欄
        self.status_bar = tk.Label(self.root, text="就緒", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # --- 初始渲染代辦事項 (延遲調用，傳遞參數) ---
        # 使用 after 延遲初始渲染，確保所有元件已創建且 __init__ 已完成
        self.root.after(10, self.render_all_tasks) # 直接傳遞方法引用
        # ----------------------

    # --- 代辦事項相關方法 ---
    def render_all_tasks(self):
        """渲染所有代辦事項到 Canvas"""
        # 檢查必要的屬性是否存在
        if not all([self.todo_canvas, self.task_groups is not None,
                    self.todo_main_task_font, self.todo_sub_task_font, self.todo_sub_task_time_font]):
            print("錯誤：缺少渲染所需的屬性 (canvas, task_groups, 或字體)")
            return
        # 調用渲染模組函數，傳遞必要的參數
        render_all_tasks(
            canvas=self.todo_canvas,
            task_groups=self.task_groups,
            main_task_font=self.todo_main_task_font,
            sub_task_font=self.todo_sub_task_font,
            sub_task_time_font=self.todo_sub_task_time_font,
            tool_instance=self # 仍然需要傳遞 self 以便處理事件回呼
        )

    def add_main_task(self):
        """處理新增主任務按鈕點擊"""
        handle_add_main_task_click(self) # 調用處理函數

    def add_sub_task(self, group_index):
        """處理新增子任務按鈕點擊"""
        handle_add_subtask_click(self, group_index) # 調用處理函數

    def edit_sub_task(self, group_index, subtask_id):
        """處理編輯子任務點擊"""
        handle_edit_subtask_click(self, group_index, subtask_id) # 調用處理函數

    def archive_sub_task(self, group_index, subtask_id):
        """處理封存子任務點擊"""
        handle_archive_subtask_click(self, group_index, subtask_id) # 調用處理函數

    def load_tasks(self):
        """載入代辦事項"""
        if load_tasks_from_json(self): # 調用資料模組函數
            self.render_all_tasks() # 載入成功後重新渲染 (會自動傳遞參數)

    def save_tasks(self):
        """儲存代辦事項"""
        save_tasks_to_json(self) # 調用資料模組函數

    def view_archived_tasks(self):
        """顯示封存區視窗"""
        show_archived_tasks_window(self) # 調用對話框模組函數
    # --------------------------

    # --- 文字修正相關方法 ---
    def undo_last_action(self):
        """還原上一步文字編輯操作"""
        try:
            self.text_area.edit_undo()
            self.status_bar.config(text="已還原上一步操作")
        except tk.TclError:
            self.status_bar.config(text="沒有可還原的操作")

    def add_shortcut(self):
        """新增快捷字: 跳出輸入視窗，將輸入文字變成新按鈕加到工具欄下層，並儲存到設定檔"""
        from utils_02_shortcuts import add_shortcut
        add_shortcut(self)

    def delete_shortcut(self):
        """刪除工具欄下層最右側的自訂快捷字按鈕及其設定"""
        from utils_02_shortcuts import delete_shortcut
        delete_shortcut(self)

    def load_custom_shortcut_buttons(self):
        """從設定檔載入並創建自訂快捷字按鈕"""
        load_custom_shortcut_buttons(self)

    def apply_theme_to_widget(self, widget):
        """Apply the current theme to a specific widget."""
        from config_01_settings import apply_theme_to_widget
        apply_theme_to_widget(self, widget)

    def apply_theme(self):
        """應用主題設定"""
        from config_01_settings import apply_theme
        apply_theme(self)

    def insert_text_at_cursor(self, text_to_insert, move_cursor=False):
        """在目前游標位置插入文字"""
        current_pos = self.text_area.index(tk.INSERT)
        self.text_area.insert(current_pos, text_to_insert)
        
        # 如果需要移動游標（例如插入引號時）
        if move_cursor and len(text_to_insert) >= 2:
            # 假設游標應該在倒數第二個字符後
            new_pos = f"{current_pos}+{len(text_to_insert) - 1}c"
            self.text_area.mark_set(tk.INSERT, new_pos)

    def on_image_container_configure(self, event):
        """當圖片容器大小變化時，更新畫布的滾動區域"""
        self.image_canvas.configure(scrollregion=self.image_canvas.bbox("all"))

    def download_images(self):
        """下載圖片：調用 file_02_image_handler 模組中的 download_images 函數"""
        from file_02_image_handler import download_images
        download_images(self)

    def choose_download_path(self):
        """選擇下載路徑：調用 file_02_image_handler 模組中的 choose_download_path 函數"""
        from file_02_image_handler import choose_download_path
        choose_download_path(self)

    def adjust_indentation(self, event=None):
        """調整文字縮進：調用 text_02_formatting 模組中的 adjust_indentation 函數"""
        from text_02_formatting import adjust_indentation
        adjust_indentation(self, event)

    def setup_drag_drop(self):
        """設置拖放功能"""
        # 檢查是否有 TkinterDnD2 支援
        if hasattr(self.root, 'drop_target_register'):
            self.root.drop_target_register('*')
            self.root.dnd_bind('<<Drop>>', self.handle_drop)
            self.status_bar.config(text="拖放功能已啟用")
        else:
            # 如果沒有拖放支援，綁定剪貼板事件作為替代
            self.root.bind('<Control-v>', self.check_clipboard)
            self.status_bar.config(text="拖放功能未啟用，請使用 Ctrl+V 貼上檔案路徑")

    def check_clipboard(self, event=None):
        """檢查剪貼簿是否有檔案路徑"""
        try:
            clipboard = self.root.clipboard_get()
            if os.path.isfile(clipboard):
                self.handle_drop(clipboard)
            return "break"  # 防止默認的貼上行為
        except:
            pass  # 剪貼簿內容不是文字或不是檔案路徑

    def handle_drop(self, event):
        """處理檔案拖放事件"""
        # 獲取檔案路徑
        if isinstance(event, str):
            # 從剪貼簿獲取的路徑
            file_path = event
        else:
            # 從拖放事件獲取的路徑
            file_path = event.data
            
            # 移除可能的 {} 或 引號
            if file_path.startswith("{") and file_path.endswith("}"):
                file_path = file_path[1:-1]
            if file_path.startswith('"') and file_path.endswith('"'):
                file_path = file_path[1:-1]

        # 檢查檔案是否存在
        if not os.path.isfile(file_path):
            messagebox.showerror("錯誤", f"找不到檔案: {file_path}")
            return

        # 根據檔案類型處理
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext in ['.docx', '.doc']:
            # 更新狀態欄，提示使用者正在開啟檔案
            self.status_bar.config(text=f"正在開啟 Word 檔案: {os.path.basename(file_path)}...")
            
            # 嘗試檢測檔案是否加密
            try:
                with open(file_path, 'rb') as file:
                    try:
                        ms_file = msoffcrypto.OfficeFile(file)
                        if ms_file.is_encrypted():
                            # 如果檔案已加密，調用處理加密檔案的函數
                            handle_password_protected_file(self, file_path)
                            return
                    except Exception as e:
                        # 如果檢測加密狀態失敗，嘗試正常載入
                        pass
            except Exception as e:
                # 如果打開檔案失敗，顯示錯誤訊息
                messagebox.showerror("錯誤", f"無法開啟檔案: {str(e)}")
                return
                
            # 如果檔案未加密或無法確定加密狀態，嘗試正常載入
            self.load_and_display_word_content(file_path)
        else:
            messagebox.showinfo("提示", f"不支援的檔案類型: {file_ext}\n目前僅支援 .docx 和 .doc 檔案")

    def load_and_display_word_content(self, file_path, password=None):
        """載入並顯示 Word 文件內容"""
        from file_01_word_processor import load_and_display_word_content
        load_and_display_word_content(self, file_path, password)

    def open_file(self):
        """開啟檔案對話框"""
        file_path = filedialog.askopenfilename(
            title="選擇檔案",
            filetypes=[("Word 文件", "*.docx *.doc"), ("所有檔案", "*.*")]
        )
        if file_path:
            self.handle_drop(file_path)

    def save_file(self):
        """儲存檔案對話框"""
        file_path = filedialog.asksaveasfilename(
            title="儲存檔案",
            defaultextension=".txt",
            filetypes=[("文字檔", "*.txt"), ("所有檔案", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(self.text_area.get("1.0", tk.END))
                self.status_bar.config(text=f"檔案已儲存: {file_path}")
            except Exception as e:
                error_msg = f"儲存檔案時發生錯誤: {str(e)}"
                messagebox.showerror("錯誤", error_msg)
                log_error(self, "File Save Error", error_msg, traceback.format_exc())

    def correct_text(self):
        """校正文字內容"""
        from text_01_correction import correct_text
        correct_text(self)

    def clear_correction_highlights(self):
        """清除所有校正標記"""
        self.text_area.tag_remove("corrected", "1.0", tk.END)
        self.status_bar.config(text="已清除所有校正標記")

    def manage_protected_words(self):
        """管理保護詞彙的視窗"""
        from config_02_protected_words import manage_protected_words
        manage_protected_words(self)

    def open_text_settings(self):
        """開啟文字格式設定視窗"""
        from config_01_settings import open_text_settings
        open_text_settings(self)

    def toggle_dark_mode(self):
        """切換深色模式"""
        from config_01_settings import toggle_dark_mode
        toggle_dark_mode(self)

    def view_error_logs(self):
        """檢視錯誤日誌"""
        from utils_01_error_handler import view_error_logs
        view_error_logs(self)
