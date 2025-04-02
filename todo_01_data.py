# -*- coding: utf-8 -*-
"""
代辦事項功能的資料結構、常數和 JSON 處理
"""
import json
import uuid
from tkinter import filedialog, messagebox

# --- 全域變數 (將由主應用程式實例管理) ---
# task_groups = [] # 改為實例變數 self.task_groups
# archived_tasks = [] # 改為實例變數 self.archived_tasks

# --- 顏色定義 ---
SUBTASK_COLORS = {
    "紅色": "#FFCDD2", # 淺紅
    "黃色": "#FFF9C4", # 淺黃
    "藍色": "#BBDEFB", # 淺藍
    "綠色": "#C8E6C9", # 淺綠
    "橘色": "#FFECB3", # 淺橘
    "預設": "#F5F5F5"  # 預設灰色
}
COLOR_NAMES = list(SUBTASK_COLORS.keys()) # ["紅色", "黃色", ...]

# --- 尺寸和間距常數 ---
# 注意：這些值可能需要根據整合後的 UI 佈局進行調整
MAIN_TASK_WIDTH = 270        # 主任務區塊寬度
MAIN_TASK_PADDING_Y = 10     # 主任務名稱上下邊距
MAIN_TASK_SPACING_X = 15     # 主任務區塊之間的水平間距
MAIN_TASK_SPACING_Y = 15     # 主任務區塊之間的垂直間距 (如果換行)

SUBTASK_HEIGHT_BASE = 60    # 子任務基礎高度 (估計值，會動態調整)
SUBTASK_WIDTH_MARGIN = 20    # 子任務左右邊距總和
SUBTASK_INTERNAL_PADDING_Y = 8 # 子任務內部上下邊距
SUBTASK_SPACING_Y = 8        # 子任務之間的垂直間距
SUBTASK_CIRCLE_RADIUS = 8    # 封存按鈕圓圈半徑
SUBTASK_CIRCLE_MARGIN = 10   # 封存按鈕左邊距
NAME_TIME_GAP = 3            # 子任務名稱和時間之間的垂直間距

# --- 輔助函數：獲取顏色值 ---
def get_color_code(name):
    """根據顏色名稱獲取十六進位顏色碼"""
    return SUBTASK_COLORS.get(name, SUBTASK_COLORS["預設"])

# --- 儲存 JSON ---
def save_tasks_to_json(tool_instance):
    """
    將目前的任務資料儲存到 JSON 檔案。
    :param tool_instance: TextCorrectionTool 的實例，用於訪問 task_groups, archived_tasks 和 root。
    """
    file_path = filedialog.asksaveasfilename(
        defaultextension=".json", filetypes=[("JSON files", "*.json")], title="選擇儲存路徑",
        parent=tool_instance.root # 指定父視窗
    )
    if not file_path: return

    # 準備儲存的數據 (過濾掉臨時 UI 數據)
    data_to_save = {
        "task_groups": [
            {
                "main_task_name": group.get("main_task_name", ""),
                "sub_tasks": [
                    # 過濾掉 y_pos 和 height，只儲存未封存的
                    {k: v for k, v in sub.items() if k not in ['y_pos', 'height']}
                    for sub in group.get("sub_tasks", []) if not sub.get("archived")
                ]
            } for group in tool_instance.task_groups # 從實例獲取
        ],
        "archived_tasks": [
             # 過濾掉 y_pos 和 height
             {k: v for k, v in sub.items() if k not in ['y_pos', 'height']}
             for sub in tool_instance.archived_tasks # 從實例獲取
        ] # 儲存封存的任務
    }
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "代辦事項資料已成功儲存！", parent=tool_instance.root)
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存代辦事項失敗: {e}", parent=tool_instance.root)
        # 可以考慮在這裡也記錄錯誤日誌
        # from utils_01_error_handler import log_error
        # log_error(tool_instance, "Save Tasks Error", f"儲存代辦事項失敗: {e}", traceback.format_exc())

# --- 載入 JSON ---
def load_tasks_from_json(tool_instance):
    """
    從 JSON 檔案載入任務資料。
    :param tool_instance: TextCorrectionTool 的實例，用於更新 task_groups 和 archived_tasks。
    :return: 如果成功載入則返回 True，否則返回 False。
    """
    file_path = filedialog.askopenfilename(
        defaultextension=".json", filetypes=[("JSON files", "*.json")], title="選擇載入的代辦事項檔案",
        parent=tool_instance.root
    )
    if not file_path: return False

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # 基本驗證
        if not isinstance(data, dict) or "task_groups" not in data or "archived_tasks" not in data:
            raise ValueError("無效的 JSON 檔案格式")

        # 清空現有數據
        tool_instance.task_groups = data.get("task_groups", [])
        tool_instance.archived_tasks = data.get("archived_tasks", [])

        # (可選) 添加更詳細的驗證，例如檢查子任務結構

        messagebox.showinfo("成功", "代辦事項資料已成功載入！", parent=tool_instance.root)
        return True # 載入成功

    except FileNotFoundError:
        messagebox.showerror("錯誤", f"找不到檔案: {file_path}", parent=tool_instance.root)
        return False
    except json.JSONDecodeError:
        messagebox.showerror("錯誤", f"無法解析 JSON 檔案: {file_path}", parent=tool_instance.root)
        return False
    except ValueError as ve:
         messagebox.showerror("錯誤", f"檔案內容錯誤: {ve}", parent=tool_instance.root)
         return False
    except Exception as e:
        messagebox.showerror("錯誤", f"載入代辦事項失敗: {e}", parent=tool_instance.root)
        # from utils_01_error_handler import log_error
        # log_error(tool_instance, "Load Tasks Error", f"載入代辦事項失敗: {e}", traceback.format_exc())
        return False

# --- 生成唯一 ID ---
def generate_uuid():
    """生成一個唯一的十六進位字串 ID"""
    return uuid.uuid4().hex
