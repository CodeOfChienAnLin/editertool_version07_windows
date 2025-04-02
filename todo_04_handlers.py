# -*- coding: utf-8 -*-
"""
代辦事項功能的事件處理函數
"""
import tkinter as tk
from tkinter import simpledialog, messagebox

# 導入對話框模組
from todo_02_dialogs import show_subtask_dialog
# 導入渲染模組 (處理後需要重新渲染)
# 避免循環導入，渲染將由主應用程式調用
# from todo_03_rendering import render_all_tasks

# --- 處理主任務 "+" 按鈕點擊事件 ---
def handle_add_main_task_click(tool_instance):
    """
    彈出視窗讓使用者輸入主任務區塊名稱，並添加到 task_groups。
    :param tool_instance: TextCorrectionTool 的實例。
    """
    task_name = simpledialog.askstring("新增任務區塊", "請輸入任務區塊名稱:", parent=tool_instance.root)
    if task_name:
        # 添加到實例的 task_groups 列表
        tool_instance.task_groups.append({
            "main_task_name": task_name,
            "sub_tasks": []
        })
        tool_instance.render_all_tasks() # 調用實例的渲染方法

# --- 處理子任務 "+" 按鈕點擊事件 ---
def handle_add_subtask_click(tool_instance, group_index):
    """
    彈出視窗讓使用者輸入子任務詳細資訊。
    :param tool_instance: TextCorrectionTool 的實例。
    :param group_index: 主任務的索引。
    """
    # 調用對話框函數，傳遞實例和索引
    show_subtask_dialog(tool_instance, group_index=group_index)
    # 渲染會在對話框儲存成功後自動調用

# --- 處理編輯子任務點擊事件 ---
def handle_edit_subtask_click(tool_instance, group_index, subtask_id):
    """
    根據 group_index 和 subtask_id 找到子任務並打開編輯視窗。
    :param tool_instance: TextCorrectionTool 的實例。
    :param group_index: 主任務的索引。
    :param subtask_id: 子任務的唯一 ID。
    """
    try:
        # 從實例的 task_groups 查找
        subtask_to_edit = next(sub for sub in tool_instance.task_groups[group_index]['sub_tasks'] if sub['id'] == subtask_id)
        # 調用對話框函數，傳遞實例、索引和現有數據
        show_subtask_dialog(tool_instance, group_index=group_index, subtask_data=subtask_to_edit)
        # 渲染會在對話框儲存成功後自動調用
    except (IndexError, StopIteration):
        messagebox.showerror("錯誤", "找不到要編輯的子任務。", parent=tool_instance.root)
    except Exception as e:
        messagebox.showerror("錯誤", f"編輯子任務時發生錯誤: {e}", parent=tool_instance.root)
        # log_error(...)

# --- 處理封存子任務點擊事件 ---
def handle_archive_subtask_click(tool_instance, group_index, subtask_id):
    """
    將子任務標記為封存並移動到封存列表。
    :param tool_instance: TextCorrectionTool 的實例。
    :param group_index: 主任務的索引。
    :param subtask_id: 子任務的唯一 ID。
    """
    try:
        group = tool_instance.task_groups[group_index]
        subtask_index = -1
        # 查找子任務在列表中的索引
        for i, sub in enumerate(group['sub_tasks']):
            if sub['id'] == subtask_id:
                subtask_index = i
                break

        if subtask_index != -1:
            # 從原列表移除
            subtask_to_archive = group['sub_tasks'].pop(subtask_index)
            subtask_to_archive['archived'] = True # 標記為封存
            # 添加到實例的 archived_tasks 列表
            tool_instance.archived_tasks.append(subtask_to_archive)
            tool_instance.render_all_tasks() # 調用實例的渲染方法
            messagebox.showinfo("成功", f"子任務 '{subtask_to_archive.get('name', '')}' 已封存。", parent=tool_instance.root) # 提供反饋
        else:
             messagebox.showerror("錯誤", "找不到要封存的子任務。", parent=tool_instance.root)

    except IndexError:
        messagebox.showerror("錯誤", "找不到指定的任務組。", parent=tool_instance.root)
    except Exception as e:
        messagebox.showerror("錯誤", f"封存子任務時發生錯誤: {e}", parent=tool_instance.root)
        # log_error(...)
