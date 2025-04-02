# -*- coding: utf-8 -*-
"""
代辦事項功能的對話框（新增/編輯子任務、顯示封存區）
"""
import tkinter as tk
from tkinter import simpledialog, messagebox, font, ttk, Text
from datetime import datetime

# 嘗試匯入 tkcalendar
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False
    # 注意：這裡不再 print 警告，主應用程式會處理

# 導入資料模組中的常數和函數
from todo_01_data import COLOR_NAMES, generate_uuid, load_tasks_from_json, save_tasks_to_json # 載入/儲存也可能觸發對話框
# 導入渲染模組 (因為儲存/取消封存後需要重新渲染)
# 避免循環導入，渲染函數將由主應用程式調用
# from todo_03_rendering import render_all_tasks

# --- 顯示/編輯子任務的彈出視窗 ---
def show_subtask_dialog(tool_instance, group_index, subtask_data=None):
    """
    創建或編輯子任務的彈出視窗。
    :param tool_instance: TextCorrectionTool 的實例。
    :param group_index: 子任務所屬的主任務索引。
    :param subtask_data: 如果是編輯，傳入現有的子任務字典；否則為 None。
    """
    is_editing = subtask_data is not None
    dialog_title = "編輯子任務" if is_editing else "新增子任務"

    dialog = tk.Toplevel(tool_instance.root)
    dialog.title(dialog_title)
    # 調整大小以適應可能的內容
    dialog.geometry("450x350") # 稍微加大
    dialog.resizable(False, False)
    dialog.grab_set() # 鎖定焦點
    dialog.transient(tool_instance.root) # 依附主視窗

    # --- 創建輸入元件 ---
    frame = ttk.Frame(dialog, padding="15") # 增加內邊距
    frame.pack(fill="both", expand=True)

    # 名稱
    ttk.Label(frame, text="子任務名稱:").grid(row=0, column=0, sticky="w", pady=3)
    name_var = tk.StringVar(value=subtask_data['name'] if is_editing else "")
    name_entry = ttk.Entry(frame, textvariable=name_var, width=45) # 調整寬度
    name_entry.grid(row=0, column=1, columnspan=4, sticky="ew", pady=3)
    name_entry.focus_set() # 設定焦點

    # 到期日期
    ttk.Label(frame, text="到期日期:").grid(row=1, column=0, sticky="w", pady=3)
    initial_date = subtask_data.get('due_date', '') if is_editing else ""
    date_var = tk.StringVar(value=initial_date)
    if HAS_TKCALENDAR:
        date_entry = DateEntry(frame, textvariable=date_var, date_pattern='yyyy-mm-dd', width=12,
                                background='darkblue', foreground='white', borderwidth=2,
                                firstweekday='sunday', showweeknumbers=False, state='normal')
        if not initial_date:
             date_entry.delete(0, tk.END) # 清空初始值
        date_entry.grid(row=1, column=1, columnspan=4, sticky="w", pady=3)
    else:
        date_entry = ttk.Entry(frame, textvariable=date_var, width=15)
        date_entry.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(frame, text="(YYYY-MM-DD)").grid(row=1, column=2, columnspan=3, sticky="w", padx=5)

    # 到期時間
    ttk.Label(frame, text="到期時間:").grid(row=2, column=0, sticky="w", pady=3)
    time_frame = ttk.Frame(frame)
    time_frame.grid(row=2, column=1, columnspan=4, sticky="w")

    hour_var = tk.StringVar()
    hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, wrap=True, width=3, format="%02.0f")
    hour_spinbox.pack(side=tk.LEFT, padx=(0, 2))
    ttk.Label(time_frame, text=":").pack(side=tk.LEFT, padx=0)
    minute_var = tk.StringVar()
    minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, wrap=True, width=3, format="%02.0f")
    minute_spinbox.pack(side=tk.LEFT, padx=(2, 0))

    initial_hour, initial_minute = "", ""
    if is_editing and subtask_data.get('due_time'):
        try:
            h, m = map(int, subtask_data['due_time'].split(':'))
            initial_hour, initial_minute = f"{h:02d}", f"{m:02d}"
        except (ValueError, TypeError): pass # 忽略格式錯誤
    hour_var.set(initial_hour)
    minute_var.set(initial_minute)
    if not initial_hour and not initial_minute:
        hour_spinbox.delete(0, tk.END)
        minute_spinbox.delete(0, tk.END)

    # 顏色選擇
    ttk.Label(frame, text="區塊顏色:").grid(row=3, column=0, sticky="w", pady=3)
    color_var = tk.StringVar(value=subtask_data.get('color_name', COLOR_NAMES[0]) if is_editing else COLOR_NAMES[0]) # 使用 .get 提供預設值
    color_combo = ttk.Combobox(frame, textvariable=color_var, values=COLOR_NAMES, state="readonly", width=10)
    color_combo.grid(row=3, column=1, columnspan=4, sticky="w", pady=3)

    # 詳細內容
    ttk.Label(frame, text="詳細內容:").grid(row=4, column=0, sticky="nw", pady=3)
    details_text = Text(frame, width=40, height=6, wrap="word") # 增加高度
    details_text.grid(row=4, column=1, columnspan=3, sticky="nsew", pady=3)
    if is_editing and subtask_data.get('details'):
        details_text.insert("1.0", subtask_data['details'])
    details_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=details_text.yview)
    details_scrollbar.grid(row=4, column=4, sticky="ns")
    details_text['yscrollcommand'] = details_scrollbar.set

    # --- 儲存按鈕回呼函數 ---
    def save_subtask_callback():
        name = name_var.get().strip()
        due_date_str = date_var.get().strip()
        hour_str = hour_var.get().strip()
        minute_str = minute_var.get().strip()
        color_name = color_var.get()
        details = details_text.get("1.0", tk.END).strip()

        if not name:
            messagebox.showwarning("警告", "子任務名稱不能為空！", parent=dialog)
            return

        # 日期和時間驗證
        final_due_date, final_due_time = "", ""
        if due_date_str:
            try:
                datetime.strptime(due_date_str, '%Y-%m-%d')
                final_due_date = due_date_str
            except ValueError:
                 messagebox.showwarning("警告", "日期格式不正確 (應為 YYYY-MM-DD)。", parent=dialog)
                 return

        if hour_str or minute_str:
            if not final_due_date:
                 messagebox.showwarning("警告", "請先設定到期日期，才能設定時間。", parent=dialog)
                 return
            if hour_str and minute_str:
                try:
                    hour_val, minute_val = int(hour_str), int(minute_str)
                    if not (0 <= hour_val <= 23 and 0 <= minute_val <= 59): raise ValueError("時間範圍錯誤")
                    final_due_time = f"{hour_val:02d}:{minute_val:02d}"
                except ValueError:
                    messagebox.showwarning("警告", "時間格式不正確 (HH:MM 應在 00:00 - 23:59 之間)。", parent=dialog)
                    return
            else:
                 messagebox.showwarning("警告", "請輸入完整時間 (小時和分鐘)。", parent=dialog)
                 return

        # 準備資料
        sub_data = {
            "id": subtask_data['id'] if is_editing else generate_uuid(), # 使用導入的函數
            "name": name,
            "due_date": final_due_date,
            "due_time": final_due_time,
            "color_name": color_name,
            "details": details,
            "archived": False # 新增或編輯時，狀態都是未封存
        }

        try:
            # 更新 tool_instance 中的資料
            if is_editing:
                # 找到對應的子任務並更新
                subtask_index = next(i for i, sub in enumerate(tool_instance.task_groups[group_index]['sub_tasks']) if sub['id'] == sub_data['id'])
                tool_instance.task_groups[group_index]['sub_tasks'][subtask_index].update(sub_data)
            else:
                # 添加新的子任務
                tool_instance.task_groups[group_index]['sub_tasks'].append(sub_data)

            dialog.destroy() # 關閉對話框
            tool_instance.render_all_tasks() # 調用主實例的渲染方法
        except (IndexError, StopIteration):
             messagebox.showerror("錯誤", "儲存子任務時出錯，找不到對應的任務。", parent=dialog)
        except Exception as e:
             messagebox.showerror("錯誤", f"儲存子任務時發生未預期錯誤: {e}", parent=dialog)
             # log_error(tool_instance, "Save Subtask Error", ...)

    # --- 按鈕 ---
    button_frame = ttk.Frame(frame)
    button_frame.grid(row=5, column=0, columnspan=5, pady=10)

    save_btn = ttk.Button(button_frame, text="確定", command=save_subtask_callback)
    save_btn.pack(side=tk.LEFT, padx=5)
    cancel_btn = ttk.Button(button_frame, text="取消", command=dialog.destroy)
    cancel_btn.pack(side=tk.LEFT, padx=5)

    # --- 設定列和行的權重 ---
    frame.grid_columnconfigure(1, weight=1) # 讓輸入框可以擴展
    frame.grid_rowconfigure(4, weight=1)    # 讓詳細內容區域可以擴展

    # --- 讓 Enter 也能觸發儲存 (避免在 Text 輸入時觸發) ---
    dialog.bind('<Return>', lambda event=None: save_subtask_callback() if dialog.focus_get() != details_text else None)

    tool_instance.root.wait_window(dialog) # 等待對話框關閉

# --- 顯示封存區視窗 ---
def show_archived_tasks_window(tool_instance):
    """
    顯示包含已封存任務的視窗。
    :param tool_instance: TextCorrectionTool 的實例。
    """
    archive_window = tk.Toplevel(tool_instance.root)
    archive_window.title("封存區")
    archive_window.geometry("450x550") # 調整大小
    archive_window.grab_set()
    archive_window.transient(tool_instance.root)

    list_frame = ttk.Frame(archive_window, padding="10")
    list_frame.pack(fill="both", expand=True)

    if not tool_instance.archived_tasks:
        ttk.Label(list_frame, text="封存區目前是空的。").pack(pady=20)
        # 添加關閉按鈕
        close_button = ttk.Button(archive_window, text="關閉", command=archive_window.destroy)
        close_button.pack(pady=10)
        tool_instance.root.wait_window(archive_window)
        return

    # 創建帶捲軸的列表框
    scrollbar = ttk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")
    listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, width=60, height=20) # 加寬

    # 填充列表框
    for index, task in enumerate(tool_instance.archived_tasks):
        # 顯示更詳細的資訊，並添加索引以便追蹤
        date_str = task.get('due_date', '無日期')
        time_str = task.get('due_time', '')
        display_text = f"{index+1}. {task.get('name', '無名稱')} ({date_str} {time_str})"
        listbox.insert(tk.END, display_text)
        # 根據顏色設定背景色 (可選)
        color_code = get_color_code(task.get('color_name'))
        listbox.itemconfig(index, {'bg': color_code})


    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # --- 取消封存和刪除按鈕 ---
    button_frame = ttk.Frame(archive_window)
    button_frame.pack(pady=10)

    def unarchive_selected():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "請先選擇要取消封存的任務。", parent=archive_window)
            return
        selected_index = selected_indices[0]

        try:
            # 從封存列表移除，使用索引
            task_to_unarchive = tool_instance.archived_tasks.pop(selected_index)
            task_to_unarchive['archived'] = False

            # 簡單起見：添加到第一個任務組或創建新組
            if tool_instance.task_groups:
                 tool_instance.task_groups[0]['sub_tasks'].append(task_to_unarchive)
            else:
                 tool_instance.task_groups.append({"main_task_name": "未分類", "sub_tasks": [task_to_unarchive]})

            # 關閉舊視窗，重新渲染主視窗，再打開新視窗
            archive_window.destroy()
            tool_instance.render_all_tasks()
            show_archived_tasks_window(tool_instance) # 遞迴調用以刷新

        except IndexError:
             messagebox.showerror("錯誤", "選擇的索引無效。", parent=archive_window)
        except Exception as e:
             messagebox.showerror("錯誤", f"取消封存時發生錯誤: {e}", parent=archive_window)

    def delete_selected_permanently():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "請先選擇要永久刪除的任務。", parent=archive_window)
            return
        selected_index = selected_indices[0]

        confirm = messagebox.askyesno("確認刪除", "確定要永久刪除選定的封存任務嗎？此操作無法復原。", parent=archive_window)
        if confirm:
            try:
                # 從封存列表永久移除
                del tool_instance.archived_tasks[selected_index]

                # 關閉舊視窗，重新渲染主視窗，再打開新視窗
                archive_window.destroy()
                tool_instance.render_all_tasks() # 主視窗不需要重繪，但習慣上保留
                show_archived_tasks_window(tool_instance) # 遞迴調用以刷新

            except IndexError:
                 messagebox.showerror("錯誤", "選擇的索引無效。", parent=archive_window)
            except Exception as e:
                 messagebox.showerror("錯誤", f"刪除任務時發生錯誤: {e}", parent=archive_window)


    unarchive_button = ttk.Button(button_frame, text="取消封存選定項", command=unarchive_selected)
    unarchive_button.pack(side=tk.LEFT, padx=5)

    delete_button = ttk.Button(button_frame, text="永久刪除選定項", command=delete_selected_permanently)
    delete_button.pack(side=tk.LEFT, padx=5)

    close_button = ttk.Button(button_frame, text="關閉", command=archive_window.destroy)
    close_button.pack(side=tk.LEFT, padx=5)

    tool_instance.root.wait_window(archive_window)
