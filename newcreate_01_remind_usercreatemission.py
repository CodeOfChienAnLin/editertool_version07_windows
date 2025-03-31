# 匯入 Tkinter 相關模組
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, font, ttk, Text # 引入 Text
import json
import uuid # 用於生成唯一 ID
from datetime import datetime # 用於處理日期時間

# 嘗試匯入 tkcalendar，如果失敗則提示安裝
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False
    print("警告：未找到 tkcalendar 庫，日期選擇將使用普通輸入框。")
    print("請使用 'pip install tkcalendar' 安裝。")

# --- 全域變數 ---
task_groups = [] # 改名為 task_groups，每個元素代表一個主任務區塊及其子任務
archived_tasks = [] # 儲存已封存的子任務
content_canvas = None # Canvas 物件
root = None
main_task_font = None # 主任務名稱字體
sub_task_font = None # 子任務名稱字體
sub_task_time_font = None # 子任務時間字體

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
MAIN_TASK_WIDTH = 270        # 主任務區塊寬度 (稍微加寬以容納子任務)
MAIN_TASK_PADDING_Y = 10     # 主任務名稱上下邊距
MAIN_TASK_SPACING_X = 15     # 主任務區塊之間的水平間距
MAIN_TASK_SPACING_Y = 15     # 主任務區塊之間的垂直間距 (如果換行)

SUBTASK_HEIGHT_BASE = 60    # 子任務基礎高度 (估計值，會動態調整)
SUBTASK_WIDTH_MARGIN = 20    # 子任務左右邊距總和
SUBTASK_INTERNAL_PADDING_Y = 8 # 子任務內部上下邊距 (增加)
SUBTASK_SPACING_Y = 8        # 子任務之間的垂直間距
SUBTASK_CIRCLE_RADIUS = 8    # 封存按鈕圓圈半徑
SUBTASK_CIRCLE_MARGIN = 10   # 封存按鈕左邊距
NAME_TIME_GAP = 3            # 子任務名稱和時間之間的垂直間距

# ADD_BUTTON_SIZE constant removed, will measure dynamically

# --- 輔助函數：獲取顏色值 ---
def get_color_code(name):
    return SUBTASK_COLORS.get(name, SUBTASK_COLORS["預設"])

# --- 儲存 JSON ---
def save_to_json():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".json", filetypes=[("JSON files", "*.json")], title="選擇儲存路徑"
    )
    if not file_path: return

    # 準備儲存的數據 (過濾掉臨時 UI 數據)
    data_to_save = {
        "task_groups": [
            {
                "main_task_name": group.get("main_task_name", ""),
                "sub_tasks": [
                    {k: v for k, v in sub.items() if k not in ['y_pos', 'height']} # 過濾掉 y_pos 和 height
                    for sub in group.get("sub_tasks", []) if not sub.get("archived") # 只儲存未封存的
                ]
            } for group in task_groups
        ],
        "archived_tasks": [
             {k: v for k, v in sub.items() if k not in ['y_pos', 'height']}
             for sub in archived_tasks
        ] # 儲存封存的任務
    }
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "資料已成功儲存至 JSON 檔案！", parent=root)
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存失敗: {e}", parent=root)

# --- 更新捲動區域 ---
def update_scroll_region():
    global content_canvas
    if content_canvas:
        content_canvas.update_idletasks()
        bbox = content_canvas.bbox("all") # 獲取所有元素的邊界
        if bbox:
            # 添加一些額外空間
            scroll_region = (0, 0, bbox[2] + MAIN_TASK_SPACING_X, bbox[3] + MAIN_TASK_SPACING_Y + 50)
            content_canvas.config(scrollregion=scroll_region)
        else:
             content_canvas.config(scrollregion=(0, 0, content_canvas.winfo_width(), content_canvas.winfo_height()))

# --- 處理主任務 "+" 按鈕點擊事件 ---
def handle_add_main_task_click():
    """ 彈出視窗讓使用者輸入主任務區塊名稱 """
    task_name = simpledialog.askstring("新增任務區塊", "請輸入任務區塊名稱:", parent=root)
    if task_name:
        task_groups.append({
            "main_task_name": task_name,
            "sub_tasks": []
        })
        render_all() # 重新渲染所有內容

# --- 處理子任務 "+" 按鈕點擊事件 ---
def handle_add_subtask_click(group_index):
    """ 彈出視窗讓使用者輸入子任務詳細資訊 """
    show_subtask_dialog(group_index=group_index) # 傳遞主任務索引

# --- 處理編輯子任務點擊事件 ---
def handle_edit_subtask_click(group_index, subtask_id):
    """ 根據 group_index 和 subtask_id 找到子任務並打開編輯視窗 """
    try:
        subtask_to_edit = next(sub for sub in task_groups[group_index]['sub_tasks'] if sub['id'] == subtask_id)
        show_subtask_dialog(group_index=group_index, subtask_data=subtask_to_edit) # 傳遞現有數據
    except (IndexError, StopIteration):
        messagebox.showerror("錯誤", "找不到要編輯的子任務。", parent=root)

# --- 處理封存子任務點擊事件 ---
def handle_archive_subtask_click(group_index, subtask_id):
    """ 將子任務標記為封存並移動到封存列表 """
    global archived_tasks
    try:
        group = task_groups[group_index]
        subtask_index = -1
        for i, sub in enumerate(group['sub_tasks']):
            if sub['id'] == subtask_id:
                subtask_index = i
                break

        if subtask_index != -1:
            subtask_to_archive = group['sub_tasks'].pop(subtask_index) # 從原列表移除
            subtask_to_archive['archived'] = True # 標記為封存
            archived_tasks.append(subtask_to_archive) # 添加到封存列表
            render_all() # 重新渲染
        else:
             messagebox.showerror("錯誤", "找不到要封存的子任務。", parent=root)

    except IndexError:
        messagebox.showerror("錯誤", "找不到指定的任務組。", parent=root)

# --- 顯示/編輯子任務的彈出視窗 ---
def show_subtask_dialog(group_index, subtask_data=None):
    """
    創建或編輯子任務的彈出視窗。
    :param group_index: 子任務所屬的主任務索引。
    :param subtask_data: 如果是編輯，傳入現有的子任務字典；否則為 None。
    """
    global root

    is_editing = subtask_data is not None
    dialog_title = "編輯子任務" if is_editing else "新增子任務"

    dialog = tk.Toplevel(root)
    dialog.title(dialog_title)
    dialog.geometry("400x300") # <--- 修改視窗大小
    dialog.resizable(False, False)
    dialog.grab_set()
    dialog.transient(root)

    # --- 創建輸入元件 ---
    frame = ttk.Frame(dialog, padding="10")
    frame.pack(fill="both", expand=True)

    # 名稱
    ttk.Label(frame, text="子任務名稱:").grid(row=0, column=0, sticky="w", pady=2)
    name_var = tk.StringVar(value=subtask_data['name'] if is_editing else "")
    name_entry = ttk.Entry(frame, textvariable=name_var, width=40)
    # Span across columns 1 to 4 (Label, Hour, Colon, Minute, Scrollbar)
    name_entry.grid(row=0, column=1, columnspan=4, sticky="ew", pady=2)
    name_entry.focus_set()

    # 到期日期
    ttk.Label(frame, text="到期日期:").grid(row=1, column=0, sticky="w", pady=2)
    # Initialize empty for new tasks
    initial_date = subtask_data.get('due_date', '') if is_editing else ""
    date_var = tk.StringVar(value=initial_date)
    if HAS_TKCALENDAR:
        # Use DateEntry, allow direct editing, set date_pattern for saving
        date_entry = DateEntry(frame, textvariable=date_var, date_pattern='yyyy-mm-dd', width=12,
                                background='darkblue', foreground='white', borderwidth=2,
                                firstweekday='sunday', showweeknumbers=False, # Optional styling
                                state='normal' # Allow typing
                               )
        # Clear the entry initially if not editing and no initial value
        if not initial_date:
             date_entry.delete(0, tk.END)
        date_entry.grid(row=1, column=1, columnspan=4, sticky="w", pady=2)
    else:
        # Fallback to Entry if tkcalendar not installed
        date_entry = ttk.Entry(frame, textvariable=date_var, width=15)
        date_entry.grid(row=1, column=1, sticky="w", pady=2)
        ttk.Label(frame, text="(YYYY-MM-DD)").grid(row=1, column=2, columnspan=3, sticky="w", padx=5)

    # 到期時間 (使用 Spinbox in an inner frame)
    ttk.Label(frame, text="到期時間:").grid(row=2, column=0, sticky="w", pady=2)
    # --- Inner frame for time widgets ---
    time_frame = ttk.Frame(frame)
    # Place the inner frame in the grid, aligned left
    time_frame.grid(row=2, column=1, columnspan=4, sticky="w")

    # -- 小時 Spinbox --
    hour_var = tk.StringVar()
    hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, wrap=True, width=3, format="%02.0f")
    hour_spinbox.pack(side=tk.LEFT, padx=(0, 2))
    # -- 分隔符 --
    ttk.Label(time_frame, text=":").pack(side=tk.LEFT, padx=0)
    # -- 分鐘 Spinbox --
    minute_var = tk.StringVar()
    minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, wrap=True, width=3, format="%02.0f")
    minute_spinbox.pack(side=tk.LEFT, padx=(2, 0))

    # --- 設定初始時間 (or empty for new) ---
    initial_hour = ""
    initial_minute = ""
    if is_editing and subtask_data.get('due_time'):
        try:
            h, m = map(int, subtask_data['due_time'].split(':'))
            initial_hour = f"{h:02d}"
            initial_minute = f"{m:02d}"
        except (ValueError, TypeError):
             initial_hour = ""
             initial_minute = ""
    hour_var.set(initial_hour)
    minute_var.set(initial_minute)
    # Clear spinboxes if not editing and no initial value
    if not initial_hour and not initial_minute:
        hour_spinbox.delete(0, tk.END)
        minute_spinbox.delete(0, tk.END)


    # 顏色選擇
    ttk.Label(frame, text="區塊顏色:").grid(row=3, column=0, sticky="w", pady=2)
    color_var = tk.StringVar(value=subtask_data['color_name'] if is_editing else COLOR_NAMES[0])
    color_combo = ttk.Combobox(frame, textvariable=color_var, values=COLOR_NAMES, state="readonly", width=10)
    color_combo.grid(row=3, column=1, columnspan=4, sticky="w", pady=2)

    # 詳細內容
    ttk.Label(frame, text="詳細內容:").grid(row=4, column=0, sticky="nw", pady=2)
    details_text = Text(frame, width=40, height=5, wrap="word") # Set Text size
    details_text.grid(row=4, column=1, columnspan=3, sticky="nsew", pady=2) # Span 3, make sticky nsew
    if is_editing and subtask_data.get('details'):
        details_text.insert("1.0", subtask_data['details'])
    # 添加捲軸到 Text 元件 (放在第4列)
    details_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=details_text.yview)
    details_scrollbar.grid(row=4, column=4, sticky="ns")
    details_text['yscrollcommand'] = details_scrollbar.set

    # --- 儲存按鈕 ---
    def save_subtask():
        name = name_var.get().strip()
        due_date_str = date_var.get().strip()
        hour_str = hour_var.get().strip()
        minute_str = minute_var.get().strip()
        color_name = color_var.get()
        details = details_text.get("1.0", tk.END).strip()

        if not name:
            messagebox.showwarning("警告", "子任務名稱不能為空！", parent=dialog)
            return

        # --- Date and Time Validation ---
        final_due_date = ""
        final_due_time = ""

        # Validate Date
        if due_date_str:
            try:
                datetime.strptime(due_date_str, '%Y-%m-%d')
                final_due_date = due_date_str
            except ValueError:
                 messagebox.showwarning("警告", "日期格式不正確 (應為 YYYY-MM-DD)。", parent=dialog)
                 return
        # If date is empty, time must also be empty (checked later)

        # Validate Time
        if hour_str or minute_str: # If either time field has input
            if not final_due_date: # Date must be set if time is set
                 messagebox.showwarning("警告", "請先設定到期日期，才能設定時間。", parent=dialog)
                 return
            if hour_str and minute_str: # Both hour and minute must be set
                try:
                    hour_val = int(hour_str)
                    minute_val = int(minute_str)
                    if not (0 <= hour_val <= 23 and 0 <= minute_val <= 59):
                        raise ValueError("Hour or minute out of range")
                    final_due_time = f"{hour_val:02d}:{minute_val:02d}"
                except ValueError:
                    messagebox.showwarning("警告", "時間格式不正確 (HH:MM 應在 00:00 - 23:59 之間)。", parent=dialog)
                    return
            else: # Only one part of time is filled
                 messagebox.showwarning("警告", "請輸入完整時間 (小時和分鐘)。", parent=dialog)
                 return
        elif final_due_date: # Date is set, but time is completely empty
            pass # Allow saving with only date
        # else: Both date and time are empty, which is allowed

        # --- Prepare data ---
        sub_data = {
            "id": subtask_data['id'] if is_editing else uuid.uuid4().hex,
            "name": name,
            "due_date": final_due_date, # Use validated or empty date
            "due_time": final_due_time, # Use validated or empty time
            "color_name": color_name,
            "details": details,
            "archived": False
        }

        try:
            if is_editing:
                subtask_index = next(i for i, sub in enumerate(task_groups[group_index]['sub_tasks']) if sub['id'] == sub_data['id'])
                task_groups[group_index]['sub_tasks'][subtask_index].update(sub_data)
            else:
                task_groups[group_index]['sub_tasks'].append(sub_data)

            dialog.destroy()
            render_all()
        except (IndexError, StopIteration):
             messagebox.showerror("錯誤", "儲存子任務時出錯。", parent=dialog)

    # 調整按鈕位置
    save_btn = ttk.Button(frame, text="確定", command=save_subtask)
    save_btn.grid(row=5, column=0, columnspan=5, pady=10) # Span all 5 columns

    # --- 設定列和行的權重 ---
    frame.grid_columnconfigure(1, weight=1) # Allow first input column to expand
    frame.grid_columnconfigure(3, weight=0) # Minute spinbox column fixed width
    frame.grid_rowconfigure(4, weight=1)    # Allow details row to expand vertically

    # --- 讓 Enter 也能觸發儲存 ---
    dialog.bind('<Return>', lambda event=None: save_subtask() if dialog.focus_get() != details_text else None)

    root.wait_window(dialog)

# --- 顯示封存區視窗 ---
def show_archived_tasks_window():
    global root, archived_tasks

    archive_window = tk.Toplevel(root)
    archive_window.title("封存區")
    archive_window.geometry("400x500")
    archive_window.grab_set()
    archive_window.transient(root)

    list_frame = ttk.Frame(archive_window, padding="10")
    list_frame.pack(fill="both", expand=True)

    if not archived_tasks:
        ttk.Label(list_frame, text="封存區是空的。").pack(pady=20)
        return

    # 創建一個帶捲軸的列表區域
    scrollbar = ttk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")
    listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, width=50, height=20)

    for task in archived_tasks:
        display_text = f"{task.get('name', '無名稱')} ({task.get('due_date', '')} {task.get('due_time', '')})"
        listbox.insert(tk.END, display_text)

    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # --- (可選) 添加取消封存按鈕 ---
    def unarchive_selected():
        selected_indices = listbox.curselection()
        if not selected_indices:
            return
        selected_index = selected_indices[0]

        task_to_unarchive = archived_tasks.pop(selected_index) # 從封存列表移除
        task_to_unarchive['archived'] = False

        # 尋找原始的主任務組 (如果需要放回原位，這會複雜)
        # 簡單起見：添加到第一個任務組
        if task_groups:
             task_groups[0]['sub_tasks'].append(task_to_unarchive)
        else: # 如果沒有主任務組了，創建一個新的
             task_groups.append({"main_task_name": "未分類", "sub_tasks": [task_to_unarchive]})

        archive_window.destroy() # 關閉封存視窗
        render_all() # 重新渲染主視窗
        show_archived_tasks_window() # 重新打開更新後的封存視窗 (可選)


    unarchive_button = ttk.Button(archive_window, text="取消封存選定項", command=unarchive_selected)
    unarchive_button.pack(pady=10)


    root.wait_window(archive_window)


# --- 渲染所有內容到 Canvas ---
def render_all():
    """
    重新繪製 Canvas 上的所有主任務、子任務和按鈕。
    """
    global content_canvas, task_groups, main_task_font, sub_task_font, sub_task_time_font
    if not content_canvas: return

    # --- Measure button height beforehand ---
    try:
        temp_btn = tk.Button(content_canvas, text="+", font=("Arial", 24)) # Use updated font size for measurement
        content_canvas.update_idletasks() # Ensure widget is processed
        measured_button_height = temp_btn.winfo_reqheight()
        temp_btn.destroy()
    except tk.TclError: # Handle case where canvas might not be ready initially
        measured_button_height = 60 # Fallback height
    # Add spacing needed above the button
    actual_button_space_needed = measured_button_height + SUBTASK_SPACING_Y

    # 確保字體物件已創建
    # ... (字體創建邏輯) ...

    # 1. 清空 Canvas
    content_canvas.delete("all")

    # 2. 計算並繪製主任務區塊及其內容
    current_x = MAIN_TASK_SPACING_X
    current_y = MAIN_TASK_SPACING_Y
    max_y_in_row = current_y

    for group_index, group_data in enumerate(task_groups):
        main_task_name = group_data.get("main_task_name", "未命名區塊")
        sub_tasks = group_data.get("sub_tasks", [])

        # A. 計算主任務區塊的總高度
        total_block_height = MAIN_TASK_PADDING_Y
        temp_main_id = content_canvas.create_text(0, 0, text=main_task_name, font=main_task_font, width=MAIN_TASK_WIDTH - 10, anchor='nw', tags="temp")
        main_name_bbox = content_canvas.bbox(temp_main_id)
        main_name_height = main_name_bbox[3] - main_name_bbox[1] if main_name_bbox else main_task_font.metrics('linespace')
        content_canvas.delete("temp")
        total_block_height += main_name_height + MAIN_TASK_PADDING_Y

        subtask_area_height = 0
        text_area_width_sub = MAIN_TASK_WIDTH - (SUBTASK_WIDTH_MARGIN + SUBTASK_CIRCLE_MARGIN + SUBTASK_CIRCLE_RADIUS * 2)
        for sub_index, sub_data in enumerate(sub_tasks):
            temp_sub_name_id = content_canvas.create_text(0,0, text=sub_data['name'], font=sub_task_font, width=text_area_width_sub, anchor='nw', tags="temp")
            sub_name_bbox = content_canvas.bbox(temp_sub_name_id)
            sub_name_height = sub_name_bbox[3] - sub_name_bbox[1] if sub_name_bbox else sub_task_font.metrics('linespace')
            content_canvas.delete("temp")
            # Display date and time only if they exist and are valid
            time_display = ""
            due_date = sub_data.get('due_date')
            due_time = sub_data.get('due_time')
            if due_date: # Only show time if date exists
                time_display = due_date
                if due_time:
                    time_display += f" {due_time}"

            temp_sub_time_id = content_canvas.create_text(0,0, text=time_display, font=sub_task_time_font, width=text_area_width_sub, anchor='nw', tags="temp")
            sub_time_bbox = content_canvas.bbox(temp_sub_time_id)
            sub_time_height = (sub_time_bbox[3] - sub_time_bbox[1]) if sub_time_bbox and time_display else 0
            content_canvas.delete("temp")
            # Add small vertical gap between name and time only if time is displayed
            name_time_gap = NAME_TIME_GAP if time_display else 0
            single_sub_height = SUBTASK_INTERNAL_PADDING_Y + sub_name_height + name_time_gap + sub_time_height + SUBTASK_INTERNAL_PADDING_Y
            sub_data['height'] = max(SUBTASK_HEIGHT_BASE / 2, single_sub_height)
            subtask_area_height += sub_data['height']
            if sub_index > 0: subtask_area_height += SUBTASK_SPACING_Y
        total_block_height += subtask_area_height # Add height of subtasks area

        # Add space for the button using the measured height
        total_block_height += actual_button_space_needed
        total_block_height += MAIN_TASK_PADDING_Y # Bottom padding


        # B. 繪製主任務區塊外框 (Draw first with calculated height)
        x1, y1 = current_x, current_y
        x2, y2 = current_x + MAIN_TASK_WIDTH, current_y + total_block_height # Use calculated total height
        content_canvas.create_rectangle(
            x1, y1, x2, y2, outline="black", fill="white", width=2,
            tags=("main_task_rect", f"group_{group_index}")
        )

        # C. 繪製主任務名稱
        content_canvas.create_text(
            x1 + MAIN_TASK_WIDTH / 2, y1 + MAIN_TASK_PADDING_Y,
            text=main_task_name, font=main_task_font, width=MAIN_TASK_WIDTH - 10,
            anchor='n', tags=("main_task_text", f"group_{group_index}", f"main_title_{group_index}") # Add specific tag
        )
        # --- Get actual title bottom Y after drawing ---
        content_canvas.update_idletasks() # Ensure bbox is ready
        title_bbox = content_canvas.bbox(f"main_title_{group_index}")
        # Use actual bottom if available, otherwise fallback to calculated position
        title_bottom_y = title_bbox[3] if title_bbox else (y1 + MAIN_TASK_PADDING_Y + main_name_height)
        # --- Revert: Start content right below title + padding ---
        current_internal_y = title_bottom_y + MAIN_TASK_PADDING_Y # Start below actual title bottom + padding

        # D. 繪製子任務
        for sub_index, sub_data in enumerate(sub_tasks):
            sub_x1 = x1 + SUBTASK_WIDTH_MARGIN / 2
            sub_y1 = current_internal_y
            sub_x2 = x2 - SUBTASK_WIDTH_MARGIN / 2
            sub_y2 = sub_y1 + sub_data['height']
            sub_id = sub_data['id']
            sub_color_code = get_color_code(sub_data.get('color_name'))
            rect_tag = f"subtask_rect_{sub_id}"
            content_canvas.create_rectangle(sub_x1, sub_y1, sub_x2, sub_y2, fill=sub_color_code, outline="black", width=1, tags=("subtask", rect_tag, f"group_{group_index}"))
            content_canvas.tag_bind(rect_tag, "<Button-1>", lambda e, gi=group_index, si=sub_id: handle_edit_subtask_click(gi, si))
            circle_cx = sub_x1 + SUBTASK_CIRCLE_MARGIN + SUBTASK_CIRCLE_RADIUS
            circle_cy = sub_y1 + sub_data['height'] / 2
            circle_tag = f"archive_btn_{sub_id}"
            content_canvas.create_oval(circle_cx - SUBTASK_CIRCLE_RADIUS, circle_cy - SUBTASK_CIRCLE_RADIUS, circle_cx + SUBTASK_CIRCLE_RADIUS, circle_cy + SUBTASK_CIRCLE_RADIUS, fill="white", outline="gray", width=1, tags=("subtask", "archive_button", circle_tag, f"group_{group_index}"))
            content_canvas.tag_bind(circle_tag, "<Button-1>", lambda e, gi=group_index, si=sub_id: handle_archive_subtask_click(gi, si))
            text_x = circle_cx + SUBTASK_CIRCLE_RADIUS + 10 # Increase space between circle and text
            text_area_w = sub_x2 - text_x
            name_tag = f"subtask_name_{sub_id}"
            content_canvas.create_text(text_x, sub_y1 + SUBTASK_INTERNAL_PADDING_Y, text=sub_data['name'], font=sub_task_font, anchor='nw', width=text_area_w, tags=("subtask", "subtask_text", name_tag, f"group_{group_index}"))
            name_bbox = content_canvas.bbox(name_tag)
            name_h = name_bbox[3] - name_bbox[1] if name_bbox else sub_task_font.metrics('linespace')
            # Display date and time only if they exist and are valid
            time_display = ""
            due_date = sub_data.get('due_date')
            due_time = sub_data.get('due_time')
            if due_date: # Only show time if date exists
                time_display = due_date
                if due_time:
                    time_display += f" {due_time}"

            if time_display: # Only draw time text if it exists
                time_tag = f"subtask_time_{sub_id}"
                # Add name_time_gap to the Y coordinate of the time text
                name_time_gap = NAME_TIME_GAP # Use the constant defined earlier
                time_y = sub_y1 + SUBTASK_INTERNAL_PADDING_Y + name_h + name_time_gap
                content_canvas.create_text(text_x, time_y, text=time_display, font=sub_task_time_font, anchor='nw', width=text_area_w, tags=("subtask", "subtask_text", time_tag, f"group_{group_index}"))
            current_internal_y = sub_y2 + SUBTASK_SPACING_Y


        # --- E. 繪製內部 "+" 按鈕 ---
        internal_add_btn_widget = tk.Button(
            content_canvas,
            text="+",
            font=("Arial", 24), # Font size reduced
            command=lambda gi=group_index: handle_add_subtask_click(gi),
            relief=tk.FLAT,
            borderwidth=0,
            bg='white',
            activebackground='white',
            fg="#D3D3D3", # Light gray color
            cursor="hand2"
        )
        # --- 定位按鈕在最後一個子任務下方 (或主標題下方) ---
        button_y_position = current_internal_y # Button top Y
        content_canvas.create_window(
            x1 + MAIN_TASK_WIDTH / 2, # Center horizontally
            button_y_position,        # Vertical position (button top)
            anchor='n',               # Use top anchor
            window=internal_add_btn_widget,
            tags=("subtask_add_button", f"group_{group_index}")
        )

        # 更新下一個主任務區塊的 X 座標
        current_x += MAIN_TASK_WIDTH + MAIN_TASK_SPACING_X
        max_y_in_row = max(max_y_in_row, y2) # Use calculated y2 for row height

        # --- (可選) 簡單的水平換行 ---
        # ...

    # --- F. 繪製主 "+" 按鈕 ---
    main_add_btn_widget = tk.Button(
        content_canvas,
        text="+",
        font=("Arial", 36), # Main button size unchanged
        command=handle_add_main_task_click,
        relief=tk.FLAT,
        borderwidth=0,
        bg='white',
        activebackground='white',
        fg="black",
        cursor="hand2"
    )
    main_add_btn_offset = 50 # Example fixed offset
    content_canvas.create_window(
        current_x + main_add_btn_offset / 2, # Adjust X based on offset/size
        current_y + main_add_btn_offset / 2, # Adjust Y based on offset/size
        anchor='center',
        window=main_add_btn_widget,
        tags="main_add_button"
    )

    # --- G. 更新捲動區域 ---
    update_scroll_region()

# --- 主視窗設定 ---
root = tk.Tk()
root.title("編審神器")
root.geometry("1000x700")

# --- 初始化字體 ---
try: main_task_font = font.Font(family="新細明體", size=14, weight="bold")
except: main_task_font = font.Font(size=14, weight="bold") # Fallback
try: sub_task_font = font.Font(family="標楷體", size=12) # Change to DFKai-SB
except: sub_task_font = font.Font(size=12) # Fallback
try: sub_task_time_font = font.Font(family="標楷體", size=10) # Change to DFKai-SB
except: sub_task_time_font = font.Font(size=10) # Fallback


# --- 左側選單 ---
menu_frame = tk.Frame(root, relief="solid", borderwidth=1, width=100)
menu_frame.pack(side="left", fill="y")
menu_frame.pack_propagate(False)
tk.Button(menu_frame, text="任務區", font=("Arial", 10), command=None).pack(fill="x", pady=2, padx=5)
tk.Button(menu_frame, text="封存區", font=("Arial", 10), command=show_archived_tasks_window).pack(fill="x", pady=2, padx=5)
tk.Button(menu_frame, text="下載", font=("Arial", 10), command=save_to_json).pack(side="bottom", fill="x", pady=5, padx=5)

# --- 可捲動的畫布區域 ---
canvas_frame = tk.Frame(root)
canvas_frame.pack(side="right", fill="both", expand=True)
content_canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=content_canvas.yview)
h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=content_canvas.xview)
content_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

# --- Grid 佈局 ---
content_canvas.grid(row=0, column=0, sticky="nsew")
v_scrollbar.grid(row=0, column=1, sticky="ns")
h_scrollbar.grid(row=1, column=0, sticky="ew")
canvas_frame.grid_rowconfigure(0, weight=1)
canvas_frame.grid_columnconfigure(0, weight=1)


# --- 初始渲染 ---
render_all()

# --- 主迴圈 ---
root.mainloop()
