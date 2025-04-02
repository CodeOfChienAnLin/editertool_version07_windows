# -*- coding: utf-8 -*-
"""
代辦事項功能的 Canvas 渲染邏輯
"""
import tkinter as tk
from tkinter import font, ttk

# 導入資料模組中的常數和函數
from todo_01_data import (
    get_color_code,
    MAIN_TASK_WIDTH, MAIN_TASK_PADDING_Y, MAIN_TASK_SPACING_X, MAIN_TASK_SPACING_Y,
    SUBTASK_WIDTH_MARGIN, SUBTASK_INTERNAL_PADDING_Y, SUBTASK_SPACING_Y,
    SUBTASK_CIRCLE_RADIUS, SUBTASK_CIRCLE_MARGIN, NAME_TIME_GAP, SUBTASK_HEIGHT_BASE
)
# 導入處理函數模組 (用於綁定事件)
# 避免循環導入，事件處理將由主應用程式的包裝方法調用
# from todo_04_handlers import handle_add_subtask_click, handle_edit_subtask_click, handle_archive_subtask_click, handle_add_main_task_click

# --- 更新捲動區域 ---
def update_todo_scroll_region(canvas): # 直接接收 canvas
    """更新代辦事項 Canvas 的捲動區域"""
    # canvas = tool_instance.todo_canvas # 改為直接使用傳入的 canvas
    if canvas:
        canvas.update_idletasks() # 確保所有元件尺寸已更新
        bbox = canvas.bbox("all") # 獲取所有元素的邊界
        if bbox:
            # 添加一些額外空間，特別是底部和右側
            scroll_region = (0, 0, bbox[2] + MAIN_TASK_SPACING_X + 50, bbox[3] + MAIN_TASK_SPACING_Y + 100) # 增加底部空間
            canvas.config(scrollregion=scroll_region)
        else:
            # 如果畫布為空，設定為畫布的當前可見區域大小
            canvas.config(scrollregion=(0, 0, canvas.winfo_width(), canvas.winfo_height()))

# --- 渲染所有代辦事項內容到 Canvas ---
def render_all_tasks(canvas, task_groups, main_task_font, sub_task_font, sub_task_time_font, tool_instance): # 保留 tool_instance 用於回呼
    """
    重新繪製代辦事項 Canvas 上的所有主任務、子任務和按鈕。
    :param canvas: 要繪製的 Tkinter Canvas。
    :param task_groups: 包含任務資料的列表。
    :param main_task_font: 主任務字體。
    :param sub_task_font: 子任務字體。
    :param sub_task_time_font: 子任務時間字體。
    :param tool_instance: TextCorrectionTool 的實例，用於事件回呼。
    """
    # 直接使用傳入的參數，不再從 tool_instance 獲取
    # canvas = tool_instance.todo_canvas
    # task_groups = tool_instance.task_groups
    # main_task_font = tool_instance.todo_main_task_font
    # sub_task_font = tool_instance.todo_sub_task_font
    # sub_task_time_font = tool_instance.todo_sub_task_time_font

    # 檢查必要參數是否存在
    if not all([canvas, task_groups is not None, main_task_font, sub_task_font, sub_task_time_font, tool_instance]):
        # task_groups 可以是空列表，所以檢查 is not None
        print("錯誤：render_all_tasks 缺少必要的參數 (canvas, task_groups, 字體, 或 tool_instance)")
        return # 如果缺少必要參數則返回

    # --- 預先測量按鈕高度 ---
    # 注意：按鈕現在是 ttk.Button，字體可能不同
    try:
        # 使用 ttk 按鈕和可能的應用程式字體進行測量
        temp_btn_font = font.Font(size=18) # 假設 "+" 按鈕字體大小
        temp_btn = ttk.Button(canvas, text="+", style="Toolbutton") # 使用 ttk 按鈕和可能的樣式
        canvas.update_idletasks()
        measured_button_height = temp_btn.winfo_reqheight()
        temp_btn.destroy()
    except tk.TclError:
        measured_button_height = 40 # 後備高度
    actual_button_space_needed = measured_button_height + SUBTASK_SPACING_Y

    # 1. 清空 Canvas
    canvas.delete("all")

    # 2. 計算並繪製主任務區塊及其內容
    # 獲取畫布可用寬度，用於可能的換行邏輯
    canvas_width = canvas.winfo_width()
    if canvas_width <= 1: canvas_width = 800 # 如果寬度尚未確定，給一個預設值

    current_x = MAIN_TASK_SPACING_X
    current_y = MAIN_TASK_SPACING_Y
    max_y_in_row = current_y # 追蹤目前行的最大 Y 座標

    for group_index, group_data in enumerate(task_groups):
        main_task_name = group_data.get("main_task_name", "未命名區塊")
        sub_tasks = group_data.get("sub_tasks", [])

        # A. 計算主任務區塊的總高度
        total_block_height = MAIN_TASK_PADDING_Y # 頂部邊距

        # 計算主任務名稱高度 (考慮換行)
        # 創建臨時文字項以獲取邊界框
        temp_main_id = canvas.create_text(0, 0, text=main_task_name, font=main_task_font,
                                          width=MAIN_TASK_WIDTH - 10, anchor='nw', tags="temp_measure")
        main_name_bbox = canvas.bbox(temp_main_id)
        main_name_height = main_name_bbox[3] - main_name_bbox[1] if main_name_bbox else main_task_font.metrics('linespace')
        canvas.delete("temp_measure") # 刪除臨時項
        total_block_height += main_name_height + MAIN_TASK_PADDING_Y # 名稱高度 + 下方邊距

        # 計算所有子任務佔用的總高度
        subtask_area_height = 0
        # 子任務文字區域寬度
        text_area_width_sub = MAIN_TASK_WIDTH - (SUBTASK_WIDTH_MARGIN + SUBTASK_CIRCLE_MARGIN + SUBTASK_CIRCLE_RADIUS * 2 + 10) # 考慮封存按鈕和邊距

        for sub_index, sub_data in enumerate(sub_tasks):
            # 計算子任務名稱高度
            temp_sub_name_id = canvas.create_text(0, 0, text=sub_data['name'], font=sub_task_font,
                                                  width=text_area_width_sub, anchor='nw', tags="temp_measure")
            sub_name_bbox = canvas.bbox(temp_sub_name_id)
            sub_name_height = sub_name_bbox[3] - sub_name_bbox[1] if sub_name_bbox else sub_task_font.metrics('linespace')
            canvas.delete("temp_measure")

            # 計算時間顯示高度
            time_display = ""
            due_date = sub_data.get('due_date')
            due_time = sub_data.get('due_time')
            if due_date:
                time_display = due_date
                if due_time: time_display += f" {due_time}"

            sub_time_height = 0
            if time_display:
                temp_sub_time_id = canvas.create_text(0, 0, text=time_display, font=sub_task_time_font,
                                                      width=text_area_width_sub, anchor='nw', tags="temp_measure")
                sub_time_bbox = canvas.bbox(temp_sub_time_id)
                sub_time_height = sub_time_bbox[3] - sub_time_bbox[1] if sub_time_bbox else 0
                canvas.delete("temp_measure")

            # 計算單個子任務總高度
            name_time_gap = NAME_TIME_GAP if time_display else 0
            single_sub_height = SUBTASK_INTERNAL_PADDING_Y + sub_name_height + name_time_gap + sub_time_height + SUBTASK_INTERNAL_PADDING_Y
            # 確保有個最小高度，避免太扁
            sub_data['height'] = max(SUBTASK_HEIGHT_BASE * 0.6, single_sub_height) # 使用常數的比例作為最小高度
            subtask_area_height += sub_data['height']
            if sub_index > 0: subtask_area_height += SUBTASK_SPACING_Y # 添加子任務間距

        total_block_height += subtask_area_height # 添加子任務區域總高度
        total_block_height += actual_button_space_needed # 添加內部 "+" 按鈕所需空間
        total_block_height += MAIN_TASK_PADDING_Y # 底部邊距

        # --- 檢查是否需要換行 ---
        if current_x + MAIN_TASK_WIDTH > canvas_width and current_x > MAIN_TASK_SPACING_X:
            # 如果添加此區塊會超出畫布寬度，並且目前不是第一列
            current_x = MAIN_TASK_SPACING_X # X 回到最左邊
            current_y = max_y_in_row + MAIN_TASK_SPACING_Y # Y 移動到目前行的最大高度下方
            max_y_in_row = current_y # 重設目前行的最大 Y

        # B. 繪製主任務區塊外框
        x1, y1 = current_x, current_y
        x2, y2 = current_x + MAIN_TASK_WIDTH, current_y + total_block_height
        canvas.create_rectangle(
            x1, y1, x2, y2, outline="gray", fill="#FAFAFA", width=1, # 稍微調整顏色和線寬
            tags=("main_task_rect", f"group_{group_index}")
        )

        # C. 繪製主任務名稱
        title_tag = f"main_title_{group_index}"
        canvas.create_text(
            x1 + MAIN_TASK_WIDTH / 2, y1 + MAIN_TASK_PADDING_Y,
            text=main_task_name, font=main_task_font, width=MAIN_TASK_WIDTH - 10,
            anchor='n', tags=("main_task_text", f"group_{group_index}", title_tag)
        )
        # 獲取實際繪製後標題的底部 Y
        canvas.update_idletasks()
        title_bbox = canvas.bbox(title_tag)
        title_bottom_y = title_bbox[3] if title_bbox else (y1 + MAIN_TASK_PADDING_Y + main_name_height)
        current_internal_y = title_bottom_y + MAIN_TASK_PADDING_Y # 子任務從標題下方開始

        # D. 繪製子任務
        for sub_index, sub_data in enumerate(sub_tasks):
            sub_x1 = x1 + SUBTASK_WIDTH_MARGIN / 2
            sub_y1 = current_internal_y
            sub_x2 = x2 - SUBTASK_WIDTH_MARGIN / 2
            sub_y2 = sub_y1 + sub_data['height']
            sub_id = sub_data['id']
            sub_color_code = get_color_code(sub_data.get('color_name'))

            # 子任務背景矩形 (綁定編輯事件)
            rect_tag = f"subtask_rect_{sub_id}"
            canvas.create_rectangle(sub_x1, sub_y1, sub_x2, sub_y2, fill=sub_color_code, outline="#E0E0E0", width=1, tags=("subtask", rect_tag, f"group_{group_index}"))
            # 使用 lambda 傳遞參數給事件處理器
            canvas.tag_bind(rect_tag, "<Button-1>", lambda e, gi=group_index, si=sub_id: tool_instance.edit_sub_task(gi, si)) # 調用實例方法

            # 封存按鈕 (圓圈)
            circle_cx = sub_x1 + SUBTASK_CIRCLE_MARGIN + SUBTASK_CIRCLE_RADIUS
            circle_cy = sub_y1 + sub_data['height'] / 2
            circle_tag = f"archive_btn_{sub_id}"
            canvas.create_oval(circle_cx - SUBTASK_CIRCLE_RADIUS, circle_cy - SUBTASK_CIRCLE_RADIUS,
                               circle_cx + SUBTASK_CIRCLE_RADIUS, circle_cy + SUBTASK_CIRCLE_RADIUS,
                               fill="white", outline="gray", width=1, activefill="#E0E0E0", # 添加點擊效果
                               tags=("subtask", "archive_button", circle_tag, f"group_{group_index}"))
            canvas.tag_bind(circle_tag, "<Button-1>", lambda e, gi=group_index, si=sub_id: tool_instance.archive_sub_task(gi, si)) # 調用實例方法

            # 子任務文字 (名稱和時間)
            text_x = circle_cx + SUBTASK_CIRCLE_RADIUS + 10 # 文字起始 X
            text_area_w = sub_x2 - text_x # 文字可用寬度

            # 繪製名稱
            name_tag = f"subtask_name_{sub_id}"
            canvas.create_text(text_x, sub_y1 + SUBTASK_INTERNAL_PADDING_Y, text=sub_data['name'],
                               font=sub_task_font, anchor='nw', width=text_area_w,
                               tags=("subtask", "subtask_text", name_tag, f"group_{group_index}"))
            name_bbox = canvas.bbox(name_tag)
            name_h = name_bbox[3] - name_bbox[1] if name_bbox else sub_task_font.metrics('linespace')

            # 繪製時間 (如果存在)
            time_display = ""
            due_date = sub_data.get('due_date')
            due_time = sub_data.get('due_time')
            if due_date:
                time_display = due_date
                if due_time: time_display += f" {due_time}"

            if time_display:
                time_tag = f"subtask_time_{sub_id}"
                name_time_gap = NAME_TIME_GAP
                time_y = sub_y1 + SUBTASK_INTERNAL_PADDING_Y + name_h + name_time_gap
                canvas.create_text(text_x, time_y, text=time_display, font=sub_task_time_font,
                                   anchor='nw', width=text_area_w, fill="gray", # 時間用灰色
                                   tags=("subtask", "subtask_text", time_tag, f"group_{group_index}"))

            # 更新下一個子任務的 Y 座標
            current_internal_y = sub_y2 + SUBTASK_SPACING_Y

        # --- E. 繪製內部 "+" 按鈕 (使用 Canvas 創建的按鈕外觀) ---
        # 改為在 Canvas 上繪製 "+" 形狀，並綁定事件
        button_y_position = current_internal_y # 按鈕頂部 Y
        button_center_x = x1 + MAIN_TASK_WIDTH / 2
        button_center_y = button_y_position + measured_button_height / 2
        button_size = measured_button_height * 0.3 # 加號大小基於按鈕高度
        add_sub_tag = f"add_sub_btn_{group_index}"

        # 繪製按鈕背景 (可選，使其看起來像按鈕)
        canvas.create_rectangle(button_center_x - measured_button_height/2, button_y_position,
                                button_center_x + measured_button_height/2, button_y_position + measured_button_height,
                                fill="#F0F0F0", outline="gray", width=1, activefill="#E0E0E0",
                                tags=("subtask_add_button_bg", add_sub_tag, f"group_{group_index}"))
        # 繪製加號
        canvas.create_line(button_center_x - button_size, button_center_y,
                           button_center_x + button_size, button_center_y,
                           width=2, fill="gray", tags=("subtask_add_button_icon", add_sub_tag, f"group_{group_index}"))
        canvas.create_line(button_center_x, button_center_y - button_size,
                           button_center_x, button_center_y + button_size,
                           width=2, fill="gray", tags=("subtask_add_button_icon", add_sub_tag, f"group_{group_index}"))
        # 綁定事件到整個按鈕區域 (背景和圖標)
        canvas.tag_bind(add_sub_tag, "<Button-1>", lambda e, gi=group_index: tool_instance.add_sub_task(gi)) # 調用實例方法


        # 更新下一個主任務區塊的 X 座標
        current_x += MAIN_TASK_WIDTH + MAIN_TASK_SPACING_X
        max_y_in_row = max(max_y_in_row, y2) # 更新目前行的最大 Y

    # --- F. 繪製主 "+" 按鈕 (固定在左上角) ---
    main_add_btn_size = 30 # 主按鈕大小 (可調整)
    main_add_btn_x = MAIN_TASK_SPACING_X + main_add_btn_size / 2
    main_add_btn_y = MAIN_TASK_SPACING_Y + main_add_btn_size / 2
    main_add_tag = "main_add_button"

    # 繪製主按鈕背景 (可選，或只繪製加號)
    # canvas.create_rectangle(main_add_btn_x - main_add_btn_size / 2, main_add_btn_y - main_add_btn_size / 2,
    #                         main_add_btn_x + main_add_btn_size / 2, main_add_btn_y + main_add_btn_size / 2,
    #                         fill="#E8E8E8", outline="darkgray", width=1, activefill="#D8D8D8",
    #                         tags=("main_add_button_bg", main_add_tag))

    # 繪製主加號 (黑色，更明顯)
    main_add_icon_size = main_add_btn_size * 0.4 # 加號大小
    canvas.create_line(main_add_btn_x - main_add_icon_size, main_add_btn_y,
                       main_add_btn_x + main_add_icon_size, main_add_btn_y,
                       width=3, fill="black", tags=("main_add_button_icon", main_add_tag)) # 黑色加號
    canvas.create_line(main_add_btn_x, main_add_btn_y - main_add_icon_size,
                       main_add_btn_x, main_add_btn_y + main_add_icon_size,
                       width=3, fill="black", tags=("main_add_button_icon", main_add_tag)) # 黑色加號

    # 綁定事件到加號圖標
    canvas.tag_bind(main_add_tag, "<Button-1>", lambda e: tool_instance.add_main_task()) # 調用實例方法
    # 添加手型游標
    canvas.tag_bind(main_add_tag, "<Enter>", lambda e: canvas.config(cursor="hand2"))
    canvas.tag_bind(main_add_tag, "<Leave>", lambda e: canvas.config(cursor=""))


    # --- G. 更新捲動區域 ---
    # 將主 "+" 按鈕的位置也考慮進去，確保它總是在可見區域
    # 獲取所有任務區塊的邊界
    tasks_bbox = canvas.bbox("main_task_rect")
    if tasks_bbox:
        # 如果有任務區塊，滾動區域包含它們和 "+" 按鈕
        final_width = max(tasks_bbox[2] + MAIN_TASK_SPACING_X + 50, main_add_btn_x + main_add_btn_size)
        final_height = max(tasks_bbox[3] + MAIN_TASK_SPACING_Y + 100, main_add_btn_y + main_add_btn_size)
        scroll_region = (0, 0, final_width, final_height)
        canvas.config(scrollregion=scroll_region)
    else:
        # 如果沒有任務區塊，滾動區域至少包含 "+" 按鈕
        scroll_region = (0, 0, main_add_btn_x + main_add_btn_size + 50, main_add_btn_y + main_add_btn_size + 50)
        canvas.config(scrollregion=scroll_region)
        # 確保即使沒有任務，也能更新滾動區域
        update_todo_scroll_region(canvas) # 調用更新函數，傳遞 canvas
