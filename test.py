# 匯入 Tkinter 相關模組
import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk # ttk 提供主題化元件 (例如捲軸)
import json # 匯入 json 模組，用於處理 JSON 檔案

# --- 全域變數 ---
task_list = [] # 用於儲存任務資料的列表
inner_frame = None # 先宣告內部框架變數
content_canvas = None # 先宣告畫布變數，方便在函數中使用

# 儲存 JSON 資料到檔案 (此功能保留，供 "下載" 按鈕使用)
def save_to_json():
    # 彈出儲存檔案對話框，讓使用者選擇儲存路徑與檔名
    file_path = filedialog.asksaveasfilename(
        defaultextension=".json", # 預設副檔名
        filetypes=[("JSON files", "*.json")], # 檔案類型過濾
        title="選擇儲存路徑" # 對話框標題
    )
    # 如果使用者有選擇路徑 (而不是取消)
    if file_path:
        # 準備要儲存的資料 (這裡使用 task_list)
        data_to_save = {"tasks": task_list}
        try:
            # 開啟檔案進行寫入 (使用 utf-8 編碼)
            with open(file_path, "w", encoding="utf-8") as f:
                # 將 Python 字典轉換為 JSON 格式並寫入檔案
                json.dump(data_to_save, f, ensure_ascii=False, indent=4)
            # 顯示成功訊息框
            messagebox.showinfo("成功", "任務資料已成功儲存至 JSON 檔案！")
        except Exception as e:
            # 若儲存過程中發生錯誤，顯示錯誤訊息框
            messagebox.showerror("錯誤", f"儲存失敗: {e}")

# --- 更新畫布捲動區域的函數 ---
def update_scroll_region():
    """
    當 inner_frame 的內容改變後，重新計算其邊界並更新畫布的 scrollregion
    """
    global content_canvas, inner_frame
    if content_canvas and inner_frame:
        # update_idletasks() 確保所有待處理的 UI 更新完成，尺寸計算正確
        inner_frame.update_idletasks()
        # 設定畫布的 scrollregion 等於 inner_frame 的實際邊界框
        # canvas.bbox(tk.ALL) 會計算畫布上所有元件 (包含 create_window 放入的 frame) 的邊界
        content_canvas.config(scrollregion=content_canvas.bbox(tk.ALL))
        # 或者，如果確定 (0,0) 是起點，可以用:
        # content_canvas.config(scrollregion=(0, 0, inner_frame.winfo_reqwidth(), inner_frame.winfo_reqheight()))

# --- 渲染任務區塊到 inner_frame ---
def render_tasks():
    """
    根據 task_list 中的資料，在 inner_frame 中建立或更新任務區塊顯示
    """
    global inner_frame, task_list
    if not inner_frame: # 如果 inner_frame 還沒建立，就返回
        return

    # 1. 清空 inner_frame 中現有的所有任務區塊元件
    for widget in inner_frame.winfo_children():
        widget.destroy()

    # 2. 遍歷 task_list，為每個任務創建顯示元件
    for i, task_data in enumerate(task_list):
        # 建立一個帶邊框的 Frame 來代表單個任務區塊
        task_frame = tk.Frame(inner_frame, relief="solid", borderwidth=1, padx=5, pady=5)
        # 使用 pack 將任務區塊水平排列在 inner_frame 中
        # padx 和 pady 在 Frame 外部增加間距
        task_frame.pack(side="left", padx=10, pady=10, anchor="nw") # anchor='nw' 確保從左上角開始排列

        # 在任務區塊 Frame 內放置標籤顯示任務名稱 (這裡用 task_data['name'])
        task_label = tk.Label(task_frame, text=task_data['name'], width=20, anchor='w') # width 約略控制寬度, anchor='w' 文字靠左
        task_label.pack()

        # --- 您可以在這裡為 task_frame 加入更多元件，例如按鈕等 ---

    # 3. 渲染完畢後，非常重要：更新畫布的捲動區域
    update_scroll_region()

# --- 新增任務的函數 ---
def add_new_task():
    """
    簡單地新增一個帶有編號的任務到 task_list，並重新渲染
    """
    global task_list
    task_count = len(task_list) + 1
    # 新增任務資料到列表 (這裡用簡單的字典)
    new_task_data = {"name": f"任務 {task_count}"} # 您可以改成更複雜的資料結構
    task_list.append(new_task_data)

    # 重新渲染所有任務區塊到 inner_frame
    render_tasks()

# --- 主視窗設定 ---
root = tk.Tk() # 建立主視窗物件
root.title("編審神器") # 設定主視窗標題
root.geometry("1000x700") # 設定主視窗初始大小 (寬x高)
# root.resizable(False, False) # 暫時允許調整大小，方便觀察捲動效果

# --- 左側選單區域 ---
menu_frame = tk.Frame(root, relief="solid", borderwidth=1, width=100) # 設定邊框樣式和固定寬度
menu_frame.pack(side="left", fill="y") # 將選單框架放置在主視窗左側，並填滿垂直方向
menu_frame.pack_propagate(False) # 防止框架因內部元件大小而自動縮小

# --- 將新增任務功能綁定到 "任務區" 按鈕 (範例) ---
#tk.Button(menu_frame, text="任務區", font=("Arial", 10)).pack(fill="x", pady=2, padx=5)
tk.Button(menu_frame, text="新增任務", font=("Arial", 10), command=add_new_task).pack(fill="x", pady=2, padx=5) # 修改按鈕文字和命令
tk.Button(menu_frame, text="封存區", font=("Arial", 10)).pack(fill="x", pady=2, padx=5) # 同上
tk.Button(menu_frame, text="下載", font=("Arial", 10), command=save_to_json).pack(side="bottom", fill="x", pady=5, padx=5) # 靠底部，填滿水平，上下左右留白

# --- 可捲動的畫布區域 ---
canvas_frame = tk.Frame(root)
canvas_frame.pack(side="right", fill="both", expand=True)

# --- 建立畫布和捲軸 ---
# 將 content_canvas 宣告為全域變數可以在其他函數中存取
content_canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=content_canvas.yview)
h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=content_canvas.xview)
content_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

# --- 建立內部框架 (inner_frame) ---
# 將 inner_frame 宣告為全域變數可以在其他函數中存取
# 父容器是 content_canvas
inner_frame = tk.Frame(content_canvas, bg='white') # 背景設為白色以匹配畫布

# --- 將 inner_frame 嵌入畫布中 ---
# 使用 create_window 將 inner_frame 放置在畫布的 (0,0) 位置
# anchor='nw' 表示 inner_frame 的西北角 (左上角) 對齊畫布的 (0,0) 點
# tags='inner_frame' 給這個嵌入的視窗一個標籤，方便之後查找或操作 (雖然這裡沒直接用到)
content_canvas.create_window((0, 0), window=inner_frame, anchor='nw', tags='inner_frame')

# --- 排列畫布和捲軸 ---
content_canvas.grid(row=0, column=0, sticky="nsew")
v_scrollbar.grid(row=0, column=1, sticky="ns")
h_scrollbar.grid(row=1, column=0, sticky="ew")
canvas_frame.grid_rowconfigure(0, weight=1)
canvas_frame.grid_columnconfigure(0, weight=1)

# --- 初始渲染與設定捲動區域 ---
# 可以在啟動時渲染一次已有的任務 (如果有的話)
render_tasks() # 初始呼叫一次，即使 task_list 是空的，也會設定好 scrollregion

# --- (可選) 綁定 inner_frame 的大小改變事件 ---
# 當 inner_frame 的大小因為某些原因改變時 (例如視窗大小調整導致內部元件重新排列)，
# 自動呼叫 update_scroll_region 來更新捲動範圍。
# 這對於更複雜的佈局或動態內容很有用。
inner_frame.bind("<Configure>", lambda event: update_scroll_region())

# --- 啟動 Tkinter 的事件迴圈 ---
root.mainloop()