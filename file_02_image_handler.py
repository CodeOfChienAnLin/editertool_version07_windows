"""
圖片處理相關功能模組
"""
import os
import traceback
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
from docx import Document
import io

def extract_images_from_docx(self, file_path):
    """從Word文件中提取圖片

    參數:
        file_path: Word檔案路徑
    """
    try:
        # 清空現有圖片
        clear_images(self)
        
        # 打開Word文檔
        doc = Document(file_path)
        
        # 計數器，用於跟踪圖片索引
        image_index = 0
        
        # 遍歷文檔中的所有關係
        for rel in doc.part.rels.values():
            # 檢查關係類型是否為圖片
            if "image" in rel.reltype:
                # 獲取圖片數據
                image_data = rel.target_part.blob
                
                try:
                    # 使用PIL打開圖片
                    image = Image.open(io.BytesIO(image_data))
                    
                    # 存儲原始圖片
                    self.images.append(image)
                    
                    # 顯示圖片
                    display_image(self, image, image_index)
                    
                    # 增加索引
                    image_index += 1
                except Exception as img_error:
                    print(f"無法處理圖片: {str(img_error)}")
        
        # 更新狀態欄
        self.status_bar.config(text=f"已從文件中提取 {image_index} 張圖片")
        
    except Exception as e:
        error_msg = f"提取圖片時發生錯誤: {str(e)}"
        messagebox.showerror("錯誤", error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Image Extraction Error", error_msg, traceback.format_exc())

def display_image(self, image, index):
    """在圖片區域顯示圖片

    參數:
        image: PIL Image 對象
        index: 圖片索引
    """
    try:
        # 創建圖片框架
        img_frame = tk.Frame(self.image_container)
        img_frame.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 縮放圖片以適應顯示區域
        max_height = 100  # 最大高度
        
        # 計算縮放比例
        ratio = max_height / image.height
        new_width = int(image.width * ratio)
        new_height = max_height
        
        # 縮放圖片
        resized_image = image.resize((new_width, new_height), Image.LANCZOS)
        
        # 轉換為Tkinter可用的格式
        tk_image = ImageTk.PhotoImage(resized_image)
        
        # 存儲引用，防止被垃圾回收
        self.image_refs.append(tk_image)
        
        # 創建標籤顯示圖片
        img_label = tk.Label(img_frame, image=tk_image)
        img_label.pack()
        
        # 添加點擊事件，顯示原始大小圖片
        img_label.bind("<Button-1>", lambda event, img=image, idx=index: show_full_image(self, img, idx))
        
        # 添加圖片索引標籤
        tk.Label(img_frame, text=f"圖片 {index + 1}").pack()
        
    except Exception as e:
        print(f"顯示圖片時發生錯誤: {str(e)}")

def show_full_image(self, image, index):
    """顯示原始大小的圖片

    參數:
        image: PIL Image 對象
        index: 圖片索引
    """
    try:
        # 創建新窗口
        img_window = tk.Toplevel(self.root)
        img_window.title(f"圖片 {index + 1}")
        
        # 設置窗口大小
        window_width = min(image.width, 800)
        window_height = min(image.height, 600)
        img_window.geometry(f"{window_width}x{window_height}")
        
        # 創建畫布和滾動條
        canvas = tk.Canvas(img_window, width=window_width, height=window_height)
        
        # 水平滾動條
        h_scrollbar = tk.Scrollbar(img_window, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 垂直滾動條
        v_scrollbar = tk.Scrollbar(img_window, orient=tk.VERTICAL, command=canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 配置畫布
        canvas.config(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 設置滾動區域
        canvas.config(scrollregion=(0, 0, image.width, image.height))
        
        # 轉換為Tkinter可用的格式
        tk_image = ImageTk.PhotoImage(image)
        
        # 在畫布上顯示圖片
        canvas.create_image(0, 0, anchor=tk.NW, image=tk_image)
        
        # 存儲引用，防止被垃圾回收
        img_window.tk_image = tk_image
        
        # 添加關閉按鈕
        close_button = tk.Button(img_window, text="關閉", command=img_window.destroy)
        close_button.pack(side=tk.BOTTOM, pady=5)
        
    except Exception as e:
        error_msg = f"顯示完整圖片時發生錯誤: {str(e)}"
        messagebox.showerror("錯誤", error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Image Display Error", error_msg, traceback.format_exc())

def clear_images(self):
    """清空圖片區域"""
    # 清空圖片列表
    self.images = []
    
    # 清空圖片引用
    self.image_refs = []
    
    # 清空圖片容器
    for widget in self.image_container.winfo_children():
        widget.destroy()

def download_images(self):
    """下載所有圖片到指定路徑"""
    if not self.images:
        messagebox.showinfo("提示", "沒有可下載的圖片")
        return
    
    try:
        # 確保下載路徑存在
        os.makedirs(self.download_path, exist_ok=True)
        
        # 下載所有圖片
        for i, image in enumerate(self.images):
            # 生成檔案名稱
            file_name = f"image_{i + 1}.png"
            file_path = os.path.join(self.download_path, file_name)
            
            # 儲存圖片
            image.save(file_path)
        
        # 更新狀態欄
        self.status_bar.config(text=f"已下載 {len(self.images)} 張圖片到 {self.download_path}")
        
        # 顯示成功訊息
        messagebox.showinfo("成功", f"已下載 {len(self.images)} 張圖片到:\n{self.download_path}")
        
    except Exception as e:
        error_msg = f"下載圖片時發生錯誤: {str(e)}"
        messagebox.showerror("錯誤", error_msg)
        
        # 記錄錯誤
        from utils_01_error_handler import log_error
        log_error(self, "Image Download Error", error_msg, traceback.format_exc())

def choose_download_path(self):
    """選擇圖片下載路徑"""
    path = filedialog.askdirectory(
        title="選擇圖片下載路徑",
        initialdir=self.download_path
    )
    
    if path:
        self.download_path = path
        self.status_bar.config(text=f"已設定下載路徑: {path}")
