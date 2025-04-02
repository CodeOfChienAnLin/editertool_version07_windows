import tkinter as tk
from tkinter import scrolledtext, messagebox
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD # 需要安裝 tkinterdnd2
    HAS_TKINTERDND2 = True
except ImportError:
    HAS_TKINTERDND2 = False
from pathlib import Path
import os
import traceback # 用於獲取詳細錯誤信息
import tempfile
import shutil # 用於刪除目錄樹
import subprocess
import sys # 用於檢查平台

# --- ODF 相關匯入 (用於快速解析) ---
try:
    from odf import opendocument, text as odf_text, teletype
    HAS_ODF = True
except ImportError:
    HAS_ODF = False

# --- COM 相關匯入和檢查 (用於 Windows + Word 解析) ---
HAS_PYWIN32 = False
if sys.platform == 'win32': # COM 只在 Windows 上可用
    try:
        import win32com.client as win32
        import pythoncom # 需要初始化 COM
        HAS_PYWIN32 = True
    except ImportError:
        # 在 GUI 啟動前打印錯誤可能更有用
        print("警告：未找到 pywin32 模組。COM 功能需要 Windows、Microsoft Word 和 'pip install pywin32'。")
        pass # 即使沒有 pywin32，程式仍可嘗試啟動 (但功能受限)

# --- 依賴項檢查 ---
MISSING_DEPENDENCIES = []
if not HAS_TKINTERDND2:
    MISSING_DEPENDENCIES.append("tkinterdnd2 (用於拖放功能)")
if sys.platform == 'win32' and not HAS_PYWIN32:
    MISSING_DEPENDENCIES.append("pywin32 (用於透過 MS Word 解析)")
if not HAS_ODF:
     MISSING_DEPENDENCIES.append("python-odf (用於快速解析)")


# --- 快速解析函數 (使用unoconv轉ODT) ---
def parse_word_fast(filepath: Path):
    """
    嘗試使用unoconv將Word轉為ODT格式後快速解析內容。
    注意：此方法對自動編號的解析可能不準確。

    Args:
        filepath (Path): Word 文件的路徑。

    Returns:
        str | None: 解析後的文字內容，或在失敗時返回 None。
    """
    if not HAS_ODF:
        print("錯誤: 缺少 python-odf 模組，無法執行快速解析。")
        return None # 返回 None 表示失敗

    # 檢查 unoconv 是否可用
    unoconv_path = shutil.which("unoconv") # 檢查 unoconv 是否在 PATH 中
    if not unoconv_path:
        print("錯誤: 找不到 unoconv 命令。")
        messagebox.showerror("依賴錯誤", "找不到 'unoconv' 命令。\n請先安裝 LibreOffice 或 OpenOffice，並確保 'unoconv' 在系統 PATH 中。")
        return None

    tmp_dir = None # 初始化為 None
    try:
        # 建立更安全的臨時目錄
        # 修復點 1：使用 finally 確保清理
        tmp_dir = tempfile.mkdtemp(prefix="word_parser_")
        odt_path = Path(tmp_dir) / "temp.odt"

        # 使用絕對路徑轉換文件
        abs_path_str = str(filepath.resolve())
        if not filepath.exists():
            print(f"錯誤: 文件不存在: {abs_path_str}")
            return None

        # 調用 unoconv 轉換
        try:
            result = subprocess.run(
                [unoconv_path, "-f", "odt", "-o", str(odt_path), abs_path_str],
                capture_output=True,
                text=True,
                encoding='utf-8', # 嘗試指定編碼
                errors='replace', # 處理潛在的編碼錯誤
                timeout=60  # 增加超時時間
            )

            if result.returncode != 0:
                error_msg = f"unoconv 轉換失敗 (返回碼 {result.returncode}):\n{result.stderr}"
                print(error_msg)
                # 可以考慮在 GUI 中顯示更友好的錯誤
                # messagebox.showerror("轉換失敗", f"使用 unoconv 轉換 '{filepath.name}' 失敗。\n\n錯誤:\n{result.stderr[:500]}...")
                return None
        except subprocess.TimeoutExpired:
            print("unoconv 轉換超時")
            messagebox.showerror("轉換超時", f"使用 unoconv 轉換 '{filepath.name}' 超時。\n文件可能過大或 unoconv 無回應。")
            return None
        except (subprocess.SubprocessError, OSError) as e:
            error_msg = f"unoconv 執行錯誤: {str(e)}"
            print(error_msg)
            messagebox.showerror("轉換錯誤", f"執行 unoconv 時發生錯誤:\n{str(e)}")
            return None
        except Exception as e: # 捕捉其他潛在錯誤
            error_msg = f"unoconv 轉換時發生未知錯誤: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            messagebox.showerror("轉換未知錯誤", f"轉換時發生未知錯誤:\n{str(e)}")
            return None


        if not odt_path.exists():
             print(f"錯誤: unoconv 聲稱成功，但未找到輸出的 ODT 文件: {odt_path}")
             return None

        # 載入 ODT 文件並解析
        try:
            doc = opendocument.load(str(odt_path))
            content = []

            # 修復點 2：改進文本提取，但保留對列表層級的猜測（並註明其不可靠性）
            all_paras = doc.getElementsByType(odf_text.P)
            for para_elem in all_paras:
                # 提取段落文本
                para_text = teletype.extractText(para_elem).strip()

                # --- 嘗試獲取列表信息 (基於樣式屬性 - 可能不準確) ---
                # ODF 列表通常通過 text:list-style-id 屬性應用於 <text:p>
                # 或通過 <text:p> 的父級 <text:list-item> 的層級來判斷
                # 這裡簡化處理：嘗試檢查 style:list-level 屬性（如果存在）
                list_level = None
                try:
                    # 檢查段落自身的列表層級屬性（不常見，但可能存在）
                    list_level_str = para_elem.getAttribute("listlevel", "text") # odfpy 使用 (ns, name)
                    if list_level_str:
                         list_level = int(list_level_str) - 1 # ODF level 通常從 1 開始
                    else:
                        # 更常見的是檢查父級 list-item
                        parent = para_elem.parentNode
                        # 檢查是否為 ListItem 且有 level
                        if hasattr(parent, 'qname') and parent.qname == (odf_text.ListItem.qname[0], odf_text.ListItem.qname[1]):
                           # odfpy 中 ListItem 沒有直接的 level 屬性，層級依賴 list 結構
                           # 這裡無法簡單獲取，保持原樣的猜測邏輯作為備用
                           pass # 無法簡單獲取父級 ListItem 的層級

                except (AttributeError, ValueError, TypeError):
                    pass # 忽略獲取層級時的錯誤

                # --- 格式化輸出 (基於猜測的層級) ---
                # **警告**: 這種基於 level 屬性的編號猜測非常不可靠
                if list_level is not None and list_level >= 0:
                    indent = "  " * list_level # 使用空格模擬縮進
                    # 簡單的編號樣式猜測
                    if list_level == 0: num_text = "1." # 假設第一級是數字
                    elif list_level == 1: num_text = "a." # 假設第二級是字母
                    elif list_level == 2: num_text = "i." # 假設第三級是羅馬數字
                    else: num_text = "-"         # 其他層級用 -
                    content.append(f"{indent}{num_text} {para_text}")
                elif para_text: # 只添加非空文本段落
                    content.append(para_text)

            return "\n".join(content)

        except Exception as e:
            error_msg = f"ODT 解析失敗: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            # 不在此處彈出消息框，讓主調用函數決定
            return None # 返回 None 表示失敗

    except Exception as e: # 捕捉創建臨時目錄等早期錯誤
        error_msg = f"快速解析準備階段出錯: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return None # 返回 None 表示失敗
    finally:
        # 修復點 1：確保無論成功或失敗都清理臨時目錄
        if tmp_dir and Path(tmp_dir).exists():
            try:
                shutil.rmtree(tmp_dir) # 使用 shutil.rmtree 更安全
                # print(f"清理臨時目錄: {tmp_dir}")
            except OSError as e:
                print(f"警告：無法刪除臨時目錄 {tmp_dir}: {e}")

# --- Word 文件解析邏輯 (使用 Windows COM) ---
def parse_word_document_com(filepath: Path):
    """
    使用 Windows COM 與 Microsoft Word 互動來解析 .docx 文件，
    以嘗試獲取包括自動編號在內的渲染後文字。

    Args:
        filepath (Path): Word 文件的路徑。

    Returns:
        str | None: 解析後的文字內容，包含自動編號和縮排。
                      如果發生錯誤、缺少依賴或無法執行則返回 None。
    """
    if not sys.platform == 'win32':
        print("錯誤：COM 功能僅在 Windows 上受支持。")
        return None
    if not HAS_PYWIN32:
        # 這個訊息應該在 GUI 初始化時已經提示過，這裡可以只記錄日誌
        print("錯誤：缺少 pywin32 模組，無法使用 COM 功能。")
        return None

    if not filepath.is_file():
        print(f"錯誤：找不到檔案 {filepath}")
        return None

    word_app = None
    doc = None
    com_initialized = False
    try:
        # 初始化 COM 環境 (對於某些線程模型是必要的)
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception as e:
            print(f"警告：初始化 COM 失敗: {e} (嘗試繼續)")

        parsed_content = []
        # 啟動 Word 應用程式 (嘗試後台執行)
        try:
            word_app = win32.Dispatch("Word.Application")
        except pythoncom.com_error as ce:
             hr, msg, exc, arg = ce.args
             print(f"COM 錯誤 (Dispatch): HRESULT={hr}, Message={msg}")
             # 特定的 HRESULT 可能表示 Word 未安裝或註冊問題
             if hr == -2147221005: # CO_E_CLASSSTRING
                 messagebox.showerror("COM 錯誤", "無法啟動 Microsoft Word。\n請確認已正確安裝 Word。")
             else:
                 messagebox.showerror("COM 錯誤", f"啟動 Word 時發生 COM 錯誤:\n{msg}")
             return None # 無法啟動 Word，直接返回

        word_app.Visible = False # 不顯示 Word 視窗
        try:
            # 設置 DisplayAlerts 為 False 可能有助於防止 Word 彈出對話框
            word_app.DisplayAlerts = 0 # wdAlertsNone = 0
        except Exception as e:
            print(f"警告：無法設置 DisplayAlerts 屬性: {e}")

        # 打開文件 (需要絕對路徑字串)
        try:
             doc = word_app.Documents.Open(str(filepath.resolve()), ReadOnly=True)
        except pythoncom.com_error as ce:
            hr, msg, exc, arg = ce.args
            print(f"COM 錯誤 (Open): HRESULT={hr}, Message={msg}")
            # 可能的錯誤：文件找不到、權限不足、文件損壞、Word 打不開此類型文件等
            messagebox.showerror("文件開啟錯誤", f"無法透過 Word 開啟檔案 '{filepath.name}':\n{msg}\n\n請檢查文件是否存在、未損壞且 Word 可以開啟它。")
            # 確保在返回前嘗試關閉已部分啟動的 Word
            if word_app:
                try: word_app.Quit()
                except: pass
            return None

        # --- 處理縮排 ---
        POINTS_PER_INDENT_LEVEL = 18 # 每 18 磅 (Point) 算一級縮排 (可調整)
        SPACES_PER_INDENT_LEVEL = 3  # 每級縮排對應的空格數 (可調整)

        # 迭代文件中的段落
        # 使用 try/except 包裹迭代過程，防止單一段落錯誤導致整個解析失敗
        try:
            for i, para_com in enumerate(doc.Paragraphs):
                indent_space = "" # 預設無縮進
                formatted_line = "[讀取段落時發生錯誤]" # 預設錯誤訊息
                try:
                    para_range = para_com.Range
                    # 獲取 Word 渲染的列表字串 (編號或項目符號)
                    list_string = para_range.ListFormat.ListString
                    # 獲取段落的完整文字 (通常包含列表字串和一個結尾的 \r)
                    full_text = para_range.Text
                    # 清理文字，移除結尾的換行符 \r 或 \r\n
                    actual_text = full_text.rstrip('\r\n')

                    # --- 計算縮排 ---
                    indent_points = 0.0
                    try:
                        # LeftIndent 單位是磅 (Points)
                        indent_points = para_com.Format.LeftIndent
                    except AttributeError:
                        pass # 某些特殊段落可能沒有這個屬性
                    except pythoncom.com_error as ce:
                         # 有時訪問特定段落格式會觸發 COM 錯誤
                         print(f"警告：獲取段落 {i+1} 縮排時 COM 錯誤: {ce}")
                    except Exception as indent_err:
                        print(f"警告：獲取段落 {i+1} 縮排時出錯: {indent_err}")

                    # 根據磅值計算縮排層級和對應空格
                    if indent_points > 0:
                        indent_level = int(indent_points / POINTS_PER_INDENT_LEVEL)
                        if indent_level < 0: indent_level = 0 # 防止負數
                        indent_space = " " * (indent_level * SPACES_PER_INDENT_LEVEL)

                    # --- 組合輸出 ---
                    separator = "\t" # 使用 Tab 分隔編號和文字

                    if list_string:
                        # 嘗試更可靠地移除文本開頭的列表字符串和分隔符
                        temp_text = actual_text
                        # 確保 list_string 不是空的才進行比較
                        if list_string and temp_text.startswith(list_string):
                            temp_text = temp_text[len(list_string):]
                            # 移除緊隨其後的可能的分隔符（空格或製表符）
                            temp_text = temp_text.lstrip(' \t')

                        formatted_line = f"{indent_space}{list_string}{separator}{temp_text}"
                    else:
                        # 沒有列表字串，就是普通段落
                        formatted_line = f"{indent_space}{actual_text}"

                    parsed_content.append(formatted_line)

                except pythoncom.com_error as para_ce:
                     print(f"警告：讀取段落 {i+1} 時發生 COM 錯誤: {para_ce}")
                     parsed_content.append(f"{indent_space}[讀取段落 COM 錯誤]")
                except Exception as para_exc:
                    print(f"警告：讀取段落 {i+1} 時發生未知錯誤: {para_exc}")
                    # 嘗試獲取原始文本作為回退
                    try:
                        raw_text = para_com.Range.Text.rstrip('\r\n')
                        parsed_content.append(f"{indent_space}[讀取錯誤] {raw_text}")
                    except:
                        parsed_content.append(f"{indent_space}[讀取錯誤且無法獲取原始文本]")

        except Exception as iter_exc:
            print(f"錯誤：迭代段落時發生嚴重錯誤: {iter_exc}\n{traceback.format_exc()}")
            messagebox.showerror("解析錯誤", f"處理文件 '{filepath.name}' 段落時發生錯誤:\n{iter_exc}\n\n可能只能顯示部分內容。")
            # 即使迭代出錯，也嘗試返回已解析的部分內容
            pass

        # 添加提示：不顯示圖片
        parsed_content.append("\n\n--- (注意：文件中的圖片無法在此顯示) ---")

        return "\n".join(parsed_content)

    except pythoncom.com_error as ce:
        # 捕捉在 Dispatch 或 Open 之外發生的 COM 錯誤
        hr, msg, exc, arg = ce.args
        print(f"處理 Word 文件時發生 COM 錯誤: HRESULT={hr}, Message={msg}\n{traceback.format_exc()}")
        messagebox.showerror("COM 交互錯誤", f"與 Word 交互時發生 COM 錯誤：\n{msg}\n(請確認 Word 可正常運作)")
        return None
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"解析 Word 文件時發生未知錯誤: {e}\n{error_details}")
        messagebox.showerror("未知解析錯誤", f"解析 Word 文件時發生未知錯誤：\n{e}\n(詳細資訊請查看控制台輸出)")
        return None
    finally:
        # --- 無論成功或失敗，都要確保關閉 Word 和清理 COM ---
        # print("正在清理 Word COM...")
        try:
            if doc:
                # print("正在關閉 Word 文件...")
                doc.Close(SaveChanges=0) # 關閉文件，不保存更改 (wdDoNotSaveChanges = 0)
                doc = None # 釋放引用
        except Exception as e_close:
            print(f"關閉 Word 文件時發生錯誤: {e_close}")
        try:
            if word_app:
                # print("正在退出 Word 應用程式...")
                word_app.Quit()
                word_app = None # 釋放引用
        except Exception as e_quit:
            print(f"退出 Word 應用程式時發生錯誤: {e_quit}")

        # 清理 COM 環境
        if com_initialized:
            try:
                # print("正在反初始化 COM...")
                pythoncom.CoUninitialize()
            except Exception as e_uninit:
                print(f"警告：清理 COM 環境失敗: {e_uninit}")


# [Keep all previous imports and functions: parse_word_fast, parse_word_document_com, etc.]
# ... (rest of the imports and functions remain the same) ...

# --- GUI 應用程式 ---
class WordViewerApp:
    def __init__(self, master):
        self.master = master
        self.current_filepath = None # 追蹤當前顯示的文件

        # --- 檢查並報告缺失的依賴 ---
        if MISSING_DEPENDENCIES:
            missing_str = "\n- ".join(MISSING_DEPENDENCIES)
            warning_title = "依賴警告"
            warning_message = f"啟動警告：檢測到缺少以下依賴項，部分功能可能無法使用：\n\n- {missing_str}\n\n請根據需要安裝它們 (例如使用 pip)。"
            print(warning_message) # 同時打印到控制台
            messagebox.showwarning(warning_title, warning_message)

        # --- 設置窗口標題 ---
        title = "Word 文件檢視器"
        # Determine primary method indication based on priority
        primary_method = ""
        if HAS_ODF: # Fast method is prioritized
            primary_method = "快速解析優先"
            # Check if unoconv is likely available (basic check)
            if not shutil.which("unoconv"):
                 primary_method += " (unoconv未找到!)"
        elif sys.platform == 'win32' and HAS_PYWIN32: # Fallback is COM
            primary_method = "COM 模式"
        else:
             primary_method = "功能受限"

        if HAS_TKINTERDND2:
             title += f" ({primary_method}) [支援拖放]"
        else:
             title += f" ({primary_method})"

        master.title(title)
        master.geometry("800x600")

        # --- 建立可滾動的文字區域 ---
        self.text_area = scrolledtext.ScrolledText(
            master,
            wrap=tk.WORD,
            font=("Arial", 11)
            # font=("Consolas", 11) # Consider monospace for better indent alignment
        )
        self.text_area.pack(expand=True, fill='both', padx=10, pady=10)

        # --- 設定初始提示文字並禁用編輯 ---
        self.show_initial_message()

        # --- 設定拖放目標 ---
        if HAS_TKINTERDND2:
            try:
                self.text_area.drop_target_register(DND_FILES)
                self.text_area.dnd_bind('<<Drop>>', self.handle_drop)
            except tk.TclError as e:
                 messagebox.showerror("拖放功能錯誤", f"無法註冊拖放目標:\n{e}\n\n請確認 tkinterdnd2 已正確安裝。")
                 self.text_area.insert(tk.END, "\n\n錯誤：拖放功能初始化失敗。")
                 self.text_area.configure(state='disabled')
        else:
            # Check done in show_initial_message now
            pass # Message is handled in show_initial_message

    def show_initial_message(self):
        """顯示初始的歡迎和說明信息"""
        self.text_area.configure(state='normal')
        self.text_area.delete('1.0', tk.END)
        if HAS_TKINTERDND2:
             self.text_area.insert(tk.END, "請將 Word (.docx) 檔案拖曳到這裡...\n\n")
        else:
             self.text_area.insert(tk.END, "提示：未安裝 tkinterdnd2，拖放功能不可用。\n請執行 'pip install tkinterdnd2'。\n\n")

        self.text_area.insert(tk.END, "** 使用說明 (優先使用快速解析) **\n")

        # Describe Fast Method First (if available)
        if HAS_ODF:
             self.text_area.insert(tk.END, "\n- [主要] 快速解析方法 (需 LibreOffice/OpenOffice + unoconv + python-odf):\n")
             self.text_area.insert(tk.END, "  - 透過轉換為 ODT 格式解析，速度較快。\n")
             self.text_area.insert(tk.END, "  - **注意：自動編號的顯示可能不準確。**\n")
             if not shutil.which("unoconv"):
                  self.text_area.insert(tk.END, "  - **警告：未在 PATH 中找到 'unoconv' 命令！此方法將失敗。**\n")
        else:
             self.text_area.insert(tk.END, "\n- 快速解析方法不可用 (缺少 python-odf)。\n")

        # Describe COM Method Second (if available)
        if sys.platform == 'win32' and HAS_PYWIN32:
            self.text_area.insert(tk.END, "\n- [備選] COM 方法 (需 Word 和 pywin32):\n")
            self.text_area.insert(tk.END, "  - 若快速解析失敗，將嘗試透過 Word 本身讀取。\n")
            self.text_area.insert(tk.END, "  - 能較好顯示自動編號和縮進。\n")
            self.text_area.insert(tk.END, "  - 處理速度相對較慢。\n")
        elif sys.platform != 'win32':
             self.text_area.insert(tk.END, "\n- COM 方法在此平台 (非 Windows) 不可用。\n")
        else: # Windows 但沒有 pywin32
             self.text_area.insert(tk.END, "\n- COM 方法不可用 (缺少 pywin32)。\n")


        self.text_area.insert(tk.END, "\n- 無法顯示文件中的圖片。\n")
        if MISSING_DEPENDENCIES:
             self.text_area.insert(tk.END, "\n**警告：缺少部分依賴，功能受限。**\n")

        self.text_area.configure(state='disabled')
        # Reset title after showing message, in case it was showing processing state
        current_title = self.master.title()
        if "正在處理" in current_title or "解析失敗" in current_title or "Word 文件檢視器 -" in current_title:
             base_title = current_title.split(" - ")[0] # Get the base part like "Word 文件檢視器 (快速解析優先) [支援拖放]"
             self.master.title(base_title)


    def handle_drop(self, event):
        """處理拖放事件"""
        filepath_str = event.data
        filepath = None

        try:
            possible_paths = self.master.tk.splitlist(filepath_str)
            if possible_paths:
                first_path_str = possible_paths[0]
                candidate = Path(first_path_str)
                if candidate.exists() and candidate.is_file():
                    filepath = candidate
                else:
                    cleaned_p = first_path_str.strip('\'"')
                    candidate = Path(cleaned_p)
                    if candidate.exists() and candidate.is_file():
                         filepath = candidate
        except Exception as e:
            print(f"解析拖放路徑時出錯: {e}\n原始數據: {filepath_str}")
            messagebox.showerror("錯誤", f"無法解析拖放的檔案路徑:\n{filepath_str}")
            return

        if filepath is None:
            messagebox.showerror("錯誤", f"未找到有效的檔案路徑或文件不存在:\n{filepath_str}")
            return

        if filepath.suffix.lower() == '.docx':
            self.current_filepath = filepath
            self.display_word_content(filepath)
        else:
            messagebox.showwarning("不支援的檔案", f"目前只支援 .docx 檔案。\n您拖放的是：{filepath.name}")

    # ***** MODIFIED METHOD *****
    def display_word_content(self, filepath: Path):
        """
        解析並顯示 Word 文件內容。
        優先嘗試快速解析 (unoconv+ODF)，失敗則回退到 COM 方法 (若可用)。
        """
        self.text_area.configure(state='normal')
        self.text_area.delete('1.0', tk.END)
        self.master.title(f"Word 文件檢視器 - 正在處理 {filepath.name}...")
        self.master.update_idletasks() # Ensure title update is visible

        parsed_text = None
        method_used = "None"

        # --- 策略：優先嘗試 Fast，其次嘗試 COM ---

        # 1. 嘗試快速解析 (如果有 ODF 依賴)
        if HAS_ODF:
            self.text_area.insert(tk.END, f"--- 正在嘗試快速解析 (unoconv): {filepath.name} ---\n"
                                          f"--- (注意：此方法的自動編號可能不準確) ---\n\n")
            self.master.update_idletasks()
            parsed_text = parse_word_fast(filepath) # Returns None on failure
            if parsed_text is not None:
                method_used = "Fast (unoconv+ODF)"
                # Optionally add the warning again at the end, or keep it in the status message
                # parsed_text += "\n\n--- (注意：快速解析方法的自動編號可能不準確) ---"
            else:
                 self.text_area.insert(tk.END, "\n--- 快速解析失敗 ---\n")
                 # Don't return yet, fall through to try COM if possible

        else:
             # If ODF is not available, mention it in the status area
             self.text_area.insert(tk.END, f"--- 快速解析不可用 (缺少 python-odf 模組) ---\n")


        # 2. 如果快速解析失敗 或 不可用，嘗試 COM (僅 Windows 且有 pywin32)
        if parsed_text is None: # Only try COM if fast method failed or wasn't available
            if sys.platform == 'win32' and HAS_PYWIN32:
                self.text_area.insert(tk.END, f"\n--- 正在嘗試透過 Word (COM) 載入: {filepath.name} (可能需要幾秒鐘) ---\n\n")
                self.master.update_idletasks()
                com_result = parse_word_document_com(filepath) # Returns None on failure
                if com_result is not None:
                    parsed_text = com_result
                    method_used = "COM (MS Word)"
                else:
                    # COM 也失敗了
                    self.text_area.insert(tk.END, "\n--- COM 方法也失敗了 ---\n")
            elif HAS_ODF: # Only show this if fast parse was attempted but failed, and COM is unavailable
                 # If fast parse failed, and we are not on Windows or don't have pywin32
                 self.text_area.insert(tk.END, f"\n--- COM 方法不可用 (非 Windows 或缺少 pywin32) ---\n")


        # --- 顯示結果或錯誤信息 ---
        self.text_area.delete('1.0', tk.END) # Clear status messages and old content
        if parsed_text is not None:
            # Success: Display content and method used
            final_title = f"Word 文件檢視器 - {filepath.name} ({method_used})"
            self.master.title(final_title)
            self.text_area.insert(tk.END, f"--- 顯示檔案: {filepath.name} (使用 {method_used} 方法) ---\n")
            if method_used == "Fast (unoconv+ODF)":
                 self.text_area.insert(tk.END, "--- (注意：此方法的自動編號可能不準確) ---\n\n")
            else:
                 self.text_area.insert(tk.END, "\n") # Just a newline for COM method
            self.text_area.insert(tk.END, parsed_text)

        else:
            # Failure: Both methods failed or were unavailable
            final_title = f"Word 文件檢視器 - 解析失敗: {filepath.name}"
            self.master.title(final_title)
            self.text_area.insert(tk.END, f"--- 無法解析檔案: {filepath.name} ---\n\n")
            self.text_area.insert(tk.END, "所有可用的解析方法均失敗。\n")

            # Provide specific reasons based on availability checks
            if not HAS_ODF and not (sys.platform == 'win32' and HAS_PYWIN32):
                 self.text_area.insert(tk.END, "錯誤：未找到任何可用的解析方法 (缺少 python-odf 和 pywin32)。\n")
            elif not HAS_ODF:
                 self.text_area.insert(tk.END, "錯誤：快速解析不可用 (缺少 python-odf)，且 COM 方法失敗或不可用。\n")
            elif not (sys.platform == 'win32' and HAS_PYWIN32):
                 self.text_area.insert(tk.END, "錯誤：快速解析失敗，且 COM 方法不可用 (非 Windows 或缺少 pywin32)。\n")
            else: # Both methods available but failed
                 self.text_area.insert(tk.END, "快速解析和 COM 方法均嘗試失敗。\n")


            self.text_area.insert(tk.END, "\n請檢查控制台輸出獲取詳細錯誤信息。\n")
            self.text_area.insert(tk.END, f"可能的原因：\n")
            if HAS_ODF:
                 self.text_area.insert(tk.END, f"- unoconv (LibreOffice/OpenOffice) 未安裝、未在PATH中或無法正常工作。\n")
            if sys.platform == 'win32':
                 self.text_area.insert(tk.END, f"- Microsoft Word 未安裝、未運行、權限不足或無法開啟該文件。\n")
            self.text_area.insert(tk.END, f"- 文件已損壞或格式不受支持。\n")


        self.text_area.configure(state='disabled') # Disable editing after completion
        self.text_area.yview_moveto(0.0) # Scroll to top


# --- 主程式入口 ---
if __name__ == "__main__":
    root = None # 初始化為 None
    try:
        # --- TkinterDnD 需要特殊的主窗口 ---
        if HAS_TKINTERDND2:
            # 使用 TkinterDnD.Tk() 初始化以支持拖放
            root = TkinterDnD.Tk()
        else:
            # 如果沒有 tkinterdnd2，使用標準 Tk
            print("提示: 未加載 tkinterdnd2，拖放功能將不可用。")
            root = tk.Tk()

        app = WordViewerApp(root)
        root.mainloop()

    except tk.TclError as e:
         # 捕捉 Tkinter 相關的 Tcl 錯誤，例如與 TkinterDnD 初始化相關的問題
         print(f"啟動應用程式時發生 Tcl 錯誤: {e}\n{traceback.format_exc()}")
         try:
             # 嘗試用標準 Tk 顯示錯誤彈窗
             error_root = tk.Tk()
             error_root.withdraw() # 隱藏主視窗
             messagebox.showerror("啟動錯誤", f"無法啟動應用程式 (Tcl 錯誤):\n{e}\n\n請檢查 Tkinter 和 tkinterdnd2 是否安裝正確且兼容。")
             error_root.destroy()
         except Exception:
             pass # 如果連 Tk 都無法建立，只能在終端顯示了
    except Exception as e:
        # 捕捉其他可能的啟動錯誤
        print(f"啟動應用程式時發生未知錯誤: {e}\n{traceback.format_exc()}")
        try:
            error_root = tk.Tk()
            error_root.withdraw()
            messagebox.showerror("啟動錯誤", f"無法啟動應用程式:\n{e}\n\n請檢查 Python 環境和所有依賴項是否已正確安裝。")
            error_root.destroy()
        except Exception:
            pass