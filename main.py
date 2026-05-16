import sys
import os
import datetime

# 【終極防護】強制將 PyInstaller 的暫存目錄與執行檔目錄加入系統路徑，解決找不到模組的問題
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.abspath("."))

import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import customtkinter as ctk

from pdf_tools import (process_merge_pdfs, process_protect_pdf, process_split_pdf, process_pdf_to_images, 
                       process_compress_pdf, process_remove_watermark, process_pdf_to_ppt, 
                       process_images_to_pdf, process_unlock_pdf, process_rotate_pdf, process_add_watermark,
                       process_remove_pages, process_to_grayscale, process_extract_text, process_insert_blank_page,
                       process_add_page_numbers, process_reorder_pages, process_extract_original_images,
                       process_flatten_pdf, process_add_image_watermark,
                       process_image_ocr, process_image_remove_text,
                       process_pdf_to_word, process_pdf_to_excel, process_redact_text, process_crop_pdf)

from utils import check_poppler_exists, open_file_or_folder, get_base_path

def check_single_instance():
    mutex_name = "Global\\PDF_TOOL_MUTEX"
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, mutex_name)
    if kernel32.GetLastError() == 183:
        messagebox.showwarning("提示", "程式已經在執行中，請勿重複開啟。")
        sys.exit(0)
    return mutex

def set_dpi_awareness():
    try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except: pass

def check_is_encrypted(file_path):
    try:
        from pypdf import PdfReader
        return PdfReader(file_path).is_encrypted
    except: return False

class CTkinterDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class ListManagerWindow(ctk.CTkToplevel):
    def __init__(self, app, initial_files, mode, start_callback):
        super().__init__(app.root)
        self.title("檔案合併管理器")
        self.geometry("600x350")
        self.start_callback = start_callback
        self.mode = mode
        self.app = app
        self.transient(app.root)
        self.app.set_ui_state("disabled")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        list_frame = ctk.CTkFrame(self, fg_color="transparent")
        list_frame.pack(side="left", fill="both", expand=True, padx=(15, 5), pady=15)
        
        ctk.CTkLabel(list_frame, text="合併檔案列表 (支援直接拖曳檔案進來)：", font=("Microsoft JhengHei", 14, "bold")).pack(anchor="w", pady=(0, 5))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Microsoft JhengHei", 11), yscrollcommand=scrollbar.set, activestyle="none")
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        for f in initial_files: self.listbox.insert(tk.END, f)
            
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(side="right", fill="y", padx=(5, 15), pady=15)
        
        file_types = [("支援的檔案", "*.pdf;*.jpg;*.jpeg;*.png;*.docx;*.doc")] if mode == "MERGE" else [("Images", "*.jpg;*.jpeg;*.png")]
            
        ctk.CTkButton(btn_frame, text="➕ 新增檔案", command=lambda: [self.listbox.insert(tk.END, f) for f in filedialog.askopenfilenames(filetypes=file_types)]).pack(pady=5)
        ctk.CTkButton(btn_frame, text="⬆️ 上移", command=self.move_up).pack(pady=5)
        ctk.CTkButton(btn_frame, text="⬇️ 下移", command=self.move_down).pack(pady=5)
        ctk.CTkButton(btn_frame, text="❌ 移除", command=self.remove_item, fg_color="#cc3333", hover_color="#aa2222").pack(pady=5)
        ctk.CTkButton(btn_frame, text="🚀 開始處理", command=self.start_merge, fg_color="#28a745", hover_color="#218838", height=35).pack(side="bottom", pady=15)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        valid_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.docx', '.doc') if self.mode == "MERGE" else ('.jpg', '.jpeg', '.png')
        for f in files:
            if f.lower().endswith(valid_exts): self.listbox.insert(tk.END, f)

    def on_close(self):
        self.app.set_ui_state("normal")
        self.destroy()
    def move_up(self):
        try:
            idx = self.listbox.curselection()[0]
            if idx > 0: val = self.listbox.get(idx); self.listbox.delete(idx); self.listbox.insert(idx-1, val); self.listbox.select_set(idx-1)
        except: pass
    def move_down(self):
        try:
            idx = self.listbox.curselection()[0]
            if idx < self.listbox.size()-1: val = self.listbox.get(idx); self.listbox.delete(idx); self.listbox.insert(idx+1, val); self.listbox.select_set(idx+1)
        except: pass
    def remove_item(self):
        try:
            idx = self.listbox.curselection()[0]; self.listbox.delete(idx)
            if self.listbox.size() > 0: self.listbox.select_set(min(idx, self.listbox.size()-1))
        except: pass
    def start_merge(self):
        final_list = list(self.listbox.get(0, tk.END))
        if len(final_list) < 2 and self.mode == "MERGE": return messagebox.showwarning("警告", "需要至少兩個檔案！")
        self.start_callback(final_list)
        self.destroy()

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 辦公室全能工具箱 PRO")
        self.root.geometry("850x650") 
        ctk.set_appearance_mode("light")  
        ctk.set_default_color_theme("blue")  

        icon_path = os.path.join(get_base_path(), "icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        if not check_poppler_exists():
            messagebox.showwarning("元件缺失", "找不到 Poppler，部分轉檔功能可能受限。")

        self.mode_var = ctk.StringVar(value="PDF2WORD")
        self.stop_event = threading.Event() 

        # 頂部：標題與主題切換
        top_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        top_frame.pack(fill="x", padx=15, pady=(10, 0))
        ctk.CTkLabel(top_frame, text="PDF 專業工具箱", font=("Microsoft JhengHei", 20, "bold")).pack(side="left")
        self.theme_switch = ctk.CTkSwitch(top_frame, text="深色模式", command=self.toggle_theme, font=("Microsoft JhengHei", 12))
        self.theme_switch.pack(side="right")

        # 核心：標籤頁設計 (TabView)
        self.tabview = ctk.CTkTabview(self.root, height=140)
        self.tabview.pack(fill="x", padx=15, pady=5)
        
        tab1 = self.tabview.add("📂 格式轉換")
        tab2 = self.tabview.add("✂️ 頁面編輯")
        tab3 = self.tabview.add("🛡️ 處理與安全")
        tab4 = self.tabview.add("🤖 AI與進階工具")

        # 定義各分頁功能
        modes_t1 = [("PDF 轉 Word", "PDF2WORD"), ("PDF 轉 Excel", "PDF2EXCEL"), ("PDF 轉 PPT", "PPT"), ("PDF 轉圖片", "PDF2IMG"), ("圖片轉 PDF", "IMG2PDF")]
        modes_t2 = [("提取/分割 PDF", "SPLIT"), ("刪除指定頁", "REMOVE_PAGES"), ("插入空白頁", "INSERT_BLANK"), ("重新排序頁", "REORDER"), ("萬能合併", "MERGE"), ("裁切白邊", "CROP")]
        modes_t3 = [("轉黑白/灰階", "GRAYSCALE"), ("扁平化(防篡改)", "FLATTEN"), ("PDF 壓縮", "COMPRESS"), ("PDF 旋轉", "ROTATE"), ("加密 PDF", "PROTECT"), ("解鎖 PDF", "UNLOCK")]
        modes_t4 = [("提取文字/OCR", "EXTRACT_TXT"), ("提取內嵌圖", "EXTRACT_IMGS"), ("加文字浮水印", "ADD_WM"), ("印章/圖片浮水印", "IMG_WM"), ("機密文字塗黑", "REDACT"), ("浮水印/文字抹除", "RMWATERMARK"), ("添加頁碼", "ADD_PAGE_NUM")]

        def create_radio_buttons(parent, modes_list):
            for i, (text, val) in enumerate(modes_list):
                row, col = divmod(i, 4)
                ctk.CTkRadioButton(parent, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13)).grid(row=row, column=col, padx=15, pady=10, sticky="w")

        create_radio_buttons(tab1, modes_t1)
        create_radio_buttons(tab2, modes_t2)
        create_radio_buttons(tab3, modes_t3)
        create_radio_buttons(tab4, modes_t4)

        # 動態選項區
        self.opt_frame = ctk.CTkFrame(self.root, fg_color="#eef5fa", corner_radius=8)
        self.opt_frame.pack(fill="x", padx=15, pady=5)
        
        self.extract_mode_var = ctk.StringVar(value="PDF 原生文字提取")
        self.rm_mode_var = ctk.StringVar(value="PDF 區域去浮水印")
        self.quality_var = ctk.StringVar(value="原畫質 (300 DPI)")
        self.wm_pos_var = ctk.StringVar(value="右下角")
        self.stamp_page_var = ctk.StringVar(value="全部頁面")
        self.rotate_var = ctk.StringVar(value="90度")
        self.ppt_mode_var = ctk.StringVar(value="純圖片簡報 (較快)")

        # 選項元件
        self.lbl_ext_mode = ctk.CTkLabel(self.opt_frame, text="📄 來源選項:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_ext_mode = ctk.CTkOptionMenu(self.opt_frame, variable=self.extract_mode_var, values=["PDF 原生文字提取", "圖片 AI OCR 辨識"], width=160, command=self.update_options_ui)
        self.lbl_rm_mode = ctk.CTkLabel(self.opt_frame, text="🧹 抹除模式:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_rm_mode = ctk.CTkOptionMenu(self.opt_frame, variable=self.rm_mode_var, values=["PDF 區域去浮水印", "圖片 AI 智慧抹除文字"], width=170, command=self.update_options_ui)
        self.lbl_dpi = ctk.CTkLabel(self.opt_frame, text="⚙️ 畫質:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_dpi = ctk.CTkOptionMenu(self.opt_frame, variable=self.quality_var, values=["原畫質 (300 DPI)", "高畫質 (200 DPI)", "中畫質 (150 DPI)", "低畫質 (72 DPI)"], width=130)
        self.lbl_wm = ctk.CTkLabel(self.opt_frame, text="📍 位置:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_wm = ctk.CTkOptionMenu(self.opt_frame, variable=self.wm_pos_var, values=["右下角", "左下角", "右上角", "左上角", "正中央"], width=100)
        self.lbl_stamp = ctk.CTkLabel(self.opt_frame, text="📑 目標頁面:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_stamp = ctk.CTkOptionMenu(self.opt_frame, variable=self.stamp_page_var, values=["全部頁面", "僅第一頁", "僅最後一頁"], width=100)
        self.lbl_rot = ctk.CTkLabel(self.opt_frame, text="🔄 旋轉:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_rot = ctk.CTkOptionMenu(self.opt_frame, variable=self.rotate_var, values=["90度", "180度", "270度"], width=90)
        self.lbl_ppt = ctk.CTkLabel(self.opt_frame, text="📄 模式:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_ppt = ctk.CTkOptionMenu(self.opt_frame, variable=self.ppt_mode_var, values=["圖文排版 (智慧 OCR)", "純圖片簡報 (較快)"], width=160)
        
        self.mode_var.trace_add("write", self.update_options_ui)
        self.update_options_ui() 

        # 拖曳區
        self.drop_frame = ctk.CTkFrame(self.root, fg_color="#f9f9f9", border_width=2, border_color="#3a7ebf", corner_radius=15, height=70)
        self.drop_frame.pack(fill="x", padx=15, pady=5)
        self.drop_frame.pack_propagate(False)
        self.status_label = ctk.CTkLabel(self.drop_frame, text="📁 將檔案拖曳至此 或 點擊選擇檔案", font=("Microsoft JhengHei", 15, "bold"), text_color="#555")
        self.status_label.pack(expand=True)
        self.drop_frame.bind("<Button-1>", lambda e: self.browse_file()); self.status_label.bind("<Button-1>", lambda e: self.browse_file())
        self.root.drop_target_register(DND_FILES); self.root.dnd_bind('<<Drop>>', self.on_drop)

        # 進度與取消
        action_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        action_frame.pack(fill="x", padx=15, pady=5)
        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.pack(side="left", expand=True, fill="x", padx=(0, 10)); self.progress_bar.set(0)
        self.cancel_btn = ctk.CTkButton(action_frame, text="取消任務", fg_color="#cc3333", hover_color="#aa2222", state="disabled", command=self.cancel_task, width=90)
        self.cancel_btn.pack(side="right")

        # 日誌終端區 (Console)
        self.log_box = ctk.CTkTextbox(self.root, height=120, font=("Consolas", 12), state="disabled", fg_color="#1e1e1e", text_color="#00ff00")
        self.log_box.pack(fill="both", expand=True, padx=15, pady=(5, 15))
        self.write_log("✅ 系統初始化完成，等待任務輸入...")

    def toggle_theme(self):
        mode = "dark" if self.theme_switch.get() == 1 else "light"
        ctk.set_appearance_mode(mode)
        bg = "#2b2b2b" if mode == "dark" else "#eef5fa"
        self.opt_frame.configure(fg_color=bg)

    def write_log(self, text):
        time_str = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{time_str}] {text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        # 同步更新上方精簡狀態列
        self.status_label.configure(text=text)

    def update_options_ui(self, *args):
        for widget in self.opt_frame.winfo_children(): widget.pack_forget()
        mode = self.mode_var.get()
        
        if mode == "EXTRACT_TXT":
            self.lbl_ext_mode.pack(side="left", padx=(10, 5), pady=5); self.menu_ext_mode.pack(side="left", padx=(0, 15))
        if mode == "RMWATERMARK":
            self.lbl_rm_mode.pack(side="left", padx=(10, 5), pady=5); self.menu_rm_mode.pack(side="left", padx=(0, 15))
            if self.rm_mode_var.get() == "PDF 區域去浮水印":
                self.lbl_dpi.pack(side="left", padx=(10, 5), pady=5); self.menu_dpi.pack(side="left", padx=(0, 15))
                self.lbl_wm.pack(side="left", padx=(5, 5), pady=5); self.menu_wm.pack(side="left", padx=(0, 15))
        if mode in ["PPT", "PDF2IMG", "GRAYSCALE", "FLATTEN"]:
            self.lbl_dpi.pack(side="left", padx=(10, 5), pady=5); self.menu_dpi.pack(side="left", padx=(0, 15))
        if mode == "PPT":
            self.lbl_ppt.pack(side="left", padx=(10, 5), pady=5); self.menu_ppt.pack(side="left", padx=(0, 15))
        if mode == "IMG_WM":
            self.lbl_stamp.pack(side="left", padx=(10, 5), pady=5); self.menu_stamp.pack(side="left", padx=(0, 15))
            self.lbl_wm.pack(side="left", padx=(5, 5), pady=5); self.menu_wm.pack(side="left", padx=(0, 15))
        if mode == "ROTATE":
            self.lbl_rot.pack(side="left", padx=(10, 5), pady=5); self.menu_rot.pack(side="left", padx=(0, 15))

    def cancel_task(self):
        self.stop_event.set()
        self.write_log("⚠️ 正在發送中斷訊號...")

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files: self.process_selected_files(files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode == "EXTRACT_TXT":
            if self.extract_mode_var.get() == "PDF 原生文字提取": file_paths = filedialog.askopenfilenames(title="選擇", filetypes=[("PDF", "*.pdf")])
            else: file_paths = filedialog.askopenfilenames(title="選擇圖片", filetypes=[("Images", "*.jpg;*.jpeg;*.png")])
        elif mode == "RMWATERMARK":
            if self.rm_mode_var.get() == "PDF 區域去浮水印": file_paths = filedialog.askopenfilenames(title="選擇", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.png")])
            else: file_paths = filedialog.askopenfilenames(title="選擇圖片", filetypes=[("Images", "*.jpg;*.jpeg;*.png")])
        elif mode in ["PPT", "PDF2IMG"]: file_paths = filedialog.askopenfilenames(title="選擇", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.png")])
        elif mode == "IMG2PDF": file_paths = filedialog.askopenfilenames(title="選擇", filetypes=[("Images", "*.jpg;*.jpeg;*.png")])
        elif mode == "MERGE": file_paths = filedialog.askopenfilenames(title="選擇", filetypes=[("支援格式", "*.pdf;*.jpg;*.png;*.docx;*.doc")])
        else: file_paths = filedialog.askopenfilenames(title="選擇 PDF", filetypes=[("PDF", "*.pdf")])
            
        if file_paths: self.process_selected_files(file_paths)

    def process_selected_files(self, file_paths):
        mode = self.mode_var.get()
        
        valid_exts = ('.pdf', '.jpg', '.jpeg', '.png')
        if mode == "MERGE": valid_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.docx', '.doc')
            
        valid_files = [f for f in file_paths if f.lower().endswith(valid_exts)]
        if not valid_files: return messagebox.showerror("錯誤", "請提供支援的檔案格式！")

        first_file = valid_files[0]
        first_file_name = os.path.splitext(os.path.basename(first_file))[0]

        if mode != "UNLOCK" and first_file.lower().endswith(".pdf"):
            if check_is_encrypted(first_file):
                return messagebox.showwarning("檔案已加密", "⚠️ 此 PDF 受到密碼保護！請先使用「解鎖 PDF」功能。")

        if mode in ["MERGE", "IMG2PDF"]: 
            return ListManagerWindow(self, valid_files, mode, lambda sf: self.trigger_list_process(mode, sf, first_file_name))

        output_path = None
        extra_args = {}
        
        if mode == "PDF2WORD": output_path = filedialog.asksaveasfilename(title="儲存 Word", initialfile=f"{first_file_name}.docx", defaultextension=".docx")
        elif mode == "PDF2EXCEL": output_path = filedialog.asksaveasfilename(title="儲存 Excel", initialfile=f"{first_file_name}.xlsx", defaultextension=".xlsx")
        elif mode == "CROP": output_path = filedialog.asksaveasfilename(title="儲存裁切檔", initialfile=f"{first_file_name}_cropped.pdf", defaultextension=".pdf")
        elif mode == "REDACT":
            keyword = ctk.CTkInputDialog(text="請輸入要塗黑遮蔽的關鍵字 (例如姓名/身分證)：", title="機密遮蔽").get_input()
            if not keyword: return
            output_path = filedialog.asksaveasfilename(title="儲存安全檔", initialfile=f"{first_file_name}_redacted.pdf", defaultextension=".pdf")
            extra_args["keyword"] = keyword
            
        elif mode == "SPLIT": 
            ranges = ctk.CTkInputDialog(text="輸入提取頁碼 (如 1-3,5)\n留白則全部分割：", title="分割").get_input()
            if ranges is None: return 
            if ranges.strip() == "": output_path = filedialog.askdirectory(title="選擇資料夾")
            else: output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_extract.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode in ["REMOVE_PAGES", "INSERT_BLANK", "REORDER"]:
            ranges = ctk.CTkInputDialog(text="請輸入指定頁碼：", title="頁面編輯").get_input()
            if not ranges: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_edited.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode == "EXTRACT_TXT":
            extra_args["extract_mode"] = self.extract_mode_var.get()
            suffix = "_OCR" if "OCR" in extra_args["extract_mode"] else ""
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}{suffix}.txt", defaultextension=".txt")
            
        elif mode == "EXTRACT_IMGS": output_path = filedialog.askdirectory(title="選擇儲存資料夾")
        elif mode == "GRAYSCALE": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_bw.pdf", defaultextension=".pdf")
        elif mode == "FLATTEN": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_flat.pdf", defaultextension=".pdf")
        elif mode == "PDF2IMG": output_path = filedialog.askdirectory(title="選擇儲存資料夾")
        elif mode in ["PROTECT", "UNLOCK"]:
            pwd = ctk.CTkInputDialog(text="請輸入密碼：", title="密碼").get_input()
            if not pwd: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_secure.pdf", defaultextension=".pdf")
            extra_args["pwd"] = pwd
            
        elif mode == "ADD_WM":
            txt = ctk.CTkInputDialog(text="請輸入浮水印文字：", title="文字浮水印").get_input()
            if not txt: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_wm.pdf", defaultextension=".pdf")
            extra_args["text"] = txt
            
        elif mode == "IMG_WM":
            img_path = filedialog.askopenfilename(title="選擇印章/圖片", filetypes=[("Images", "*.png;*.jpg")])
            if not img_path: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_stamp.pdf", defaultextension=".pdf")
            extra_args["img_path"] = img_path
            extra_args["position"] = self.wm_pos_var.get()
            extra_args["target_page"] = self.stamp_page_var.get()
            
        elif mode == "ADD_PAGE_NUM": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_pages.pdf", defaultextension=".pdf")
        elif mode == "ROTATE":
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_rot.pdf", defaultextension=".pdf")
            extra_args["angle"] = self.rotate_var.get()
            
        elif mode == "COMPRESS": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_comp.pdf", defaultextension=".pdf")
            
        elif mode == "RMWATERMARK": 
            extra_args["rm_mode"] = self.rm_mode_var.get()
            if "PDF" in extra_args["rm_mode"]:
                output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_clean", defaultextension=".pdf", filetypes=[("PDF", "*.pdf"), ("PPT", "*.pptx")])
                extra_args["position"] = self.wm_pos_var.get()
            else: output_path = filedialog.asksaveasfilename(title="儲存圖片", initialfile=f"{first_file_name}_clean.jpg", defaultextension=".jpg")
            
        elif mode == "PPT":
            extra_args["ppt_mode"] = self.ppt_mode_var.get()
            if len(valid_files) == 1: output_path = filedialog.asksaveasfilename(title="儲存 PPT", initialfile=first_file_name, defaultextension=".pptx")
            else:
                output_path = filedialog.askdirectory(title="選擇資料夾")
                valid_files = ["BATCH_MODE"] + list(valid_files)

        if output_path:
            dpi_map = {"原畫質 (300 DPI)": 300, "高畫質 (200 DPI)": 200, "中畫質 (150 DPI)": 150, "低畫質 (72 DPI)": 72}
            extra_args["dpi"] = dpi_map.get(self.quality_var.get(), 300)
            self.start_thread(mode, valid_files, output_path, extra_args)

    def trigger_list_process(self, mode, sorted_files, first_file_name):
        ext = ".pdf"
        out = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_merged.pdf", defaultextension=ext)
        if out: self.start_thread(mode, sorted_files, out, {})

    def start_thread(self, mode, input_data, output_data, extra_args):
        self.set_ui_state("disabled")
        self.stop_event.clear(); self.progress_bar.set(0)
        self.log_box.configure(state="normal"); self.log_box.delete("1.0", "end"); self.log_box.configure(state="disabled")
        threading.Thread(target=self.run_task_router, args=(mode, input_data, output_data, extra_args), daemon=True).start()

    def set_ui_state(self, state):
        self.drop_frame.configure(border_color="#3a7ebf" if state != "disabled" else "#cccccc")
        if state == "disabled":
            self.cancel_btn.configure(state="normal"); self.drop_frame.unbind("<Button-1>"); self.status_label.unbind("<Button-1>")
        else:
            self.cancel_btn.configure(state="disabled"); self.drop_frame.bind("<Button-1>", lambda e: self.browse_file()); self.status_label.bind("<Button-1>", lambda e: self.browse_file())

    def update_status(self, text, progress=None):
        self.root.after(0, lambda: self.write_log(text))
        if progress is not None: self.root.after(0, lambda: self.progress_bar.set(progress))

    def run_task_router(self, mode, input_data, output_data, extra):
        try:
            kwargs = {'status_callback': self.update_status, 'stop_event': self.stop_event}
            result_msg = None
            
            if mode == "MERGE": process_merge_pdfs(input_data, output_data, **kwargs)
            elif mode == "IMG2PDF": process_images_to_pdf(input_data, output_data, **kwargs)
            elif mode == "SPLIT": process_split_pdf(input_data[0], output_data, extra["ranges"], **kwargs)
            elif mode == "REMOVE_PAGES": process_remove_pages(input_data[0], output_data, extra["ranges"], **kwargs)
            elif mode == "INSERT_BLANK": process_insert_blank_page(input_data[0], output_data, extra["ranges"], **kwargs)
            elif mode == "REORDER": process_reorder_pages(input_data[0], output_data, extra["ranges"], **kwargs)
            elif mode == "PDF2WORD": process_pdf_to_word(input_data[0], output_data, **kwargs)
            elif mode == "PDF2EXCEL": process_pdf_to_excel(input_data[0], output_data, **kwargs)
            elif mode == "REDACT": process_redact_text(input_data[0], output_data, extra["keyword"], **kwargs)
            elif mode == "CROP": process_crop_pdf(input_data[0], output_data, **kwargs)
            elif mode == "EXTRACT_TXT":
                if "OCR" in extra.get("extract_mode", ""): process_image_ocr(input_data[0], output_data, **kwargs)
                else: process_extract_text(input_data[0], output_data, **kwargs)
            elif mode == "EXTRACT_IMGS": result_msg = process_extract_original_images(input_data[0], output_data, **kwargs)
            elif mode == "GRAYSCALE": process_to_grayscale(input_data[0], output_data, dpi=extra["dpi"], **kwargs)
            elif mode == "FLATTEN": process_flatten_pdf(input_data[0], output_data, dpi=extra["dpi"], **kwargs)
            elif mode == "PROTECT": process_protect_pdf(input_data[0], output_data, extra["pwd"], **kwargs)
            elif mode == "UNLOCK": process_unlock_pdf(input_data[0], output_data, extra["pwd"], **kwargs)
            elif mode == "ROTATE": process_rotate_pdf(input_data[0], output_data, extra["angle"], **kwargs)
            elif mode == "ADD_WM": process_add_watermark(input_data[0], output_data, extra["text"], **kwargs)
            elif mode == "IMG_WM": process_add_image_watermark(input_data[0], output_data, extra["img_path"], extra["position"], extra["target_page"], **kwargs)
            elif mode == "ADD_PAGE_NUM": process_add_page_numbers(input_data[0], output_data, **kwargs)
            elif mode == "PDF2IMG": process_pdf_to_images(input_data[0], output_data, dpi=extra["dpi"], **kwargs)
            elif mode == "COMPRESS": result_msg = process_compress_pdf(input_data[0], output_data, **kwargs)
            elif mode == "RMWATERMARK":
                if "AI" in extra.get("rm_mode", ""): process_image_remove_text(input_data[0], output_data, **kwargs)
                else: process_remove_watermark(input_data[0], output_data, dpi=extra["dpi"], position=extra["position"], **kwargs)
            elif mode == "PPT":
                if input_data[0] == "BATCH_MODE":
                    files = input_data[1:]
                    for idx, file in enumerate(files):
                        if self.stop_event.is_set(): break
                        base = os.path.splitext(os.path.basename(file))[0]
                        self.update_status(f"🔄 批次處理中 ({idx+1}/{len(files)}): {base}", idx/len(files))
                        process_pdf_to_ppt(file, os.path.join(output_data, base + ".pptx"), dpi=extra["dpi"], ppt_mode=extra["ppt_mode"], **kwargs)
                else: process_pdf_to_ppt(input_data[0], output_data, dpi=extra["dpi"], ppt_mode=extra["ppt_mode"], **kwargs)

            if self.stop_event.is_set(): self.write_log("⛔ 任務已中斷")
            else:
                self.write_log("✅ 任務執行成功！")
                self.root.after(0, lambda: self.progress_bar.set(1))
                msg = result_msg if result_msg else "作業成功！檔案已處理完成。"
                messagebox.showinfo("完成", msg)
                open_file_or_folder(output_data) 
        except Exception as e:
            self.write_log(f"❌ 發生錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"執行錯誤：\n{e}")
        finally: self.root.after(0, lambda: self.set_ui_state("normal"))

def main():
    set_dpi_awareness()
    _mutex = check_single_instance()
    root = CTkinterDnD()
    app = PDFToolApp(root)
    root.mainloop()

if __name__ == "__main__": main()
