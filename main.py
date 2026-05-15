import sys
import os
import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import customtkinter as ctk
import fitz  # PyMuPDF 用於智慧判斷

from pdf_tools import process_merge_pdfs, process_protect_pdf, process_split_pdf, process_pdf_to_images, process_compress_pdf, process_remove_watermark
from ocr_to_ppt import process_ocr_to_ppt
from ocr_to_word import process_ocr_to_word

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

class CTkinterDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class MergeManagerWindow(ctk.CTkToplevel):
    def __init__(self, parent, initial_files, start_callback):
        super().__init__(parent)
        self.title("PDF 合併管理器")
        self.geometry("600x350")
        self.start_callback = start_callback
        self.transient(parent)
        self.grab_set()

        list_frame = ctk.CTkFrame(self, fg_color="transparent")
        list_frame.pack(side="left", fill="both", expand=True, padx=(15, 5), pady=15)
        ctk.CTkLabel(list_frame, text="合併檔案列表 (由上到下)：", font=("Microsoft JhengHei", 14, "bold")).pack(anchor="w", pady=(0, 5))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Microsoft JhengHei", 11), yscrollcommand=scrollbar.set, activestyle="none")
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        for f in initial_files: self.listbox.insert(tk.END, f)
            
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(side="right", fill="y", padx=(5, 15), pady=15)
        ctk.CTkButton(btn_frame, text="➕ 新增檔案", command=lambda: [self.listbox.insert(tk.END, f) for f in filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])]).pack(pady=5)
        ctk.CTkButton(btn_frame, text="⬆️ 上移", command=self.move_up).pack(pady=5)
        ctk.CTkButton(btn_frame, text="⬇️ 下移", command=self.move_down).pack(pady=5)
        ctk.CTkButton(btn_frame, text="❌ 移除", command=self.remove_item, fg_color="#cc3333", hover_color="#aa2222").pack(pady=5)
        ctk.CTkButton(btn_frame, text="🚀 開始合併", command=self.start_merge, fg_color="#28a745", hover_color="#218838", height=35).pack(side="bottom", pady=15)

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
        if len(final_list) < 2: return messagebox.showwarning("警告", "需要至少兩個檔案！")
        self.start_callback(final_list)
        self.destroy()

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF工具")
        # 略微增加寬度以容納畫質選單
        self.root.geometry("860x290")
        ctk.set_appearance_mode("light")  
        ctk.set_default_color_theme("blue")  

        self.mode_var = ctk.StringVar(value="PPT")
        self.stop_event = threading.Event() 

        # 功能選擇區塊 (修正名稱)
        mode_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        mode_frame.pack(fill="x", padx=15, pady=(10, 5))
        for i in range(4): mode_frame.grid_columnconfigure(i, weight=1, uniform="col_group")
        
        modes1 = [("PDF 轉 PPT", "PPT"), ("PDF 轉 Word", "WORD"), ("NotebookLM 去浮水印", "RMWATERMARK"), ("合併 PDF", "MERGE")]
        modes2 = [("分割 PDF", "SPLIT"), ("加密 PDF", "PROTECT"), ("PDF 轉圖片", "PDF2IMG"), ("PDF 壓縮", "COMPRESS")]
        
        for i, (text, val) in enumerate(modes1):
            ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13)).grid(row=0, column=i, padx=5, pady=8, sticky="w")
        for i, (text, val) in enumerate(modes2):
            ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13)).grid(row=1, column=i, padx=5, pady=8, sticky="w")

        # 進階設定區塊 (加入畫質選項)
        opt_frame = ctk.CTkFrame(self.root, fg_color="#eef5fa", corner_radius=10)
        opt_frame.pack(fill="x", padx=15, pady=5)
        
        self.use_gpu_var = ctk.BooleanVar(value=False)
        self.white_bg_var = ctk.BooleanVar(value=False)
        self.quality_var = ctk.StringVar(value="原畫質 (300 DPI)")
        
        ctk.CTkLabel(opt_frame, text="⚙️ 設定:", font=("Microsoft JhengHei", 12, "bold")).pack(side="left", padx=10, pady=5)
        
        ctk.CTkOptionMenu(opt_frame, variable=self.quality_var, values=["原畫質 (300 DPI)", "高畫質 (200 DPI)", "中畫質 (150 DPI)", "低畫質 (72 DPI)"], font=("Microsoft JhengHei", 12), width=130).pack(side="left", padx=5)
        ctk.CTkCheckBox(opt_frame, text="啟用 GPU", variable=self.use_gpu_var, font=("Microsoft JhengHei", 12)).pack(side="left", padx=10)
        ctk.CTkCheckBox(opt_frame, text="PPT 白底覆蓋", variable=self.white_bg_var, font=("Microsoft JhengHei", 12)).pack(side="left", padx=10)

        # 拖曳區塊
        self.drop_frame = ctk.CTkFrame(self.root, fg_color="#f9f9f9", border_width=2, border_color="#3a7ebf", corner_radius=15, height=70)
        self.drop_frame.pack(fill="x", padx=15, pady=10)
        self.drop_frame.pack_propagate(False)
        self.status_label = ctk.CTkLabel(self.drop_frame, text="📁 將檔案拖曳至此 或 點擊選擇檔案", font=("Microsoft JhengHei", 14, "bold"), text_color="#555")
        self.status_label.pack(expand=True)
        self.drop_frame.bind("<Button-1>", lambda e: self.browse_file()); self.status_label.bind("<Button-1>", lambda e: self.browse_file())
        self.root.drop_target_register(DND_FILES); self.root.dnd_bind('<<Drop>>', self.on_drop)

        # 底部區塊
        bottom_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=15, pady=(0, 10))
        self.progress_bar = ctk.CTkProgressBar(bottom_frame)
        self.progress_bar.pack(side="left", expand=True, fill="x", padx=(0, 10)); self.progress_bar.set(0)
        self.cancel_btn = ctk.CTkButton(bottom_frame, text="取消任務", fg_color="#cc3333", hover_color="#aa2222", state="disabled", command=self.cancel_task, width=90)
        self.cancel_btn.pack(side="right")

    def check_is_native_pdf(self, filepath):
        """ 智慧偵測是否為含有文字的原生 PDF """
        if not filepath.lower().endswith(".pdf"): return False
        try:
            doc = fitz.open(filepath)
            # 檢查前三頁是否有足夠的文字量
            text_length = sum(len(doc[i].get_text("text").strip()) for i in range(min(3, len(doc))))
            doc.close()
            return text_length > 50
        except: return False

    def cancel_task(self):
        self.stop_event.set()
        self.update_status("⚠️ 正在停止任務中...", 0)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files: self.process_selected_files(files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode in ["PPT", "WORD", "PDF2IMG", "RMWATERMARK"]:
            file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.png")])
        else: file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF Files", "*.pdf")])
        if file_paths: self.process_selected_files(file_paths)

    def process_selected_files(self, file_paths):
        mode = self.mode_var.get()
        valid_files = [f for f in file_paths if f.lower().endswith(('.pdf', '.jpg', '.png'))]
        if not valid_files: return messagebox.showerror("錯誤", "請提供有效的檔案！")

        first_file = valid_files[0]
        first_file_name = os.path.splitext(os.path.basename(first_file))[0]

        if mode == "MERGE": return MergeManagerWindow(self.root, valid_files, lambda sf: self.trigger_merge_process(sf, first_file_name))

        # 智慧判斷是否需要 OCR
        use_ocr = True
        if mode in ["PPT", "WORD"] and len(valid_files) == 1 and self.check_is_native_pdf(first_file):
            answer = messagebox.askyesno("智慧偵測", "偵測到此 PDF 包含原生文字！\n\n是否使用「極速文字提取」？\n(速度極快且精準，無需使用 AI 圖片辨識)\n\n按 [是] 極速轉換\n按 [否] 強制圖片 OCR 辨識")
            use_ocr = not answer

        output_path = None
        if mode in ["SPLIT", "PDF2IMG"]: output_path = filedialog.askdirectory(title="選擇儲存資料夾")
        elif mode == "PROTECT":
            pwd = ctk.CTkInputDialog(text="請輸入密碼：", title="加密").get_input()
            if not pwd: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_protected.pdf", defaultextension=".pdf")
            valid_files = [first_file, pwd] 
        elif mode == "COMPRESS": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_compressed.pdf", defaultextension=".pdf")
        elif mode == "RMWATERMARK": output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_clean", defaultextension=".pdf", filetypes=[("PDF", "*.pdf"), ("PPT", "*.pptx")])
        elif mode in ["PPT", "WORD"]:
            if len(valid_files) == 1:
                ext = ".pptx" if mode == "PPT" else ".docx"
                output_path = filedialog.asksaveasfilename(title="儲存", initialfile=first_file_name, defaultextension=ext)
            else:
                output_path = filedialog.askdirectory(title="選擇批次轉檔的儲存資料夾")
                valid_files = ["BATCH_MODE"] + list(valid_files)

        if output_path:
            dpi_map = {"原畫質 (300 DPI)": 300, "高畫質 (200 DPI)": 200, "中畫質 (150 DPI)": 150, "低畫質 (72 DPI)": 72}
            dpi_val = dpi_map.get(self.quality_var.get(), 300)
            self.start_thread(mode, valid_files, output_path, use_ocr, dpi_val)

    def trigger_merge_process(self, sorted_files, first_file_name):
        out = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_merged.pdf", defaultextension=".pdf")
        if out: self.start_thread("MERGE", sorted_files, out, True, 300)

    def start_thread(self, mode, input_data, output_data, use_ocr, dpi):
        self.set_ui_state("disabled")
        self.stop_event.clear(); self.progress_bar.set(0)
        threading.Thread(target=self.run_task_router, args=(mode, input_data, output_data, use_ocr, dpi), daemon=True).start()

    def set_ui_state(self, state):
        self.drop_frame.configure(border_color="#3a7ebf" if state != "disabled" else "#cccccc")
        if state == "disabled":
            self.cancel_btn.configure(state="normal"); self.drop_frame.unbind("<Button-1>"); self.status_label.unbind("<Button-1>")
        else:
            self.cancel_btn.configure(state="disabled"); self.drop_frame.bind("<Button-1>", lambda e: self.browse_file()); self.status_label.bind("<Button-1>", lambda e: self.browse_file())

    def update_status(self, text, progress=None):
        self.root.after(0, lambda: self.status_label.configure(text=text))
        if progress is not None: self.root.after(0, lambda: self.progress_bar.set(progress))

    def run_task_router(self, mode, input_data, output_data, use_ocr, dpi):
        try:
            kwargs = {'status_callback': self.update_status, 'stop_event': self.stop_event}
            
            if mode == "MERGE": process_merge_pdfs(input_data, output_data, **kwargs)
            elif mode == "SPLIT": process_split_pdf(input_data[0], output_data, **kwargs)
            elif mode == "PROTECT": process_protect_pdf(input_data[0], output_data, input_data[1], **kwargs)
            elif mode == "PDF2IMG": process_pdf_to_images(input_data[0], output_data, dpi=dpi, **kwargs)
            elif mode == "COMPRESS": process_compress_pdf(input_data[0], output_data, **kwargs)
            elif mode == "RMWATERMARK": process_remove_watermark(input_data[0], output_data, dpi=dpi, **kwargs)
            elif mode in ["PPT", "WORD"]:
                ocr_kwargs = kwargs.copy()
                ocr_kwargs.update({'use_gpu': self.use_gpu_var.get(), 'use_ocr': use_ocr, 'dpi': dpi})
                if mode == "PPT": ocr_kwargs['white_bg'] = self.white_bg_var.get()
                
                if input_data[0] == "BATCH_MODE":
                    files = input_data[1:]
                    ext = ".pptx" if mode == "PPT" else ".docx"
                    for idx, file in enumerate(files):
                        if self.stop_event.is_set(): break
                        base = os.path.splitext(os.path.basename(file))[0]
                        self.update_status(f"🔄 批次處理中 ({idx+1}/{len(files)}): {base}", idx/len(files))
                        # 批次模式預設使用 OCR (因為難以逐一詢問)
                        if mode == "PPT": process_ocr_to_ppt(file, os.path.join(output_data, base + ext), **ocr_kwargs)
                        else: process_ocr_to_word(file, os.path.join(output_data, base + ext), **ocr_kwargs)
                else:
                    if mode == "PPT": process_ocr_to_ppt(input_data[0], output_data, **ocr_kwargs)
                    else: process_ocr_to_word(input_data[0], output_data, **ocr_kwargs)

            if self.stop_event.is_set(): self.update_status("⛔ 任務已取消", 0)
            else:
                self.update_status("✅ 任務完成！", 1)
                messagebox.showinfo("完成", "作業成功！檔案已處理完成。")
        except Exception as e:
            self.update_status("❌ 發生錯誤", 0); messagebox.showerror("錯誤", f"執行錯誤：\n{e}")
        finally: self.root.after(0, lambda: self.set_ui_state("normal"))

def main():
    set_dpi_awareness()
    _mutex = check_single_instance()
    root = CTkinterDnD()
    app = PDFToolApp(root)
    root.mainloop()

if __name__ == "__main__": main()
