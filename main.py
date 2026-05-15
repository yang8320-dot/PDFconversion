import sys
import os
import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import customtkinter as ctk

from pdf_tools import process_merge_pdfs, process_protect_pdf, process_split_pdf, process_pdf_to_images, process_compress_pdf
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
    except:
        try: ctypes.windll.user32.SetProcessDPIAware()
        except: pass

class CTkinterDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFconversion - 專業多功能工具 (Pro)")
        self.root.geometry("850x650")
        ctk.set_appearance_mode("light")  
        ctk.set_default_color_theme("blue")  

        self.mode_var = ctk.StringVar(value="PPT")
        self.stop_event = threading.Event() # 用於中斷任務

        # 頂部標題
        ctk.CTkLabel(self.root, text="PDFconversion 專業工具集", font=("Microsoft JhengHei", 24, "bold"), text_color="#333").pack(pady=(15, 5))

        # 功能選擇區塊
        mode_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        mode_frame.pack(fill="x", padx=20, pady=5)
        
        modes1 = [("OCR 轉 PPT", "PPT"), ("OCR 轉 Word", "WORD"), ("合併 PDF", "MERGE"), ("分割 PDF", "SPLIT")]
        modes2 = [("加密 PDF", "PROTECT"), ("PDF 轉圖片", "PDF2IMG"), ("PDF 壓縮", "COMPRESS")]
        
        for i, (text, val) in enumerate(modes1):
            ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13)).grid(row=0, column=i, padx=10, pady=5)
        for i, (text, val) in enumerate(modes2):
            ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13)).grid(row=1, column=i, padx=10, pady=5)

        # 進階設定區塊
        opt_frame = ctk.CTkFrame(self.root, fg_color="#eef5fa", corner_radius=10)
        opt_frame.pack(fill="x", padx=30, pady=5)
        
        self.use_gpu_var = ctk.BooleanVar(value=True)
        self.white_bg_var = ctk.BooleanVar(value=True)
        
        ctk.CTkLabel(opt_frame, text="⚙️ 進階設定 (僅限 OCR):", font=("Microsoft JhengHei", 12, "bold")).pack(side="left", padx=15, pady=5)
        ctk.CTkCheckBox(opt_frame, text="啟用 GPU 加速 (需 NVIDIA 顯卡)", variable=self.use_gpu_var, font=("Microsoft JhengHei", 12)).pack(side="left", padx=10)
        ctk.CTkCheckBox(opt_frame, text="PPT 自動白底覆蓋(適合白底文件)", variable=self.white_bg_var, font=("Microsoft JhengHei", 12)).pack(side="left", padx=10)

        # 拖曳放置區塊
        self.drop_frame = ctk.CTkFrame(self.root, fg_color="#f9f9f9", border_width=2, border_color="#3a7ebf", corner_radius=15)
        self.drop_frame.pack(expand=True, fill="both", padx=30, pady=10)
        
        self.status_label = ctk.CTkLabel(self.drop_frame, text="📁 將檔案拖曳至此 或 點擊選擇檔案\n支援批次處理", font=("Microsoft JhengHei", 16, "bold"), text_color="#555")
        self.status_label.pack(expand=True)
        
        self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
        self.status_label.bind("<Button-1>", lambda e: self.browse_file())
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        # 底部進度條與取消按鈕
        bottom_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        self.progress_bar = ctk.CTkProgressBar(bottom_frame)
        self.progress_bar.pack(side="left", expand=True, fill="x", padx=(0, 10))
        self.progress_bar.set(0)
        
        self.cancel_btn = ctk.CTkButton(bottom_frame, text="取消任務", fg_color="#cc3333", hover_color="#aa2222", state="disabled", command=self.cancel_task, width=100)
        self.cancel_btn.pack(side="right")

    def cancel_task(self):
        self.stop_event.set()
        self.update_status("⚠️ 正在停止任務中...", 0)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files: self.process_selected_files(files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode in ["PPT", "WORD", "PDF2IMG"]:
            file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.jpeg;*.png")])
        else:
            file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF Files", "*.pdf")])
        if file_paths: self.process_selected_files(file_paths)

    def process_selected_files(self, file_paths):
        mode = self.mode_var.get()
        valid_files = [f for f in file_paths if f.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png'))]
        if not valid_files: return messagebox.showerror("錯誤", "請提供有效的檔案！")

        first_file_name = os.path.splitext(os.path.basename(valid_files[0]))[0]
        output_path = None

        if mode == "MERGE":
            if len(valid_files) < 2: return messagebox.showwarning("警告", "合併需要至少兩個檔案！")
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_merged.pdf", defaultextension=".pdf")
        elif mode in ["SPLIT", "PDF2IMG"]:
            output_path = filedialog.askdirectory(title="選擇儲存資料夾")
        elif mode == "PROTECT":
            pwd = ctk.CTkInputDialog(text="請輸入密碼：", title="加密").get_input()
            if not pwd: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_protected.pdf", defaultextension=".pdf")
            valid_files = [valid_files[0], pwd] # 借用陣列傳密碼
        elif mode == "COMPRESS":
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_compressed.pdf", defaultextension=".pdf")
        elif mode in ["PPT", "WORD"]:
            if len(valid_files) == 1:
                ext = ".pptx" if mode == "PPT" else ".docx"
                output_path = filedialog.asksaveasfilename(title="儲存", initialfile=first_file_name, defaultextension=ext)
            else:
                output_path = filedialog.askdirectory(title="選擇批次轉檔的儲存資料夾")
                valid_files = ["BATCH_MODE"] + list(valid_files)

        if output_path:
            self.start_thread(mode, valid_files, output_path)

    def start_thread(self, mode, input_data, output_data):
        self.set_ui_state("disabled")
        self.stop_event.clear()
        self.progress_bar.set(0)
        threading.Thread(target=self.run_task_router, args=(mode, input_data, output_data), daemon=True).start()

    def set_ui_state(self, state):
        color = "#3a7ebf" if state != "disabled" else "#cccccc"
        self.drop_frame.configure(border_color=color)
        if state == "disabled":
            self.cancel_btn.configure(state="normal")
            self.drop_frame.unbind("<Button-1>")
            self.status_label.unbind("<Button-1>")
            self.root.dnd_bind('<<Drop>>', '')
        else:
            self.cancel_btn.configure(state="disabled")
            self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
            self.status_label.bind("<Button-1>", lambda e: self.browse_file())
            self.root.dnd_bind('<<Drop>>', self.on_drop)

    def update_status(self, text, progress=None):
        self.root.after(0, lambda: self.status_label.configure(text=text))
        if progress is not None:
            self.root.after(0, lambda: self.progress_bar.set(progress))

    def run_task_router(self, mode, input_data, output_data):
        try:
            kwargs = {'status_callback': self.update_status, 'stop_event': self.stop_event}
            
            if mode == "MERGE": process_merge_pdfs(input_data, output_data, **kwargs)
            elif mode == "SPLIT": process_split_pdf(input_data[0], output_data, **kwargs)
            elif mode == "PROTECT": process_protect_pdf(input_data[0], output_data, input_data[1], **kwargs)
            elif mode == "PDF2IMG": process_pdf_to_images(input_data[0], output_data, **kwargs)
            elif mode == "COMPRESS": process_compress_pdf(input_data[0], output_data, **kwargs)
            elif mode in ["PPT", "WORD"]:
                ocr_kwargs = kwargs.copy()
                ocr_kwargs['use_gpu'] = self.use_gpu_var.get()
                if mode == "PPT": ocr_kwargs['white_bg'] = self.white_bg_var.get()
                
                if input_data[0] == "BATCH_MODE":
                    files = input_data[1:]
                    ext = ".pptx" if mode == "PPT" else ".docx"
                    for idx, file in enumerate(files):
                        if self.stop_event.is_set(): break
                        base = os.path.splitext(os.path.basename(file))[0]
                        self.update_status(f"🔄 批次處理中 ({idx+1}/{len(files)}): {base}", idx/len(files))
                        out = os.path.join(output_data, base + ext)
                        if mode == "PPT": process_ocr_to_ppt(file, out, **ocr_kwargs)
                        else: process_ocr_to_word(file, out, **ocr_kwargs)
                else:
                    if mode == "PPT": process_ocr_to_ppt(input_data[0], output_data, **ocr_kwargs)
                    else: process_ocr_to_word(input_data[0], output_data, **ocr_kwargs)

            if self.stop_event.is_set():
                self.update_status("⛔ 任務已取消", 0)
                messagebox.showinfo("取消", "已成功中斷任務。")
            else:
                self.update_status("✅ 任務完成！可以繼續處理下一個檔案", 1)
                messagebox.showinfo("完成", "作業成功！檔案已處理完成。")
        except Exception as e:
            self.update_status("❌ 發生錯誤，請重試", 0)
            messagebox.showerror("錯誤", f"執行錯誤：\n{e}")
        finally:
            self.root.after(0, lambda: self.set_ui_state("normal"))

def main():
    set_dpi_awareness()
    _mutex = check_single_instance()
    root = CTkinterDnD()
    app = PDFToolApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
