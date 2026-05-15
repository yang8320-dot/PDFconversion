import sys
import os
import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import customtkinter as ctk

# 匯入我們拆分出去的模組
from pdf_tools import process_merge_pdfs, process_protect_pdf, process_split_pdf
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
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

class CTkinterDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFconversion - 現代化多功能工具")
        self.root.geometry("750x550")
        
        # ====== UI 顏色調整為明亮主題 (白底) ======
        ctk.set_appearance_mode("light")  
        ctk.set_default_color_theme("blue")  

        self.mode_var = ctk.StringVar(value="PPT")

        # 頂部標題
        title_label = ctk.CTkLabel(self.root, text="PDFconversion 工具集", font=("Microsoft JhengHei", 24, "bold"), text_color="#333333")
        title_label.pack(pady=(20, 10))

        # 功能選擇區塊
        mode_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        mode_frame.pack(fill="x", padx=30, pady=10)

        ctk.CTkLabel(mode_frame, text="請選擇功能：", font=("Microsoft JhengHei", 14, "bold"), text_color="#333333").grid(row=0, column=0, padx=10, pady=10)
        
        modes = [
            ("OCR 轉 PPT", "PPT"),
            ("OCR 轉 Word", "WORD"),
            ("合併 PDF", "MERGE"),
            ("分割 PDF", "SPLIT"),
            ("加密 PDF", "PROTECT")
        ]
        
        col = 1
        for text, val in modes:
            rb = ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 13), text_color="#333333")
            rb.grid(row=0, column=col, padx=10, pady=10)
            col += 1

        # 拖曳放置區塊 (調整為淺色背景與深色邊框)
        self.drop_frame = ctk.CTkFrame(self.root, fg_color="#f9f9f9", border_width=2, border_color="#3a7ebf", corner_radius=15)
        self.drop_frame.pack(expand=True, fill="both", padx=30, pady=(10, 30))
        
        self.status_label = ctk.CTkLabel(
            self.drop_frame, 
            text="📁 將檔案拖曳至此\n或\n點擊這裡選擇檔案\n\n支援 PDF 與 圖片(JPG/PNG)\n支援批次處理多個檔案", 
            font=("Microsoft JhengHei", 16, "bold"),
            justify="center",
            text_color="#555555" # 深灰色文字，白底上更易讀
        )
        self.status_label.pack(expand=True)
        
        self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
        self.status_label.bind("<Button-1>", lambda e: self.browse_file())
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if not files: return
        self.process_selected_files(files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode in ["PPT", "WORD"]:
            file_paths = filedialog.askopenfilenames(title="選擇要轉換的檔案", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.jpeg;*.png")])
        else:
            file_paths = filedialog.askopenfilenames(title="選擇 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])
            
        if file_paths:
            self.process_selected_files(file_paths)

    def process_selected_files(self, file_paths):
        mode = self.mode_var.get()
        valid_files = [f for f in file_paths if f.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png'))]
        
        if not valid_files:
            messagebox.showerror("錯誤", "請提供有效的檔案！")
            return

        first_file_name = os.path.splitext(os.path.basename(valid_files[0]))[0]

        if mode == "MERGE":
            if len(valid_files) < 2:
                messagebox.showwarning("警告", "合併功能需要至少兩個檔案！")
                return
            output_path = filedialog.asksaveasfilename(title="儲存合併後的 PDF", initialfile=f"{first_file_name}_merged.pdf", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if output_path:
                self.start_thread(mode, valid_files, output_path, None)

        elif mode == "SPLIT":
            output_dir = filedialog.askdirectory(title="選擇分割後 PDF 的儲存資料夾")
            if output_dir:
                self.start_thread(mode, valid_files[0], output_dir, None)

        elif mode == "PROTECT":
            dialog = ctk.CTkInputDialog(text="請輸入要設定的 PDF 密碼：", title="加密 PDF")
            pwd = dialog.get_input()
            if pwd:
                output_path = filedialog.asksaveasfilename(title="儲存加密後的 PDF", initialfile=f"{first_file_name}_protected.pdf", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
                if output_path:
                    self.start_thread(mode, valid_files[0], output_path, pwd)

        elif mode in ["PPT", "WORD"]:
            if len(valid_files) == 1:
                ext = ".pptx" if mode == "PPT" else ".docx"
                ftype = [("PowerPoint", "*.pptx")] if mode == "PPT" else [("Word", "*.docx")]
                output_path = filedialog.asksaveasfilename(title="儲存檔案", initialfile=first_file_name, defaultextension=ext, filetypes=ftype)
                if output_path:
                    self.start_thread(mode, valid_files, output_path, None)
            else:
                output_dir = filedialog.askdirectory(title="選擇批次轉檔的儲存資料夾")
                if output_dir:
                    self.start_thread(mode, valid_files, output_dir, "BATCH")

    def start_thread(self, mode, input_data, output_data, extra_param):
        self.set_ui_state("disabled")
        threading.Thread(target=self.run_task_router, args=(mode, input_data, output_data, extra_param), daemon=True).start()

    def set_ui_state(self, state):
        color = "#3a7ebf" if state != "disabled" else "#cccccc"
        self.drop_frame.configure(border_color=color)
        
        if state == "disabled":
            self.drop_frame.unbind("<Button-1>")
            self.status_label.unbind("<Button-1>")
            self.root.dnd_bind('<<Drop>>', '')
        else:
            self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
            self.status_label.bind("<Button-1>", lambda e: self.browse_file())
            self.root.dnd_bind('<<Drop>>', self.on_drop)

    def update_status(self, text):
        self.root.after(0, lambda: self.status_label.configure(text=text))

    def run_task_router(self, mode, input_data, output_data, extra_param):
        try:
            if mode == "MERGE":
                process_merge_pdfs(input_data, output_data, self.update_status)
            elif mode == "SPLIT":
                process_split_pdf(input_data, output_data, self.update_status)
            elif mode == "PROTECT":
                process_protect_pdf(input_data, output_data, extra_param, self.update_status)
            
            elif mode in ["PPT", "WORD"]:
                if extra_param == "BATCH":
                    total = len(input_data)
                    ext = ".pptx" if mode == "PPT" else ".docx"
                    for idx, file in enumerate(input_data):
                        base_name = os.path.splitext(os.path.basename(file))[0]
                        out_path = os.path.join(output_data, base_name + ext)
                        self.update_status(f"🔄 批次處理中 ({idx+1}/{total}): {base_name}")
                        if mode == "PPT": process_ocr_to_ppt(file, out_path, self.update_status)
                        else: process_ocr_to_word(file, out_path, self.update_status)
                else:
                    if mode == "PPT": process_ocr_to_ppt(input_data[0], output_data, self.update_status)
                    else: process_ocr_to_word(input_data[0], output_data, self.update_status)

            self.update_status("✅ 任務完成！\n可以繼續選擇或拖曳下一個檔案")
            messagebox.showinfo("完成", f"作業成功！\n檔案已處理完成。")
        except Exception as e:
            self.update_status("❌ 發生錯誤，請重試")
            messagebox.showerror("錯誤", f"執行過程中發生錯誤：\n{e}")
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
