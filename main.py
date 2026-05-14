import sys
import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

# 匯入我們拆分出去的模組
from merge_pdf import process_merge_pdfs
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

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多功能 PDF 與 OCR 工具")
        self.root.geometry("650x450")
        self.root.configure(bg="#f0f0f0")
        
        self.mode_var = tk.StringVar(value="PPT")
        mode_frame = tk.Frame(self.root, bg="#f0f0f0")
        mode_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(mode_frame, text="請選擇功能：", bg="#f0f0f0", font=("Microsoft JhengHei", 12, "bold")).pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="OCR 轉 PPT", variable=self.mode_var, value="PPT", bg="#f0f0f0", font=("Microsoft JhengHei", 11)).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(mode_frame, text="OCR 轉 Word", variable=self.mode_var, value="WORD", bg="#f0f0f0", font=("Microsoft JhengHei", 11)).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(mode_frame, text="合併多個 PDF", variable=self.mode_var, value="MERGE", bg="#f0f0f0", font=("Microsoft JhengHei", 11)).pack(side=tk.LEFT, padx=5)

        self.drop_frame = tk.Frame(self.root, bg="#ffffff", bd=2, relief="groove")
        self.drop_frame.pack(expand=True, fill=tk.BOTH, padx=30, pady=(0, 30))
        
        self.status_label = tk.Label(
            self.drop_frame, 
            text="📁 將 PDF 檔案拖曳至此\n或\n點擊這裡選擇檔案\n\n(若選擇合併功能，支援一次拖曳多個檔案)", 
            bg="#ffffff", fg="#555555", font=("Microsoft JhengHei", 13, "bold"),
            justify="center"
        )
        self.status_label.pack(expand=True)
        
        self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
        self.status_label.bind("<Button-1>", lambda e: self.browse_file())
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if not files: return
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showerror("錯誤", "請提供有效的 PDF 檔案！")
            return
            
        mode = self.mode_var.get()
        if mode in ["PPT", "WORD"]:
            if len(pdf_files) > 1:
                messagebox.showinfo("提示", "轉檔模式一次只能處理一個檔案，將僅處理第一個檔案。")
            self.ask_save_and_process(mode, [pdf_files[0]])
        elif mode == "MERGE":
            if len(pdf_files) < 2:
                messagebox.showwarning("警告", "合併功能需要至少提供兩個 PDF 檔案！")
                return
            pdf_files.sort()
            self.ask_save_and_process(mode, pdf_files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode in ["PPT", "WORD"]:
            file_path = filedialog.askopenfilename(title="選擇要轉換的 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])
            if file_path:
                self.ask_save_and_process(mode, [file_path])
        elif mode == "MERGE":
            file_paths = filedialog.askopenfilenames(title="選擇多個要合併的 PDF 檔案", filetypes=[("PDF Files", "*.pdf")])
            if file_paths:
                if len(file_paths) < 2:
                    messagebox.showwarning("警告", "合併功能需要至少選擇兩個 PDF 檔案！")
                    return
                sorted_paths = list(file_paths)
                sorted_paths.sort()
                self.ask_save_and_process(mode, sorted_paths)

    def ask_save_and_process(self, mode, input_files):
        if mode == "PPT":
            ext, title, ftypes = ".pptx", "選擇 PPT 儲存位置", [("PowerPoint Files", "*.pptx")]
        elif mode == "WORD":
            ext, title, ftypes = ".docx", "選擇 Word 儲存位置", [("Word Files", "*.docx")]
        else: 
            ext, title, ftypes = ".pdf", "選擇合併後的 PDF 儲存位置", [("PDF Files", "*.pdf")]

        output_path = filedialog.asksaveasfilename(title=title, defaultextension=ext, filetypes=ftypes)
        if not output_path: return
            
        self.set_ui_state("disabled")
        threading.Thread(target=self.run_task_router, args=(mode, input_files, output_path), daemon=True).start()

    def set_ui_state(self, state):
        if state == "disabled":
            self.drop_frame.unbind("<Button-1>")
            self.status_label.unbind("<Button-1>")
            self.root.dnd_bind('<<Drop>>', '')
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame) and widget != self.drop_frame:
                    for child in widget.winfo_children():
                        child.configure(state="disabled")
        else:
            self.drop_frame.bind("<Button-1>", lambda e: self.browse_file())
            self.status_label.bind("<Button-1>", lambda e: self.browse_file())
            self.root.dnd_bind('<<Drop>>', self.on_drop)
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame) and widget != self.drop_frame:
                    for child in widget.winfo_children():
                        child.configure(state="normal")

    def update_status(self, text):
        self.root.after(0, lambda: self.status_label.config(text=text))

    def run_task_router(self, mode, input_files, output_path):
        try:
            # 根據模式呼叫不同的外部模組，並把更新介面的方法 (self.update_status) 當作參數傳進去
            if mode == "MERGE":
                process_merge_pdfs(input_files, output_path, self.update_status)
            elif mode == "PPT":
                process_ocr_to_ppt(input_files[0], output_path, self.update_status)
            elif mode == "WORD":
                process_ocr_to_word(input_files[0], output_path, self.update_status)
                
            self.update_status("✅ 任務完成！\n可以繼續選擇或拖曳下一個檔案")
            messagebox.showinfo("完成", f"作業成功！\n檔案已儲存至：\n{output_path}")
        except Exception as e:
            self.update_status("❌ 發生錯誤，請重試")
            messagebox.showerror("錯誤", f"執行過程中發生錯誤：\n{e}")
        finally:
            self.root.after(0, lambda: self.set_ui_state("normal"))

def main():
    set_dpi_awareness()
    _mutex = check_single_instance()
    root = TkinterDnD.Tk()
    app = PDFToolApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
