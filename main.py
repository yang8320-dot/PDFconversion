import sys
import os
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
                       process_add_page_numbers, process_reorder_pages, process_extract_original_images)
from utils import check_poppler_exists, open_file_or_folder

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

class ListManagerWindow(ctk.CTkToplevel):
    def __init__(self, parent, initial_files, mode, start_callback):
        super().__init__(parent)
        self.title("檔案合併管理器")
        self.geometry("600x350")
        self.start_callback = start_callback
        self.mode = mode
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
        
        file_types = [("PDF", "*.pdf")] if mode == "MERGE" else [("Images", "*.jpg;*.png")]
        ctk.CTkButton(btn_frame, text="➕ 新增檔案", command=lambda: [self.listbox.insert(tk.END, f) for f in filedialog.askopenfilenames(filetypes=file_types)]).pack(pady=5)
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
        if len(final_list) < 2 and self.mode == "MERGE": return messagebox.showwarning("警告", "需要至少兩個檔案！")
        self.start_callback(final_list)
        self.destroy()

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 辦公室全能工具箱")
        # 因應 5x4 排版，將寬度稍微拉寬，確保字體不擁擠
        self.root.geometry("900x420") 
        ctk.set_appearance_mode("light")  
        ctk.set_default_color_theme("blue")  

        if not check_poppler_exists():
            messagebox.showwarning("元件缺失", "找不到 Poppler 渲染元件，圖片轉檔相關功能可能受限。")

        self.mode_var = ctk.StringVar(value="PPT")
        self.stop_event = threading.Event() 

        # 功能選擇區塊 (5x4 排版)
        mode_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        mode_frame.pack(fill="x", padx=15, pady=(10, 5))
        for i in range(5): mode_frame.grid_columnconfigure(i, weight=1, uniform="col_group")
        
        # Row 1: 格式轉換
        modes1 = [("PDF 轉 PPT", "PPT"), ("PDF 轉圖片", "PDF2IMG"), ("圖片轉 PDF", "IMG2PDF"), ("提取純文字", "EXTRACT_TXT"), ("提取內嵌圖", "EXTRACT_IMGS")]
        # Row 2: 頁面管理
        modes2 = [("提取/分割 PDF", "SPLIT"), ("刪除指定頁", "REMOVE_PAGES"), ("插入空白頁", "INSERT_BLANK"), ("重新排序頁", "REORDER"), ("合併 PDF", "MERGE")]
        # Row 3: 內容修改
        modes3 = [("轉黑白/灰階", "GRAYSCALE"), ("PDF 壓縮", "COMPRESS"), ("PDF 旋轉", "ROTATE"), ("添加頁碼", "ADD_PAGE_NUM"), ("加文字水印", "ADD_WM")]
        # Row 4: 進階保全
        modes4 = [("加密 PDF", "PROTECT"), ("解鎖 PDF", "UNLOCK"), ("去浮水印", "RMWATERMARK"), ("", ""), ("", "")]
        
        for i, (text, val) in enumerate(modes1): ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 12)).grid(row=0, column=i, padx=2, pady=6, sticky="w")
        for i, (text, val) in enumerate(modes2): ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 12)).grid(row=1, column=i, padx=2, pady=6, sticky="w")
        for i, (text, val) in enumerate(modes3): ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 12)).grid(row=2, column=i, padx=2, pady=6, sticky="w")
        for i, (text, val) in enumerate(modes4): 
            if text: ctk.CTkRadioButton(mode_frame, text=text, variable=self.mode_var, value=val, font=("Microsoft JhengHei", 12)).grid(row=3, column=i, padx=2, pady=6, sticky="w")

        # 進階設定區塊
        self.opt_frame = ctk.CTkFrame(self.root, fg_color="#eef5fa", corner_radius=10)
        self.opt_frame.pack(fill="x", padx=15, pady=5)
        
        self.quality_var = ctk.StringVar(value="原畫質 (300 DPI)")
        self.wm_pos_var = ctk.StringVar(value="右下角")
        self.rotate_var = ctk.StringVar(value="90度")

        self.lbl_dpi = ctk.CTkLabel(self.opt_frame, text="⚙️ 畫質:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_dpi = ctk.CTkOptionMenu(self.opt_frame, variable=self.quality_var, values=["原畫質 (300 DPI)", "高畫質 (200 DPI)", "中畫質 (150 DPI)", "低畫質 (72 DPI)"], width=130)
        
        self.lbl_wm = ctk.CTkLabel(self.opt_frame, text="📍 浮水印位置:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_wm = ctk.CTkOptionMenu(self.opt_frame, variable=self.wm_pos_var, values=["右下角", "左下角", "右上角", "左上角"], width=100)

        self.lbl_rot = ctk.CTkLabel(self.opt_frame, text="🔄 旋轉角度:", font=("Microsoft JhengHei", 12, "bold"))
        self.menu_rot = ctk.CTkOptionMenu(self.opt_frame, variable=self.rotate_var, values=["90度", "180度", "270度"], width=90)

        self.mode_var.trace_add("write", self.update_options_ui)
        self.update_options_ui() 

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

    def update_options_ui(self, *args):
        for widget in self.opt_frame.winfo_children(): widget.pack_forget()
        mode = self.mode_var.get()
        if mode in ["PPT", "PDF2IMG", "RMWATERMARK", "GRAYSCALE"]:
            self.lbl_dpi.pack(side="left", padx=(10, 5), pady=5); self.menu_dpi.pack(side="left", padx=(0, 15))
        if mode == "RMWATERMARK":
            self.lbl_wm.pack(side="left", padx=(5, 5), pady=5); self.menu_wm.pack(side="left", padx=(0, 15))
        if mode == "ROTATE":
            self.lbl_rot.pack(side="left", padx=(10, 5), pady=5); self.menu_rot.pack(side="left", padx=(0, 15))

    def cancel_task(self):
        self.stop_event.set()
        self.update_status("⚠️ 正在停止任務中...", 0)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files: self.process_selected_files(files)

    def browse_file(self):
        mode = self.mode_var.get()
        if mode in ["PPT", "PDF2IMG", "RMWATERMARK", "IMG2PDF"]:
            file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF/Images", "*.pdf;*.jpg;*.png")])
        else: file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("PDF Files", "*.pdf")])
        if file_paths: self.process_selected_files(file_paths)

    def process_selected_files(self, file_paths):
        mode = self.mode_var.get()
        valid_files = [f for f in file_paths if f.lower().endswith(('.pdf', '.jpg', '.png'))]
        if not valid_files: return messagebox.showerror("錯誤", "請提供有效的檔案！")

        first_file = valid_files[0]
        first_file_name = os.path.splitext(os.path.basename(first_file))[0]

        if mode in ["MERGE", "IMG2PDF"]: 
            return ListManagerWindow(self.root, valid_files, mode, lambda sf: self.trigger_list_process(mode, sf, first_file_name))

        output_path = None
        extra_args = {}
        
        if mode == "SPLIT": 
            ranges = ctk.CTkInputDialog(text="請輸入提取頁碼 (如 1-3,5)\n若要獨立全部分割請留白：", title="提取/分割").get_input()
            if ranges is None: return 
            if ranges.strip() == "": output_path = filedialog.askdirectory(title="選擇分割檔儲存資料夾")
            else: output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_extracted.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode == "REMOVE_PAGES":
            ranges = ctk.CTkInputDialog(text="請輸入要【刪除】的頁碼 (如 12, 15-20)：", title="刪除頁面").get_input()
            if not ranges: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_removed.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode == "INSERT_BLANK":
            ranges = ctk.CTkInputDialog(text="請輸入要在哪些頁碼【之後】插入空白頁\n(例如輸入 1,3 代表在第1頁與第3頁後插入)：", title="插入空白頁").get_input()
            if not ranges: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_inserted.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode == "REORDER":
            ranges = ctk.CTkInputDialog(text="請輸入新的排序頁碼 (如 3, 1, 2)：", title="重新排序").get_input()
            if not ranges: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_reordered.pdf", defaultextension=".pdf")
            extra_args["ranges"] = ranges
            
        elif mode == "EXTRACT_TXT":
            output_path = filedialog.asksaveasfilename(title="儲存純文字檔", initialfile=f"{first_file_name}.txt", defaultextension=".txt")
            
        elif mode == "EXTRACT_IMGS":
            output_path = filedialog.askdirectory(title="選擇圖片提取的儲存資料夾")
            
        elif mode == "GRAYSCALE":
            output_path = filedialog.asksaveasfilename(title="儲存黑白 PDF", initialfile=f"{first_file_name}_bw.pdf", defaultextension=".pdf")

        elif mode == "PDF2IMG": output_path = filedialog.askdirectory(title="選擇儲存資料夾")
        
        elif mode == "PROTECT":
            pwd = ctk.CTkInputDialog(text="請輸入要設定的密碼：", title="加密").get_input()
            if not pwd: return
            output_path = filedialog.asksaveasfilename(title="儲存加密 PDF", initialfile=f"{first_file_name}_protected.pdf", defaultextension=".pdf")
            extra_args["pwd"] = pwd
            
        elif mode == "UNLOCK":
            pwd = ctk.CTkInputDialog(text="請輸入目前 PDF 的密碼：", title="解鎖").get_input()
            if not pwd: return
            output_path = filedialog.asksaveasfilename(title="儲存解鎖 PDF", initialfile=f"{first_file_name}_unlocked.pdf", defaultextension=".pdf")
            extra_args["pwd"] = pwd
            
        elif mode == "ADD_WM":
            txt = ctk.CTkInputDialog(text="請輸入要添加的文字：", title="添加浮水印").get_input()
            if not txt: return
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_wm.pdf", defaultextension=".pdf")
            extra_args["text"] = txt
            
        elif mode == "ADD_PAGE_NUM":
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_pages.pdf", defaultextension=".pdf")
            
        elif mode == "ROTATE":
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_rotated.pdf", defaultextension=".pdf")
            extra_args["angle"] = self.rotate_var.get()
            
        elif mode == "COMPRESS": 
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_compressed.pdf", defaultextension=".pdf")
            
        elif mode == "RMWATERMARK": 
            output_path = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_clean", defaultextension=".pdf", filetypes=[("PDF", "*.pdf"), ("PPT", "*.pptx")])
            extra_args["position"] = self.wm_pos_var.get()
            
        elif mode == "PPT":
            if len(valid_files) == 1: output_path = filedialog.asksaveasfilename(title="儲存 PPT", initialfile=first_file_name, defaultextension=".pptx")
            else:
                output_path = filedialog.askdirectory(title="選擇批次轉檔的儲存資料夾")
                valid_files = ["BATCH_MODE"] + list(valid_files)

        if output_path:
            dpi_map = {"原畫質 (300 DPI)": 300, "高畫質 (200 DPI)": 200, "中畫質 (150 DPI)": 150, "低畫質 (72 DPI)": 72}
            extra_args["dpi"] = dpi_map.get(self.quality_var.get(), 300)
            self.start_thread(mode, valid_files, output_path, extra_args)

    def trigger_list_process(self, mode, sorted_files, first_file_name):
        ext = ".pdf"
        suffix = "merged" if mode == "MERGE" else "from_images"
        out = filedialog.asksaveasfilename(title="儲存", initialfile=f"{first_file_name}_{suffix}.pdf", defaultextension=ext)
        if out: self.start_thread(mode, sorted_files, out, {})

    def start_thread(self, mode, input_data, output_data, extra_args):
        self.set_ui_state("disabled")
        self.stop_event.clear(); self.progress_bar.set(0)
        threading.Thread(target=self.run_task_router, args=(mode, input_data, output_data, extra_args), daemon=True).start()

    def set_ui_state(self, state):
        self.drop_frame.configure(border_color="#3a7ebf" if state != "disabled" else "#cccccc")
        if state == "disabled":
            self.cancel_btn.configure(state="normal"); self.drop_frame.unbind("<Button-1>"); self.status_label.unbind("<Button-1>")
        else:
            self.cancel_btn.configure(state="disabled"); self.drop_frame.bind("<Button-1>", lambda e: self.browse_file()); self.status_label.bind("<Button-1>", lambda e: self.browse_file())

    def update_status(self, text, progress=None):
        self.root.after(0, lambda: self.status_label.configure(text=text))
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
            elif mode == "EXTRACT_TXT": process_extract_text(input_data[0], output_data, **kwargs)
            elif mode == "EXTRACT_IMGS": result_msg = process_extract_original_images(input_data[0], output_data, **kwargs)
            elif mode == "GRAYSCALE": process_to_grayscale(input_data[0], output_data, dpi=extra["dpi"], **kwargs)
            elif mode == "PROTECT": process_protect_pdf(input_data[0], output_data, extra["pwd"], **kwargs)
            elif mode == "UNLOCK": process_unlock_pdf(input_data[0], output_data, extra["pwd"], **kwargs)
            elif mode == "ROTATE": process_rotate_pdf(input_data[0], output_data, extra["angle"], **kwargs)
            elif mode == "ADD_WM": process_add_watermark(input_data[0], output_data, extra["text"], **kwargs)
            elif mode == "ADD_PAGE_NUM": process_add_page_numbers(input_data[0], output_data, **kwargs)
            elif mode == "PDF2IMG": process_pdf_to_images(input_data[0], output_data, dpi=extra["dpi"], **kwargs)
            elif mode == "COMPRESS": result_msg = process_compress_pdf(input_data[0], output_data, **kwargs)
            elif mode == "RMWATERMARK": process_remove_watermark(input_data[0], output_data, dpi=extra["dpi"], position=extra["position"], **kwargs)
            elif mode == "PPT":
                if input_data[0] == "BATCH_MODE":
                    files = input_data[1:]
                    for idx, file in enumerate(files):
                        if self.stop_event.is_set(): break
                        base = os.path.splitext(os.path.basename(file))[0]
                        self.update_status(f"🔄 批次處理中 ({idx+1}/{len(files)}): {base}", idx/len(files))
                        process_pdf_to_ppt(file, os.path.join(output_data, base + ".pptx"), dpi=extra["dpi"], **kwargs)
                else:
                    process_pdf_to_ppt(input_data[0], output_data, dpi=extra["dpi"], **kwargs)

            if self.stop_event.is_set(): self.update_status("⛔ 任務已取消", 0)
            else:
                self.update_status("✅ 任務完成！", 1)
                msg = result_msg if result_msg else "作業成功！檔案已處理完成。"
                messagebox.showinfo("完成", msg)
                open_file_or_folder(output_data) 
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
