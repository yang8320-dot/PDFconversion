import os
from pypdf import PdfWriter, PdfReader
import fitz  # PyMuPDF 用於壓縮
from pdf2image import convert_from_path
from utils import get_poppler_path

def process_merge_pdfs(input_files, output_path, status_callback, stop_event):
    """
    合併多個 PDF 檔案，並明確保留原有的書籤 (Bookmarks/Outlines)
    """
    merger = PdfWriter()
    total = len(input_files)
    
    for i, pdf in enumerate(input_files):
        if stop_event.is_set(): return
        
        # 更新狀態與進度條
        status_callback(f"📑 正在合併 PDF... ({i+1}/{total})", (i+1)/total)
        
        # append 預設行為已支援書籤匯入，但明確指定 import_outline=True 以確保標籤完整保留
        merger.append(pdf, import_outline=True) 
        
    merger.write(output_path)
    merger.close()

def process_protect_pdf(input_file, output_path, password, status_callback, stop_event):
    status_callback("🔒 正在讀取並加密 PDF...", 0.3)
    reader = PdfReader(input_file)
    writer = PdfWriter()
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
    status_callback("🔒 正在寫入加密檔案...", 0.7)
    writer.encrypt(password)
    with open(output_path, "wb") as f:
        writer.write(f)

def process_split_pdf(input_file, output_dir, status_callback, stop_event):
    reader = PdfReader(input_file)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    total = len(reader.pages)
    for i, page in enumerate(reader.pages):
        if stop_event.is_set(): return
        status_callback(f"✂️ 正在分割第 {i+1} / {total} 頁...", (i+1)/total)
        writer = PdfWriter()
        writer.add_page(page)
        with open(os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf"), "wb") as f:
            writer.write(f)

def process_pdf_to_images(input_file, output_dir, status_callback, stop_event):
    """ 新功能：PDF 每頁轉出為高畫質 JPG """
    status_callback("🖼️ 正在解析 PDF...", 0.1)
    poppler_path = get_poppler_path()
    pages = convert_from_path(input_file, dpi=300, poppler_path=poppler_path)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    total = len(pages)
    
    for i, page_img in enumerate(pages):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在儲存圖片 {i+1} / {total}...", (i+1)/total)
        page_img.save(os.path.join(output_dir, f"{base_name}_{i+1}.jpg"), 'JPEG')

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    """ 新功能：利用 PyMuPDF 垃圾回收與壓縮演算法對 PDF 進行瘦身 """
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    # garbage=4 (清理重複物件與未使用的流), deflate=True (重新壓縮資料)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
