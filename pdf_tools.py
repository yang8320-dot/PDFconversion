import os
from pypdf import PdfWriter, PdfReader
import fitz
from pdf2image import convert_from_path
from utils import get_poppler_path, apply_watermark_removal
from PIL import Image

def process_merge_pdfs(input_files, output_path, status_callback, stop_event):
    merger = PdfWriter()
    total = len(input_files)
    for i, pdf in enumerate(input_files):
        if stop_event.is_set(): return
        status_callback(f"📑 正在合併 PDF... ({i+1}/{total})", (i+1)/total)
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
    with open(output_path, "wb") as f: writer.write(f)

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

def process_pdf_to_images(input_file, output_dir, status_callback, stop_event, rm_wm=False):
    status_callback("🖼️ 正在解析高畫質 PDF...", 0.1)
    # 強制使用 300 DPI
    pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    total = len(pages)
    for i, page_img in enumerate(pages):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在儲存圖片 {i+1} / {total}...", (i+1)/total)
        if rm_wm: page_img = apply_watermark_removal(page_img)
        page_img.save(os.path.join(output_dir, f"{base_name}_{i+1}.jpg"), 'JPEG')

def process_rebuild_pdf(input_file, output_path, status_callback, stop_event, rm_wm=False):
    """ 新功能：去浮水印並原汁原味輸出成 PDF """
    status_callback("📄 正在轉換高解析度圖片...", 0.2)
    pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    total = len(pages)
    processed_pages = []
    
    for i, page_img in enumerate(pages):
        if stop_event.is_set(): return
        status_callback(f"🖌️ 正在處理頁面 {i+1} / {total}...", 0.2 + 0.6*(i/total))
        if rm_wm: page_img = apply_watermark_removal(page_img)
        processed_pages.append(page_img)
        
    status_callback("💾 正在組合寫入新的 PDF 檔案...", 0.9)
    if processed_pages:
        processed_pages[0].save(output_path, "PDF", resolution=300.0, save_all=True, append_images=processed_pages[1:])

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
