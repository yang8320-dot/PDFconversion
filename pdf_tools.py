import os
import gc
import tempfile
import shutil
from pypdf import PdfWriter, PdfReader
import fitz
from pdf2image import convert_from_path, pdfinfo_from_path
from pptx import Presentation
from utils import get_poppler_path, apply_watermark_removal

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

def process_pdf_to_images(input_file, output_dir, status_callback, stop_event):
    poppler = get_poppler_path()
    info = pdfinfo_from_path(input_file, poppler_path=poppler)
    total = info["Pages"]
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    
    # 【記憶體優化】每次迴圈只從 PDF 抽取 1 頁進記憶體，儲存後立刻回收
    for i in range(1, total + 1):
        if stop_event.is_set(): break
        status_callback(f"🖼️ 正在處理並儲存圖片 {i} / {total}...", i/total)
        page_img = convert_from_path(input_file, dpi=300, first_page=i, last_page=i, poppler_path=poppler)[0]
        page_img.save(os.path.join(output_dir, f"{base_name}_{i}.jpg"), 'JPEG')
        del page_img
        gc.collect()

def process_remove_watermark(input_file, output_path, status_callback, stop_event):
    poppler = get_poppler_path()
    info = pdfinfo_from_path(input_file, poppler_path=poppler)
    total = info["Pages"]
    
    # 使用暫存資料夾存放圖片，避免記憶體爆炸
    temp_dir = tempfile.mkdtemp()
    temp_images = []
    
    try:
        for i in range(1, total + 1):
            if stop_event.is_set(): break
            status_callback(f"🖌️ 正在抹除浮水印 {i} / {total}...", 0.1 + 0.7*(i/total))
            page_img = convert_from_path(input_file, dpi=300, first_page=i, last_page=i, poppler_path=poppler)[0]
            page_img = apply_watermark_removal(page_img)
            
            temp_path = os.path.join(temp_dir, f"page_{i}.jpg")
            page_img.save(temp_path, "JPEG", quality=95)
            temp_images.append(temp_path)
            
            del page_img
            gc.collect()
            
        if stop_event.is_set(): return
        status_callback("💾 正在組合寫入檔案...", 0.9)
        
        if output_path.lower().endswith('.pptx'):
            prs = Presentation()
            from PIL import Image
            with Image.open(temp_images[0]) as first_img:
                width_px, height_px = first_img.size
            prs.slide_width = int(width_px * 914400 / 300) 
            prs.slide_height = int(height_px * 914400 / 300)
            
            for img_path in temp_images:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)
            prs.save(output_path)
        else:
            # 建立高效的 PDF 寫入
            doc = fitz.open()
            for img_path in temp_images:
                img_doc = fitz.open(img_path)
                pdf_bytes = img_doc.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdf_bytes)
                doc.insert_pdf(img_pdf)
                img_doc.close()
                img_pdf.close()
            doc.save(output_path)
            doc.close()
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
