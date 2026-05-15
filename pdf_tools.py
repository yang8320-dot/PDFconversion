import os
from pypdf import PdfWriter, PdfReader
import fitz
from pdf2image import convert_from_path
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
    status_callback("🖼️ 正在解析高畫質 PDF...", 0.1)
    pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    total = len(pages)
    for i, page_img in enumerate(pages):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在儲存圖片 {i+1} / {total}...", (i+1)/total)
        page_img.save(os.path.join(output_dir, f"{base_name}_{i+1}.jpg"), 'JPEG')

def process_remove_watermark(input_file, output_path, status_callback, stop_event):
    """ 純淨去浮水印功能 (無 OCR)：依據存檔副檔名導出為 PDF 或 PPT """
    status_callback("📄 正在讀取並轉換高解析度圖片...", 0.2)
    pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    total = len(pages)
    processed_pages = []
    
    for i, page_img in enumerate(pages):
        if stop_event.is_set(): return
        status_callback(f"🖌️ 正在抹除浮水印 {i+1} / {total}...", 0.2 + 0.6*(i/total))
        page_img = apply_watermark_removal(page_img) # 呼叫 utils 中的 125% 遮蔽邏輯
        processed_pages.append(page_img)
        
    status_callback("💾 正在組合寫入檔案...", 0.9)
    if not processed_pages: return
    
    if output_path.lower().endswith('.pptx'):
        # 導出為無可編輯文字的純圖片 PPT
        prs = Presentation()
        width_px, height_px = processed_pages[0].size
        prs.slide_width = int(width_px * 914400 / 300) # EMU 單位轉換
        prs.slide_height = int(height_px * 914400 / 300)
        
        for i, img in enumerate(processed_pages):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            temp_img = f"temp_clean_{i}.jpg"
            img.save(temp_img, "JPEG", quality=95)
            slide.shapes.add_picture(temp_img, 0, 0, prs.slide_width, prs.slide_height)
            os.remove(temp_img)
        prs.save(output_path)
    else:
        # 預設導出為 PDF
        processed_pages[0].save(output_path, "PDF", resolution=300.0, save_all=True, append_images=processed_pages[1:])

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
