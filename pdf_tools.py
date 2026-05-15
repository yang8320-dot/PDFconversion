import os
import gc
import tempfile
from utils import get_poppler_path, apply_watermark_removal, format_size

# ⚠️ 啟動優化：已將 pypdf, fitz, pptx, pdf2image 移至各別函式內部 (延遲載入 Lazy Import)
# 這樣一來，程式雙擊啟動時不需要將所有模組載入記憶體，能大幅改善「卡頓等一下下」的問題。

def process_merge_pdfs(input_files, output_path, status_callback, stop_event):
    from pypdf import PdfWriter  # 延遲載入
    merger = PdfWriter()
    total = len(input_files)
    for i, pdf in enumerate(input_files):
        if stop_event.is_set(): return
        status_callback(f"📑 正在合併 PDF... ({i+1}/{total})", (i+1)/total)
        merger.append(pdf, import_outline=True) 
    merger.write(output_path)
    merger.close()

def process_images_to_pdf(input_files, output_path, status_callback, stop_event):
    import fitz  # 延遲載入
    """將多張圖片合併成單一 PDF"""
    doc = fitz.open()
    total = len(input_files)
    for i, img_path in enumerate(input_files):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在將圖片轉為 PDF... ({i+1}/{total})", (i+1)/total)
        img_doc = fitz.open(img_path)
        pdf_bytes = img_doc.convert_to_pdf()
        img_pdf = fitz.open("pdf", pdf_bytes)
        doc.insert_pdf(img_pdf)
        img_doc.close(); img_pdf.close()
    doc.save(output_path)
    doc.close()

def process_protect_pdf(input_file, output_path, password, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    status_callback("🔒 正在讀取並加密 PDF...", 0.3)
    reader = PdfReader(input_file)
    writer = PdfWriter()
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
    status_callback("🔒 正在寫入加密檔案...", 0.7)
    writer.encrypt(password)
    with open(output_path, "wb") as f: writer.write(f)

def process_unlock_pdf(input_file, output_path, password, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    """解除 PDF 保全密碼"""
    status_callback("🔓 正在嘗試解鎖 PDF...", 0.3)
    reader = PdfReader(input_file)
    if reader.is_encrypted:
        success = reader.decrypt(password)
        if not success: raise Exception("密碼錯誤，解鎖失敗！")
    writer = PdfWriter()
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
    status_callback("🔓 正在寫入無密碼檔案...", 0.7)
    with open(output_path, "wb") as f: writer.write(f)

def parse_page_ranges(range_str, total_pages):
    """解析如 1-3,5 的頁碼範圍字串，轉換為 0-based 索引列表"""
    if not range_str.strip(): return list(range(total_pages))
    pages = set()
    for part in range_str.replace(" ", "").split(","):
        if "-" in part:
            start, end = map(int, part.split("-"))
            pages.update(range(start - 1, end))
        else:
            pages.add(int(part) - 1)
    return sorted([p for p in pages if 0 <= p < total_pages])

def process_split_pdf(input_file, output_path, page_ranges, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    """分割或提取指定範圍的 PDF 頁面"""
    reader = PdfReader(input_file)
    total_pages = len(reader.pages)
    target_pages = parse_page_ranges(page_ranges, total_pages)
    
    if os.path.isdir(output_path): # 分割成多個檔案
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        for i, p_idx in enumerate(target_pages):
            if stop_event.is_set(): return
            status_callback(f"✂️ 正在分割頁面 ({i+1}/{len(target_pages)})...", (i+1)/len(target_pages))
            writer = PdfWriter()
            writer.add_page(reader.pages[p_idx])
            with open(os.path.join(output_path, f"{base_name}_page_{p_idx+1}.pdf"), "wb") as f:
                writer.write(f)
    else: # 提取合併成一個檔案
        writer = PdfWriter()
        for i, p_idx in enumerate(target_pages):
            if stop_event.is_set(): return
            status_callback(f"✂️ 正在提取頁面 ({i+1}/{len(target_pages)})...", (i+1)/len(target_pages))
            writer.add_page(reader.pages[p_idx])
        with open(output_path, "wb") as f: writer.write(f)

def process_rotate_pdf(input_file, output_path, angle, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    """旋轉 PDF 頁面"""
    status_callback("🔄 正在旋轉 PDF...", 0.5)
    reader = PdfReader(input_file)
    writer = PdfWriter()
    angle_int = int(angle.replace("度", ""))
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
        writer.pages[-1].rotate(angle_int)
    with open(output_path, "wb") as f: writer.write(f)

def process_add_watermark(input_file, output_path, text, status_callback, stop_event):
    import fitz
    """添加文字浮水印"""
    doc = fitz.open(input_file)
    total = len(doc)
    for i, page in enumerate(doc):
        if stop_event.is_set(): return
        status_callback(f"🖋️ 正在加入浮水印... ({i+1}/{total})", (i+1)/total)
        # 簡單在頁面中心稍微偏上方壓印半透明紅色文字
        rect = page.rect
        page.insert_text(fitz.Point(50, rect.height / 2), text, fontsize=48, color=(1, 0, 0))
    doc.save(output_path)
    doc.close()

def process_pdf_to_images(input_file, output_dir, status_callback, stop_event, dpi=300):
    from pdf2image import convert_from_path, pdfinfo_from_path
    poppler = get_poppler_path()
    info = pdfinfo_from_path(input_file, poppler_path=poppler)
    total = info["Pages"]
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    
    for i in range(1, total + 1):
        if stop_event.is_set(): break
        status_callback(f"🖼️ 正在處理並儲存圖片 {i} / {total}...", i/total)
        page_img = convert_from_path(input_file, dpi=dpi, first_page=i, last_page=i, poppler_path=poppler)[0]
        page_img.save(os.path.join(output_dir, f"{base_name}_{i}.jpg"), 'JPEG')
        del page_img; gc.collect()

def process_pdf_to_ppt(input_file, output_path, status_callback, stop_event, dpi=300):
    from pptx import Presentation
    from pdf2image import convert_from_path, pdfinfo_from_path
    
    poppler = get_poppler_path()
    is_pdf = input_file.lower().endswith('.pdf')
    
    with tempfile.TemporaryDirectory() as temp_dir:
        prs = Presentation()
        if is_pdf:
            info = pdfinfo_from_path(input_file, poppler_path=poppler)
            total = info["Pages"]
            for i in range(1, total + 1):
                if stop_event.is_set(): break
                status_callback(f"🖼️ 正在轉換圖片並寫入 PPT (第 {i} / {total} 頁)...", i/total)
                page_img = convert_from_path(input_file, dpi=dpi, first_page=i, last_page=i, poppler_path=poppler)[0]
                
                temp_path = os.path.join(temp_dir, f"page_{i}.jpg")
                page_img.save(temp_path, "JPEG", quality=95)
                
                if i == 1:
                    prs.slide_width = int(page_img.width * 914400 / dpi)
                    prs.slide_height = int(page_img.height * 914400 / dpi)
                    
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(temp_path, 0, 0, prs.slide_width, prs.slide_height)
                del page_img; gc.collect()
        else:
            status_callback("🖼️ 正在處理圖片檔案...", 0.5)
            from PIL import Image
            page_img = Image.open(input_file).convert('RGB')
            temp_path = os.path.join(temp_dir, "page.jpg")
            page_img.save(temp_path, "JPEG", quality=95)
            prs.slide_width = int(page_img.width * 914400 / dpi)
            prs.slide_height = int(page_img.height * 914400 / dpi)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(temp_path, 0, 0, prs.slide_width, prs.slide_height)
            
        if not stop_event.is_set():
            status_callback("💾 正在儲存 PPT 檔案...", 0.95)
            prs.save(output_path)

def process_remove_watermark(input_file, output_path, status_callback, stop_event, dpi=300, position="右下角"):
    from pdf2image import convert_from_path, pdfinfo_from_path
    from pptx import Presentation
    
    poppler = get_poppler_path()
    info = pdfinfo_from_path(input_file, poppler_path=poppler)
    total = info["Pages"]
    temp_images = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        for i in range(1, total + 1):
            if stop_event.is_set(): break
            status_callback(f"🖌️ 正在抹除浮水印 {i} / {total}...", 0.1 + 0.7*(i/total))
            page_img = convert_from_path(input_file, dpi=dpi, first_page=i, last_page=i, poppler_path=poppler)[0]
            page_img = apply_watermark_removal(page_img, position)
            
            temp_path = os.path.join(temp_dir, f"page_{i}.jpg")
            page_img.save(temp_path, "JPEG", quality=95)
            temp_images.append(temp_path)
            del page_img; gc.collect()
            
        if stop_event.is_set(): return
        status_callback("💾 正在組合寫入檔案...", 0.9)
        
        if output_path.lower().endswith('.pptx'):
            from PIL import Image
            prs = Presentation()
            with Image.open(temp_images[0]) as first_img: width_px, height_px = first_img.size
            prs.slide_width = int(width_px * 914400 / dpi) 
            prs.slide_height = int(height_px * 914400 / dpi)
            
            for img_path in temp_images:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)
            prs.save(output_path)
        else:
            import fitz
            doc = fitz.open()
            for img_path in temp_images:
                img_doc = fitz.open(img_path)
                pdf_bytes = img_doc.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdf_bytes)
                doc.insert_pdf(img_pdf)
                img_doc.close(); img_pdf.close()
            doc.save(output_path)
            doc.close()

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    import fitz
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    orig_size = os.path.getsize(input_file)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    new_size = os.path.getsize(output_path)
    return f"壓縮完成！\n原大小: {format_size(orig_size)} ➡️ 新大小: {format_size(new_size)}"
