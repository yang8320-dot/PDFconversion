import os
import gc
import tempfile
from utils import get_poppler_path, apply_watermark_removal, format_size

# --- 基礎合併與轉換 ---

def process_merge_pdfs(input_files, output_path, status_callback, stop_event):
    from pypdf import PdfWriter  
    merger = PdfWriter()
    total = len(input_files)
    for i, pdf in enumerate(input_files):
        if stop_event.is_set(): return
        status_callback(f"📑 正在合併 PDF... ({i+1}/{total})", (i+1)/total)
        merger.append(pdf, import_outline=True) 
    merger.write(output_path)
    merger.close()

def process_images_to_pdf(input_files, output_path, status_callback, stop_event):
    import fitz  
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

# --- 加密與解鎖 ---

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

# --- 頁面操作 ---

def parse_page_ranges(range_str, total_pages):
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
    reader = PdfReader(input_file)
    total_pages = len(reader.pages)
    target_pages = parse_page_ranges(page_ranges, total_pages)
    
    if os.path.isdir(output_path):
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        for i, p_idx in enumerate(target_pages):
            if stop_event.is_set(): return
            status_callback(f"✂️ 正在分割頁面 ({i+1}/{len(target_pages)})...", (i+1)/len(target_pages))
            writer = PdfWriter()
            writer.add_page(reader.pages[p_idx])
            with open(os.path.join(output_path, f"{base_name}_page_{p_idx+1}.pdf"), "wb") as f:
                writer.write(f)
    else:
        writer = PdfWriter()
        for i, p_idx in enumerate(target_pages):
            if stop_event.is_set(): return
            status_callback(f"✂️ 正在提取頁面 ({i+1}/{len(target_pages)})...", (i+1)/len(target_pages))
            writer.add_page(reader.pages[p_idx])
        with open(output_path, "wb") as f: writer.write(f)

def process_remove_pages(input_file, output_path, page_ranges, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    reader = PdfReader(input_file)
    total_pages = len(reader.pages)
    pages_to_remove = parse_page_ranges(page_ranges, total_pages)
    writer = PdfWriter()
    for i in range(total_pages):
        if stop_event.is_set(): return
        status_callback(f"🗑️ 正在掃描並剔除頁面 ({i+1}/{total_pages})...", (i+1)/total_pages)
        if i not in pages_to_remove: writer.add_page(reader.pages[i])
    with open(output_path, "wb") as f: writer.write(f)

def process_insert_blank_page(input_file, output_path, page_ranges, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    reader = PdfReader(input_file)
    total_pages = len(reader.pages)
    insert_after = parse_page_ranges(page_ranges, total_pages)
    writer = PdfWriter()
    for i in range(total_pages):
        if stop_event.is_set(): return
        status_callback(f"📎 正在處理並插入空白頁 ({i+1}/{total_pages})...", (i+1)/total_pages)
        page = reader.pages[i]
        writer.add_page(page)
        if i in insert_after: writer.add_blank_page(width=page.mediabox.width, height=page.mediabox.height)
    with open(output_path, "wb") as f: writer.write(f)

def process_reorder_pages(input_file, output_path, order_str, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    reader = PdfReader(input_file)
    writer = PdfWriter()
    indices = [int(x.strip())-1 for x in order_str.split(",") if x.strip().isdigit()]
    for i, idx in enumerate(indices):
        if stop_event.is_set(): return
        status_callback(f"🔀 正在重新排序... ({i+1}/{len(indices)})", (i+1)/len(indices))
        if 0 <= idx < len(reader.pages): writer.add_page(reader.pages[idx])
    with open(output_path, "wb") as f: writer.write(f)

def process_rotate_pdf(input_file, output_path, angle, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    status_callback("🔄 正在旋轉 PDF...", 0.5)
    reader = PdfReader(input_file)
    writer = PdfWriter()
    angle_int = int(angle.replace("度", ""))
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
        writer.pages[-1].rotate(angle_int)
    with open(output_path, "wb") as f: writer.write(f)

# --- 辦公室修改/編輯工具 ---

def process_add_page_numbers(input_file, output_path, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    for i in range(total):
        if stop_event.is_set(): return
        status_callback(f"🔢 正在加入頁碼 ({i+1}/{total})...", (i+1)/total)
        page = doc[i]
        rect = page.rect
        text = f"- {i+1} -"
        p = fitz.Point(rect.width / 2 - 15, rect.height - 20)
        page.insert_text(p, text, fontsize=11, color=(0, 0, 0))
    doc.save(output_path)
    doc.close()

def process_to_grayscale(input_file, output_path, status_callback, stop_event, dpi=200):
    import fitz
    doc = fitz.open(input_file)
    new_doc = fitz.open()
    total = len(doc)
    for i in range(total):
        if stop_event.is_set(): return
        status_callback(f"🖨️ 正在轉換為黑白/灰階 ({i+1}/{total})...", (i+1)/total)
        page = doc[i]
        pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(page.rect, pixmap=pix)
    new_doc.save(output_path, garbage=4, deflate=True)
    doc.close(); new_doc.close()

def process_flatten_pdf(input_file, output_path, status_callback, stop_event, dpi=200):
    """將 PDF 轉為純圖片以防止竄改 (扁平化)"""
    import fitz
    doc = fitz.open(input_file)
    new_doc = fitz.open()
    total = len(doc)
    for i in range(total):
        if stop_event.is_set(): return
        status_callback(f"🥞 正在扁平化 PDF 防止篡改 ({i+1}/{total})...", (i+1)/total)
        page = doc[i]
        pix = page.get_pixmap(dpi=dpi) # 保持彩色
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(page.rect, pixmap=pix)
    new_doc.save(output_path, garbage=4, deflate=True)
    doc.close(); new_doc.close()

def process_extract_text(input_file, output_path, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    with open(output_path, "w", encoding="utf-8") as f:
        for i, page in enumerate(doc):
            if stop_event.is_set(): return
            status_callback(f"📄 正在提取文字 ({i+1}/{total})...", (i+1)/total)
            text = page.get_text("text")
            f.write(f"--- 第 {i+1} 頁 ---\n")
            f.write(text + "\n\n")
    doc.close()

def process_extract_original_images(input_file, output_dir, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    count = 0
    for i in range(total):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在提取內嵌圖片 (掃描第 {i+1} 頁)...", (i+1)/total)
        for img_idx, img in enumerate(doc[i].get_images(True)):
            xref = img[0]
            base_img = doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext = base_img["ext"]
            with open(os.path.join(output_dir, f"page{i+1}_img{img_idx+1}.{ext}"), "wb") as f:
                f.write(img_bytes)
            count += 1
    doc.close()
    return f"提取完成！\n共從檔案中成功提取了 {count} 張原始圖片。"

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

# --- 浮水印處理 ---

def process_add_watermark(input_file, output_path, text, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    for i, page in enumerate(doc):
        if stop_event.is_set(): return
        status_callback(f"🖋️ 正在加入浮水印... ({i+1}/{total})", (i+1)/total)
        rect = page.rect
        page.insert_text(fitz.Point(50, rect.height / 2), text, fontsize=48, color=(1, 0, 0))
    doc.save(output_path)
    doc.close()

def process_add_image_watermark(input_file, output_path, img_path, position, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    # 打開圖片以獲取比例
    img_doc = fitz.open(img_path)
    img_rect = img_doc[0].rect
    img_doc.close()

    for i, page in enumerate(doc):
        if stop_event.is_set(): return
        status_callback(f"🖼️ 正在壓印圖片浮水印... ({i+1}/{total})", (i+1)/total)
        page_rect = page.rect
        
        # 假設浮水印寬度為頁面寬度的 25%
        w = page_rect.width * 0.25
        h = w * (img_rect.height / img_rect.width)
        
        if position == "右下角": target_rect = fitz.Rect(page_rect.width - w - 20, page_rect.height - h - 20, page_rect.width - 20, page_rect.height - 20)
        elif position == "左下角": target_rect = fitz.Rect(20, page_rect.height - h - 20, 20 + w, page_rect.height - 20)
        elif position == "右上角": target_rect = fitz.Rect(page_rect.width - w - 20, 20, page_rect.width - 20, 20 + h)
        else: target_rect = fitz.Rect(20, 20, 20 + w, 20 + h) # 左上角
            
        page.insert_image(target_rect, filename=img_path)
    doc.save(output_path)
    doc.close()

def process_remove_watermark(input_file, output_path, status_callback, stop_event, dpi=300, position="右下角"):
    from pdf2image import convert_from_path, pdfinfo_from_path
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
            from pptx import Presentation
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

# --- 各式轉檔 ---

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
