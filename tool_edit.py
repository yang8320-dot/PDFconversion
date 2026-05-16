import os
import gc
import tempfile

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

def process_merge_pdfs(input_files, output_path, status_callback, stop_event):
    from pypdf import PdfWriter
    import fitz
    import pythoncom
    import sys
    import io
    
    pythoncom.CoInitialize()
    try:
        merger = PdfWriter()
        total = len(input_files)
        with tempfile.TemporaryDirectory() as temp_dir:
            for i, file_path in enumerate(input_files):
                if stop_event.is_set(): return
                base_name = os.path.basename(file_path)
                status_callback(f"📑 處理合併: {base_name} ({i+1}/{total})", (i+1)/total)
                ext = file_path.lower().split('.')[-1]
                
                if ext == 'pdf': merger.append(file_path)
                elif ext in ['jpg', 'jpeg', 'png', 'bmp']:
                    img_doc = fitz.open(file_path)
                    img_pdf = fitz.open("pdf", img_doc.convert_to_pdf())
                    temp_pdf = os.path.join(temp_dir, f"temp_{i}.pdf")
                    img_pdf.save(temp_pdf)
                    img_pdf.close(); img_doc.close(); del img_doc; gc.collect()
                    merger.append(temp_pdf)
                elif ext in ['docx', 'doc']:
                    from docx2pdf import convert
                    temp_pdf = os.path.join(temp_dir, f"temp_{i}.pdf")
                    dummy_out = io.StringIO()
                    old_out, old_err = sys.stdout, sys.stderr
                    sys.stdout, sys.stderr = dummy_out, dummy_out
                    try:
                        convert(file_path, temp_pdf)
                        merger.append(temp_pdf)
                    except Exception as e: raise Exception(f"Word 轉換失敗 ({base_name})\n確認沒開啟 Word 卡住。\n錯誤: {e}")
                    finally: sys.stdout, sys.stderr = old_out, old_err
                else: raise Exception(f"不支援的檔案格式: {ext}")
            
            if stop_event.is_set(): return
            status_callback("💾 正在輸出最終合併檔案...", 0.95)
            merger.write(output_path)
            merger.close()
    finally:
        pythoncom.CoUninitialize()

def process_split_pdf(input_file, output_path, page_ranges, status_callback, stop_event):
    from pypdf import PdfWriter, PdfReader
    reader = PdfReader(input_file)
    target_pages = parse_page_ranges(page_ranges, len(reader.pages))
    if os.path.isdir(output_path):
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        for i, p_idx in enumerate(target_pages):
            if stop_event.is_set(): return
            status_callback(f"✂️ 正在分割頁面 ({i+1}/{len(target_pages)})...", (i+1)/len(target_pages))
            writer = PdfWriter()
            writer.add_page(reader.pages[p_idx])
            with open(os.path.join(output_path, f"{base_name}_page_{p_idx+1}.pdf"), "wb") as f: writer.write(f)
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

def process_crop_pdf(input_file, output_path, status_callback, stop_event):
    import fitz
    doc = fitz.open(input_file)
    total = len(doc)
    for i, page in enumerate(doc):
        if stop_event.is_set(): return
        status_callback(f"✂️ 正在裁切第 {i+1} 頁白邊...", (i+1)/total)
        page.set_cropbox(page.get_bbox()) 
    doc.save(output_path, garbage=3, deflate=True)
    doc.close()
