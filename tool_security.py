import os
import gc
from utils import format_size

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
        if not reader.decrypt(password): raise Exception("密碼錯誤，解鎖失敗！")
    writer = PdfWriter()
    for page in reader.pages:
        if stop_event.is_set(): return
        writer.add_page(page)
    status_callback("🔓 正在寫入無密碼檔案...", 0.7)
    with open(output_path, "wb") as f: writer.write(f)

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
        del pix; gc.collect()
    new_doc.save(output_path, garbage=4, deflate=True)
    doc.close(); new_doc.close()

def process_flatten_pdf(input_file, output_path, status_callback, stop_event, dpi=200):
    import fitz
    doc = fitz.open(input_file)
    new_doc = fitz.open()
    total = len(doc)
    for i in range(total):
        if stop_event.is_set(): return
        status_callback(f"🥞 正在扁平化 PDF ({i+1}/{total})...", (i+1)/total)
        page = doc[i]
        pix = page.get_pixmap(dpi=dpi)
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(page.rect, pixmap=pix)
        del pix; gc.collect()
    new_doc.save(output_path, garbage=4, deflate=True)
    doc.close(); new_doc.close()

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

def process_compress_pdf(input_file, output_path, status_callback, stop_event):
    import fitz
    status_callback("🗜️ 正在掃描與壓縮 PDF 檔案...", 0.5)
    orig_size = os.path.getsize(input_file)
    doc = fitz.open(input_file)
    if stop_event.is_set(): return
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return f"壓縮完成！\n原大小: {format_size(orig_size)} ➡️ 新大小: {format_size(os.path.getsize(output_path))}"
