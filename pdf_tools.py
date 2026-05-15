from pypdf import PdfWriter, PdfReader
import os

def process_merge_pdfs(input_files, output_path, status_callback):
    """ 合併多個 PDF 檔案 """
    status_callback("📑 正在合併 PDF 檔案...")
    merger = PdfWriter()
    for pdf in input_files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def process_protect_pdf(input_file, output_path, password, status_callback):
    """ 為 PDF 加入密碼保護 """
    status_callback("🔒 正在加密 PDF 檔案...")
    reader = PdfReader(input_file)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    writer.encrypt(password)
    with open(output_path, "wb") as f:
        writer.write(f)

def process_split_pdf(input_file, output_dir, status_callback):
    """ 分割 PDF 檔案，每頁獨立存成一個檔案 """
    status_callback("✂️ 正在分割 PDF 檔案...")
    reader = PdfReader(input_file)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    
    total_pages = len(reader.pages)
    for i, page in enumerate(reader.pages):
        status_callback(f"✂️ 正在分割第 {i+1} / {total_pages} 頁...")
        writer = PdfWriter()
        writer.add_page(page)
        
        out_name = os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf")
        with open(out_name, "wb") as f:
            writer.write(f)
