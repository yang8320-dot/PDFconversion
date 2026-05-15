import os
import easyocr
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt as DocxPt
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_word(input_file, output_path, status_callback):
    status_callback("⏳ 正在初始化 OCR 模型 (本地離線模式)...")
    
    # 載入離線模型，避免 Proxy 錯誤
    model_path = get_model_path()
    reader = easyocr.Reader(['ch_tra', 'ch_sim', 'en'], 
                            model_storage_directory=model_path, 
                            download_enabled=False)
    
    is_pdf = input_file.lower().endswith('.pdf')
    if is_pdf:
        poppler_path = get_poppler_path()
        status_callback("📄 正在將 PDF 轉換為高解析度圖片...")
        pages = convert_from_path(input_file, dpi=200, poppler_path=poppler_path)
    else:
        status_callback("🖼️ 正在讀取圖片檔案...")
        pages = [Image.open(input_file).convert('RGB')]

    total_pages = len(pages)
    doc = Document()

    for i, page_img in enumerate(pages):
        status_callback(f"🔍 正在辨識第 {i+1} / {total_pages} 頁...")
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        # 開啟 paragraph=True 來實作基礎的「段落保留」，避免每行強制斷句
        result = reader.readtext(temp_img, paragraph=True)
        
        for bbox, text in result:
            p = doc.add_paragraph(text)
            p.style.font.name = 'Microsoft JhengHei'
            for run in p.runs:
                run.font.size = DocxPt(12)
        
        if i < total_pages - 1:
            doc.add_page_break()
                
        os.remove(temp_img)
    
    status_callback("💾 正在寫入 Word 檔案...")
    doc.save(output_path)
