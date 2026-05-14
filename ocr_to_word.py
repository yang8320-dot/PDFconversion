import os
import easyocr
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt as DocxPt
from utils import get_poppler_path

def process_ocr_to_word(pdf_path, output_path, status_callback):
    status_callback("⏳ 正在初始化 OCR 模型 (中英雙語)...")
    reader = easyocr.Reader(['ch_tra', 'ch_sim', 'en'])
    
    poppler_path = get_poppler_path()
    status_callback("📄 正在將 PDF 轉換為高解析度圖片...")
    pages = convert_from_path(pdf_path, dpi=200, poppler_path=poppler_path)
    total_pages = len(pages)

    doc = Document()

    for i, page_img in enumerate(pages):
        status_callback(f"🔍 正在辨識第 {i+1} / {total_pages} 頁...")
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        result = reader.readtext(temp_img, detail=0)
        extracted_text = "\n".join(result)
        
        p = doc.add_paragraph(extracted_text)
        p.style.font.name = 'Microsoft JhengHei'
        for run in p.runs:
            run.font.size = DocxPt(12)
        
        # 每頁 PDF 對應 Word 中的一頁 (最後一頁不加分頁符)
        if i < total_pages - 1:
            doc.add_page_break()
                
        os.remove(temp_img)
    
    status_callback("💾 正在寫入 Word 檔案...")
    doc.save(output_path)
