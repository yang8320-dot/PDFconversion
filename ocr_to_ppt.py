import os
import easyocr
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt
from utils import get_poppler_path

def process_ocr_to_ppt(pdf_path, output_path, status_callback):
    status_callback("⏳ 正在初始化 OCR 模型 (中英雙語)...")
    reader = easyocr.Reader(['ch_tra', 'ch_sim', 'en'])
    
    poppler_path = get_poppler_path()
    status_callback("📄 正在將 PDF 轉換為高解析度圖片...")
    pages = convert_from_path(pdf_path, dpi=200, poppler_path=poppler_path)
    total_pages = len(pages)

    prs = Presentation()

    for i, page_img in enumerate(pages):
        status_callback(f"🔍 正在辨識第 {i+1} / {total_pages} 頁...")
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        result = reader.readtext(temp_img, detail=0)
        extracted_text = "\n".join(result)
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), prs.slide_width - Inches(1), prs.slide_height - Inches(1))
        txBox.text_frame.word_wrap = True
        p = txBox.text_frame.add_paragraph()
        p.text = extracted_text
        p.font.size = Pt(14)
                
        os.remove(temp_img)
    
    status_callback("💾 正在寫入 PPT 檔案...")
    prs.save(output_path)
