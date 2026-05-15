import os
import easyocr
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt as DocxPt, Inches
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_word(input_file, output_path, status_callback, stop_event, use_gpu):
    status_callback("⏳ 正在初始化 OCR 模型...", 0.05)
    model_path = get_model_path()
    reader = easyocr.Reader(['ch_tra', 'en'], model_storage_directory=model_path, download_enabled=False, gpu=use_gpu)
    
    if input_file.lower().endswith('.pdf'):
        status_callback("📄 正在將 PDF 轉為圖片...", 0.1)
        pages = convert_from_path(input_file, dpi=200, poppler_path=get_poppler_path())
    else:
        pages = [Image.open(input_file).convert('RGB')]

    total_pages = len(pages)
    doc = Document()

    for i, page_img in enumerate(pages):
        if stop_event.is_set(): break
        status_callback(f"🔍 正在辨識 Word 圖文 (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
        
        temp_img = f"temp_word_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        # 【優化：先將原圖插入 Word 中作為對照】
        doc.add_paragraph(f"--- 第 {i+1} 頁 原始圖片 ---").style.font.name = 'Microsoft JhengHei'
        doc.add_picture(temp_img, width=Inches(6.0)) # 寬度限制在 6 吋避免爆框
        
        # 開啟 paragraph=True 合併段落
        result = reader.readtext(temp_img, paragraph=True)
        
        doc.add_paragraph(f"--- 第 {i+1} 頁 辨識文字 ---").style.font.name = 'Microsoft JhengHei'
        for bbox, text in result:
            p = doc.add_paragraph(text)
            p.style.font.name = 'Microsoft JhengHei'
            for run in p.runs:
                run.font.size = DocxPt(11)
        
        if i < total_pages - 1:
            doc.add_page_break() # 換頁處理下一張
                
        os.remove(temp_img)
    
    if not stop_event.is_set():
        status_callback("💾 正在寫入 Word 檔案...", 0.95)
        doc.save(output_path)
