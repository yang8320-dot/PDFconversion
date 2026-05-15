import os
import gc
import easyocr
from pdf2image import convert_from_path, pdfinfo_from_path
from docx import Document
from docx.shared import Pt as DocxPt, Inches
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_word(input_file, output_path, status_callback, stop_event, use_gpu):
    status_callback("⏳ 正在初始化 OCR 模型...", 0.05)
    reader = easyocr.Reader(['ch_tra', 'en'], model_storage_directory=get_model_path(), download_enabled=False, gpu=use_gpu)
    
    is_pdf = input_file.lower().endswith('.pdf')
    poppler = get_poppler_path()
    
    if is_pdf:
        status_callback("📄 正在分析 PDF 結構...", 0.1)
        info = pdfinfo_from_path(input_file, poppler_path=poppler)
        total_pages = info["Pages"]
    else:
        status_callback("🖼️ 正在讀取圖片檔案...", 0.1)
        img_cache = [Image.open(input_file).convert('RGB')]
        total_pages = 1

    doc = Document()

    # 【記憶體優化】一頁一頁讀取並釋放
    for i in range(total_pages):
        if stop_event.is_set(): break
        status_callback(f"🔍 正在辨識 Word 圖文 (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
        
        if is_pdf:
            page_img = convert_from_path(input_file, dpi=300, first_page=i+1, last_page=i+1, poppler_path=poppler)[0]
        else:
            page_img = img_cache[0]
            
        temp_img = f"temp_word_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        doc.add_paragraph(f"--- 第 {i+1} 頁 原始圖片 ---").style.font.name = 'Microsoft JhengHei'
        doc.add_picture(temp_img, width=Inches(6.0))
        
        result = reader.readtext(temp_img, paragraph=True)
        
        doc.add_paragraph(f"--- 第 {i+1} 頁 辨識文字 ---").style.font.name = 'Microsoft JhengHei'
        for bbox, text in result:
            p = doc.add_paragraph(text)
            p.style.font.name = 'Microsoft JhengHei'
            for run in p.runs:
                run.font.size = DocxPt(11)
        
        if i < total_pages - 1:
            doc.add_page_break() 
                
        os.remove(temp_img)
        
        if is_pdf:
            del page_img
            gc.collect() # 立刻清空記憶體
    
    if not stop_event.is_set():
        status_callback("💾 正在寫入 Word 檔案...", 0.95)
        doc.save(output_path)
