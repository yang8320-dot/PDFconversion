import os
import easyocr
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt as DocxPt, Inches
from utils import get_poppler_path, get_model_path, apply_watermark_removal
from PIL import Image

def process_ocr_to_word(input_file, output_path, status_callback, stop_event, use_gpu, rm_wm):
    status_callback("⏳ 正在初始化 OCR 模型...", 0.05)
    
    model_path = get_model_path()
    # 支援 GPU 加速切換與強制離線模式
    reader = easyocr.Reader(['ch_tra', 'en'], model_storage_directory=model_path, download_enabled=False, gpu=use_gpu)
    
    if input_file.lower().endswith('.pdf'):
        status_callback("📄 正在將 PDF 轉為高解析度圖片 (300 DPI)...", 0.1)
        pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    else:
        status_callback("🖼️ 正在讀取圖片檔案...", 0.1)
        pages = [Image.open(input_file).convert('RGB')]

    total_pages = len(pages)
    doc = Document()

    for i, page_img in enumerate(pages):
        if stop_event.is_set(): break
        status_callback(f"🔍 正在辨識 Word 圖文 (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
        
        # 【去浮水印處理】如果使用者有勾選，則進行覆蓋抹除
        if rm_wm:
            page_img = apply_watermark_removal(page_img)
            
        temp_img = f"temp_word_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        # 【優化：先將原圖插入 Word 中作為對照】
        doc.add_paragraph(f"--- 第 {i+1} 頁 原始圖片 ---").style.font.name = 'Microsoft JhengHei'
        doc.add_picture(temp_img, width=Inches(6.0)) # 限制最大寬度避免爆框
        
        # 進行 OCR 辨識，開啟 paragraph=True 實作段落合併
        result = reader.readtext(temp_img, paragraph=True)
        
        doc.add_paragraph(f"--- 第 {i+1} 頁 辨識文字 ---").style.font.name = 'Microsoft JhengHei'
        for bbox, text in result:
            p = doc.add_paragraph(text)
            p.style.font.name = 'Microsoft JhengHei'
            for run in p.runs:
                run.font.size = DocxPt(11)
        
        # 若不是最後一頁，則加入分頁符號
        if i < total_pages - 1:
            doc.add_page_break() 
                
        os.remove(temp_img)
    
    if not stop_event.is_set():
        status_callback("💾 正在寫入 Word 檔案...", 0.95)
        doc.save(output_path)
