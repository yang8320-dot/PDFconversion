import os
import easyocr
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Pt
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_ppt(input_file, output_path, status_callback):
    status_callback("⏳ 正在初始化 OCR 模型 (本地離線模式)...")
    
    # 載入離線模型，修正為繁英雙語避免衝突
    model_path = get_model_path()
    reader = easyocr.Reader(['ch_tra', 'en'], 
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
    prs = Presentation()
    
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    for i, page_img in enumerate(pages):
        status_callback(f"🔍 正在辨識第 {i+1} / {total_pages} 頁 (版面還原中)...")
        img_width, img_height = page_img.size
        
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        result = reader.readtext(temp_img)
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 使用空白投影片
        
        # 版面保留 (Layout Retention) 核心邏輯
        for (bbox, text, prob) in result:
            x_tl, y_tl = bbox[0]
            x_br, y_br = bbox[2]
            
            # 依據圖片原始座標按比例轉換成 PPT 投影片的座標
            left = int((x_tl / img_width) * slide_width)
            top = int((y_tl / img_height) * slide_height)
            width = int(((x_br - x_tl) / img_width) * slide_width)
            height = int(((y_br - y_tl) / img_height) * slide_height)
            
            # 建立文字方塊在精準位置上
            txBox = slide.shapes.add_textbox(left, top, width, height)
            txBox.text_frame.word_wrap = True
            p = txBox.text_frame.add_paragraph()
            p.text = text
            p.font.size = Pt(12)
                
        os.remove(temp_img)
    
    status_callback("💾 正在寫入 PPT 檔案...")
    prs.save(output_path)
