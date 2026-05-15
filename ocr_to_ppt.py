import os
import easyocr
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_ppt(input_file, output_path, status_callback, stop_event, use_gpu, white_bg):
    status_callback("⏳ 正在初始化 OCR 模型...", 0.05)
    reader = easyocr.Reader(['ch_tra', 'en'], model_storage_directory=get_model_path(), download_enabled=False, gpu=use_gpu)
    
    if input_file.lower().endswith('.pdf'):
        status_callback("📄 正在將 PDF 轉為高解析度圖片 (300 DPI)...", 0.1)
        pages = convert_from_path(input_file, dpi=300, poppler_path=get_poppler_path())
    else:
        status_callback("🖼️ 正在讀取圖片檔案...", 0.1)
        pages = [Image.open(input_file).convert('RGB')]

    total_pages = len(pages)
    prs = Presentation()
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    for i, page_img in enumerate(pages):
        if stop_event.is_set(): break
        status_callback(f"🔍 正在辨識與排版 PPT (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
            
        img_width, img_height = page_img.size
        temp_img = f"temp_ppt_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        result = reader.readtext(temp_img)
        slide = prs.slides.add_slide(prs.slide_layouts[6]) 
        slide.shapes.add_picture(temp_img, 0, 0, width=slide_width, height=slide_height)
        
        for (bbox, text, prob) in result:
            x_tl, y_tl = bbox[0]
            x_br, y_br = bbox[2]
            
            left = int((x_tl / img_width) * slide_width)
            top = int((y_tl / img_height) * slide_height)
            width = int(((x_br - x_tl) / img_width) * slide_width)
            height = int(((y_br - y_tl) / img_height) * slide_height)
            
            txBox = slide.shapes.add_textbox(left, top, width, height)
            
            if white_bg:
                fill = txBox.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(255, 255, 255)

            tf = txBox.text_frame
            # 防止文字自動換行與內邊距，並將錨點設在頂部，確保「絕對不跑針」
            tf.word_wrap = False
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            
            p = tf.add_paragraph()
            p.text = text
            
            # 使用更精準的公式：文字框高度轉換為 Pt，大約乘上 0.75 扣除行距
            estimated_pt = max(6, min(int((height / 12700) * 0.75), 96))
            p.font.size = Pt(estimated_pt)
            
            # 強制使用微軟正黑體，解決 EasyOCR 無法識別字體名稱的問題
            p.font.name = '微軟正黑體'
                
        os.remove(temp_img)
    
    if not stop_event.is_set():
        status_callback("💾 正在寫入 PPT 檔案...", 0.95)
        prs.save(output_path)
