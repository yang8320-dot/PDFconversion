import os
import gc
import easyocr
import fitz
from pdf2image import convert_from_path, pdfinfo_from_path
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_ppt(input_file, output_path, status_callback, stop_event, use_gpu, white_bg, use_ocr, dpi):
    is_pdf = input_file.lower().endswith('.pdf')
    prs = Presentation()
    
    # === 智慧極速模式 (非 OCR) ===
    if not use_ocr and is_pdf:
        status_callback("⚡ 啟動極速文字提取模式...", 0.1)
        doc = fitz.open(input_file)
        total_pages = len(doc)
        
        for i in range(total_pages):
            if stop_event.is_set(): break
            status_callback(f"⚡ 正在極速排版 PPT (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
            
            page = doc[i]
            # 產生背景圖片 (利用 fitz 內建轉圖更快速)
            pix = page.get_pixmap(dpi=dpi)
            temp_img = f"temp_fast_bg_{i}.jpg"
            pix.save(temp_img)
            
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(temp_img, 0, 0, width=prs.slide_width, height=prs.slide_height)
            os.remove(temp_img)

            # 抓取原生文字區塊並排版
            blocks = page.get_text("dict")["blocks"]
            scale_w = prs.slide_width / page.rect.width
            scale_h = prs.slide_height / page.rect.height
            
            for b in blocks:
                if b["type"] == 0:  # 0代表文字
                    text = "".join([s["text"] for l in b["lines"] for s in l["spans"]]).strip()
                    if not text: continue
                    
                    x0, y0, x1, y1 = b["bbox"]
                    left, top = int(x0 * scale_w), int(y0 * scale_h)
                    width, height = int((x1 - x0) * scale_w), int((y1 - y0) * scale_h)
                    
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    if white_bg:
                        fill = txBox.fill; fill.solid(); fill.fore_color.rgb = RGBColor(255, 255, 255)
                    
                    tf = txBox.text_frame
                    tf.word_wrap = False; tf.vertical_anchor = MSO_ANCHOR.TOP
                    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
                    
                    p = tf.add_paragraph()
                    p.text = text
                    estimated_pt = max(6, min(int((height / 12700) * 0.75), 96))
                    p.font.size = Pt(estimated_pt); p.font.name = '微軟正黑體'
        doc.close()
    
    # === 傳統 OCR 圖片辨識模式 ===
    else:
        status_callback("⏳ 正在初始化 OCR 模型...", 0.05)
        reader = easyocr.Reader(['ch_tra', 'en'], model_storage_directory=get_model_path(), download_enabled=False, gpu=use_gpu)
        poppler = get_poppler_path()
        
        if is_pdf:
            total_pages = pdfinfo_from_path(input_file, poppler_path=poppler)["Pages"]
        else:
            img_cache = [Image.open(input_file).convert('RGB')]
            total_pages = 1

        for i in range(total_pages):
            if stop_event.is_set(): break
            status_callback(f"🔍 正在辨識與排版 PPT (第 {i+1} / {total_pages} 頁)...", 0.1 + 0.8 * ((i+1)/total_pages))
            
            if is_pdf: page_img = convert_from_path(input_file, dpi=dpi, first_page=i+1, last_page=i+1, poppler_path=poppler)[0]
            else: page_img = img_cache[0]
                
            img_width, img_height = page_img.size
            temp_img = f"temp_ppt_page_{i}.jpg"
            page_img.save(temp_img, 'JPEG')
            
            result = reader.readtext(temp_img)
            slide = prs.slides.add_slide(prs.slide_layouts[6]) 
            slide.shapes.add_picture(temp_img, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            for (bbox, text, prob) in result:
                x_tl, y_tl = bbox[0]; x_br, y_br = bbox[2]
                left = int((x_tl / img_width) * prs.slide_width)
                top = int((y_tl / img_height) * prs.slide_height)
                width = int(((x_br - x_tl) / img_width) * prs.slide_width)
                height = int(((y_br - y_tl) / img_height) * prs.slide_height)
                
                txBox = slide.shapes.add_textbox(left, top, width, height)
                if white_bg:
                    fill = txBox.fill; fill.solid(); fill.fore_color.rgb = RGBColor(255, 255, 255)
                tf = txBox.text_frame
                tf.word_wrap = False; tf.vertical_anchor = MSO_ANCHOR.TOP
                tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
                p = tf.add_paragraph(); p.text = text
                estimated_pt = max(6, min(int((height / 12700) * 0.75), 96))
                p.font.size = Pt(estimated_pt); p.font.name = '微軟正黑體'
                    
            os.remove(temp_img)
            if is_pdf: del page_img; gc.collect()
    
    if not stop_event.is_set():
        status_callback("💾 正在寫入 PPT 檔案...", 0.95)
        prs.save(output_path)
