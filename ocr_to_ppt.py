import os
import easyocr
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from utils import get_poppler_path, get_model_path
from PIL import Image

def process_ocr_to_ppt(input_file, output_path, status_callback):
    status_callback("⏳ 正在初始化 OCR 模型 (本地離線模式)...")
    
    # 載入離線模型
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
        status_callback(f"🔍 正在辨識第 {i+1} / {total_pages} 頁 (保留原圖背景與文字還原中)...")
        img_width, img_height = page_img.size
        
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        result = reader.readtext(temp_img)
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 使用空白投影片
        
        # 【全新功能：將原本的頁面截圖當作 PPT 的最底層背景】
        # 因為是最先加入的物件，所以它會自動在最底層，不會蓋住後續生成的文字
        slide.shapes.add_picture(temp_img, 0, 0, width=slide_width, height=slide_height)
        
        for (bbox, text, prob) in result:
            x_tl, y_tl = bbox[0]
            x_br, y_br = bbox[2]
            
            # 依據圖片原始座標按比例轉換成 PPT 投影片的座標
            left = int((x_tl / img_width) * slide_width)
            top = int((y_tl / img_height) * slide_height)
            width = int(((x_br - x_tl) / img_width) * slide_width)
            height = int(((y_br - y_tl) / img_height) * slide_height)
            
            # 建立文字方塊 (會疊加在背景圖上面)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            
            # --- 💡 進階選項 ---
            # 如果你覺得疊加在背景上原本的字看起來會重疊/模糊，
            # 你可以解除下面三行的註解，這會為文字方塊加上「白色不透明底色」來蓋掉底圖原本的字
            # fill = txBox.fill
            # fill.solid()
            # fill.fore_color.rgb = RGBColor(255, 255, 255)
            # --------------------

            tf = txBox.text_frame
            tf.word_wrap = True
            
            # 【優化 1】將文字方塊內邊距歸零，確保定位不跑偏
            tf.margin_left = 0
            tf.margin_top = 0
            tf.margin_right = 0
            tf.margin_bottom = 0
            
            p = tf.add_paragraph()
            p.text = text
            
            # 【優化 2】根據 Bounding Box 高度動態估算字體大小
            estimated_pt = (height / 12700) * 0.75 
            estimated_pt = max(8, min(int(estimated_pt), 96))
            p.font.size = Pt(estimated_pt)
            
            # 【優化 3】強制設定預設字型為微軟正黑體
            p.font.name = '微軟正黑體'
                
        os.remove(temp_img)
    
    status_callback("💾 正在寫入 PPT 檔案...")
    prs.save(output_path)
