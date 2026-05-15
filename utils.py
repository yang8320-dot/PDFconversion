import os
import sys
from PIL import ImageDraw

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

def get_poppler_path():
    return os.path.join(get_base_path(), "Library", "poppler_bin")

def get_model_path():
    return os.path.join(get_base_path(), "Library", "model")

def apply_watermark_removal(img):
    """ 去除 NotebookLM 浮水印：精準鎖定右下角，遮蔽範圍擴大至 125% """
    draw = ImageDraw.Draw(img)
    width, height = img.size
    
    # 預估 NotebookLM 浮水印的預設大小 (約寬度 15%, 高度 5%)
    wm_w = int(width * 0.15)
    wm_h = int(height * 0.05)
    
    # 依使用者需求，將遮蔽範圍放大至 125%
    mask_w = int(wm_w * 1.25)
    mask_h = int(wm_h * 1.25)
    
    box = [width - mask_w, height - mask_h, width, height]
    
    # 在浮水印遮蔽區塊的「左上方」一點點採集背景顏色
    sample_x = width - mask_w - 10
    sample_y = height - mask_h - 10
    try:
        bg_color = img.getpixel((sample_x, sample_y))
    except:
        bg_color = (255, 255, 255) # 萬一取樣失敗預設為白色
        
    # 用採集到的底色進行覆蓋填滿
    draw.rectangle(box, fill=bg_color, outline=bg_color)
    return img
