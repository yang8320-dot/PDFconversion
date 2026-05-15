import os
import sys
from PIL import ImageDraw

def get_base_path():
    """ 取得程式執行時的根目錄 (支援 PyInstaller 打包的 sys._MEIPASS 暫存目錄) """
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS # 打包成單一執行檔時，檔案會被解壓縮到這個暫存目錄
    else:
        return os.path.abspath(".")

def get_poppler_path():
    return os.path.join(get_base_path(), "Library", "poppler_bin")

def get_model_path():
    return os.path.join(get_base_path(), "Library", "model")

def apply_watermark_removal(img):
    """ 去除 NotebookLM 右下角浮水印：採樣背景色並進行智慧覆蓋 """
    draw = ImageDraw.Draw(img)
    width, height = img.size
    
    # 鎖定 NotebookLM 浮水印位置 (大約在右下角 25% 寬度, 8% 高度的範圍內)
    wm_w = int(width * 0.25)
    wm_h = int(height * 0.08)
    box = [width - wm_w, height - wm_h, width, height]
    
    # 在浮水印正上方不遠處，採集原本的「背景底色」
    sample_x = width - wm_w + 10
    sample_y = height - wm_h - 15
    try:
        bg_color = img.getpixel((sample_x, sample_y))
    except:
        bg_color = (255, 255, 255) # 萬一超出邊界，預設白色
        
    # 用採集到的底色畫一個無邊框矩形，完美覆蓋浮水印
    draw.rectangle(box, fill=bg_color, outline=bg_color)
    return img
