import os
import sys
import platform
import subprocess
from PIL import ImageDraw

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.abspath(".")

def get_poppler_path():
    return os.path.join(get_base_path(), "Library", "poppler_bin")

def get_model_path():
    return os.path.join(get_base_path(), "Library", "models")

def check_poppler_exists():
    return os.path.exists(get_poppler_path())

def open_file_or_folder(path):
    """完成後自動打開檔案所在資料夾"""
    try:
        target = path if os.path.isdir(path) else os.path.dirname(path)
        if platform.system() == "Windows":
            os.startfile(target)
        elif platform.system() == "Darwin":
            subprocess.call(["open", target])
        else:
            subprocess.call(["xdg-open", target])
    except: pass

def format_size(size_in_bytes):
    """格式化檔案大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_in_bytes < 1024: return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024
    return f"{size_in_bytes:.2f} TB"

def apply_watermark_removal(img, position="右下角"):
    """ 去除浮水印：可自訂角落位置，遮蔽範圍為 125% """
    draw = ImageDraw.Draw(img)
    width, height = img.size
    
    mask_w = int(width * 0.15 * 1.25)
    mask_h = int(height * 0.05 * 1.25)
    
    if position == "右下角":
        box = [width - mask_w, height - mask_h, width, height]
        sample_pt = (width - mask_w - 10, height - mask_h - 10)
    elif position == "左下角":
        box = [0, height - mask_h, mask_w, height]
        sample_pt = (mask_w + 10, height - mask_h - 10)
    elif position == "右上角":
        box = [width - mask_w, 0, width, mask_h]
        sample_pt = (width - mask_w - 10, mask_h + 10)
    else: # 左上角
        box = [0, 0, mask_w, mask_h]
        sample_pt = (mask_w + 10, mask_h + 10)
    
    try: bg_color = img.getpixel(sample_pt)
    except: bg_color = (255, 255, 255) 
        
    draw.rectangle(box, fill=bg_color, outline=bg_color)
    return img
