import os
import sys

def get_poppler_path():
    """ 取得 Poppler 執行檔路徑，相容開發與打包環境 """
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
        return os.path.join(base_dir, "Library", "poppler_bin")
    else:
        return os.path.join(os.path.abspath("."), "Library", "poppler_bin")
