import os
import sys

def get_base_path():
    """ 取得程式執行時的根目錄 """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath(".")

def get_poppler_path():
    """ 取得 Poppler 執行檔路徑 """
    return os.path.join(get_base_path(), "Library", "poppler_bin")

def get_model_path():
    """ 取得 EasyOCR 離線模型路徑，解決 407 Proxy 問題 """
    return os.path.join(get_base_path(), "Library", "model")
