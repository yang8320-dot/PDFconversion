import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import easyocr
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt

def get_poppler_path():
    """ 
    動態獲取 Poppler 執行檔路徑 
    處理 PyInstaller 打包後的 Library 目錄結構
    """
    if getattr(sys, 'frozen', False):
        # 如果是被打包的執行檔，取得 .exe 所在的目錄
        base_dir = os.path.dirname(sys.executable)
        return os.path.join(base_dir, "Library", "poppler_bin")
    else:
        # 如果是本地 Python 開發環境 (假設 poppler_bin 放在專案下的 Library 資料夾)
        return os.path.join(os.path.abspath("."), "Library", "poppler_bin")

def pdf_to_ppt_with_ocr(pdf_path, ppt_output_path):
    # 支援繁體中文、簡體中文、英文 (數字會自動包含在內)
    languages = ['ch_tra', 'ch_sim', 'en']
    
    print(f"啟動 OCR 引擎，載入模型: {languages}...")
    reader = easyocr.Reader(languages)
    
    # 取得 Poppler 路徑
    poppler_path = get_poppler_path()
    if not os.path.exists(poppler_path):
        print(f"警告: 找不到 Poppler 路徑 {poppler_path}，若系統已安裝可略過此警告。")
    
    print(f"正在將 PDF 轉換為圖片，這可能需要一點時間...")
    # 若執行時報錯，請確保 Poppler 已經正確下載並位於正確路徑
    try:
        pages = convert_from_path(pdf_path, dpi=200, poppler_path=poppler_path)
    except Exception as e:
        messagebox.showerror("錯誤", f"PDF 讀取失敗，請確認 Poppler 是否存在：\n{e}")
        return

    prs = Presentation()
    total_pages = len(pages)
    
    for i, page_img in enumerate(pages):
        print(f"正在辨識第 {i+1} / {total_pages} 頁...")
        temp_img = f"temp_page_{i}.jpg"
        page_img.save(temp_img, 'JPEG')
        
        # 執行 OCR (detail=0 僅回傳文字陣列)
        result = reader.readtext(temp_img, detail=0)
        extracted_text = "\n".join(result)
        
        # 建立 PPT 頁面 (Layout 6 通常為空白頁面)
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        # 設定文字方塊大小與位置
        left = top = Inches(0.5)
        width = prs.slide_width - Inches(1)
        height = prs.slide_height - Inches(1)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        
        # 寫入文字並設定樣式
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.add_paragraph()
        p.text = extracted_text
        p.font.size = Pt(14)
        
        # 清理暫存圖片
        os.remove(temp_img)
    
    prs.save(ppt_output_path)
    print(f"\n成功！PPT 已儲存至: {ppt_output_path}")
    messagebox.showinfo("完成", f"轉換成功！\n檔案已儲存至：\n{ppt_output_path}")

def main():
    # 隱藏 Tkinter 的主視窗
    root = tk.Tk()
    root.withdraw()

    print("請在彈出的視窗中選擇要轉換的 PDF 檔案...")
    pdf_path = filedialog.askopenfilename(
        title="選擇要轉換的 PDF 檔案",
        filetypes=[("PDF Files", "*.pdf")]
    )
    
    if not pdf_path:
        print("未選擇 PDF 檔案，程式結束。")
        return

    print("請選擇 PPT 儲存位置...")
    ppt_path = filedialog.asksaveasfilename(
        title="儲存 PPT 檔案",
        defaultextension=".pptx",
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    
    if not ppt_path:
        print("未指定儲存位置，程式結束。")
        return

    # 開始執行轉換
    pdf_to_ppt_with_ocr(pdf_path, ppt_path)

if __name__ == "__main__":
    main()
