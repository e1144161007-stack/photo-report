import os
import re
import docx
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import streamlit as st
from PIL import Image

def natural_sort_key(s):
    """自然排序邏輯：確保 1, 2, 10 與 2F-A, 2F-B 能按正確順序由小到大排列"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

# 頁面標題設定
st.set_page_config(page_title="施工照片紀錄表生成器", layout="centered")
st.title("📷 施工照片紀錄表生成器 (永久免費版)")

# 表頭輸入區
col1, col2 = st.columns(2)
with col1:
    company_name = st.text_input("🏢 公司名稱", value="漢振機電股份有限公司")
    check_date = st.text_input("📅 檢查日期", value="1150725")
with col2:
    project_name = st.text_input("🏗️ 工程名稱", value="龜山區精忠段新建工程")
    item_type = st.text_input("📌 工程項目", value="滿水試驗")

# 照片上傳區
uploaded_files = st.file_uploader(
    "📁 請選擇施工照片 (支援多選/分批選取，系統會自動依檔名由小到大排序)",
    type=["jpg", "png", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"已成功選擇 {len(uploaded_files)} 張照片！")

# 轉換與下載按鈕
if uploaded_files and st.button("🚀 開始自動排序並生成 Word 檔", type="primary"):
    # 自然排序
    sorted_files = sorted(uploaded_files, key=lambda x: natural_sort_key(x.name))
    
    doc = docx.Document()
    
    # A4 直式設定 (210 x 297 mm)
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    
    # 精確設定頁邊界，完美放滿頁面
    section.top_margin = Inches(0.40)
    section.bottom_margin = Inches(0.40)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)
    
    photos_per_page = 6
    
    for i in range(0, len(sorted_files), photos_per_page):
        batch = sorted_files[i:i + photos_per_page]
        
        # 1. 頁首大標題
        title_p = doc.add_paragraph()
        title_p.paragraph_format.space_before = Pt(0)
        title_p.paragraph_format.space_after = Pt(2)
        title_run = title_p.add_run(f"{company_name} - 施工照片紀錄表")
        title_run.font.size = Pt(14)
        title_run.bold = True
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 2. 表頭單欄資訊表
        info_table = doc.add_table(rows=1, cols=1)
        info_table.style = 'Table Grid'
        info_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_table.autofit = False
        info_table.columns[0].width = Inches(7.57)
        
        p0 = info_table.cell(0, 0).paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(f"工程名稱：{project_name}  │  項目：{item_type}  │  日期：{check_date} (第 {i // photos_per_page + 1} 頁)")
        r0.font.size = Pt(10.5)
        r0.bold = True
        
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
        
        # 3. 建立原生 6列 x 2欄 主表格 (徹底消除嵌套表格造成的跑版與空白)
        main_table = doc.add_table(rows=6, cols=2)
        main_table.style = 'Table Grid'
        main_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        main_table.autofit = False
        
        main_table.columns[0].width = Inches(3.78)
        main_table.columns[1].width = Inches(3.78)
        
        # 設定列高：照片列固定 2.85 吋，說明列固定 0.38 吋
        for r_idx in range(6):
            row = main_table.rows[r_idx]
            if r_idx % 2 == 0:
                row.height = Inches(2.85)
            else:
                row.height = Inches(0.38)
        
        for idx, photo_file in enumerate(batch):
            filename = photo_file.name
            raw_title = os.path.splitext(filename)[0]
            
            col_idx = idx % 2
            photo_row_idx = (idx // 2) * 2
            desc_row_idx = photo_row_idx + 1
            
            # --- 上方照片格 (位置 ❷) ---
            cell_photo = main_table.cell(photo_row_idx, col_idx)
            cell_photo.width = Inches(3.78)
            cell_photo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            p_img = cell_photo.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(0)
            p_img.paragraph_format.space_after = Pt(0)
            
            # 讀取圖片尺寸並等比例放滿照片區
            photo_file.seek(0)
            try:
                img = Image.open(photo_file)
                w, h = img.size
                photo_file.seek(0)
            except:
                w, h = 4, 3
            
            max_w = 3.65
            max_h = 2.75
            
            run_img = p_img.add_run()
            if (w / h) > (max_w / max_h):
                run_img.add_picture(photo_file, width=Inches(max_w))
            else:
                run_img.add_picture(photo_file, height=Inches(max_h))
            
            # --- 下方說明格 (位置 ❶) ---
            cell_desc = main_table.cell(desc_row_idx, col_idx)
            cell_desc.width = Inches(3.78)
            cell_desc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            p_desc = cell_desc.paragraphs[0]
            p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_desc.paragraph_format.space_before = Pt(0)
            p_desc.paragraph_format.space_after = Pt(0)
            
            r_desc = p_desc.add_run(f"說明：{raw_title}")
            r_desc.font.size = Pt(9.5)
            r_desc.font.bold = True
            
        if i + photos_per_page < len(sorted_files):
            doc.add_page_break()
            
    output_path = "施工照片紀錄表.docx"
    doc.save(output_path)
    
    with open(output_path, "rb") as f:
        st.download_button(
            "📥 下載排版完成的 Word 紀錄表 (.docx)",
            f,
            file_name="施工照片紀錄表.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
