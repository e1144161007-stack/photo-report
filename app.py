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

def natural_sort_key(s):
    """自然排序邏輯：確保 1, 2, 10 與 2F-A, 2F-B 能按正確順序由小到大排列"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def set_cell_border_top_dashed(cell):
    """設定儲存格頂部虛線分隔線"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="dashed" w:sz="4" w:space="0" w:color="auto"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

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
    
    # A4 橫式設定
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    
    # 邊界設定 (確保 3x2 共 6 張卡片完美裝入一頁)
    section.top_margin = Inches(0.3)
    section.bottom_margin = Inches(0.3)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)
    
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
        
        # 2. 表頭雙欄資訊表
        info_table = doc.add_table(rows=1, cols=2)
        info_table.style = 'Table Grid'
        info_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_table.autofit = False
        
        info_table.columns[0].width = Inches(7.2)
        info_table.columns[1].width = Inches(3.6)
        
        p0 = info_table.cell(0, 0).paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(f"工程名稱：{project_name}")
        r0.font.size = Pt(12)
        r0.bold = True
        
        p1 = info_table.cell(0, 1).paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(f"檢查日期：{check_date} (第 {i // photos_per_page + 1} 頁)")
        r1.font.size = Pt(12)
        r1.bold = True
        
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
        
        # 3. 建立主表格 (A4 橫式 3x2 6 張卡片)
        main_table = doc.add_table(rows=3, cols=2)
        main_table.style = 'Table Grid'
        main_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        main_table.autofit = False
        
        for idx, photo_file in enumerate(batch):
            filename = photo_file.name
            raw_title = os.path.splitext(filename)[0]
            
            row_idx = idx // 2
            col_idx = idx % 2
            
            parent_cell = main_table.cell(row_idx, col_idx)
            
            # 單欄上下 2 列結構 (上照片 80%、下說明 20%)
            sub_table = parent_cell.add_table(rows=2, cols=1)
            sub_table.style = 'Table Grid'
            sub_table.autofit = False
            sub_table.columns[0].width = Inches(5.3)
            
            # 位置 ❷：上方照片欄 (80% 高度比例)
            cell_photo = sub_table.cell(0, 0)
            cell_photo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_img = cell_photo.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(0)
            p_img.paragraph_format.space_after = Pt(0)
            
            run_img = p_img.add_run()
            # 控制照片高度以精準維持 8:2 比例
            run_img.add_picture(photo_file, height=Inches(1.65))
            
            # 位置 ❶：下方說明欄 (20% 高度比例)
            cell_desc = sub_table.cell(1, 0)
            set_cell_border_top_dashed(cell_desc)
            cell_desc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            p_desc = cell_desc.paragraphs[0]
            p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_desc.paragraph_format.space_before = Pt(2)
            p_desc.paragraph_format.space_after = Pt(2)
            
            r_item = p_desc.add_run(f"項目：{item_type}  │  說明：{raw_title}")
            r_item.font.size = Pt(9.5)
            r_item.font.bold = True
            
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
