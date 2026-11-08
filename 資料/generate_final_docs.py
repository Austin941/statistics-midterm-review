# -*- coding: utf-8 -*-
"""
商用統計學 期中考精選習題與解析 (第 2 至 5 章共 40 題)
雙軌輸出生成器：生成排版精美的 DOCX 並自動調用 Word COM 轉為 PDF
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=180, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_left_border_only(cell, color_hex="1B365D", sz="24"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def format_cell_text(cell, text, bold=False, color_rgb=(0x22, 0x22, 0x22), size_pt=9.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Microsoft JhengHei"
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    return p

def create_table_grid(doc, headers, data_rows, col_widths=None):
    table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        set_cell_background(hdr_cells[i], "1F4E79")
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=100, right=100)
        format_cell_text(hdr_cells[i], title, bold=True, color_rgb=(0xFF, 0xFF, 0xFF), size_pt=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        if col_widths and i < len(col_widths):
            hdr_cells[i].width = col_widths[i]
            
    # Data rows
    for r_idx, row_data in enumerate(data_rows):
        row_cells = table.rows[r_idx + 1].cells
        bg_col = "FFFFFF" if r_idx % 2 == 0 else "F9FAFC"
        for c_idx, val in enumerate(row_data):
            set_cell_background(row_cells[c_idx], bg_col)
            set_cell_margins(row_cells[c_idx], top=60, bottom=60, left=100, right=100)
            align = WD_ALIGN_PARAGRAPH.RIGHT if (c_idx > 0 and any(char.isdigit() for char in str(val)) and not any(k in str(val) for k in ["小於", "以上", "至"])) else WD_ALIGN_PARAGRAPH.LEFT
            if c_idx == 0:
                align = WD_ALIGN_PARAGRAPH.CENTER
            format_cell_text(row_cells[c_idx], str(val), bold=False, color_rgb=(0x33, 0x33, 0x33), size_pt=9, align=align)
            if col_widths and c_idx < len(col_widths):
                row_cells[c_idx].width = col_widths[c_idx]
                
    # Table borders
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="1F4E79"/>
            <w:left w:val="none"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="1F4E79"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def add_callout(doc, title, text_paras, bg_color="F4F7FB", border_color="1B365D"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.8)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_left_border_only(cell, border_color, sz="24")
    set_cell_margins(cell, top=100, bottom=100, left=150, right=120)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run_t = p.add_run(f"【{title}】\n")
    run_t.bold = True
    run_t.font.name = "Microsoft JhengHei"
    run_t.font.size = Pt(10)
    run_t.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    for i, t in enumerate(text_paras):
        if i > 0 or len(text_paras) > 1:
            p2 = cell.add_paragraph()
            p2.paragraph_format.space_before = Pt(1)
            p2.paragraph_format.space_after = Pt(1)
            p2.paragraph_format.line_spacing = 1.15
            run = p2.add_run(t)
        else:
            run = p.add_run(t)
        run.font.name = "Microsoft JhengHei"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def add_solution_block(doc, methodology, steps, answer):
    # Methodology
    p_m = doc.add_paragraph()
    p_m.paragraph_format.space_before = Pt(3)
    p_m.paragraph_format.space_after = Pt(1)
    p_m.paragraph_format.line_spacing = 1.15
    r_mt = p_m.add_run("◆ 核心觀念與統計公式：\n")
    r_mt.bold = True
    r_mt.font.name = "Microsoft JhengHei"
    r_mt.font.size = Pt(9.5)
    r_mt.font.color.rgb = RGBColor(0x0F, 0x4C, 0x81)
    
    r_mc = p_m.add_run(methodology)
    r_mc.font.name = "Microsoft JhengHei"
    r_mc.font.size = Pt(9.5)
    r_mc.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    # Steps
    p_s = doc.add_paragraph()
    p_s.paragraph_format.space_before = Pt(3)
    p_s.paragraph_format.space_after = Pt(1)
    p_s.paragraph_format.line_spacing = 1.15
    r_st = p_s.add_run("◆ 詳細解題步驟與推導：\n")
    r_st.bold = True
    r_st.font.name = "Microsoft JhengHei"
    r_st.font.size = Pt(9.5)
    r_st.font.color.rgb = RGBColor(0x0F, 0x4C, 0x81)
    
    r_sc = p_s.add_run(steps)
    r_sc.font.name = "Microsoft JhengHei"
    r_sc.font.size = Pt(9.5)
    r_sc.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    
    # Answer box
    tbl_ans = doc.add_table(rows=1, cols=1)
    tbl_ans.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ans.autofit = False
    tbl_ans.columns[0].width = Inches(6.8)
    cell_ans = tbl_ans.cell(0, 0)
    set_cell_background(cell_ans, "EBF3FA")
    set_cell_left_border_only(cell_ans, "2B6CB0", sz="18")
    set_cell_margins(cell_ans, top=80, bottom=80, left=150, right=120)
    
    p_a = cell_ans.paragraphs[0]
    p_a.paragraph_format.space_before = Pt(0)
    p_a.paragraph_format.space_after = Pt(0)
    p_a.paragraph_format.line_spacing = 1.15
    r_at = p_a.add_run("★ 標準解答與統計結論 (Final Answer)：\n")
    r_at.bold = True
    r_at.font.name = "Microsoft JhengHei"
    r_at.font.size = Pt(9.5)
    r_at.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    r_ac = p_a.add_run(answer)
    r_ac.bold = True
    r_ac.font.name = "Microsoft JhengHei"
    r_ac.font.size = Pt(9.5)
    r_ac.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)
    
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(6)

print("Setup completed.")
