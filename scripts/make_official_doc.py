# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_official_doc(out_path):
    doc = docx.Document()
    
    # Page setup
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.2)
    sec.header_distance = Cm(1.2)
    sec.footer_distance = Cm(1.2)
    
    # Helper for font
    def set_run_font(run, name_cjk="標楷體", name_ascii="Times New Roman", size_pt=12, bold=False, color_rgb=None):
        run.font.name = name_ascii
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        if color_rgb:
            run.font.color.rgb = color_rgb
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:eastAsia'), name_cjk)
        rFonts.set(qn('w:ascii'), name_ascii)
        rFonts.set(qn('w:hAnsi'), name_ascii)

    # 1. Top row: 檔號 / 保存年限 (Right aligned)
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_top.paragraph_format.space_before = Pt(0)
    p_top.paragraph_format.space_after = Pt(2)
    p_top.paragraph_format.line_spacing = Pt(14)
    
    r_top = p_top.add_run("檔　　號：\n保存年限：")
    set_run_font(r_top, size_pt=10)

    # 2. Agency Title: 臺南市柳營區重溪國民小學 函 (Center)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(8)
    r_title = p_title.add_run("臺南市柳營區重溪國民小學　函")
    set_run_font(r_title, size_pt=22, bold=True)

    # 3. Contact Info (Right aligned box / paragraph)
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    p_contact.paragraph_format.line_spacing = Pt(13)
    contact_text = (
        "地　　址：736004 臺南市柳營區義士路二段2號\n"
        "承 辦 人：謝璨鴻\n"
        "電　　話：(06)623-1824\n"
        "傳　　真：(06)623-0725\n"
        "電子信箱：tonlyteacher@gmail.com"
    )
    r_contact = p_contact.add_run(contact_text)
    set_run_font(r_contact, size_pt=9.5)

    # 4. Recipient Address & Recipient Name
    p_recip_addr = doc.add_paragraph()
    p_recip_addr.paragraph_format.space_before = Pt(4)
    p_recip_addr.paragraph_format.space_after = Pt(2)
    p_recip_addr.paragraph_format.line_spacing = Pt(14)
    r_recip_addr = p_recip_addr.add_run("702008\n臺南市南區樹林街2段253號9樓")
    set_run_font(r_recip_addr, size_pt=11)

    p_recip = doc.add_paragraph()
    p_recip.paragraph_format.space_before = Pt(2)
    p_recip.paragraph_format.space_after = Pt(4)
    r_recip_label = p_recip.add_run("受文者：")
    set_run_font(r_recip_label, size_pt=12.5, bold=True)
    r_recip_val = p_recip.add_run("財團法人吳修齊紀念先嚴吳克讀公先慈吳陳氏勤娘文教公益基金會")
    set_run_font(r_recip_val, size_pt=12.5, bold=True)

    # 5. Metadata block (發文日期、發文字號、速別、密等、附件)
    meta_lines = [
        ("發文日期：", "中華民國116年1月2日"),
        ("發文字號：", "重溪小教字第1160000001號"),
        ("速　　別：", "普通件"),
        ("密等及解密條件或保密期限：", "普通"),
        ("附　　件：", "如文（第一期款領據乙紙、農會公庫存摺封面影本乙份）")
    ]
    for label, val in meta_lines:
        p_m = doc.add_paragraph()
        p_m.paragraph_format.space_before = Pt(0)
        p_m.paragraph_format.space_after = Pt(1)
        p_m.paragraph_format.line_spacing = Pt(14)
        r_l = p_m.add_run(label)
        set_run_font(r_l, size_pt=11)
        r_v = p_m.add_run(val)
        set_run_font(r_v, size_pt=11)

    # 6. 主旨
    p_zhuzhi = doc.add_paragraph()
    p_zhuzhi.paragraph_format.space_before = Pt(8)
    p_zhuzhi.paragraph_format.space_after = Pt(4)
    p_zhuzhi.paragraph_format.line_spacing = Pt(18)
    # Hanging indent for 主旨
    p_zhuzhi.paragraph_format.left_indent = Cm(1.2)
    p_zhuzhi.paragraph_format.first_line_indent = Cm(-1.2)
    
    r_zz_l = p_zhuzhi.add_run("主旨：")
    set_run_font(r_zz_l, size_pt=12.5, bold=True)
    r_zz_v = p_zhuzhi.add_run("檢送本校辦理116年「吳修齊獎」課後輔導計畫第一期款領據及匯款公庫帳戶資料各乙份，請惠予辦理撥款，請查照。")
    set_run_font(r_zz_v, size_pt=12.5, bold=True)

    # 7. 說明 Header
    p_sm = doc.add_paragraph()
    p_sm.paragraph_format.space_before = Pt(4)
    p_sm.paragraph_format.space_after = Pt(2)
    p_sm.paragraph_format.line_spacing = Pt(16)
    r_sm = p_sm.add_run("說明：")
    set_run_font(r_sm, size_pt=12, bold=True)

    # 8. 說明條項
    instructions = [
        "一、復貴會115年11月19日(115)齊字第013號函。",
        "二、本校提報「115學年度補助課後輔導計畫」承蒙貴會第十六屆第四次董事會議審查通過，核定116年「吳修齊獎」課後輔導計畫補助新臺幣40萬4,890元整，至紉公誼，感荷無極。",
        "三、依貴會來函規定，旨揭計畫第一期款（70%，計新臺幣28萬3,423元整）預定於116年1月撥付。本校已開立116年1月第一期款領據乙紙（如附件一），並提供柳營區農會代收款專戶匯款資料如下：\n"
        "　　(一)金融機構：臺南市柳營區農會（金融機構代號：618-0092）\n"
        "　　(二)公庫戶名：臺南市柳營區重溪國民小學代收款專戶\n"
        "　　(三)帳　　號：00092160094086\n"
        "　　(四)統一編號：72046589\n"
        "　　(五)檢附柳營區農會存摺封面影本乙份（如附件二）。",
        "四、本校定當恪遵貴會計畫規範與各項審查意見，落實專款專用、深耕鄉村教學，並依規於116年9月15日前繳交期中成果報告；隨文檢附上開請款文件，敬請惠予惠撥第一期款，俾利課照教學順利推展。"
    ]

    for item in instructions:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.space_before = Pt(2)
        p_item.paragraph_format.space_after = Pt(2)
        p_item.paragraph_format.line_spacing = Pt(16)
        p_item.paragraph_format.left_indent = Cm(1.2)
        p_item.paragraph_format.first_line_indent = Cm(-1.2)
        r_item = p_item.add_run(item)
        set_run_font(r_item, size_pt=11.5)

    # 9. 正本 / 副本
    p_zb = doc.add_paragraph()
    p_zb.paragraph_format.space_before = Pt(6)
    p_zb.paragraph_format.space_after = Pt(1)
    p_zb.paragraph_format.line_spacing = Pt(14)
    r_zb_l = p_zb.add_run("正本：")
    set_run_font(r_zb_l, size_pt=11, bold=True)
    r_zb_v = p_zb.add_run("財團法人吳修齊紀念先嚴吳克讀公先慈吳陳氏勤娘文教公益基金會")
    set_run_font(r_zb_v, size_pt=11)

    p_fb = doc.add_paragraph()
    p_fb.paragraph_format.space_before = Pt(0)
    p_fb.paragraph_format.space_after = Pt(8)
    p_fb.paragraph_format.line_spacing = Pt(14)
    r_fb_l = p_fb.add_run("副本：")
    set_run_font(r_fb_l, size_pt=11, bold=True)
    r_fb_v = p_fb.add_run("本校教導處、本校總務處、本校會計室")
    set_run_font(r_fb_v, size_pt=11)

    # 10. Principal Signature block (Centered / Bottom)
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.space_before = Pt(12)
    p_sign.paragraph_format.space_after = Pt(0)
    r_sign = p_sign.add_run("校長　吳　建　邦")
    set_run_font(r_sign, size_pt=16, bold=True)

    # 11. Footer: 第 1 頁 共 1 頁
    ftr = sec.footer
    p_ftr = ftr.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ftr = p_ftr.add_run("第 1 頁　共 1 頁")
    set_run_font(r_ftr, size_pt=9)

    doc.save(out_path)
    print("Official doc created successfully at:", out_path)

if __name__ == '__main__':
    create_official_doc("output/1160102_吳修齊基金會_重溪國小第一期款函復公文_v2.docx")
