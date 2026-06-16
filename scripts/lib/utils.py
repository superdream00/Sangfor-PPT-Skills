"""
深信服 PPT 生成器基础工具模块 (utils.py)
提供颜色转换、双语字体设置、段落间距调整等底层辅助工具。
"""
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree
from scripts.lib.constants import SangforFonts, SangforColors

def hex_to_rgb(hex_color: str) -> RGBColor:
    """将 hex 颜色字符串转为 RGBColor 对象"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)

def set_font(run, font_name='微软雅黑', size=Pt(12), color='#0E0E0E', 
             bold=False, italic=False):
    """设置 run 的完整字体属性（包括中文和英文字体）"""
    font = run.font
    font.name = font_name
    font.size = size
    font.color.rgb = hex_to_rgb(color)
    font.bold = bold
    font.italic = italic
    
    # 通过 XML 设置中文字体 <a:ea typeface="微软雅黑"/>
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
    if ea is None:
        ea = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
    ea.set('typeface', font_name)
    
    # 同时设置拉丁字体
    latin = rpr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
    if latin is None:
        latin = etree.SubElement(rpr, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
    latin.set('typeface', font_name)

def set_paragraph_spacing(paragraph, line_spacing=1.5, space_before=Pt(0), space_after=Pt(0)):
    """设置段落行间距与段前后间距"""
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pPr = paragraph._p.get_or_add_pPr()

    # OOXML 要求 lnSpc 必须是 pPr 的第一个子元素，且只能有一个。
    # 先移除已有的 lnSpc（避免重复调用产生多个节点），再 insert 到首位。
    for existing in pPr.findall(f'{{{A_NS}}}lnSpc'):
        pPr.remove(existing)

    lnSpc = etree.SubElement(pPr, f'{{{A_NS}}}lnSpc')
    spcPct = etree.SubElement(lnSpc, f'{{{A_NS}}}spcPct')
    spcPct.set('val', str(int(line_spacing * 100000)))
    pPr.remove(lnSpc)
    pPr.insert(0, lnSpc)

def _set_cell_color(cell, hex_color):
    """设置表格单元格背景色"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
    srgbClr = etree.SubElement(solidFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
    srgbClr.set('val', hex_color.lstrip('#'))

def _fill_shape_text(shape, text_data, default_font_size=Pt(10), default_color='#0E0E0E', default_bold=False, alignment=PP_ALIGN.LEFT):
    """向形状中填充多行格式化文字（画布组件使用）"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.1)
    tf.margin_bottom = Cm(0.1)
    
    if not text_data:
        return
        
    if isinstance(text_data, str):
        lines = text_data.split('\n')
    else:
        lines = text_data
        
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = alignment
        set_paragraph_spacing(p, 1.2)
        
        if isinstance(line, str):
            run = p.add_run()
            run.text = line
            set_font(run, font_name=SangforFonts.CHINESE, size=default_font_size, color=default_color, bold=default_bold)
        elif isinstance(line, dict):
            run = p.add_run()
            run.text = line.get('text', '')
            font_size = Pt(line.get('font_size_pt', line.get('size', 10)))
            font_color = line.get('font_color', line.get('color', default_color))
            bold = line.get('bold', default_bold)
            set_font(run, font_name=SangforFonts.CHINESE, size=font_size, color=font_color, bold=bold)
