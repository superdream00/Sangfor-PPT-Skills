"""
深信服 PPT 生成工具库 (utils.py)
提供颜色、字体、形状、图表等工具函数，遵循深信服品牌视觉规范
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from lxml import etree
import os

# ============================================================
# XML 命名空间
# ============================================================
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


# ============================================================
# 深信服品牌颜色常量
# ============================================================
class SangforColors:
    """深信服品牌色彩体系"""
    # 主色系 (蓝色)
    BLUE_PRIMARY = '#006CD9'       # 深信服蓝（主色）
    BLUE_DARK = '#003592'          # 深蓝（标题装饰竖条）
    BLUE_DEEP = '#00479D'          # 深蓝辅助
    BLUE_LIGHT1 = '#0587F5'        # 品牌浅蓝1
    BLUE_LIGHT2 = '#3BA5FF'        # 品牌浅蓝2
    BLUE_LIGHT3 = '#51C0F9'        # 品牌浅蓝3（渐变浅端）
    BLUE_HIGHLIGHT = '#0185FF'     # 高亮蓝
    BLUE_SOFT = '#65B5FF'          # 柔和蓝
    BLUE_PALE = '#80D8FD'          # 淡蓝（渐变端点）
    
    # 强调色 (绿色)
    GREEN_PRIMARY = '#53C800'      # 深信服绿（强调色）
    GREEN_VARIANT1 = '#6FBA2C'     # 绿色变体1（引号装饰）
    GREEN_VARIANT2 = '#6CCD24'     # 绿色变体2（渐变端点）
    GREEN_LIGHT = '#9BF48F'        # 浅绿（渐变浅端）
    GREEN_SOFT = '#CCE28E'         # 柔和绿（渐变端点）
    
    # 中性色
    TEXT_PRIMARY = '#0E0E0E'       # 正文主色（使用240次）
    TEXT_SECONDARY = '#404040'     # 辅助灰
    TEXT_TERTIARY = '#595959'      # 说明灰
    TEXT_DARK = '#0D0D0D'          # 深黑
    TEXT_MEDIUM = '#262626'        # 中灰
    BG_LIGHT_GRAY = '#ECECEC'      # 浅灰背景
    BORDER_GRAY = '#E7E6E6'        # 边框灰
    BG_SOFT_GRAY = '#E8E8E9'       # 柔和灰背景
    GRAY_MEDIUM = '#808080'        # 中灰
    WHITE = '#FFFFFF'
    BLACK = '#000000'
    
    # 图表推荐色序
    CHART_COLORS = [
        '#006CD9', '#53C800', '#0587F5', '#3BA5FF', 
        '#6FBA2C', '#51C0F9', '#0185FF', '#65B5FF'
    ]
    
    # 表格配色
    TABLE_HEADER_BG = '#006CD9'
    TABLE_HEADER_TEXT = '#FFFFFF'
    TABLE_ODD_ROW = '#F2F7FC'
    TABLE_EVEN_ROW = '#FFFFFF'
    TABLE_BORDER = '#E7E6E6'


# ============================================================
# 深信服字体规范
# ============================================================
class SangforFonts:
    """深信服字体规范"""
    CHINESE = '微软雅黑'
    ENGLISH = '微软雅黑'
    
    # 各场景字号
    COVER_TITLE = Pt(26)       # 封面大标题
    COVER_SUBTITLE = Pt(14)    # 封面小标题
    PAGE_TITLE = Pt(22)        # 页面标题
    SUBTITLE = Pt(16)          # 副标题
    BODY = Pt(12)              # 正文
    BODY_LARGE = Pt(14)        # 大号正文
    SMALL = Pt(10)             # 小号文字
    CAPTION = Pt(9)            # 脚注/说明
    TABLE_HEADER = Pt(12)      # 表格标题行
    TABLE_BODY = Pt(11)        # 表格正文
    CHART_LABEL = Pt(10)       # 图表标签
    NUMBER_HIGHLIGHT = Pt(44)  # 高亮数字
    SECTION_TITLE = Pt(44)     # 章节标题


# ============================================================
# 颜色工具函数
# ============================================================
def hex_to_rgb(hex_color: str) -> RGBColor:
    """将 hex 颜色字符串转为 RGBColor 对象"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


# ============================================================
# 字体设置函数
# ============================================================
def set_font(run, font_name='微软雅黑', size=Pt(12), color='#0E0E0E', 
             bold=False, italic=False):
    """设置 run 的完整字体属性（包括中文和英文字体）
    
    Args:
        run: pptx Run 对象
        font_name: 字体名称（同时用于中英文）
        size: 字号 (Pt 对象)
        color: hex 颜色值
        bold: 是否加粗
        italic: 是否斜体
    """
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
    """设置段落间距
    
    Args:
        paragraph: pptx Paragraph 对象
        line_spacing: 行距倍数
        space_before: 段前间距
        space_after: 段后间距
    """
    pPr = paragraph._p.get_or_add_pPr()
    lnSpc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lnSpc')
    spcPct = etree.SubElement(lnSpc, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPct')
    spcPct.set('val', str(int(line_spacing * 100000)))


# ============================================================
# 形状创建函数
# ============================================================
def add_textbox(slide, left_cm, top_cm, width_cm, height_cm, text,
                font_name='微软雅黑', font_size=Pt(12), font_color='#0E0E0E',
                bold=False, alignment=PP_ALIGN.LEFT, line_spacing=1.5,
                vertical_anchor=MSO_ANCHOR.TOP):
    """添加格式化文本框
    
    Args:
        slide: 幻灯片对象
        left_cm, top_cm, width_cm, height_cm: 位置和尺寸（厘米）
        text: 文本内容
        font_name: 字体名称
        font_size: 字号
        font_color: 文字颜色
        bold: 是否加粗
        alignment: 对齐方式
        line_spacing: 行距
        vertical_anchor: 垂直对齐
    Returns:
        添加的文本框 Shape 对象
    """
    txBox = slide.shapes.add_textbox(
        Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    set_font(run, font_name, font_size, font_color, bold)
    
    if line_spacing != 1.0:
        set_paragraph_spacing(p, line_spacing)
    
    return txBox


def add_multiline_textbox(slide, left_cm, top_cm, width_cm, height_cm, 
                          lines, default_font='微软雅黑', default_size=Pt(12),
                          default_color='#0E0E0E', alignment=PP_ALIGN.LEFT,
                          line_spacing=1.5):
    """添加多行/多格式文本框
    
    Args:
        lines: 文本行列表，每行可以是字符串或字典：
            - 字符串: 使用默认格式
            - 字典: {'text': '...', 'color': '#xxx', 'size': Pt(x), 'bold': True}
    Returns:
        添加的文本框 Shape 对象
    """
    txBox = slide.shapes.add_textbox(
        Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        p.alignment = alignment
        
        if isinstance(line, str):
            run = p.add_run()
            run.text = line
            set_font(run, default_font, default_size, default_color)
        elif isinstance(line, dict):
            run = p.add_run()
            run.text = line.get('text', '')
            set_font(
                run,
                line.get('font', default_font),
                line.get('size', default_size),
                line.get('color', default_color),
                line.get('bold', False)
            )
        
        if line_spacing != 1.0:
            set_paragraph_spacing(p, line_spacing)
    
    return txBox


def add_styled_rectangle(slide, left_cm, top_cm, width_cm, height_cm,
                         fill_color='#006CD9', line_color=None, 
                         line_width=Pt(0), corner_radius=None):
    """添加样式化矩形
    
    Args:
        slide: 幻灯片对象
        fill_color: 填充颜色 hex
        line_color: 描边颜色 hex（None 则无描边）
        line_width: 描边宽度
        corner_radius: 圆角半径（厘米，None 则无圆角）
    Returns:
        矩形 Shape 对象
    """
    if corner_radius:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
        )
    else:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
        )
    
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
    
    if line_color:
        shape.line.fill.solid()
        shape.line.fill.fore_color.rgb = hex_to_rgb(line_color)
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    
    return shape


def add_styled_circle(slide, left_cm, top_cm, diameter_cm, fill_color='#006CD9'):
    """添加样式化圆形
    
    Args:
        diameter_cm: 直径（厘米）
    Returns:
        圆形 Shape 对象
    """
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Cm(left_cm), Cm(top_cm), Cm(diameter_cm), Cm(diameter_cm)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
    shape.line.fill.background()
    return shape


def add_gradient_fill(shape, angle_deg, stops):
    """为形状添加渐变填充
    
    Args:
        shape: 形状对象
        angle_deg: 渐变角度（度数）
        stops: 渐变色标列表 [(position_pct, hex_color, alpha_pct_optional), ...]
            position_pct: 0-100 的百分比
            hex_color: '#006CD9' 格式的颜色
            alpha_pct: 可选，0-100 的透明度百分比（100=不透明）
    """
    spPr = shape._element.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}spPr')
    if spPr is None:
        spPr = shape._element.spPr
    
    # 移除现有填充
    for child in list(spPr):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('solidFill', 'gradFill', 'noFill', 'pattFill'):
            spPr.remove(child)
    
    # 创建渐变填充 XML
    gradFill = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}gradFill')
    gsLst = etree.SubElement(gradFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}gsLst')
    
    for stop in stops:
        pos = stop[0]
        color = stop[1]
        alpha = stop[2] if len(stop) > 2 else None
        
        gs = etree.SubElement(gsLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}gs')
        gs.set('pos', str(int(pos * 1000)))
        
        srgbClr = etree.SubElement(gs, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
        srgbClr.set('val', color.lstrip('#'))
        
        if alpha is not None and alpha < 100:
            alphaElem = etree.SubElement(srgbClr, '{http://schemas.openxmlformats.org/drawingml/2006/main}alpha')
            alphaElem.set('val', str(int(alpha * 1000)))
    
    lin = etree.SubElement(gradFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}lin')
    lin.set('ang', str(int(angle_deg * 60000)))
    lin.set('scaled', '1')


# ============================================================
# 高级组件函数
# ============================================================
def add_title_area(slide, title_text):
    """添加标准标题区域（左上角蓝色标题+深蓝竖条装饰）
    
    遵循模板设计：
    - 位置: (1.39cm, 0.54cm), 宽22cm, 高1.5cm
    - 文字: #006CD9, 微软雅黑, 22pt, 加粗
    - 左侧有 #003592 深蓝竖条装饰
    
    Returns:
        (竖条Shape, 文本框Shape) 元组
    """
    # 添加深蓝竖条装饰
    bar = add_styled_rectangle(
        slide, 0.8, 0.54, 0.3, 1.5,
        fill_color=SangforColors.BLUE_DARK
    )
    
    # 添加标题文本
    title_box = add_textbox(
        slide, 1.39, 0.54, 22.03, 1.5, title_text,
        font_size=SangforFonts.PAGE_TITLE,
        font_color=SangforColors.BLUE_PRIMARY,
        bold=True
    )
    
    return bar, title_box


def add_bullet_list(slide, left_cm, top_cm, width_cm, items,
                    bullet_color='#53C800', title_color='#006CD9', 
                    text_color='#0E0E0E', item_height_cm=2.5,
                    bullet_size_cm=0.3):
    """添加带圆点标记的列表
    
    Args:
        items: [{'title': '小标题', 'text': '正文内容'}, ...]
        bullet_color: 圆点颜色
        title_color: 小标题颜色
        text_color: 正文颜色
        item_height_cm: 每项高度
        bullet_size_cm: 圆点直径
    """
    shapes = []
    for i, item in enumerate(items):
        y = top_cm + i * item_height_cm
        
        # 添加绿色圆点
        circle = add_styled_circle(
            slide, left_cm, y + 0.4, bullet_size_cm, bullet_color
        )
        shapes.append(circle)
        
        # 添加文本（标题+正文）
        lines = []
        if item.get('title'):
            lines.append({
                'text': item['title'],
                'color': title_color,
                'size': SangforFonts.BODY_LARGE,
                'bold': True
            })
        if item.get('text'):
            lines.append({
                'text': item['text'],
                'color': text_color,
                'size': SangforFonts.BODY
            })
        
        txBox = add_multiline_textbox(
            slide, left_cm + 0.8, y, width_cm - 0.8, item_height_cm,
            lines, line_spacing=1.3
        )
        shapes.append(txBox)
    
    return shapes


def add_card(slide, left_cm, top_cm, width_cm, height_cm,
             header_text='', body_text='',
             header_color='#006CD9', bg_color='#ECECEC',
             bg_alpha=62):
    """添加内容卡片（标题条+正文区）
    
    Args:
        header_text: 标题文字
        body_text: 正文文字
        header_color: 标题条颜色
        bg_color: 卡片背景色
        bg_alpha: 背景透明度（0-100）
    Returns:
        创建的形状列表
    """
    shapes = []
    header_height = 1.3
    
    # 卡片背景
    bg = add_styled_rectangle(
        slide, left_cm, top_cm, width_cm, height_cm,
        fill_color=bg_color
    )
    if bg_alpha < 100:
        # 设置透明度
        spPr = bg._element.spPr
        solidFill = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
        if solidFill is not None:
            srgbClr = solidFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
            if srgbClr is not None:
                alpha = etree.SubElement(srgbClr, '{http://schemas.openxmlformats.org/drawingml/2006/main}alpha')
                alpha.set('val', str(int(bg_alpha * 1000)))
    shapes.append(bg)
    
    # 标题条
    header_bar = add_styled_rectangle(
        slide, left_cm, top_cm, width_cm, header_height,
        fill_color=header_color
    )
    shapes.append(header_bar)
    
    # 标题文字
    if header_text:
        header_tb = add_textbox(
            slide, left_cm + 0.5, top_cm, width_cm - 1, header_height,
            header_text,
            font_size=SangforFonts.BODY_LARGE,
            font_color=SangforColors.WHITE,
            bold=True
        )
        shapes.append(header_tb)
    
    # 正文文字
    if body_text:
        body_tb = add_textbox(
            slide, left_cm + 0.5, top_cm + header_height + 0.3,
            width_cm - 1, height_cm - header_height - 0.6,
            body_text,
            font_size=SangforFonts.BODY,
            font_color=SangforColors.TEXT_PRIMARY,
            line_spacing=1.5
        )
        shapes.append(body_tb)
    
    return shapes


def add_table(slide, left_cm, top_cm, width_cm, height_cm, data,
              header_color='#006CD9', odd_row_color='#F2F7FC',
              even_row_color='#FFFFFF'):
    """添加样式化表格
    
    Args:
        data: 二维列表 [['Header1', 'Header2'], ['row1col1', 'row1col2'], ...]
        header_color: 表头背景色
        odd_row_color: 奇数行背景色
        even_row_color: 偶数行背景色
    Returns:
        Table 对象
    """
    rows = len(data)
    cols = len(data[0]) if data else 0
    
    if rows == 0 or cols == 0:
        return None
    
    table_shape = slide.shapes.add_table(
        rows, cols, Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    table = table_shape.table
    
    for row_idx in range(rows):
        for col_idx in range(cols):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(data[row_idx][col_idx])
            
            # 设置字体
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    if row_idx == 0:
                        # 表头样式
                        set_font(run, '微软雅黑', SangforFonts.TABLE_HEADER,
                                SangforColors.WHITE, bold=True)
                    else:
                        # 正文样式
                        set_font(run, '微软雅黑', SangforFonts.TABLE_BODY,
                                SangforColors.TEXT_PRIMARY)
            
            # 设置单元格背景色
            if row_idx == 0:
                _set_cell_color(cell, header_color)
            elif row_idx % 2 == 1:
                _set_cell_color(cell, odd_row_color)
            else:
                _set_cell_color(cell, even_row_color)
    
    return table


def _set_cell_color(cell, hex_color):
    """设置表格单元格背景色"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
    srgbClr = etree.SubElement(solidFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
    srgbClr.set('val', hex_color.lstrip('#'))


def add_chart(slide, left_cm, top_cm, width_cm, height_cm,
              chart_type, chart_data, colors=None):
    """添加图表（增强样式版）
    
    Args:
        chart_type: 图表类型 ('column', 'bar', 'line', 'pie', 'doughnut',
                    'stacked_column', 'stacked_bar', 'area')
        chart_data: {
            'categories': ['分类1', '分类2', ...],
            'series': [
                {'name': '系列1', 'values': [1, 2, 3]},
                {'name': '系列2', 'values': [4, 5, 6]},
            ],
            'title': '图表标题（可选）'
        }
        colors: 颜色列表（默认使用深信服品牌色）
    Returns:
        Chart Shape 对象
    """
    if colors is None:
        colors = SangforColors.CHART_COLORS
    
    # 映射图表类型
    type_map = {
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'line': XL_CHART_TYPE.LINE_MARKERS,
        'pie': XL_CHART_TYPE.PIE,
        'doughnut': XL_CHART_TYPE.DOUGHNUT,
        'stacked_column': XL_CHART_TYPE.COLUMN_STACKED,
        'stacked_bar': XL_CHART_TYPE.BAR_STACKED,
        'area': XL_CHART_TYPE.AREA,
    }
    
    xl_type = type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    
    # 构建图表数据
    cd = CategoryChartData()
    cd.categories = chart_data['categories']
    for series in chart_data['series']:
        cd.add_series(series['name'], series['values'])
    
    # 添加图表
    chart_shape = slide.shapes.add_chart(
        xl_type, Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm), cd
    )
    
    chart = chart_shape.chart
    
    # === 通用样式 ===
    
    # 图例
    if len(chart_data['series']) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.name = '微软雅黑'
        chart.legend.font.size = Pt(9)
    else:
        chart.has_legend = False
    
    # 图表标题
    chart_title = chart_data.get('title', '')
    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = chart_title
        chart.chart_title.text_frame.paragraphs[0].font.name = '微软雅黑'
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        chart.chart_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(SangforColors.TEXT_PRIMARY)
    else:
        chart.has_title = False
    
    # 系列颜色
    plot = chart.plots[0]
    for i, series in enumerate(plot.series):
        color_idx = i % len(colors)
        fill = series.format.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(colors[color_idx])
    
    # === 按图表类型设置专属样式 ===
    
    if chart_type in ('pie', 'doughnut'):
        _style_pie_chart(chart, plot, colors)
    elif chart_type in ('line',):
        _style_line_chart(chart, plot, colors)
    else:
        _style_bar_chart(chart, plot)
    
    # 全局字体
    try:
        chart.font.name = '微软雅黑'
        chart.font.size = SangforFonts.CHART_LABEL
    except:
        pass
    
    return chart_shape


def _style_bar_chart(chart, plot):
    """柱状图/条形图专属样式"""
    # 显示数据标签
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = '微软雅黑'
    data_labels.font.size = Pt(9)
    data_labels.font.color.rgb = hex_to_rgb(SangforColors.TEXT_SECONDARY)
    data_labels.number_format = '#,##0'
    data_labels.number_format_is_linked = False
    
    # 设置坐标轴
    try:
        # 类别轴
        category_axis = chart.category_axis
        category_axis.has_major_gridlines = False
        category_axis.tick_labels.font.name = '微软雅黑'
        category_axis.tick_labels.font.size = Pt(9)
        category_axis.tick_labels.font.color.rgb = hex_to_rgb(SangforColors.TEXT_SECONDARY)
        
        # 值轴
        value_axis = chart.value_axis
        value_axis.has_major_gridlines = True
        value_axis.major_gridlines.format.line.color.rgb = hex_to_rgb('#E0E0E0')
        value_axis.major_gridlines.format.line.width = Pt(0.5)
        value_axis.tick_labels.font.name = '微软雅黑'
        value_axis.tick_labels.font.size = Pt(9)
        value_axis.tick_labels.font.color.rgb = hex_to_rgb(SangforColors.TEXT_SECONDARY)
    except:
        pass


def _style_pie_chart(chart, plot, colors):
    """饼图/环形图专属样式"""
    # 显示数据标签（带百分比）
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = '微软雅黑'
    data_labels.font.size = Pt(10)
    data_labels.number_format = '0.0%'
    data_labels.number_format_is_linked = False
    
    # 设置每个数据点的颜色（饼图系列只有一个，颜色按数据点设置）
    try:
        series = plot.series[0]
        for i in range(len(series.values)):
            point = series.points[i]
            color_idx = i % len(colors)
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = hex_to_rgb(colors[color_idx])
    except:
        pass


def _style_line_chart(chart, plot, colors):
    """折线图专属样式"""
    # 线宽和标记
    for i, series in enumerate(plot.series):
        series.format.line.width = Pt(2.5)
        series.smooth = False
        # 标记样式
        try:
            marker = series.marker
            marker.style = 8  # circle
            marker.size = 8
            color_idx = i % len(colors)
            marker.format.fill.solid()
            marker.format.fill.fore_color.rgb = hex_to_rgb(colors[color_idx])
        except:
            pass
    
    # 显示数据标签
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = '微软雅黑'
    data_labels.font.size = Pt(9)
    data_labels.font.color.rgb = hex_to_rgb(SangforColors.TEXT_SECONDARY)
    
    # 坐标轴
    try:
        value_axis = chart.value_axis
        value_axis.has_major_gridlines = True
        value_axis.major_gridlines.format.line.color.rgb = hex_to_rgb('#E0E0E0')
        value_axis.major_gridlines.format.line.width = Pt(0.5)
    except:
        pass


def add_image(slide, image_path, left_cm, top_cm, width_cm=None, height_cm=None):
    """添加图片，支持自适应尺寸
    
    Args:
        image_path: 图片文件路径
        left_cm, top_cm: 位置
        width_cm, height_cm: 尺寸（都为None时使用原始尺寸，
                              只指定一个时按比例缩放）
    Returns:
        Picture Shape 对象
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")
    
    if width_cm is not None and height_cm is not None:
        pic = slide.shapes.add_picture(
            image_path, Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
        )
    elif width_cm is not None:
        pic = slide.shapes.add_picture(
            image_path, Cm(left_cm), Cm(top_cm), Cm(width_cm)
        )
    elif height_cm is not None:
        pic = slide.shapes.add_picture(
            image_path, Cm(left_cm), Cm(top_cm), height=Cm(height_cm)
        )
    else:
        pic = slide.shapes.add_picture(
            image_path, Cm(left_cm), Cm(top_cm)
        )
    
    return pic


def add_number_highlight(slide, left_cm, top_cm, number_text, label_text,
                         number_color='#006CD9', label_color='#0E0E0E',
                         width_cm=6):
    """添加数字高亮展示（大号数字+说明文字）
    
    Args:
        number_text: 数字文字（如 "99%", "500+"）
        label_text: 说明文字
        number_color: 数字颜色
        label_color: 说明文字颜色
    """
    # 大号数字
    num_box = add_textbox(
        slide, left_cm, top_cm, width_cm, 2.5,
        number_text,
        font_size=SangforFonts.NUMBER_HIGHLIGHT,
        font_color=number_color,
        bold=True,
        alignment=PP_ALIGN.CENTER
    )
    
    # 说明文字
    label_box = add_textbox(
        slide, left_cm, top_cm + 2.5, width_cm, 1.5,
        label_text,
        font_size=SangforFonts.BODY,
        font_color=label_color,
        alignment=PP_ALIGN.CENTER
    )
    
    return num_box, label_box


# ============================================================
# 页面布局构建器
# ============================================================
# 内容块视觉权重配置
BLOCK_WEIGHT_CONFIG = {
    'text':             {'min_height': 1.5, 'preferred_height': 3.0, 'weight': 1.0},
    'bullet_list':      {'min_height': 3.0, 'preferred_height': None, 'weight': 2.0},  # 按项数动态
    'card_grid':        {'min_height': 5.0, 'preferred_height': 8.0, 'weight': 2.5},
    'table':            {'min_height': 4.0, 'preferred_height': 8.0, 'weight': 2.0},
    'chart':            {'min_height': 7.0, 'preferred_height': 11.0, 'weight': 3.0},
    'number_highlight': {'min_height': 4.0, 'preferred_height': 5.5, 'weight': 1.5},
    'image':            {'min_height': 5.0, 'preferred_height': 10.0, 'weight': 2.5},
    'two_column':       {'min_height': 5.0, 'preferred_height': 10.0, 'weight': 2.5},
}


def build_standard_page(slide, title, content_blocks):
    """在空白页上构建标准内容页（智能布局版）
    
    根据内容块类型的视觉权重动态分配空间，而非均分。
    图表和表格获得更大的空间，文字获得较少的空间。
    
    Args:
        slide: 空白幻灯片对象
        title: 页面标题
        content_blocks: 内容块列表
    """
    # 添加标题区
    add_title_area(slide, title)
    
    # 内容区边界
    content_top = 3.5
    content_left = 2.0
    content_width = 29.5
    content_bottom = 17.0
    available_height = content_bottom - content_top
    gap = 0.5  # 块间间距
    
    if not content_blocks:
        return
    
    # 如果只有一个内容块，给它全部空间
    if len(content_blocks) == 1:
        block = content_blocks[0]
        _render_content_block(slide, block, content_left, content_top, 
                            content_width, available_height)
        return
    
    # 多个内容块：按权重分配空间
    total_gaps = (len(content_blocks) - 1) * gap
    distributable_height = available_height - total_gaps
    
    # 计算每个块的权重
    weights = []
    for block in content_blocks:
        block_type = block.get('type', 'text')
        config = BLOCK_WEIGHT_CONFIG.get(block_type, BLOCK_WEIGHT_CONFIG['text'])
        
        weight = config['weight']
        
        # bullet_list 按项数调整权重
        if block_type == 'bullet_list':
            item_count = len(block.get('items', []))
            weight = max(1.0, item_count * 0.6)
        
        # card_grid 多行时增加权重
        elif block_type == 'card_grid':
            cards = block.get('cards', [])
            columns = block.get('columns', min(len(cards), 3))
            rows = (len(cards) + columns - 1) // columns if columns > 0 else 1
            weight = 2.0 + (rows - 1) * 1.5
        
        weights.append(weight)
    
    total_weight = sum(weights)
    
    # 分配高度（按权重比例），但不小于最小高度
    heights = []
    for i, block in enumerate(content_blocks):
        block_type = block.get('type', 'text')
        config = BLOCK_WEIGHT_CONFIG.get(block_type, BLOCK_WEIGHT_CONFIG['text'])
        
        proportional_height = distributable_height * (weights[i] / total_weight)
        min_height = config['min_height']
        allocated = max(proportional_height, min_height)
        heights.append(allocated)
    
    # 如果总高度超出可用空间，按比例缩减
    total_allocated = sum(heights)
    if total_allocated > distributable_height:
        scale = distributable_height / total_allocated
        heights = [h * scale for h in heights]
    
    # 渲染每个块
    current_y = content_top
    for i, block in enumerate(content_blocks):
        _render_content_block(slide, block, content_left, current_y,
                            content_width, heights[i])
        current_y += heights[i] + gap


def _render_content_block(slide, block, left, top, width, height):
    """渲染单个内容块"""
    block_type = block.get('type', 'text')
    
    if block_type == 'text':
        add_textbox(
            slide, left, top, width, height,
            block.get('content', ''),
            font_size=SangforFonts.BODY,
            font_color=SangforColors.TEXT_PRIMARY,
            line_spacing=1.5
        )
    
    elif block_type == 'bullet_list':
        items = block.get('items', [])
        item_h = min(height / max(len(items), 1), 3.0)
        add_bullet_list(
            slide, left, top, width, items,
            item_height_cm=item_h
        )
    
    elif block_type == 'card_grid':
        cards = block.get('cards', [])
        columns = block.get('columns', min(len(cards), 3))
        if columns == 0:
            return
        
        card_width = (width - (columns - 1) * 0.5) / columns
        card_height = height
        card_colors = [SangforColors.BLUE_PRIMARY, SangforColors.BLUE_HIGHLIGHT, 
                      SangforColors.BLUE_LIGHT2]
        
        for i, card in enumerate(cards):
            col = i % columns
            row = i // columns
            x = left + col * (card_width + 0.5)
            y = top + row * (card_height + 0.5)
            color = card_colors[col % len(card_colors)]
            
            add_card(
                slide, x, y, card_width, card_height,
                header_text=card.get('header', ''),
                body_text=card.get('body', ''),
                header_color=color
            )
    
    elif block_type == 'table':
        data = block.get('data', [])
        if data:
            add_table(slide, left, top, width, height, data)
    
    elif block_type == 'chart':
        chart_data = block.get('data', {})
        chart_type = block.get('chart_type', 'column')
        if chart_data:
            add_chart(slide, left, top, width, height, chart_type, chart_data)
    
    elif block_type == 'number_highlight':
        numbers = block.get('numbers', [])
        if not numbers:
            return
        num_width = width / len(numbers)
        for i, num in enumerate(numbers):
            x = left + i * num_width
            add_number_highlight(
                slide, x, top,
                num.get('value', ''),
                num.get('label', ''),
                width_cm=num_width
            )
    
    elif block_type == 'image':
        path = block.get('path', '')
        position = block.get('position', 'center')
        if path and os.path.exists(path):
            if position == 'full':
                add_image(slide, path, 0, 0, 33.87, 19.05)
            elif position == 'left':
                add_image(slide, path, left, top, width * 0.45, height)
            elif position == 'right':
                add_image(slide, path, left + width * 0.55, top, width * 0.45, height)
            else:
                add_image(slide, path, left + width * 0.1, top, width * 0.8, height)
    
    elif block_type == 'two_column':
        col_width = (width - 1) / 2
        left_block = block.get('left', {})
        right_block = block.get('right', {})
        _render_content_block(slide, left_block, left, top, col_width, height)
        _render_content_block(slide, right_block, left + col_width + 1, top, col_width, height)
