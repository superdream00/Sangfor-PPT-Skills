#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 图片重建工具 - 从结构化 JSON 生成可编辑的 PowerPoint

将 AI 生成的 PPT 图片转换为原生 PowerPoint 形状，支持：
- 基础形状：矩形、圆角矩形、圆形
- 文本：标题、正文、多段文字
- 颜色：纯色填充、边框颜色
- 布局：精确位置控制

用法:
    python reconstruct_from_image.py layout.json output.pptx
"""

import json
import sys
import os
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor


def parse_color(color_str):
    """解析颜色字符串为 RGBColor

    Args:
        color_str: 十六进制颜色（如 "#9B59B6"）或颜色名（如 "purple"）

    Returns:
        RGBColor 对象
    """
    # 颜色名称映射（常用色）
    COLOR_NAMES = {
        "white": "#FFFFFF",
        "black": "#000000",
        "purple": "#9B59B6",
        "blue": "#3498DB",
        "green": "#2ECC71",
        "orange": "#E67E22",
        "red": "#E74C3C",
        "gray": "#95A5A6",
        "darkgray": "#7F8C8D",
    }

    # 如果是颜色名，转为十六进制
    if not color_str.startswith("#"):
        color_str = COLOR_NAMES.get(color_str.lower(), "#000000")

    # 去掉 # 并转为 RGB
    hex_color = color_str.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)

    return RGBColor(r, g, b)


def _draw_text(slide, elem):
    """绘制纯文本框

    Args:
        slide: 幻灯片对象
        elem: 元素描述（包含 position, size, style）
    """
    pos = elem['position']
    size = elem['size']
    style = elem['style']

    textbox = slide.shapes.add_textbox(
        Cm(pos['left_cm']), Cm(pos['top_cm']),
        Cm(size['width_cm']), Cm(size['height_cm'])
    )

    text_frame = textbox.text_frame
    text_frame.text = elem['content']

    # 设置样式
    para = text_frame.paragraphs[0]
    para.font.size = Pt(style.get('font_size_pt', 14))
    para.font.color.rgb = parse_color(style.get('font_color', '#000000'))
    para.font.bold = style.get('bold', False)
    para.font.name = '微软雅黑'

    # 对齐方式
    alignment = style.get('alignment', 'left')
    if alignment == 'center':
        para.alignment = PP_ALIGN.CENTER
    elif alignment == 'right':
        para.alignment = PP_ALIGN.RIGHT
    else:
        para.alignment = PP_ALIGN.LEFT

    text_frame.word_wrap = True


def _draw_card(slide, elem):
    """绘制卡片（圆角矩形 + 图标 + 标题 + 说明）

    Args:
        slide: 幻灯片对象
        elem: 卡片元素描述
    """
    pos = elem['position']
    size = elem['size']
    style = elem['style']

    # 1. 绘制圆角矩形背景
    card_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Cm(pos['left_cm']), Cm(pos['top_cm']),
        Cm(size['width_cm']), Cm(size['height_cm'])
    )
    card_shape.fill.solid()
    card_shape.fill.fore_color.rgb = parse_color(style['fill_color'])
    card_shape.line.width = Pt(0)  # 无边框

    # 2. 绘制图标（白色圆形）
    if 'icon' in elem:
        icon_info = elem['icon']
        icon_x = pos['left_cm'] + icon_info['position_offset']['x_cm']
        icon_y = pos['top_cm'] + icon_info['position_offset']['y_cm']
        icon_size = icon_info['size_cm']

        icon_circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Cm(icon_x), Cm(icon_y),
            Cm(icon_size), Cm(icon_size)
        )
        icon_circle.fill.solid()
        icon_circle.fill.fore_color.rgb = parse_color(icon_info['fill_color'])
        icon_circle.line.width = Pt(0)

    # 3. 标题文字
    if 'title' in elem:
        title_info = elem['title']
        title_box = slide.shapes.add_textbox(
            Cm(pos['left_cm'] + 0.5),
            Cm(pos['top_cm'] + title_info['y_offset_cm']),
            Cm(size['width_cm'] - 1),
            Cm(1)
        )
        title_frame = title_box.text_frame
        title_frame.text = title_info['content']

        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(title_info.get('font_size_pt', 18))
        title_para.font.color.rgb = parse_color(title_info.get('font_color', '#FFFFFF'))
        title_para.font.bold = title_info.get('bold', True)
        title_para.font.name = '微软雅黑'
        title_para.alignment = PP_ALIGN.CENTER

    # 4. 说明文字
    if 'description' in elem:
        desc_info = elem['description']
        desc_box = slide.shapes.add_textbox(
            Cm(pos['left_cm'] + 0.5),
            Cm(pos['top_cm'] + desc_info['y_offset_cm']),
            Cm(size['width_cm'] - 1),
            Cm(1.5)
        )
        desc_frame = desc_box.text_frame
        desc_frame.text = desc_info['content']

        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(desc_info.get('font_size_pt', 12))
        desc_para.font.color.rgb = parse_color(desc_info.get('font_color', '#FFFFFF'))
        desc_para.font.bold = False
        desc_para.font.name = '微软雅黑'
        desc_para.alignment = PP_ALIGN.CENTER
        desc_frame.word_wrap = True


def _draw_rounded_rectangle(slide, elem):
    """绘制圆角矩形（带可选文字）

    Args:
        slide: 幻灯片对象
        elem: 元素描述
    """
    pos = elem['position']
    size = elem['size']
    style = elem['style']

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Cm(pos['left_cm']), Cm(pos['top_cm']),
        Cm(size['width_cm']), Cm(size['height_cm'])
    )

    # 填充色
    shape.fill.solid()
    shape.fill.fore_color.rgb = parse_color(style.get('fill_color', '#ECECEC'))

    # 边框
    if 'line_color' in style:
        shape.line.color.rgb = parse_color(style['line_color'])
        shape.line.width = Pt(style.get('line_width_pt', 1))
    else:
        shape.line.width = Pt(0)

    # 内嵌文字（可选）
    if 'text' in elem:
        text_info = elem['text']
        text_frame = shape.text_frame
        text_frame.text = text_info.get('content', '')

        para = text_frame.paragraphs[0]
        para.font.size = Pt(text_info.get('font_size_pt', 14))
        para.font.color.rgb = parse_color(text_info.get('font_color', '#000000'))
        para.font.bold = text_info.get('bold', False)
        para.font.name = '微软雅黑'

        alignment = text_info.get('alignment', 'center')
        if alignment == 'center':
            para.alignment = PP_ALIGN.CENTER
        elif alignment == 'right':
            para.alignment = PP_ALIGN.RIGHT

        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.word_wrap = True


def reconstruct_ppt_from_json(layout_json_path, output_ppt_path):
    """从结构化 JSON 生成 PPT

    Args:
        layout_json_path: JSON 文件路径
        output_ppt_path: 输出 PPT 路径

    Returns:
        生成的 PPT 文件路径
    """
    # 读取 JSON
    with open(layout_json_path, 'r', encoding='utf-8') as f:
        layout = json.load(f)

    # 创建 PPT
    prs = Presentation()
    prs.slide_width = Cm(layout['page_info']['width_cm'])
    prs.slide_height = Cm(layout['page_info']['height_cm'])

    # 添加空白页
    blank_layout = prs.slide_layouts[6]  # 空白布局
    slide = prs.slides.add_slide(blank_layout)

    # 设置背景色（如果需要）
    bg_color = layout['page_info'].get('background_color', '#FFFFFF')
    if bg_color != '#FFFFFF':
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = parse_color(bg_color)

    # 逐个绘制元素
    for elem in layout['elements']:
        elem_type = elem['type']

        if elem_type == 'text':
            _draw_text(slide, elem)
        elif elem_type == 'card':
            _draw_card(slide, elem)
        elif elem_type == 'rounded_rectangle':
            _draw_rounded_rectangle(slide, elem)
        else:
            print(f"警告: 不支持的元素类型 '{elem_type}'，已跳过")

    # 保存 PPT
    prs.save(output_ppt_path)
    print(f"PPT generated: {output_ppt_path}")

    return output_ppt_path


def main():
    """命令行入口"""
    if len(sys.argv) < 3:
        print("用法: python reconstruct_from_image.py <layout.json> <output.pptx>")
        print("\n示例:")
        print("  python reconstruct_from_image.py test_layout.json output/reconstructed.pptx")
        sys.exit(1)

    layout_json = sys.argv[1]
    output_ppt = sys.argv[2]

    # 检查输入文件
    if not os.path.exists(layout_json):
        print(f"错误: JSON 文件不存在: {layout_json}")
        sys.exit(1)

    # 确保输出目录存在
    output_dir = os.path.dirname(output_ppt)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 生成 PPT
    try:
        reconstruct_ppt_from_json(layout_json, output_ppt)
        print(f"\nSuccess! Open {output_ppt} to view the result.")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
