#!/usr/bin/env python3
"""
将 icons/ 下所有 SVG 预转换为高分辨率 PNG（512px，透明背景）
运行时直接插 PNG，零运行时依赖。
用法: python convert_icons_to_png.py
"""
import os
from pathlib import Path
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM
from PIL import Image

TARGET_PX = 512  # 输出分辨率（正方形），保证矢量级清晰度

def _white_to_transparent(png_path):
    """把渲染出的白底转为透明（线条图标专用）"""
    img = Image.open(png_path).convert('RGBA')
    datas = img.getdata()
    new_data = []
    for r, g, b, a in datas:
        # 接近纯白的像素 -> 透明
        if r > 245 and g > 245 and b > 245:
            new_data.append((r, g, b, 0))
        else:
            new_data.append((r, g, b, a))
    img.putdata(new_data)
    img.save(png_path, 'PNG')

def convert_one(svg_path):
    """转换单个 SVG -> 同名 PNG（透明背景）"""
    png_path = svg_path.with_suffix('.png')
    try:
        drawing = svg2rlg(str(svg_path))
        if drawing is None:
            return False, "svg2rlg returned None"
        # 计算缩放，使最长边 = TARGET_PX
        w = drawing.width or 24
        h = drawing.height or 24
        scale = TARGET_PX / max(w, h)
        drawing.width = w * scale
        drawing.height = h * scale
        drawing.scale(scale, scale)
        renderPM.drawToFile(drawing, str(png_path), fmt='PNG', bg=0xFFFFFF)
        # 后处理：白底转透明
        _white_to_transparent(png_path)
        return True, None
    except Exception as e:
        return False, str(e)

def main():
    icons_root = Path(__file__).parent / "icons"
    svgs = list(icons_root.rglob("*.svg"))
    print(f"找到 {len(svgs)} 个 SVG，开始转换为 {TARGET_PX}px PNG...")

    ok_count = 0
    failed = []
    for i, svg in enumerate(svgs, 1):
        rel = svg.relative_to(icons_root)
        success, err = convert_one(svg)
        if success:
            ok_count += 1
        else:
            failed.append((str(rel), err))
        if i % 20 == 0:
            print(f"  进度 {i}/{len(svgs)}")

    print(f"\n完成: {ok_count}/{len(svgs)} 成功")
    if failed:
        print(f"失败 {len(failed)} 个:")
        for rel, err in failed[:30]:
            print(f"  {rel}: {err}")

if __name__ == '__main__':
    main()
