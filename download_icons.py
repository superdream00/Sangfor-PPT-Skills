#!/usr/bin/env python3
"""
批量下载 Tabler Icons SVG 并修改为深信服品牌色
从 jsDelivr CDN 下载，速度快且无需克隆整个仓库
"""
import os
import re
import urllib.request
import urllib.error
from pathlib import Path

# Tabler Icons CDN (jsDelivr)
CDN_BASE = "https://cdn.jsdelivr.net/npm/@tabler/icons@latest/icons/"
SANGFOR_BLUE = "#006CD9"

def download_and_recolor_icon(icon_name, category, output_dir):
    """下载单个图标并改色"""
    # Tabler Icons 使用 kebab-case 文件名
    url = f"{CDN_BASE}{icon_name}"
    output_path = output_dir / category / icon_name

    try:
        # 下载
        with urllib.request.urlopen(url, timeout=10) as response:
            svg_content = response.read().decode('utf-8')

        # 改色：替换常见的颜色属性为深信服蓝
        # Tabler Icons 使用 stroke="currentColor"
        svg_content = svg_content.replace('stroke="currentColor"', f'stroke="{SANGFOR_BLUE}"')
        svg_content = svg_content.replace('fill="currentColor"', f'fill="{SANGFOR_BLUE}"')
        # 有些图标用 stroke="#000" 或其他颜色
        svg_content = re.sub(r'stroke="#[0-9a-fA-F]{3,6}"', f'stroke="{SANGFOR_BLUE}"', svg_content)
        svg_content = re.sub(r'fill="#[0-9a-fA-F]{3,6}"', f'fill="{SANGFOR_BLUE}"', svg_content)

        # 保存
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(svg_content)

        return True, None
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}"
    except Exception as e:
        return False, str(e)

def main():
    script_dir = Path(__file__).parent
    icons_dir = script_dir / "icons"
    list_file = icons_dir / "icon_list.txt"

    if not list_file.exists():
        print(f"错误: {list_file} 不存在")
        return

    # 解析清单
    icons_to_download = []
    with open(list_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '|' not in line:
                continue
            parts = [p.strip() for p in line.split('|')]
            if len(parts) >= 2:
                icon_name, category = parts[0], parts[1]
                icons_to_download.append((icon_name, category))

    print(f"准备下载 {len(icons_to_download)} 个图标...")

    success_count = 0
    failed = []

    for i, (icon_name, category) in enumerate(icons_to_download, 1):
        print(f"[{i}/{len(icons_to_download)}] {category}/{icon_name} ...", end=' ')
        ok, err = download_and_recolor_icon(icon_name, category, icons_dir)
        if ok:
            print("OK")
            success_count += 1
        else:
            print(f"FAIL ({err})")
            failed.append((icon_name, category, err))

    print(f"\n完成: {success_count}/{len(icons_to_download)} 成功")
    if failed:
        print(f"\n失败列表 ({len(failed)} 个):")
        for icon_name, category, err in failed:
            print(f"  {category}/{icon_name}: {err}")

if __name__ == '__main__':
    main()
