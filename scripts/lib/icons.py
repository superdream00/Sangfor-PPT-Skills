"""
深信服 PPT 生成器矢量图标模块 (icons.py)
解析 SVG 矢量图标并直接在 PPTX 中构建出原生 Custom Geometry，支持调整色彩和轮廓。
"""
import os
import re
import math
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree
from scripts.lib.utils import hex_to_rgb
from scripts.lib.constants import SangforColors

def tokenize_path(d_str):
    """提取 SVG 路径指令和数值"""
    token_pattern = re.compile(r'([A-Za-z]|-?\d*\.?\d+(?:[eE][-+]?\d+)?)')
    return token_pattern.findall(d_str)

def parse_tokens(tokens):
    """解析 token 列表为指令和参数对"""
    commands = []
    i = 0
    curr_cmd = None
    cmd_arg_counts = {
        'M': 2, 'm': 2, 'L': 2, 'l': 2, 'H': 1, 'h': 1, 'V': 1, 'v': 1,
        'C': 6, 'c': 6, 'S': 4, 's': 4, 'Q': 4, 'q': 4, 'T': 2, 't': 2,
        'A': 7, 'a': 7, 'Z': 0, 'z': 0
    }
    
    while i < len(tokens):
        token = tokens[i]
        if token.isalpha():
            curr_cmd = token
            i += 1
            arg_count = cmd_arg_counts.get(curr_cmd, 0)
            args = []
            for _ in range(arg_count):
                if i < len(tokens) and not tokens[i].isalpha():
                    args.append(float(tokens[i]))
                    i += 1
                else:
                    break
            commands.append((curr_cmd, args))
        else:
            if curr_cmd is None:
                i += 1
                continue
            implicit_cmd = curr_cmd
            if curr_cmd == 'M': implicit_cmd = 'L'
            elif curr_cmd == 'm': implicit_cmd = 'l'
            
            arg_count = cmd_arg_counts.get(implicit_cmd, 0)
            args = []
            for _ in range(arg_count):
                if i < len(tokens) and not tokens[i].isalpha():
                    args.append(float(tokens[i]))
                    i += 1
                else:
                    break
            commands.append((implicit_cmd, args))
            
    return commands

def svg_arc_to_points(x1, y1, rx, ry, phi, large_arc, sweep, x2, y2, num_segments=16):
    """将 SVG 弧线参数转换为逼近的折线段点集"""
    if x1 == x2 and y1 == y2:
        return []
    if rx == 0 or ry == 0:
        return [(x2, y2)]
    
    rx = abs(rx)
    ry = abs(ry)
    phi_rad = math.radians(phi)
    
    dx = (x1 - x2) / 2.0
    dy = (y1 - y2) / 2.0
    cos_phi = math.cos(phi_rad)
    sin_phi = math.sin(phi_rad)
    
    x1_prime = cos_phi * dx + sin_phi * dy
    y1_prime = -sin_phi * dx + cos_phi * dy
    
    radii_check = (x1_prime**2) / (rx**2) + (y1_prime**2) / (ry**2)
    if radii_check > 1.0:
        rx *= math.sqrt(radii_check)
        ry *= math.sqrt(radii_check)
        
    rx_sq, ry_sq = rx**2, ry**2
    x1_pr_sq, y1_pr_sq = x1_prime**2, y1_prime**2
    
    sign = -1 if large_arc == sweep else 1
    num = rx_sq * ry_sq - rx_sq * y1_pr_sq - ry_sq * x1_pr_sq
    den = rx_sq * y1_pr_sq + ry_sq * x1_pr_sq
    
    if num < 0: num = 0
    coef = sign * math.sqrt(num / den) if den != 0 else 0
    
    cx_prime = coef * (rx * y1_prime / ry)
    cy_prime = coef * (-ry * x1_prime / rx)
    
    cx = cos_phi * cx_prime - sin_phi * cy_prime + (x1 + x2) / 2.0
    cy = sin_phi * cx_prime + cos_phi * cy_prime + (y1 + y2) / 2.0
    
    def angle(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        len_u = math.sqrt(ux**2 + uy**2)
        len_v = math.sqrt(vx**2 + vy**2)
        if len_u == 0 or len_v == 0: return 0
        val = max(-1.0, min(1.0, dot / (len_u * len_v)))
        ang = math.acos(val)
        if (ux * vy - uy * vx) < 0: ang = -ang
        return ang
        
    theta1 = angle(1.0, 0.0, (x1_prime - cx_prime) / rx, (y1_prime - cy_prime) / ry)
    d_theta = angle((x1_prime - cx_prime) / rx, (y1_prime - cy_prime) / ry,
                    (-x1_prime - cx_prime) / rx, (-y1_prime - cy_prime) / ry)
                    
    if sweep == 0 and d_theta > 0: d_theta -= 2 * math.pi
    elif sweep == 1 and d_theta < 0: d_theta += 2 * math.pi
        
    points = []
    for step in range(1, num_segments + 1):
        t = theta1 + d_theta * (step / num_segments)
        xt = rx * math.cos(t)
        yt = ry * math.sin(t)
        x = cos_phi * xt - sin_phi * yt + cx
        y = sin_phi * xt + cos_phi * yt + cy
        points.append((x, y))
        
    return points

def svg_path_to_ooxml(d_str, viewBox_w=24, viewBox_h=24):
    """将 SVG 路径转换为等比例放大的 PPTX 矢量线段段 XML 片段"""
    tokens = tokenize_path(d_str)
    commands = parse_tokens(tokens)
    
    xml_parts = []
    curr_x, curr_y = 0.0, 0.0
    start_x, start_y = 0.0, 0.0
    prev_c_x, prev_c_y = 0.0, 0.0
    prev_cmd = None
    
    for cmd, args in commands:
        if cmd == 'M':
            curr_x, curr_y = args[0], args[1]
            start_x, start_y = curr_x, curr_y
            xml_parts.append(f'<a:moveTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:moveTo>')
        elif cmd == 'm':
            curr_x += args[0]
            curr_y += args[1]
            start_x, start_y = curr_x, curr_y
            xml_parts.append(f'<a:moveTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:moveTo>')
        elif cmd == 'L':
            curr_x, curr_y = args[0], args[1]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'l':
            curr_x += args[0]
            curr_y += args[1]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'H':
            curr_x = args[0]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'h':
            curr_x += args[0]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'V':
            curr_y = args[0]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'v':
            curr_y += args[0]
            xml_parts.append(f'<a:lnTo><a:pt x="{int(curr_x * 1000)}" y="{int(curr_y * 1000)}"/></a:lnTo>')
        elif cmd == 'C':
            x1, y1, x2, y2, x, y = args
            xml_parts.append(f'<a:cubicBezTo><a:pt x="{int(x1 * 1000)}" y="{int(y1 * 1000)}"/><a:pt x="{int(x2 * 1000)}" y="{int(y2 * 1000)}"/><a:pt x="{int(x * 1000)}" y="{int(y * 1000)}"/></a:cubicBezTo>')
            prev_c_x, prev_c_y = x2, y2
            curr_x, curr_y = x, y
        elif cmd == 'c':
            dx1, dy1, dx2, dy2, dx, dy = args
            x1, y1 = curr_x + dx1, curr_y + dy1
            x2, y2 = curr_x + dx2, curr_y + dy2
            x, y = curr_x + dx, curr_y + dy
            xml_parts.append(f'<a:cubicBezTo><a:pt x="{int(x1 * 1000)}" y="{int(y1 * 1000)}"/><a:pt x="{int(x2 * 1000)}" y="{int(y2 * 1000)}"/><a:pt x="{int(x * 1000)}" y="{int(y * 1000)}"/></a:cubicBezTo>')
            prev_c_x, prev_c_y = x2, y2
            curr_x, curr_y = x, y
        elif cmd == 'S':
            x2, y2, x, y = args
            if prev_cmd in ('C', 'c', 'S', 's'):
                x1 = 2 * curr_x - prev_c_x
                y1 = 2 * curr_y - prev_c_y
            else:
                x1, y1 = curr_x, curr_y
            xml_parts.append(f'<a:cubicBezTo><a:pt x="{int(x1 * 1000)}" y="{int(y1 * 1000)}"/><a:pt x="{int(x2 * 1000)}" y="{int(y2 * 1000)}"/><a:pt x="{int(x * 1000)}" y="{int(y * 1000)}"/></a:cubicBezTo>')
            prev_c_x, prev_c_y = x2, y2
            curr_x, curr_y = x, y
        elif cmd == 's':
            dx2, dy2, dx, dy = args
            x2, y2 = curr_x + dx2, curr_y + dy2
            x, y = curr_x + dx, curr_y + dy
            if prev_cmd in ('C', 'c', 'S', 's'):
                x1 = 2 * curr_x - prev_c_x
                y1 = 2 * curr_y - prev_c_y
            else:
                x1, y1 = curr_x, curr_y
            xml_parts.append(f'<a:cubicBezTo><a:pt x="{int(x1 * 1000)}" y="{int(y1 * 1000)}"/><a:pt x="{int(x2 * 1000)}" y="{int(y2 * 1000)}"/><a:pt x="{int(x * 1000)}" y="{int(y * 1000)}"/></a:cubicBezTo>')
            prev_c_x, prev_c_y = x2, y2
            curr_x, curr_y = x, y
        elif cmd == 'A':
            rx, ry, phi, large_arc, sweep, x, y = args
            pts = svg_arc_to_points(curr_x, curr_y, rx, ry, phi, large_arc, sweep, x, y)
            for px, py in pts:
                xml_parts.append(f'<a:lnTo><a:pt x="{int(px * 1000)}" y="{int(py * 1000)}"/></a:lnTo>')
            curr_x, curr_y = x, y
        elif cmd == 'a':
            rx, ry, phi, large_arc, sweep, dx, dy = args
            x, y = curr_x + dx, curr_y + dy
            pts = svg_arc_to_points(curr_x, curr_y, rx, ry, phi, large_arc, sweep, x, y)
            for px, py in pts:
                xml_parts.append(f'<a:lnTo><a:pt x="{int(px * 1000)}" y="{int(py * 1000)}"/></a:lnTo>')
            curr_x, curr_y = x, y
        elif cmd in ('Z', 'z'):
            xml_parts.append('<a:close/>')
            curr_x, curr_y = start_x, start_y
        prev_cmd = cmd
        
    return xml_parts

def _find_icon_file(icons_root, filename, category=None):
    """查找图标文件（跨 7 个分类目录）"""
    categories = [category] if category else \
        ['business', 'cloud', 'automotive', 'chip', 'electronics', 'energy', 'pharma']

    for cat in categories:
        if not cat:
            continue
        candidate = os.path.join(icons_root, cat, filename)
        if os.path.exists(candidate):
            return candidate
    return None

def add_icon(slide, icon_name, left_cm, top_cm, size_cm=1.5,
             color=None, category=None):
    """添加图标（解析 SVG 路径，直接构建为 PPTX 原生自定义图形）"""
    if color is None:
        color = SangforColors.BLUE_PRIMARY
        
    # 定位图标文件 (icons/ 目录位于项目根目录，即 lib/ 的上上级)
    lib_dir = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.dirname(lib_dir)
    project_root = os.path.dirname(scripts_dir)
    icons_root = os.path.join(project_root, 'icons')

    # 规范化文件名
    base = icon_name.replace('.svg', '').replace('.png', '')
    svg_name = base + '.svg'

    # 检测 SVG 源是否存在
    svg_path = _find_icon_file(icons_root, svg_name, category)

    if not svg_path:
        raise FileNotFoundError(
            f"图标 '{base}' 不存在（找不到 SVG 文件）。"
            f"请检查 icons/ 目录 (当前解析根路径: {icons_root})。"
        )

    # 读取并解析 SVG 文件的路径数据
    with open(svg_path, 'r', encoding='utf-8') as f:
        svg_content = f.read()
    
    # 提取所有 <path> 元素的 d 属性
    paths = re.findall(r'<path[^>]*d="([^"]+)"', svg_content)
    
    if not paths:
        raise ValueError(f"在图标 SVG 中未找到路径数据: {svg_path}")
        
    # 提取 viewBox
    viewBox_match = re.search(r'viewBox="([^"]+)"', svg_content)
    v_w, v_h = 24.0, 24.0
    if viewBox_match:
        parts = viewBox_match.group(1).split()
        if len(parts) == 4:
            v_w = float(parts[2])
            v_h = float(parts[3])
            
    # 创建一个标准矩形作为矢量图形容器
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(left_cm), Cm(top_cm), Cm(size_cm), Cm(size_cm)
    )
    
    # 禁用图形填充，设置描边轮廓
    shape.fill.background()
    shape.line.fill.solid()
    shape.line.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.width = Pt(1.5)
    
    # 将 SVG path 转换为 PPTX Custom Geometry XML
    paths_xml_list = []
    for d in paths:
        # 过滤掉 tabler icons 中常见的无填充全屏背景矩形路径 (stroke="none" d="M0 0h24v24H0z")
        if 'M0 0h24v24H0z' in d or 'M0 0h24v24H0z' in d.replace(' ', ''):
            continue
        cmds = svg_path_to_ooxml(d, v_w, v_h)
        if cmds:
            cmds_str = "".join(cmds)
            # DrawingML 标准写法：省略 fill 和 stroke 属性。
            paths_xml_list.append(f'<a:path w="{int(v_w * 1000)}" h="{int(v_h * 1000)}">{cmds_str}</a:path>')
            
    paths_xml = "".join(paths_xml_list)
    
    # 修改形状底层几何 XML
    spPr = shape.element.spPr
    prstGeom = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
    if prstGeom is not None:
        spPr.remove(prstGeom)
        
    custGeom_xml = f"""
    <a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:avLst/>
        <a:gdLst/>
        <a:ahLst/>
        <a:cxnLst/>
        <a:rect l="0" t="0" r="{int(v_w * 1000)}" b="{int(v_h * 1000)}"/>
        <a:pathLst>
            {paths_xml}
        </a:pathLst>
    </a:custGeom>
    """
    custGeom_el = etree.fromstring(custGeom_xml)
    
    # 按照 DrawingML Schema 规范的顺序插入元素：
    # spPr 子元素顺序必须为：xfrm -> EG_Geometry (custGeom) -> EG_FillProperties -> ln
    xfrm = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    if xfrm is not None:
        idx = list(spPr).index(xfrm)
        spPr.insert(idx + 1, custGeom_el)
    else:
        spPr.insert(0, custGeom_el)
    
    return shape
