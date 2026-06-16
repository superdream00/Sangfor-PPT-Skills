"""
深信服 PPT 生成工具库兼容外壳 (utils.py)
重构为指向 scripts.lib 的导入壳，以保证与现有所有旧脚本的 100% 向后兼容性。
"""
# 引入 lib 包中的所有公共接口
from scripts.lib import *

# 显式引入并重定义一些可能被旧脚本访问的辅助/内部函数
from scripts.lib.utils import _set_cell_color, _fill_shape_text
from scripts.lib.charts import _style_bar_chart, _style_pie_chart, _style_line_chart
from scripts.lib.icons import (
    tokenize_path,
    parse_tokens,
    svg_arc_to_points,
    svg_path_to_ooxml,
    _find_icon_file
)
from scripts.lib.components import (
    _add_timeline_horizontal,
    _add_timeline_vertical,
    _render_canvas_block
)
from scripts.lib.layouts import _render_content_block
