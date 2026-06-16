"""
深信服 PPT 生成工具库公共接口 (lib)
整合颜色、字体、几何形状、矢量图标、图表及排版引擎，提供统一调用入口。
"""
# 动态加载并执行配置读取
from scripts.lib.config import load_config

# 暴露常量与命名空间
from scripts.lib.constants import (
    NSMAP,
    SKIP_PAGES,
    SangforColors,
    SangforFonts,
    BLOCK_WEIGHT_CONFIG
)

# 暴露基础底层工具
from scripts.lib.utils import (
    hex_to_rgb,
    set_font,
    set_paragraph_spacing
)

# 暴露几何图形绘制
from scripts.lib.shapes import (
    add_textbox,
    add_multiline_textbox,
    add_styled_rectangle,
    add_styled_circle,
    add_gradient_fill,
    add_connector_arrow,
    add_connector_line,
    add_image,
    set_no_border
)

# 暴露矢量图标
from scripts.lib.icons import (
    add_icon
)

# 暴露数据图表
from scripts.lib.charts import (
    add_chart
)

# 暴露业务高级组件
from scripts.lib.components import (
    add_title_area,
    add_bullet_list,
    add_card,
    add_table,
    add_comparison_table,
    add_process_steps,
    add_timeline,
    add_number_highlight,
    add_icon_row
)

# 暴露智能排版引擎
from scripts.lib.layouts import (
    build_standard_page
)
