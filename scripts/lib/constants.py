"""
深信服 PPT 生成器常量模块 (constants.py)
定义品牌视觉色彩、字体字号规范、DrawingML 命名空间以及布局引擎的默认权重配置。
"""
from pptx.util import Pt

# DrawingML XML 命名空间
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

# 默认跳过的模板页索引（0-based）
SKIP_PAGES = [0, 4, 5]

class SangforColors:
    # 主色系 (蓝色与背景衬底)
    BLUE_PRIMARY = '#006CD9'       # 深信服蓝
    BLUE_DARK = '#003592'          # 深蓝
    BLUE_DEEP = '#00479D'          # 深蓝辅助
    BLUE_NAVY = '#02288C'          # 深藏青蓝
    BLUE_LIGHT1 = '#0587F5'        # 品牌浅蓝1
    BLUE_LIGHT2 = '#3BA5FF'        # 品牌浅蓝2
    BLUE_LIGHT3 = '#51C0F9'        # 品牌浅蓝3
    BLUE_HIGHLIGHT = '#0185FF'     # 高亮蓝
    BLUE_SOFT = '#65B5FF'          # 柔和蓝
    BLUE_PALE = '#80D8FD'          # 淡蓝
    BLUE_BG_SOFT = '#CDD1F6'       # 柔和蓝背景（淡青蓝衬底）
    
    # 强调与警示色 (绿、橙、紫、黄)
    GREEN_PRIMARY = '#53C800'      # 深信服绿
    GREEN_VARIANT1 = '#6FBA2C'     # 绿色变体1
    GREEN_VARIANT2 = '#6CCD24'     # 绿色变体2
    GREEN_LIGHT = '#9BF48F'        # 浅绿
    GREEN_SOFT = '#CCE28E'         # 柔和绿
    ORANGE_LIGHT = '#FFD6BF'       # 珊瑚粉/浅橙
    PURPLE_ACCENT = '#6159E5'      # 智能紫
    YELLOW_HIGHLIGHT = '#FFC000'   # 警告黄/高亮金
    
    # 中性色
    TEXT_PRIMARY = '#0E0E0E'       # 正文主色
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

class SangforFonts:
    CHINESE = '微软雅黑'
    ENGLISH = '微软雅黑'
    
    # 各场景默认字号
    COVER_TITLE = Pt(26)
    COVER_SUBTITLE = Pt(14)
    PAGE_TITLE = Pt(22)
    SUBTITLE = Pt(16)
    BODY = Pt(12)
    BODY_LARGE = Pt(14)
    SMALL = Pt(10)
    CAPTION = Pt(9)
    TABLE_HEADER = Pt(12)
    TABLE_BODY = Pt(11)
    CHART_LABEL = Pt(10)
    NUMBER_HIGHLIGHT = Pt(44)
    SECTION_TITLE = Pt(44)

# 内容块默认高度分配权重
BLOCK_WEIGHT_CONFIG = {
    'text':             {'min_height': 1.5, 'preferred_height': 3.0, 'weight': 1.0},
    'bullet_list':      {'min_height': 3.0, 'preferred_height': None, 'weight': 2.0},
    'card_grid':        {'min_height': 5.0, 'preferred_height': 8.0, 'weight': 2.5},
    'table':            {'min_height': 4.0, 'preferred_height': 8.0, 'weight': 2.0},
    'chart':            {'min_height': 7.0, 'preferred_height': 11.0, 'weight': 3.0},
    'number_highlight': {'min_height': 4.0, 'preferred_height': 5.5, 'weight': 1.5},
    'image':            {'min_height': 5.0, 'preferred_height': 10.0, 'weight': 2.5},
    'two_column':       {'min_height': 5.0, 'preferred_height': 10.0, 'weight': 2.5},
    'icon_row':         {'min_height': 2.5, 'preferred_height': 3.0, 'weight': 1.2},
    'canvas':           {'min_height': 5.0, 'preferred_height': 12.0, 'weight': 3.0},
    'grid_matrix':      {'min_height': 5.0, 'preferred_height': 8.0, 'weight': 2.5},
    'timeline':         {'min_height': 4.0, 'preferred_height': 6.0, 'weight': 2.0},
}
