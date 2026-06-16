"""
深信服 PPT 生成器配置加载模块 (config.py)
从同一目录下的 ppt_config.json 动态加载配置并覆盖 constants 中的默认值。
"""
import os
import json
from pptx.util import Pt
from scripts.lib import constants

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ppt_config.json')

def load_config():
    """从 json 文件加载配置并动态修改 constants 模块"""
    if not os.path.exists(CONFIG_PATH):
        return False
        
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
            
        # 1. 覆盖颜色
        colors_cfg = config_data.get('colors', {})
        for key, val in colors_cfg.items():
            if hasattr(constants.SangforColors, key):
                setattr(constants.SangforColors, key, val)
                
        # 2. 覆盖字体和字号
        fonts_cfg = config_data.get('fonts', {})
        if 'CHINESE' in fonts_cfg:
            constants.SangforFonts.CHINESE = fonts_cfg['CHINESE']
        if 'ENGLISH' in fonts_cfg:
            constants.SangforFonts.ENGLISH = fonts_cfg['ENGLISH']
            
        sizes_cfg = fonts_cfg.get('sizes', {})
        for key, val in sizes_cfg.items():
            if hasattr(constants.SangforFonts, key):
                setattr(constants.SangforFonts, key, Pt(val))
                
        # 3. 覆盖排版引擎权重
        weights_cfg = config_data.get('block_weight_config', {})
        for key, val in weights_cfg.items():
            if key in constants.BLOCK_WEIGHT_CONFIG:
                constants.BLOCK_WEIGHT_CONFIG[key].update(val)
                
        return True
    except Exception as e:
        print(f"  警告: 加载配置文件 {CONFIG_PATH} 失败: {e}，将使用系统默认常量。")
        return False

# 导入时自动执行配置覆盖
load_config()
