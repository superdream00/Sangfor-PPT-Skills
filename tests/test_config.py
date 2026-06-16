"""
测试配置加载与常量模块
"""
import pytest
from scripts.lib import constants
from scripts.lib.config import load_config

def test_constants_default():
    """验证默认常量是否正确"""
    assert constants.SangforColors.BLUE_PRIMARY == '#006CD9'
    assert constants.SangforColors.GREEN_PRIMARY == '#53C800'
    assert constants.SangforFonts.CHINESE == '微软雅黑'

def test_load_config():
    """验证配置加载是否成功（或回退成功）"""
    res = load_config()
    # load_config 应返回 True (已加载配置文件) 或 False (文件不存在但安全回退)
    assert res in (True, False)
