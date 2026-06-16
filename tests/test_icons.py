"""
测试 SVG 图标解析与 OOXML 转换逻辑
"""
import pytest
from scripts.lib.icons import tokenize_path, parse_tokens, svg_path_to_ooxml

def test_tokenize_path():
    """测试 SVG 路径字符串分词"""
    d = "M10 20L30,40 Z"
    tokens = tokenize_path(d)
    assert tokens == ['M', '10', '20', 'L', '30', '40', 'Z']

def test_parse_tokens():
    """测试将 Token 序列解析为指令列表"""
    tokens = ['M', '10.5', '-20', 'l', '30', '40.2', 'Z']
    commands = parse_tokens(tokens)
    assert commands == [
        ('M', [10.5, -20.0]),
        ('l', [30.0, 40.2]),
        ('Z', [])
    ]

def test_svg_path_to_ooxml():
    """测试 SVG 路径指令到 DrawingML 的 XML 转换"""
    d = "M 10 20 L 30 40 Z"
    xml_parts = svg_path_to_ooxml(d)
    assert len(xml_parts) == 3
    assert xml_parts[0] == '<a:moveTo><a:pt x="10000" y="20000"/></a:moveTo>'
    assert xml_parts[1] == '<a:lnTo><a:pt x="30000" y="40000"/></a:lnTo>'
    assert xml_parts[2] == '<a:close/>'
