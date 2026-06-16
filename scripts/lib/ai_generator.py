"""
深信服 PPT 生成器 AI 增强模块 (ai_generator.py)
利用大语言模型 (LLM) 将用户的文字大纲或描述翻译为 PPT 生成计划 JSON 字节流。
采用纯 urllib 实现，零外部依赖，支持 OpenAI 兼容 API（如 DeepSeek）以及 Google Gemini API。
"""
import json
import urllib.request
import urllib.error
import re
import os

# 默认提示词，指导大模型如何将文字转化为符合深信服 PPT 生成计划的 JSON 数据
SYSTEM_PROMPT = """
你是一个专门负责深信服 (Sangfor) 企业演示文稿排版大纲生成的 AI 专家。你的任务是将用户的文字描述、大纲或主题翻译为符合 Sangfor PPT 渲染引擎格式要求的生成计划 JSON 结构。

### 深信服 PPT 模板与设计规范
1. **视觉配色**：
   - 主色调（深信服蓝）：`#006CD9`
   - 强调色（深信服绿）：`#53C800`
   - 其他辅助色：`#0587F5`（浅蓝1）、`#3BA5FF`（浅蓝2）、`#0E0E0E`（正文黑）
2. **矢量图标库**：
   - 支持的图标名称示例（均支持描边变色）：`cloud`（云安全/云数据）、`shield`（网络安全/防御）、`server`（服务器/超融合）、`database`（数据库/存储）、`device`（安全终端/设备）、`cpu`（芯片/算力）、`arrow-right`（流程步骤）、`user`（用户/客户）、`chart-bar`（分析/业务指标）、`activity`（系统监控/日志）。

### 生成计划 JSON 规范
你的输出必须是符合以下 Schema 的纯 JSON 对象：
{
  "title": "演示文稿总体标题",
  "slides": [
    // 页面类型1：克隆模板页 (action = "clone")。用于封面页、目录页、章节过渡页、尾页等固定模板。
    {
      "action": "clone",
      "source_index": 1, // 模板页的索引。常见索引建议：第1页=封面（通常克隆2或3）、第6页=目录（克隆10）、第11-13页=纯列表、第14页=两栏对比、第15-16页=三栏并列、第49或50页=封底（谢谢聆听）。
      "replacements": {
        // 进行文本占位符替换。键为模板里的文字，值为实际要替换的文字。
        "大标题大标题大标题大标题": "实际的主标题",
        "小标题小标题小标题小标题小标题": "实际的副标题或时间部门信息"
      },
      "notes": "本页的演讲者备注或大纲提示（可选）"
    },
    // 页面类型2：智能排版页 (action = "blank_with_content")。在此页面上添加符合视觉规范的业务组件。
    {
      "action": "blank_with_content",
      "layout_name": "标题和内容", // 默认版式
      "title": "页面标题",
      "content_blocks": [
        // 模块A：多列卡片组
        {
          "type": "card_grid",
          "columns": 3, // 支持 2, 3, 4 列
          "cards": [
            {"header": "卡片1标题", "body": "卡片1的详细内容描述...", "icon": "cloud"},
            {"header": "卡片2标题", "body": "卡片2的详细内容描述...", "icon": "shield"},
            {"header": "卡片3标题", "body": "卡片3的详细内容描述...", "icon": "server"}
          ]
        },
        // 模块B：带图标或绿色圆点的项目列表
        {
          "type": "bullet_list",
          "items": [
            {"title": "条目1标题", "text": "条目的内容描述...", "icon": "cpu"},
            {"title": "条目2标题", "text": "条目的内容描述...", "icon": "device"}
          ]
        },
        // 模块C：竞品或方案对比表
        {
          "type": "comparison_table",
          "headers": ["功能特性", "深信服托管云", "传统公有云", "自建私有云"],
          "highlight_col": 1, // 突出深信服那一列（列索引从0开始）
          "rows": [
            ["安全合规", "内置安全服务，合规度高", "责任共担，合规需自行购买", "合规建设周期长，投入大"],
            ["运维响应", "专享贴身管家，分钟级响应", "工单流转慢，自助式服务", "需自主配备专业安全团队"]
          ]
        },
        // 模块D：水平流程步骤条
        {
          "type": "process_steps",
          "step_width_cm": 4.5,
          "spacing_cm": 0.8,
          "items": [
            {"title": "需求评估", "description": "识别业务资产与合规要求"},
            {"title": "方案设计", "description": "出具网络与安全拓扑方案"},
            {"title": "部署交付", "description": "一键激活策略并上线运行"}
          ]
        },
        // 模块E：高亮数字统计块
        {
          "type": "number_highlight",
          "numbers": [
            {"value": "99.99%", "label": "系统可用性指标"},
            {"value": "30分钟", "label": "应急安全响应时间"},
            {"value": "100+", "label": "全球服务网点"}
          ]
        },
        // 模块F：数据分析图表（支持 column, bar, line, pie, doughnut 等）
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "title": "安全事件平均响应时间趋势 (分钟)",
            "categories": ["2022", "2023", "2024"],
            "series": [
              {"name": "传统方式", "values": [120, 95, 80]},
              {"name": "Sangfor托管方式", "values": [30, 20, 15]}
            ]
          }
        },
        // 模块G：里程碑/时间轴（水平或垂直）
        {
          "type": "timeline",
          "orientation": "horizontal", // horizontal 或 vertical
          "items": [
            {"date": "第一阶段", "title": "安全架构评估", "description": "漏洞扫描与漏洞修复验证"},
            {"date": "第二阶段", "title": "安全策略统合", "description": "多分支机构防火墙规则拉齐"},
            {"date": "第三阶段", "title": "持续监测感知", "description": "态势感知系统24小时值守"}
          ]
        }
      ],
      "notes": "本页的演讲大纲提示"
    }
  ]
}

### 约束原则
1. **结构精炼**：每页的 `content_blocks` 数量建议为 1-2 个，最多不超过 3 个，以防页面内容拥挤堆叠。
2. **严禁废话**：你的回答必须是纯 JSON 格式，绝不允许带有任何额外的解释文字、Markdown 包裹块（除了 standard json markdown block）或前后置引导语。
3. **语言要求**：PPT 页面里的内容使用**中文**生成。
"""

def generate_slides_plan(prompt, api_key=None, base_url=None, model=None):
    """
    通过调用 LLM API 将自然语言 prompt 翻译为 slides_plan JSON 对象。
    
    Args:
        prompt: 用户的文字大纲或要求
        api_key: LLM API key。若不传，则尝试从环境变量加载。
        base_url: LLM API base URL。若不传，则默认尝试 OpenAI 或环境变量。
        model: LLM 调用的模型名称。若不传，则默认尝试环境变量或使用默认值。
    """
    # 1. 解析 API 密钥和参数
    api_key = api_key or os.environ.get('LLM_API_KEY') or os.environ.get('OPENAI_API_KEY') or os.environ.get('GEMINI_API_KEY')
    if not api_key:
        raise ValueError("未配置 API Key。请设置环境变量 LLM_API_KEY 或在请求中传入。")
        
    model = model or os.environ.get('LLM_MODEL') or 'deepseek-chat'
    
    # 检测是否为 Google Gemini API 方式
    is_gemini = False
    if 'gemini' in model.lower() or (base_url and 'googleapis' in base_url.lower()) or api_key.startswith('AIza'):
        is_gemini = True
        
    # 设置默认 base_url
    if not base_url:
        if is_gemini:
            base_url = "https://generativelanguage.googleapis.com/v1beta"
        else:
            base_url = os.environ.get('LLM_API_BASE') or "https://api.deepseek.com/v1"
            
    # 去除尾部斜杠
    base_url = base_url.rstrip('/')
    
    # 2. 构造请求 Payload
    if is_gemini:
        # Gemini API 发送格式
        url = f"{base_url}/models/{model}:generateContent?key={api_key}"
        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": f"{SYSTEM_PROMPT}\n\n用户生成主题或大纲要求：\n{prompt}"}
                    ]
                }
            ],
            "generationConfig": {
                "responseMimeType": "application/json"
            }
        }
        headers = {"Content-Type": "application/json"}
    else:
        # OpenAI 兼容 API 发送格式
        url = f"{base_url}/chat/completions"
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3
        }
        # 如果不是一些特殊的小模型，尝试开启 JSON Mode
        if "deepseek" in model.lower() or "gpt-" in model.lower() or "claude-" in model.lower():
            payload["response_format"] = {"type": "json_object"}
            
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers=headers, method='POST')
    
    # 3. 发送请求并获取结果
    print(f"正在发起 LLM 请求 (Model: {model})...")
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            res_body = response.read().decode('utf-8')
            res_json = json.loads(res_body)
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode('utf-8')
        raise RuntimeError(f"LLM API HTTP 错误 ({e.code}): {err_msg}")
    except Exception as e:
        raise RuntimeError(f"连接 LLM 接口时发生未知错误: {e}")
        
    # 4. 解析大模型返回的数据
    if is_gemini:
        try:
            content_text = res_json['candidates'][0]['content']['parts'][0]['text']
        except KeyError:
            raise ValueError(f"Gemini API 返回格式异常: {res_json}")
    else:
        try:
            content_text = res_json['choices'][0]['message']['content']
        except KeyError:
            raise ValueError(f"OpenAI API 返回格式异常: {res_json}")
            
    # 5. 精确提取 JSON 内容并校验格式
    try:
        # 大模型有时仍会加上 ```json ... ``` 标签，使用正则清洗
        json_clean = content_text.strip()
        if json_clean.startswith('```'):
            match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', json_clean)
            if match:
                json_clean = match.group(1).strip()
                
        plan_dict = json.loads(json_clean)
        return plan_dict
    except Exception as e:
        print("解析大模型 JSON 失败。大模型原始返回内容如下:")
        print(content_text)
        raise ValueError(f"无法将 LLM 的输出解析为 JSON 对象: {e}")
