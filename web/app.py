"""
深信服 PPT 生成器 Flask Web 服务 (app.py)
提供在线生成、API 接口服务，并支持通过 LLM 将文本直接生成演示文稿。
"""
import os
import sys
import json
import tempfile
from flask import Flask, request, jsonify, send_file, render_template

# 添加项目根目录到 Path 确保能正确导入 scripts
web_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(web_dir)
sys.path.insert(0, project_root)

from scripts.generate_ppt import generate_ppt
from scripts.lib.ai_generator import generate_slides_plan

app = Flask(__name__, 
            template_folder=os.path.join(web_dir, 'templates'),
            static_folder=os.path.join(web_dir, 'static'))

# 默认使用的 PPT 模板文件路径
DEFAULT_TEMPLATE = os.path.join(project_root, 'templates', 'sangfor_template_2024.pptx')

@app.route('/')
def index():
    """首页面"""
    return render_template('index.html')

@app.route('/api/icons', methods=['GET'])
def get_icons():
    """获取所有可用图标列表（按分类）"""
    icons_root = os.path.join(project_root, 'icons')
    icons_data = {}
    
    if not os.path.exists(icons_root):
        return jsonify({"error": f"图标目录不存在: {icons_root}"}), 404
        
    try:
        for cat in os.listdir(icons_root):
            cat_path = os.path.join(icons_root, cat)
            if os.path.isdir(cat_path) and not cat.startswith('.'):
                icons_data[cat] = []
                for file in os.listdir(cat_path):
                    if file.endswith('.svg') or file.endswith('.png'):
                        name = file.rsplit('.', 1)[0]
                        if name not in icons_data[cat]:
                            icons_data[cat].append(name)
        return jsonify(icons_data)
    except Exception as e:
        return jsonify({"error": f"列出图标失败: {e}"}), 500

@app.route('/api/generate', methods=['POST'])
def generate():
    """从提交的 slides_plan JSON 在线生成 PPT 并提供下载"""
    try:
        plan_data = request.json
        if not plan_data:
            return jsonify({"error": "缺少生成计划 JSON 数据"}), 400
            
        # 1. 寻找模板文件
        template_path = DEFAULT_TEMPLATE
        if not os.path.exists(template_path):
            # 兼容性寻找：如果 .pptx 不存在，寻找 .potx
            potx_path = DEFAULT_TEMPLATE.replace('.pptx', '.potx')
            if os.path.exists(potx_path):
                template_path = potx_path
            else:
                return jsonify({"error": "在服务器端找不到默认的深信服 PPT 模板"}), 500

        # 2. 创建临时输出路径
        fd, temp_out = tempfile.mkstemp(suffix='.pptx')
        os.close(fd)
        
        # 3. 运行生成引擎
        generate_ppt(template_path, plan_data, temp_out)
        
        # 4. 返回下载
        response = send_file(
            temp_out,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="sangfor_generated_presentation.pptx"
        )
        
        # 注册清理回调，在发送完毕后删除临时文件
        @response.call_on_close
        def remove_temp():
            try:
                if os.path.exists(temp_out):
                    os.remove(temp_out)
            except Exception as e:
                app.logger.error(f"清除临时文件 {temp_out} 失败: {e}")
                
        return response
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"生成 PPT 失败: {str(e)}"}), 500

@app.route('/api/ai-generate', methods=['POST'])
def ai_generate():
    """结合大语言模型，从文字描述中直接智能生成并下载 PPT"""
    try:
        req_data = request.json or {}
        prompt = req_data.get('prompt')
        if not prompt:
            return jsonify({"error": "缺少文字描述字段 'prompt'"}), 400
            
        api_key = req_data.get('api_key')
        base_url = req_data.get('base_url')
        model = req_data.get('model')
        
        # 1. 调用 LLM 生成计划 JSON
        plan_data = generate_slides_plan(
            prompt=prompt,
            api_key=api_key,
            base_url=base_url,
            model=model
        )
        
        # 2. 验证和微调生成的 plan 结构
        if 'slides' not in plan_data:
            return jsonify({"error": "大语言模型未能返回合法的 slides 数据列表", "llm_output": plan_data}), 500
            
        # 3. 渲染成 PPT 物理文件
        template_path = DEFAULT_TEMPLATE
        if not os.path.exists(template_path):
            potx_path = DEFAULT_TEMPLATE.replace('.pptx', '.potx')
            if os.path.exists(potx_path):
                template_path = potx_path
            else:
                return jsonify({"error": "在服务器端找不到默认的深信服 PPT 模板"}), 500

        fd, temp_out = tempfile.mkstemp(suffix='.pptx')
        os.close(fd)
        
        generate_ppt(template_path, plan_data, temp_out)
        
        # 4. 返回下载并将大纲计划也作为 Header 发送（供前端展示）
        response = send_file(
            temp_out,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="sangfor_ai_generated.pptx"
        )
        
        # 把 JSON 计划写进自定义 Response 标头（经过 url 编码以防中文报错）
        import urllib.parse
        encoded_plan = urllib.parse.quote(json.dumps(plan_dict_clean(plan_data)))
        response.headers['X-Slides-Plan'] = encoded_plan

        @response.call_on_close
        def remove_temp():
            try:
                if os.path.exists(temp_out):
                    os.remove(temp_out)
            except Exception as e:
                app.logger.error(f"清除临时文件 {temp_out} 失败: {e}")
                
        return response

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"AI 生成 PPT 失败: {str(e)}"}), 500

def plan_dict_clean(plan):
    """提取一个精简版的 plan，避免过长导致 header 报错"""
    cleaned = {"title": plan.get("title", ""), "slides": []}
    for slide in plan.get("slides", []):
        cleaned["slides"].append({
            "title": slide.get("title", "未命名页面"),
            "action": slide.get("action", "")
        })
    return cleaned

if __name__ == '__main__':
    # 支持命令行自定义端口
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f"深信服 PPT 在线服务已启动，请在浏览器中打开: http://127.0.0.1:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
