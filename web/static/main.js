/**
 * 深信服 PPT 在线智能生成平台前端逻辑 (main.js)
 */

// 全局变量存储所有图标数据，方便搜索过滤
let allIcons = {};

// 初始化
window.addEventListener('DOMContentLoaded', () => {
    // 默认加载基本的 JSON Demo 到编辑器中
    loadPreset('basic');
    // 获取矢量图标库
    fetchIcons();
    // 从本地存储或环境变量读取 LLM API Key (如果有)
    loadLlmConfig();
});

// 日志记录函数，向控制台添加一行日志
function log(message, type = 'info') {
    const consoleLog = document.getElementById('console-log');
    if (!consoleLog) return;

    const line = document.createElement('div');
    line.className = 'console-line';
    
    const time = new Date().toLocaleTimeString();
    
    let typeClass = 'console-info';
    let prefix = '[INFO]';
    if (type === 'warn') {
        typeClass = 'console-warn';
        prefix = '[WARN]';
    } else if (type === 'error') {
        typeClass = 'console-error';
        prefix = '[ERROR]';
    } else if (type === 'system') {
        typeClass = 'console-info';
        prefix = '[SYSTEM]';
    }

    line.innerHTML = `
        <span class="console-prefix">${time} ${prefix}</span>
        <span class="${typeClass}">${escapeHtml(message)}</span>
    `;
    
    consoleLog.appendChild(line);
    consoleLog.scrollTop = consoleLog.scrollHeight;
}

function escapeHtml(text) {
    return text.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// 切换选项卡模式
function switchMode(mode) {
    // 移除 active 类
    document.querySelectorAll('.mode-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(panel => panel.classList.remove('active'));
    
    if (mode === 'ai') {
        document.getElementById('btn-mode-ai').classList.add('active');
        document.getElementById('panel-ai').classList.add('active');
        log('切换到 AI 描述生成模式', 'system');
    } else {
        document.getElementById('btn-mode-json').classList.add('active');
        document.getElementById('panel-json').classList.add('active');
        log('切换到 JSON 计划配置生成模式', 'system');
    }
}

// 可选：加载/保存本地大模型配置
function loadLlmConfig() {
    const savedKey = localStorage.getItem('llm_key');
    const savedBase = localStorage.getItem('llm_base');
    const savedModel = localStorage.getItem('llm_model');
    
    if (savedKey) document.getElementById('llm-key').value = savedKey;
    if (savedBase) document.getElementById('llm-base').value = savedBase;
    if (savedModel) document.getElementById('llm-model').value = savedModel;
}

function saveLlmConfig(key, base, model) {
    localStorage.setItem('llm_key', key);
    localStorage.setItem('llm_base', base);
    localStorage.setItem('llm_model', model);
}

// 加载图标列表
async function fetchIcons() {
    try {
        const response = await fetch('/api/icons');
        if (!response.ok) {
            throw new Error(`HTTP 错误: ${response.status}`);
        }
        allIcons = await response.json();
        renderIcons(allIcons);
        log(`成功加载图标库：共 ${Object.keys(allIcons).length} 个分类`, 'info');
    } catch (e) {
        log(`加载图标库失败: ${e.message}`, 'error');
        document.getElementById('icon-list-container').innerHTML = `
            <div style="font-size: 12px; color: var(--text-sub);">加载图标库失败，请检查服务器运行状态。</div>
        `;
    }
}

// 渲染图标列表到界面上
function renderIcons(iconsData) {
    const container = document.getElementById('icon-list-container');
    container.innerHTML = '';
    
    for (const [category, icons] of Object.entries(iconsData)) {
        if (icons.length === 0) continue;
        
        const group = document.createElement('div');
        group.className = 'icon-category-group';
        
        const title = document.createElement('div');
        title.className = 'icon-category-title';
        title.innerText = translateCategory(category);
        group.appendChild(title);
        
        const list = document.createElement('div');
        list.className = 'icon-list';
        
        icons.forEach(icon => {
            const badge = document.createElement('span');
            badge.className = 'icon-badge';
            badge.innerText = icon;
            badge.title = '点击复制图标名称';
            badge.onclick = () => copyToClipboard(icon);
            list.appendChild(badge);
        });
        
        group.appendChild(list);
        container.appendChild(group);
    }
}

function translateCategory(cat) {
    const mapping = {
        'business': '商业与政企',
        'cloud': '云计算与虚拟化',
        'automotive': '汽车与工业',
        'chip': '芯片与硬件',
        'electronics': '电子消费品',
        'energy': '能源与化工',
        'pharma': '医疗与生物'
    };
    return mapping[cat.toLowerCase()] || cat;
}

// 图标搜索过滤
function filterIcons() {
    const query = document.getElementById('icon-search').value.toLowerCase().trim();
    if (!query) {
        renderIcons(allIcons);
        return;
    }
    
    const filtered = {};
    for (const [category, icons] of Object.entries(allIcons)) {
        const matching = icons.filter(icon => icon.toLowerCase().includes(query));
        if (matching.length > 0) {
            filtered[category] = matching;
        }
    }
    renderIcons(filtered);
}

// 点击图标名称复制到剪贴板
function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        log(`已复制图标名称: "${text}" 到剪贴板`, 'info');
    }).catch(err => {
        log(`复制失败: ${err}`, 'error');
    });
}

// JSON Editor 格式化
function formatJson() {
    const textarea = document.getElementById('json-code');
    try {
        const obj = JSON.parse(textarea.value);
        textarea.value = JSON.stringify(obj, null, 2);
        log('JSON 格式化成功', 'info');
    } catch (e) {
        log(`JSON 语法错误，无法格式化: ${e.message}`, 'error');
    }
}

// JSON 预设配置
const presets = {
    basic: {
        "title": "深信服智能产品特性介绍",
        "slides": [
            {
                "action": "clone",
                "source_index": 1,
                "replacements": {
                    "大标题大标题大标题大标题": "深信服智能大盘系统",
                    "小标题小标题小标题小标题小标题": "安全与云计算联合解决方案运营汇报"
                }
            },
            {
                "action": "blank_with_content",
                "title": "核心方案三大核心组件",
                "content_blocks": [
                    {
                        "type": "card_grid",
                        "columns": 3,
                        "cards": [
                            {
                                "header": "超融合计算平台",
                                "body": "极简架构交付，承载核心应用，提供极高的资源弹性与异构兼容性。",
                                "icon": "server"
                            },
                            {
                                "header": "全局态势感知系统",
                                "body": "全网资产可视，AI 行为引擎关联，毫秒级响应未知威胁和高危入侵行为。",
                                "icon": "shield"
                            },
                            {
                                "header": "云端数据托管服务",
                                "body": "三副本物理备份，安全网关隔离，金融级存储合规与端到端高强度加密保护。",
                                "icon": "cloud"
                            }
                        ]
                    }
                ]
            },
            {
                "action": "clone",
                "source_index": 49,
                "replacements": {
                    "谢谢聆听": "谢谢聆听",
                    "汇报人：深信服项目团队": "汇报人：深信服智能解决方案部门"
                }
            }
        ]
    },
    
    timeline: {
        "title": "业务迁移与演练路线规划",
        "slides": [
            {
                "action": "blank_with_content",
                "title": "三阶段项目落地流程图",
                "content_blocks": [
                    {
                        "type": "process_steps",
                        "step_width_cm": 5.0,
                        "spacing_cm": 1.0,
                        "items": [
                            {
                                "title": "架构评估与拓扑设计",
                                "description": "识别业务资产，定位安全等级，规划子网"
                            },
                            {
                                "title": "策略激活与同步",
                                "description": "数据镜像迁移，零信任网关部署，双活运行"
                            },
                            {
                                "title": "实兵演练与交付",
                                "description": "高风险红蓝对抗检测，容灾切换测试演习"
                            }
                        ]
                    }
                ]
            },
            {
                "action": "blank_with_content",
                "title": "年度安全大盘里程碑规划",
                "content_blocks": [
                    {
                        "type": "timeline",
                        "orientation": "horizontal",
                        "items": [
                            {
                                "date": "Q1 季度",
                                "title": "漏洞资产大盘排查",
                                "description": "完成500+子系统全量漏洞排障与修复"
                            },
                            {
                                "date": "Q2 季度",
                                "title": "云原生微隔离上线",
                                "description": "实现容器网络级别的细粒度安全感知"
                            },
                            {
                                "date": "Q3 季度",
                                "title": "两地三中心演练",
                                "description": "模拟机房断电演习，验证业务不停摆"
                            }
                        ]
                    }
                ]
            }
        ]
    },
    
    chart: {
        "title": "业务运营指标与增长态势",
        "slides": [
            {
                "action": "blank_with_content",
                "title": "安全事件响应与效率提升大盘",
                "content_blocks": [
                    {
                        "type": "number_highlight",
                        "numbers": [
                            {
                                "value": "99.95%",
                                "label": "托管云核心服务可用性"
                            },
                            {
                                "value": "15分钟",
                                "label": "P0级故障平均恢复时间"
                            },
                            {
                                "value": "1.2万次",
                                "label": "拦截高风险外部网络入侵"
                            }
                        ]
                    },
                    {
                        "type": "chart",
                        "chart_type": "column",
                        "data": {
                            "title": "传统灾备 vs 深信服一键双活响应时间对比 (分钟)",
                            "categories": ["数据同步", "应用拉起", "业务接管"],
                            "series": [
                                {
                                    "name": "传统备份恢复",
                                    "values": [240, 180, 60]
                                },
                                {
                                    "name": "深信服托管双活",
                                    "values": [3, 5, 2]
                                }
                            ]
                        }
                    }
                ]
            }
        ]
    }
};

// 加载 Preset JSON 到编辑器
function loadPreset(key) {
    const code = presets[key];
    if (code) {
        document.getElementById('json-code').value = JSON.stringify(code, null, 2);
        log(`成功加载预设 JSON 示例: ${key}`, 'info');
    }
}

// 按钮动作：JSON 渲染 PPT
async function runJsonGeneration() {
    const jsonText = document.getElementById('json-code').value;
    let planData;
    
    try {
        planData = JSON.parse(jsonText);
    } catch (e) {
        log(`JSON 语法解析错误: ${e.message}。请核实逗号或双引号！`, 'error');
        alert(`JSON 语法解析错误: ${e.message}`);
        return;
    }
    
    log('准备发送 JSON 生成请求到服务器端...', 'info');
    showOverlay('正在渲染您的 PPT 配置计划...', '请稍候，服务器端正在克隆模板并绘制几何组件...');
    
    try {
        const response = await fetch('/api/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(planData)
        });
        
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({}));
            throw new Error(errJson.error || `HTTP 异常: ${response.status}`);
        }
        
        // 接收文件二进制流并下载
        const blob = await response.blob();
        triggerDownload(blob, "sangfor_presentation.pptx");
        log('🎉 PPT 生成成功，已开始下载！', 'info');
    } catch (e) {
        log(`PPT 生成失败: ${e.message}`, 'error');
        alert(`PPT 生成失败: ${e.message}`);
    } finally {
        hideOverlay();
    }
}

// 按钮动作：AI 描述一键生成
async function runAIGeneration() {
    const promptText = document.getElementById('ai-prompt').value.trim();
    if (!promptText) {
        alert('请输入关于 PPT 主题或大纲的文字描述！');
        return;
    }
    
    const apiKey = document.getElementById('llm-key').value.trim();
    const apiBase = document.getElementById('llm-base').value.trim();
    const apiModel = document.getElementById('llm-model').value.trim();
    
    // 保存到本地存储
    saveLlmConfig(apiKey, apiBase, apiModel);
    
    log('🚀 正在调用大语言模型 (LLM) 进行 PPT 排版大纲智能编排...', 'info');
    showOverlay('大模型正在编排大纲计划...', '这需要 10-30 秒，大模型正在规划每页幻灯片的组件和图标选择...');
    
    try {
        const response = await fetch('/api/ai-generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                prompt: promptText,
                api_key: apiKey,
                base_url: apiBase,
                model: apiModel
            })
        });
        
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({}));
            throw new Error(errJson.error || `HTTP 异常: ${response.status}`);
        }
        
        // 尝试从 Header 提取 AI 生成的原始大纲计划 (JSON 字符串)
        const encodedPlan = response.headers.get('X-Slides-Plan');
        if (encodedPlan) {
            try {
                const planJson = JSON.parse(decodeURIComponent(encodedPlan));
                log(`大模型大纲编排完成，页面结构如下:`, 'info');
                log(`  文稿标题: "${planJson.title}"`, 'info');
                planJson.slides.forEach((sl, idx) => {
                    log(`  第 ${idx+1} 页: [操作: ${sl.action}] ${sl.title || ''}`, 'info');
                });
            } catch (e) {
                // 解码失败不阻断下载
            }
        }
        
        const blob = await response.blob();
        triggerDownload(blob, "sangfor_ai_presentation.pptx");
        log('🎉 AI 描述直接生成 PPT 成功，已开始下载！', 'info');
    } catch (e) {
        log(`AI 直接生成 PPT 失败: ${e.message}`, 'error');
        alert(`AI 直接生成 PPT 失败: ${e.message}`);
    } finally {
        hideOverlay();
    }
}

// 触发浏览器下载 Blob 文件
function triggerDownload(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

// UI 遮罩层控制
function showOverlay(title, subtitle) {
    const overlay = document.getElementById('loading-overlay');
    const titleEl = document.getElementById('loading-text');
    const subEl = document.getElementById('loading-subtext');
    
    titleEl.innerText = title;
    subEl.innerText = subtitle;
    overlay.style.display = 'flex';
}

function hideOverlay() {
    const overlay = document.getElementById('loading-overlay');
    overlay.style.display = 'none';
}
