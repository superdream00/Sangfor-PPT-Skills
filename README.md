# Sangfor PPT Generator Skill / 深信服 PPT 保真生成技能

基于深信服（Sangfor）2024 浅色企业 PPT 模板，**程序化、保真**生成演示文稿的 Agent Skill。核心策略：**幻灯片克隆 + 文本精确替换 + 空白页智能布局**，确保生成的 PPT 100% 保留模板的品牌字体、配色与排版。

> A code-driven Agent Skill that generates PowerPoint decks from the Sangfor 2024 light corporate template while preserving 100% of the brand's fonts, colors, and layout. Strategy: **slide cloning + format-preserving text replacement + weighted blank-page layout**.

---

## 为什么用它 / Why

市面上的 AI 工具导入企业模板后，常常丢失全局字体、配色和排版，生成的页面凌乱。本 Skill 直接在底层 OOXML 上克隆模板页、按品牌规范填充空白页，从根本上避免这个问题。

---

## 环境要求 / Requirements

- Python 3.10+（开发于 3.12）
- 依赖见 `requirements.txt`：`python-pptx`、`lxml`
- 字体：模板使用 **微软雅黑（Microsoft YaHei）**。Windows 自带；Linux/macOS 上若无该字体，PowerPoint/WPS 打开时会回退到默认字体（生成不报错，仅显示有差异）。

```bash
pip install -r requirements.txt
```

---

## 目录结构 / Layout

```
sangfor-ppt-generator/
├── SKILL.md                  # Skill 说明书（Anthropic Agent Skills 格式：触发词、设计规范、页表、三种模式）
├── README.md                 # 本文件
├── requirements.txt          # Python 依赖
├── scripts/
│   ├── generate_ppt.py       # 核心：克隆/替换/空白页填充/物理瘦身/校验
│   └── utils.py              # 工具库：品牌色字体常量、形状/表格/图表/卡片组件、权重布局引擎
├── references/
│   ├── page_catalog.json     # 50 页模板的 15 类页面索引与可替换占位符映射
│   └── text_replacement_rules.json
├── icons/                    # 172 个深信服蓝图标（Tabler Icons, MIT），7 大分类
│   ├── business/ cloud/ automotive/ chip/ electronics/ energy/ pharma/
│   └── icon_list.txt         # 图标清单
├── download_icons.py         # 图标下载+改色工具（从 Tabler CDN）
├── convert_icons_to_png.py   # SVG→512px透明PNG预转换工具
├── templates/
│   └── 【常用】深信服--PPT浅色模板2024.pptx   # 必需资源（Skill 依赖此模板运行）
├── test_plan.json            # 涵盖全部组件类型的示例生成计划
├── comp_analysis_plan.json   # 真实案例：桌面云竞争分析（6 页）
├── demo_icons.json           # 图标库演示：汽车行业云桌面方案（4 页）
└── output/                   # 生成产物（已 .gitignore）
```

---

## 快速开始 / Usage

在仓库根目录执行（命令依赖相对路径，请先 `cd` 到本目录）：

```bash
# 1) 从 JSON 计划生成完整 PPT
python scripts/generate_ppt.py \
  --template "templates/【常用】深信服--PPT浅色模板2024.pptx" \
  --plan test_plan.json \
  --output "output/result.pptx"

# 2) 单页快速克隆 + 文本替换
python scripts/generate_ppt.py \
  --template "templates/【常用】深信服--PPT浅色模板2024.pptx" \
  --clone 1 --replace "大标题大标题大标题大标题:实际标题" \
  --output "output/single.pptx"
```

生成计划（`plan.json`）的完整字段、`clone` / `blank_with_content` 两种 action、以及 `content_blocks` 支持的块类型（text / bullet_list / card_grid / table / chart / number_highlight / image / two_column / **timeline** / **process_steps** / **comparison_table** / icon_row），详见 **`SKILL.md`**。

---

## 场景化组件（第二期新增）★

新增 3 个高频业务场景组件：

1. **timeline** — 时间轴（水平/垂直）：展示里程碑、发展历程、项目阶段
2. **process_steps** — 流程步骤（带序号和箭头）：业务流程、操作步骤、实施路径
3. **comparison_table** — 对比表（支持高亮列）：竞品对比、方案选型、功能对比

演示：`python scripts/generate_ppt.py --template "templates/【常用】深信服--PPT浅色模板2024.pptx" --plan test_scene_components.json --output "output/test_scene_components.pptx"`

---

## 图标库 / Icon Library ★

内置 **172 个深信服蓝图标**（来源 [Tabler Icons](https://tabler.io/icons)，MIT 许可），覆盖 7 大场景：通用商务、云计算/IT、汽车、芯片、电子、新能源、创新药。

**第三期升级（2026）**：原生 PowerPoint 自定义几何矢量图形（iSlide 样式）
- **完全可编辑与改色**：图标在底层被直接转换为 PPT 原生的自由图形（`FREEFORM`）及自定义几何路径（`custGeom`）。用户在 PowerPoint 界面中只需选中图标，即可直接在顶部的“形状格式 -> 形状轮廓”中随意修改图标的颜色、线宽和线型。
- **100% 自包含**：生成的 PPTX 文件内不含任何 SVG 图像资源包（`media/image1.svg` 等），完全是轻量级的矢量数学指令，极大地降低了文件体积，避免了拷贝或发送到其他设备时出现“红叉”或“图片链接断开”的问题。
- **零外部依赖**：生成 PPT 时不需要任何图片渲染库，100% 纯 Python + python-pptx 原生 XML 解析支持。

三种用法（详见 `SKILL.md` 4.4 节）：
- `bullet_list` 列表项加 `"icon": "shield-lock"` → 图标替代圆点
- `card_grid` 卡片加 `"icon": "lock"` → 图标显示在标题上方
- 新增 `icon_row` 组件 → 水平排列"图标+标签"展示产品特性

演示：`python scripts/generate_ppt.py --template "templates/【常用】深信服--PPT浅色模板2024.pptx" --plan demo_icons.json --output "output/demo_icons.pptx"`

**维护与扩展图标**：若需增加图标，只需将 `.svg` 格式的图标放置到 `icons/` 相应的分类文件夹中，生成器在运行时会自动读取、解析其路径并实时绘制。

---

## 在其他 Agent 工具中使用 / Portability

**核心机制通用**：Python 脚本 + python-pptx，任何 AI 编程工具都能调用。
**适配其他工具只需 3 步**：

### 步骤 1：安装到本地

```bash
# 方法 A：从 GitHub 克隆（推荐）
git clone https://github.com/<你的用户名>/sangfor-ppt-generator.git
cd sangfor-ppt-generator

# 方法 B：解压 ZIP（从 GitHub Release 下载）
unzip sangfor-ppt-generator.zip
cd sangfor-ppt-generator

# 安装依赖
pip install -r requirements.txt
```

### 步骤 2：配置 Agent 工具识别

将 `AGENT_PROMPT.md` 的内容复制到对应工具的规则文件：

| Agent 工具 | 配置文件位置 | 操作 |
|-----------|-------------|------|
| **Cursor** | 项目根目录 `.cursorrules` | 复制 `AGENT_PROMPT.md` 全文到该文件 |
| **Cline (VS Code)** | 设置 → Extensions → Cline → Custom Instructions | 粘贴 `AGENT_PROMPT.md` 内容 |
| **Windsurf** | 项目根目录 `rules.md` | 复制 `AGENT_PROMPT.md` 全文 |
| **Claude Code** | 无需配置 | 原生支持 `SKILL.md`（本仓库已包含）|
| **其他工具** | 查阅对应工具的"自定义指令"文档 | 粘贴 `AGENT_PROMPT.md` |

> **`AGENT_PROMPT.md` 是什么？** 跨工具通用的使用说明，告诉 AI："当用户说'生成 PPT'时，执行 `python scripts/generate_ppt.py ...`"。

### 步骤 3：测试运行

在对应工具的对话框中输入：

```
生成一个汽车行业云桌面方案 PPT，要有封面、痛点分析、核心价值、尾页
```

Agent 应该自动：
1. 帮你编写 `plan.json`
2. 执行 `python scripts/generate_ppt.py --template ... --plan plan.json --output output/result.pptx`
3. 告诉你生成结果

---

### 跨工具使用注意事项

✅ **路径自适应**：脚本用 `__file__` 定位模板/图标，无需手动配置路径  
✅ **零额外依赖**：图标已预转换 PNG，不需要 cairosvg/svglib  
⚠️ **字体依赖**：需要微软雅黑（Windows 默认自带，macOS/Linux 用户需自行安装）  
⚠️ **模板文件**：13MB 模板文件必须随仓库一起（已在 `templates/` 下）

---

## 已知限制 / Known Limitations

- **图表 part 共享**：克隆图表页时多页共享同一个图表 part，在 PowerPoint 中修改一页图表数据可能影响其他页（纯展示无影响）。彻底解耦需复制内嵌 `xl/embeddings/*.xlsx`。
- 文本渐变（gradFill）目前仅形状支持，文字 run 暂不支持。
- 单项文字明显超字数时暂无自动断行/分页。

更多技术细节与开发历史见 `handover_readme.md` 与 `skill_creation_history.md`。
