# 深信服 PPT 生成器 - Agent 工具通用指令
# Agent Prompt for Sangfor PPT Generator (Cross-tool Compatible)

当用户说出以下关键词时，识别为 PPT 生成需求并使用本工具：
- "生成 PPT" / "制作 PPT" / "做个演示文稿"
- "深信服 PPT" / "深信服模板"
- "产品赋能 PPT" / "方案 PPT"

---

## 核心能力

本工具基于深信服 2024 浅色 PPT 企业模板，提供两种生成模式：

1. **模板克隆模式**：精确复制模板中 50 页的任意页面，通过正则替换占位文字
2. **智能布局模式**：在空白页上用代码生成内容，支持 9 种组件类型 + 172 个深信服蓝图标

---

## 使用方式

### 步骤 1：用户需求 → 生成计划 JSON

当用户提出 PPT 需求时，帮用户编写生成计划 JSON 文件（如 `user_plan.json`），格式：

```json
{
  "title": "PPT 标题",
  "slides": [
    {
      "action": "clone",
      "source_index": 1,
      "replacements": {
        "大标题大标题大标题大标题": "实际标题",
        "小标题小标题小标题小标题小标题": "实际副标题"
      },
      "notes": "演讲者备注（可选）"
    },
    {
      "action": "blank_with_content",
      "title": "页面标题",
      "content_blocks": [
        {
          "type": "bullet_list",
          "items": [
            {"title": "要点1", "text": "说明文字", "icon": "shield-lock"},
            {"title": "要点2", "text": "说明文字", "icon": "rocket"}
          ]
        }
      ]
    }
  ]
}
```

### 步骤 2：执行生成命令

```bash
cd /path/to/sangfor-ppt-generator
python scripts/generate_ppt.py \
  --template "templates/【常用】深信服--PPT浅色模板2024.pptx" \
  --plan user_plan.json \
  --output "output/用户需求.pptx"
```

### 步骤 3：告知用户结果

生成成功后，告诉用户：
- 文件位置：`output/用户需求.pptx`
- 页数、包含的组件类型
- 建议打开检查效果

---

## 9 种内容组件（content_blocks）

| 组件类型 | 用途 | 必需字段 |
|----------|------|----------|
| `text` | 纯文字段落 | `text` |
| `bullet_list` | 带圆点/图标的列表 | `items: [{"title", "text", "icon"(可选)}]` |
| `card_grid` | 多卡片网格 | `cards: [{"header", "body", "icon"(可选)}]`, `columns` |
| `table` | 数据表格 | `data: [[row1], [row2], ...]` |
| `chart` | 图表（柱状/折线/饼图） | `chart_type`, `data: {"categories": [...], "series": [...]}` |
| `number_highlight` | 数字高亮展示 | `numbers: [{"value", "label"}]` |
| `image` | 图片插入 | `path`, `position` |
| `two_column` | 两栏布局 | `left: {...}`, `right: {...}` |
| `icon_row` | 图标+标签水平排列 | `items: [{"icon", "label"}]` |

---

## 图标库（172 个深信服蓝图标）

7 大分类，共 172 个：
- `business/` (48) — chart-line, target, shield, users, rocket, coin, clock, trophy, lock...
- `cloud/` (37) — cloud-computing, server, network, cpu, database, key, shield-check...
- `automotive/` (24) — car, truck, robot, battery-charging, shield-lock, gauge...
- `chip/` (20) — dna, atom, calculator, temperature, checklist, stack...
- `electronics/` (15) — device-laptop, qrcode, printer, scan...
- `energy/` (16) — battery, sun, wind, leaf, recycle, bulb...
- `pharma/` (12) — pill, vaccine, building-hospital, heartbeat...

**使用示例**：
```json
{"type": "bullet_list", "items": [
  {"title": "安全合规", "icon": "shield"},
  {"title": "高效敏捷", "icon": "clock"}
]}
```

图标会自动查找（跨 7 个分类），找不到时回退为绿色圆点（不报错）。

完整图标清单见 `icons/icon_list.txt`，或浏览 `icons/*/` 目录下的 `.png` 文件名。

---

## 模板克隆模式 - 50 页可用索引

深信服模板共 50 页，按用途分 15 类。常用页索引：

| 索引 | 类型 | 用途 |
|------|------|------|
| 1 | 封面 | 标题 + 副标题 + 日期 |
| 2-5 | 目录 | 1-4 级目录页 |
| 7, 9, 12, 14 | 过渡页 | 章节分隔 |
| 19, 21, 23, 25 | 卡片页 | 2-5 列卡片 |
| 32 | 时间轴 | 里程碑展示 |
| 40-43 | 数据对比 | 数字高亮 |
| 48 | 尾页 | "谢谢聆听" |

详细索引和可替换占位符见 `references/page_catalog.json`。

---

## 典型工作流示例

**场景：汽车行业云桌面解决方案（4 页）**

1. **用户需求**："做个汽车行业云桌面的产品方案 PPT，要有封面、痛点分析、核心价值、尾页"
2. **Agent 生成计划**：
   - 第 1 页：克隆模板索引 1（封面），替换标题为"汽车行业云桌面解决方案"
   - 第 2 页：空白页 + bullet_list（4 个痛点，每个配图标）
   - 第 3 页：空白页 + card_grid（3 张卡片展示价值，顶部配图标）+ icon_row（4 个特性）
   - 第 4 页：克隆模板索引 48（尾页）
3. **写入 `demo_icons.json`**（参考仓库里的示例）
4. **执行生成命令**
5. **告知用户**："已生成 4 页 PPT，位于 `output/汽车行业云桌面.pptx`，包含 15 张图标。"

---

## 错误处理

| 错误提示 | 原因 | 解决方法 |
|----------|------|----------|
| `模板文件不存在` | 模板路径错误 | 检查 `--template` 参数是否正确 |
| `图标 'xxx' 不存在` | 图标名拼写错误或不在库中 | 查看 `icons/icon_list.txt`，或让图标自动回退（不报错） |
| `source_index 超出范围` | 克隆的页码 >50 | 检查 `references/page_catalog.json` 的有效索引 |
| `ModuleNotFoundError: pptx` | 依赖未安装 | `pip install -r requirements.txt` |

---

## 质量检查清单（生成后提醒用户检查）

✅ 所有占位文字已替换（无"大标题大标题"残留）  
✅ 图标正确显示（无白框或错位）  
✅ 中文字体正常（微软雅黑）  
✅ 颜色符合深信服品牌（蓝 #006CD9、绿 #53C800）  
✅ 页数与计划一致  
✅ 演讲者备注已添加（如有）

---

## 技术要求

- **Python**: 3.8+
- **依赖**: `python-pptx>=0.6.21`, `lxml>=4.9.0`（见 `requirements.txt`）
- **字体**: 微软雅黑（中文）、Arial（英文）— Windows 默认自带
- **模板**: `templates/【常用】深信服--PPT浅色模板2024.pptx`（13MB，已在仓库）
- **图标**: `icons/` 目录下 172 个预转换 PNG（无需运行时依赖 cairosvg/svglib）

---

## 文档参考

完整技术规范见 `SKILL.md`（包含 50 页模板详细索引、颜色常量、字数上限等）。

---

**配置说明（针对不同 Agent 工具）**

- **Cursor**: 将本文件内容复制到项目根目录的 `.cursorrules` 文件
- **Cline**: 复制到 `AGENTS.md` 或在设置中添加自定义指令
- **Windsurf**: 复制到 `rules.md`
- **Claude Code**: 已原生支持 `SKILL.md`，无需额外配置
- **其他工具**: 复制到对应的 AI 指令/规则文件中

---

生成时记得 `cd` 到 `sangfor-ppt-generator/` 目录再执行命令，或使用绝对路径。
