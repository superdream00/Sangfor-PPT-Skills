# Sangfor PPT Generator Skill - Handover & Explanation Guide

> **面向 AI 代理 / 模型的交接说明书**
> 
> 本文档旨在指导下一阶段接手的 AI 编码助手，帮助其快速理解项目现状、代码设计、已知策略，以及未来的优化方向。
> 
> **项目谱系与贡献历史 (Model Lineage & Contributions)**
> - **核心创建 (Phase 1-3)**: 由 **Claude Opus** 完成，涵盖原始模板 XML 解析、可替换页表元素提取（`page_catalog.json`）、以及 XML 节点深拷贝克隆与首行格式保留替换代码。
> - **重构与优化 (Phase 4-7)**: 由 **Gemini 3.5 Flash** 完成，涵盖 92% 冗余物理瘦身、rId 智能复制避免图片丢失与图表冲突、智能布局权重分摊引擎、演讲者备注通道、以及空白正文页使用 `'标题和内容'` 默认布局并去重蓝色小竖条的最终打磨与交付。

---

## 1. 项目基础信息

- **目标与核心痛点**: 本项目构建了一个自定义的 Antigravity IDE "Skill"，用于基于深信服（Sangfor）2024浅色 PPT 企业模板，程序化生成高保真的演示文稿。传统 AI 工具常因忽略模板背景/色彩/字体/特定排版而被抛弃，本项目采用“**幻灯片克隆 + 文本精确替换 + 空白页智能布局**”策略，实现 100% 品牌设计保真。
- **模板属性**: 尺寸 16:9 宽屏（33.87cm × 19.05cm）。共有 50 页预设，包含各种柱状图、饼图、卡片、两栏对比和数字高亮等。
- **开发语言与依赖**: Python 3.12+，依赖库为 `python-pptx` 与 `lxml`（用于 XML 树的深拷贝及属性注入）。

---

## 2. 目录结构与说明文件清单

全部项目文件均存放在 Skill 根目录 `sangfor-ppt-generator/` 下（路径相对于该目录）：

| 相对路径 | 类型 | 用途 |
|---|---|---|
| **[SKILL.md](SKILL.md)** | 文档 | **核心规范与技能定义**。包含全局字体/颜色/渐变/图形规范、50页目录细节、三模式触发条件、字数上限、质检清单及内容组合最佳实践。 |
| **[scripts/generate_ppt.py](scripts/generate_ppt.py)** | 源码 | **核心生成流程控制**。处理命令行参数、验证生成计划 (`validate_plan`)、执行幻灯片克隆、支持演讲者备注、遍历替换文本，以及通过关系解绑批量物理删除未用模板页（防臃肿）。 |
| **[scripts/utils.py](scripts/utils.py)** | 源码 | **工具函数与底层组件**。提供品牌色及字号常量、支持中英文双重字体的 `set_font`、支持透明度的形状绘制、智能权重布局 (`build_standard_page`)、表格生成以及原生图表样式美化。 |
| **[references/page_catalog.json](references/page_catalog.json)** | 配置 | 模板中 15 类关键页面（如封面、目录、卡片、数字等）对应的模板索引与可替换文本占位符正则映射。 |
| **[references/text_replacement_rules.json](references/text_replacement_rules.json)** | 配置 | 常见占位文字的映射规则。 |
| **templates/【常用】深信服--PPT浅色模板2024.pptx** | 资源 | 13MB 官方主模板文件。 |
| **[test_plan.json](test_plan.json)** | 配置 | 涵盖全部组件类型的多页测试生成计划。 |
| **[comp_analysis_plan.json](comp_analysis_plan.json)** | 配置 | 对桌面云竞争分析测试用的真实生成计划（6页）。 |
| **[.gitignore](.gitignore)** | 配置 | Git 忽略规则。 |

---

## 3. 核心设计机制与技术实现

接手模型需要理解以下关键实现，这关系到系统架构的正确运行：

### A. 克隆策略 (Slide Cloning with Relationship Remap)
- **原理**: `python-pptx` 本身不支持跨幻灯片克隆。我们通过 `lxml.etree` 进行幻灯片 `spTree`（形状树）与 `bg`（背景）的深拷贝。
- **关系映射**: 源幻灯片内部的图片和图表存在 `rId`（Relationship ID）指向具体的 `media` 或 `chart` 文件。直接复制 XML 会导致 `rId` 冲突或丢失。
- **改进实现**: 复制关系时，我们调用 `new_part.rels.get_or_add(rel.reltype, target_part)` 自动生成无冲突的新 `rId`，并维护一个 `rid_map` 映射字典，随后递归更新新幻灯片 XML 中所有对应的 `embed`、`link` 和 `id` 属性。

### B. 文本替换精度控制
- **问题**: 原版文本框中可能有多段文字，直接修改 `paragraph.text` 会丢失段落内部各 run 的精细格式（如粗体、字号）。
- **解决**: `_replace_text_preserve_format` 找到对应段落，如果是占位符，仅修改第一个 `run.text`，并清空其他 run；支持 `\n` 换行——对于换行内容会克隆底层段落 XML 节点（`_p`）进行拼接，从而实现完美的段落换行并继承首行格式。
- **特定实例替换**: 支持传递特定索引。例如，当页面有多个同名文本框时，可以通过指定 `shape_index` 仅替换指定的那一个（防止误替换）。

### C. 输出体积优化 (92% 空间瘦身)
- **策略**: 模板有 13MB 且含有大量背景图。如果直接生成，未使用的页虽然被删了，但底层的 media 部分仍然驻留在 zip 压缩包中。
- **实现**: 在 `remove_slides` 中，我们首先遍历被删除幻灯片的 `rId`，调用 `prs.part.drop_rel(rId)` 强行从底层关系链条中卸载。保存后，生成的文件大小约为 **1.0MB**。

### D. 空白页智能权重布局与去重竖条
- **背景**: 模板正文页面（如 `标题和内容`）的背景图片中已经预置了左上角的蓝色小竖条。
- **布局原理**: 早期版本在空白正文页上同时渲染了 programmatic 的蓝色竖条，导致了重影。已于最新版本中修复——**将默认空白页版式设置为 `'标题和内容'`**，同时**剔除了 `add_title_area` 内部绘制竖条的代码**，依靠版式背景实现完美的无冲突视觉。
- **高度分摊**: 引入 `BLOCK_WEIGHT_CONFIG`，根据内容块类型（图表=3.0，表格=2.0，文字=1.0）加权分摊可用高度空间（13.5厘米），并进行过大自动缩放，保证页面呼吸感。

---

## 4. 下一步优化空间建议

如果您想进一步优化该 Skill，可以朝以下三个方向迭代：
1. **多媒体复制的物理克隆 (Chart/Excel Data Deep Copy)**: 当前对于图表（Chart），为了避免版本不兼容，我们共享了同一个 Chart Part 指针（能正确呈现图表但修改一方数据可能影响全局）。您可以将其升级为通过 `zipfile` 复制内部的 `xl/embeddings/Microsoft_Excel_Worksheet.xlsx` 彻底解耦图表源数据。
2. **多栏智能断排 (Smart Auto-wrapping)**: 对于 `card_grid` 或 `bullet_list`，当单项文字明显超出建议字数限制时，提供自动断行或多卡片折行排列的动态机制。
3. **更丰富的 XML 文本渐变渲染**: 目前只有形状（Shape）能进行完美的渐变渲染。您可以实现针对文本 run 本身的渐变（如大号高亮数字的渐变色填充，需要构造文本属性的 `gradFill` XML 节点）。
