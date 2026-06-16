# 深信服 PPT 核心 Skill 开发编年史与技术内幕

> **交接文档（第二部分）**
> 
> 本文档记录了本 Skill 从原始模板开始，如何一步一步设计、测试、排坑并演进至当前版本的完整开发历史。它将帮助接手模型了解“我们是如何走到今天的”，以及每一个技术抉择背后的原因。
> 
> **模型开发历史与谱系 (Model Lineage & Contributions)**
> - **第一至第三阶段（起步与框架搭建）**: 由 **Claude Opus** 完成了模板的 XML 树初步解析、`page_catalog.json` 页表元数据设计、克隆与格式保留替换的核心逻辑编码。
> - **第四至第七阶段（优化、重构与漏洞修复）**: 由 **Gemini 3.5 Flash** 实现了 92% 体积物理瘦身（drop_rel）、图表/媒体 rId 深度重映射避免冲突、页面智能布局权重引擎、Speakers Notes 支持以及空白页 `'标题和内容'` 默认版式去重（修复左上角双竖条重影）的最终打磨与发布。

---

## 目录
1. [第一阶段：原始模板的深度剖析与特征提取](#1-第一阶段原始模板的深度剖析与特征提取)
2. [第二阶段：Skill 设计规范与配置元数据确立](#2-第二阶段skill-设计规范与配置元数据确立)
3. [第三阶段：克隆与替换的核心架构开发](#3-第三阶段克隆与替换的核心架构开发)
4. [第四阶段：文件瘦身与关系卸载的突破](#4-第四阶段文件瘦身与关系卸载的突破)
5. [第五阶段：空白页智能权重布局与组件库编写](#5-第五阶段空白页智能权重布局与组件库编写)
6. [第六阶段：健壮性改造、校验与演讲者备注](#6-第六阶段健壮性改造校验与演讲者备注)
7. [第七阶段：细节打磨（解决双竖条重影Bug）](#7-第七阶段细节打磨解决双竖条重影bug)
8. [总结：后续演进的思维导图](#总结后续演进的思维导图)

---

## 1. 第一阶段：原始模板的深度剖析与特征提取

### 1.1 背景任务
最初，用户提供了一个名为 `sangfor_template_2024.pptx` 的文件（大小约 13MB，共 50 页）。用户痛点是：**市面上的 AI 工具在导入该模板后，生成的新 PPT 完全忽略了深信服的全局字体、配色规范和排版，排版极其凌乱。**

### 1.2 编写分析脚本
我们编写了 `analyze_template.py` 和 `analyze_template_deep.py` 脚本，递归扫描了这 50 页 PPT 的底层 XML 结构。

```python
# 分析原理片段：提取形状与文本框类型，确定它是否使用占位符
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.is_placeholder:
            # 检查是否为标准 Placeholder
            print(f"Slide {i} has placeholder type: {shape.placeholder_format.type}")
```

### 1.3 核心扫描发现
1. **非占位符设计 (No Placeholders)**: 模板中的所有文字（标题、副标题、正文）都是用普通的 `TEXT_BOX`（文本框形状）硬编码堆叠而成的，**没有使用 PowerPoint 官方的 Slide Layout Placeholders**。
   - *影响*: 这意味着直接用 `python-pptx` 常见的 `slide.placeholders[0].text = ...` 会完全失效，必须使用 **“幻灯片 XML 克隆 + 逐个形状文本精确匹配替换”** 策略。
2. **字体体系 (Typography)**: 中文字体和英文字体均统一使用 `微软雅黑`。在 XML 底层表现为：
   - 拉丁字体: `<a:latin typeface="微软雅黑"/>`
   - 东亚字体: `<a:ea typeface="微软雅黑"/>`
3. **色彩方案 (Color Palette)**:
   - 主色蓝（深信服蓝）: `#006CD9`
   - 辅助深蓝（竖条及深色端）: `#003592` 和 `#00479D`
   - 强调绿（列表项圆点）: `#53C800`
   - 正文黑（主文本）: `#0E0E0E`（模板中出现达 240 次）
4. **渐变配置 (Gradients)**: 发现了 17 种微调渐变，主要是蓝绿渐变（G1-G9），用于章节分隔底条和高亮修饰。

---

## 2. 第二阶段：Skill 设计规范与配置元数据确立

有了分析数据后，我们开始确立 Skill 的格式，并在项目根目录下创建了三个指导/配置文件。

### 2.1 编写 `SKILL.md`
这是 Skill 的全局大纲与规则书。它向 IDE 的 AI 代理声明了如何触发 Skill，以及深信服品牌一致性所需的全局颜色常量、字体规格和渐变编号，为 AI 规划页面提供了直观的设计指南。

### 2.2 编写页面目录配置 (`page_catalog.json`)
我们将 50 页模板划分为 15 种类型（封面、目录、卡片网格、图表页、两栏对比等），提取出每一页可供替换的文本占位符正则。
例如：
```json
"cover": {
  "template_indices": [1, 2, 3],
  "replaceable_elements": [
    {
      "role": "main_title",
      "match_pattern": "大标题大标题大标题大标题"
    }
  ]
}
```

### 2.3 编写文本替换规则 (`text_replacement_rules.json`)
定义了将用户的自然语言需求转译为模板占位符的转换网格。

---

## 3. 第三阶段：克隆与替换的核心架构开发

这是项目中最核心的技术攻坚阶段，我们在 `generate_ppt.py` 中实现了这些核心逻辑。

### 3.1 幻灯片克隆 (Slide Cloning)
* **技术痛点**: `python-pptx` 官方库没有提供 `clone_slide` 的方法。
* **解决思路**:
  1. 通过 `prs.slides.add_slide(source_layout)` 创建一个干净的新页。
  2. 移除新页上所有由版式默认带入的占位符形状。
  3. 通过 `copy.deepcopy()` 将源幻灯片 XML 里的 `spTree`（形状树）和 `bg`（背景树）强行复制并附加到新幻灯片上。

### 3.2 攻克 rId 冲突（图片/图表引用丢失）
* **二次痛点**: 在复制 XML 后，发现图片变成了白块，图表直接损坏。
* **原因剖析**: XML 中的形状（如 Blip 贴图、Chart 图表）是通过 `rId`（如 `rId2`）关联到 `slide.xml.rels` 中的媒体资源的。直接深拷贝 XML 会将旧的 `rId` 带到新幻灯片中，而新幻灯片的关系字典中并没有配置这些 `rId`，或者它们指向了别的东西。
* **终极方案**:
  1. 重新设计 `_copy_slide_relationships_with_remap`。在克隆时，不再原样复制 `rId`，而是调用 `new_part.rels.get_or_add(rel.reltype, target_part)` 建立关系。
  2. `get_or_add` 会在内部生成一个全新且唯一的 `new_rid` 并返回。
  3. 建立 `rid_map = {old_rid: new_rid}` 关系映射表。
  4. 编写 `_remap_rids_in_element` 函数，递归遍历克隆后的 XML 树，将所有 `r:embed`、`r:link` 和 `r:id` 属性的值替换为新分配的 `rId`：
     ```python
     def _remap_rids_in_element(element, rid_map):
         for attr_name in ('embed', 'link', 'id'):
             full_attr = f'{{http://schemas.openxmlformats.org/officeDocument/2006/relationships}}{attr_name}'
             old_val = element.get(full_attr)
             if old_val in rid_map:
                 element.set(full_attr, rid_map[old_val])
     ```

### 3.3 文本替换不丢失格式
* **痛点**: 在 PowerPoint 中，一个文本框内的不同字符可以有不同的格式（如前三个字是蓝色，后面是黑色）。直接替换 `paragraph.text` 会导致段落里分裂的所有 `run` 节点被抹平，格式全部丢失。
* **方案**: 编写 `_replace_text_preserve_format`。在匹配到文本后：
  1. 计算替换后的完整字符串。
  2. 将新字符串赋值给段落的第一个 `run`（即 `runs[0].text = replaced_text`）。
  3. 将段落内的其余 `run` 全部清空（`run.text = ""`）。由于第一个 `run` 保留了该段落最核心的字体和颜色，从而实现了高保真的格式继承。
  4. **换行支持**: 如果新文本中含有 `\n`，则第一行替换当前段落，随后的每一行都通过克隆底层段落 XML 节点并调用 `paragraph._p.addnext(new_para)` 动态插入新段落，完美支持了列表或长正文的自动换行排版。

---

## 4. 第四阶段：文件瘦身与关系卸载的突破

### 4.1 痛点：生成的文件太大 (13MB 困局)
由于我们是从含有 50 页大图的原始模板开始生成的，最终就算删除了不需要的幻灯片，输出的 PPT 体积依然是 13MB。这对于网络传输极不友好。

### 4.2 原因剖析
在 PowerPoint 的 `.zip` 打包结构中，即使你从 `presentation.xml` 中删除了某页的 `sldId`，但只要该页关联的图片、媒体关系没有在关系层被卸载，PowerPoint 在保存时就不会自动垃圾回收这些 media 文件，它们依然躺在包内。

### 4.3 彻底清理逻辑
我们重构了 `remove_slides` 函数，先收集要删除幻灯片的 `rId`，显式调用 `prs.part.drop_rel(rId)` 强行切断关联，最后从 `sldIdLst` 中移除节点。
```python
# 核心清理动作
for sldId, rId in reversed(items_to_remove):
    if rId:
        prs.part.drop_rel(rId) # 强行解除 slide-part 关系，触发底层资源回收
    sldIdLst.remove(sldId)
```
* **效果**: 生成的 7 页测试 PPT 瞬间从 13,000KB 降为 **1,040KB**，成功剔除了所有未用大图！

---

## 5. 第五阶段：空白页智能权重布局与组件库编写

正文页的排版往往需要动态插入表格、列表或图表，不能全部依赖克隆。我们在 `utils.py` 中编写了组件库。

### 5.1 颜色与字体包装
在 `utils.py` 中声明了 `SangforColors` 和 `SangforFonts` 两个静态类，并在 `set_font` 函数中封装了底层的中英文字体 XML 设置（解决 `python-pptx` 中英文字体不同步的顽疾）。

### 5.2 列表、卡片、表格渲染
- `add_bullet_list`: 自动在文本左侧绘制绿色的圆形强调微章（`#53C800`，直径0.3cm），右侧自动创建多行格式化文本框。
- `add_card`: 程序化绘制卡片背景、卡片高亮标题条（高度 1.3cm，填充品牌蓝）和卡片正文。
- `add_table`: 根据输入的二维数组自动绘制表格，并应用交替奇偶行背景色 (`#F2F7FC` vs `#FFFFFF`) 和主蓝表头。

### 5.3 智能高度布局引擎
* **痛点**: 原来在空白页上堆叠多个组件（例如“一段文字”+“一个柱状图”）时，代码采取的是“按个数均分高度”策略。这导致文字文本框过高，而需要展示大量数据的柱状图高度被压缩到极扁，十分难看。
* **解决**: 在 `utils.py` 中定义了 `BLOCK_WEIGHT_CONFIG`，为不同组件配置了**视觉权重**与**最小高度限制**：
  ```python
  BLOCK_WEIGHT_CONFIG = {
      'text':             {'min_height': 1.5, 'weight': 1.0},
      'bullet_list':      {'min_height': 3.0, 'weight': 2.0},
      'card_grid':        {'min_height': 5.0, 'weight': 2.5},
      'table':            {'min_height': 4.0, 'weight': 2.0},
      'chart':            {'min_height': 7.0, 'weight': 3.0}, # 图表权重最大
  }
  ```
  根据当页所有 content_blocks 的权重比例，动态算出每个组件的实际高度。如果总高度超过页面可用范围（13.5cm），则按比例进行安全缩放，实现了高智能的页面自适应布局。

### 5.4 图表美化
- 重构了 `add_chart`。柱状图和折线图会自动启用 `plot.has_data_labels = True` 呈现数据数值；
- 自动为值轴添加了淡淡的灰色网格线（`#E0E0E0`，宽度 0.5pt）；
- 系列颜色完全使用深信服品牌推荐色序。

---

## 6. 第六阶段：健壮性改造、校验与演讲者备注

### 6.1 计划文件校验 (`validate_plan`)
新增了 `validate_plan`，在开始解析前做断言校验：
- 克隆的 `source_index` 必须在 0 - 49 之间。
- 不允许克隆 `SKIP_PAGES` (0, 4, 5) 等模板自带的使用规范和配色参考页。
- 检查 content_blocks 的键名合法性，确保 action 只有 `clone` 和 `blank_with_content`。

### 6.2 演讲者备注 (Speaker Notes)
在 Slide 遍历生成中，检查 slide 计划中是否有 `"notes"` 字段。如果有，调用底层 `slide.notes_slide.notes_text_frame.text = notes_text` 将其写入备注。这可供汇报人员在演播/演示者模式下读取。

---

## 7. 第七阶段：细节打磨（解决双竖条重影Bug）

### 7.1 痛点表现
在生成包含标准标题的正文空白页时，发现左上角的蓝色小竖条（高1.5cm，宽0.3cm）经常出现**重影、颜色变深或微微错位**的现象。

### 7.2 问题诊断
1. 模板中的 content layouts（如 `标题和内容`，索引 1）其背景图片（Master Background）中**已经预先绘制了**左上角的蓝色装饰小竖条。
2. 我们在 `utils.py` 的 `add_title_area` 函数中，又通过 `add_styled_rectangle` 在相同坐标上**程序化地画了一个形状竖条**。
3. 此外，空白页生成时的默认版式被设为了 `'标题幻灯片'`（Cover layout），这也是不合理的（内容页应该使用内容版式以确保继承正文页的底部横线等背景）。

### 7.3 解决方案
- **步骤 1**: 将 `generate_ppt.py` 中空白页的默认 layout 从 `'标题幻灯片'` 调整为 `'标题和内容'`（Layout 1）。这样空白页将完美继承深信服正文页的页眉页脚与背景图。
- **步骤 2**: 在 [utils.py](scripts/utils.py) 的 `add_title_area` 中，**去掉了创建竖条形状的代码**。
- **效果**: 彻底消除了重影，生成的空白正文页背景非常干净，自然融入模板背景。

---

## 总结：后续演进的思维导图

如果接手此项目的模型想要继续进行更深度的功能挖掘，我们建议以下方向：

```
                              ┌── 核心数据彻底解耦：ZipFile 深度解压复制内嵌的 xl/embeddings/*.xlsx
                              │
Sangfor PPT Skill 进一步优化 ──┼── 文字渐变填充：构造文字 Run 底层的 gradFill XML 节点
                              │
                              └── 文本过载智能处理：自动收缩行距或将过多的卡片自动换行/分页
```
