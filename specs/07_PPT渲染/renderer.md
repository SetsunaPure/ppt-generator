# PPT模板映射、内容渲染引擎与排版模板生成器

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 10.1 设计原则

shape索引映射配置化，替代硬编码 `shapes[0/1/2]`。换模板只改配置。

## 10.2 配置结构

```python
from dataclasses import dataclass
from typing import Optional

@dataclass(frozen=True)
class ShapeMapping:
    """单个shape的映射"""
    shape_index: int               # 在slide.shapes中的索引
    field: str                     # 填充字段: title/subtitle/header/content

@dataclass(frozen=True)
class SlideLayout:
    """一种slide类型的shape映射"""
    source_slide_index: int        # 模板中源slide索引
    shapes: list[ShapeMapping]     # shape映射列表
    auto_create_textbox: bool = False  # 是否自动创建正文文本框

SLIDE_LAYOUTS = {
    "title": SlideLayout(
        source_slide_index=0,
        shapes=[
            ShapeMapping(shape_index=1, field="title"),
            ShapeMapping(shape_index=2, field="subtitle"),
        ],
    ),
    "catalog": SlideLayout(
        source_slide_index=2,
        shapes=[
            ShapeMapping(shape_index=2, field="content"),
        ],
    ),
    "content": SlideLayout(
        source_slide_index=1,
        shapes=[
            ShapeMapping(shape_index=2, field="header"),
        ],
        auto_create_textbox=True,
    ),
}
```

## 10.3 文本框自动创建参数

```python
TEXT_BOX_CONFIG = {
    "left": Inches(0.2),
    "top": Inches(0.7),
    "width": Inches(10.8),
    "height": Inches(5),
    "margin": Pt(10),
    "font_name": "黑体",
    "word_wrap": True,
    "auto_size": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
}
```

---

## 11. 内容渲染引擎 (renderer/content_renderer.py)

## 11.1 设计原则：Block驱动渲染

渲染引擎接收分页器输出的 `List[List[Block]]`，每个Block已携带绝对坐标(left, top)和尺寸(width, height)。

**渲染策略按Block类型分道：**
- **行内元素(text/latex)**：写入文本框的段落流，文本框本身由布局引擎统一管理
- **块级元素(image/table)**：用绝对坐标直接放置到slide上，不嵌入段落流

这样图片/表格不会被"塞进"段落，而是像积木一样摆在确定位置。

## 11.2 混排协调原则（关键）

**问题**：表格/图片绝对定位在slide上，文本框段落流不知道它们占了多少空间，导致文字和块级元素重叠。

**解法**：布局引擎统一管理Y轴游标，渲染时按Y坐标顺序遍历所有Block，遇到块级元素就在文本框里插入**占位空段落**（精确高度=块级元素高度），让段落流"跳过"块级元素占用的空间。

```
示意（同一页内文字→表格→文字混排）：

文本框段落流：                slide绝对定位：
┌─────────────────┐
│ 文字段落1        │
│ 文字段落2        │
│ ████████ 空段落 █│  ← 占位高度=表格高度  │  [表格] 绝对坐标放在此处
│ ████ (预留空间) █│                         │
│ 文字段落3        │  ← 空段落之后继续       │
│ 文字段落4        │
└─────────────────┘

关键：空段落的高度精确等于块级元素的高度，Y轴对齐不重叠
```

## 11.3 渲染入口

```python
def render_page(slide, blocks: list[Block], env: EnvConfig,
                font_size_pt: int, formula_cache: dict[str, str],
                text_box_origin: tuple[int, int]):
    """
    渲染一页的所有Block，处理混排协调

    参数:
    - slide: python-pptx的Slide对象
    - blocks: 本页Block列表（已有left/top/width/height，按y坐标排序）
    - env: 环境配置
    - font_size_pt: 字号
    - formula_cache: 公式转换缓存
    - text_box_origin: (left_emu, top_emu) 文本框在slide上的起点坐标

    流程（按y坐标顺序遍历，保证混排对齐）：
    1. 将blocks按y坐标(top)排序，同y按x排序
    2. 遍历每个block：
       a) text/latex → 写入text_frame当前段落
       b) newline → 添加新段落
       c) image/table →
          - 在text_frame中插入占位空段落（space_before=0, line_spacing=精确高度）
          - 在slide上绝对定位放置该元素
       d) 记录当前y游标，确保后续行内元素从正确位置开始
    3. 所有块级元素绝对定位坐标 = text_box_origin + block.(left, top)
    """
```

## 11.4 占位空段落实现

```python
def add_placeholder_paragraph(text_frame, height_emu: int):
    """
    在text_frame中插入占位空段落，为块级元素预留空间

    实现：
    1. text_frame.add_paragraph() 创建空段落
    2. 设置段落行距为精确高度：
       paragraph.line_spacing = Emu(height_emu)  # 精确行距=块级元素高度
       paragraph.line_spacing_rule = MSO_LINE_SPACING.EXACTLY
    3. paragraph.space_before = Pt(0)
    4. paragraph.space_after = Pt(0)
    5. 段落内容为空（无run），纯占位
    """
```

## 11.5 行内元素 → 文本框段落流

```python
def render_inline_blocks(text_frame, inline_blocks: list[Block],
                         font_size_pt: int, formula_cache: dict[str, str]):
    """
    将行内Block写入text_frame的段落流

    规则：
    1. 遍历inline_blocks，按y坐标分行（同一行的Block的top值相同）
    2. 同一行的Block按x坐标排序，依次写入同一段落
    3. text Block → add_run() 设置文本内容
    4. latex Block → render_formula() 插入OMML
    5. 遇到不同y坐标（换行） → 添加新段落
    6. 合并相邻同类型text Block减少run碎片
    """
```

## 11.4 文本渲染

```python
def render_text(paragraph, block: Block, font_size_pt: int):
    """
    1. 用 lxml.html.clean.Cleaner 清洗HTML
    2. clean_html_for_pptx() 处理不兼容标签
    3. 构造 <a:r><a:t>text</a:t></a:r> XML插入段落
    4. 设置字体：黑体、font_size_pt大小
    5. block.is_bold时设置加粗
    """
```

## 11.5 公式渲染

```python
def render_formula(paragraph, omml_str: str):
    """
    将OMML XML包裹在 a14:m 标签中插入段落：
    <a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
        {omml_str}
    </a14:m>
    """
```

## 11.6 图片渲染（绝对坐标定位）

```python
def render_image(slide, block: Block, env: EnvConfig, text_box_origin: tuple[int, int]):
    """
    图片用绝对坐标放置到slide上，不嵌入段落流

    参数:
    - block: image类型的Block（含image_url, image_width_inches, image_height_inches, left, top）
    - text_box_origin: 文本框在slide上的起点(left_emu, top_emu)

    流程：
    1. 通过 make_request 下载图片（自动代理）
    2. 保存为临时文件
    3. 计算slide上的绝对坐标：
       absolute_left = text_box_origin[0] + block.left
       absolute_top = text_box_origin[1] + block.top
    4. slide.shapes.add_picture(
           image_path, absolute_left, absolute_top,
           Inches(block.image_width_inches), Inches(block.image_height_inches)
       )
    5. 图片尺寸由Block携带，不再收缩
    """
```

## 11.7 表格渲染（绝对坐标定位）

```python
def render_table(slide, block: Block, font_size_pt: int, text_box_origin: tuple[int, int]):
    """
    表格用绝对坐标放置到slide上

    参数:
    - block: table类型的Block（含table_data, table_rows, table_cols, left, top, width, height）

    流程：
    1. 计算slide上的绝对坐标：
       absolute_left = text_box_origin[0] + block.left
       absolute_top = text_box_origin[1] + block.top
    2. add_table(rows, cols, absolute_left, absolute_top, block.width, block.height)
    3. 按内容估算列宽（中文字符×2 + 其他×1）× font_size × 0.6
    4. 填充单元格内容（递归处理text/latex/image）
    5. 单行数据且多列时合并单元格
    6. 应用样式：黑体、16pt、居中、灰色背景(RGB 221,221,221)
    """
```

## 11.8 答案组渲染（AnswerGroupBlock）

```python
def render_answer_group(slide, block: Block, font_size_pt: int,
                        text_box_origin: tuple[int, int], formula_cache: dict):
    """
    答案组：内部已确定2列/4列布局，渲染时按列数分配X坐标

    参数:
    - block: answer_group类型的Block（含answer_group字段）
    - answer_group.columns: 2 或 4
    - answer_group.items: 答案项inline Block列表

    流程：
    1. 取 group = block.answer_group
    2. absolute_top = text_box_origin[1] + block.top
    3. 根据 columns 计算每个答案项的 X 坐标：
       - columns=4: 均分4列，每列宽度 = TEXT_BOX_WIDTH / 4
       - columns=2: 均分2列，每列宽度 = TEXT_BOX_WIDTH / 2
    4. 遍历 items，按列数和行号计算每个item的坐标：
       - col = i % columns
       - row = i // columns
       - item.left = text_box_origin[0] + col * (TEXT_BOX_WIDTH / columns)
       - item.top = absolute_top + row * group.row_height
    5. 每个item作为inline text写入text_frame对应段落
       （在文本框的段落流中，用空段落占位到目标行，再在目标位置写入答案文本）
    6. 字体：黑体、font_size_pt大小
    """
```

---

## 12. 排版模板生成器 (renderer/layout_generator.py)

## 12.1 流程

```python
def generate_layout_template(
    origin_template_path: Path,
    output_path: Path,
    pages: list[Page],
    logo_path: Optional[str],
    env: EnvConfig,
) -> None:
    """
    1. 解析logo路径（URL则通过make_request下载到临时文件，本地路径直用）
    2. 计算logo缩放尺寸（固定高度1.07cm，按原始宽高比算宽度）
    3. 遍历pages，按page.type从SLIDE_LAYOUTS取source_slide_index：
       - PageType.TITLE → source_slide_index=0
       - PageType.CATALOG → source_slide_index=2
       - 其他 → source_slide_index=1
    4. 复制对应源slide到新Presentation
    5. 每页右上角添加logo（left=幻灯片宽度-logo宽度-0.2in, top=0.2in）
    6. 保存为排版模板文件
    7. 清理临时logo文件（如果是URL下载的）
    """
```

## 12.2 Slide复制逻辑

```
1. 创建目标Presentation，设置与源相同的slide宽高
2. 使用源slide的slide_layout创建新slide
3. 深拷贝源slide的spTree子元素(sp/pic/grpSp/cxnSp)到新slide
4. 修复关系ID映射(rId)：遍历源slide的rels，建立old_rId→new_rId映射
5. 修复blip中的embed rId引用
```

---
