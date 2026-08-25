# HTML内容分页器与积木流式布局引擎

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 7.1 设计原则

用 **BeautifulSoup4** 替代逐字符遍历，将HTML先解析为DOM树，再按节点类型生成**内容块(Block)**。

**核心架构变更：从"行计数分页"升级为"积木流式布局"**

旧方案问题：
- 固定18行/页是基于Pt(24)算出来的临时方案，字号变了就不对
- 图片收缩塞进去是为适配分页的无奈之举，排版效果差
- "每页几行"这个概念本身就是错的——混排场景下，文字、图片、表格共享垂直空间，不可能用行数来衡量页满

新方案思路：
- **页满判断 = 空间耗尽**：`cursor_y + next_block.height > TEXT_BOX_HEIGHT` 就换页，不存在"行数"概念
- **积木式流式布局**：每个元素是一个带实际尺寸的Block，从左到右、从上到下排列，塞不下就换行/换页
- **图片不收缩**：按原始比例缩放到文本框宽度，占多少高度就占多少，自然分页
- **字号只决定单行高度**：字号影响的是文字行的高度增量，不影响"一页放多少"的判断

## 7.2 布局常量与字号参数

```python
from pptx.util import Inches, Pt, Emu
from pptx.dml import Length

# ── 固定常量 ──
TEXT_BOX_WIDTH = Inches(10.8)                # 文本框宽度
TEXT_BOX_HEIGHT = Inches(5)                  # 文本框高度（一页的垂直空间上限）
PADDING_LEFT_RIGHT = Inches(0.1)             # 左右内边距
AVAILABLE_WIDTH = TEXT_BOX_WIDTH - PADDING_LEFT_RIGHT * 2  # 实际可用内容宽度

# ── 字号驱动的参数（仅用于计算文字行高） ──
LINE_SPACING_FACTOR = 1.3                    # 行距系数（1.3倍行距）

def calc_text_line_height(font_size_pt: int) -> int:
    """
    根据字号计算单行文字的高度(EMU)

    注意：这只是一个文字行的高度增量，不是"每页几行"。
    在混排中，图片/表格占用的垂直空间与字号无关。
    页满判断由布局引擎统一用 cursor_y vs TEXT_BOX_HEIGHT 决定。
    """
    return int(Pt(font_size_pt) * LINE_SPACING_FACTOR)  # EMU
```

## 7.3 内容块(Block)数据模型

```python
from dataclasses import dataclass, field
from typing import Optional, Any

@dataclass
class Block:
    """
    内容块 —— 积木流式布局的基本单元
    每个Block都有确定的宽高，由布局引擎按尺寸排列
    """
    type: str                   # "text" | "latex" | "image" | "table" | "newline"
    content: Any                # 具体内容数据

    # ── 尺寸（EMU单位） ──
    width: int = 0              # 块宽度
    height: int = 0             # 块高度

    # ── 行内/块级标记 ──
    inline: bool = True         # True=行内元素(可和别的行内元素共处一行)，False=独占一行的块级元素

    # ── 图片专用 ──
    image_url: Optional[str] = None
    image_width_inches: Optional[float] = None
    image_height_inches: Optional[float] = None
    image_size_level: Optional[str] = None     # "normal" | "large" | "fullpage"，布局引擎据此决定排页策略

    # ── 表格专用 ──
    table_data: Optional[list] = None
    table_rows: int = 0
    table_cols: int = 0

    # ── 文本专用 ──
    is_bold: bool = False
    font_name: str = "黑体"

    # ── 答案组专用 ──
    answer_group: Optional['AnswerGroupBlock'] = None   # 非None时，本Block是答案组的容器
```

## 7.4 字符宽度规则

| 字符类型 | 宽度系数 | Unicode范围 |
|---------|---------|-------------|
| 中文汉字 | 2.0 | \u4e00-\u9fff |
| 中文标点 | 2.0 | \u3000-\u303f |
| 全角符号 | 2.0 | \uff00-\uffef |
| 中文引号 | 2.0 | \u2018-\u201d |
| 中文间隔号\u00b7 | 2.0 | — |
| 中文省略号\u2026 | 2.0 | — |
| 中文破折号\u2014 | 2.0 | — |
| 带圈数字/字母 | 2.66 | \u2460-\u24ff |
| 其他字符 | 1.0 | 其余 |

## 7.5 HTML标签 → Block 转换规则

用BS4解析后，按标签类型生成Block对象：

| BS4元素 | 处理方式 | 产物Block |
|---------|---------|-----------|
| `<table>` | 提取表格数据，计算实际占高 | `Block(type="table", inline=False, height=行数×行高, width=TEXT_BOX_WIDTH)` |
| `<img>` | 提取src，按原始比例计算实际尺寸（宽度不超TEXT_BOX_WIDTH，按比例缩放） | `Block(type="image", inline=False, height=缩放后高度, width=缩放后宽度)` |
| `\(` ... `\)` | 提取LaTeX公式 | `Block(type="latex", inline=True, height=line_height, width=估算宽度)` |
| `<span class="fill">` | 替换为5个下划线 | `Block(type="text", inline=True, content="_____")` |
| `<span class="longFill">` | 替换为20个下划线 | `Block(type="text", inline=True, content="____________________")` |
| `<span class="brack">` | 替换为空格 | `Block(type="text", inline=True, content=" ")` |
| `<br/>` / `<p>` | 换行 | `Block(type="newline", inline=False, height=line_height)` |
| `<b>` / `<sub>` / `<sup>` | 保留文本内容，标记加粗/上下标 | `Block(type="text", inline=True, is_bold=True)` |
| `&nbsp;` | 空格 | `Block(type="text", inline=True, content=" ")` |
| `&lt;` | 转为LaTeX输出 `<` | `Block(type="latex", inline=True, content="<")` |
| `&ldquo;` / `&rdquo;` | 转为LaTeX输出对应引号 | `Block(type="latex", inline=True)` |
| `&hellip;` | 省略号 | `Block(type="text", inline=True, content="\u2026")` |
| `&#160;` | 短空格 | `Block(type="text", inline=True, content="\u00a0")` |
| `<u>...</u>` | 下划线内容替换为等量`_` | `Block(type="text", inline=True, content="_"*len)` |
| 纯文本 | 按字符宽度计算 | `Block(type="text", inline=True, width=总宽度emu)` |
| 选择题答案区 | 4个答案项包装为AnswerGroupBlock，内部决定2列/4列 | `Block(type="answer_group", inline=False, width=TEXT_BOX_WIDTH, height=行数×line_height)` |

## 7.6 积木流式布局算法（核心）

```
输入：HTML字符串, font_size_pt
输出：List[List[Block]]，每个子列表为一页的Block

算法流程：
1. calc_text_line_height(font_size_pt) → text_line_height (文字行高，仅文字行用)
2. BeautifulSoup解析HTML为DOM树
3. 遍历DOM节点，按7.5规则转为 Block 列表（此时每个Block已有实际宽高）
4. 初始化布局状态：
   - cursor_y = 0          # 当前Y坐标(EMU)，从0开始
   - cursor_x = 0          # 当前X坐标(EMU)，从左开始
   - current_page = []     # 当前页的Block列表
   - pages = []            # 所有页
5. 遍历Block列表，逐个放置（核心：空间判断，不行数判断）：

   a) 块级元素(inline=False)：图片/表格/换行
      - cursor_x = 0（新起一行）
      - 判断剩余垂直空间：TEXT_BOX_HEIGHT - cursor_y
      - block.height > 剩余空间 → 换页，但需要处理孤行问题：
        
        **孤行保护（Orphan Control）**：
        场景：当前页尾部排了几行文字，紧接着图片放不下要换页。
        如果直接换页，上一页尾部只留1-2行文字 + 大片空白，排版很丑。
        
        解决：换页前回头检查，当前页在本次块级元素之前只有少量文字行
        （y游标 < 3 × text_line_height），就把这些文字也一起移到新页。
        新页从头开始排文字，紧接着排图片，整体紧凑不割裂。

        ```
        换页逻辑：
        if cursor_y + block.height > TEXT_BOX_HEIGHT:
            # 需要换页
            orphan_threshold = 3 * text_line_height
            if cursor_y <= orphan_threshold and current_page有行内元素:
                # 孤行：当前页文字太少，全部移到新页
                从current_page中移除本次块级元素之前的行内Block
                cursor_y = 这些行内Block重新排列后的y值（新页从0开始）
            else:
                # 非孤行：当前页内容够多，正常换页
                cursor_y = 0
            新开一页
      - 放得下 → 设置Block.position(left=0, top=cursor_y)
                 cursor_y += block.height
                 cursor_x = 0

   b) 行内元素(inline=True)：文字/公式
      - 判断当前行剩余水平宽度：AVAILABLE_WIDTH - cursor_x
      - 放得下 → 设置Block.position(left=cursor_x, top=cursor_y)
                 cursor_x += block.width
      - 放不下 → 换行（cursor_y += text_line_height, cursor_x = 0）
                 判断剩余垂直空间：cursor_y + text_line_height > TEXT_BOX_HEIGHT
                 超出 → 换页（cursor_y=0）
                 设置Block.position(left=0, top=cursor_y)
                 cursor_x += block.width

6. 每个Block记录绝对坐标(left, top)
7. 合并相邻同类型inline text Block（减少渲染碎片）
8. 返回 pages

关键认知：
- 页满条件 = cursor_y + next_block.height > TEXT_BOX_HEIGHT
- 不存在"每页几行"，因为图片/表格占的空间跟字号无关
- 字号只影响 text/latex 行内元素换行时的 Y 轴增量(text_line_height)
- 孤行保护阈值 = 3 × text_line_height，当前页内容不足此阈值时，内容跟随块级元素一起换页
```

## 7.7 图片尺寸计算与分级策略

```python
from enum import Enum

class ImageSizeLevel(str, Enum):
    NORMAL = "normal"       # 常规图：缩放后高度 < 60% TEXT_BOX_HEIGHT，可跟文字同页
    LARGE = "large"         # 大图：缩放后高度 ≥ 60% TEXT_BOX_HEIGHT，独占一页
    FULLPAGE = "fullpage"   # 超大图：原始高度本身就超过TEXT_BOX_HEIGHT，居中撑满


def calc_image_block(image_url: str, env: EnvConfig, text_line_height: int) -> Block:
    """
    图片分级处理 —— 不是所有图片都一个待遇

    三级策略：
    1. NORMAL（常规图）：缩放后高度 < 60% TEXT_BOX_HEIGHT
       - 宽度缩到TEXT_BOX_WIDTH内，按比例算高度
       - 可以跟文字同页，正常流式排列

    2. LARGE（大图）：缩放后高度 ≥ 60% TEXT_BOX_HEIGHT
       - 同样按比例缩放
       - 但布局引擎必须让它独占一页（当前页已有内容就换新页，后续文字也换新页）
       - 不跟文字挤，图片在上文字在下看着也丑

    3. FULLPAGE（超大图）：按比例缩放后高度仍超过TEXT_BOX_HEIGHT
       - 继续等比缩放直到高度 = TEXT_BOX_HEIGHT（这是适配显示区域，不是收缩塞行）
       - 独占一页，slide上居中显示
    """
    from PIL import Image
    import io
    from utils.http import make_request

    resp = make_request("GET", image_url, env)
    img = Image.open(io.BytesIO(resp.content))
    orig_w, orig_h = img.size  # pixels

    # pixel → inch（96dpi）
    width_inches = orig_w / 96.0
    height_inches = orig_h / 96.0
    tb_width_in = TEXT_BOX_WIDTH / 914400
    tb_height_in = TEXT_BOX_HEIGHT / 914400

    # 第一步：宽度不超文本框
    if width_inches > tb_width_in:
        scale = tb_width_in / width_inches
        width_inches *= scale
        height_inches *= scale

    # 第二步：判断级别
    height_ratio = height_inches / tb_height_in

    if height_ratio >= 1.0:
        # FULLPAGE：高度超了，继续等比缩放到TEXT_BOX_HEIGHT
        level = ImageSizeLevel.FULLPAGE
        scale = tb_height_in / height_inches
        width_inches *= scale
        height_inches *= scale
    elif height_ratio >= 0.6:
        # LARGE：高度占60%以上，独占一页但不需继续缩放
        level = ImageSizeLevel.LARGE
    else:
        # NORMAL：正常图
        level = ImageSizeLevel.NORMAL

    return Block(
        type="image",
        inline=False,
        content=None,
        width=int(Inches(width_inches)),
        height=int(Inches(height_inches)),
        image_url=image_url,
        image_width_inches=width_inches,
        image_height_inches=height_inches,
        image_size_level=level,        # 新增字段，布局引擎据此决定排页策略
    )
```

**布局引擎对图片分级的处理**（在7.6算法步骤5.a中）：

```
遇到 image Block 时：

if block.image_size_level == NORMAL:
    # 常规图：跟其他元素正常流式排列
    剩余空间不够 → 换页
    
elif block.image_size_level == LARGE:
    # 大图：独占一页
    if cursor_y > 0:
        # 当前页已有内容 → 图片换到新页
        结束当前页，新开一页
    放置图片
    # 图片之后的内容也换到新页（大图下面挤文字不好看）
    结束当前页
    
elif block.image_size_level == FULLPAGE:
    # 超大图：独占一页 + 居中
    if cursor_y > 0:
        结束当前页，新开一页
    放置图片（居中定位，left居中计算）
    结束当前页
```

**图片居中定位**：

```python
def calc_image_center_position(block: Block, text_box_origin: tuple) -> tuple[int, int]:
    """
    FULLPAGE级别图片居中放置

    - horizontal: text_box_origin[0] + (TEXT_BOX_WIDTH - block.width) / 2
    - vertical: text_box_origin[1] + (TEXT_BOX_HEIGHT - block.height) / 2
    """
    abs_left = text_box_origin[0] + (TEXT_BOX_WIDTH - block.width) // 2
    abs_top = text_box_origin[1] + (TEXT_BOX_HEIGHT - block.height) // 2
    return abs_left, abs_top
```

## 7.8 表格尺寸计算

```python
def calc_table_block(table_data: list, line_height: int) -> Block:
    """
    计算表格Block的实际尺寸

    规则：
    1. 宽度 = TEXT_BOX_WIDTH（表格撑满）
    2. 高度 = 行数 × 行高（每行约0.35英寸，含内边距）
    3. 最少2行（表头+1行数据）
    """
    rows = len(table_data)
    row_height_inches = 0.35
    height_inches = max(rows, 2) * row_height_inches
    # 不超过文本框高度
    height_inches = min(height_inches, TEXT_BOX_HEIGHT / 914400)

    return Block(
        type="table",
        inline=False,
        width=int(TEXT_BOX_WIDTH),
        height=int(Inches(height_inches)),
        table_data=table_data,
        table_rows=rows,
        table_cols=max(len(row) for row in table_data) if table_data else 0,
    )
```

## 7.9 表格跨页拆分

当表格高度超过当前页剩余空间时，必须拆分表格为多个Block：

```python
def split_table_if_needed(table_block: Block, remaining_height_emu: int,
                          line_height: int) -> list[Block]:
    """
    表格跨页拆分

    规则：
    1. 如果表格高度 ≤ remaining_height → 不拆分，返回原Block
    2. 如果表格高度 > remaining_height → 按行拆分：
       a) 第一页：放得下的行数（含表头），高度=行数×0.35in
       b) 后续页：每页放剩余行，首行重复表头
       c) 每个拆分后的Block都是独立的table Block，各自有正确的height
    3. 表头行在每个拆分块的table_data首行重复，确保每页表格都有表头
    4. 拆分后每个Block的table_rows/height重新计算
    """
    available_rows = max(1, int(remaining_height_emu / Inches(0.35)))
    if table_block.table_rows <= available_rows:
        return [table_block]

    header = table_block.table_data[0]  # 表头行
    data_rows = table_block.table_data[1:]  # 数据行
    blocks = []

    # 第一页：表头 + 尽可能多的数据行
    first_chunk = [header] + data_rows[:available_rows - 1]
    remaining = data_rows[available_rows - 1:]

    # ... 后续页同理，每页开头重复表头
    return blocks
```

**布局引擎对表格跨页的处理**：在7.6的布局算法步骤5.a中，遇到table Block时：
1. 先判断剩余空间 `remaining = TEXT_BOX_HEIGHT - cursor_y`
2. 如果 `block.height > remaining` → 调用 `split_table_if_needed()` 拆分
3. 第一个拆分Block放当前页，后续Block分别放到新页

---
