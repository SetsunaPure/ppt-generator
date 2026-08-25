"""
PPT生成器 - 积木流式布局引擎

核心模块：将HTML内容解析为Block列表，然后按积木流式布局算法分页

关键设计：
- 页满判断 = 空间耗尽：cursor_y + block.height > TEXT_BOX_HEIGHT
- 积木式流式布局：每个元素是一个带实际尺寸的Block，从左到右、从上到下排列
- 字号只决定单行高度，不影响"一页放多少"的判断
- 图片三级策略：NORMAL/LARGE/FULLPAGE
- 孤行保护：当前页内容不足3行时，内容跟随块级元素一起换页
- 表格跨页拆分：按行拆分，每页重复表头
"""

from typing import List, Optional, Any, Tuple
from dataclasses import dataclass, field

from bs4 import BeautifulSoup, NavigableString, Tag
from pptx.util import Inches, Pt, Emu

from .config import (
    TEXT_BOX_WIDTH,
    TEXT_BOX_HEIGHT,
    PADDING_LEFT_RIGHT,
    AVAILABLE_WIDTH,
    LINE_SPACING_FACTOR,
    calc_text_line_height,
    char_width_factor,
    ImageSizeLevel,
)
from .models import Block, BlockPosition, AnswerGroupBlock
from .utils import get_logger

logger = get_logger()


# ============================================================
# 字符宽度计算
# ============================================================

def calc_text_width(text: str, font_size_emu: int) -> int:
    """
    按字符宽度规则计算文本的实际宽度(EMU)

    规则见 config.char_width_factor
    """
    total = 0
    for ch in text:
        total += char_width_factor(ch) * font_size_emu
    return int(total)


# ============================================================
# HTML → Block 转换
# ============================================================

def split_text_block_by_width(
    text: str,
    font_size_emu: int,
    available_width: int,
    text_line_height: int,
    is_bold: bool = False,
    font_name: str = "黑体"
) -> List[Block]:
    """
    按字符宽度拆分文本块，确保每个Block宽度不超过available_width

    返回拆分后的Block列表，相邻行之间插入newline Block。
    适用于word_wrap=True的场景，保证布局引擎和渲染器对行数认知一致。

    算法：
    1. 从头开始贪心扫描
    2. 遇到超宽字符强制拆行
    3. 中文字符连续挤压时按比例估算
    """
    if not text:
        return []

    blocks: List[Block] = []
    current_line = ""
    current_width = 0

    for ch in text:
        char_width = int(char_width_factor(ch) * font_size_emu)

        # 检查是否需要换行
        if current_width + char_width > available_width and current_line:
            # 保存当前行
            blocks.append(Block(
                type="text",
                content=current_line,
                width=current_width,
                height=text_line_height,
                inline=True,
                is_bold=is_bold,
                font_name=font_name,
            ))
            # 插入换行符
            blocks.append(Block(
                type="newline",
                width=0,
                height=text_line_height,
                inline=False,
            ))
            # 开始新行
            current_line = ch
            current_width = char_width
        else:
            current_line += ch
            current_width += char_width

    # 保存最后一行
    if current_line:
        blocks.append(Block(
            type="text",
            content=current_line,
            width=current_width,
            height=text_line_height,
            inline=True,
            is_bold=is_bold,
            font_name=font_name,
        ))

    return blocks


# ============================================================
# HTML → Block 转换
# ============================================================

def html_to_blocks(
    html: str,
    font_size_pt: int = 16,
    env: Optional['EnvConfig'] = None
) -> List[Block]:
    """
    将HTML字符串解析为Block列表

    使用BeautifulSoup解析HTML，按标签类型生成Block对象

    参数:
        html: HTML字符串
        font_size_pt: 字号（用于计算文字行高和文本宽度）
        env: 环境配置（用于图片下载，可选）

    返回:
        Block列表
    """
    if not html or not html.strip():
        return []

    text_line_height = calc_text_line_height(font_size_pt)
    font_size_emu = Pt(font_size_pt)

    # BeautifulSoup解析，使用lxml以支持更完整的HTML解析
    soup = BeautifulSoup(html, 'lxml')

    # 预处理：将 $...$ 和 $$...$$ LaTeX 标记转为 <latex> 标签
    import re
    html_processed = html
    # 先处理 $$...$$ (display math)
    html_processed = re.sub(r'\$\$(.+?)\$\$', r'<latex>\1</latex>', html_processed, flags=re.DOTALL)
    # 再处理 $...$ (inline math)
    html_processed = re.sub(r'\$([^$\n]+?)\$', r'<latex>\1</latex>', html_processed)
    # 处理 \(...\) 格式
    html_processed = re.sub(r'\\\((.+?)\\\)', r'<latex>\1</latex>', html_processed, flags=re.DOTALL)

    soup = BeautifulSoup(html_processed, 'lxml')

    blocks: List[Block] = []
    pending_inline: List[Block] = []  # 同一行的行内元素

    def flush_inline_blocks():
        """将累积的行内元素合并"""
        nonlocal pending_inline
        if pending_inline:
            # 简单合并策略：将同类型文本合并
            merged_content = []
            merged_width = 0
            for block in pending_inline:
                if block.type == 'text' and merged_content and merged_content[-1].type == 'text':
                    merged_content[-1].content += block.content
                    merged_content[-1].width += block.width
                else:
                    merged_content.append(block)
                    merged_width += block.width
            blocks.extend(merged_content)
            pending_inline = []

    def process_element(element):
        """递归处理HTML元素"""
        nonlocal pending_inline

        if isinstance(element, NavigableString):
            # 纯文本节点
            text = str(element).strip()
            if text:
                width = calc_text_width(text, font_size_emu)
                # 检查宽度是否超过AVAILABLE_WIDTH，如果超了就拆分
                if width > AVAILABLE_WIDTH:
                    split_blocks = split_text_block_by_width(
                        text, font_size_emu, AVAILABLE_WIDTH, text_line_height
                    )
                    pending_inline.extend(split_blocks)
                else:
                    pending_inline.append(Block(
                        type="text",
                        content=text,
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
            return

        if not isinstance(element, Tag):
            return

        tag_name = element.name.lower() if element.name else ''

        # LaTeX标签处理（由预处理 $...$ 转换而来）
        if tag_name == 'latex':
            latex_content = element.get_text()
            if latex_content:
                width = calc_text_width(f"[{latex_content}]", font_size_emu)
                # LaTeX块也检查宽度（虽然通常LaTeX较短）
                if width > AVAILABLE_WIDTH:
                    split_blocks = split_text_block_by_width(
                        f"[{latex_content}]", font_size_emu, AVAILABLE_WIDTH, text_line_height,
                        is_bold=False, font_name="黑体"
                    )
                    # 将文本块标记为latex类型
                    for sb in split_blocks:
                        if sb.type == "text":
                            sb.type = "latex"
                    pending_inline.extend(split_blocks)
                else:
                    pending_inline.append(Block(
                        type="latex",
                        content=latex_content,
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
            return

        # 换行处理
        if tag_name in ('br', 'p', 'div'):
            if tag_name == 'br':
                flush_inline_blocks()
                blocks.append(Block(
                    type="newline",
                    width=0,
                    height=text_line_height,
                    inline=False,
                ))
            else:
                # 对于p和div，先flush当前行内元素，然后处理子元素
                flush_inline_blocks()
                # 添加换行
                blocks.append(Block(
                    type="newline",
                    width=0,
                    height=text_line_height,
                    inline=False,
                ))
                # 递归处理子元素
                for child in element.children:
                    process_element(child)
            return

        # 表格处理
        if tag_name == 'table':
            flush_inline_blocks()
            table_data = _parse_table(element)
            table_block = _create_table_block(table_data, text_line_height)
            if table_block:
                blocks.append(table_block)
            return

        # 图片处理
        if tag_name == 'img':
            # 检查是否为行内LaTeX公式图片（class="math-tex" 或有 latexdata 属性）
            class_attr = element.get('class', [])
            latexdata = element.get('latexdata', '')
            is_math_tex = ('math-tex' in class_attr) if isinstance(class_attr, list) else ('math-tex' in str(class_attr))
            
            if is_math_tex or latexdata:
                # 行内LaTeX公式图片 → 识别为行内LaTeX元素
                from urllib.parse import unquote
                latex_text = unquote(latexdata) if latexdata else ''
                if latex_text:
                    # 估算公式宽度（约为2个中文字符宽度）
                    width = int(2.0 * font_size_emu)
                    pending_inline.append(Block(
                        type="latex",
                        content=latex_text,
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
                else:
                    # 无latexdata，用占位文本
                    width = int(2.0 * font_size_emu)
                    pending_inline.append(Block(
                        type="text",
                        content="[公式]",
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
                return
            
            # 非公式图片 → 块级图片处理
            flush_inline_blocks()
            img_block = _create_image_block(element, env, text_line_height)
            if img_block:
                blocks.append(img_block)
            return

        # span处理
        if tag_name == 'span':
            class_attr = element.get('class', [])
            if 'fill' in class_attr:
                # 填空题标记 → 5个下划线
                flush_inline_blocks()
                fill_text = "_____"
                width = calc_text_width(fill_text, font_size_emu)
                if width > AVAILABLE_WIDTH:
                    # 填空题超宽也要拆分
                    split_blocks = split_text_block_by_width(
                        fill_text, font_size_emu, AVAILABLE_WIDTH, text_line_height
                    )
                    blocks.extend(split_blocks)
                else:
                    blocks.append(Block(
                        type="text",
                        content=fill_text,
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
                return
            elif 'longFill' in class_attr:
                # 长填空题 → 20个下划线
                flush_inline_blocks()
                fill_text = "____________________"
                width = calc_text_width(fill_text, font_size_emu)
                if width > AVAILABLE_WIDTH:
                    split_blocks = split_text_block_by_width(
                        fill_text, font_size_emu, AVAILABLE_WIDTH, text_line_height
                    )
                    blocks.extend(split_blocks)
                else:
                    blocks.append(Block(
                        type="text",
                        content=fill_text,
                        width=width,
                        height=text_line_height,
                        inline=True,
                    ))
                return
            elif 'brack' in class_attr:
                # 括号标记 → 空格
                blocks.append(Block(
                    type="text",
                    content=" ",
                    width=font_size_emu,
                    height=text_line_height,
                    inline=True,
                ))
                return

        # 加粗处理
        is_bold = tag_name in ('b', 'strong', 'u')
        sub_content = element.get_text()

        # 处理HTML实体
        sub_content = _decode_html_entities(sub_content)

        # 处理特殊实体
        if '&lt;' in str(element) or '&gt;' in str(element):
            # LaTeX输出
            blocks.append(Block(
                type="latex",
                content=sub_content,
                width=calc_text_width(sub_content, font_size_emu),
                height=text_line_height,
                inline=True,
                is_bold=is_bold,
            ))
            return

        # 处理LaTeX公式
        latex_pattern = element.find_all(string=lambda text: text and '\\(' in text)
        if latex_pattern or '\\(' in html:
            # 提取LaTeX公式
            import re
            latex_matches = re.findall(r'\\\((.+?)\\\)', str(element))
            for latex in latex_matches:
                blocks.append(Block(
                    type="latex",
                    content=latex,
                    width=calc_text_width(f"[{latex}]", font_size_emu),
                    height=text_line_height,
                    inline=True,
                ))
            if latex_matches:
                return

        # 上下标处理
        if tag_name in ('sub', 'sup'):
            prefix = ''
            suffix = ''
            if tag_name == 'sub':
                suffix = ''  # 下标直接跟在后面
            elif tag_name == 'sup':
                suffix = ''  # 上标直接跟在后面

            blocks.append(Block(
                type="text",
                content=sub_content,
                width=calc_text_width(sub_content, font_size_emu),
                height=text_line_height,
                inline=True,
                is_bold=is_bold,
            ))
            return

        # 递归处理子元素
        for child in element.children:
            process_element(child)

    # 遍历解析
    for element in soup.children:
        process_element(element)

    # 处理剩余的行内元素
    flush_inline_blocks()

    return blocks


def _decode_html_entities(text: str) -> str:
    """解码HTML实体"""
    entities = {
        '&nbsp;': ' ',
        '&lt;': '<',
        '&gt;': '>',
        '&amp;': '&',
        '&quot;': '"',
        '&#160;': ' ',
        '&#160;': '\u00a0',
        '&hellip;': '\u2026',
        '&ldquo;': '\u201c',
        '&rdquo;': '\u201d',
        '&lsquo;': '\u2018',
        '&rsquo;': '\u2019',
    }
    for entity, char in entities.items():
        text = text.replace(entity, char)
    return text


def _parse_table(table: Tag) -> List[List[str]]:
    """解析HTML表格为二维数组"""
    data = []
    rows = table.find_all('tr')
    for row in rows:
        cells = row.find_all(['td', 'th'])
        row_data = [cell.get_text(strip=True) for cell in cells]
        if row_data:
            data.append(row_data)
    return data


def _create_table_block(table_data: List[List[str]], line_height: int) -> Optional[Block]:
    """创建表格Block"""
    if not table_data:
        return None

    rows = len(table_data)
    cols = max((len(row) for row in table_data), default=0)

    # 计算表格高度：每行约0.35英寸
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
        table_cols=cols,
    )


def _create_image_block(
    img_tag: Tag,
    env: Optional['EnvConfig'],
    text_line_height: int
) -> Optional[Block]:
    """创建图片Block（分级处理）"""
    from .utils import get_image_dimensions, download_image_to_bytes, save_temp_image

    # 提取图片URL
    src = img_tag.get('src') or img_tag.get('data-src')
    if not src:
        return None

    # 获取图片尺寸
    if env:
        orig_w, orig_h = get_image_dimensions(src, env)
    else:
        # 默认尺寸（当无法获取时）
        orig_w, orig_h = 800, 600

    # 像素 → 英寸（96dpi）
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
    height_ratio = height_inches / tb_height_in if tb_height_in > 0 else 0

    if height_ratio >= 1.0:
        # FULLPAGE：高度超了，等比缩放到TEXT_BOX_HEIGHT
        level = ImageSizeLevel.FULLPAGE
        scale = tb_height_in / height_inches
        width_inches *= scale
        height_inches *= scale
    elif height_ratio >= 0.6:
        # LARGE：高度占60%以上，独占一页
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
        image_url=src,
        image_width_inches=width_inches,
        image_height_inches=height_inches,
        image_size_level=level.value if isinstance(level, ImageSizeLevel) else level,
    )


# ============================================================
# 积木流式布局引擎
# ============================================================

def layout_blocks(
    blocks: List[Block],
    font_size_pt: int = 16
) -> List[List[Block]]:
    """
    积木流式布局算法

    输入：Block列表，每个Block已有实际宽高
    输出：List[List[Block]]，每个子列表为一页的Block

    算法流程：
    1. 初始化布局状态
    2. 遍历Block列表，逐个放置（核心：空间判断，不行数判断）
    3. 每个Block记录绝对坐标(left, top)
    4. 返回分页结果

    关键认知：
    - 页满条件 = cursor_y + next_block.height > TEXT_BOX_HEIGHT
    - 不存在"每页几行"，因为图片/表格占的空间跟字号无关
    - 字号只影响 text/latex 行内元素换行时的 Y 轴增量(text_line_height)
    """
    if not blocks:
        return []

    text_line_height = calc_text_line_height(font_size_pt)
    orphan_threshold = 3 * text_line_height  # 孤行保护阈值

    # 初始化布局状态
    cursor_y = 0      # 当前Y坐标(EMU)，从0开始
    cursor_x = 0      # 当前X坐标(EMU)，从左开始
    current_page: List[Block] = []
    pages: List[List[Block]] = []

    def start_new_page():
        """开始新的一页"""
        nonlocal cursor_y, cursor_x, current_page, pages
        pages.append(current_page)
        current_page = []
        cursor_y = 0
        cursor_x = 0

    def get_current_page_inline_blocks() -> List[Block]:
        """获取当前页的行内Block"""
        return [b for b in current_page if b.inline]

    def check_and_handle_orphan(new_block: Block):
        """
        孤行保护检查

        场景：当前页尾部排了几行文字，紧接着块级元素放不下要换页。
        如果直接换页，上一页尾部只留1-2行文字 + 大片空白，排版很丑。

        解决：换页前回头检查，当前页在本次块级元素之前只有少量文字行
        （y游标 < 3 × text_line_height），就把这些文字也一起移到新页。
        """
        nonlocal cursor_y, current_page

        if cursor_y <= orphan_threshold:
            inline_blocks = get_current_page_inline_blocks()
            if inline_blocks:
                # 有孤行，从current_page中移除这些行内Block
                for ib in inline_blocks:
                    if ib in current_page:
                        current_page.remove(ib)
                return True  # 表示发生了孤行保护
        return False

    for block in blocks:
        if block.inline:
            # === 行内元素处理 ===
            remaining_width = AVAILABLE_WIDTH - cursor_x

            if block.width <= remaining_width:
                # 放得下，更新位置
                block.position = BlockPosition(left=cursor_x, top=cursor_y)
                cursor_x += block.width
                current_page.append(block)
            else:
                # 放不下，换行
                cursor_y += text_line_height
                cursor_x = 0

                # 检查是否需要换页
                if cursor_y + text_line_height > TEXT_BOX_HEIGHT:
                    start_new_page()

                # 放置
                block.position = BlockPosition(left=cursor_x, top=cursor_y)
                cursor_x = block.width
                current_page.append(block)

        else:
            # === 块级元素处理 ===
            # 图片分级处理
            if block.type == "image":
                img_level = block.image_size_level
                if img_level == ImageSizeLevel.LARGE.value or img_level == "large":
                    # LARGE图片：独占一页
                    if cursor_y > 0:
                        start_new_page()
                    # 图片之后的内容也换到新页
                    block.position = BlockPosition(left=0, top=0)
                    current_page.append(block)
                    start_new_page()
                    continue
                elif img_level == ImageSizeLevel.FULLPAGE.value or img_level == "fullpage":
                    # FULLPAGE图片：独占一页 + 居中
                    if cursor_y > 0:
                        start_new_page()
                    # 居中放置
                    abs_left = (TEXT_BOX_WIDTH - block.width) // 2
                    block.position = BlockPosition(left=abs_left, top=0)
                    current_page.append(block)
                    start_new_page()
                    continue

            # 表格跨页检查
            if block.type == "table":
                if block.height > TEXT_BOX_HEIGHT - cursor_y:
                    # 表格需要拆分
                    split_blocks = split_table_if_needed(block, TEXT_BOX_HEIGHT - cursor_y, text_line_height)
                    for sb in split_blocks:
                        if cursor_y + sb.height > TEXT_BOX_HEIGHT:
                            start_new_page()
                        sb.position = BlockPosition(left=0, top=cursor_y)
                        cursor_y += sb.height
                        current_page.append(sb)
                    continue

            # 普通块级元素
            if cursor_y + block.height > TEXT_BOX_HEIGHT:
                # 需要换页
                # 先检查孤行保护
                if check_and_handle_orphan(block):
                    # 孤行保护：行内Block已移到新页
                    pass
                else:
                    start_new_page()

                # 在新页放置
                cursor_y = 0

            # 放置块级元素（从行首开始）
            block.position = BlockPosition(left=0, top=cursor_y)
            cursor_y += block.height
            cursor_x = 0
            current_page.append(block)

    # 处理最后一页
    if current_page:
        pages.append(current_page)

    return pages


def split_table_if_needed(
    table_block: Block,
    remaining_height_emu: int,
    line_height: int
) -> List[Block]:
    """
    表格跨页拆分

    规则：
    1. 如果表格高度 ≤ remaining_height → 不拆分，返回原Block
    2. 如果表格高度 > remaining_height → 按行拆分：
       a) 第一页：放得下的行数（含表头），高度=行数×0.35in
       b) 后续页：每页放剩余行，首行重复表头
       c) 每个拆分后的Block都是独立的table Block，各自有正确的height
    3. 表头行在每个拆分块的table_data首行重复，确保每页表格都有表头
    """
    if not table_block.table_data or len(table_block.table_data) <= 1:
        return [table_block]

    available_rows = max(1, int(remaining_height_emu / Inches(0.35)))

    if table_block.table_rows <= available_rows:
        return [table_block]

    header = table_block.table_data[0]  # 表头行
    data_rows = table_block.table_data[1:]  # 数据行

    blocks = []
    current_row = 0
    current_page_rows = [header]
    current_page_height = Inches(0.35)  # 表头高度

    for i, row in enumerate(data_rows):
        row_height = Inches(0.35)
        if current_page_height + row_height > remaining_height_emu:
            # 当前页放不下，保存并开始新页
            blocks.append(Block(
                type="table",
                inline=False,
                width=table_block.width,
                height=int(current_page_height),
                table_data=current_page_rows,
                table_rows=len(current_page_rows),
                table_cols=max(len(r) for r in current_page_rows) if current_page_rows else 0,
            ))

            # 新页从表头开始
            current_page_rows = [header]
            current_page_height = Inches(0.35)
            remaining_height_emu = TEXT_BOX_HEIGHT  # 重置为整页高度

        current_page_rows.append(row)
        current_page_height += row_height

    # 保存最后一页
    if current_page_rows:
        blocks.append(Block(
            type="table",
            inline=False,
            width=table_block.width,
            height=int(current_page_height),
            table_data=current_page_rows,
            table_rows=len(current_page_rows),
            table_cols=max(len(r) for r in current_page_rows) if current_page_rows else 0,
        ))

    return blocks


# ============================================================
# 内容分页入口
# ============================================================

def split_page(
    html_content: str,
    font_size_pt: int = 16,
    env: Optional['EnvConfig'] = None
) -> List[List[Block]]:
    """
    内容分页入口函数

    将HTML内容解析为Block，然后按积木流式布局分页

    参数:
        html_content: HTML内容字符串
        font_size_pt: 字号
        env: 环境配置（用于图片处理）

    返回:
        List[List[Block]]，每页一个Block列表
    """
    # HTML → Block
    blocks = html_to_blocks(html_content, font_size_pt, env)

    # Block → 分页
    pages = layout_blocks(blocks, font_size_pt)

    return pages
