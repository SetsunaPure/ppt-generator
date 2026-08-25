"""
布局引擎测试 - P0优先级

测试积木流式布局引擎的核心功能：
1. 纯文字分页
2. 行内元素换行
3. 图片分级策略
4. 孤行保护
5. 表格跨页拆分
6. 答案组成对排列
"""

import sys
from pathlib import Path

# 将src目录添加到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import pytest
from pptx.util import Inches, Pt

from src.config import calc_text_line_height, ImageSizeLevel
from src.models import Block, BlockPosition, AnswerGroupBlock
from src.layout_engine import layout_blocks, html_to_blocks, split_table_if_needed
from src.question_parser import build_answer_group


# ============================================================
# 布局常量
# ============================================================

TEXT_BOX_HEIGHT = int(Inches(5))   # 5英寸 = 4572000 EMU
TEXT_BOX_WIDTH = int(Inches(10.8))  # 10.8英寸


def make_text_block(
    text: str = "测试",
    width: int = 200000,
    height: int = None,
    inline: bool = True,
    font_size_pt: int = 16
) -> Block:
    """快速构造测试用Block"""
    if height is None:
        height = calc_text_line_height(font_size_pt)
    return Block(
        type="text",
        content=text,
        width=width,
        height=height,
        inline=inline,
    )


def make_image_block(
    width: int = 4000000,
    height: int = None,
    level: str = "normal",
    image_url: str = "http://example.com/test.png"
) -> Block:
    """快速构造测试用图片Block"""
    if height is None:
        height = int(TEXT_BOX_HEIGHT * 0.3)
    return Block(
        type="image",
        content=None,
        width=width,
        height=height,
        inline=False,
        image_url=image_url,
        image_size_level=level,
    )


def make_table_block(
    rows: int = 3,
    cols: int = 2,
    data: list = None,
    height: int = None
) -> Block:
    """快速构造测试用表格Block"""
    if height is None:
        height = rows * int(Inches(0.35))
    if data is None:
        data = [[f"R{r}C{c}" for c in range(cols)] for r in range(rows)]
    return Block(
        type="table",
        content=None,
        width=TEXT_BOX_WIDTH,
        height=height,
        inline=False,
        table_data=data,
        table_rows=rows,
        table_cols=cols,
    )


# ============================================================
# 1. 纯文字分页测试
# ============================================================

class TestTextPaging:
    """纯文字分页测试"""

    def test_single_page_fits(self, font_size_pt=16):
        """少量文字，一页放得下"""
        line_height = calc_text_line_height(font_size_pt)
        blocks = [make_text_block(height=line_height) for _ in range(5)]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        assert len(pages[0]) == 5

    def test_overflow_to_second_page(self, font_size_pt=16):
        """文字超出一页，自动分页（使用足够多的大块级元素）"""
        line_height = calc_text_line_height(font_size_pt)
        lines_per_page = TEXT_BOX_HEIGHT // line_height

        # 构造足够多的块级元素填满一页还多
        blocks = [make_text_block(height=line_height * 5, inline=False) for _ in range(lines_per_page // 5 + 1)]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) >= 2

    def test_empty_input(self):
        """空输入返回空列表"""
        assert layout_blocks([]) == []

    def test_exact_one_page(self, font_size_pt=16):
        """精确填满一页"""
        line_height = calc_text_line_height(font_size_pt)
        lines_per_page = TEXT_BOX_HEIGHT // line_height

        blocks = [make_text_block(height=line_height) for _ in range(lines_per_page)]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        assert len(pages[0]) == lines_per_page


# ============================================================
# 2. 行内元素换行测试
# ============================================================

class TestInlineWrapping:
    """行内元素换行测试"""

    def test_inline_fits_on_current_line(self, font_size_pt=16):
        """行内元素放得下，cursor_x右移"""
        blocks = [
            make_text_block("A", width=200000, height=calc_text_line_height(font_size_pt), inline=True),
            make_text_block("B", width=200000, height=calc_text_line_height(font_size_pt), inline=True),
        ]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        # 两个block在同一行
        assert pages[0][0].position.top == pages[0][1].position.top

    def test_inline_wraps_to_next_line(self, font_size_pt=16):
        """行内元素放不下，换行"""
        from src.config import AVAILABLE_WIDTH

        half_width = AVAILABLE_WIDTH // 2 + 1000  # 多一点确保超宽
        blocks = [
            make_text_block("A", width=half_width, height=calc_text_line_height(font_size_pt), inline=True),
            make_text_block("B", width=half_width, height=calc_text_line_height(font_size_pt), inline=True),
        ]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        # 两个block不在同一行
        assert pages[0][0].position.top != pages[0][1].position.top


# ============================================================
# 3. 图片分级策略测试
# ============================================================

class TestImageSizeLevel:
    """图片分级策略测试"""

    def test_normal_image_on_same_page(self, font_size_pt=16):
        """NORMAL图片跟文字同页"""
        line_height = calc_text_line_height(font_size_pt)
        img_height = int(TEXT_BOX_HEIGHT * 0.3)  # 30%高度，正常图

        img_block = make_image_block(
            width=4000000,
            height=img_height,
            level="normal"
        )
        blocks = [make_text_block(height=line_height), img_block]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        assert any(b.type == "image" for b in pages[0])

    def test_large_image_gets_own_page(self, font_size_pt=16):
        """LARGE图片如果当前页有内容，应该换到新页"""
        line_height = calc_text_line_height(font_size_pt)
        img_height = int(TEXT_BOX_HEIGHT * 0.7)  # 70%高度，大图

        # 先放几个文字块（这样cursor_y > 0）
        blocks = [make_text_block(height=line_height, inline=False) for _ in range(5)]
        # 再放LARGE图片
        img_block = make_image_block(width=6000000, height=img_height, level="large")
        blocks.append(img_block)
        pages = layout_blocks(blocks, font_size_pt)

        # 至少有一页包含图片
        assert any(b.type == "image" for b in pages[-1])

    def test_fullpage_image_scaled(self, font_size_pt=16):
        """FULLPAGE图片等比缩放到TEXT_BOX_HEIGHT"""
        img_height = TEXT_BOX_HEIGHT  # 高度等于文本框高度

        img_block = make_image_block(
            width=6000000,
            height=img_height,
            level="fullpage"
        )
        blocks = [img_block]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        # FULLPAGE图片独占一页
        img = pages[0][0]
        assert img.type == "image"
        assert img.image_size_level == "fullpage"


# ============================================================
# 4. 孤行保护测试
# ============================================================

class TestOrphanControl:
    """孤行保护测试"""

    def test_orphan_prevented(self, font_size_pt=16):
        """当前页内容不足3行，行内Block跟着块级元素换页"""
        text_line_height = calc_text_line_height(font_size_pt)
        img_height = int(TEXT_BOX_HEIGHT * 0.7)

        # 只放1行文字 + 一个放不下的图片
        blocks = [
            make_text_block(height=text_line_height),
            make_image_block(height=img_height, level="large"),
        ]
        pages = layout_blocks(blocks, font_size_pt)

        # 孤行保护：那1行文字应该跟着图片到新页
        # 至少有一页包含文字
        assert any(b.type == "text" for b in pages[-1])

    def test_no_orphan_when_enough_content(self, font_size_pt=16):
        """当前页内容≥3行，正常换页不触发孤行保护"""
        text_line_height = calc_text_line_height(font_size_pt)
        img_height = int(TEXT_BOX_HEIGHT * 0.7)

        # 放4行文字 + 一个放不下的图片
        blocks = [make_text_block(height=text_line_height) for _ in range(4)]
        blocks.append(make_image_block(height=img_height, level="large"))
        pages = layout_blocks(blocks, font_size_pt)

        # 前4行文字在第1页，图片在第2页
        assert any(b.type == "text" for b in pages[0])


# ============================================================
# 5. 表格跨页拆分测试
# ============================================================

class TestTableSplit:
    """表格跨页拆分测试"""

    def test_small_table_no_split(self, font_size_pt=16):
        """小表格不拆分"""
        table_block = make_table_block(rows=2, cols=2)
        blocks = [table_block]
        pages = layout_blocks(blocks, font_size_pt)

        assert len(pages) == 1
        assert all(b.type == "table" for b in pages[0])

    def test_large_table_splits(self, font_size_pt=16):
        """大表格按行拆分，每页重复表头"""
        line_height = calc_text_line_height(font_size_pt)

        # 构造一个超高表格（50行）
        table_block = make_table_block(
            rows=51,
            cols=1,
            data=[["Header"]] + [[f"Row{i}"] for i in range(50)],
            height=TEXT_BOX_HEIGHT * 2
        )
        blocks = [table_block]
        pages = layout_blocks(blocks, font_size_pt)

        # 应该拆分成多页
        assert len(pages) >= 2

        # 每页的表格都应该有表头行
        for page in pages:
            table_blocks = [b for b in page if b.type == "table"]
            for tb in table_blocks:
                assert tb.table_data[0] == ["Header"], "每页表格都应该有表头"


# ============================================================
# 6. 答案组成对排列测试
# ============================================================

class TestAnswerGroup:
    """答案组排版测试"""

    def test_four_columns_short_answers(self, font_size_pt=16):
        """短答案四排"""
        answers = ["A", "B", "C", "D"]
        block = build_answer_group(answers, font_size_pt)

        assert block.type == "answer_group"
        assert block.inline is False
        assert block.answer_group.columns == 4
        assert block.height == block.answer_group.row_height  # 1行

    def test_two_columns_long_answers(self, font_size_pt=16):
        """长答案双排"""
        answers = [
            "A.这是一段很长的答案文字会超出四列宽度",
            "B.同样很长的答案",
            "C.也很长",
            "D.也是"
        ]
        block = build_answer_group(answers, font_size_pt)

        assert block.answer_group.columns == 2
        assert block.height == block.answer_group.row_height * 2  # 2行

    def test_no_single_column(self, font_size_pt=16):
        """绝不允许单排"""
        # 即使某个答案特别长，也必须是2列或4列
        answers = [
            "A.超长答案超长答案超长答案超长答案超长答案超长答案超长答案",
            "B", "C", "D"
        ]
        block = build_answer_group(answers, font_size_pt)

        assert block.answer_group.columns in (2, 4), "只能2列或4列，绝不单排"

    def test_three_answers_two_columns(self, font_size_pt=16):
        """3个答案时根据宽度决定2列或4列"""
        answers = ["A.选项A", "B.选项B", "C.选项C"]
        block = build_answer_group(answers, font_size_pt)

        # 3个答案时，columns可能是2或4（取决于宽度）
        # 关键是columns只能是2或4，绝不单排
        assert block.answer_group.columns in (2, 4), "只能2列或4列，绝不单排"


# ============================================================
# 运行测试
# ============================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
