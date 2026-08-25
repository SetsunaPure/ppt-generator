# 单元测试策略

> 来源：备课PPT生成-AI代码驱动规格书 v1.2

## 测试分层

| 优先级 | 模块 | 性质 | 测试方式 |
|--------|------|------|----------|
| P0 | 布局引擎 (content_splitter.py) | 纯函数，输入→输出 | pytest参数化 |
| P0 | HTML→Block转换 | 纯函数 | pytest参数化 |
| P1 | 公式批量转换 | 需mock LaTeX接口 | pytest + mock |
| P2 | PPT渲染 | 依赖python-pptx | 快照对比 |
| P3 | 业务服务层 | 集成测试 | 端到端 |

## P0：布局引擎测试

```python
import pytest
from parser.content_splitter import layout_blocks, Block, AnswerGroupBlock

# ── 布局常量（与content_splitter.py一致） ──
TEXT_BOX_HEIGHT = 5486400   # EMU
TEXT_BOX_WIDTH = 8686400

def make_text_block(text="测试", width=200000, height=200000, inline=True):
    """快速构造测试用Block"""
    return Block(type="text", content=text, width=width, height=height, inline=inline)

# ── 1. 纯文字分页 ──

class TestTextPaging:
    def test_single_page_fits(self):
        """少量文字，一页放得下"""
        blocks = [make_text_block(height=200000) for _ in range(5)]
        pages = layout_blocks(blocks)
        assert len(pages) == 1
        assert len(pages[0]) == 5

    def test_overflow_to_second_page(self):
        """文字超出一页，自动分页"""
        # 构造足够多的Block填满一页还多
        line_height = 200000
        lines_per_page = TEXT_BOX_HEIGHT // line_height
        blocks = [make_text_block(height=line_height) for _ in range(lines_per_page + 3)]
        pages = layout_blocks(blocks)
        assert len(pages) == 2
        assert len(pages[0]) == lines_per_page
        assert len(pages[1]) == 3

    def test_empty_input(self):
        """空输入返回空列表"""
        assert layout_blocks([]) == []

# ── 2. 行内元素换行 ──

class TestInlineWrapping:
    def test_inline_fits_on_current_line(self):
        """行内元素放得下，cursor_x右移"""
        blocks = [
            make_text_block("A", width=200000, height=200000, inline=True),
            make_text_block("B", width=200000, height=200000, inline=True),
        ]
        pages = layout_blocks(blocks)
        assert len(pages) == 1
        # 两个block在同一行
        assert pages[0][0].position.top == pages[0][1].position.top

    def test_inline_wraps_to_next_line(self):
        """行内元素放不下，换行"""
        half_width = TEXT_BOX_WIDTH // 2 + 1
        blocks = [
            make_text_block("A", width=half_width, height=200000, inline=True),
            make_text_block("B", width=half_width, height=200000, inline=True),
        ]
        pages = layout_blocks(blocks)
        assert len(pages) == 1
        # 两个block不在同一行
        assert pages[0][0].position.top != pages[0][1].position.top

# ── 3. 图片分级策略 ──

class TestImageSizeLevel:
    def test_normal_image_on_same_page(self):
        """NORMAL图片跟文字同页"""
        img_block = Block(
            type="image", content=None, inline=False,
            width=4000000, height=int(TEXT_BOX_HEIGHT * 0.5),
            image_size_level="normal"
        )
        blocks = [make_text_block(height=200000), img_block]
        pages = layout_blocks(blocks)
        assert len(pages) == 1

    def test_large_image_gets_own_page(self):
        """LARGE图片独占一页"""
        img_block = Block(
            type="image", content=None, inline=False,
            width=6000000, height=int(TEXT_BOX_HEIGHT * 0.7),
            image_size_level="large"
        )
        # 先放几个文字block
        blocks = [make_text_block(height=200000), img_block]
        pages = layout_blocks(blocks)
        # 文字在第1页，图片独占第2页
        assert len(pages) == 2
        assert any(b.type == "image" for b in pages[1])

    def test_fullpage_image_scaled(self):
        """FULLPAGE图片等比缩放到TEXT_BOX_HEIGHT"""
        img_block = Block(
            type="image", content=None, inline=False,
            width=6000000, height=TEXT_BOX_HEIGHT,
            image_size_level="fullpage"
        )
        blocks = [img_block]
        pages = layout_blocks(blocks)
        assert len(pages) == 1
        # 缩放后高度应≤TEXT_BOX_HEIGHT
        img = pages[0][0]
        assert img.height <= TEXT_BOX_HEIGHT

# ── 4. 孤行保护 ──

class TestOrphanControl:
    def test_orphan_prevented(self):
        """当前页内容不足3行，行内Block跟着块级元素换页"""
        text_line_height = 200000
        # 只放1行文字 + 一个放不下的图片
        blocks = [
            make_text_block(height=text_line_height),
            Block(type="image", content=None, inline=False,
                  width=6000000, height=int(TEXT_BOX_HEIGHT * 0.7),
                  image_size_level="large"),
        ]
        pages = layout_blocks(blocks)
        # 孤行保护：那1行文字应该跟着图片到新页
        # 第1页应该为空或没有内容
        assert any(b.type == "text" for b in pages[-1])  # 文字跟图片在一起

    def test_no_orphan_when_enough_content(self):
        """当前页内容≥3行，正常换页不触发孤行保护"""
        text_line_height = 200000
        # 放4行文字 + 一个放不下的图片
        blocks = [make_text_block(height=text_line_height) for _ in range(4)]
        blocks.append(Block(type="image", content=None, inline=False,
                           width=6000000, height=int(TEXT_BOX_HEIGHT * 0.7),
                           image_size_level="large"))
        pages = layout_blocks(blocks)
        # 前4行文字在第1页，图片在第2页
        assert any(b.type == "text" for b in pages[0])

# ── 5. 表格跨页拆分 ──

class TestTableSplit:
    def test_small_table_no_split(self):
        """小表格不拆分"""
        table_block = Block(
            type="table", content=None, inline=False,
            width=TEXT_BOX_WIDTH, height=400000,
            table_data=[["A", "B"], ["C", "D"]],
            table_rows=2, table_cols=2,
        )
        blocks = [table_block]
        pages = layout_blocks(blocks)
        assert len(pages) == 1

    def test_large_table_splits(self):
        """大表格按行拆分，每页重复表头"""
        # 构造一个超高表格
        table_block = Block(
            type="table", content=None, inline=False,
            width=TEXT_BOX_WIDTH, height=TEXT_BOX_HEIGHT * 2,
            table_data=[["Header"]] + [[f"Row{i}"] for i in range(50)],
            table_rows=51, table_cols=1,
        )
        blocks = [table_block]
        pages = layout_blocks(blocks)
        assert len(pages) >= 2
        # 每页的表格都应该有表头行
        for page in pages:
            table_blocks = [b for b in page if b.type == "table"]
            if table_blocks:
                assert table_blocks[0].table_data[0] == ["Header"]

# ── 6. 答案组成对排列 ──

class TestAnswerGroup:
    def test_four_columns_short_answers(self):
        """短答案四排"""
        answers = ["A", "B", "C", "D"]
        block = build_answer_group(answers, font_size_pt=16)
        assert block.type == "answer_group"
        assert block.inline is False
        assert block.answer_group.columns == 4
        assert block.height == block.answer_group.row_height  # 1行

    def test_two_columns_long_answers(self):
        """长答案双排"""
        answers = ["A.这是一段很长的答案文字会超出四列宽度", "B.同样很长的答案", "C.也很长", "D.也是"]
        block = build_answer_group(answers, font_size_pt=16)
        assert block.answer_group.columns == 2
        assert block.height == block.answer_group.row_height * 2  # 2行

    def test_no_single_column(self):
        """绝不允许单排"""
        # 即使某个答案特别长，也必须是2列
        answers = ["A.超长答案超长答案超长答案超长答案超长答案超长答案超长答案", "B", "C", "D"]
        block = build_answer_group(answers, font_size_pt=16)
        assert block.answer_group.columns in (2, 4)  # 只能2或4
```

## P0：HTML→Block转换测试

```python
import pytest
from parser.content_splitter import html_to_blocks, Block

class TestHtmlToBlock:
    def test_table_conversion(self):
        """<table> → Block(type=table)"""
        html = '<table><tr><td>A</td><td>B</td></tr></table>'
        blocks = html_to_blocks(html)
        table_blocks = [b for b in blocks if b.type == "table"]
        assert len(table_blocks) == 1
        assert table_blocks[0].table_rows == 1
        assert table_blocks[0].table_cols == 2

    def test_image_conversion(self):
        """<img> → Block(type=image)"""
        html = '<img src="http://example.com/test.png" style="width:200px;height:150px;">'
        blocks = html_to_blocks(html)
        img_blocks = [b for b in blocks if b.type == "image"]
        assert len(img_blocks) == 1
        assert img_blocks[0].image_url is not None

    def test_latex_conversion(self):
        """\\(...\\) → Block(type=latex)"""
        html = '<p>公式 \\(x^2+y^2=1\\) 继续</p>'
        blocks = html_to_blocks(html)
        latex_blocks = [b for b in blocks if b.type == "latex"]
        assert len(latex_blocks) == 1

    def test_fill_span(self):
        """<span class='fill'> → Block(content=_____)"""
        html = '<span class="fill">5</span>'
        blocks = html_to_blocks(html)
        fill_blocks = [b for b in blocks if b.type == "text" and "___" in (b.content or "")]
        assert len(fill_blocks) >= 1

    def test_br_newline(self):
        """<br/> → Block(type=newline)"""
        html = '第一行<br/>第二行'
        blocks = html_to_blocks(html)
        newline_blocks = [b for b in blocks if b.type == "newline"]
        assert len(newline_blocks) == 1
```

## P1：公式转换测试（mock LaTeX接口）

```python
import pytest
from unittest.mock import patch
from renderer.formula_batch import batch_convert_formulas

class TestFormulaConversion:
    @patch("renderer.formula_batch.make_request")
    def test_batch_convert_success(self, mock_request):
        """批量转换正常返回"""
        mock_request.return_value.status_code = 200
        mock_request.return_value.json.return_value = {"data": "<omml>...</omml>"}
        formulas = ["x^2", "y=mx+b"]
        result = batch_convert_formulas(formulas, env=EnvConfig(...))
        assert len(result) == 2

    @patch("renderer.formula_batch.make_request")
    def test_fallback_to_local(self, mock_request):
        """接口失败时回退到本地转换"""
        mock_request.side_effect = Exception("timeout")
        formulas = ["x^2"]
        result = batch_convert_formulas(formulas, env=EnvConfig(...))
        # 本地兜底应该也能返回结果
        assert len(result) == 1
```

## P2：渲染层快照测试

```python
import pytest
from renderer.ppt_builder import build_ppt

class TestRenderSnapshot:
    def test_snapshot_matches(self, tmp_path, snapshot):
        """生成的PPTX与基准快照对比"""
        output = tmp_path / "test_output.pptx"
        build_ppt(test_pages, output, env=test_env)
        if snapshot.update_mode:
            snapshot.update(output)
        else:
            assert output.read_bytes() == snapshot.path.read_bytes()
```

## 运行方式

```bash
# 全部测试
pytest tests/ -v

# 只跑P0布局引擎
pytest tests/test_layout_engine.py -v

# 只跑HTML转换
pytest tests/test_html_to_block.py -v

# 快照更新（改了渲染逻辑后）
pytest tests/test_render_snapshot.py --snapshot-update
```
