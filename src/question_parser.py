"""
PPT生成器 - 题目解析器

处理题目HTML解析与答案排版
AnswerGroupBlock成对排列：答案只允许双排或四排，绝不单排
"""

import json
import logging
from typing import Optional, List, Any, Tuple
from dataclasses import dataclass

from pptx.util import Pt, Inches

from .config import TEXT_BOX_WIDTH, calc_text_line_height
from .models import Block, AnswerGroupBlock, QuestionData
from .compat import md_to_text

logger = logging.getLogger("ppt_generator")


# ============================================================
# 答案格式化
# ============================================================

def format_answers(answers: Any, ques_type: Optional[str]) -> Optional[str]:
    """
    格式化答案文本

    - answers为JSON字符串时先json.loads解析为list
    - 判断题(ques_type=="判断题")：1→"对", 0→"错"
    - 二维数组：每行逗号拼接，行间用";&nbsp;"分隔
    - 一维数组：逗号拼接
    - None/空列表 → 返回None
    """
    if answers is None or answers == "":
        return None

    # JSON字符串解析
    if isinstance(answers, str):
        try:
            answers = json.loads(answers)
        except json.JSONDecodeError:
            return answers

    # 判断题
    if ques_type == "判断题":
        if isinstance(answers, list) and len(answers) > 0:
            return "对" if answers[0] == 1 else "错"
        return None

    # 二维数组
    if isinstance(answers, list) and answers and isinstance(answers[0], list):
        rows = []
        for row in answers:
            rows.append("，".join(str(a) for a in row))
        return "； ".join(rows)

    # 一维数组
    if isinstance(answers, list):
        return "，".join(str(a) for a in answers)

    return str(answers)


# ============================================================
# 答案组构建（关键：双排/四排，绝不单排）
# ============================================================

def calc_text_width(text: str, font_size_emu: int) -> int:
    """
    按字符宽度规则计算文本的实际宽度(EMU)
    """
    from .config import char_width_factor
    total = 0
    for ch in text:
        total += char_width_factor(ch) * font_size_emu
    return int(total)


def build_answer_group(answers_list: List[str], font_size_pt: int = 16) -> Block:
    """
    选择题答案只允许双排或四排，不允许单排

    决策逻辑：
    1. 计算每个答案项的实际宽度
    2. 计算一行放4个的总宽度：sum(4个宽度) + 3×间隔
    3. 4个放得下 → 四排(columns=4, 1行)
    4. 4个放不下 → 双排(columns=2, 2行，每行2个)
    5. 生成AnswerGroupBlock，包装成一个块级Block交给布局引擎

    布局引擎视角：这就是一个带确定高度和宽度的普通Block，
    不需要知道内部是答案，正常流式排列即可
    """
    text_line_height = calc_text_line_height(font_size_pt)
    font_size_emu = Pt(font_size_pt)
    gap_width = int(font_size_emu * 4)  # 答案间隔4空格宽度

    # 1. 生成每个答案项的inline Block
    items = []
    for i, answer in enumerate(answers_list):
        text = f"{i+1}.{answer}"
        width = calc_text_width(text, font_size_emu)
        items.append(Block(
            type="text",
            inline=True,
            content=text,
            width=width,
            height=text_line_height,
        ))

    # 2. 判断四排还是双排
    total_items_width = sum(item.width for item in items)
    total_gaps = gap_width * (len(items) - 1)
    four_col_width = total_items_width + total_gaps

    if four_col_width <= TEXT_BOX_WIDTH:
        # 四排：一行放4个
        columns = 4
        rows = 1
    else:
        # 双排：每行2个，拆成2行
        # 注意：即使答案数量是1或3，也按2列排，确保美观
        columns = 2
        rows = (len(items) + columns - 1) // columns  # 向上取整

    group = AnswerGroupBlock(
        items=items,
        columns=columns,
        row_height=text_line_height,
        total_height=rows * text_line_height,
        total_width=TEXT_BOX_WIDTH,
    )

    # 3. 包装为块级Block
    return Block(
        type="answer_group",
        inline=False,
        content=group,
        width=TEXT_BOX_WIDTH,
        height=group.total_height,
        answer_group=group,
    )


# ============================================================
# 题目解析
# ============================================================

def parse_question(question_vo: dict, is_origin: bool = True) -> QuestionData:
    """
    从questionVo中提取题目数据

    is_origin=True (备课场景):
        answers = question_vo["explain"][0]["answers"]
        analysis = question_vo["explain"][0]["analysis"]
        ques_type = question_vo.get("quesType", {}).get("name")

    is_origin=False (讲题场景):
        answers = question_vo["explain"]["answers"]
        analysis = question_vo["explain"]["analysis"]
        ques_type = ""

    通用:
        stem = question_vo.get("context", {}).get("stem")
        options = question_vo.get("context", {}).get("options")

    递归: children = question_vo.get("children")，逐个递归parse_question
    """
    if is_origin:
        # 备课场景
        explain_list = question_vo.get("explain", [])
        if explain_list and len(explain_list) > 0:
            explain = explain_list[0]
        else:
            explain = {}
        answers = explain.get("answers", [])
        analysis = explain.get("analysis", "")
        ques_type = question_vo.get("quesType", {}).get("name", "")
    else:
        # 讲题场景
        explain = question_vo.get("explain", {})
        answers = explain.get("answers", [])
        analysis = explain.get("analysis", "")
        ques_type = ""

    # 解析题干
    context = question_vo.get("context", {})
    stem = context.get("stem", "")
    options = context.get("options", [])

    # 格式化答案
    answer_str = format_answers(answers, ques_type)

    # 递归处理子题
    children = []
    for child_vo in question_vo.get("children", []):
        children.append(parse_question(child_vo, is_origin))

    return QuestionData(
        stem=stem,
        options=options,
        answer=answer_str,
        analysis=analysis,
        children=children,
        ques_type=ques_type
    )


# ============================================================
# 答案元素生成
# ============================================================

def create_answer_element(answer: str, analysis: str, children: list) -> str:
    """
    生成答案格式HTML

    返回格式：【答案】<br />{answer}<br />【解析】<br />{analysis}

    子题处理：
    - 每个子题编号 (1) (2) ...
    - 子子题用带圈数字 ①②... (超过20个用(21)(22)...)
    - 无子题时 answer/analysis 前加 &nbsp;
    """
    parts = []

    # 答案部分
    if answer:
        parts.append(f"【答案】<br />{answer}")
    else:
        parts.append("【答案】<br />&nbsp;")

    # 解析部分
    if analysis:
        parts.append(f"【解析】<br />{analysis}")
    else:
        parts.append("【解析】<br />&nbsp;")

    # 子题处理
    if children:
        sub_parts = []
        for i, child in enumerate(children):
            if i < 20:
                # 带圈数字
                circled = chr(0x2460 + i) if i < 20 else f"({i+1})"
            else:
                circled = f"({i+1})"

            child_answer = format_answers(child.answer, child.ques_type) if child.answer else "&nbsp;"
            sub_parts.append(f"{circled} {child_answer}")

        if sub_parts:
            parts.append("【子题答案】<br />" + "<br />".join(sub_parts))

    return "<br />".join(parts)


# ============================================================
# 题目内容组装
# ============================================================

def assemble_question_content(
    question: QuestionData,
    include_answer: bool = True,
    include_analysis: bool = True
) -> str:
    """
    组装题目完整内容为HTML

    包含：题干 + 选项 + 答案 + 解析
    """
    html_parts = []

    # 题干
    if question.stem:
        html_parts.append(question.stem)

    # 选项
    if question.options:
        options_html = "<br />".join(
            f"{chr(65+i)}.{opt}" for i, opt in enumerate(question.options)
        )
        html_parts.append(options_html)

    # 答案和解析
    if include_answer or include_analysis:
        answer_element = create_answer_element(
            question.answer or "",
            question.analysis or "",
            []
        )
        html_parts.append(answer_element)

    return "<br />".join(html_parts)


def get_options_for_layout(options: List[str], font_size_pt: int = 16) -> Block:
    """
    将选项转换为布局用的Block

    适用于选择题选项的排版处理
    """
    if not options:
        return None

    # 选项作为纯文本处理
    text_line_height = calc_text_line_height(font_size_pt)
    font_size_emu = Pt(font_size_pt)

    blocks = []
    for i, opt in enumerate(options):
        text = f"{chr(65+i)}.{opt}"  # A.B.C.D...
        width = calc_text_width(text, font_size_emu)
        blocks.append(Block(
            type="text",
            content=text,
            width=width,
            height=text_line_height,
            inline=True,
        ))

    # 如果有4个选项，尝试四排
    if len(options) == 4:
        # 计算四排是否放得下
        total_width = sum(b.width for b in blocks)
        gap_width = int(font_size_emu * 4) * 3
        if total_width + gap_width <= TEXT_BOX_WIDTH:
            # 四排可行
            group = AnswerGroupBlock(
                items=blocks,
                columns=4,
                row_height=text_line_height,
                total_height=text_line_height,
                total_width=TEXT_BOX_WIDTH,
            )
            return Block(
                type="answer_group",
                inline=False,
                content=group,
                width=TEXT_BOX_WIDTH,
                height=group.total_height,
                answer_group=group,
            )

    # 否则作为普通文本块返回第一项，后续按需处理
    return blocks[0] if blocks else None
