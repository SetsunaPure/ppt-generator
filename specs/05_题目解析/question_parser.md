# 题目解析器与答案排版

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 8.1 题目数据模型

```python
from dataclasses import dataclass, field
from typing import Optional, List

@dataclass
class QuestionData:
    stem: str                           # 题干HTML
    options: Optional[List[str]]        # 选项列表
    answer: Optional[str]               # 答案文本
    analysis: Optional[str]             # 解析文本
    children: List['QuestionData'] = field(default_factory=list)  # 子题列表
    ques_type: Optional[str] = None     # 题型名称
```

## 8.2 解析规则

```python
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
```

## 8.3 答案格式化

```python
def format_answers(answers, ques_type: Optional[str]) -> Optional[str]:
    """
    - answers为JSON字符串时先json.loads解析为list
    - 判断题(ques_type=="判断题")：1→"对", 0→"错"
    - 二维数组：每行逗号拼接，行间用";&nbsp;"分隔
    - 一维数组：逗号拼接
    - None/空列表 → 返回None
    """
```

## 8.4 选择题答案排版（AnswerGroupBlock成对排列）

```python
@dataclass
class AnswerGroupBlock:
    """
    选择题答案组 —— 保证答案只出现双排或四排，绝不单排
    答案项成对排列：一行放得下4个就四排，放不下4个就拆成2+2双排
    """
    items: list[Block]          # 每个答案项的inline Block
    columns: int                # 2 或 4，由宽度计算决定
    row_height: int             # 单行高度(text_line_height)
    total_height: int           # 总高度 = 行数 × row_height
    total_width: int            # = TEXT_BOX_WIDTH（撑满文本框宽度）

def build_answer_group(answers_list: list[str], font_size_pt: int) -> Block:
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
        columns = 2
        rows = (len(items) + 1) // 2

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


def calc_text_width(text: str, font_size_emu: int) -> int:
    """
    按字符宽度规则计算文本的实际宽度(EMU)
    宽度系数表见7.4节
    """
    total = 0
    for ch in text:
        total += char_width_factor(ch) * font_size_emu
    return int(total)
```

## 8.5 答案元素生成

```python
def create_answer_element(answer: str, analysis: str, children: list) -> str:
    """
    生成格式：【答案】<br />{answer}<br />【解析】<br />{analysis}

    子题处理：
    - 每个子题编号 (1) (2) ...
    - 子子题用带圈数字 ①②... (超过20个用(21)(22)...)
    - 无子题时 answer/analysis 前加 &nbsp;
    """
```

---
