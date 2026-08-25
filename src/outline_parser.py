"""
PPT生成器 - 大纲解析与讲题模板解析

严格对齐老代码 lesson_split_json 的结构逻辑：
outlineCode=6  教学设计
  └ outlineCode=7  教学规划
      ├ outlineCode=13 教学目标 → formula页
      ├ outlineCode=14 重点难点 → formula页
      └ outlineCode=15 教学方法 → formula页
  └ outlineCode=8  教学过程
      ├ outlineCode=16 新课导入 → formula页
      ├ 知识点节点 → ml目录页
      │   ├ outlineCode=18 知识梳理 → formula页
      │   └ 考点节点
      │       ├ outlineCode=20 解法探究 → formula页
      │       └ 例题节点 → question页 + answer页
      │           └ outlineCode=22 课内练习 → question页
      │           └ outlineCode=23 课后作业 → question页
  └ outlineCode=9  教学总结 → formula页
"""

import json
import logging
from typing import Optional, List, Any, Dict, Tuple
from dataclasses import dataclass, field

from .config import PageType
from .models import Page, PageData, QuestionData
from .layout_engine import split_page
from .compat import md_to_text

logger = logging.getLogger("ppt_generator")


# ============================================================
# 大纲解析器
# ============================================================

class OutlineParser:
    """
    课时大纲JSON解析器

    严格对齐老代码 lesson_split_json 的遍历逻辑：
    1. 标题页（固定第一页）
    2. 目录页（收集教学规划子节点名称）
    3. 教学规划（教学目标/重点难点/教学方法）
    4. 教学过程（新课导入 → 知识点遍历）
    5. 教学总结
    """

    def parse(self, json_data: Dict[str, Any], font_size: int = 16) -> List[Page]:
        """
        解析大纲JSON，生成Page列表

        严格对齐老代码 lesson_split_json 的遍历逻辑：
        1. 标题页（固定第一页）
        2. 教学目标/重点难点/教学方法（直接从教学规划取，无目录页）
        3. 教学过程（新课导入 → 知识点遍历，每个知识点先ml目录页）
        4. 教学总结
        """
        pages: List[Page] = []

        title = json_data.get("title", "")
        sub_title = json_data.get("subTitle", "")

        # ── 1. 标题页 ──
        pages.append(Page(
            type=PageType.TITLE,
            data=PageData(
                title=f"课时名称：{title}",
                subTitle=f"单元名称：{sub_title}" if sub_title else None
            )
        ))

        # 定位教学设计节点（outlineCode=6）
        outline_list = json_data.get("outlineList", [])
        jxsj = self._find_node(outline_list, 6)
        if not jxsj:
            logger.warning("未找到教学设计节点(outlineCode=6)")
            return pages

        # ── 2. 教学规划（outlineCode=7）→ 教学目标/重点难点/教学方法 ──
        # 对齐老代码：直接append_formula_item，没有目录页
        jxgh = self._find_node(jxsj.get("children", []), 7)
        if jxgh:
            self._append_formula_item(pages, jxgh, 13, font_size=font_size)
            self._append_formula_item(pages, jxgh, 14, font_size=font_size)
            self._append_formula_item(pages, jxgh, 15, font_size=font_size)

        # ── 3. 教学过程（outlineCode=8） ──
        jxgc = self._find_node(jxsj.get("children", []), 8)
        if jxgc:
            lt_index = 0  # 例题计数器
            for zsd in jxgc.get("children", []):
                code = zsd.get("outlineCode")

                # 新课导入（outlineCode=16）
                # 对齐老代码：append_formula_content(list, xkdr_name, zsd["content"])
                if code == 16:
                    xkdr_name = zsd.get("name", "")
                    self._append_formula_content(pages, xkdr_name, zsd.get("content", []), font_size)
                    continue

                # 知识点节点 → ml目录页 + 子节点遍历
                # 对齐老代码：list.append({"type": "ml", "data": {"content": zsd["name"]}})
                zsd_name = zsd.get("name", "")
                pages.append(Page(
                    type=PageType.CATALOG,
                    data=PageData(title=None, content=None)
                ))
                # 老代码把zsd_name写到目录页的content字段（shapes[2]）
                # 这里暂存在PageData.title中，builder会填充到catalog slide的content shape
                if pages:
                    pages[-1].data.title = zsd_name

                # 遍历知识点子节点
                for kx in zsd.get("children", []):
                    kx_code = kx.get("outlineCode")

                    # 知识梳理（outlineCode=18）
                    # 对齐老代码：append_formula_content(list, zssl_name, kx["content"])
                    if kx_code == 18:
                        zssl_name = kx.get("name", "")
                        self._append_formula_content(pages, zssl_name, kx.get("content", []), font_size)
                        continue

                    # 考点节点
                    kx_name = kx.get("name", "")
                    kx_children = kx.get("children", [])

                    for kx_child in kx_children:
                        kx_child_code = kx_child.get("outlineCode")

                        # 解法探究（outlineCode=20）
                        # 对齐老代码：append_formula_item(list, kx, 20, title=kx_name, text_title=kx_child["name"])
                        if kx_child_code == 20:
                            self._append_formula_item(
                                pages, kx, 20,
                                title=kx_name,
                                font_size=font_size
                            )
                            continue

                        # 例题节点（有questionVo的content）
                        lt_name = kx_child.get("name", "")
                        content_list = kx_child.get("content", [])

                        if content_list:
                            qvo = content_list[0].get("questionVo") if content_list else None
                            if qvo:
                                self._append_question_and_answer(pages, qvo, kx_name, lt_index + 1, font_size)
                                lt_index += 1

                                # 例题的子节点：课内练习(22)、课后作业(23)
                                for lt_child in kx_child.get("children", []):
                                    lt_child_code = lt_child.get("outlineCode")
                                    if lt_child_code in (22, 23):
                                        self._append_exercise_items(pages, lt_child, kx_name, font_size)

        # ── 4. 教学总结（outlineCode=9） ──
        self._append_formula_item(pages, jxsj, 9, font_size=font_size)

        return pages

    # ============================================================
    # 辅助方法
    # ============================================================

    @staticmethod
    def _find_node(children: List[Dict], code: int) -> Optional[Dict]:
        """在children列表中按outlineCode查找节点"""
        return next((i for i in children if i.get("outlineCode") == code), None)

    def _append_formula_content(self, pages: List[Page], item_name: str,
                                 content_data: list, font_size: int):
        """
        从item_name和content_data生成formula页
        对齐老代码 append_formula_content(list, item_name, content_data)
        取content_data[0]["content"]再分页
        """
        if content_data and len(content_data) > 0 and content_data[0].get("content"):
            content = content_data[0]["content"]
            blocks_pages = split_page(content, font_size)
            for i, block_page in enumerate(blocks_pages):
                pages.append(Page(
                    type=PageType.FORMULA,
                    data=PageData(
                        title=item_name,
                        content=block_page
                    )
                ))
        else:
            pages.append(Page(
                type=PageType.FORMULA,
                data=PageData(title=item_name, content=None)
            ))

    def _append_formula_item(self, pages: List[Page], context: Dict, outline_code: int,
                              title: Optional[str] = None, font_size: int = 16):
        """
        从父节点的children中按outlineCode查找子节点，生成formula页
        对齐老代码 append_formula_item(list, context, outline_conf_code, title, text_title)
        """
        children = context.get("children", [])
        item = self._find_node(children, outline_code)

        if not item:
            # 节点不存在，生成空页
            final_title = title if title else ""
            pages.append(Page(
                type=PageType.FORMULA,
                data=PageData(title=final_title, content=None)
            ))
            return

        final_title = title if title is not None else item.get("name", "")
        content_list = item.get("content", [])

        if content_list and len(content_list) > 0 and "content" in content_list[0]:
            content = content_list[0]["content"]
            if content:
                blocks_pages = split_page(content, font_size)
                for i, block_page in enumerate(blocks_pages):
                    pages.append(Page(
                        type=PageType.FORMULA,
                        data=PageData(
                            title=final_title,
                            content=block_page
                        )
                    ))
                return

        pages.append(Page(
            type=PageType.FORMULA,
            data=PageData(title=final_title, content=None)
        ))

    def _append_question_and_answer(self, pages: List[Page], question_vo: Dict,
                                     kx_name: str, question_index: int, font_size: int):
        """
        从questionVo生成题目页+答案页
        对齐老代码中例题处理逻辑
        """
        # 解析题目
        stem, options, answer, analysis, children = self._parse_question(question_vo)

        # 题目页
        question_content = f"【例题{question_index}】{self._create_question_element(stem, options, children)}"
        if question_content:
            blocks_pages = split_page(question_content, font_size)
            for i, block_page in enumerate(blocks_pages):
                pages.append(Page(
                    type=PageType.QUESTION,
                    data=PageData(
                        title=kx_name if i == 0 else "",
                        content=block_page
                    )
                ))

        # 答案页
        answer_content = self._create_answer_element(answer, analysis, children)
        if answer_content:
            blocks_pages = split_page(answer_content, font_size)
            for i, block_page in enumerate(blocks_pages):
                pages.append(Page(
                    type=PageType.ANSWER,
                    data=PageData(
                        title=kx_name if i == 0 else "",
                        content=block_page
                    )
                ))

    def _append_exercise_items(self, pages: List[Page], lt_child: Dict,
                                kx_name: str, font_size: int):
        """
        处理课内练习/课后作业
        对齐老代码 append_question_item
        """
        ex_name = lt_child.get("name", "")
        content_list = lt_child.get("content", [])
        index = 0

        for ex in content_list:
            qvo = ex.get("questionVo")
            if not qvo:
                continue

            stem, options, answer, analysis, children = self._parse_question(qvo)
            question_content = f"【{ex_name}{index + 1}】{self._create_question_element(stem, options, children)}"
            if question_content:
                blocks_pages = split_page(question_content, font_size)
                for block_page in blocks_pages:
                    pages.append(Page(
                        type=PageType.QUESTION,
                        data=PageData(title=kx_name, content=block_page)
                    ))

            answer_content = self._create_answer_element(answer, analysis, children)
            if answer_content:
                blocks_pages = split_page(answer_content, font_size)
                for block_page in blocks_pages:
                    pages.append(Page(
                        type=PageType.ANSWER,
                        data=PageData(title=kx_name, content=block_page)
                    ))

            index += 1

    # ============================================================
    # 题目解析方法（对齐老代码 parse_question）
    # ============================================================

    def _parse_question(self, lt: Dict, name: str = "", index: int = 0):
        """解析题目数据"""
        explain_list = lt.get("explain", [{}])
        explain = explain_list[0] if explain_list else {}
        answers = explain.get("answers", None)
        analyses = explain.get("analysis", None)
        ques_type_name = lt.get("quesType", {}).get("name", None) if lt.get("quesType") else None

        stem = lt.get("context", {}).get("stem", None)
        options = lt.get("context", {}).get("options", None)
        answer = self._phase_answer(answers, ques_type_name)
        analysis = self._phase_analysis(analyses)

        # 递归处理子题
        lt_list = lt.get("children", None)
        tmp_children = []
        if lt_list is not None:
            for child_lt in lt_list:
                s, o, a, an, c = self._parse_question(child_lt, name, index)
                tmp_children.append({
                    "stem": s, "options": o, "answer": a, "analysis": an, "children": c
                })

        return stem, options, answer, analysis, tmp_children

    @staticmethod
    def _phase_answer(src, ques_type=None):
        """对齐老代码 phase_answer"""
        if isinstance(src, str):
            try:
                answers = json.loads(src)
            except json.JSONDecodeError:
                return src
        else:
            answers = src

        if answers is None or answers == "":
            return ""

        if isinstance(answers, list) and len(answers) > 0:
            # 二维数组
            if all(isinstance(item, list) for item in answers):
                if ques_type == "判断题":
                    processed = [
                        ', '.join('对' if i == 1 or i == '1' else '错' for i in sublist)
                        for sublist in answers
                    ]
                else:
                    processed = [', '.join(str(x) for x in sublist) for sublist in answers]
                return ';\u0026nbsp;'.join(processed)
            # 一维数组
            else:
                return ', '.join(str(x) for x in answers)

        return str(answers)

    @staticmethod
    def _phase_analysis(src):
        """对齐老代码 phase_analysis"""
        if isinstance(src, str):
            try:
                analyses = json.loads(src)
            except json.JSONDecodeError:
                return src
        else:
            analyses = src

        if analyses is None or analyses == "":
            return ""

        if isinstance(analyses, list) and len(analyses) > 0:
            return '\n'.join(str(a) for a in analyses)

        return str(analyses)

    @staticmethod
    def _create_question_element(stem, options, children):
        """对齐老代码 create_question_element"""
        if not stem:
            return ""

        result = str(stem)

        # 选项（可能是list[dict]、list[str]或None）
        if options and isinstance(options, list):
            for i, opt in enumerate(options):
                if isinstance(opt, dict):
                    label = opt.get("label", "")
                    content = opt.get("content", "")
                    if label or content:
                        result += f"\n{label}. {content}" if label else f"\n{content}"
                elif isinstance(opt, str):
                    # 选项是纯文本/HTML字符串，直接附加
                    prefix = chr(65 + i) + "."  # A. B. C. D.
                    result += f"\n{prefix} {opt}"

        # 子题
        for i, child in enumerate(children or []):
            child_stem = child.get("stem", "")
            if child_stem:
                result += f"\n({i+1}) {child_stem}"
            child_options = child.get("options", [])
            if child_options and isinstance(child_options, list):
                for j, opt in enumerate(child_options):
                    if isinstance(opt, dict):
                        label = opt.get("label", "")
                        content = opt.get("content", "")
                        if label or content:
                            result += f"\n  {label}. {content}" if label else f"\n  {content}"
                    elif isinstance(opt, str):
                        prefix = chr(65 + j) + "."
                        result += f"\n  {prefix} {opt}"

        return result

    @staticmethod
    def _create_answer_element(answer, analysis, children):
        """对齐老代码 create_answer_element"""
        answer = answer if answer else ""
        analysis = analysis if analysis else ""

        # 子题答案和解析
        for i, child in enumerate(children or []):
            child_answer = child.get("answer", "")
            child_analysis = child.get("analysis", "")
            if child_answer:
                answer += f'\u0026nbsp;({i+1}){child_answer}<br />'
            if child_analysis:
                analysis += f'\u0026nbsp;({i+1}){child_analysis}<br />'

            # 孙题
            for si, sub_child in enumerate(child.get("children", [])):
                symbol = f'({si+1})' if si >= 20 else chr(0x2460 + si)
                sub_answer = sub_child.get("answer", "")
                sub_analysis = sub_child.get("analysis", "")
                if sub_answer:
                    answer += f'\u0026nbsp;{symbol}{sub_answer}<br />'
                if sub_analysis:
                    analysis += f'\u0026nbsp;{symbol}{sub_analysis}<br />'

        if not children:
            answer = f'\u0026nbsp;{answer}'
            analysis = f'\u0026nbsp;{analysis}'

        return '【答案】<br />' + answer + '<br />【解析】<br />' + analysis


# 保留向后兼容的常量引用
LESSON_OUTLINE_MAP = None  # 不再使用配置驱动，改为硬编码逻辑


# ============================================================
# 讲题模板配置
# ============================================================

@dataclass(frozen=True)
class TopicSection:
    """讲题模板的一个版块"""
    field_name: str        # 从题目JSON中取值的key
    title: str             # PPT中显示的标题


@dataclass(frozen=True)
class TopicTemplate:
    """讲题模板定义"""
    template_type: int               # topicTemplateType值
    name: str                         # 模板名称
    sections: Tuple[TopicSection, ...]     # 按顺序渲染的版块列表


TOPIC_TEMPLATES: Dict[int, TopicTemplate] = {
    1: TopicTemplate(
        template_type=1,
        name="基础模板",
        sections=(
            TopicSection(field_name="testPointAnalysis", title="考点分析"),
            TopicSection(field_name="explorationOfSolutions", title="解法探究"),
            TopicSection(field_name="explorationOfSublimate", title="解法升华"),
        )
    ),
    2: TopicTemplate(
        template_type=2,
        name="卡点模板",
        sections=(
            TopicSection(field_name="overallDesignAssessment", title="整体设计评估"),
            TopicSection(field_name="stuckPointAndSolution", title="核心卡点与破题路径"),
            TopicSection(field_name="similarQuestionDesignPattern", title="同类题模式识别"),
        )
    ),
    3: TopicTemplate(
        template_type=3,
        name="全流程模板",
        sections=(
            TopicSection(field_name="testPointAnalysis", title="考点分析"),
            TopicSection(field_name="solutionToTheProblem", title="破题思路"),
            TopicSection(field_name="solutionProcess", title="解题过程"),
            TopicSection(field_name="cautionaryNote", title="易错警示"),
            TopicSection(field_name="explorationOfSublimate", title="解法升华"),
        )
    ),
}


# ============================================================
# 讲题模板解析器
# ============================================================

class TopicParser:
    """
    讲题模板JSON解析器

    根据TOPIC_TEMPLATES配置遍历题目JSON，生成Page列表
    """

    def __init__(self, config: Dict[int, TopicTemplate] = TOPIC_TEMPLATES):
        self.config = config

    def parse(
        self,
        json_data: Dict[str, Any],
        font_size: int = 16,
        is_origin: bool = True
    ) -> List[Page]:
        """
        解析讲题JSON

        参数:
            json_data: 题目详情JSON
            font_size: 字号
            is_origin: 是否为备课场景（True）还是讲题场景（False）

        返回:
            Page列表
        """
        pages: List[Page] = []

        # 解析题目数据
        question = self._parse_question_vo(json_data, is_origin)

        # 生成题目页
        if question.stem:
            stem_blocks = split_page(question.stem, font_size)
            for i, block_page in enumerate(stem_blocks):
                pages.append(Page(
                    type=PageType.QUESTION,
                    data=PageData(
                        title="题目" if i == 0 else "",
                        content=block_page
                    )
                ))

        # 生成答案页
        if question.answer:
            pages.append(Page(
                type=PageType.ANSWER,
                data=PageData(title="答案", content=None)
            ))

        # 按模板生成各版块页
        template_type = json_data.get("topicTemplateType", 1)
        template = self.config.get(template_type)

        if template:
            logger.info(f"使用讲题模板: {template.name} (type={template_type})")
            for section in template.sections:
                content = json_data.get(section.field_name, "")
                if content:
                    # md转纯文本再分页
                    plain_text = md_to_text(content)
                    section_blocks = split_page(plain_text, font_size)
                    for block_page in section_blocks:
                        pages.append(Page(
                            type=PageType.QUESTION,
                            data=PageData(title=section.title, content=block_page)
                        ))
        else:
            logger.warning(f"未找到匹配的讲题模板: topicTemplateType={template_type}")

        # 递归处理子题
        for child in question.children:
            child_json = child  # 简化处理
            child_pages = self.parse(child_json, font_size, is_origin)
            pages.extend(child_pages)

        return pages

    def _parse_question_vo(
        self,
        question_vo: Dict[str, Any],
        is_origin: bool = True
    ) -> QuestionData:
        """从questionVo中提取题目数据"""
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

        # 处理答案格式化
        answer_str = self._format_answers(answers, ques_type)

        # 解析子题
        children = []
        for child_vo in question_vo.get("children", []):
            children.append(self._parse_question_vo(child_vo, is_origin))

        return QuestionData(
            stem=stem,
            options=options,
            answer=answer_str,
            analysis=analysis,
            children=children,
            ques_type=ques_type
        )

    def _format_answers(self, answers: Any, ques_type: Optional[str]) -> Optional[str]:
        """
        格式化答案

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
