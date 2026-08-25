# 大纲解析配置表与讲题模板配置表

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 5.1 设计原则

用配置表替代硬编码outlineCode，新增大纲类型只需追加配置项，不改源码。

## 5.2 配置结构

```python
from dataclasses import dataclass
from typing import Optional, List

@dataclass(frozen=True)
class OutlineNode:
    """大纲节点配置"""
    code: int                              # outlineCode，None表示遍历所有未匹配子节点
    name: str                              # 节点名称（日志/调试用）
    page_type: str                         # 产物页类型: title/formula/ml/question/answer
    children: Optional[List['OutlineNode']] = None  # 子节点配置
    content_key: str = "content"           # 内容取值路径，默认直接取content字段
    is_formula_content: bool = False       # True时取content[0]["content"]再分页

# ── 教学规划子节点 ──
TEACHING_PLAN_CHILDREN = [
    OutlineNode(code=13, name="教学目标", page_type="formula", is_formula_content=True),
    OutlineNode(code=14, name="重点难点", page_type="formula", is_formula_content=True),
    OutlineNode(code=15, name="教学方法", page_type="formula", is_formula_content=True),
]

# ── 考点子节点 ──
EXAM_POINT_CHILDREN = [
    OutlineNode(code=20, name="解法探究", page_type="formula", is_formula_content=True),
    # code=None 的节点：遍历剩余子节点作为例题处理
    OutlineNode(code=None, name="例题", page_type="question"),
]

# ── 知识点配置 ──
KNOWLEDGE_POINT_CONFIG = [
    OutlineNode(code=18, name="知识梳理", page_type="formula", is_formula_content=True),
    OutlineNode(code=None, name="考点", page_type="formula", children=EXAM_POINT_CHILDREN),
]

# ── 例题子节点（练习/作业） ──
EXERCISE_CHILDREN = [
    OutlineNode(code=22, name="课内练习", page_type="question"),
    OutlineNode(code=23, name="课后作业", page_type="question"),
]

# ── 教学过程子节点 ──
TEACHING_PROCESS_CHILDREN = [
    OutlineNode(code=16, name="新课导入", page_type="formula", is_formula_content=True),
    # 其余为知识点，动态遍历
]

# ── 完整大纲树 ──
LESSON_OUTLINE_MAP = OutlineNode(
    code=6, name="教学设计", page_type="formula",
    children=[
        OutlineNode(code=7, name="教学规划", page_type="formula",
                    children=TEACHING_PLAN_CHILDREN),
        OutlineNode(code=8, name="教学过程", page_type="formula",
                    children=TEACHING_PROCESS_CHILDREN),
        OutlineNode(code=9, name="教学总结", page_type="formula", is_formula_content=True),
    ]
)
```

## 5.3 解析器行为规范

1. 遍历大纲树时，按配置表匹配 outlineCode
2. 匹配到节点后，根据 page_type 生成对应 Page 对象
3. `code=None` 的节点表示"遍历父节点的所有未匹配子节点"
4. **`next()` 查找失败必须返回 `None`，绝不返回1或任何非None默认值**
5. 找不到节点时记录 warning 日志，跳过该节点，不中断整体流程
6. `is_formula_content=True` 时，取值路径为 `item["content"][0]["content"]`，再经 split_page 分页
7. 知识点遍历时：每个知识点先产生目录页(ml)，再按 KNOWLEDGE_POINT_CONFIG 处理子节点
8. 例题处理：产生题目页(question) + 答案页(answer)，然后处理练习/作业子节点

---

## 6. 讲题模板配置表 (config/topic_template_map.py)

## 6.1 设计原则

3种讲题模板定义为数据结构，用配置驱动替代 if-elif-else。新增模板类型只需追加配置项。

## 6.2 配置结构

```python
from dataclasses import dataclass
from typing import List

@dataclass(frozen=True)
class TopicSection:
    """讲题模板的一个版块"""
    field_name: str        # 从题目JSON中取值的key
    title: str             # PPT中显示的标题

@dataclass(frozen=True)
class TopicTemplate:
    """讲题模板定义"""
    template_type: int               # topicTemplateType值
    name: str                        # 模板名称
    sections: List[TopicSection]     # 按顺序渲染的版块列表

TOPIC_TEMPLATES = {
    1: TopicTemplate(
        template_type=1,
        name="基础模板",
        sections=[
            TopicSection(field_name="testPointAnalysis", title="考点分析"),
            TopicSection(field_name="explorationOfSolutions", title="解法探究"),
            TopicSection(field_name="explorationOfSublimate", title="解法升华"),
        ]
    ),
    2: TopicTemplate(
        template_type=2,
        name="卡点模板",
        sections=[
            TopicSection(field_name="overallDesignAssessment", title="整体设计评估"),
            TopicSection(field_name="stuckPointAndSolution", title="核心卡点与破题路径"),
            TopicSection(field_name="similarQuestionDesignPattern", title="同类题模式识别"),
        ]
    ),
    3: TopicTemplate(
        template_type=3,
        name="全流程模板",
        sections=[
            TopicSection(field_name="testPointAnalysis", title="考点分析"),
            TopicSection(field_name="solutionToTheProblem", title="破题思路"),
            TopicSection(field_name="solutionProcess", title="解题过程"),
            TopicSection(field_name="cautionaryNote", title="易错警示"),
            TopicSection(field_name="explorationOfSublimate", title="解法升华"),
        ]
    ),
}
```

## 6.3 解析器行为规范

1. 从JSON中取 `topicTemplateType`，在 `TOPIC_TEMPLATES` 中查找模板配置
2. 未匹配到模板时记录 warning 日志，跳过模板版块（题目+答案仍正常生成）
3. 遍历 `sections`，按 `field_name` 从JSON取值，经 `md_to_text()` 清洗后分页
4. 每个section生成 `Page(type=PageType.QUESTION, data=PageData(title=section.title, content=page))`

---
