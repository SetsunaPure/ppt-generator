# PPT构建主流程、业务服务层与API入口

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 13.1 构建器

```python
from pathlib import Path
from pptx.util import Pt

class PptBuilder:
    def __init__(self, env: EnvConfig, pages: list[Page], font_size: int,
                 file_content_style: str, logo_path: str, template_prefix: str):
        self.env = env
        self.pages = pages
        self.font_size = Pt(font_size)
        self.style = file_content_style
        self.logo_path = logo_path
        self.prefix = template_prefix  # "lesson" | "topic"

    def build(self) -> Path:
        """
        1. 确保目录存在：file/output/, file/template/, file/layout_template/
        2. 生成路径：
           - output: file/output/{prefix}_output_{timestamp}.pptx
           - template: file/template/{prefix}_template_{style}.pptx
           - layout: file/layout_template/{prefix}_template_layout_{timestamp}.pptx
        3. generate_layout_template() → 生成排版模板
        4. 加载排版模板 Presentation(layout_path)
        5. 遍历pages渲染内容
        6. 保存PPT到output路径
        7. 返回output路径
        """
```

## 13.2 内容填充逻辑

```python
def _fill_slide(self, slide, page: Page, slide_idx: int):
    """
    根据 page.type 获取 SLIDE_LAYOUTS 配置，按映射填充：

    PageType.TITLE:
        - 从SLIDE_LAYOUTS["title"]取映射
        - 填充 title → shapes[shape_index].text_frame
        - 填充 subtitle → shapes[shape_index].text_frame
        - 注意：标题页只填充一次（取pages中第一个title类型的数据）

    PageType.CATALOG:
        - 从SLIDE_LAYOUTS["catalog"]取映射
        - 填充 content → shapes[shape_index].text_frame

    其他(PageType.FORMULA/QUESTION/ANSWER):
        - 从SLIDE_LAYOUTS["content"]取映射
        - 填充 header → shapes[shape_index].text_frame
        - auto_create_textbox=True时：
          add_textbox() 创建正文框
          调用 render_page(slide, page.data.content, env, font_size_pt, formula_cache, text_box_origin)
          其中 text_box_origin = (textbox.left, textbox.top) 文本框在slide上的坐标
          render_page 内部会把行内元素写入text_frame，块级元素(图片/表格)用绝对坐标直接放在slide上

    文本填充时保留模板原有字体/颜色/大小格式（copy_text逻辑）
    """
```

## 13.3 文本填充（保留模板格式）

```python
def copy_text(text_frame, title: str, font_size: Pt):
    """
    1. 深拷贝模板段落的格式
    2. 清空text_frame
    3. 用模板第一个段落的run格式创建新run
    4. 设置文本内容，继承模板的字体名/大小/颜色/粗体/斜体/下划线
    """
```

---

## 14. 业务服务层

## 14.1 备课PPT服务 (service/lesson_service.py)

```python
def create_lesson_ppt(request: LessonPptRequest) -> Optional[str]:
    """
    返回OSS文件链接，异常返回None

    流程：
    1. env = get_env_config(request.activeProfile)
    2. json_data = json.loads(request.lessonDetail)
    3. pages = OutlineParser(LESSON_OUTLINE_MAP).parse(json_data)
       - 第一项固定为标题页：title="课时名称："+json_data["title"], subTitle="单元名称："+json_data["subTitle"]
    4. ppt_path = PptBuilder(env, pages, request.fontSize,
                             request.fileContentStyle, request.schoolLogo,
                             "lesson").build()
    5. oss_region = env.oss_region
       oss_prefix = f"lesson/download/pptx/{json_data['subTitle']}_{timestamp}_{request.fileContentStyle}.pptx"
    6. oss_url = OssOperationHandler(oss_region).upload_file(ppt_path, oss_prefix)
    7. 清理临时文件(ppt_path, layout_template_path)
    8. 返回oss_url

    异常处理：
    - logger.error() 记录完整异常
    - send_wechat_notify(env, message=f"lessonId:{request.lessonId} 备课转换异常：{e}")
    - 返回None
    """
```

## 14.2 讲题PPT服务 (service/topic_service.py)

```python
def create_topic_ppt(request: TopicPptRequest) -> Optional[str]:
    """
    返回OSS文件链接，异常返回None

    流程：
    1. env = get_env_config(request.activeProfile)
    2. json_data = json.loads(request.detail)
    3. pages = TopicParser(config=TOPIC_TEMPLATES).parse(json_data)
       - 先生成题目页+答案页
       - 再按topicTemplateType匹配模板生成各版块页
    4. ppt_path = PptBuilder(env, pages, request.fontSize,
                             request.fileContentStyle, "",
                             "topic").build()
    5. oss_prefix = f"topic/download/pptx/{request.questionId}_{timestamp}_{request.fileContentStyle}.pptx"
    6. oss_url = OssOperationHandler(env.oss_region).upload_file(ppt_path, oss_prefix)
    7. 清理临时文件
    8. 返回oss_url

    异常处理：同上
    """
```

---

## 15. FastAPI入口 (main.py)

```python
from fastapi import FastAPI, Body
from models.request import LessonPptRequest, TopicPptRequest
from service.lesson_service import create_lesson_ppt
from service.topic_service import create_topic_ppt

app = FastAPI(root_path="/api")

@app.post("/lesson/create_ppt")
def create_ppt(request: LessonPptRequest = Body(...)):
    """备课PPT生成接口"""
    return create_lesson_ppt(request)

@app.post("/exam/topic/create_ppt")
def exam_topic_create_ppt(request: TopicPptRequest = Body(...)):
    """讲题PPT生成接口"""
    return create_topic_ppt(request)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
```

---
