"""
PPT生成器 - 备课PPT生成系统

基于积木流式布局引擎的PPT生成系统

主要模块：
- config: 配置（代理环境、shapes映射、布局常量）
- models: 数据模型（Block、AnswerGroupBlock、Page、EnvConfig）
- compat: 兼容性处理（MathML修正、OMML适配、HTML清洗）
- utils: 工具函数（HTTP请求、OSS、日志、图片处理）
- layout_engine: 积木流式布局引擎（核心）
- outline_parser: 大纲解析与讲题模板解析
- question_parser: 题目解析与答案排版
- formula_converter: 公式批量转换
- renderer: PPT渲染引擎
- builder: PPT构建主流程
- service: 业务服务层

使用示例：
    from src.service import create_lesson_ppt, create_topic_ppt
    from src.models import LessonPptRequest, TopicPptRequest

    # 备课PPT
    request = LessonPptRequest(
        lessonId="123456",
        lessonDetail='{"title": "测试", "outlineList": []}',
        activeProfile="dev"
    )
    url = create_lesson_ppt(request)

    # 讲题PPT
    request = TopicPptRequest(
        questionId="789",
        detail='{"topicTemplateType": 1, ...}',
        activeProfile="dev"
    )
    url = create_topic_ppt(request)
"""

__version__ = "1.2.0"
__author__ = "PPT Generator Team"

# 导入主要组件
from .models import (
    EnvConfig,
    get_env_config,
    ENV_PRESETS,
)
from .config import (
    get_proxy_config,
    TEXT_BOX_WIDTH,
    TEXT_BOX_HEIGHT,
    AVAILABLE_WIDTH,
    LINE_SPACING_FACTOR,
    calc_text_line_height,
    char_width_factor,
    ImageSizeLevel,
    PageType,
    SLIDE_LAYOUTS,
    ShapeMapping,
    SlideLayout,
    TEXT_BOX_CONFIG,
    FilePaths,
)

from .models import (
    Block,
    BlockPosition,
    AnswerGroupBlock,
    Page,
    PageData,
    QuestionData,
    LessonPptRequest,
    TopicPptRequest,
)

from .layout_engine import (
    layout_blocks,
    html_to_blocks,
    split_page,
    split_table_if_needed,
    calc_text_width,
)

from .outline_parser import (
    OutlineParser,
    TopicParser,
    TOPIC_TEMPLATES,
    TopicSection,
    TopicTemplate,
)

from .question_parser import (
    parse_question,
    format_answers,
    build_answer_group,
    create_answer_element,
    assemble_question_content,
    get_options_for_layout,
)

from .formula_converter import (
    batch_latex_to_omml,
    batch_convert_formulas,
    convert_single_formula,
    FormulaCache,
)

from .renderer import (
    render_page,
    render_content_page,
)

from .builder import (
    PptBuilder,
    build_ppt,
    generate_layout_template,
)

from .service import (
    create_lesson_ppt,
    create_topic_ppt,
    health_check,
)

from .compat import (
    revise_mathml,
    adapt_omml,
    clean_html_for_pptx,
    md_to_text,
)

from .utils import (
    get_logger,
    make_request,
    send_wechat_notify,
    get_image_dimensions,
    download_image_to_bytes,
    save_temp_image,
    get_scaled_dimensions,
    OssOperationHandler,
    convert_to_internal_url,
    ensure_dir,
    cleanup_temp_files,
    get_timestamp_filename,
)

__all__ = [
    # 配置
    "EnvConfig",
    "get_env_config",
    "ENV_PRESETS",
    "PROXY_SETTINGS",
    "TEXT_BOX_WIDTH",
    "TEXT_BOX_HEIGHT",
    "AVAILABLE_WIDTH",
    "LINE_SPACING_FACTOR",
    "calc_text_line_height",
    "char_width_factor",
    "ImageSizeLevel",
    "PageType",
    "SLIDE_LAYOUTS",
    "ShapeMapping",
    "SlideLayout",
    "TEXT_BOX_CONFIG",
    "FilePaths",
    # 模型
    "Block",
    "BlockPosition",
    "AnswerGroupBlock",
    "Page",
    "PageData",
    "QuestionData",
    "LessonPptRequest",
    "TopicPptRequest",
    # 布局引擎
    "layout_blocks",
    "html_to_blocks",
    "split_page",
    "split_table_if_needed",
    "calc_text_width",
    # 大纲解析
    "OutlineParser",
    "TopicParser",
    "TOPIC_TEMPLATES",
    "TopicSection",
    "TopicTemplate",
    # 题目解析
    "parse_question",
    "format_answers",
    "build_answer_group",
    "create_answer_element",
    "assemble_question_content",
    "get_options_for_layout",
    # 公式转换
    "batch_latex_to_omml",
    "batch_convert_formulas",
    "convert_single_formula",
    "FormulaCache",
    # 渲染
    "render_page",
    "render_content_page",
    # 构建
    "PptBuilder",
    "build_ppt",
    "generate_layout_template",
    # 服务
    "create_lesson_ppt",
    "create_topic_ppt",
    "health_check",
    # 兼容
    "revise_mathml",
    "adapt_omml",
    "clean_html_for_pptx",
    "md_to_text",
    # 工具
    "get_logger",
    "make_request",
    "send_wechat_notify",
    "get_image_dimensions",
    "download_image_to_bytes",
    "save_temp_image",
    "get_scaled_dimensions",
    "OssOperationHandler",
    "convert_to_internal_url",
    "ensure_dir",
    "cleanup_temp_files",
    "get_timestamp_filename",
]
