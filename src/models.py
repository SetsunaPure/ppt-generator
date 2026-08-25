"""
PPT生成器 - 数据模型模块

包含:
- 环境配置 (EnvConfig)
- Block数据模型
- AnswerGroupBlock数据模型
- Page相关数据模型
- 请求参数模型
"""

from dataclasses import dataclass, field
from typing import Optional, Any, List
from enum import Enum

from .config import PageType


# ============================================================
# 环境配置
# ============================================================

@dataclass(frozen=True)
class EnvConfig:
    """环境配置数据类"""
    profile: str                              # 环境标识: dev/test/product
    proxy: Optional[str]                      # 代理地址，dev/test为None
    oss_region: str                           # "beijing" | "hangzhou"
    latex_api_url: str                        # LaTeX转换接口地址
    wechat_webhook: str                        # 企微机器人webhook
    wechat_mentioned: tuple[str, ...]          # @手机号列表

    @property
    def proxies(self) -> Optional[dict]:
        """获取requests代理配置"""
        if self.proxy:
            return {"http": self.proxy, "https": self.proxy}
        return None


# 环境配置预设
ENV_PRESETS: dict[str, EnvConfig] = {
    "dev": EnvConfig(
        profile="dev",
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://exam.canpoint.cn/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=("15392878433",),
    ),
    "test": EnvConfig(
        profile="test",
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://exam.canpoint.cn/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=("15392878433",),
    ),
    "product": EnvConfig(
        profile="product",
        proxy="http://10.1.1.219:8500",
        oss_region="hangzhou",
        latex_api_url="http://10.1.1.231/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=("15392878433",),
    ),
}


def get_env_config(profile: str) -> EnvConfig:
    """根据环境标识获取环境配置"""
    if profile not in ENV_PRESETS:
        raise ValueError(f"未知环境标识: {profile}，可选: {list(ENV_PRESETS.keys())}")
    return ENV_PRESETS[profile]


# ============================================================
# Block位置信息
# ============================================================

@dataclass
class BlockPosition:
    """Block在页面上的位置"""
    left: int = 0       # 左侧偏移(EMU)
    top: int = 0        # 顶部偏移(EMU)


# ============================================================
# Block数据模型 - 积木流式布局的基本单元
# ============================================================

@dataclass
class Block:
    """
    内容块 —— 积木流式布局的基本单元
    每个Block都有确定的宽高，由布局引擎按尺寸排列
    """
    type: str                   # "text" | "latex" | "image" | "table" | "newline" | "answer_group"
    content: Any = None         # 具体内容数据

    # ── 尺寸（EMU单位） ──
    width: int = 0              # 块宽度
    height: int = 0             # 块高度

    # ── 位置（由布局引擎计算） ──
    position: BlockPosition = field(default_factory=BlockPosition)

    # ── 行内/块级标记 ──
    inline: bool = True         # True=行内元素(可和别的行内元素共处一行)，False=独占一行的块级元素

    # ── 图片专用 ──
    image_url: Optional[str] = None
    image_width_inches: Optional[float] = None
    image_height_inches: Optional[float] = None
    image_size_level: Optional[str] = None     # "normal" | "large" | "fullpage"

    # ── 表格专用 ──
    table_data: Optional[list] = None
    table_rows: int = 0
    table_cols: int = 0

    # ── 文本专用 ──
    is_bold: bool = False
    font_name: str = "黑体"

    # ── 答案组专用 ──
    answer_group: Optional['AnswerGroupBlock'] = None   # 非None时，本Block是答案组的容器


# ============================================================
# AnswerGroupBlock - 选择题答案组，保证只出现双排或四排
# ============================================================

@dataclass
class AnswerGroupBlock:
    """
    选择题答案组 —— 保证答案只出现双排或四排，绝不单排
    答案项成对排列：一行放得下4个就四排，放不下4个就拆成2+2双排
    """
    items: List[Block]          # 每个答案项的inline Block
    columns: int                 # 2 或 4，由宽度计算决定
    row_height: int              # 单行高度(text_line_height)
    total_height: int            # 总高度 = 行数 × row_height
    total_width: int             # = TEXT_BOX_WIDTH（撑满文本框宽度）


# ============================================================
# Page数据模型
# ============================================================

@dataclass
class PageData:
    """Page的内容数据"""
    title: str
    content: Optional[Any] = None       # 分页后的Block列表 或 None
    subTitle: Optional[str] = None      # 仅title类型使用
    text_title: Optional[str] = None    # 仅formula类型，副标题


@dataclass
class Page:
    """一个PPT页面"""
    type: PageType
    data: PageData


# ============================================================
# 题目数据模型
# ============================================================

@dataclass
class QuestionData:
    """题目数据模型"""
    stem: str                           # 题干HTML
    options: Optional[List[str]]        # 选项列表
    answer: Optional[str]               # 答案文本
    analysis: Optional[str]              # 解析文本
    children: List['QuestionData'] = field(default_factory=list)  # 子题列表
    ques_type: Optional[str] = None      # 题型名称


# ============================================================
# 请求参数模型（Pydantic模型，供API使用）
# ============================================================

class LessonPptRequest:
    """备课PPT请求"""
    def __init__(
        self,
        lessonId: str = "123456",
        stage: int = 12,
        subject: int = 9912,
        lessonDetail: str = "{}",
        fontSize: int = 16,
        activeProfile: str = "dev",
        fileContentStyle: str = "0",
        schoolLogo: str = "",
    ):
        self.lessonId = lessonId
        self.stage = stage
        self.subject = subject
        self.lessonDetail = lessonDetail
        self.fontSize = fontSize
        self.activeProfile = activeProfile
        self.fileContentStyle = fileContentStyle
        self.schoolLogo = schoolLogo


class TopicPptRequest:
    """讲题PPT请求"""
    def __init__(
        self,
        questionId: str = "123456",
        stage: int = 12,
        subject: int = 9912,
        detail: str = "{}",
        fontSize: int = 16,
        activeProfile: str = "dev",
        fileContentStyle: str = "0",
    ):
        self.questionId = questionId
        self.stage = stage
        self.subject = subject
        self.detail = detail
        self.fontSize = fontSize
        self.activeProfile = activeProfile
        self.fileContentStyle = fileContentStyle
