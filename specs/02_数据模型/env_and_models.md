# 环境配置、请求模型与Pages模型

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 2.1 配置结构

```python
from dataclasses import dataclass
from typing import Optional

@dataclass(frozen=True)
class EnvConfig:
    proxy: Optional[str]       # 代理地址，dev/test为None
    oss_region: str            # "beijing" | "hangzhou"
    latex_api_url: str         # LaTeX转换接口地址
    wechat_webhook: str        # 企微机器人webhook
    wechat_mentioned: list     # @手机号列表

ENV_PRESETS = {
    "dev": EnvConfig(
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://exam.canpoint.cn/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=["15392878433"],
    ),
    "test": EnvConfig(
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://exam.canpoint.cn/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=["15392878433"],
    ),
    "product": EnvConfig(
        proxy="http://10.1.1.219:8500",
        oss_region="hangzhou",
        latex_api_url="http://10.1.1.231/python_api_latex/latex/latex2mml",
        wechat_webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4857f76c-c273-4cd3-ad0a-36b3e8e2af87",
        wechat_mentioned=["15392878433"],
    ),
}

def get_env_config(profile: str) -> EnvConfig:
    """根据activeProfile获取环境配置，未知profile抛ValueError"""
    if profile not in ENV_PRESETS:
        raise ValueError(f"未知环境标识: {profile}，可选: {list(ENV_PRESETS.keys())}")
    return ENV_PRESETS[profile]
```

## 2.2 代理请求工具函数

```python
import requests

def make_request(method: str, url: str, env: EnvConfig, **kwargs):
    """统一请求函数，自动处理代理。dev/test直连，product走代理"""
    proxies = {"http": env.proxy, "https": env.proxy} if env.proxy else None
    return requests.request(method, url, proxies=proxies, **kwargs)
```

---

## 3. 请求参数模型 (models/request.py)

## 3.1 备课PPT请求

```python
from pydantic import BaseModel, Field

class LessonPptRequest(BaseModel):
    lessonId: str = Field(default="123456", description="课时ID")
    stage: int = Field(default=12, description="学段")
    subject: int = Field(default=9912, description="学科")
    lessonDetail: str = Field(default="{}", description="课时大纲JSON字符串")
    fontSize: int = Field(default=16, description="正文字号")
    activeProfile: str = Field(default="dev", description="环境标识: dev/test/product")
    fileContentStyle: str = Field(default="0", description="模板风格编号")
    schoolLogo: str = Field(default="", description="学校Logo URL或路径")
```

## 3.2 讲题PPT请求

```python
class TopicPptRequest(BaseModel):
    questionId: str = Field(default="123456", description="题目ID")
    stage: int = Field(default=12, description="学段")
    subject: int = Field(default=9912, description="学科")
    detail: str = Field(default="{}", description="题目详情JSON字符串")
    fontSize: int = Field(default=16, description="正文字号")
    activeProfile: str = Field(default="dev", description="环境标识: dev/test/product")
    fileContentStyle: str = Field(default="0", description="模板风格编号")
```

---

## 4. Pages数据模型 (models/pages.py)

## 4.1 Page类型枚举

```python
from enum import Enum

class PageType(str, Enum):
    TITLE = "title"         # 标题页
    CATALOG = "ml"          # 目录页
    FORMULA = "formula"     # 公式内容页
    QUESTION = "question"   # 题目页
    ANSWER = "answer"       # 答案页
```

## 4.2 Page数据结构

```python
from dataclasses import dataclass
from typing import Optional, Any

@dataclass
class PageData:
    title: str
    content: Optional[Any] = None       # 分页后的元素列表 或 None
    subTitle: Optional[str] = None      # 仅title类型使用
    text_title: Optional[str] = None    # 仅formula类型，副标题

@dataclass
class Page:
    type: PageType
    data: PageData
```

---
