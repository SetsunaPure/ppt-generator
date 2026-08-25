# 企微告警、日志、缺陷修复、兼容性

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


```python
def send_wechat_notify(env: EnvConfig, message: str) -> bool:
    """
    通过企微机器人发送text消息，@指定手机号

    - 使用 make_request("POST", env.wechat_webhook, env, ...) 发送
    - 请求头：Content-Type: application/json
    - 请求体：{"msgtype":"text","text":{"content":message,"mentioned_mobile_list":env.wechat_mentioned}}
    - 成功返回True，失败返回False（不抛异常，告警不能影响主流程）
    """
```

---

## 17. 日志模块 (utils/logger.py)

```python
import logging.handlers  # 必须显式import，原代码缺失此行导致运行报错
import json
from datetime import datetime

def get_logger(log_filename: str = "./log/gen_ppt.log") -> logging.Logger:
    """
    JSON格式日志，RotatingFileHandler
    - 5MB轮转，保留2个备份，UTF-8编码
    - 日志字段：timestamp, level, message, module, function, line, thread, process
    - 有异常时追加exception字段
    - 防重复初始化：按log_filename生成唯一logger名称，检查已有handlers
    """
```

---

## 18. Markdown转纯文本 (utils/md_helper.py)

```python
import re

def md_to_text(md: str) -> str:
    """
    去除Markdown格式，返回纯文本
    - 移除图片 ![alt](url) → 保留alt
    - 移除链接 [text](url) → 保留text
    - 移除粗体 **text** / __text__ → text
    - 移除斜体 *text* / _text_ → text
    - 移除行内代码 `code` → code
    - 移除代码块 ```...``` → 内容
    - 移除标题 # 前缀
    - 移除列表 * - + 前缀
    - 移除引用 > 前缀
    - 合并多余空行
    - 返回strip后的结果
    """
```

---

## 19. 已知缺陷修复清单

实现时**必须**修复以下原代码缺陷：

| # | 原缺陷 | 修复方式 |
|---|--------|---------|
| 1 | `logging.handlers` 未import | utils/logger.py 显式 `import logging.handlers` |
| 2 | `next()` 找不到节点返回整数1 | 返回None，调用方做None检查 |
| 3 | 变量名 `list` 覆盖内置函数 | 改为 `pages` 或 `result` |
| 4 | 硬编码questionId调试代码 | 删除所有硬编码questionId判断 |
| 5 | 废弃调试pass块 | 删除 `lt_index+1==2`、`index+1==1` 等无效判断 |
| 6 | 代理地址硬编码 | 统一走 EnvConfig + make_request |
| 7 | shapes索引硬编码 | 走 SLIDE_LAYOUTS 配置映射 |

---

## 20. 保持不变的部分

以下模块/逻辑从原代码直接迁移，不做改动：

| 模块/函数 | 说明 |
|-----------|------|
| `oss_con.py` → OssOperationHandler | OSS连接与上传 |
| `revise_mathml()` | MathML修正函数 |
| `clean_html_for_pptx()` | HTML清洗函数 |
| `adapt_omml()` | OMML适配函数（已在本规格书9.4描述增强逻辑） |
| `convert_to_internal_url()` | OSS URL内外网转换 |
| `get_scaled_dimensions()` | Logo缩放尺寸计算 |

---

## 21. 接口兼容性

## 21.1 入参完全兼容

两个HTTP接口的请求参数名、类型、默认值与原代码完全一致，上游调用方无需改动。

## 21.2 返回值

正常：返回OSS文件链接字符串
异常：返回None（与原代码行为一致）

## 21.3 模板文件

模板文件命名规则不变：`lesson_template_{fileContentStyle}.pptx` / `topic_template_{fileContentStyle}.pptx`
