# 项目概述与环境约束

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 1.1 功能定义

教育备课PPT生成微服务，接收课时大纲JSON，自动生成包含公式、图片、表格、题目的PPT文件并上传OSS返回链接。

## 1.2 核心流程

```
HTTP请求 → 参数校验 → JSON大纲解析(pages列表) → 排版模板生成 → 内容分页 → PPT渲染填充 → OSS上传 → 返回链接
```

## 1.3 环境约束（强制）

| 环境 | 外网 | 代理 | OSS区域 | LaTeX转换接口 |
|------|------|------|---------|--------------|
| dev | 有 | 不需要 | beijing | http://exam.canpoint.cn/python_api_latex/latex/latex2mml |
| test | 有 | 不需要 | beijing | http://exam.canpoint.cn/python_api_latex/latex/latex2mml |
| product | 无 | http://10.1.1.219:8500 | hangzhou | http://10.1.1.231/python_api_latex/latex/latex2mml |

**规则**：所有外网请求（企微消息、图片下载、LaTeX接口调用）必须根据环境配置决定是否走代理。零硬编码。

## 1.4 技术栈

- Python 3.13
- FastAPI + uvicorn（端口8001，root_path=/api）
- python-pptx（PPT生成）
- BeautifulSoup4 + lxml（HTML解析，替代逐字符遍历）
- latex2mathml + mathml2omml（公式转换，本地兜底）
- Pillow（图片尺寸计算）
- requests（HTTP请求，支持代理）

## 1.5 目录结构

```
lesson/
├── main.py                    # FastAPI入口，路由定义
├── config/
│   ├── __init__.py
│   ├── env.py                 # 环境配置（代理/OSS/LaTeX接口地址）
│   ├── outline_map.py         # 大纲解析配置表
│   ├── topic_template_map.py  # 讲题模板配置表
│   └── slide_layout_map.py    # PPT模板shape映射配置
├── models/
│   ├── __init__.py
│   ├── request.py             # 请求参数模型(Pydantic)
│   ├── pages.py               # Pages数据模型
│   └── question.py            # 题目数据模型
├── parser/
│   ├── __init__.py
│   ├── outline_parser.py      # 大纲解析器（配置驱动）
│   ├── content_splitter.py    # HTML内容分页器（BS4重构）
│   ├── question_parser.py     # 题目解析器
│   └── topic_template.py      # 讲题模板解析器（配置驱动）
├── renderer/
│   ├── __init__.py
│   ├── ppt_builder.py         # PPT构建主流程
│   ├── layout_generator.py    # 排版模板生成（Logo注入）
│   ├── content_renderer.py    # 内容渲染引擎（文本/公式/图片/表格）
│   └── formula_batch.py       # 公式批量转换器
├── service/
│   ├── __init__.py
│   ├── lesson_service.py      # 备课PPT业务逻辑
│   └── topic_service.py       # 讲题PPT业务逻辑
├── utils/
│   ├── __init__.py
│   ├── oss_handler.py         # OSS上传
│   ├── wechat_notify.py       # 企微告警通知
│   ├── logger.py              # 日志模块
│   └── image_helper.py        # 图片下载与处理
├── oss_con.py                 # OSS连接（已有，保持）
└── file/                      # 运行时文件目录
    ├── template/              # PPT模板文件
    ├── layout_template/       # 排版模板（临时）
    ├── output/                # 输出文件（临时）
    └── log/                   # 日志文件
```

---
