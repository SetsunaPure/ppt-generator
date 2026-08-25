# 备课PPT自动生成微服务

一个将教学大纲JSON自动转换为排版精良的PPT的微服务。

核心能力：输入一份课时大纲（含标题、知识点、公式、图片、表格、例题），输出一份排版完成的PPT文件并上传OSS。

## 架构设计

### 整体流水线

```
                    ┌──────────────┐
   HTTP Request ──→ │  FastAPI 层   │  参数校验 + 环境路由
                    └──────┬───────┘
                           │
                    ┌──────▼───────┐
                   │ OutlineParser │  将嵌套JSON按教学结构拆解为Page列表
                    └──────┬───────┘
                           │
                    ┌──────▼───────┐
                   │ LayoutEngine  │  积木流式布局：HTML→Block→分页
                    └──────┬───────┘
                           │
                    ┌──────▼───────┐
                   │  PptBuilder   │  加载模板 → 生成Slide骨架
                    └──────┬───────┘
                           │
                    ┌──────▼───────┐
                   │   Renderer    │  Block驱动渲染：文本/图片/公式/表格混排
                    └──────┬───────┘
                           │
                    ┌──────▼───────┐
                   │  OSS Upload   │  上传文件 + 微信通知
                    └──────────────┘
```

每个阶段职责单一、接口清晰，中间产物（Page / Block）是纯数据结构，可独立测试。

### 核心设计思路：积木流式布局

这是整个系统最关键的设计决策。

**问题**：PPT没有HTML的自动排版能力，python-pptx只支持绝对坐标放置元素。但教学内容是流式的——一段文字、一张图、一个公式、一个表格，它们的大小不固定，需要像网页一样"流"下去，满了就翻页。

**方案**：借鉴CSS的流式布局思想，建立一套"Block → 分页 → 绝对坐标"的三层转换：

```
HTML内容 → Block列表(带语义尺寸) → 分页裁切(空间耗尽判定) → 绝对坐标渲染
```

**Block 是核心抽象单元**：
- `text` Block：文本段，宽度按字符宽度因子计算（中文字符=1.0×字号，英文=0.55×字号）
- `image` Block：三级尺寸策略（NORMAL / LARGE / FULLPAGE）
- `table` Block：支持跨页拆分，每页重复表头
- `latex` Block：LaTeX公式，异步批量转SVG后作为图片Block插入
- `newline` Block：控制换行节奏

**页满判断 = 空间耗尽**：
```python
cursor_y + block.height > TEXT_BOX_HEIGHT → 翻页
```

**孤行保护**：当前页剩余空间不足3行时，整块内容跟随换页，避免出现"一页最后一行"的孤立内容。

**混排协调**：文本框内的段落流中，为块级元素（图片/表格）插入占位空段落预留空间，确保视觉对齐。

### 大纲解析：教学结构映射

大纲JSON是嵌套树结构，解析器按教学语义拆解为有序的Page列表：

```
课时大纲（JSON树）
  ├─ 标题页（固定首页）
  ├─ 教学设计
  │   ├─ 教学规划
  │   │   ├─ 教学目标 → formula页
  │   │   ├─ 重点难点 → formula页
  │   │   └─ 教学方法 → formula页
  │   └─ 教学过程
  │       ├─ 新课导入 → formula页
  │       ├─ 知识点A
  │       │   ├─ 知识梳理 → formula页
  │       │   ├─ 例题 → question页 + answer页
  │       │   └─ 课后作业 → question页
  │       └─ 知识点B ...
  └─ 教学总结 → formula页
```

每个outlineCode映射到特定的PageType，PageType决定使用哪种Slide布局模板。

### 渲染策略：Block分道

渲染引擎按Block类型走不同渲染路径：

| Block类型 | 渲染方式 |
|-----------|---------|
| text / latex | 写入文本框段落流（相对坐标） |
| image | 绝对坐标直接放置到Slide |
| table | 按行拆分，绝对坐标放置，跨页时重复表头 |
| AnswerGroupBlock | 内部决定2列/4列布局，渲染为组合文本框 |

**公式处理**：LaTeX公式先经过批量转换（batch_convert_formulas），缓存结果，避免重复请求转换服务。转换失败时降级为纯文本显示。

## 技术栈

- **Python 3.13** + **FastAPI**：HTTP服务层
- **python-pptx**：PPT文件操作（模板加载、元素放置、样式控制）
- **BeautifulSoup4**：HTML内容解析，提取文本/图片/表格/公式节点
- **latex2mathml**：LaTeX到MathML转换（公式渲染链路的一环）
- **lxml**：XML直接操作（处理pptx底层OOXML结构）
- **uvicorn**：ASGI服务器

## 项目结构

```
├── main.py                  # FastAPI 入口，HTTP层
├── cli.py                   # CLI工具，本地调试/演示
├── src/
│   ├── models.py            # 数据结构定义（Block/Page/EnvConfig）
│   ├── config.py            # 常量配置（页尺寸、布局参数、模板映射）
│   ├── outline_parser.py    # 大纲JSON → Page列表（教学结构映射）
│   ├── layout_engine.py     # 积木流式布局引擎（HTML → Block → 分页）
│   ├── renderer.py          # Block → PPT元素渲染
│   ├── builder.py           # 主流程编排（模板加载 → 渲染 → 输出）
│   ├── service.py           # 业务服务层（入口函数 + OSS上传）
│   ├── formula_converter.py # LaTeX公式批量转换
│   ├── utils.py             # 工具函数（日志/OSS/通知/文件处理）
│   └── compat.py            # 兼容层（HTML清洗/Markdown转换）
├── specs/                   # 详细设计文档（10个模块）
├── tests/                   # 单元测试 + 集成测试
└── file/template/           # PPT模板文件
```

## 快速开始

### CLI演示模式

```bash
# 用内置示例数据生成PPT（跳过OSS上传）
python cli.py demo

# 指定自定义大纲JSON
python cli.py lesson --json examples/lesson_outline.json

# 讲题PPT
python cli.py topic --json examples/question_detail.json
```

### 作为服务运行

```bash
# 启动FastAPI服务
python main.py
# 或
uvicorn main:app --host 0.0.0.0 --port 8001
```

### API接口

```
POST /api/lesson/create_ppt     备课PPT生成
POST /api/exam/topic/create_ppt  讲题PPT生成
GET  /api/health                 健康检查
```

请求示例（备课PPT）：
```json
{
  "lessonId": "lesson_001",
  "stage": 12,
  "subject": 9912,
  "lessonDetail": "{\"title\":\"二次函数\",\"subTitle\":\"九年级数学\",...}",
  "fontSize": 16,
  "activeProfile": "dev",
  "fileContentStyle": "0"
}
```

## 设计文档

`specs/` 目录包含每个模块的详细设计文档：

| 文档 | 内容 |
|------|------|
| 01_项目概述 | 系统架构、技术选型、部署方案 |
| 02_数据模型 | Block/Page/EnvConfig 等核心数据结构 |
| 03_大纲解析 | JSON树遍历逻辑、outlineCode映射规则 |
| 04_内容分页 | 积木流式布局算法、空间耗尽判定 |
| 05_题目解析 | 讲题PPT的模板匹配逻辑 |
| 06_公式转换 | LaTeX批量转换、缓存策略、降级方案 |
| 07_PPT渲染 | Block分道渲染、混排协调、OOXML操作 |
| 08_构建与服务 | 模板加载、主流程编排、OSS上传 |
| 09_工具与配置 | 字符宽度计算、兼容层、环境配置 |
| 10_单元测试 | 测试策略、Mock方案、覆盖范围 |

## License

MIT
