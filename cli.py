"""
PPT生成器 - CLI命令行工具

用于本地调试和演示，不依赖FastAPI服务运行
用法:
  python cli.py lesson --json <大纲JSON文件> [--font-size 16] [--style 0] [--profile dev]
  python cli.py topic  --json <题目JSON文件> [--font-size 16] [--style 0] [--profile dev]
  python cli.py demo   [--font-size 16] [--style 0]
"""

import argparse
import json
import sys
from pathlib import Path

# 确保项目根目录在sys.path中
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.models import LessonPptRequest, TopicPptRequest, get_env_config
from src.outline_parser import OutlineParser, TopicParser, TOPIC_TEMPLATES
from src.builder import PptBuilder


def cmd_lesson(args):
    """备课PPT生成"""
    json_path = Path(args.json)
    if not json_path.exists():
        print(f"错误: 文件不存在 {json_path}")
        sys.exit(1)

    with open(json_path, "r", encoding="utf-8") as f:
        lesson_detail = f.read()

    # 验证JSON格式
    try:
        json.loads(lesson_detail)
    except json.JSONDecodeError as e:
        print(f"错误: JSON格式无效 - {e}")
        sys.exit(1)

    request = LessonPptRequest(
        lessonId=args.lesson_id or "cli_test",
        lessonDetail=lesson_detail,
        fontSize=args.font_size,
        activeProfile=args.profile,
        fileContentStyle=args.style,
        schoolLogo=args.logo or "",
    )

    print(f"[CLI] 开始生成备课PPT...")
    print(f"  环境: {request.activeProfile}")
    print(f"  字号: {request.fontSize}")
    print(f"  风格: {request.fileContentStyle}")

    from src.service import create_lesson_ppt
    result = create_lesson_ppt(request)

    if result:
        print(f"[CLI] 生成成功! OSS链接: {result}")
    else:
        print("[CLI] 生成失败，请查看日志")
        sys.exit(1)


def cmd_topic(args):
    """讲题PPT生成"""
    json_path = Path(args.json)
    if not json_path.exists():
        print(f"错误: 文件不存在 {json_path}")
        sys.exit(1)

    with open(json_path, "r", encoding="utf-8") as f:
        detail = f.read()

    try:
        json.loads(detail)
    except json.JSONDecodeError as e:
        print(f"错误: JSON格式无效 - {e}")
        sys.exit(1)

    request = TopicPptRequest(
        questionId=args.question_id or "cli_test",
        detail=detail,
        fontSize=args.font_size,
        activeProfile=args.profile,
        fileContentStyle=args.style,
    )

    print(f"[CLI] 开始生成讲题PPT...")
    print(f"  环境: {request.activeProfile}")
    print(f"  字号: {request.fontSize}")
    print(f"  风格: {request.fileContentStyle}")

    from src.service import create_topic_ppt
    result = create_topic_ppt(request)

    if result:
        print(f"[CLI] 生成成功! OSS链接: {result}")
    else:
        print("[CLI] 生成失败，请查看日志")
        sys.exit(1)


def cmd_demo(args):
    """
    演示模式：用内置示例数据生成PPT，跳过OSS上传
    用于本地验证和培训演示
    """
    env = get_env_config(args.profile)

    # 构造完整的课时大纲（OutlineParser.parse期望dict，包含title/subTitle/outlineList）
    demo_outline = {
        "title": "二次函数的图像与性质",
        "subTitle": "九年级数学上册",
        "outlineList": [
            {
                "outlineCode": 6,
                "title": "教学设计",
                "children": [
                    {
                        "outlineCode": 7,
                        "title": "教学规划",
                        "children": [
                            {
                                "outlineCode": 13,
                                "title": "教学目标",
                                "content": [
                                    {"content": "<p>1. 掌握二次函数的基本概念</p><p>2. 理解二次函数图像的性质</p>"}
                                ]
                            },
                            {
                                "outlineCode": 14,
                                "title": "重点难点",
                                "content": [
                                    {"content": "<p>重点：二次函数图像的开口方向、对称轴、顶点坐标</p><p>难点：二次函数与一元二次方程的关系</p>"}
                                ]
                            }
                        ]
                    },
                    {
                        "outlineCode": 8,
                        "title": "教学过程",
                        "children": [
                            {
                                "outlineCode": 16,
                                "title": "新课导入",
                                "content": [
                                    {"content": "<p>回顾一次函数 $y=kx+b$ 的图像和性质，思考：当自变量最高次数为2时，函数图像会是什么样？</p>"}
                                ]
                            }
                        ]
                    },
                    {
                        "outlineCode": 9,
                        "title": "教学总结",
                        "content": [
                            {"content": "<p>本节课学习了二次函数的定义、图像特征和基本性质</p>"}
                        ]
                    }
                ]
            }
        ]
    }

    print("[CLI] Demo模式 - 用内置示例数据生成PPT（跳过OSS上传）")

    # 解析大纲
    parser = OutlineParser()
    pages = parser.parse(demo_outline, font_size=args.font_size)
    print(f"  大纲解析完成，共 {len(pages)} 页")

    # 构建PPT
    builder = PptBuilder(
        env=env,
        pages=pages,
        font_size=args.font_size,
        file_content_style=args.style,
        logo_path="",
        template_prefix="lesson"
    )
    output_path = builder.build()
    print(f"[CLI] Demo PPT生成成功: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="备课PPT生成器 CLI工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # --- lesson 命令 ---
    lesson_parser = subparsers.add_parser("lesson", help="备课PPT生成")
    lesson_parser.add_argument("--json", required=True, help="课时大纲JSON文件路径")
    lesson_parser.add_argument("--lesson-id", default="", help="课时ID")
    lesson_parser.add_argument("--font-size", type=int, default=16, help="字号")
    lesson_parser.add_argument("--style", default="0", help="模板风格(0-5)")
    lesson_parser.add_argument("--profile", default="dev", help="环境(dev/test/product)")
    lesson_parser.add_argument("--logo", default="", help="学校Logo URL")
    lesson_parser.set_defaults(func=cmd_lesson)

    # --- topic 命令 ---
    topic_parser = subparsers.add_parser("topic", help="讲题PPT生成")
    topic_parser.add_argument("--json", required=True, help="题目详情JSON文件路径")
    topic_parser.add_argument("--question-id", default="", help="题目ID")
    topic_parser.add_argument("--font-size", type=int, default=16, help="字号")
    topic_parser.add_argument("--style", default="0", help="模板风格(0-5)")
    topic_parser.add_argument("--profile", default="dev", help="环境(dev/test/product)")
    topic_parser.set_defaults(func=cmd_topic)

    # --- demo 命令 ---
    demo_parser = subparsers.add_parser("demo", help="演示模式（内置示例数据）")
    demo_parser.add_argument("--font-size", type=int, default=16, help="字号")
    demo_parser.add_argument("--style", default="0", help="模板风格(0-5)")
    demo_parser.add_argument("--profile", default="dev", help="环境(dev/test/product)")
    demo_parser.set_defaults(func=cmd_demo)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(0)

    args.func(args)


if __name__ == "__main__":
    main()
