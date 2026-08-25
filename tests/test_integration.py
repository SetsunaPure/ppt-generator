"""
PPT生成器 - 集成测试

验证完整流程：大纲解析 → 分页 → PPT构建 → 文件输出
不依赖OSS，仅验证本地生成能力
"""

import sys
import json
import tempfile
from pathlib import Path

# 确保项目根目录在sys.path中
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.models import get_env_config
from src.outline_parser import OutlineParser, TopicParser, TOPIC_TEMPLATES
from src.layout_engine import split_page
from src.builder import PptBuilder


# ============================================================
# 测试用大纲数据
# ============================================================

_OUTLINE_LIST = [
    {
        "outlineCode": 6,
        "name": "教学设计",
        "children": [
            {
                "outlineCode": 7,
                "name": "教学规划",
                "children": [
                    {
                        "outlineCode": 13,
                        "name": "教学目标",
                        "content": [
                            {
                                "content": "<p>1. 掌握二次函数的基本概念和图像特征</p>"
                                           "<p>2. 能够根据条件确定二次函数的解析式</p>"
                                           "<p>3. 理解二次函数与一元二次方程的关系</p>"
                            }
                        ]
                    },
                    {
                        "outlineCode": 14,
                        "name": "重点难点",
                        "content": [
                            {
                                "content": "<p><strong>重点：</strong>二次函数图像的开口方向、对称轴、顶点坐标</p>"
                                           "<p><strong>难点：</strong>二次函数与一元二次方程的内在联系</p>"
                            }
                        ]
                    }
                ]
            },
            {
                "outlineCode": 8,
                "name": "教学过程",
                "children": [
                    {
                        "outlineCode": 16,
                        "name": "新课导入",
                        "content": [
                            {
                                "content": "<p>回顾一次函数 y=kx+b 的图像和性质，思考：当自变量最高次数为2时，函数图像会是什么样子？</p>"
                            }
                        ]
                    },
                    {
                        "outlineCode": 18,
                        "name": "知识梳理",
                        "content": [
                            {
                                "content": "<p>一般形式：$y=ax^2+bx+c$（$a \\neq 0$）</p>"
                                           "<p>顶点式：$y=a(x-h)^2+k$，顶点坐标为 $(h,k)$</p>"
                                           "<p>交点式：$y=a(x-x_1)(x-x_2)$，其中 $x_1, x_2$ 为与x轴交点的横坐标</p>"
                            }
                        ]
                    }
                ]
            },
            {
                "outlineCode": 9,
                "name": "教学总结",
                "content": [
                    {
                        "content": "<p>本节课学习了二次函数的定义、三种表达形式及其图像的基本性质，为后续学习二次函数的应用打下基础。</p>"
                    }
                ]
            }
        ]
    }
]

# 包装为API响应格式（OutlineParser.parse 期望 dict）
LESSON_OUTLINE_JSON = {
    "title": "二次函数的概念与性质",
    "subTitle": "函数的概念与性质",
    "outlineList": _OUTLINE_LIST,
}


# ============================================================
# 集成测试
# ============================================================

class TestIntegration:
    """集成测试 - 完整流程验证"""

    def setup_method(self):
        self.env = get_env_config("dev")

    def test_lesson_outline_parse(self):
        """测试备课大纲解析"""
        parser = OutlineParser()
        pages = parser.parse(LESSON_OUTLINE_JSON, font_size=16)

        assert len(pages) > 0, "大纲解析应产生至少1页"

        # 第一页应为标题页
        from src.config import PageType
        assert pages[0].type == PageType.TITLE, "第一页应为标题页"
        assert "课时名称" in pages[0].data.title, "标题应包含'课时名称'"

        # 第二页应为内容页（教学目标/重点难点/教学方法在目录页之前，对齐老代码逻辑）
        assert pages[1].type in (PageType.FORMULA, PageType.CATALOG), "第二页应为内容页或目录页"

        # 应有目录页
        catalog_pages = [p for p in pages if p.type == PageType.CATALOG]
        assert len(catalog_pages) > 0, "应有目录页"

        # 应有内容页
        content_pages = [p for p in pages if p.type in (PageType.FORMULA, PageType.QUESTION)]
        assert len(content_pages) > 0, "应有内容页"

        print(f"  [PASS] 备课大纲解析: {len(pages)} 页")

    def test_lesson_ppt_build(self):
        """测试备课PPT构建（本地文件输出）"""
        parser = OutlineParser()
        pages = parser.parse(LESSON_OUTLINE_JSON, font_size=16)

        builder = PptBuilder(
            env=self.env,
            pages=pages,
            font_size=16,
            file_content_style="0",
            logo_path="",
            template_prefix="lesson"
        )

        output_path = builder.build()

        assert output_path is not None, "应返回输出路径"
        assert Path(output_path).exists(), f"输出文件应存在: {output_path}"
        assert output_path.endswith(".pptx"), "输出应为.pptx文件"

        # 文件大小至少1KB（空模板都大于1KB）
        file_size = Path(output_path).stat().st_size
        assert file_size > 1024, f"文件大小应大于1KB，实际: {file_size}"

        print(f"  [PASS] 备课PPT构建: {output_path} ({file_size} bytes)")

    def test_layout_engine_with_mixed_content(self):
        """测试布局引擎处理混合内容"""
        from src.models import Block

        blocks = [
            Block(type="text", content="这是一段测试文字", inline=True, width=500000, height=300000),
            Block(type="newline", inline=False, width=0, height=0),
            Block(type="text", content="另一段文字", inline=True, width=400000, height=300000),
            Block(type="latex", content="y=ax^2+bx+c", inline=True, width=600000, height=300000),
        ]

        # 分页（使用layout_blocks，因为blocks已经是Block对象）
        from src.layout_engine import layout_blocks
        pages_blocks = layout_blocks(blocks, font_size_pt=16)

        assert len(pages_blocks) > 0, "应产生至少1页"
        print(f"  [PASS] 混合内容分页: {len(pages_blocks)} 页, 首页 {len(pages_blocks[0])} blocks")

    def test_env_config(self):
        """测试环境配置"""
        # dev环境无代理
        dev_env = get_env_config("dev")
        assert dev_env.proxy is None, "dev环境不应有代理"
        assert dev_env.oss_region == "beijing", "dev环境OSS应为beijing"

        # product环境有代理
        prod_env = get_env_config("product")
        assert prod_env.proxy is not None, "product环境应有代理"
        assert prod_env.oss_region == "hangzhou", "product环境OSS应为hangzhou"

        # 无效环境应报错
        try:
            get_env_config("invalid")
            assert False, "无效环境应抛出异常"
        except ValueError:
            pass

        print("  [PASS] 环境配置验证")


def run_integration_tests():
    """运行集成测试"""
    test = TestIntegration()
    test.setup_method()

    print("=" * 50)
    print("PPT生成器 - 集成测试")
    print("=" * 50)

    tests = [
        ("环境配置", test.test_env_config),
        ("大纲解析", test.test_lesson_outline_parse),
        ("混合内容分页", test.test_layout_engine_with_mixed_content),
        ("PPT构建", test.test_lesson_ppt_build),
    ]

    passed = 0
    failed = 0

    for name, fn in tests:
        try:
            fn()
            passed += 1
        except Exception as e:
            print(f"  [FAIL] {name}: {e}")
            failed += 1

    print("=" * 50)
    print(f"结果: {passed} 通过, {failed} 失败")
    print("=" * 50)

    return failed == 0


if __name__ == "__main__":
    success = run_integration_tests()
    sys.exit(0 if success else 1)
