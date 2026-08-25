"""
测试配置和公共fixture
"""

import sys
import os
from pathlib import Path

# 将src目录添加到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import pytest
from pptx.util import Inches

# 布局常量（与config.py一致）
TEXT_BOX_HEIGHT = int(Inches(5))   # EMU
TEXT_BOX_WIDTH = int(Inches(10.8))  # EMU


@pytest.fixture
def text_box_height():
    """文本框高度"""
    return TEXT_BOX_HEIGHT


@pytest.fixture
def text_box_width():
    """文本框宽度"""
    return TEXT_BOX_WIDTH


@pytest.fixture
def sample_env():
    """测试环境配置"""
    from src.models import EnvConfig
    return EnvConfig(
        profile="test",
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://test.api/latex",
        wechat_webhook="https://test.webhook",
        wechat_mentioned=(),
    )


@pytest.fixture
def font_size():
    """默认字号"""
    return 16
