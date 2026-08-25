"""
公式转换测试 - P1优先级

测试公式批量转换功能（使用mock）
"""

import sys
from pathlib import Path

# 将src目录添加到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import pytest
from unittest.mock import patch, MagicMock

from src.formula_converter import (
    batch_latex_to_omml,
    batch_convert_formulas,
    convert_single_formula,
    FormulaCache,
    _fallback_convert,
    _create_fallback_omml,
)
from src.models import EnvConfig


@pytest.fixture
def mock_env():
    """Mock环境配置"""
    return EnvConfig(
        profile="test",
        proxy=None,
        oss_region="beijing",
        latex_api_url="http://test.api/latex",
        wechat_webhook="https://test.webhook",
        wechat_mentioned=(),
    )


class TestFormulaConversion:
    """公式转换测试"""

    @patch("src.formula_converter.make_request")
    def test_batch_convert_success(self, mock_request, mock_env):
        """批量转换正常返回"""
        # Mock API响应
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "data": ["<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>"]
        }
        mock_request.return_value = mock_response

        formulas = ["x^2"]
        result = batch_convert_formulas(formulas, mock_env)

        assert len(result) == 1
        assert "x^2" in result
        assert "<m:oMath" in result["x^2"]

    @patch("src.formula_converter.make_request")
    def test_batch_convert_multiple(self, mock_request, mock_env):
        """批量转换多个公式"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "data": [
                "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
                "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>",
            ]
        }
        mock_response.raise_for_status = MagicMock()
        mock_request.return_value = mock_response

        formulas = ["x", "y"]
        result = batch_convert_formulas(formulas, mock_env)

        assert len(result) == 2
        assert "x" in result
        assert "y" in result

    @patch("src.formula_converter.make_request")
    def test_fallback_to_local(self, mock_request, mock_env):
        """接口失败时回退到本地转换"""
        # 模拟网络错误
        mock_request.side_effect = Exception("timeout")

        formulas = ["x^2"]
        result = batch_convert_formulas(formulas, mock_env)

        # 本地兜底应该也能返回结果
        assert len(result) == 1
        assert "x^2" in result
        # 结果可能是fallback OMML
        assert "m:oMath" in result["x^2"] or "x^2" in result["x^2"]

    @patch("src.formula_converter.make_request")
    def test_empty_formulas(self, mock_request, mock_env):
        """空公式列表"""
        result = batch_convert_formulas([], mock_env)
        assert result == {}

    def test_fallback_convert(self):
        """本地兜底转换"""
        result = _fallback_convert("x^2")
        assert "m:oMath" in result or "x^2" in result

    def test_fallback_omml_structure(self):
        """兜底OMML结构正确"""
        result = _create_fallback_omml("E=mc^2")
        assert "<m:oMath" in result
        assert "E=mc^2" in result or "E" in result


class TestFormulaCache:
    """公式缓存测试"""

    def test_cache_get_put(self, mock_env):
        """缓存存取"""
        cache = FormulaCache(mock_env)
        cache.put("x^2", "<m:oMath>result</m:oMath>")

        result = cache.get("x^2")
        assert result == "<m:oMath>result</m:oMath>"

    def test_cache_miss(self, mock_env):
        """缓存未命中"""
        cache = FormulaCache(mock_env)
        result = cache.get("nonexistent")
        assert result is None

    def test_cache_clear(self, mock_env):
        """清空缓存"""
        cache = FormulaCache(mock_env)
        cache.put("x^2", "result")
        cache.clear()

        assert cache.get("x^2") is None
        assert len(cache) == 0

    def test_cache_length(self, mock_env):
        """缓存长度"""
        cache = FormulaCache(mock_env)
        cache.put("a", "1")
        cache.put("b", "2")

        assert len(cache) == 2


# ============================================================
# 运行测试
# ============================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
