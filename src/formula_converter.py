"""
PPT生成器 - 公式批量转换器

设计原则：
- 收集一页内所有LaTeX公式，一次性批量请求latex2mml接口
- 替代逐个请求，提高效率
- 接口失败时用本地latex2mathml兜底
"""

import logging
from typing import Dict, List, Optional, Any

from .models import EnvConfig
from .utils import make_request, get_logger
from .compat import revise_mathml, adapt_omml

logger = get_logger()


# ============================================================
# 公式批量转换
# ============================================================

def batch_latex_to_omml(
    formulas: List[str],
    env: EnvConfig
) -> Dict[str, str]:
    """
    批量公式转换

    入参：LaTeX公式字符串列表
    返回：{latex_str: omml_str} 映射

    流程：
    1. 构建请求体：{"formulas": [formula1, formula2, ...]}
    2. POST到 env.latex_api_url（通过make_request自动代理）
    3. 逐个解析响应，失败的用本地 latex2mathml 兜底：
       LaTeX → latex2mathml.converter.convert() → revise_mathml() → mathml2omml.convert() → adapt_omml()
    4. 批量接口不可用时，降级为逐个请求但仍统一经make_request调用
    """
    if not formulas:
        return {}

    results: Dict[str, str] = {}

    try:
        # 批量请求
        payload = {"formulas": formulas}
        response = make_request(
            "POST",
            env.latex_api_url,
            env,
            json=payload,
            timeout=30
        )
        response.raise_for_status()

        data = response.json()
        formulas_data = data.get("data", [])

        # 解析响应
        if isinstance(formulas_data, list):
            for i, formula in enumerate(formulas):
                if i < len(formulas_data):
                    omml = formulas_data[i]
                    if omml:
                        results[formula] = adapt_omml(omml)
                    else:
                        # 转换失败，尝试本地兜底
                        results[formula] = _fallback_convert(formula)
                else:
                    results[formula] = _fallback_convert(formula)
        else:
            # 响应格式不是列表，按单公式处理
            for formula in formulas:
                results[formula] = _fallback_convert(formula)

    except Exception as e:
        logger.warning(f"批量公式转换接口失败，使用本地兜底: {e}")
        # 批量接口失败，降级为逐个本地转换
        for formula in formulas:
            results[formula] = _fallback_convert(formula)

    return results


def batch_convert_formulas(
    formulas: List[str],
    env: EnvConfig
) -> Dict[str, str]:
    """
    批量转换公式（入口函数）

    与 batch_latex_to_omml 功能相同，提供更直观的命名
    """
    return batch_latex_to_omml(formulas, env)


# ============================================================
# 本地兜底转换
# ============================================================

def _fallback_convert(latex: str) -> str:
    """
    本地LaTeX到OMML的兜底转换

    流程：LaTeX → MathML → OMML
    """
    try:
        # LaTeX → MathML
        mathml = _latex_to_mathml(latex)
        if not mathml:
            return _create_fallback_omml(latex)

        # 修正MathML
        mathml = revise_mathml(mathml)

        # MathML → OMML
        omml = _mathml_to_omml(mathml)
        if omml:
            return adapt_omml(omml)

        return _create_fallback_omml(latex)

    except Exception as e:
        logger.warning(f"本地公式转换失败: {latex[:30]}... error: {e}")
        return _create_fallback_omml(latex)


def _latex_to_mathml(latex: str) -> Optional[str]:
    """
    LaTeX → MathML 转换

    使用 latex2mathml 库
    """
    try:
        import latex2mathml.converter
        mathml = latex2mathml.converter.convert(latex)
        return mathml
    except ImportError:
        logger.warning("latex2mathml 库未安装")
        return None
    except Exception as e:
        logger.warning(f"LaTeX到MathML转换失败: {e}")
        return None


def _mathml_to_omml(mathml: str) -> Optional[str]:
    """
    MathML → OMML 转换

    使用 mathml2omml 库
    """
    try:
        import mathml2omml
        omml = mathml2omml.convert(mathml)
        return omml
    except ImportError:
        logger.warning("mathml2omml 库未安装")
        return None
    except Exception as e:
        logger.warning(f"MathML到OMML转换失败: {e}")
        return None


def _create_fallback_omml(latex: str) -> str:
    """
    创建兜底OMML

    当所有转换都失败时，返回一个基本的OMML结构
    """
    # 转义特殊字符
    escaped = (
        latex
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

    return (
        f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        f'<m:r><m:t>{escaped}</m:t></m:r>'
        f'</m:oMath>'
    )


# ============================================================
# 单个公式转换（便捷函数）
# ============================================================

def convert_single_formula(latex: str, env: EnvConfig) -> str:
    """
    转换单个LaTeX公式

    优先使用批量接口（虽然是单公式）
    """
    result = batch_latex_to_omml([latex], env)
    return result.get(latex, _create_fallback_omml(latex))


# ============================================================
# 公式缓存管理
# ============================================================

class FormulaCache:
    """
    公式转换缓存

    用于在渲染时缓存已转换的公式，避免重复转换
    """

    def __init__(self, env: EnvConfig):
        self.env = env
        self._cache: Dict[str, str] = {}

    def get(self, latex: str) -> Optional[str]:
        """从缓存获取"""
        return self._cache.get(latex)

    def put(self, latex: str, omml: str):
        """放入缓存"""
        self._cache[latex] = omml

    def batch_convert(self, formulas: List[str]) -> Dict[str, str]:
        """
        批量转换并缓存

        只转换缓存中没有的公式
        """
        to_convert = [f for f in formulas if f not in self._cache]
        if not to_convert:
            return {f: self._cache[f] for f in formulas}

        results = batch_latex_to_omml(to_convert, self.env)

        # 更新缓存
        for latex, omml in results.items():
            self._cache[latex] = omml

        return {f: self._cache[f] for f in formulas}

    def clear(self):
        """清空缓存"""
        self._cache.clear()

    def __len__(self) -> int:
        return len(self._cache)
