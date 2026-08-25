# 公式批量转换器

> 来源：备课PPT生成-AI代码驱动规格书 v1.2


## 9.1 设计原则

收集一页内所有LaTeX公式，**一次性批量请求** latex2mml 接口，替代逐个请求。

## 9.2 批量接口调用

```python
def batch_latex_to_omml(formulas: list[str], env: EnvConfig) -> dict[str, str]:
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
```

## 9.3 MathML修正规则 (revise_mathml)

保持原逻辑：修正MathML命名空间和结构，确保mathml2omml能正确转换。从原 ppt_parser.py 直接迁移。

## 9.4 OMML适配规则 (adapt_omml)

```python
def adapt_omml(omml: str) -> str:
    """
    1. 修正 groupChr 标签：补全 <m:groupChrPr> 闭合
       - <m:groupChrPr><m:chr m:val="..."...><m:pos m:val="top"/></m:groupChr>
         → <m:groupChrPr><m:chr m:val="..."/><m:pos m:val="top"/></m:groupChrPr>
    2. 添加 oMath 命名空间：
       <m:oMath> → <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    """
```

---
