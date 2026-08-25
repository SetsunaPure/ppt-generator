"""
PPT生成器 - 兼容性处理模块

从原ppt_parser.py迁移的兼容性函数
- MathML修正规则
- OMML适配规则
- HTML清洗函数
"""

import re
from lxml import etree


# ============================================================
# MathML修正规则
# ============================================================

def revise_mathml(mathml: str) -> str:
    """
    修正MathML命名空间和结构，确保mathml2omml能正确转换

    规则：
    1. 补全缺失的命名空间
    2. 修正不完整的标签闭合
    """
    if not mathml or not mathml.strip():
        return mathml

    try:
        # 添加默认命名空间前缀
        mathml = mathml.strip()
        if not mathml.startswith('<?xml'):
            # 如果没有声明，添加默认的math元素包装
            if '<math' not in mathml:
                mathml = f'<math xmlns="http://www.w3.org/1998/Math/MathML">{mathml}</math>'

        # 确保根元素有命名空间
        root = etree.fromstring(mathml.encode('utf-8'))
        ns = "http://www.w3.org/1998/Math/MathML"

        # 修正没有命名空间的元素
        for elem in root.iter():
            if elem.tag.startswith('{'):
                # 已经是带命名空间的
                pass
            else:
                # 添加命名空间
                elem.tag = f'{{{ns}}}{elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag}'

        return etree.tostring(root, encoding='unicode', pretty_print=False)
    except Exception:
        # 解析失败时返回原始内容
        return mathml


# ============================================================
# OMML适配规则
# ============================================================

def adapt_omml(omml: str) -> str:
    """
    OMML适配规则

    1. 修正 groupChr 标签：补全 <m:groupChrPr> 闭合
       - <m:groupChrPr><m:chr m:val="..."...><m:pos m:val="top"/></m:groupChr>
         → <m:groupChrPr><m:chr m:val="..."/><m:pos m:val="top"/></m:groupChrPr>
    2. 添加 oMath 命名空间：
       <m:oMath> → <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    """
    if not omml or not omml.strip():
        return omml

    # 规则1: 补全 groupChr 标签闭合
    # 匹配 <m:groupChrPr>...<m:chr m:val="..."...><m:pos.../> 后面缺少 </m:groupChrPr> 的情况
    omml = re.sub(
        r'(<m:groupChrPr>)(.*?)(<m:chr[^>]*/?>)(.*?)(<m:pos[^>]*/?>)(.*?)(</m:groupChr>)(?!\s*</m:groupChrPr>)',
        r'\1\2\3\4\5\6</m:groupChrPr>\7',
        omml,
        flags=re.DOTALL
    )

    # 规则2: 添加 oMath 命名空间
    omath_ns = "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""
    omml = re.sub(
        r'<m:oMath(?!(\s+xmlns|m:|\s*>)',
        f'<m:oMath {omath_ns}',
        omml
    )

    return omml


# ============================================================
# HTML清洗函数
# ============================================================

def clean_html_for_pptx(html: str) -> str:
    """
    清洗HTML中不兼容PPT的标签和属性

    处理规则：
    1. 移除不兼容的标签但保留内容
    2. 修正常见的HTML实体
    3. 处理下划线、删除线等样式标签
    """
    if not html:
        return html

    # 移除危险的脚本和样式
    html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL | re.IGNORECASE)

    # 修正HTML实体
    html = html.replace('&nbsp;', ' ')
    html = html.replace('&lt;', '<')
    html = html.replace('&gt;', '>')
    html = html.replace('&amp;', '&')
    html = html.replace('&quot;', '"')
    html = html.replace('&#160;', ' ')

    # 处理下划线标签，保留内容
    html = re.sub(r'<u>(.*?)</u>', r'\1', html, flags=re.DOTALL)
    html = re.sub(r'<s>(.*?)</s>', r'\1', html, flags=re.DOTALL)
    html = re.sub(r'<strike>(.*?)</strike>', r'\1', html, flags=re.DOTALL)

    # 移除字体标签但保留内容
    html = re.sub(r'</?font[^>]*>', '', html, flags=re.IGNORECASE)

    # 移除span标签但保留内容（保留特定class如fill/longFill/brack）
    def preserve_span_with_class(match):
        content = match.group(1)
        class_attr = match.group(2) if len(match.groups()) > 1 else ''
        if class_attr in ('fill', 'longFill', 'brack'):
            return f'<span class="{class_attr}">{content}</span>'
        return content

    html = re.sub(r'<span([^>]*)>(.*?)</span>', preserve_span_with_class, html, flags=re.DOTALL)

    # 清理多余的空白
    html = re.sub(r'\s+', ' ', html)
    html = html.strip()

    return html


# ============================================================
# Markdown转纯文本
# ============================================================

def md_to_text(md: str) -> str:
    """
    去除Markdown格式，返回纯文本

    处理规则：
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
    """
    if not md:
        return md

    text = md

    # 移除图片，保留alt文本
    text = re.sub(r'!\[([^\]]*)\]\([^\)]*\)', r'\1', text)

    # 处理链接，保留链接文本
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)

    # 移除粗体标记
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)

    # 移除斜体标记
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)

    # 移除行内代码
    text = re.sub(r'`([^`]+)`', r'\1', text)

    # 处理代码块
    text = re.sub(r'```[\s\S]*?```', '', text)

    # 移除标题标记
    text = re.sub(r'^#+\s+', '', text, flags=re.MULTILINE)

    # 移除列表标记
    text = re.sub(r'^[\s]*[-*+]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^[\s]*\d+\.\s+', '', text, flags=re.MULTILINE)

    # 移除引用标记
    text = re.sub(r'^>\s+', '', text, flags=re.MULTILINE)

    # 移除分割线
    text = re.sub(r'^[-*_]{3,}$', '', text, flags=re.MULTILINE)

    # 合并多余空行
    text = re.sub(r'\n{3,}', '\n\n', text)

    return text.strip()
