#!/usr/bin/env python3
"""LaTeX → MathML → OMML 转换（PowerPoint 原生可编辑公式）。

链路: latex2mathml（纯 Python）→ MathML → MML2OMML.XSL（Office 自带）→ OMML
输出元素可直接 append 到 pptx 段落的 <a:p> 下（外层包装 a14:m/m:oMathPara）。
"""
import os
import pathlib

import latex2mathml.converter
from lxml import etree

_HERE = pathlib.Path(__file__).resolve().parent
_XSL_CANDIDATES = [
    os.environ.get("PPT_STUDIO_MML2OMML"),
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
]
_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
_xslt = None


def _get_xslt():
    global _xslt
    if _xslt is None:
        for c in _XSL_CANDIDATES:
            if c and pathlib.Path(c).exists():
                _xslt = etree.XSLT(etree.parse(c))
                return _xslt
        raise RuntimeError("MML2OMML.XSL 未找到（需安装 Microsoft Office）")
    return _xslt


def latex_to_omml_element(latex: str):
    """LaTeX 文本 → m:oMath XML 元素。转换失败抛异常，由调用方兜底。"""
    latex = latex.strip()
    for pat in (r"^\\\((.*)\\\)$", r"^\$\$(.*)\$\$$", r"^\$(.*)\$$"):
        m = re.match(pat, latex, re.S)
        if m:
            latex = m.group(1).strip()
            break
    mathml = latex2mathml.converter.convert(latex)
    tree = etree.fromstring(mathml)
    result = _get_xslt()(tree)
    return result.getroot()


def wrap_for_paragraph(omml_root):
    """包装为可 append 到 <a:p> 的 a14:m 容器。"""
    container = etree.Element("{%s}m" % _A14, nsmap={"a14": _A14, "m": _M})
    para = etree.SubElement(container, "{%s}oMathPara" % _M)
    para.append(omml_root)
    return container


def is_formula(text: str) -> bool:
    """整段被 \\( \\) / $$ 包围时判定为公式段。"""
    t = (text or "").strip()
    return bool(re.match(r"^\\\(.*\\\)$", t, re.S) or
                re.match(r"^\$\$.*\$\$$", t, re.S) or
                re.match(r"^\$[^$].*\$$", t, re.S))


def strip_delims(text: str) -> str:
    t = (text or "").strip()
    for pat in (r"^\\\((.*)\\\)$", r"^\$\$(.*)\$\$$", r"^\$(.*)\$$"):
        m = re.match(pat, t, re.S)
        if m:
            return m.group(1).strip()
    return t


import re  # noqa: E402
