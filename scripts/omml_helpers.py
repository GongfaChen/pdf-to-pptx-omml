#!/usr/bin/env python3
"""omml_helpers.py — PowerPoint OMML 公式 XML 构造库（配合 python-pptx 使用）。

设计原则（每条对应 references/pitfalls.md 中的实战坑）：

1. 所有构造 helper 返回 _Frag（已验证的 XML 片段）；裸 str 会被自动 T() 包裹，
   从 API 层面杜绝「裸文本塞进公式」类错误（坑 3 / 6 / 36）。
2. 空基底 / 空上标 / 空下标 / 空分子分母直接抛 ValueError（坑 1）。
3. 不提供 m:acc（矢量箭头在 PowerPoint→PDF 链路渲染成亚美尼亚标点，坑 35），
   不提供带 supHide 的 m:nary（会让 PowerPoint 丢弃整个文本框，坑 12）。
4. a14 命名空间在 mc:AlternateContent 根上声明；xmlns:m/a 声明在 oMath 根上（坑 2 / 10 / 11）。
5. display 公式（oMathPara）必须独占段落，行距建议 1.0（坑 13 / 15）——由调用方保证。

用法示例见 references/omml-cookbook.md；自检见 demo/smoke_test.py。
"""
from __future__ import annotations

import html
from xml.sax.saxutils import escape

# ---- 命名空间 ---------------------------------------------------------------

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# ---- 全局 run 样式（影响所有后续构造的 m:r） ---------------------------------

_STYLE = {"sz": None, "color": None, "font": "Cambria Math"}


def set_run_style(size=None, color=None, font=None):
    """设置后续所有公式 run 的默认样式。

    size: 字号（pt），如 20；内部转 sz 百分点（20pt -> sz="2000"）。
    color: RRGGBB 十六进制，如 "1F3B73"。浅色底 shape 里必须显式设深色（坑 18）。
    font: 默认 Cambria Math。
    """
    if size is not None:
        _STYLE["sz"] = str(int(size * 100))
    if color is not None:
        _STYLE["color"] = color
    if font is not None:
        _STYLE["font"] = font


def _arpr(italic=None, size=None, color=None, font=None):
    """a:rPr 属性串（管字号/颜色/字体/正体）。"""
    attrs = []
    sz = size * 100 if size else _STYLE["sz"]
    if sz:
        attrs.append(f'sz="{int(sz)}"')
    if italic is False:
        attrs.append('i="0"')
    col = color or _STYLE["color"]
    fnt = font or _STYLE["font"]
    inner = ""
    if col:
        inner += f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
    if fnt:
        inner += f'<a:latin typeface="{fnt}"/>'
    if not attrs and not inner:
        return ""
    body = f"<a:rPr {' '.join(attrs)}>{inner}</a:rPr>" if inner else f"<a:rPr {' '.join(attrs)}/>"
    return body


# ---- 片段标记 ----------------------------------------------------------------


class _Frag(str):
    """已构造的 OMML XML 片段。isinstance 检查用于区分「裸文本」与「XML」。"""


def _wrap(x):
    """裸 str 自动包 T()；_Frag 原样通过；None / 空串抛错。"""
    if isinstance(x, _Frag):
        return x
    if isinstance(x, str):
        if not x.strip():
            raise ValueError("空文本基底/操作数 —— 对应坑 1（PowerPoint 严格校验会拒开整个文件）")
        return T(x)
    if isinstance(x, (list, tuple)):
        return _Frag("".join(_wrap(i) for i in x))
    raise TypeError(f"期望 str/_Frag/list，收到 {type(x)}")


def _concat(x):
    if isinstance(x, (list, tuple)):
        return _Frag("".join(_wrap(i) for i in x))
    return _wrap(x)


# ---- 基础 run ----------------------------------------------------------------


def T(text, italic=None, size=None, color=None, font=None):
    """公式内文本 run。italic=False 生成正体（单位、文字说明用）。

    注意：T() 只接受纯文本；传 XML 字符串会被转义成字面文本（坑 4/36 的根源，
    本库通过 _Frag 类型检查在嵌套 helper 时自动避免）。
    """
    esc = escape(str(text))
    msty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if italic is False else ""
    arpr = _arpr(italic=italic, size=size, color=color, font=font)
    return _Frag(f'<m:r>{msty}{arpr}<m:t xml:space="preserve">{esc}</m:t></m:r>')


# ---- 结构 helper ---------------------------------------------------------------


def F(num, den):
    """分数 m:f。F(T('3'), T('5')) 或 F('3', '5') 均可。"""
    return _Frag(f"<m:f><m:num>{_wrap(num)}</m:num><m:den>{_wrap(den)}</m:den></m:f>")


def S(base, sup):
    """上标 m:sSup。S('v', '2') -> v²。"""
    return _Frag(f"<m:sSup><m:e>{_wrap(base)}</m:e><m:sup>{_wrap(sup)}</m:sup></m:sSup>")


def Q(base, sub):
    """下标 m:sSub。Q('F', 'cr') -> F_cr。注意 S=上标、Q=下标（坑 19）。"""
    return _Frag(f"<m:sSub><m:e>{_wrap(base)}</m:e><m:sub>{_wrap(sub)}</m:sub></m:sSub>")


def SQ(base, sup, sub):
    """上下标并存 m:sSubSup。"""
    return _Frag(
        f"<m:sSubSup><m:e>{_wrap(base)}</m:e><m:sup>{_wrap(sup)}</m:sup>"
        f"<m:sub>{_wrap(sub)}</m:sub></m:sSubSup>"
    )


def RAD(expr, deg=None):
    """根号 m:rad。deg=None 为平方根（degHide）。"""
    if deg is None:
        return _Frag(
            '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
            f"<m:e>{_wrap(expr)}</m:e></m:rad>"
        )
    return _Frag(f"<m:rad><m:deg>{_wrap(deg)}</m:deg><m:e>{_wrap(expr)}</m:e></m:rad>")


def LIM(base, under):
    """下极限 m:limLow（∫ 带下限、lim 下标等）。禁用 m:nary+supHide（坑 12）。"""
    return _Frag(
        f"<m:limLow><m:e>{_wrap(base)}</m:e><m:lim>{_wrap(under)}</m:lim></m:limLow>"
    )


def PAR(inner, left="(", right=")"):
    """定界符 m:d。left/right 传 '' 生成隐藏括号。"""
    if left == "(" and right == ")":
        return _Frag(f"<m:d><m:e>{_concat(inner)}</m:e></m:d>")
    dpr = (
        f'<m:dPr><m:begChr m:val="{escape(left)}"/><m:endChr m:val="{escape(right)}"/></m:dPr>'
    )
    return _Frag(f"<m:d>{dpr}<m:e>{_concat(inner)}</m:e></m:d>")


def ABS(inner):
    """绝对值 |x|。"""
    return PAR(inner, left="|", right="|")


def CV(*rows):
    """列向量 m:d > m:e > m:eqArr（坑 34）。两行 ≈1.3 行高、三行 ≈1.7 行高，预算行数时注意。

    用法：CV('x', S('v', 'T')) 或 CV(['x', 'y'], ['0', '1'])（后者为矩阵行）。
    """
    cells = []
    for r in rows:
        cell = _concat(r) if isinstance(r, (list, tuple)) else _wrap(r)
        cells.append(f"<m:e>{cell}</m:e>")
    eqarr = f"<m:eqArr>{''.join(cells)}</m:eqArr>"
    return _Frag(f"<m:d><m:e>{eqarr}</m:e></m:d>")


def MX(*rows):
    """矩阵 m:m（自动套圆括号）。

    用法：MX(['a', 'b'], ['c', 'd']) —— 每行是 list，元素为 str/_Frag。
    """
    mrs = []
    for row in rows:
        if not isinstance(row, (list, tuple)):
            raise TypeError("MX 每行必须是 list/tuple")
        mrs.append("<m:mr>" + "".join(f"<m:e>{_wrap(c)}</m:e>" for c in row) + "</m:mr>")
    return _Frag(f"<m:d><m:e><m:m>{''.join(mrs)}</m:m></m:e></m:d>")


def NORM(x):
    """范数 ‖x‖（用 U+2016 字面量，禁用 m:acc，坑 35）。"""
    return _Frag(f"{T('‖')}{_wrap(x)}{T('‖')}")


def M(*frags):
    """拼接多个片段为单个 _Frag。"""
    return _Frag("".join(_wrap(f) for f in frags))


# ---- 装配与插入 ---------------------------------------------------------------


def _omath(inner):
    return f'<m:oMath xmlns:m="{NS_M}" xmlns:a="{NS_A}">{inner}</m:oMath>'


def math_ac(inner, display=False, fallback="[formula]"):
    """生成完整的 mc:AlternateContent XML（可直接被 pptx.oxml.parse_xml 解析）。

    display=True 生成 oMathPara（必须独占段落，行距 1.0，坑 13/15）。
    fallback 为 PowerPoint 不支持 a14 时的占位文本（坑 11：Fallback 里是 a:r/a:t）。
    """
    inner = _concat(inner)
    omath = _omath(inner)
    body = f'<m:oMathPara xmlns:m="{NS_M}" xmlns:a="{NS_A}">{omath}</m:oMathPara>' if display else omath
    fb = (
        f'<mc:Fallback><a:r><a:rPr i="0"/><a:t xml:space="preserve">'
        f"{escape(fallback)}</a:t></a:r></mc:Fallback>"
    )
    return (
        f'<mc:AlternateContent xmlns:mc="{NS_MC}" xmlns:a14="{NS_A14}" xmlns:a="{NS_A}">'
        f'<mc:Choice Requires="a14"><a14:m>{body}</a14:m></mc:Choice>{fb}</mc:AlternateContent>'
    )


def insert_math(paragraph, inner, display=False, fallback="[formula]"):
    """把公式插入 python-pptx 的段落对象末尾。

    display=True 时该段落应只含此公式（独占段落）；纯公式段建议 line_spacing=1.0。
    """
    from pptx.oxml import parse_xml

    paragraph._p.append(parse_xml(math_ac(inner, display=display, fallback=fallback)))


# ---- 自检 ----------------------------------------------------------------


def self_check():
    """构造全部 helper 并做基本合法性检查，返回问题列表。"""
    problems = []
    samples = {
        "T": T("x"),
        "T italic": T("kg", italic=False),
        "F": F("a", "b"),
        "S": S("v", "2"),
        "Q": Q("a", "1"),
        "SQ": SQ("x", "2", "i"),
        "RAD": RAD("2"),
        "LIM": LIM("T", "0"),
        "PAR": PAR("x"),
        "ABS": ABS("x"),
        "CV": CV("x", "y"),
        "MX": MX(["a", "b"], ["c", "d"]),
        "NORM": NORM("v"),
        "M": M(S("E", "2"), T("="), F("1", "2")),
    }
    for name, frag in samples.items():
        xml = math_ac(frag)
        if "&lt;m:" in xml:
            problems.append(f"{name}: XML 被二次转义")
        if "<m:oMath" not in xml:
            problems.append(f"{name}: 缺 oMath")
    # 空基底必须抛错
    for bad in (lambda: S("", "2"), lambda: Q("F", ""), lambda: F("", "2")):
        try:
            bad()
            problems.append("空基底未抛 ValueError（坑 1 回归）")
        except ValueError:
            pass
    return problems


if __name__ == "__main__":
    issues = self_check()
    for p in issues:
        print("FAIL:", p)
    print("self_check:", "PASS" if not issues else f"{len(issues)} problems")
    raise SystemExit(0 if not issues else 1)
