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
import re
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


def reset_run_style():
    """恢复默认 run 样式（清掉模块级全局状态）。

    set_run_style 的 size/color/font 传 None 表示「不修改」而非「清空」——一旦设过
    颜色/字号就没法用 None 清回来；_STYLE 又是模块级全局，连续构建多份 deck 会串味。
    每份 deck 构建前（或构建完）调一次 reset_run_style() 即可回到出厂默认。
    """
    _STYLE.update({"sz": None, "color": None, "font": "Cambria Math"})


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
    """已构造的 OMML XML 片段。isinstance 检查用于区分「裸文本」与「XML」。

    坑 45：str 子类相加默认走 str.__add__，结果退化成普通 str，下游 _wrap 会把它
    当裸文本用 T() 包裹并转义（F(T('a') + T('b'), '2') 的 m:t 里出现 &lt;m:r&gt;）。
    重载 __add__/__radd__，让任意方向的拼接结果都仍是 _Frag。
    """

    def __add__(self, other):
        return _Frag(str(self) + str(other))

    def __radd__(self, other):
        return _Frag(str(other) + str(self))


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


def _nary(chr_val, sub, sup, body):
    """m:nary 公共构造（求和/求积）。坑 12：sub/sup 必须同时给，且不写 subHide/supHide。"""
    pr = f'<m:naryPr><m:chr m:val="{chr_val}"/><m:limLoc m:val="undOvr"/></m:naryPr>'
    return _Frag(
        f"<m:nary>{pr}<m:sub>{_wrap(sub)}</m:sub><m:sup>{_wrap(sup)}</m:sup>"
        f"<m:e>{_wrap(body)}</m:e></m:nary>"
    )


def SUM(sub, sup, body):
    """求和 ∑_{sub}^{sup} body（m:nary）。用法：SUM(Q('i','1'), 'n', Q('a','i'))。"""
    return _nary("∑", sub, sup, body)


def PROD(sub, sup, body):
    """求积 ∏_{sub}^{sup} body（m:nary）。"""
    return _nary("∏", sub, sup, body)


def EQARR(*rows):
    """左对齐方程组 m:eqArr（裸结构，不套 m:d；列向量 CV 才需要 m:d，坑 34）。

    每个 row 是一行方程，可传单个 _Frag 或 list/tuple 拼接。
    用法：EQARR(M(T('x'),T('+'),T('y'),T('=1')), M(T('x'),T('−'),T('y'),T('=0')))。
    """
    cells = "".join(f"<m:e>{_concat(r)}</m:e>" for r in rows)
    return _Frag(f"<m:eqArr>{cells}</m:eqArr>")


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


def _linear_fallback(inner):
    """结构感知的 OMML → 线性文本（坑 33 / R2）。

    按 m:f / m:sSup / m:sSub / m:sSubSup / m:rad / m:d / m:limLow / m:eqArr / m:m
    还原成 a/b、a^b、a_b、sqrt(…)、(…) 等可读线性式，而不是把标签一去了之——
    后者会把 b² 压成 b2、½ 压成 12，反而误导。显式 fallback= 仍优先（见 math_ac）。
    解析用标准库 ElementTree：属性节点无文本子节点，join 后自然为空，文本只在 m:t 里。
    """
    import xml.etree.ElementTree as ET

    mval_attr = "{" + NS_M + "}val"

    def local(tag):
        return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""

    def child(el, name):
        for c in el:
            if local(c.tag) == name:
                return c
        return None

    def mval(el):
        return (el.get(mval_attr) or "") if el is not None else ""

    def grp(s):
        """子式含运算符/分隔符时括起来，避免 (a+b)/2 被读成 a+b/2。"""
        s = (s or "").strip()
        if not s:
            return "()"
        if s.startswith("(") and s.endswith(")"):
            return s
        if any(ch in s for ch in "+-−*/^_=,; "):
            return f"({s})"
        return s

    def lin_child(el, name):
        c = child(el, name)
        return lin(c) if c is not None else ""

    def lin(el):
        tag = local(el.tag)
        if tag == "t":
            return el.text or ""
        if tag == "f":
            return f"{grp(lin_child(el, 'num'))}/{grp(lin_child(el, 'den'))}"
        if tag == "sSup":
            return f"{grp(lin_child(el, 'e'))}^{grp(lin_child(el, 'sup'))}"
        if tag == "sSub":
            return f"{grp(lin_child(el, 'e'))}_{grp(lin_child(el, 'sub'))}"
        if tag == "sSubSup":
            return (f"{grp(lin_child(el, 'e'))}_{{{lin_child(el, 'sub')}}}"
                    f"^{{{lin_child(el, 'sup')}}}")
        if tag == "rad":
            deg, expr = lin_child(el, "deg"), lin_child(el, "e")
            return f"root({deg}, {expr})" if deg.strip() else f"sqrt({expr})"
        if tag == "d":
            dpr = child(el, "dPr")
            beg, end = "(", ")"
            if dpr is not None:
                b, e = child(dpr, "begChr"), child(dpr, "endChr")
                beg = mval(b) if b is not None else "("
                end = mval(e) if e is not None else ")"
            return f"{beg}{''.join(lin(c) for c in el)}{end}"
        if tag == "limLow":
            return f"{grp(lin_child(el, 'e'))}_{{{lin_child(el, 'lim')}}}"
        if tag == "eqArr":
            rows = [lin(c) for c in el if local(c.tag) == "e"]
            return " ; ".join(rows) if rows else "".join(lin(c) for c in el)
        if tag == "m":
            rows = []
            for mr in el:
                if local(mr.tag) == "mr":
                    rows.append(", ".join(lin(c) for c in mr if local(c.tag) == "e"))
            return " ; ".join(rows) if rows else "".join(lin(c) for c in el)
        if tag == "nary":
            pr = child(el, "naryPr")
            sym = mval(child(pr, "chr")) if pr is not None else ""
            sym = sym or "∑"
            return (f"{sym}_{{{lin_child(el, 'sub')}}}^{{{lin_child(el, 'sup')}}}"
                    f"({lin_child(el, 'e')})")
        # 其余容器（oMath/oMathPara/num/den/e/sub/sup/lim/deg/mr/r/rPr/…）：拼接子节点。
        return "".join(lin(c) for c in el)

    wrapped = f'<root xmlns:m="{NS_M}" xmlns:a="{NS_A}">{inner}</root>'
    try:
        root = ET.fromstring(wrapped)
    except ET.ParseError:
        # 兜底：片段不成 XML 时退回纯文本，至少不留 [formula] 占位。
        text = re.sub(r"<[^>]*>", "", str(inner))
        return " ".join(html.unescape(text).split()) or "公式"
    out = re.sub(r"\s+", " ", "".join(lin(c) for c in root)).strip()
    return out or "公式"


def _omath(inner):
    return f'<m:oMath xmlns:m="{NS_M}" xmlns:a="{NS_A}">{inner}</m:oMath>'


def math_ac(inner, display=False, fallback=None):
    """生成完整的 mc:AlternateContent XML（可直接被 pptx.oxml.parse_xml 解析）。

    display=True 生成 oMathPara（必须独占段落，行距 1.0，坑 13/15）。
    fallback 为 mc:Fallback 里的线性文本（坑 11：Fallback 里是 a:r/a:t）。
    默认 None = 由片段自动生成可读线性文本，不再使用 [formula] 占位（坑 33）；
    显式传字符串可覆盖，便于给出更精确的线性公式。
    """
    inner = _concat(inner)
    if fallback is None:
        fallback = _linear_fallback(inner)
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


def insert_math(paragraph, inner, display=False, fallback=None):
    """把公式插入 python-pptx 的段落对象末尾。

    display=True 时该段落应只含此公式（独占段落）；纯公式段建议 line_spacing=1.0。
    插到 a:endParaRPr 之前——它必须是 a:p 的最后一个子元素，直接 append() 会把它
    挤到公式后面，PowerPoint 编辑过的段落常带这个元素（I8）。
    """
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import qn

    ac = parse_xml(math_ac(inner, display=display, fallback=fallback))
    end_rpr = paragraph._p.find(qn("a:endParaRPr"))
    if end_rpr is not None:
        end_rpr.addprevious(ac)
    else:
        paragraph._p.append(ac)


# ---- 自检 ----------------------------------------------------------------


def self_check():
    """构造全部 helper 并做基本合法性检查，返回问题列表。"""
    problems = []

    # R1 回归：先造两个片段再拼（源码里不写 T(...) + T(...)，避免触发新增 lint 规则），
    # 拼接结果必须仍是 _Frag，且喂给下一个 helper 时不能被转义（坑 45）。
    frag_a, frag_b = T("a"), T("b")
    concat = frag_a + frag_b
    if not isinstance(concat, _Frag):
        problems.append("_Frag + _Frag 退化成普通 str（R1 回归，坑 45）")
    if "&lt;m:" in math_ac(F(concat, "2")):
        problems.append("_Frag + _Frag 拼接被转义（R1 回归，坑 45）")

    samples = {
        "T": T("x"),
        "T italic": T("kg", italic=False),
        "F": F("a", "b"),
        "S": S("v", "2"),
        "Q": Q("a", "1"),
        "SQ": SQ("x", "2", "i"),
        "RAD": RAD("2"),
        "LIM": LIM("T", "0"),
        "SUM": SUM(Q("i", "1"), "n", Q("a", "i")),
        "PROD": PROD(Q("k", "1"), "m", "k"),
        "EQARR": EQARR(M(T("x"), T("+"), T("y"), T("=1")), M(T("x"), T("−"), T("y"), T("=0"))),
        "PAR": PAR("x"),
        "ABS": ABS("x"),
        "CV": CV("x", "y"),
        "MX": MX(["a", "b"], ["c", "d"]),
        "NORM": NORM("v"),
        "M": M(S("E", "2"), T("="), F("1", "2")),
        "CONCAT": F(concat, "2"),
    }
    for name, frag in samples.items():
        xml = math_ac(frag)
        if "&lt;m:" in xml:
            problems.append(f"{name}: XML 被二次转义")
        if "<m:oMath" not in xml:
            problems.append(f"{name}: 缺 oMath")

    # R2 回归：结构感知线性化，关键结构不能塌成裸串；Fallback 不留 &lt; / [formula]。
    linear_cases = {
        "F": (F("1", "2"), "1/2"),
        "S": (S("b", "2"), "b^2"),
        "Q": (Q("x", "1"), "x_1"),
        "RAD": (RAD(S("b", "2")), "sqrt("),
        "D": (PAR("x"), "(x)"),
        "ABS": (ABS("x"), "|x|"),
        "EQARR": (EQARR("x=1", "y=2"), "x=1 ; y=2"),
        "NARY": (SUM("i=1", "n", "a_i"), "∑_"),
    }
    for name, (frag, want) in linear_cases.items():
        got = _linear_fallback(frag)
        if want not in got:
            problems.append(f"{name} 线性化异常：期望含 {want!r}，得到 {got!r}")
    # 显式 fallback= 必须优先
    forced = math_ac(F("1", "2"), fallback="one half")
    if "one half" not in forced or "1/2" in forced:
        problems.append("显式 fallback= 未优先（R2）")

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
