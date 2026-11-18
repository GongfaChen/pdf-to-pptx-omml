#!/usr/bin/env python3
"""validate_omml.py — OMML 静态校验器（两种模式）。

模式 1 · pptx 体检：解包 .pptx，逐 slide 检查所有 m:oMath：
  - m:t 内含 '<' / 'xmlns' → 公式被误转义为纯文本（坑 4/36）
  - m:e/m:sup/m:sub/m:num/m:den 为空 → PowerPoint 严格校验拒开（坑 1）
  - 结构元素下出现裸文本（坑 3/6）
  - 存在 m:acc → PDF 导出会渲染成亚美尼亚标点（坑 35，警告）
  - 存在 supHide 的 m:nary → 整个文本框被丢弃（坑 12，警告）
  - oMath/oMathPara 缺 a14:m 祖先、含 oMath 的 mc:Choice 其 Requires 未含 a14
    → 整个 AlternateContent 被 PowerPoint 静默丢弃（坑 10）
  - oMathPara 上溯到 a:p 后与普通 a:r 同段 → display 公式被压成 inline（坑 13）
  - mc:Fallback 仍是 [formula] 占位 → 只认 Fallback 的渲染器吞掉正文（坑 33）

模式 2 · 源码 lint：扫描 .py 源文件中的高频自伤写法：
  - T('')、T(''),（空基底，坑 1/43）
  - T(' ' + F(…) 等 —— T() 拼接 helper 输出（坑 36）
  - T() 内含 '<m:'（XML 当文本，坑 4）
  - insert_math 传入裸字符串（坑 6）
  - helper 输出用 + 直接拼接（T('a') + T('b')，坑 45/R1；建议 M(...)）

用法：
    python validate_omml.py check <文件.pptx>            # 模式 1
    python validate_omml.py lint <src1.py> [src2.py ...]  # 模式 2
退出码：0 = 通过；1 = 有问题。
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def mtag(tag):
    return f"{{{NS_M}}}{tag}"


def atag(tag):
    return f"{{{NS_A}}}{tag}"


def a14tag(tag):
    return f"{{{NS_A14}}}{tag}"


# 禁用词拆开写：lint 扫描 validate_omml.py 自身时，规则表/校验代码里若出现完整字面词
# 会命中自己的规则（自指假阳性），因此用常量拼出，源文件里不留连续字面词。
_BAN_ACC = "m:" + "acc"
_BAN_SUP = "sup" + "Hide"


def check_pptx(path):
    issues = []
    with zipfile.ZipFile(path) as z:
        slides = sorted(
            n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)
        )
        if not slides:
            issues.append("文件中没有 slide（可能不是有效的 pptx）")
        for name in slides:
            root = etree.fromstring(z.read(name))
            sl = Path(name).name
            omaths = list(root.iter(mtag("oMath")))
            for om in omaths:
                # 坑 10（I2）：pptx 数学必须是 mc:Choice(Requires=a14) > a14:m > oMath(Para)。
                # 只遍历 m:oMath 抓不到「整个 AlternateContent 被 PowerPoint 静默丢弃」的形态。
                if not any(a.tag == a14tag("m") for a in om.iterancestors()):
                    issues.append(f"{sl}: m:oMath 缺 a14:m 祖先（坑 10，公式会被静默丢弃）")
                for t in om.iter(mtag("t")):
                    txt = t.text or ""
                    if "<" in txt or "xmlns" in txt or "&lt;" in txt:
                        issues.append(f"{sl}: m:t 含转义 XML（公式显示为代码）: {txt[:60]!r}")
                for tag in ("e", "sup", "sub", "num", "den"):
                    for el in om.iter(mtag(tag)):
                        if len(el) == 0 and not (el.text or "").strip():
                            issues.append(f"{sl}: 空 <m:{tag}> —— PowerPoint 将拒开文件")
                for el in om.iter():
                    if el.tag in (mtag("t"), mtag("r")):
                        continue
                    if (el.text or "").strip():
                        issues.append(f"{sl}: {etree.QName(el).localname} 下有裸文本: {el.text.strip()[:40]!r}")
            for op in root.iter(mtag("oMathPara")):
                if not any(a.tag == a14tag("m") for a in op.iterancestors()):
                    issues.append(f"{sl}: m:oMathPara 缺 a14:m 祖先（坑 10）")
                # 坑 13：display 公式必须独占段落。oMathPara 的父元素是 a14:m（不是 a:p），
                # 要顺着 a14:m / mc:Choice 上溯到 a:p 再判 a:r 兄弟，否则这段检查永远不触发。
                para = next((a for a in op.iterancestors() if a.tag == atag("p")), None)
                if para is not None:
                    for sib in para:
                        if sib.tag == atag("r") and len(sib):
                            issues.append(f"{sl}: oMathPara 与普通文字 run 同段（公式会被压成 inline 小字号）")
                            break
            for ch in root.iter(f"{{{NS_MC}}}Choice"):
                if (next(ch.iter(mtag("oMath")), None) is not None
                        and "a14" not in (ch.get("Requires") or "")):
                    issues.append(f"{sl}: 含 oMath 的 mc:Choice 其 Requires 未含 a14（坑 10）")
            # 坑 33（I1）：Fallback 仍是 [formula] 占位 → 只认 Fallback 的渲染器吞掉正文
            for fb in root.iter(f"{{{NS_MC}}}Fallback"):
                for t in fb.iter(atag("t")):
                    if "[formula]" in (t.text or ""):
                        issues.append(f"{sl}: mc:Fallback 仍是 [formula] 占位（坑 33，应改成可读线性文本）")
            for el in root.iter(mtag("acc")):
                issues.append(f"{sl}: 存在 {_BAN_ACC}（PDF 导出渲染成亚美尼亚标点，坑 35）")
            for el in root.iter(mtag("nary")):
                pr = el.find(mtag("naryPr"))
                if pr is not None and pr.find(mtag(_BAN_SUP)) is not None:
                    issues.append(f"{sl}: m:nary 带 {_BAN_SUP}（整个文本框会被丢弃，坑 12）")
    return issues


# R1 回归规则：helper 输出用 + 直接拼接（T('a') + T('b')）。
# 与「T(' ' + …)」那条区分开：本规则要求 + 左边是「)」（helper 调用收尾）、右边是 helper 调用，
# 而 T(' ' + F(…)) 里 + 左边是字符串引号、右边也是 helper 但左边无 ')', 故只命中旧规则一次。
_HELPER_NAMES = r"(?:T|F|S|Q|SQ|RAD|LIM|PAR|ABS|CV|MX|NORM|M|SUM|PROD|EQARR)"

LINT_RULES = [
    (re.compile(r"T\(\s*['\"][,)]"), "T('') 空基底/空文本（坑 1/43）"),
    (re.compile(r"T\(\s*['\"][^'\"]*['\"]\s*\+"), "T() 内做字符串拼接（会把 helper XML 当文本，坑 36）"),
    (re.compile(r"T\(\s*['\"][^'\"]*<m:"), "T() 内含 OMML XML（坑 4）"),
    (re.compile(r"insert_math\(\s*\w+\s*,\s*['\"]"), "insert_math 传裸字符串（坑 6，应 T() 包裹或用 helper）"),
    (re.compile(r"\)\s*\+\s*" + _HELPER_NAMES + r"\s*\("),
     "helper 输出用 + 直接拼接（坑 45/R1，曾退化为裸文本被转义；建议 M(...)）"),
    (re.compile(_BAN_ACC), f"使用 {_BAN_ACC}（坑 35）"),
    (re.compile(_BAN_SUP), f"使用 {_BAN_SUP}（坑 12）"),
]


DOCSTRING_RX = re.compile(r'""".*?"""|\'\'\'.*?\'\'\'', re.DOTALL)


def lint_sources(paths):
    issues = []
    for p in paths:
        src = Path(p).read_text(encoding="utf-8", errors="replace")
        src = DOCSTRING_RX.sub('""', src)  # 剥离 docstring（其中会提及 m:acc/supHide 等禁用词）
        for i, line in enumerate(src.splitlines(), 1):
            code = line.split("#")[0]  # 剥离行注释
            for rx, msg in LINT_RULES:
                if rx.search(code):
                    issues.append(f"{p}:{i}: {msg} → {line.strip()[:70]}")
    return issues


def main():
    # 先解析参数，不触发 license 闸门（P2-1）：--help 永远可用，且应返回 0
    if sys.argv[1:] and sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0
    if len(sys.argv) < 3 or sys.argv[1] not in ("check", "lint"):
        print(__doc__)
        return 2

    from license_check import require_license_or_exit
    require_license_or_exit("core")  # PAID_FEATURES 现在为空 → 免费放行

    mode, targets = sys.argv[1], sys.argv[2:]
    issues = check_pptx(targets[0]) if mode == "check" else lint_sources(targets)
    for i in issues:
        print("FAIL:", i)
    n = len(issues)
    print(f"{mode}: {'PASS' if n == 0 else f'{n} issue(s)'}")
    return 0 if n == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
