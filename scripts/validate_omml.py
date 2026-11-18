#!/usr/bin/env python3
"""validate_omml.py — OMML 静态校验器（两种模式）。

模式 1 · pptx 体检：解包 .pptx，逐 slide 检查所有 m:oMath：
  - m:t 内含 '<' / 'xmlns' → 公式被误转义为纯文本（坑 4/36）
  - m:e/m:sup/m:sub/m:num/m:den 为空 → PowerPoint 严格校验拒开（坑 1）
  - 结构元素下出现裸文本（坑 3/6）
  - 存在 m:acc → PDF 导出会渲染成亚美尼亚标点（坑 35，警告）
  - 存在 supHide 的 m:nary → 整个文本框被丢弃（坑 12，警告）

模式 2 · 源码 lint：扫描 .py 源文件中的高频自伤写法：
  - T('')、T(''),（空基底，坑 1/43）
  - T(' ' + F(…) 等 —— T() 拼接 helper 输出（坑 36）
  - T() 内含 '<m:'（XML 当文本，坑 4）
  - insert_math 传入裸字符串（坑 6）

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
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def mtag(tag):
    return f"{{{NS_M}}}{tag}"


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
            if not omaths:
                continue
            for om in omaths:
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
            for om in root.iter(mtag("oMathPara")):
                # 坑 13：oMathPara 与普通文字 run 混排会被压成 inline 样式
                parent = om.getparent()
                if parent is not None and parent.tag.endswith("}p"):
                    for sib in parent:
                        if sib.tag.endswith("}r") and len(sib):
                            issues.append(f"{sl}: oMathPara 与普通文字 run 同段（公式会被压成 inline 小字号）")
                            break
            for el in root.iter(mtag("acc")):
                issues.append(f"{sl}: 存在 m:acc（PDF 导出渲染成亚美尼亚标点，坑 35）")
            for el in root.iter(mtag("nary")):
                pr = el.find(mtag("naryPr"))
                if pr is not None and pr.find(mtag("supHide")) is not None:
                    issues.append(f"{sl}: m:nary 带 supHide（整个文本框会被丢弃，坑 12）")
    return issues


LINT_RULES = [
    (re.compile(r"T\(\s*['\"][,)]"), "T('') 空基底/空文本（坑 1/43）"),
    (re.compile(r"T\(\s*['\"][^'\"]*['\"]\s*\+"), "T() 内做字符串拼接（会把 helper XML 当文本，坑 36）"),
    (re.compile(r"T\(\s*['\"][^'\"]*<m:"), "T() 内含 OMML XML（坑 4）"),
    (re.compile(r"insert_math\(\s*\w+\s*,\s*['\"]"), "insert_math 传裸字符串（坑 6，应 T() 包裹或用 helper）"),
    (re.compile(r"m:acc"), "使用 m:acc（坑 35）"),
    (re.compile(r"supHide"), "使用 supHide（坑 12）"),
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
    if len(sys.argv) < 3 or sys.argv[1] not in ("check", "lint"):
        print(__doc__)
        return 2
    mode, targets = sys.argv[1], sys.argv[2:]
    issues = check_pptx(targets[0]) if mode == "check" else lint_sources(targets)
    for i in issues:
        print("FAIL:", i)
    n = len(issues)
    print(f"{mode}: {'PASS' if n == 0 else f'{n} issue(s)'}")
    return 0 if n == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
