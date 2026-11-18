#!/usr/bin/env python3
"""smoke_test.py — 环境自检：用 omml_helpers 构建一个 3 页 PPTX 并跑静态校验。

验证内容：helper 全套构造、公式插入、validate_omml 的 pptx 体检全零；
有 PowerPoint COM 时导出 PDF 并断言公式是原生渲染（CambriaMath 字形 / 无占位文本），
无 COM 时明确声明「公式渲染未验收」，不伪报 PASS。

用法：python smoke_test.py [输出目录]
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))

from omml_helpers import (  # noqa: E402
    ABS, CV, EQARR, F, LIM, M, MX, NORM, PAR, PROD, Q, RAD, S, SUM, T,
    insert_math, reset_run_style, self_check, set_run_style,
)
from validate_omml import check_pptx  # noqa: E402


def build(out_path):
    from pptx import Presentation
    from pptx.util import Inches, Pt

    reset_run_style()               # I10：每份 deck 先清全局样式，避免跨 deck 串味
    set_run_style(color="1F3B73")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    def prose(p, text):
        """正文一律走普通 a:r，放在 mc:AlternateContent 之外（坑 5/33）。"""
        r = p.add_run()
        r.text = text
        r.font.name = "Times New Roman"
        r.font.size = Pt(20)

    # slide 1：标题（纯正文，不再把整句塞进 oMath）
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    box = s1.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11.3), Inches(1.6))
    prose(box.text_frame.paragraphs[0], "PDF 教材 → PPT 流水线自检")

    # slide 2：正文 + inline 公式混排；正文在 a:r，公式各自一个 AlternateContent
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    tf2 = s2.shapes.add_textbox(Inches(0.6), Inches(0.8), Inches(12.1), Inches(6.0))
    p1 = tf2.text_frame.paragraphs[0]
    p1.line_spacing = 1.0
    prose(p1, "求根公式：")
    insert_math(
        p1,
        M(Q("x", "1,2"), T(" = "),
          F(M(T("−b ± "), RAD(M(S("b", "2"), T("−4ac")))), T("2"))),
        display=False,
    )
    prose(p1, "，向量 ")
    insert_math(p1, CV("x", "y"), display=False)
    prose(p1, "，范数 ")
    insert_math(p1, NORM("v"), display=False)

    p2 = tf2.text_frame.add_paragraph()
    p2.line_spacing = 1.0
    prose(p2, "定积分 ")
    insert_math(p2, LIM(T("∫"), T("0")), display=False)
    prose(p2, " 与矩阵 ")
    insert_math(p2, MX(["a", "b"], ["c", "d"]), display=False)
    prose(p2, "，绝对值 ")
    insert_math(p2, ABS(S("x", "2")), display=False)

    # display 公式独占段落（坑 13/15），保留 oMathPara 路径覆盖
    p3 = tf2.text_frame.add_paragraph()
    p3.line_spacing = 1.0
    insert_math(p3, M(S("E", "k"), T(" = "), F("1", "2"), S("mv", "2")), display=True)

    # slide 3：正文规格 + I9 新增结构（SUM / PROD / EQARR）
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    tf3 = s3.shapes.add_textbox(Inches(0.6), Inches(0.8), Inches(12.1), Inches(6.0))
    p4 = tf3.text_frame.paragraphs[0]
    p4.line_spacing = 1.4
    prose(p4, "正文规格：Times New Roman 20pt / 行距 1.4；公式段行距 1.0；display 公式独占段落。")

    p5 = tf3.text_frame.add_paragraph()
    p5.line_spacing = 1.0
    prose(p5, "求和 ")
    insert_math(p5, SUM(Q("i", "1"), "n", Q("a", "i")), display=False)
    prose(p5, "，求积 ")
    insert_math(p5, PROD(Q("k", "1"), "m", "k"), display=False)

    p6 = tf3.text_frame.add_paragraph()
    p6.line_spacing = 1.0
    prose(p6, "方程组 ")
    insert_math(
        p6,
        EQARR(M(T("x"), T(" + "), T("y"), T(" = 1")),
              M(T("x"), T(" − "), T("y"), T(" = 0"))),
        display=False,
    )

    prs.save(out_path)
    return out_path


def main():
    issues = self_check()
    print("omml_helpers self_check:", "PASS" if not issues else issues)
    if issues:
        return 1

    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE
    out = build(out_dir / "smoke_out.pptx")
    print("built:", out)

    pptx_issues = check_pptx(out)
    print("validate_omml check:", "PASS (0 issues)" if not pptx_issues else pptx_issues)
    if pptx_issues:
        return 1

    # 闸 2：有 PowerPoint COM 时导出 PDF 并核验公式是原生渲染（不是 mc:Fallback 占位）。
    # 坑 10 / 坑 33：XML 合法但公式被静默丢弃、或整页只剩占位文本，静态校验都查不出来。
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("PowerPoint COM: SKIP —— 公式渲染未验收（无 pywin32，闸 2 未执行）")
        return 0

    pdf = out.with_suffix(".pdf")
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
        except Exception as e:
            print(f"PowerPoint COM: SKIP —— 公式渲染未验收（本机无 PowerPoint COM：{e}）")
            return 0
        app.DisplayAlerts = 1  # ppAlertsNone，避免弹窗卡住无人值守自检
        try:
            pres = app.Presentations.Open(str(out.resolve()), WithWindow=False, ReadOnly=True)
            n = pres.Slides.Count
            pres.SaveAs(str(pdf.resolve()), 32)  # 32 = ppSaveAsPDF，公式原生渲染
            pres.Close()
            app.Quit()
            print(f"PowerPoint COM open: PASS ({n} slides)")
        except Exception as e:
            print("PowerPoint COM: FAIL —", e)
            try:
                app.Quit()
            except Exception:
                pass
            return 1
    finally:
        pythoncom.CoUninitialize()

    # 字形级断言：确认公式是原生 OMML，不是 [formula] 占位
    try:
        import pdfplumber
    except ImportError:
        print("公式渲染: 未验收 —— 缺 pdfplumber，无法检查导出 PDF 的字形")
        return 0

    cambria = 0
    placeholders = 0
    per_page = []
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            page_cambria = 0
            for ch in page.chars:
                if "cambriamath" in (ch.get("fontname") or "").replace(" ", "").lower():
                    page_cambria += 1
            cambria += page_cambria
            per_page.append(page_cambria)
            text = page.extract_text() or ""
            placeholders += text.count("[formula]") + text.count("□")
    detail = " / ".join(f"p{i}={c}" for i, c in enumerate(per_page, 1))
    if cambria == 0 or placeholders:
        print(f"公式渲染: FAIL —— CambriaMath 字形 {cambria} 处（{detail}）、占位文本命中 {placeholders} 处")
        return 1
    print(f"公式渲染: PASS —— CambriaMath 字形 {cambria} 处（{detail}）、无占位文本")
    return 0


if __name__ == "__main__":
    sys.exit(main())
