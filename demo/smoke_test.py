#!/usr/bin/env python3
"""smoke_test.py — 环境自检：用 omml_helpers 构建一个 3 页 PPTX 并跑静态校验。

验证内容：helper 全套构造、公式插入、validate_omml 的 pptx 体检全零、PowerPoint 能打开
（有 COM 时自动验证，没有则跳过）。

用法：python smoke_test.py [输出目录]
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))

from omml_helpers import (  # noqa: E402
    ABS, CV, F, LIM, M, MX, NORM, PAR, Q, RAD, S, T, insert_math, self_check, set_run_style,
)
from validate_omml import check_pptx  # noqa: E402


def build(out_path):
    from pptx import Presentation
    from pptx.util import Inches

    set_run_style(color="1F3B73")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    # slide 1：标题
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    box = s1.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11.3), Inches(1.6))
    p = box.text_frame.paragraphs[0]
    insert_math(p, M(T("PDF 教材 → PPT 流水线自检", italic=False)), display=True)

    # slide 2：混合公式样例
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    tf2 = s2.shapes.add_textbox(Inches(0.6), Inches(0.8), Inches(12.1), Inches(6.0))
    p1 = tf2.text_frame.paragraphs[0]
    p1.line_spacing = 1.0
    insert_math(
        p1,
        M(T("求根公式：", italic=False), Q("x", "1,2"), T(" = "),
          F(M(T("−b ± "), RAD(M(S("b", "2"), T("−4ac")))), T("2")),
          T("，向量 "), CV("x", "y"), T("，范数 "), NORM("v")),
        display=True,
    )
    p2 = tf2.text_frame.add_paragraph()
    p2.line_spacing = 1.0
    insert_math(
        p2,
        M(T("定积分 "), LIM(T("∫"), T("0")), T(" 与矩阵 "), MX(["a", "b"], ["c", "d"]),
          T("，绝对值 "), ABS(S("x", "2"))),
        display=True,
    )

    # slide 3：正文规格样例（Times 20pt / 1.4）
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    tf3 = s3.shapes.add_textbox(Inches(0.6), Inches(0.8), Inches(12.1), Inches(6.0))
    p3 = tf3.text_frame.paragraphs[0]
    p3.line_spacing = 1.4
    run = p3.add_run()
    run.text = "正文规格：Times New Roman 20pt / 行距 1.4；公式段行距 1.0；display 公式独占段落。"
    run.font.name = "Times New Roman"
    run.font.size = None  # 由主题/模板接管，此处仅示意

    prs.save(out_path)
    return out_path


def main():
    issues = self_check()
    print("omml_helpers self_check:", "PASS" if not issues else issues)

    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE
    out = build(out_dir / "smoke_out.pptx")
    print("built:", out)

    pptx_issues = check_pptx(out)
    print("validate_omml check:", "PASS (0 issues)" if not pptx_issues else pptx_issues)

    # 有 PowerPoint COM 时验证能否打开（坑：XML 合法但 PowerPoint 拒开）
    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(str(out.resolve()), WithWindow=False, ReadOnly=True)
        n = pres.Slides.Count
        pres.Close()
        app.Quit()
        print(f"PowerPoint COM open: PASS ({n} slides)")
    except ImportError:
        print("PowerPoint COM open: SKIP（未安装 pywin32）")
    except Exception as e:
        print("PowerPoint COM open: FAIL —", e)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
