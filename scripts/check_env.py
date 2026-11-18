#!/usr/bin/env python3
"""check_env.py — 前置 A · 环境闸：开工前确认依赖矩阵与 COM 闸能不能跑（只读）。

只读，不打开 PowerPoint、不导出、不修改任何文件。输出：依赖矩阵、缺失项安装命令、
PowerPoint/WPS ProgID 是否注册，以及「有/无 COM」时闸 2（公式原生渲染核验）的结论。
命名「前置 A」是为了不与 SKILL.md 已有的「第 0 步 · probe 诊断」撞号。

用法：
    python scripts/check_env.py
退出码恒为 0（环境闸是诊断，不因缺依赖报错）。
"""
import importlib
import sys
import warnings
from pathlib import Path

# (import 名, 显示名, pip 安装名, 用途)
DEPS = [
    ("fitz", "PyMuPDF (fitz)", "pymupdf", "PDF 解析 / 裁图 / 版式审计"),
    ("pptx", "python-pptx", "python-pptx", "构建 PPTX"),
    ("lxml", "lxml", "lxml", "OMML 静态校验"),
    ("PIL", "Pillow", "pillow", "拼 contact sheet / 白边 trim"),
    ("pdfplumber", "pdfplumber", "pdfplumber", "导出 PDF 字形核验"),
    ("win32com", "pywin32 (win32com)", "pywin32", "PowerPoint COM 闸 2"),
    ("pythoncom", "pywin32 (pythoncom)", "pywin32", "COM 初始化"),
]


def probe(mod):
    try:
        # 探测导入时屏蔽依赖自身的告警（如 fitz 的弃用 FutureWarning），保持闸输出干净。
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            m = importlib.import_module(mod)
        return True, getattr(m, "__version__", "")
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def progids():
    """注册表只读查询，返回 {ProgID: CLSID 或 None}；非 Windows 返回全 None。"""
    result = {"PowerPoint.Application": None, "KWPP.Application": None}
    try:
        import winreg
    except ImportError:
        return result
    for progid in result:
        try:
            k = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid)
            result[progid] = winreg.QueryValue(k, "CLSID")
            winreg.CloseKey(k)
        except OSError:
            pass
    return result


def main():
    is_windows = sys.platform.startswith("win")
    print(f"# 前置 A · 环境闸：{Path(sys.executable).name}")
    print()
    print(f"- Python：`{sys.executable}`")
    print(f"- 版本：{sys.version.split()[0]}｜平台：{sys.platform}")
    print()
    print("## 依赖矩阵")
    print()
    print("| 依赖 | 状态 | 版本/错误 | 用途 |")
    print("|---|---|---|---|")
    results = {}
    missing_pip = []
    for mod, label, pip_name, use in DEPS:
        ok, ver = probe(mod)
        results[mod] = ok
        print(f"| {label} | {'OK' if ok else 'MISSING'} | {ver or '—'} | {use} |")
        if not ok and pip_name not in missing_pip:
            if pip_name == "pywin32" and not is_windows:
                continue  # pywin32 仅 Windows 提供，非 Windows 不列入安装命令
            missing_pip.append(pip_name)

    print()
    print("## PowerPoint / WPS COM ProgID（注册表只读）")
    print()
    pids = progids()
    if not is_windows:
        print("- 非 Windows 平台：无 COM 注册表，pywin32 亦不可用。")
        com_ok = False
    else:
        for progid, clsid in pids.items():
            print(f"- {progid}：{'已注册 ' + clsid if clsid else '未注册'}")
        pywin_ok = results.get("win32com") and results.get("pythoncom")
        com_ok = bool(pywin_ok and any(pids.values()))

    print()
    print("## 闸 2（公式原生渲染核验）")
    print()
    if com_ok:
        print("- 可执行：pywin32 + PowerPoint/WPS COM 均可用。")
        print("- 交付前必须跑 pres.SaveAs(pdf, 32) 并核验 CambriaMath 字形 / 无 [formula] 占位。")
    else:
        print("- 不可执行：缺少 pywin32 或本机无 PowerPoint/WPS COM。")
        print("- 此时**公式渲染未验收**，不得据此宣布通过或交付；LibreOffice 等只认 mc:Fallback "
              "的渲染器导出的 PDF 禁止作为交付物。")

    print()
    print("## 缺失项安装命令")
    print()
    if missing_pip:
        print("```bash")
        print(f"{sys.executable} -m pip install " + " ".join(sorted(set(missing_pip))))
        print("```")
    else:
        print("- 依赖齐全，无需安装。")
    if not is_windows:
        print("- 注：pywin32 仅 Windows 可用，非 Windows 无法补齐，闸 2 只能 SKIP 并声明「未验收」。")

    print()
    print("> 本闸只读，不打开 PowerPoint、不导出文件；命名「前置 A」以免与「第 0 步 · probe 诊断」撞号。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
