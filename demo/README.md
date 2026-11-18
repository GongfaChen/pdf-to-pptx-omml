# demo

- `smoke_test.py` — 一键环境自检：用 omml_helpers 构建一个 3 页 PPTX（含分数/根号/上下标/列向量/矩阵/定积分/范数/绝对值），跑静态校验，并尝试用 PowerPoint COM 打开验证。
  ```bash
  python smoke_test.py [输出目录]
  ```
  预期输出：`self_check: PASS` → `validate_omml check: PASS (0 issues)` → `PowerPoint COM open: PASS (3 slides)`。

- `smoke_out.pptx / smoke_out.pdf` — smoke test 的产物示例（本地生成，不入库）。
  可用 PowerPoint 打开 pptx 双击任意公式，确认是原生可编辑 OMML 而非图片。

- `test_probe.pdf` — 3 页合成样张（文字层页 / 扫描版页 / OCR 乱码页），用来试跑：
  ```bash
  python ../scripts/probe_pdf.py test_probe.pdf
  ```

> 注意：PowerPoint COM 的 `Open/SaveAs` 需要绝对路径且用反斜杠（正斜杠会报"无法保存"）。
