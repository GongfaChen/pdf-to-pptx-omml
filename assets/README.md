# assets/

README 效果演示用的截图/图片。

## 版权与授权边界

- `before_matl_p60.png` —— 国内《材料力学》教材"扭转"章节原页截图（p.60，扫描版），
  **版权归原出版社所有**，此处仅作"输入长什么样"的效果展示。
- `after_matl_slide_3_3.png` —— 用本工具把同一节提取成的课件截图（公式为原生 OMML）。
- `before_textbook_p200.jpg` —— AoPS *Precalculus*（第 6 章 p.200）教材页截图，
  **版权归 Art of Problem Solving (AoPS) 所有**，此处仅作"输入长什么样"的效果展示。
- `after_slide_magnitude.png` / `after_slide_solution_6_12.jpg` —— 用本工具把同一节内容
  提取成的课件截图，仅作效果展示。
- `demo_workflow.gif` —— README 顶部流程演示的**静态兜底版**（960×540，35 帧，约 1.7 MB）。当前
  README 用的是 user-attachments 上的 H.264 视频（807 KB，27 秒）；PyPI / VS Code 等不渲染
  `<video>` 的渲染器回落到这张 GIF。两者都由真实录屏合成：画面含上述《材料力学》p.60 扫描页
  与工具产出的课件页，**版权边界同前两条**，不在 MIT 授权范围内。生成脚本在 `_gif_work/`
  （已 gitignore，不入库）。

以上图片**不在本仓库 MIT 授权范围内**，不得再分发或用于商业用途；如需在自己的展示中
使用教材截图，请自备已获授权的样本。工具本身只处理用户自有或已获授权的文件。

## 替换建议

正式发布前建议把教材页替换为**自制样张**（自己排一页公式+图形的 PDF），
既避免版权风险，也能保证截图长期可用（教材改版后页码/版式会变）。
