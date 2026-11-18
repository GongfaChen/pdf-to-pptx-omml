# 45 条实战坑：PDF 教材 → OMML 公式 PPT 流水线

> 每一条都真实导致过 PowerPoint 拒开、公式消失、渲染错误或交付返工。
> 编号与原文一致，便于溯源讨论。适用环境：Windows + python-pptx + PyMuPDF + PowerPoint COM。

## A. OMML 构造类（最常见的自伤）

1. **空基底**：`S(T(''), T('2'))` 产生空 `m:e`，PowerPoint 严格校验拒绝打开。应写 `S(T('v'), T('2'))`。
2. **a14 命名空间**：`mc:Choice Requires="a14"` 的前缀必须在 `mc:AlternateContent` 元素上显式声明 `xmlns:a14`（或提升到 slide 根元素 + `mc:Ignorable="a14"`），否则 prefix 无法解析。
3. **裸文本下标**：`m:sub` 内必须包 `m:r`，直接塞文本是无效 OMML（最隐蔽——XML 合法但 PowerPoint 打不开）。同理：希腊字母常量拼进公式必须用文本 run 包裹。
4. **helper 误包裹**：把 helper 的 XML 输出再包一层文本 run，会整体转义成纯文本，公式显示为原始 XML。检查方法：oMath tostring 后搜 `<`。
5. **整句塞进 inline oMath 不换行**：整段成为单个数学区域后渲染不换行、溢出文本框并在右缘产生碎片。应拆成「文本 run + 短数学 + 文本 run」三段。
6. **裸字符串传给 insert**：直接嵌裸文本（坑 3 的入口形态）。用静态校验器扫全部 oMath：`<` 检测转义、结构元素直接 text 检测裸文本、空元素检测。
7. **二分定位坏页**：按 slide 前缀重建子文件逐个试开，快速锁定问题页。
36. **T() 内拼接/包裹 helper 输出**：`T(' ' + F('3','5'))`、`T(F('3','5'))` 会把 XML 转义成字面文本，公式显示为代码（坑 4 的同源形态）。正确写法 `M(T(' '), F('3','5'), T(' '))` 或直接传 `_Frag`；`validate_omml.py lint` 会扫这种写法。
45. **_Frag 相加退化成 str → 下游被转义 → 公式显示成 XML**：`_Frag` 是 `str` 子类，`T('a') + T('b')` 默认走 `str.__add__`，结果不再是 `_Frag`；喂给下一个 helper 时 `_wrap` 把它当裸文本用 `T()` 包裹并转义（`F(T('a') + T('b'), '2')` 的 `m:t` 里出现 `&lt;m:r&gt;`），公式显示成原始 XML 字面文本（坑 4/36 的同源形态）。**修法**：库已重载 `_Frag.__add__/__radd__`，任意方向拼接结果都仍是 `_Frag`；但源码里拼接 helper 片段**一律用 `M(...)`、不要用 `+`**——`validate_omml.py lint` 会抓 `+` 拼接（R1 回归，跨版本也更稳）。

## B. XML / 命名空间类（PowerPoint 静默丢内容）

10. **pptx 数学必须有 `a14:m` 包裹**：结构必须是 `mc:Choice(Requires="a14") > a14:m > (m:oMathPara|m:oMath)`。若 Choice 里直接放 oMath（无 a14:m），PowerPoint 会**静默丢弃整个 AlternateContent（连 Fallback 也不显示）**——公式全部消失但文件能正常打开。
11. **命名空间声明位置**：`xmlns:a` 声明在 `mc:AlternateContent` 根上（Fallback 里有 `a:r`/`a:t`）；`xmlns:m`+`xmlns:a` 声明在 `m:oMathPara`（display）或 `m:oMath`（inline）根上。声明放错层级 → lxml `Namespace prefix not defined`。
12. **nary+undOvr+supHide 会丢弃整个文本框**：`m:nary`（∫）只带下限（limLoc="undOvr" + supHide=1，不写 m:sup）时，PDF 导出中该文本框**所有内容（含纯文本段）全部消失且无报错**。改用 `m:limLow`（m:e 放 ∫ run，m:lim 放下限）。排查手段：把各公式拆到独立 slide 二分测试。
13. **oMathPara 与普通文字 run 混在同一段落** → 公式被压成 inline 小字号（分数变小、重叠）。display 公式必须独占段落；行内短公式用 inline oMath 无碍。
33. **整段正文塞进一个 oMath**：正文与公式同处一个 `mc:AlternateContent` 时，只认 `mc:Fallback` 的渲染器会用一句占位文本整体替换，正文陪葬。正文必须走普通 `a:r`，只有公式进 AlternateContent；Fallback 默认应是片段的可读线性文本，别留 `[formula]`。

## C. 版式与布局类

14. **add_picture 尺寸参数必须是 EMU**：python-pptx 的 `add_picture(path, x, y, w, h)` 若 w/h 传英寸浮点数，会被当作 EMU ≈ 0 尺寸，图片退化为零尺寸点、肉眼不可见。务必 `Inches()` 包裹。
15. **display 公式段落行距 1.4 会把行高撑到 ~1in**（分数+行距叠乘），连续多行公式易溢出幻灯片底边。纯公式段落建议 line_spacing=1.0。
17. **PowerShell 输出捕获可能失效**：脚本一律把结果写入 UTF-8 文件再读取；`| Out-File -Encoding utf8` 重定向会产生 UTF-16/BOM 混乱。
18. **shape 内 inline 公式颜色继承主题浅色**：数学 run 无显式颜色，在浅色填充的 shape 里几乎不可见。修复：run 属性里写入 `<a:solidFill><a:srgbClr val="..."/></a:solidFill>`。
21. **样式传递遗漏**：构造段落的样式函数若不把 size 传到公式 run，公式不随方框缩放 → 方框内容溢出（文本缩了、公式没缩）。
22. **页脚双写**：slide 构造函数内部已加页脚，外部再调一次 → 两个坐标完全相同的页脚（审计报 100% 重叠）。批量生成后统计每页页脚出现次数，非 1 即有问题。
23. **版式审计用字形级 span 碰撞检测，不要用 block 级**：逐 span 取 bbox；交集面积/较小面积 ≥ 0.40 判为真压字。block 级会把 display 公式分子/分母拆成相邻块而大量误报；下标与基底轻微相交（~0.30）属正常。另两项：内容最低点 ≤ 页脚上缘；每页页脚对象数 = 1。
24. **行高预算（20pt/1.4 实测）**：正文 ≈33.9 pt/行；无分数 display 公式 ≈1.0 行，含分数 ≈1.7 行；每行字符数 ≈ 宽度(in)×72/8.7。溢出页修法优先级：合并连续 display 公式为一行 > 删冗余从句 > 拆页。**不要全局缩字号**。
26. **解答框 top 是硬编码的**：缩短引言后不会自动上移，必须同步改 top（引言结束 y ≈ 起点 + 行数×行高，解答框 top ≥ 结束 y + 8pt）。
27. **三件套全零不等于公式正确**：审计输入必须是 PowerPoint COM 原生导出的 PDF；若 PDF 来自只认 `mc:Fallback` 的渲染器，整页可能只剩占位文本而 collisions/overflow/footer 照样全零。审计前先确认 PDF 含 CambriaMath 字形，命中 `[formula]`/`□` 即判违规（坑 33）。
31. **题目密集页直接拆页**（按题组分页、配图随题组走），比无限压缩文案可靠。
32. **微溢出（<5pt）两种快速修法**：图注压线 → 图 top 上移 0.15–0.25 in；文字差 1–2pt → 精简该段为更短单行。
37. **含多个分式的解答页按小问拆页**：各小问含 2–3 个 display 分式时一页必然溢出 40pt+，直接一小问一页。
38. **超宽 display 公式会让 m:d 括号断行**（PDF 里出现单独一行的右括号）。修法：在等号处拆成两个 display 段（第二段以 `= ` 开头）。
39. **审计脚本的页脚排除要防假阳性**：页脚文案过长就不再被排除条件命中，每页报出「页脚自身高度」的假溢出。页脚保持短（≤50 字符），或按「字号 ≤10 且位于底部」排除。
43. **批量写作公式的两个高频坑**：(a) 空基底反复出现，拼接前必须清零；(b) 纯文本串塞进公式位 → 裸文本 OMML → 拒开。数值+单位尽量用 Unicode 上下标（`10⁻¹⁹`、`q₁`）写进文本 run，少造结构就少踩坑。

## D. 数学记法特例

19. **上标/下标 helper 别用反**：S=上标（sSup）、Q=下标（sSub）；写 a₁ 要用下标。PDF 核验时注意区分 a¹ vs a₁。
20. **小体量配图用 matplotlib 生成**：200dpi、与 deck 风格一致的配色，`bbox_inches='tight'` 保存后按宽计算放置尺寸。
34. **列向量/矩阵是线代教材刚需**：m:d > m:e > m:eqArr 结构，两行向量行高 ≈1.3 行、三行 ≈1.7 行。范数用 U+2016 字面量。
35. **禁止 m:acc 画矢量箭头**：`<m:acc><m:chr m:val="→"/>` 在 PowerPoint→PDF 链路渲染成亚美尼亚标点并压在字母前。矢量记法退化为纯文本 `v→`，或文字说明。

## E. PDF 提取 / 裁图类

8. **COM 崩溃后先清进程**：`Stop-Process POWERPNT -Force` 清理挂起实例，用 `DispatchEx` 起新实例。
9. **页码 0-based/1-based 差 1**：PyMuPDF `p.number` 是 0-based；核对版式时用页面文本匹配标题定位，不要靠记忆页码。
16. **图版式 PPT 的图多为嵌入位图**：先 `get_image_info(xrefs=True)` 拿精确 bbox，再按 xref 原分辨率提取（比 clip 渲染裁剪干净、无杂边）；仅矢量图形才用 clip 渲染裁剪。
25. **按块替换源码时的正则陷阱**：块边界正则匹配不上连续等号行 → 块边界判错、静默吞掉相邻内容。替换后必须核对 slide 创建调用数与页数一致。
28. **扫描版整页位图 PDF 裁图的 prose 判别三判据**：`prose_ratio ≥ 0.86` + `longest_word ≥ 4`（拒绝图形标签）+ `y1 ≥ 58`（排除页眉）。每页先算正文块最小 x0 定左边界；列判定用中缝附近阈值；个别图手工加 per-figure 覆盖字典。渲染后白边 trim 兜底。
29. **patch 脚本拆块必须「先取全部锚点、再从后往前 splice」**：否则边界吞掉中间内容，下一步报 substring not found（看似目标缺失，实为被上一步删除）。
30. **源码里的 unicode 是转义形式**（`\u2013` 六个字符而非真实字符）：patch 要用字面转义串构造，否则静默 no-op。大段 patch 不用 bash heredoc（引号转义炸），用 Write 写 .py 再执行，结尾 `ast.parse` + 调用计数校验。
40. **扫描+OCR 型 PDF**：正文可用文本层，但公式必须以渲染页目视为准；不要相信 get_text 里的乱码串。
41. **OCR 会把 caption 编号拆散**或把图形标签并进 caption 块。对策：先合并被拆数字，再严格匹配 `^Figure N\.(\d+)[:.]`（带标点才收，排除正文引用）；被合并的 caption 用手工 rect 兜底。
42. **扫描页边缘的彩色装订线会混进裁图**（属位图墨迹非矢量，get_drawings 查不到）：对整列做像素扫描定位线的 x 坐标，裁剪 x0 设到线右侧；顶/底同理。
44. **bash `-c` 一行补丁里写 `\\uXXXX` 会被 Python 解释成真实字符**，而源文件里是字面六字符转义 → replace 静默 no-op。一律用 Write 写补丁文件执行。

## F. 工程流程类（方法论）

- **视觉校验闭环是生命线**：构建 → PowerPoint COM `SaveAs(pdf, 32)` 导出 PDF（公式原生渲染，一个 PDF 同时核验版式+公式）→ PyMuPDF 渲染拼 contact sheet 逐页检查 → 字形级审计跑三件套（collisions=0 / overflow=0 / footer=1）→ 迭代到全零。
- **审计比目视快且无遗漏**；目视兜底处理审美问题（对齐、留白），审计兜底处理硬错误。
- **补丁一律走「Write 脚本 → ast.parse → 调用计数校验」流程**，禁止 bash -c 内联改源码。
