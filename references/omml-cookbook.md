# omml_helpers 使用手册（cookbook）

## 核心规则

1. **helper 的输出不要再用 T() 包**——T() 只接受纯文本；嵌套时直接传 helper 返回值。
2. 裸字符串可以当快捷写法传给任何 helper（自动包 T()）：`S('v', '2')` 等价 `S(T('v'), T('2'))`。
3. 空字符串会直接抛 `ValueError`（这是特性：坑 1 在写代码时就爆，而不是 PowerPoint 拒开时爆）。
4. **段落里的说明文字不要用 T() 塞进公式**：正文走普通 `a:r`（`p.add_run()`），只有公式交给 `insert_math`；`display=True` 的公式必须独占段落（坑 5/13/33）。
5. `mc:Fallback` 默认由片段线性化出可读文本，不再是 `[formula]` 占位；需要精确线性文本时给 `insert_math(..., fallback="...")` 显式传值（坑 33）。
6. **每份 deck 构建前先 `reset_run_style()`**：`set_run_style()` 改的是模块级全局 `_STYLE`，size/color/font 传 `None` 表示「不修改」而非「清空」——设过一次颜色/字号后就没法用 `None` 清回来，连续构建多份 deck 会串味。重置后再设本 deck 样式（见基础公式示例）。
7. **helper 片段拼接用 `M(...)`，不要用 `+`**：`T('a') + T('b')` 是坑 45，详见下方「高频反模式」。

## 基础公式

```python
from omml_helpers import (T, F, S, Q, SQ, RAD, LIM, PAR, ABS, CV, MX, NORM, M,
                          SUM, PROD, EQARR, insert_math, set_run_style, reset_run_style)

# 每份 deck 构建前先重置全局样式，再设本 deck 的样式（见核心规则 6）
reset_run_style()                      # 清掉上一份 deck 残留的 size/color/font
set_run_style(color="1F3B73")          # 浅色 shape 里必须显式设深色（坑 18）

# 3/5
frac = F('3', '5')

# v²  与  a₁
v2  = S('v', '2')        # 上标用 S
a1  = Q('a', '1')        # 下标用 Q（坑 19）

# x²ᵢ
x2i = SQ('x', '2', 'i')

# √(b²−4ac)
disc = RAD(M(S('b', '2'), T('−4'), T('ac')))

# ∫₀^∞ 用 limLow（禁 nary+supHide，坑 12）
int_inf = LIM(T('∫'), T('0'))

# 定界符与绝对值
br  = PAR(M(T('x'), T('+'), T('y')))     # (x+y)
ab  = ABS('x')                            # |x|；尖括号 PAR(inner, '<', '>')

# 矢量与范数（禁 m:acc，坑 35）
vec = CV('x', S('v', 'T'))                # 列向量
nrm = NORM('v')                            # ‖v‖

# 矩阵
mat = MX(['a', 'b'], ['c', 'd'])

# 求和 / 求积（m:nary）：sub、sup 必须同时给，不写 supHide（坑 12）
sm  = SUM(Q('i', '1'), 'n', Q('a', 'i'))       # ∑_{i=1}^{n} a_i
pd  = PROD(Q('k', '1'), 'm', 'k')              # ∏_{k=1}^{m} k

# 方程组（m:eqArr，左对齐、不套 m:d；列向量才用 CV，坑 34）
eqs = EQARR(M(T('x'), T('+'), T('y'), T('=1')),
            M(T('x'), T('−'), T('y'), T('=0')))
```

## 插入段落（python-pptx）

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])
tf = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(12.1), Inches(1.2))
p = tf.text_frame.paragraphs[0]
p.line_spacing = 1.0                      # display 公式段行距 1.0（坑 15）

# 行内公式：正文走普通 a:r，只有公式交给 insert_math（坑 5/13/33）。
# 说明文字必须用 a:r 放在 mc:AlternateContent 之外，否则只认 mc:Fallback 的渲染器
# 会把整段替换成一句占位文本，正文陪葬（坑 33）。
def prose(paragraph, text):
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Times New Roman"
    run.font.size = Pt(20)

prose(p, '动能 ')
insert_math(p, S('E', 'k'), display=False)          # inline oMath：E_k
prose(p, ' = ')
insert_math(p, F('1', '2'), display=False)           # 1/2
insert_math(p, S('mv', '2'), display=False)          # (mv)² —— 别写 S('mv', '')：空上标会抛 ValueError（坑 1）
prose(p, '。')

# display 公式：独占段落，不与 a:r 正文混排（坑 13/15）
p2 = tf.add_paragraph()
p2.line_spacing = 1.0
insert_math(p2, M(S('E', 'k'), T(' = '), F('1', '2'), S('mv', '2')), display=True)

prs.save('demo_out.pptx')
```

## 高频反模式（validate_omml.py lint 会抓）

```python
T(' ' + F('3','5') + ' ')     # ✗ 坑 36：XML 变字面文本
T(F('3', '5'))                # ✗ 同上
S('', '2')                    # ✗ 坑 1：空基底（直接抛 ValueError）
S('mv', '')                   # ✗ 坑 1：空上标同样抛 ValueError，要写 S('mv', '2')
insert_math(p, 'x=2')         # ✗ 坑 6：裸字符串 → 应 M(T('x'), T('='), T('2'))
insert_math(p, M(T('动能 ')))  # ✗ 坑 33：正文应走 a:r，不要塞进公式
T('a') + T('b')               # ✗ 坑 45：helper 片段用 + 拼接（曾退化成裸文本被转义）
```

正确写法：

```python
M(T(' '), F('3', '5'), T(' '))
insert_math(p, M(T('x'), T('='), T('2')))
M(T('a'), T('b'))             # ✓ 拼接一律用 M(...)，不要用 +
```

**关于 `T('a') + T('b')`（坑 45）**：`_Frag` 是 `str` 子类，旧版本相加会默认走 `str.__add__` 退化成普通
`str`，下游 `_wrap` 把它当裸文本用 `T()` 包裹并转义——`F(T('a') + T('b'), '2')` 的 `m:t` 里出现
`&lt;m:r&gt;`，公式显示成 XML 字面文本。当前 `omml_helpers` 已重载 `_Frag.__add__/__radd__`，任意方向
拼接结果都仍是 `_Frag`，**此写法现在不会再错**；但仍建议一律用 `M(...)` 拼接——语义明确、跨版本更稳
（旧版 helper 库无此重载），且 `validate_omml.py lint` 会把 `+` 拼接判为问题。

## 单位与希腊字母

- 单位、文字说明用 `T('kg', italic=False)`（正体）。
- 希腊字母直接用 Unicode 字符：`T('μ₀')`、`T('α')`——比造 Q/S 结构稳。
- 数值+单位尽量用 Unicode 上下标：`10⁻¹⁹`、`q₁`（坑 43）。
