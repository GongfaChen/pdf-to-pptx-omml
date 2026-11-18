# omml_helpers 使用手册（cookbook）

## 核心规则

1. **helper 的输出不要再用 T() 包**——T() 只接受纯文本；嵌套时直接传 helper 返回值。
2. 裸字符串可以当快捷写法传给任何 helper（自动包 T()）：`S('v', '2')` 等价 `S(T('v'), T('2'))`。
3. 空字符串会直接抛 `ValueError`（这是特性：坑 1 在写代码时就爆，而不是 PowerPoint 拒开时爆）。

## 基础公式

```python
from omml_helpers import T, F, S, Q, SQ, RAD, LIM, PAR, ABS, CV, MX, NORM, M, insert_math, set_run_style

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

# 行内公式：文本 run + 短数学 + 文本 run 三段式（坑 5）
insert_math(p, M(T('动能 '), S('E', 'k'), T(' = ')), display=False)
insert_math(p, F('1', '2'), display=False)
insert_math(p, M(S('mv', ''), T('²。')), display=False)

# display 公式：独占段落
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
insert_math(p, 'x=2')         # ✗ 坑 6：裸字符串 → 应 M(T('x'), T('='), T('2'))
```

正确写法：

```python
M(T(' '), F('3', '5'), T(' '))
insert_math(p, M(T('x'), T('='), T('2')))
```

## 单位与希腊字母

- 单位、文字说明用 `T('kg', italic=False)`（正体）。
- 希腊字母直接用 Unicode 字符：`T('μ₀')`、`T('α')`——比造 Q/S 结构稳。
- 数值+单位尽量用 Unicode 上下标：`10⁻¹⁹`、`q₁`（坑 43）。
