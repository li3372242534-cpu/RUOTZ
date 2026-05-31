# -*- coding: utf-8 -*-
"""
《光学原理》课程文献汇报 PPT 生成脚本
主题：色散 (Dispersion) —— 现代研究中的"消除"与"驾驭"
文献一：全可见光消色差超透镜 (Nature Communications, 2024)
文献二：色散/耗散工程化光孤子频率梳 (Light: Science & Applications, 2024)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# 全局配色（色散 / 光谱主题）
# ----------------------------------------------------------------------------
NAVY      = RGBColor(0x14, 0x1B, 0x41)   # 深靛蓝（标题栏、主色）
INK       = RGBColor(0x22, 0x26, 0x33)   # 正文深灰
GREY      = RGBColor(0x5B, 0x61, 0x70)   # 次要文字
LIGHT     = RGBColor(0xF2, 0xF4, 0xF9)   # 浅底
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
VIOLET    = RGBColor(0x6C, 0x4A, 0xB6)   # 强调紫（光谱短波端）
CYAN      = RGBColor(0x13, 0xA8, 0x9E)   # 强调青
AMBER     = RGBColor(0xE8, 0x9A, 0x1C)   # 强调琥珀（光谱长波端）
RED       = RGBColor(0xD9, 0x4A, 0x4A)
CARD_BLUE = RGBColor(0xE8, 0xEC, 0xF7)
CARD_TEAL = RGBColor(0xE2, 0xF3, 0xF1)
CARD_AMBR = RGBColor(0xFB, 0xF0, 0xDD)

# 光谱色带（用于装饰条）
SPECTRUM = [
    RGBColor(0x6C, 0x4A, 0xB6),  # 紫
    RGBColor(0x3A, 0x5B, 0xD0),  # 蓝
    RGBColor(0x13, 0xA8, 0x9E),  # 青绿
    RGBColor(0x6A, 0xB8, 0x3A),  # 绿
    RGBColor(0xE8, 0xC8, 0x1C),  # 黄
    RGBColor(0xE8, 0x9A, 0x1C),  # 橙
    RGBColor(0xD9, 0x4A, 0x4A),  # 红
]

CN_FONT = "微软雅黑"
CN_FONT_LIGHT = "微软雅黑"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# 辅助函数
# ----------------------------------------------------------------------------
def set_cn_font(run, name=CN_FONT):
    """设置中文(东亚)字体，保证 PowerPoint 中正确渲染中文。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', name)
    cs = rPr.find(qn('a:cs'))
    if cs is None:
        cs = rPr.makeelement(qn('a:cs'), {})
        rPr.append(cs)
    cs.set('typeface', name)


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill, line=None, line_w=None, shadow=False):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def round_rect(slide, x, y, w, h, fill, line=None, line_w=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def spectrum_bar(slide, x, y, w, h):
    """绘制一条光谱渐变装饰条（用 7 个色块拼接）。"""
    n = len(SPECTRUM)
    seg = int(w / n)
    for i, c in enumerate(SPECTRUM):
        rect(slide, x + seg * i, y, seg + 1, h, c)


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
            word_wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tb, tf


def add_para(tf, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT,
             space_before=0, space_after=6, first=False, line_spacing=1.1,
             font=CN_FONT, level=0):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    p.level = level
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    set_cn_font(r, font)
    return p, r


def title_bar(slide, kicker, title, page_no):
    """统一的内容页标题栏。"""
    # 顶部光谱细条
    spectrum_bar(slide, 0, 0, EMU_W, Inches(0.16))
    # 左侧竖直强调条
    rect(slide, Inches(0.0), Inches(0.16), Inches(0.16), Inches(1.34), NAVY)
    # kicker（小标签）
    tb, tf = textbox(slide, Inches(0.55), Inches(0.30), Inches(11.0), Inches(0.35))
    add_para(tf, kicker, 13, VIOLET, bold=True, first=True, space_after=0)
    # 主标题
    tb, tf = textbox(slide, Inches(0.55), Inches(0.62), Inches(12.2), Inches(0.85))
    add_para(tf, title, 27, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.0)
    # 页码
    tb, tf = textbox(slide, Inches(12.4), Inches(7.02), Inches(0.8), Inches(0.35),
                     align=PP_ALIGN.RIGHT)
    add_para(tf, str(page_no), 11, GREY, first=True, space_after=0,
             align=PP_ALIGN.RIGHT)
    # 底部细线
    rect(slide, Inches(0.55), Inches(1.52), Inches(12.2), Pt(2), LIGHT)


def bullet(tf, text, size=15, color=INK, bold=False, first=False,
           space_after=8, indent=False, marker="●", marker_color=VIOLET,
           line_spacing=1.15):
    """带圆点项目符号的段落（标记与文字同段，颜色可区分）。"""
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    rm = p.add_run()
    rm.text = ("    " if indent else "") + marker + "  "
    rm.font.size = Pt(size - 2)
    rm.font.bold = True
    rm.font.color.rgb = marker_color
    set_cn_font(rm)
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    set_cn_font(r)
    return p


def card(slide, x, y, w, h, fill, bar_color, head, head_color=None):
    """带顶部色条的卡片，返回内容文本框。"""
    round_rect(slide, x, y, w, h, fill)
    rect(slide, x, y, w, Inches(0.12), bar_color)
    tb, tf = textbox(slide, x + Inches(0.25), y + Inches(0.30),
                     w - Inches(0.5), h - Inches(0.5))
    if head:
        add_para(tf, head, 15, head_color or bar_color, bold=True, first=True,
                 space_after=7)
    return tf


# ============================================================================
# 1. 封面
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, NAVY)
# 背景装饰：右下角光谱色块
for i, c in enumerate(SPECTRUM):
    rect(s, Inches(0.0), Inches(6.6) + Pt(2), EMU_W, Inches(0.9), NAVY)
spectrum_bar(s, 0, Inches(7.16), EMU_W, Inches(0.34))
spectrum_bar(s, 0, 0, EMU_W, Inches(0.18))

# 课程标签
tb, tf = textbox(s, Inches(1.0), Inches(1.15), Inches(11), Inches(0.5))
add_para(tf, "《光学原理》 · 课程文献汇报", 18, RGBColor(0xB9, 0xC4, 0xE8),
         bold=True, first=True, space_after=0)

# 主标题
tb, tf = textbox(s, Inches(1.0), Inches(2.05), Inches(11.4), Inches(2.2))
add_para(tf, "色散：被消除，还是被驾驭？", 44, WHITE, bold=True, first=True,
         space_after=10, line_spacing=1.05)
add_para(tf, "从课本的色差与群速度色散，到超透镜与片上光孤子频率梳",
         20, RGBColor(0xC9, 0xD2, 0xEC), first=False, space_after=0,
         line_spacing=1.2)

# 副信息
tb, tf = textbox(s, Inches(1.0), Inches(4.55), Inches(11), Inches(1.6))
add_para(tf, "知识点出处：第 X 章  光的吸收、色散与散射", 15,
         RGBColor(0xAE, 0xB8, 0xDB), first=True, space_after=8)
add_para(tf, "文献一  全可见光消色差超透镜 ｜ Nature Communications, 2024", 15,
         WHITE, first=False, space_after=6)
add_para(tf, "文献二  色散/耗散工程化光孤子频率梳 ｜ Light: Sci. & Appl., 2024", 15,
         WHITE, first=False, space_after=0)

tb, tf = textbox(s, Inches(1.0), Inches(6.5), Inches(11), Inches(0.5))
add_para(tf, "汇报人：________     日期：________", 13,
         RGBColor(0x8F, 0x9B, 0xC4), first=True, space_after=0)


# ============================================================================
# 2. 目录
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "OUTLINE", "汇报提纲", 2)

items = [
    ("01", "课本知识点回顾", "色散的定义、正常/反常色散、色差与群速度色散", VIOLET),
    ("02", "为什么选这两篇文献", "围绕同一知识点的两种相反思路：消除 vs 驾驭", CYAN),
    ("03", "文献一：消色差超透镜", "把色差“工程化地消除”——全可见光平面透镜", AMBER),
    ("04", "文献二：光孤子频率梳", "把反常色散“主动利用”——片上倍频程光频梳", RED),
    ("05", "对比 · 总结 · 启示", "色散是敌是友？现代光学如何重新定义它", NAVY),
]
y = Inches(1.95)
for num, t, d, c in items:
    round_rect(s, Inches(0.7), y, Inches(0.95), Inches(0.85), c)
    tb, tf = textbox(s, Inches(0.7), y, Inches(0.95), Inches(0.85),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, num, 22, WHITE, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=0)
    tb, tf = textbox(s, Inches(1.85), y + Inches(0.04), Inches(10.6), Inches(0.8),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, t, 18, NAVY, bold=True, first=True, space_after=2)
    add_para(tf, d, 13, GREY, first=False, space_after=0)
    y += Inches(0.98)


# ============================================================================
# 3. 课本知识点回顾
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "01 · 课本知识点回顾", "色散：折射率随波长而变  n = n(λ)", 3)

# 左：核心概念
tf = card(s, Inches(0.55), Inches(1.85), Inches(6.0), Inches(4.95),
          LIGHT, VIOLET, "教材中的色散（经典内容）")
bullet(tf, "色散：介质折射率 n 随光的波长 λ（频率）变化", 14)
bullet(tf, "正常色散：dn/dλ < 0，波长越短折射率越大", 14,
       marker_color=CYAN)
bullet(tf, "     → 柯西公式  n = A + B/λ² + C/λ⁴", 13, color=GREY,
       marker="–", marker_color=GREY, indent=True)
bullet(tf, "反常色散：在介质吸收带附近 dn/dλ > 0", 14, marker_color=AMBER)
bullet(tf, "相速度 vs 群速度；群速度色散 GVD：不同频率成分传播速度不同 → 脉冲展宽",
       14, marker_color=RED)
bullet(tf, "典型后果——色差：透镜对红光、蓝光焦点不在同一处，成像发虚、有彩边",
       14, marker_color=VIOLET)

# 右：可视化（棱镜分光示意 + 色带）
round_rect(s, Inches(6.85), Inches(1.85), Inches(5.9), Inches(2.55), CARD_BLUE)
tb, tf = textbox(s, Inches(7.1), Inches(2.0), Inches(5.4), Inches(0.4))
add_para(tf, "棱镜分光：色散最直观的体现", 14, NAVY, bold=True, first=True,
         space_after=0)
# 入射白光
rect(s, Inches(7.15), Inches(3.15), Inches(1.7), Pt(4), RGBColor(0x88,0x88,0x88))
# 三角棱镜
tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(8.85),
                         Inches(2.7), Inches(1.0), Inches(1.2))
tri.rotation = 90
tri.fill.solid(); tri.fill.fore_color.rgb = RGBColor(0xD7,0xDE,0xF0)
tri.line.color.rgb = NAVY; tri.line.width = Pt(1.2); tri.shadow.inherit = False
# 出射光谱（扇形展开的彩线）
import math
ox, oy = Inches(9.7), Inches(3.2)
for i, c in enumerate(SPECTRUM):
    ln = s.shapes.add_connector(2, ox, oy,
                                ox + Inches(2.6),
                                oy + Inches(-0.55) + Inches(0.18) * i)
    ln.line.color.rgb = c; ln.line.width = Pt(3); ln.shadow.inherit = False

# 右下：为什么要找新文献
tf = card(s, Inches(6.85), Inches(4.55), Inches(5.9), Inches(2.25),
          CARD_AMBR, AMBER, "教材“老”在哪里？")
bullet(tf, "课本多停留在棱镜、透镜色差、光纤色散等经典现象", 13,
       marker_color=AMBER)
bullet(tf, "近十年：人们既能把色散“精确抹平”，也能把它“为我所用”", 13,
       marker_color=AMBER, bold=True)
bullet(tf, "下面用两篇 2024 年顶刊文献展示这两条全新路线", 13,
       marker_color=AMBER)


# ============================================================================
# 4. 为什么选这两篇
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "02 · 选题逻辑", "同一个色散，两种相反的现代思路", 4)

tb, tf = textbox(s, Inches(0.7), Inches(1.75), Inches(12), Inches(0.6))
add_para(tf, "色散在课本里常被当作“成像的麻烦”。但现代光学告诉我们：它既可以被彻底消除，也可以被精准驾驭。",
         16, INK, first=True, space_after=0, line_spacing=1.2)

# 左卡片：消除
tf = card(s, Inches(0.7), Inches(2.55), Inches(5.85), Inches(4.0),
          CARD_BLUE, VIOLET, "思路 A · 消除色散")
bullet(tf, "目标：让所有颜色聚焦到同一点（消色差）", 14, marker_color=VIOLET)
bullet(tf, "对象：超透镜 metalens —— 亚波长纳米结构构成的超薄平面透镜", 14,
       marker_color=VIOLET)
bullet(tf, "关键：同时控制相位 + 群延迟，对抗色差", 14, marker_color=VIOLET)
bullet(tf, "→ 文献一 (Nature Communications, 2024)", 14, color=NAVY, bold=True,
       marker="★", marker_color=VIOLET)

# 右卡片：驾驭
tf = card(s, Inches(6.8), Inches(2.55), Inches(5.85), Inches(4.0),
          CARD_TEAL, CYAN, "思路 B · 驾驭色散")
bullet(tf, "目标：用反常色散平衡非线性，生成稳定光孤子", 14, marker_color=CYAN)
bullet(tf, "对象：片上微环谐振腔 —— 薄膜铌酸锂 (TFLN)", 14, marker_color=CYAN)
bullet(tf, "关键：工程化色散 + 耗散，得到倍频程频率梳", 14, marker_color=CYAN)
bullet(tf, "→ 文献二 (Light: Sci. & Appl., 2024)", 14, color=NAVY, bold=True,
       marker="★", marker_color=CYAN)

# 中间“VS”
oval = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.25), Inches(4.05),
                          Inches(1.0), Inches(1.0))
oval.fill.solid(); oval.fill.fore_color.rgb = NAVY
oval.line.color.rgb = WHITE; oval.line.width = Pt(2); oval.shadow.inherit = False
tb, tf = textbox(s, Inches(6.25), Inches(4.05), Inches(1.0), Inches(1.0),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_para(tf, "VS", 20, WHITE, bold=True, first=True, align=PP_ALIGN.CENTER,
         space_after=0)


# ============================================================================
# 文献一系列
# ============================================================================
# 5. 文献一：背景
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 背景与问题", "消色差超透镜：把色差工程化地消除", 5)

# 文献信息条
round_rect(s, Inches(0.7), Inches(1.75), Inches(12.0), Inches(1.05), NAVY)
tb, tf = textbox(s, Inches(0.95), Inches(1.86), Inches(11.5), Inches(0.85),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "Achromatic metalenses for full visible spectrum with extended group delay control via dispersion-matched layers",
         15, WHITE, bold=True, first=True, space_after=3, line_spacing=1.05)
add_para(tf, "Nature Communications (2024)  ·  关键词：超透镜 / 消色差 / 群延迟 / 色散工程",
         12, RGBColor(0xC9, 0xD2, 0xEC), first=False, space_after=0)

tf = card(s, Inches(0.7), Inches(3.05), Inches(5.85), Inches(3.6),
          LIGHT, VIOLET, "问题背景")
bullet(tf, "超透镜：用亚波长“超原子”阵列在平面上逐点控制光的相位，替代厚重的曲面玻璃透镜",
       14, marker_color=VIOLET)
bullet(tf, "天然缺陷：超透镜的相位本质上依赖波长，天生有强烈色差", 14,
       marker_color=VIOLET)
bullet(tf, "要消色差，每个超原子不仅要给对相位，还要给对“群延迟” GD",
       14, marker_color=VIOLET)

tf = card(s, Inches(6.8), Inches(3.05), Inches(5.85), Inches(3.6),
          CARD_AMBR, AMBER, "核心瓶颈：群延迟极限")
bullet(tf, "可实现的最大群延迟 ∝ 材料厚度与超原子库的丰富度，存在上限", 14,
       marker_color=AMBER)
bullet(tf, "群延迟极限 → 直接限制了消色差超透镜的 口径 × 数值孔径(NA)", 14,
       marker_color=AMBER)
bullet(tf, "因此过去消色差超透镜口径很小，难以做大、难以实用", 14,
       marker_color=AMBER)


# 6. 文献一：方法
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 方法与创新", "用“色散匹配层”突破群延迟极限", 6)

tf = card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.95),
          CARD_BLUE, VIOLET, "核心创新点")
bullet(tf, "提出“色散匹配层 (dispersion-matched layers)”：在超原子层之外，叠加专门提供额外群延迟的辅助结构层",
       14, marker_color=VIOLET)
bullet(tf, "两层的色散特性相互匹配 → 大幅扩展可控群延迟范围", 14,
       marker_color=VIOLET)
bullet(tf, "等价于突破了单层超表面的群延迟极限，从而放大消色差口径",
       14, marker_color=VIOLET)
bullet(tf, "在整个可见光波段实现“同焦点”：红、绿、蓝聚焦到同一平面",
       14, marker_color=VIOLET, bold=True)

# 右：相位+群延迟双重条件示意
tf = card(s, Inches(6.8), Inches(1.85), Inches(5.9), Inches(4.95),
          LIGHT, CYAN, "为什么需要“群延迟”？")
bullet(tf, "消色差条件 = 相位条件 + 群延迟条件", 14, marker_color=CYAN, bold=True)
bullet(tf, "相位 φ(λ₀)：保证中心波长正确聚焦", 13, marker_color=CYAN)
bullet(tf, "群延迟 dφ/dω：保证不同波长的波前“同时到达”焦点", 13,
       marker_color=CYAN)
bullet(tf, "群延迟需要的范围随 口径、NA 增大而增大", 13, marker_color=CYAN)
# 公式框
round_rect(s, Inches(7.05), Inches(5.05), Inches(5.4), Inches(1.45), WHITE)
tb, tf2 = textbox(s, Inches(7.25), Inches(5.18), Inches(5.0), Inches(1.2),
                  anchor=MSO_ANCHOR.MIDDLE)
add_para(tf2, "所需群延迟  ∝  (口径 / 2) · NA / c", 15, NAVY, bold=True,
         first=True, space_after=6)
add_para(tf2, "口径越大、NA 越高 → 需要越大的群延迟范围", 12, GREY,
         first=False, space_after=0)


# 7. 文献一：结果
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 实验结果", "做出了一整套不同口径的消色差超透镜", 7)

# 结果数据卡（四个口径/NA）
data = [
    ("16 μm",  "NA 0.27"),
    ("66 μm",  "NA 0.11"),
    ("200 μm", "NA 0.04"),
    ("400 μm", "NA 0.02"),
]
x = Inches(0.7)
for i, (d, na) in enumerate(data):
    c = SPECTRUM[i * 2]
    round_rect(s, x, Inches(1.9), Inches(2.85), Inches(1.7), LIGHT)
    rect(s, x, Inches(1.9), Inches(2.85), Inches(0.12), c)
    tb, tf = textbox(s, x, Inches(2.15), Inches(2.85), Inches(1.4),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, d, 28, NAVY, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=2)
    add_para(tf, na, 15, c, bold=True, first=False, align=PP_ALIGN.CENTER,
             space_after=2)
    add_para(tf, "口径 / 数值孔径", 11, GREY, first=False, align=PP_ALIGN.CENTER,
             space_after=0)
    x += Inches(3.0)

tf = card(s, Inches(0.7), Inches(3.95), Inches(12.0), Inches(2.85),
          CARD_BLUE, VIOLET, "结论与意义")
bullet(tf, "在 16 / 66 / 200 / 400 μm 多种口径上设计、加工并实测了消色差超透镜，覆盖全可见光波段",
       15, marker_color=VIOLET)
bullet(tf, "“色散匹配层”策略可扩展群延迟、放大消色差口径，为更大尺寸、更高 NA 的平面消色差透镜提供了路径",
       15, marker_color=VIOLET)
bullet(tf, "应用前景：手机/内窥镜/AR 眼镜等超薄成像模组，替代厚重的多片消色差镜组",
       15, marker_color=VIOLET, bold=True)


# 8. 文献一：与课本联系
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 与课本知识点的联系", "课本概念 → 前沿研究", 8)

pairs = [
    ("色差 (chromatic aberration)", "正是要消除的对象——让红绿蓝重新聚焦于同一焦点", VIOLET),
    ("正常色散 / n = n(λ)", "超透镜相位天然随波长变化，是色差的根源", CYAN),
    ("相速度 vs 群速度 / 群延迟", "消色差的关键判据从“相位”升级为“相位 + 群延迟”", AMBER),
    ("色散是“成像的麻烦”", "现代手段把它工程化地“抹平”，而非简单回避", RED),
]
y = Inches(1.95)
for kbook, research, c in pairs:
    round_rect(s, Inches(0.7), y, Inches(4.7), Inches(1.05), LIGHT)
    rect(s, Inches(0.7), y, Inches(0.14), Inches(1.05), c)
    tb, tf = textbox(s, Inches(1.0), y, Inches(4.3), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, kbook, 15, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.05)
    # 箭头
    ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.5), y + Inches(0.32),
                            Inches(0.8), Inches(0.4))
    ar.fill.solid(); ar.fill.fore_color.rgb = c; ar.line.fill.background()
    ar.shadow.inherit = False
    round_rect(s, Inches(6.45), y, Inches(6.25), Inches(1.05), CARD_BLUE)
    tb, tf = textbox(s, Inches(6.7), y, Inches(5.8), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, research, 14, INK, first=True, space_after=0, line_spacing=1.1)
    y += Inches(1.2)


# ============================================================================
# 文献二系列
# ============================================================================
# 9. 文献二：背景
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 背景与问题", "光孤子频率梳：把反常色散主动利用", 9)

round_rect(s, Inches(0.7), Inches(1.75), Inches(12.0), Inches(1.05), CYAN)
tb, tf = textbox(s, Inches(0.95), Inches(1.86), Inches(11.5), Inches(0.85),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "Octave-spanning Kerr soliton frequency combs in dispersion- and dissipation-engineered lithium niobate microresonators",
         15, WHITE, bold=True, first=True, space_after=3, line_spacing=1.05)
add_para(tf, "Light: Science & Applications (2024) · Harvard, Lončar 组 · 关键词：光孤子 / 频率梳 / 反常色散 / 铌酸锂",
         12, RGBColor(0xDF, 0xF3, 0xF1), first=False, space_after=0)

tf = card(s, Inches(0.7), Inches(3.05), Inches(5.85), Inches(3.6),
          LIGHT, CYAN, "什么是光孤子频率梳？")
bullet(tf, "光频率梳：频谱上等间距、像“梳齿”一样的激光线，是精密测量/光钟/光通信的核心光源",
       13.5, marker_color=CYAN)
bullet(tf, "微腔孤子梳：在微环谐振腔内形成稳定传输的光脉冲（耗散克尔孤子 DKS），把整套梳搬上芯片",
       13.5, marker_color=CYAN)
bullet(tf, "孤子能稳定存在的物理基础：反常色散与克尔非线性相互平衡",
       13.5, marker_color=CYAN, bold=True)

tf = card(s, Inches(6.8), Inches(3.05), Inches(5.85), Inches(3.6),
          CARD_AMBR, AMBER, "难点在哪里？")
bullet(tf, "薄膜铌酸锂(TFLN)性能优异，但存在强拉曼效应，会“抢走”能量、抑制孤子形成",
       13.5, marker_color=AMBER)
bullet(tf, "要同时满足：合适的反常色散 + 抑制拉曼，难度大、器件成品率低",
       13.5, marker_color=AMBER)
bullet(tf, "目标：稳定、可重复地做出“倍频程(覆盖一个倍频)”的宽谱孤子梳",
       13.5, marker_color=AMBER)


# 10. 文献二：方法
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 方法与创新", "同时工程化“色散”与“耗散”", 10)

tf = card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.95),
          CARD_TEAL, CYAN, "核心创新点")
bullet(tf, "色散工程：调节微环几何，精确设计反常群速度色散(GVD)，为孤子提供存在条件",
       14, marker_color=CYAN)
bullet(tf, "耗散工程：设计腔的“损耗谱”，在拉曼增益频段引入更高损耗，从而抑制拉曼激射",
       14, marker_color=CYAN)
bullet(tf, "通过控制自由光谱范围(FSR)与耗散谱，压制抑制孤子的拉曼效应",
       14, marker_color=CYAN)
bullet(tf, "得到完整连通、覆盖一个倍频程的孤子梳，且器件成品率接近 100%",
       14, marker_color=CYAN, bold=True)

# 右：孤子=色散 vs 非线性平衡 示意
tf = card(s, Inches(6.8), Inches(1.85), Inches(5.9), Inches(2.6),
          LIGHT, AMBER, "孤子的物理图像")
bullet(tf, "色散：让脉冲在时间上展宽", 14, marker_color=VIOLET)
bullet(tf, "克尔非线性：让脉冲在时间上压缩", 14, marker_color=AMBER)
bullet(tf, "两者恰好平衡 → 形状不变的稳定脉冲 = 孤子", 14, marker_color=CYAN,
       bold=True)
# 平衡天平式公式条
round_rect(s, Inches(6.8), Inches(4.6), Inches(5.9), Inches(2.2), CARD_TEAL)
tb, tf2 = textbox(s, Inches(7.05), Inches(4.78), Inches(5.4), Inches(1.9),
                  anchor=MSO_ANCHOR.MIDDLE)
add_para(tf2, "反常色散  ⇄  克尔非线性", 18, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=8)
add_para(tf2, "（展宽）            （压缩）", 13, GREY, first=False,
         align=PP_ALIGN.CENTER, space_after=8)
add_para(tf2, "平衡 → 稳定光孤子 → 等间距频率梳", 14, CYAN, bold=True,
         first=False, align=PP_ALIGN.CENTER, space_after=0)


# 11. 文献二：结果
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 结果与应用", "倍频程、近乎 100% 成品率的片上孤子梳", 11)

stats = [
    ("Octave", "倍频程", "频谱覆盖一个完整倍频，利于自参考稳频", VIOLET),
    ("~100%", "器件成品率", "拉曼抑制策略让孤子器件可重复制造", CYAN),
    ("TFLN", "薄膜铌酸锂", "兼具克尔与电光效应的优异集成平台", AMBER),
]
x = Inches(0.7)
for big, small, desc, c in stats:
    round_rect(s, x, Inches(1.9), Inches(3.9), Inches(2.0), LIGHT)
    rect(s, x, Inches(1.9), Inches(3.9), Inches(0.12), c)
    tb, tf = textbox(s, x + Inches(0.2), Inches(2.15), Inches(3.5), Inches(1.65),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, big, 30, NAVY, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=0)
    add_para(tf, small, 14, c, bold=True, first=False, align=PP_ALIGN.CENTER,
             space_after=4)
    add_para(tf, desc, 11.5, GREY, first=False, align=PP_ALIGN.CENTER,
             space_after=0, line_spacing=1.05)
    x += Inches(4.05)

tf = card(s, Inches(0.7), Inches(4.25), Inches(12.0), Inches(2.55),
          CARD_TEAL, CYAN, "意义与应用")
bullet(tf, "给出了在铌酸锂上抑制拉曼、稳定生成宽谱孤子梳的明确设计准则", 15,
       marker_color=CYAN)
bullet(tf, "倍频程频谱是实现“自参考”、把光频梳锁定为光钟/频率基准的前提条件", 15,
       marker_color=CYAN)
bullet(tf, "应用：芯片级光钟、精密光谱、相干光通信、微波/毫米波信号源、激光雷达(LiDAR)",
       15, marker_color=CYAN, bold=True)


# 12. 文献二：与课本联系
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 与课本知识点的联系", "课本概念 → 前沿研究", 12)

pairs2 = [
    ("反常色散 (dn/dλ > 0)", "不再是“反常”的累赘，而是孤子能稳定存在的必要条件", VIOLET),
    ("群速度色散 GVD", "通过微腔几何被精确“设计”，决定孤子能否形成与谱宽", CYAN),
    ("色散 ⇄ 非线性 平衡", "色散与克尔效应平衡 → 形状不变的光孤子", AMBER),
    ("光的吸收 / 损耗", "“耗散工程”主动设计损耗谱，抑制拉曼、保护孤子", RED),
]
y = Inches(1.95)
for kbook, research, c in pairs2:
    round_rect(s, Inches(0.7), y, Inches(4.7), Inches(1.05), LIGHT)
    rect(s, Inches(0.7), y, Inches(0.14), Inches(1.05), c)
    tb, tf = textbox(s, Inches(1.0), y, Inches(4.3), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, kbook, 15, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.05)
    ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.5), y + Inches(0.32),
                            Inches(0.8), Inches(0.4))
    ar.fill.solid(); ar.fill.fore_color.rgb = c; ar.line.fill.background()
    ar.shadow.inherit = False
    round_rect(s, Inches(6.45), y, Inches(6.25), Inches(1.05), CARD_TEAL)
    tb, tf = textbox(s, Inches(6.7), y, Inches(5.8), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, research, 14, INK, first=True, space_after=0, line_spacing=1.1)
    y += Inches(1.2)


# ============================================================================
# 13. 对比总结
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "05 · 对比与总结", "两篇文献：色散的一体两面", 13)

# 表头
headers = ["对比维度", "文献一 · 消色差超透镜", "文献二 · 光孤子频率梳"]
colw = [Inches(3.0), Inches(4.85), Inches(4.85)]
xs = [Inches(0.7), Inches(3.7), Inches(8.55)]
hy = Inches(1.85)
hcolors = [NAVY, VIOLET, CYAN]
for i in range(3):
    rect(s, xs[i], hy, colw[i], Inches(0.6), hcolors[i])
    tb, tf = textbox(s, xs[i], hy, colw[i], Inches(0.6),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, headers[i], 14, WHITE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=0)

rows = [
    ("对色散的态度", "消除（视为敌人）", "驾驭（视为朋友）"),
    ("用到的色散", "正常色散 / 色差 / 群延迟", "反常色散 / 群速度色散"),
    ("载体平台", "亚波长超表面（平面超透镜）", "片上微环（薄膜铌酸锂）"),
    ("关键创新", "色散匹配层，突破群延迟极限", "色散+耗散工程，抑制拉曼"),
    ("典型应用", "超薄消色差成像（手机/AR/内窥镜）", "光钟/光通信/激光雷达/微波源"),
]
ry = hy + Inches(0.6)
rh = Inches(0.83)
for j, row in enumerate(rows):
    bg = WHITE if j % 2 == 0 else LIGHT
    for i in range(3):
        rect(s, xs[i], ry, colw[i], rh, bg)
        tb, tf = textbox(s, xs[i] + Inches(0.12), ry, colw[i] - Inches(0.24), rh,
                         anchor=MSO_ANCHOR.MIDDLE,
                         align=PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT)
        add_para(tf, row[i], 13, NAVY if i == 0 else INK,
                 bold=(i == 0), first=True,
                 align=PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT,
                 space_after=0, line_spacing=1.05)
    ry += rh
# 表格描边
rect(s, Inches(0.7), hy, Inches(12.0), Pt(2.5), NAVY)


# ============================================================================
# 14. 结论与启示
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, NAVY)
spectrum_bar(s, 0, 0, EMU_W, Inches(0.16))
spectrum_bar(s, 0, Inches(7.34), EMU_W, Inches(0.16))

tb, tf = textbox(s, Inches(0.9), Inches(0.7), Inches(11.5), Inches(0.7))
add_para(tf, "总结与启示", 30, WHITE, bold=True, first=True, space_after=0)

tf_card = round_rect(s, Inches(0.9), Inches(1.7), Inches(11.55), Inches(4.7),
                     RGBColor(0x1E, 0x27, 0x52))
tb, tf = textbox(s, Inches(1.3), Inches(2.0), Inches(10.8), Inches(4.2))
add_para(tf, "1.  色散不只是“成像的麻烦”", 18, SPECTRUM[4], bold=True,
         first=True, space_after=4)
add_para(tf, "     课本把色散主要当作要克服的缺陷；现代光学既能精确消除它，也能主动利用它。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=14,
         line_spacing=1.2)
add_para(tf, "2.  “消除”与“驾驭”，本质都是色散工程 (dispersion engineering)", 18,
         SPECTRUM[2], bold=True, first=False, space_after=4)
add_para(tf, "     文献一靠“色散匹配层”抹平色差；文献二靠“色散+耗散工程”稳定孤子。控制色散，是共同的核心能力。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=14,
         line_spacing=1.2)
add_para(tf, "3.  经典概念在前沿仍是基石", 18, SPECTRUM[0], bold=True,
         first=False, space_after=4)
add_para(tf, "     正常/反常色散、群速度色散、色差……这些课本概念正是读懂 2024 年顶刊工作的钥匙。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=0,
         line_spacing=1.2)

tb, tf = textbox(s, Inches(0.9), Inches(6.55), Inches(11.5), Inches(0.6))
add_para(tf, "一句话：理解了色散，你就握住了从经典棱镜到片上光梳的同一根主线。",
         16, SPECTRUM[5], bold=True, first=True, space_after=0)


# ============================================================================
# 15. 参考文献
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "REFERENCES", "参考文献", 15)

refs = [
    ("[1]  全可见光消色差超透镜（文献一）",
     "Achromatic metalenses for full visible spectrum with extended group delay "
     "control via dispersion-matched layers. Nature Communications, 2024.",
     "https://www.nature.com/articles/s41467-024-53701-8", VIOLET),
    ("[2]  色散/耗散工程化光孤子频率梳（文献二）",
     "Y. Song, Y. Hu, X. Zhu, K. Yang, M. Lončar. Octave-spanning Kerr soliton "
     "frequency combs in dispersion- and dissipation-engineered lithium niobate "
     "microresonators. Light: Science & Applications, 2024.",
     "https://www.nature.com/articles/s41377-024-01546-7", CYAN),
    ("[3]  课本知识点参考",
     "《光学原理》/《现代光学基础》 第“光的吸收、色散与散射”章 —— 色散、色差、"
     "群速度色散等经典概念。",
     "", AMBER),
]
y = Inches(1.95)
for head, body, url, c in refs:
    round_rect(s, Inches(0.7), y, Inches(12.0), Inches(1.45), LIGHT)
    rect(s, Inches(0.7), y, Inches(0.14), Inches(1.45), c)
    tb, tf = textbox(s, Inches(1.0), y + Inches(0.15), Inches(11.4), Inches(1.2))
    add_para(tf, head, 15, NAVY, bold=True, first=True, space_after=4)
    add_para(tf, body, 12.5, INK, first=False, space_after=4, line_spacing=1.15)
    if url:
        add_para(tf, url, 11.5, c, first=False, space_after=0)
    y += Inches(1.65)

tb, tf = textbox(s, Inches(0.7), Inches(6.95), Inches(12), Inches(0.4))
add_para(tf, "注：两篇均为 2024 年发表的开放获取顶刊文献，可按上方链接查阅原文与图表。",
         11.5, GREY, first=True, space_after=0)


# ----------------------------------------------------------------------------
prs.save("光学原理_色散_文献汇报.pptx")
print("Saved: 光学原理_色散_文献汇报.pptx")
print("Total slides:", len(prs.slides._sldIdLst))
