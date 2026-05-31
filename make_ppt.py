# -*- coding: utf-8 -*-
"""
《光学原理》课程文献汇报 PPT 生成脚本
主题：像差的衍射理论 (The Diffraction Theory of Aberrations)
     —— 从波像差函数、泽尼克多项式与点扩散函数(PSF)，到 AI 驱动的波前传感与像差校正

文献一：从点扩散函数直接预测泽尼克系数 (arXiv:2404.15231, 2024) —— 感知像差
文献二：物理信息图神经网络 / 频域感知的光学像差校正 (arXiv:2512.05683, 2025) —— 校正像差
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# 全局配色（波前 / 衍射主题：靛蓝—蓝—青—青绿渐变）
# ----------------------------------------------------------------------------
NAVY      = RGBColor(0x15, 0x1B, 0x40)   # 深靛蓝（标题栏、主色）
INK       = RGBColor(0x22, 0x26, 0x33)   # 正文深灰
GREY      = RGBColor(0x5B, 0x61, 0x70)   # 次要文字
LIGHT     = RGBColor(0xF1, 0xF4, 0xFA)   # 浅底
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
VIOLET    = RGBColor(0x6C, 0x4A, 0xB6)   # 强调紫
BLUE      = RGBColor(0x2E, 0x5B, 0xD0)   # 强调蓝
TEAL      = RGBColor(0x12, 0xA4, 0xA0)   # 强调青
CYAN      = RGBColor(0x2C, 0xB8, 0xC6)   # 强调青绿
AMBER     = RGBColor(0xE8, 0x95, 0x1C)   # 强调琥珀（提示/警示）
RED       = RGBColor(0xD9, 0x4A, 0x4A)
CARD_BLUE = RGBColor(0xE7, 0xEC, 0xF8)
CARD_TEAL = RGBColor(0xE1, 0xF2, 0xF2)
CARD_VIO  = RGBColor(0xEC, 0xE6, 0xF7)
CARD_AMBR = RGBColor(0xFB, 0xF0, 0xDD)

# 渐变控制色（用于装饰条）
GRAD_STOPS = [
    (0x5B, 0x3F, 0xA3),  # 紫
    (0x2E, 0x5B, 0xD0),  # 蓝
    (0x12, 0xA4, 0xA0),  # 青
    (0x36, 0xC6, 0xCC),  # 青绿
]

CN_FONT = "微软雅黑"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# 辅助函数
# ----------------------------------------------------------------------------
def lerp(a, b, t):
    return int(round(a + (b - a) * t))


def grad_color(t):
    """t in [0,1] -> RGBColor，沿 GRAD_STOPS 线性插值。"""
    n = len(GRAD_STOPS) - 1
    if t <= 0:
        r, g, b = GRAD_STOPS[0]
        return RGBColor(r, g, b)
    if t >= 1:
        r, g, b = GRAD_STOPS[-1]
        return RGBColor(r, g, b)
    pos = t * n
    i = int(pos)
    f = pos - i
    c0, c1 = GRAD_STOPS[i], GRAD_STOPS[i + 1]
    return RGBColor(lerp(c0[0], c1[0], f), lerp(c0[1], c1[1], f),
                    lerp(c0[2], c1[2], f))


def set_cn_font(run, name=CN_FONT):
    """设置中文(东亚)字体，保证 PowerPoint 中正确渲染中文。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', name)


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill, line=None, line_w=None):
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


def grad_bar(slide, x, y, w, h, n=48):
    """用 n 个细切片绘制平滑渐变条。"""
    seg = int(w / n)
    for i in range(n):
        rect(slide, x + seg * i, y, seg + 1, h, grad_color(i / (n - 1)))


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
             font=CN_FONT):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    set_cn_font(r, font)
    return p, r


def title_bar(slide, kicker, title, page_no):
    """统一的内容页标题栏。"""
    grad_bar(slide, 0, 0, EMU_W, Inches(0.16))
    rect(slide, Inches(0.0), Inches(0.16), Inches(0.16), Inches(1.34), NAVY)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.30), Inches(11.8), Inches(0.35))
    add_para(tf, kicker, 13, BLUE, bold=True, first=True, space_after=0)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.62), Inches(12.3), Inches(0.85))
    add_para(tf, title, 26, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.0)
    tb, tf = textbox(slide, Inches(12.4), Inches(7.02), Inches(0.8), Inches(0.35),
                     align=PP_ALIGN.RIGHT)
    add_para(tf, str(page_no), 11, GREY, first=True, space_after=0,
             align=PP_ALIGN.RIGHT)
    rect(slide, Inches(0.55), Inches(1.52), Inches(12.25), Pt(2), LIGHT)


def bullet(tf, text, size=15, color=INK, bold=False, first=False,
           space_after=8, indent=False, marker="●", marker_color=BLUE,
           line_spacing=1.15):
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
    round_rect(slide, x, y, w, h, fill)
    rect(slide, x, y, w, Inches(0.12), bar_color)
    tb, tf = textbox(slide, x + Inches(0.25), y + Inches(0.30),
                     w - Inches(0.5), h - Inches(0.5))
    if head:
        add_para(tf, head, 15, head_color or bar_color, bold=True, first=True,
                 space_after=7)
    return tf


def arrow(slide, x, y, w, h, color, shape=MSO_SHAPE.RIGHT_ARROW):
    ar = slide.shapes.add_shape(shape, x, y, w, h)
    ar.fill.solid(); ar.fill.fore_color.rgb = color
    ar.line.fill.background(); ar.shadow.inherit = False
    return ar


# ============================================================================
# 1. 封面
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, NAVY)
grad_bar(s, 0, 0, EMU_W, Inches(0.20))
grad_bar(s, 0, Inches(7.30), EMU_W, Inches(0.20))

# 装饰：右侧同心环（衍射 / PSF 意象）
for i, rr in enumerate([3.4, 2.7, 2.0, 1.35, 0.8]):
    cc = grad_color(i / 4)
    ov = s.shapes.add_shape(MSO_SHAPE.OVAL,
                            Inches(10.4 - rr), Inches(3.75 - rr),
                            Inches(rr * 2), Inches(rr * 2))
    ov.fill.background()
    ov.line.color.rgb = cc; ov.line.width = Pt(1.6)
    ov.shadow.inherit = False

tb, tf = textbox(s, Inches(0.9), Inches(1.2), Inches(9.0), Inches(0.5))
add_para(tf, "《光学原理》 · 课程文献汇报", 18, RGBColor(0xB9, 0xC4, 0xE8),
         bold=True, first=True, space_after=0)

tb, tf = textbox(s, Inches(0.9), Inches(2.0), Inches(9.4), Inches(2.4))
add_para(tf, "像差的衍射理论", 46, WHITE, bold=True, first=True,
         space_after=6, line_spacing=1.0)
add_para(tf, "The Diffraction Theory of Aberrations", 20,
         RGBColor(0x9F, 0xB4, 0xE8), bold=True, first=False, space_after=12)
add_para(tf, "从波像差函数、泽尼克多项式与点扩散函数，\n到 AI 驱动的波前传感与像差校正",
         17, RGBColor(0xC9, 0xD2, 0xEC), first=False, space_after=0,
         line_spacing=1.25)

tb, tf = textbox(s, Inches(0.9), Inches(5.05), Inches(9.6), Inches(1.6))
add_para(tf, "知识点出处：像差的衍射理论（波像差 / 泽尼克 / PSF / 斯特列尔比 / 容差）",
         14, RGBColor(0xAE, 0xB8, 0xDB), first=True, space_after=8)
add_para(tf, "文献一  从 PSF 直接预测泽尼克像差系数（深度学习）｜ arXiv, 2024", 14,
         WHITE, first=False, space_after=6)
add_para(tf, "文献二  物理信息图神经网络的频域像差校正 ｜ arXiv, 2025", 14,
         WHITE, first=False, space_after=0)

tb, tf = textbox(s, Inches(0.9), Inches(6.7), Inches(9), Inches(0.4))
add_para(tf, "汇报人：________     日期：________", 13,
         RGBColor(0x8F, 0x9B, 0xC4), first=True, space_after=0)


# ============================================================================
# 2. 目录
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "OUTLINE", "汇报提纲", 2)

items = [
    ("01", "课本知识点回顾", "波像差函数、泽尼克多项式、PSF、斯特列尔比与容差理论", VIOLET),
    ("02", "选题逻辑：正问题 vs 反问题", "课本算“像差→衍射像”；前沿反推“衍射像→像差”", BLUE),
    ("03", "文献一：从 PSF 预测泽尼克系数", "深度学习直接读出像差——“感知像差”", TEAL),
    ("04", "文献二：物理信息网络校正像差", "把衍射-像差物理模型嵌入网络——“校正像差”", CYAN),
    ("05", "对比 · 总结 · 启示", "经典衍射理论如何成为 AI 时代的语言与先验", NAVY),
]
y = Inches(1.95)
for num, t, d, c in items:
    round_rect(s, Inches(0.7), y, Inches(0.95), Inches(0.85), c)
    tb, tf = textbox(s, Inches(0.7), y, Inches(0.95), Inches(0.85),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, num, 22, WHITE, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=0)
    tb, tf = textbox(s, Inches(1.85), y + Inches(0.04), Inches(10.7), Inches(0.8),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, t, 18, NAVY, bold=True, first=True, space_after=2)
    add_para(tf, d, 13, GREY, first=False, space_after=0)
    y += Inches(0.98)


# ============================================================================
# 3. 课本知识点回顾
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "01 · 课本知识点回顾", "像差的衍射理论：波像差如何决定衍射像", 3)

# 左：核心概念
tf = card(s, Inches(0.55), Inches(1.8), Inches(6.05), Inches(5.0),
          LIGHT, VIOLET, "教材中的核心框架")
bullet(tf, "波像差函数 W(ρ,θ)：实际波前相对理想球面波前的光程差", 13.5,
       marker_color=VIOLET)
bullet(tf, "衍射像 = 出瞳函数的夫琅禾费衍射（衍射积分）", 13.5,
       marker_color=BLUE)
bullet(tf, "     PSF = | F { P(ρ,θ)·e^( i k W ) } |²", 13, color=NAVY,
       bold=True, marker="–", marker_color=BLUE, indent=True)
bullet(tf, "泽尼克多项式 Zₙᵐ：单位圆上正交基，分解像差（离焦/像散/彗差/球差…）",
       13.5, marker_color=TEAL)
bullet(tf, "斯特列尔比 S：像差使焦点峰值强度下降的程度", 13.5,
       marker_color=AMBER)
bullet(tf, "     S ≈ 1 − (2π/λ)²·σ_W²  （马雷夏尔近似）", 13, color=NAVY,
       bold=True, marker="–", marker_color=AMBER, indent=True)
bullet(tf, "容差理论：瑞利 λ/4 判据；马雷夏尔判据 σ_W ≤ λ/14 → S ≳ 0.8",
       13.5, marker_color=RED)

# 右上：正问题流程示意（出瞳 → 衍射 → PSF）
round_rect(s, Inches(6.8), Inches(1.8), Inches(5.95), Inches(2.5), CARD_BLUE)
tb, tf = textbox(s, Inches(7.05), Inches(1.95), Inches(5.5), Inches(0.4))
add_para(tf, "正问题：已知像差 W → 算出衍射像 PSF", 14, NAVY, bold=True,
         first=True, space_after=0)
# 出瞳（带像差的圆）
ov = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.15), Inches(2.65),
                        Inches(1.25), Inches(1.25))
ov.fill.solid(); ov.fill.fore_color.rgb = RGBColor(0xCE, 0xD9, 0xF2)
ov.line.color.rgb = NAVY; ov.line.width = Pt(1.2); ov.shadow.inherit = False
tb, tf = textbox(s, Inches(7.15), Inches(3.95), Inches(1.25), Inches(0.3),
                 align=PP_ALIGN.CENTER)
add_para(tf, "出瞳 P·e^(ikW)", 9.5, GREY, first=True, align=PP_ALIGN.CENTER,
         space_after=0)
arrow(s, Inches(8.6), Inches(3.05), Inches(1.35), Inches(0.45), TEAL)
tb, tf = textbox(s, Inches(8.55), Inches(2.66), Inches(1.5), Inches(0.3),
                 align=PP_ALIGN.CENTER)
add_para(tf, "衍射", 10, TEAL, bold=True, first=True, align=PP_ALIGN.CENTER,
         space_after=0)
# PSF（同心环）
for i, rr in enumerate([0.62, 0.42, 0.24]):
    cx, cy = Inches(10.95), Inches(3.28)
    o2 = s.shapes.add_shape(MSO_SHAPE.OVAL, cx - Inches(rr), cy - Inches(rr),
                            Inches(rr * 2), Inches(rr * 2))
    o2.fill.background(); o2.line.color.rgb = grad_color(i / 2)
    o2.line.width = Pt(1.4); o2.shadow.inherit = False
tb, tf = textbox(s, Inches(10.2), Inches(3.95), Inches(1.6), Inches(0.3),
                 align=PP_ALIGN.CENTER)
add_para(tf, "PSF 衍射像", 9.5, GREY, first=True, align=PP_ALIGN.CENTER,
         space_after=0)

# 右下：为什么找新文献
tf = card(s, Inches(6.8), Inches(4.45), Inches(5.95), Inches(2.35),
          CARD_AMBR, AMBER, "教材“老”在哪里？")
bullet(tf, "课本以解析推导“正问题”为主：给定像差，预测衍射像与分辨率下降",
       13, marker_color=AMBER)
bullet(tf, "现实更想解“反问题”：从观测到的衍射像/模糊像，反推并校正像差",
       13, marker_color=AMBER, bold=True)
bullet(tf, "近两年 AI 让这一反演又快又准——下面两篇文献正是代表",
       13, marker_color=AMBER)


# ============================================================================
# 4. 选题逻辑：正问题 vs 反问题
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "02 · 选题逻辑", "同一套衍射理论：正问题（课本） vs 反问题（前沿）", 4)

tb, tf = textbox(s, Inches(0.7), Inches(1.72), Inches(12), Inches(0.6))
add_para(tf, "像差的衍射理论给出“像差 W → 衍射像 PSF”的正向映射；现代研究关心它的反演——从衍射像反推、并校正像差。",
         15.5, INK, first=True, space_after=0, line_spacing=1.2)

# 正问题条
round_rect(s, Inches(0.7), Inches(2.5), Inches(12.0), Inches(1.35), CARD_BLUE)
rect(s, Inches(0.7), Inches(2.5), Inches(0.14), Inches(1.35), VIOLET)
tb, tf = textbox(s, Inches(1.0), Inches(2.62), Inches(2.4), Inches(1.1),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "正问题\n（课本）", 16, VIOLET, bold=True, first=True,
         space_after=0, line_spacing=1.05)
b1 = round_rect(s, Inches(3.5), Inches(2.78), Inches(2.7), Inches(0.8), WHITE)
tb, tf = textbox(s, Inches(3.5), Inches(2.78), Inches(2.7), Inches(0.8),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_para(tf, "像差 W\n(泽尼克系数)", 13, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line_spacing=1.0)
arrow(s, Inches(6.4), Inches(2.95), Inches(1.4), Inches(0.46), BLUE)
tb, tf = textbox(s, Inches(6.3), Inches(2.55), Inches(1.6), Inches(0.3),
                 align=PP_ALIGN.CENTER)
add_para(tf, "衍射积分", 11, BLUE, bold=True, first=True, align=PP_ALIGN.CENTER,
         space_after=0)
b2 = round_rect(s, Inches(8.0), Inches(2.78), Inches(2.7), Inches(0.8), WHITE)
tb, tf = textbox(s, Inches(8.0), Inches(2.78), Inches(2.7), Inches(0.8),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_para(tf, "衍射像 PSF\n(分辨率下降)", 13, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0, line_spacing=1.0)

# 反问题条
round_rect(s, Inches(0.7), Inches(4.0), Inches(12.0), Inches(1.35), CARD_TEAL)
rect(s, Inches(0.7), Inches(4.0), Inches(0.14), Inches(1.35), TEAL)
tb, tf = textbox(s, Inches(1.0), Inches(4.12), Inches(2.4), Inches(1.1),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "反问题\n（前沿）", 16, TEAL, bold=True, first=True,
         space_after=0, line_spacing=1.05)
round_rect(s, Inches(3.5), Inches(4.28), Inches(2.7), Inches(0.8), WHITE)
tb, tf = textbox(s, Inches(3.5), Inches(4.28), Inches(2.7), Inches(0.8),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_para(tf, "衍射像 / 模糊像", 13, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
arrow(s, Inches(6.4), Inches(4.45), Inches(1.4), Inches(0.46), TEAL)
tb, tf = textbox(s, Inches(6.25), Inches(4.05), Inches(1.7), Inches(0.3),
                 align=PP_ALIGN.CENTER)
add_para(tf, "AI 反演/校正", 11, TEAL, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
round_rect(s, Inches(8.0), Inches(4.28), Inches(2.7), Inches(0.8), WHITE)
tb, tf = textbox(s, Inches(8.0), Inches(4.28), Inches(2.7), Inches(0.8),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_para(tf, "像差 W → 校正", 13, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=0)

# 两篇文献定位
tf = card(s, Inches(0.7), Inches(5.55), Inches(5.9), Inches(1.25),
          CARD_VIO, VIOLET, "文献一 · 感知像差")
bullet(tf, "用 CNN 从 PSF 直接读出泽尼克系数（反演正问题）", 13,
       marker_color=VIOLET)
tf = card(s, Inches(6.8), Inches(5.55), Inches(5.9), Inches(1.25),
          CARD_TEAL, TEAL, "文献二 · 校正像差")
bullet(tf, "把衍射-像差物理模型作先验，校正显微像差（频域）", 13,
       marker_color=TEAL)


# ============================================================================
# 文献一
# ============================================================================
# 5. 背景
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 背景与问题", "从衍射像直接“读出”像差：感知像差", 5)

round_rect(s, Inches(0.7), Inches(1.75), Inches(12.0), Inches(1.05), NAVY)
tb, tf = textbox(s, Inches(0.95), Inches(1.86), Inches(11.5), Inches(0.85),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "Direct Zernike Coefficient Prediction from Point Spread Functions and Extended Images using Deep Learning",
         15, WHITE, bold=True, first=True, space_after=3, line_spacing=1.05)
add_para(tf, "arXiv:2404.15231 (2024) · 关键词：泽尼克系数 / PSF / 相位分集 / 卷积神经网络 / 自适应光学",
         12, RGBColor(0xC9, 0xD2, 0xEC), first=False, space_after=0)

tf = card(s, Inches(0.7), Inches(3.05), Inches(5.85), Inches(3.65),
          LIGHT, VIOLET, "问题背景")
bullet(tf, "像差会严重退化成像质量，需要先“测出”像差才能校正", 14,
       marker_color=VIOLET)
bullet(tf, "传统自适应光学多靠迭代搜索来逼近最佳校正，速度慢、易陷局部最优",
       14, marker_color=VIOLET)
bullet(tf, "波前传感器件复杂；能否直接“看图识像差”？", 14,
       marker_color=VIOLET)

tf = card(s, Inches(6.8), Inches(3.05), Inches(5.85), Inches(3.65),
          CARD_AMBR, AMBER, "核心思路")
bullet(tf, "把“PSF → 泽尼克像差系数”当作一个回归/反演问题", 14,
       marker_color=AMBER)
bullet(tf, "用卷积神经网络(CNN)直接预测泽尼克系数，取代迭代搜索", 14,
       marker_color=AMBER)
bullet(tf, "借助“相位分集”：拍焦面上方/焦面/下方多张图，提供深度信息",
       14, marker_color=AMBER, bold=True)


# 6. 方法
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 方法与创新", "CNN + 相位分集，直接回归泽尼克系数", 6)

tf = card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.95),
          CARD_VIO, VIOLET, "核心方法")
bullet(tf, "输入：2–3 张相位分集图像（焦面下方 / 焦面 / 上方）", 14,
       marker_color=VIOLET)
bullet(tf, "网络：卷积神经网络 (CNN) 提取特征", 14, marker_color=VIOLET)
bullet(tf, "输出：直接回归出一组泽尼克系数 → 即完整波像差 W", 14,
       marker_color=VIOLET)
bullet(tf, "无需迭代搜索，一次前向推理即得像差估计", 14, marker_color=VIOLET,
       bold=True)
bullet(tf, "对“点源 PSF”和“扩展物体图像”都适用", 14, marker_color=VIOLET)

# 右：相位分集示意
tf = card(s, Inches(6.8), Inches(1.85), Inches(5.9), Inches(2.55),
          LIGHT, BLUE, "为什么要“相位分集”？")
bullet(tf, "只用单张焦面 PSF，正负离焦等情形会混淆（解不唯一）", 13,
       marker_color=BLUE)
bullet(tf, "额外加入已知离焦量（离焦本身就是一个泽尼克模式 Z₂⁰）", 13,
       marker_color=BLUE)
bullet(tf, "多张分集图像 → 反演更稳定、更准确", 13, marker_color=BLUE, bold=True)

# 右下：三张分集图标
labels = ["焦面下方", "焦面", "焦面上方"]
for i, lab in enumerate(labels):
    bx = Inches(6.95 + i * 1.97)
    round_rect(s, bx, Inches(4.6), Inches(1.75), Inches(1.4), CARD_BLUE)
    cx, cy = bx + Inches(0.875), Inches(5.05)
    rr = [0.5, 0.28, 0.5][i]
    o = s.shapes.add_shape(MSO_SHAPE.OVAL, cx - Inches(rr), cy - Inches(rr),
                           Inches(rr * 2), Inches(rr * 2))
    o.fill.solid(); o.fill.fore_color.rgb = grad_color(i / 2)
    o.line.fill.background(); o.shadow.inherit = False
    tb, tf2 = textbox(s, bx, Inches(5.62), Inches(1.75), Inches(0.32),
                      align=PP_ALIGN.CENTER)
    add_para(tf2, lab, 11, NAVY, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=0)


# 7. 结果
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 结果", "三张分集图像即可高精度反演像差", 7)

stats = [
    ("0.10 rad", "波前 RMSE", "3 张分集 PSF（振幅 1）在仿真集上的均方根误差", VIOLET),
    ("2–3 张", "所需图像", "仅用焦面附近少量图像即可完成反演", BLUE),
    ("一次推理", "无需迭代", "CNN 前向一次即出像差，远快于迭代搜索", TEAL),
]
x = Inches(0.7)
for big, small, desc, c in stats:
    round_rect(s, x, Inches(1.9), Inches(3.9), Inches(2.0), LIGHT)
    rect(s, x, Inches(1.9), Inches(3.9), Inches(0.12), c)
    tb, tf = textbox(s, x + Inches(0.2), Inches(2.15), Inches(3.5), Inches(1.65),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, big, 28, NAVY, bold=True, first=True, align=PP_ALIGN.CENTER,
             space_after=0)
    add_para(tf, small, 14, c, bold=True, first=False, align=PP_ALIGN.CENTER,
             space_after=4)
    add_para(tf, desc, 11.5, GREY, first=False, align=PP_ALIGN.CENTER,
             space_after=0, line_spacing=1.05)
    x += Inches(4.05)

tf = card(s, Inches(0.7), Inches(4.25), Inches(12.0), Inches(2.55),
          CARD_VIO, VIOLET, "结论与意义")
bullet(tf, "用焦面上方/焦面/下方 3 张相位分集图像，在仿真 PSF 数据集上达到约 0.10 弧度的低波前 RMSE",
       15, marker_color=VIOLET)
bullet(tf, "直接回归泽尼克系数，省去传统自适应光学的迭代搜索，更快、更直接",
       15, marker_color=VIOLET)
bullet(tf, "对点源与扩展物体均可用，为快速波前传感、实时像差校正提供新路径",
       15, marker_color=VIOLET, bold=True)


# 8. 与课本联系
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "03 · 文献一 / 与课本知识点的联系", "课本概念 → 前沿研究", 8)

pairs = [
    ("PSF = |F{P·e^(ikW)}|²", "网络学习的正是这条映射的“反函数”：PSF → W", VIOLET),
    ("泽尼克多项式 Zₙᵐ", "网络的输出就是泽尼克系数，沿用课本的像差正交基", BLUE),
    ("离焦是一种像差 (Z₂⁰)", "相位分集 = 人为加入已知离焦像差，帮助反演", TEAL),
    ("波前误差 / 容差理论", "0.10 rad 的 RMSE 直接对应斯特列尔比与马雷夏尔判据", AMBER),
]
y = Inches(1.95)
for kbook, research, c in pairs:
    round_rect(s, Inches(0.7), y, Inches(4.9), Inches(1.05), LIGHT)
    rect(s, Inches(0.7), y, Inches(0.14), Inches(1.05), c)
    tb, tf = textbox(s, Inches(1.0), y, Inches(4.5), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, kbook, 14.5, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.05)
    arrow(s, Inches(5.7), y + Inches(0.32), Inches(0.8), Inches(0.4), c)
    round_rect(s, Inches(6.65), y, Inches(6.05), Inches(1.05), CARD_BLUE)
    tb, tf = textbox(s, Inches(6.9), y, Inches(5.6), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, research, 14, INK, first=True, space_after=0, line_spacing=1.1)
    y += Inches(1.2)


# ============================================================================
# 文献二
# ============================================================================
# 9. 背景
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 背景与问题", "物理 + AI：在显微成像中校正像差", 9)

round_rect(s, Inches(0.7), Inches(1.75), Inches(12.0), Inches(1.05), TEAL)
tb, tf = textbox(s, Inches(0.95), Inches(1.86), Inches(11.5), Inches(0.85),
                 anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "Physics-Informed Graph Neural Networks for Frequency-Aware Optical Aberration Correction",
         15, WHITE, bold=True, first=True, space_after=3, line_spacing=1.05)
add_para(tf, "arXiv:2512.05683 (2025) · 关键词：像差校正 / 泽尼克 / 物理信息神经网络 / 频域(OTF) / 显微成像",
         12, RGBColor(0xDF, 0xF3, 0xF1), first=False, space_after=0)

tf = card(s, Inches(0.7), Inches(3.05), Inches(5.85), Inches(3.65),
          LIGHT, TEAL, "问题背景")
bullet(tf, "光学像差会显著退化显微成像质量，尤其在“往样品深处成像”时更严重",
       13.5, marker_color=TEAL)
bullet(tf, "像差来自光学波前的畸变，可用泽尼克多项式数学描述", 13.5,
       marker_color=TEAL)
bullet(tf, "纯数据驱动的校正常缺乏物理约束，泛化差、易过拟合", 13.5,
       marker_color=TEAL)

tf = card(s, Inches(6.8), Inches(3.05), Inches(5.85), Inches(3.65),
          CARD_AMBR, AMBER, "核心思路")
bullet(tf, "物理信息(Physics-Informed)：把衍射-像差的物理前向模型嵌入网络作先验",
       13.5, marker_color=AMBER)
bullet(tf, "频域感知(Frequency-Aware)：不同空间频率受像差影响不同（即 OTF 视角）",
       13.5, marker_color=AMBER)
bullet(tf, "图神经网络(GNN)：建模各模式/频段间的关联，更有效地校正像差",
       13.5, marker_color=AMBER, bold=True)


# 10. 方法
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 方法与创新", "把经典衍射理论“写进”网络结构", 10)

tf = card(s, Inches(0.7), Inches(1.85), Inches(5.85), Inches(4.95),
          CARD_TEAL, TEAL, "核心创新点")
bullet(tf, "物理信息先验：以泽尼克展开 + 像差→PSF/OTF 前向模型约束网络输出",
       14, marker_color=TEAL)
bullet(tf, "频域感知：显式区分不同空间频率成分的退化（高频受像差影响更大）",
       14, marker_color=TEAL)
bullet(tf, "图神经网络：把像差模式 / 频段建成图节点，建模其相互关系", 14,
       marker_color=TEAL)
bullet(tf, "面向深层显微成像，提升校正的准确性与鲁棒性", 14, marker_color=TEAL,
       bold=True)

# 右：频域(OTF)视角示意
tf = card(s, Inches(6.8), Inches(1.85), Inches(5.9), Inches(2.45),
          LIGHT, BLUE, "频域(OTF)视角：像差为何降分辨率")
bullet(tf, "OTF = 出瞳函数的自相关；像差使 OTF 在高频处下降", 13,
       marker_color=BLUE)
bullet(tf, "高频 = 细节；高频被压低 → 图像变模糊、分辨率下降", 13,
       marker_color=BLUE)
bullet(tf, "“频域感知”正是针对不同频率分别校正", 13, marker_color=BLUE,
       bold=True)

# 右下：物理+AI 融合
round_rect(s, Inches(6.8), Inches(4.5), Inches(5.9), Inches(2.3), CARD_VIO)
tb, tf2 = textbox(s, Inches(7.05), Inches(4.68), Inches(5.4), Inches(2.0),
                  anchor=MSO_ANCHOR.MIDDLE)
add_para(tf2, "物理模型  +  神经网络", 17, NAVY, bold=True, first=True,
         align=PP_ALIGN.CENTER, space_after=8)
add_para(tf2, "（衍射理论/泽尼克）   （图神经网络）", 12, GREY, first=False,
         align=PP_ALIGN.CENTER, space_after=8)
add_para(tf2, "= 物理信息网络：更准、更稳、更省数据", 14, TEAL, bold=True,
         first=False, align=PP_ALIGN.CENTER, space_after=0)


# 11. 结果与应用
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 结果与应用", "物理信息 + 频域感知带来更优校正", 11)

tf = card(s, Inches(0.7), Inches(1.9), Inches(5.95), Inches(4.9),
          CARD_TEAL, TEAL, "主要结果（定性）")
bullet(tf, "相较纯数据驱动方法，引入物理先验后像差校正更准确、更稳定", 14,
       marker_color=TEAL)
bullet(tf, "频域感知设计能更好恢复被像差压低的高频细节 → 分辨率回升", 14,
       marker_color=TEAL)
bullet(tf, "在“深层成像”这类强像差场景下优势更明显", 14, marker_color=TEAL)
bullet(tf, "物理约束降低对大规模标注数据的依赖，泛化性更好", 14,
       marker_color=TEAL, bold=True)

tf = card(s, Inches(6.85), Inches(1.9), Inches(5.85), Inches(4.9),
          LIGHT, BLUE, "应用场景")
bullet(tf, "生物显微 / 活体深层成像：穿过组织时的像差校正", 14,
       marker_color=BLUE)
bullet(tf, "超分辨显微：去除像差以逼近衍射极限分辨率", 14, marker_color=BLUE)
bullet(tf, "自适应光学的“数字/离线”版本：先估计再后处理校正", 14,
       marker_color=BLUE)
bullet(tf, "天文、视觉科学等同样存在波前像差的成像系统", 14,
       marker_color=BLUE)
bullet(tf, "为“物理 + AI”融合的计算成像提供了范式", 14, marker_color=BLUE,
       bold=True)


# 12. 与课本联系
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "04 · 文献二 / 与课本知识点的联系", "课本概念 → 前沿研究", 12)

pairs2 = [
    ("泽尼克多项式展开", "作为网络的“物理先验”，约束像差的表示", VIOLET),
    ("光学传递函数 OTF", "“频域感知”= 针对 OTF 在不同频率的衰减分别校正", BLUE),
    ("像差 → 分辨率下降", "高频被像差压低 → 模糊；校正即恢复高频细节", TEAL),
    ("衍射理论的前向模型", "physics-informed = 把 W→PSF/OTF 模型写进网络", AMBER),
]
y = Inches(1.95)
for kbook, research, c in pairs2:
    round_rect(s, Inches(0.7), y, Inches(4.9), Inches(1.05), LIGHT)
    rect(s, Inches(0.7), y, Inches(0.14), Inches(1.05), c)
    tb, tf = textbox(s, Inches(1.0), y, Inches(4.5), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, kbook, 14.5, NAVY, bold=True, first=True, space_after=0,
             line_spacing=1.05)
    arrow(s, Inches(5.7), y + Inches(0.32), Inches(0.8), Inches(0.4), c)
    round_rect(s, Inches(6.65), y, Inches(6.05), Inches(1.05), CARD_TEAL)
    tb, tf = textbox(s, Inches(6.9), y, Inches(5.6), Inches(1.05),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, research, 14, INK, first=True, space_after=0, line_spacing=1.1)
    y += Inches(1.2)


# ============================================================================
# 13. 对比总结
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "05 · 对比与总结", "两篇文献：感知像差 与 校正像差", 13)

headers = ["对比维度", "文献一 · 感知像差", "文献二 · 校正像差"]
colw = [Inches(2.9), Inches(4.9), Inches(4.9)]
xs = [Inches(0.7), Inches(3.6), Inches(8.5)]
hy = Inches(1.85)
hcolors = [NAVY, VIOLET, TEAL]
for i in range(3):
    rect(s, xs[i], hy, colw[i], Inches(0.6), hcolors[i])
    tb, tf = textbox(s, xs[i], hy, colw[i], Inches(0.6),
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_para(tf, headers[i], 14, WHITE, bold=True, first=True,
             align=PP_ALIGN.CENTER, space_after=0)

rows = [
    ("任务", "从 PSF 反推泽尼克像差", "校正显微图像中的像差"),
    ("对衍射理论的用法", "反演正问题 PSF → W", "把 PSF/OTF 物理模型作先验"),
    ("主要方法", "CNN + 相位分集", "物理信息图神经网络 + 频域感知"),
    ("工作域", "空域（PSF / 图像）", "频域（OTF / 空间频率）"),
    ("典型应用", "快速波前传感 / 自适应光学", "深层显微 / 超分辨 / 数字AO"),
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
        add_para(tf, row[i], 12.5, NAVY if i == 0 else INK,
                 bold=(i == 0), first=True,
                 align=PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT,
                 space_after=0, line_spacing=1.05)
    ry += rh
rect(s, Inches(0.7), hy, Inches(12.0), Pt(2.5), NAVY)


# ============================================================================
# 14. 总结与启示
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, NAVY)
grad_bar(s, 0, 0, EMU_W, Inches(0.16))
grad_bar(s, 0, Inches(7.34), EMU_W, Inches(0.16))

tb, tf = textbox(s, Inches(0.9), Inches(0.7), Inches(11.5), Inches(0.7))
add_para(tf, "总结与启示", 30, WHITE, bold=True, first=True, space_after=0)

round_rect(s, Inches(0.9), Inches(1.7), Inches(11.55), Inches(4.7),
           RGBColor(0x1E, 0x26, 0x52))
tb, tf = textbox(s, Inches(1.3), Inches(2.0), Inches(10.8), Inches(4.2))
add_para(tf, "1.  像差的衍射理论是“正问题”，前沿热点是它的“反问题”", 18,
         grad_color(1.0), bold=True, first=True, space_after=4)
add_para(tf, "     课本推导 像差 W → 衍射像 PSF；两篇文献则从衍射像反推与校正像差，互为表里。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=14,
         line_spacing=1.2)
add_para(tf, "2.  泽尼克 / PSF / OTF / 斯特列尔比，是 AI 时代的“通用语言”", 18,
         grad_color(0.55), bold=True, first=False, space_after=4)
add_para(tf, "     现代算法的输入输出与先验，依然建立在这些经典衍射理论工具之上。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=14,
         line_spacing=1.2)
add_para(tf, "3.  “物理 + AI”是趋势：把衍射理论写进网络", 18,
         grad_color(0.0), bold=True, first=False, space_after=4)
add_para(tf, "     文献一用数据反演、文献二用物理先验约束，越懂衍射理论，越能把 AI 用准、用稳。",
         15, RGBColor(0xD4, 0xDB, 0xF0), first=False, space_after=0,
         line_spacing=1.2)

tb, tf = textbox(s, Inches(0.9), Inches(6.55), Inches(11.5), Inches(0.6))
add_para(tf, "一句话：读懂像差的衍射理论，才能在 AI 时代既“看见”像差，又“消除”像差。",
         16, grad_color(0.8), bold=True, first=True, space_after=0)


# ============================================================================
# 15. 参考文献
# ============================================================================
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, WHITE)
title_bar(s, "REFERENCES", "参考文献", 15)

refs = [
    ("[1]  文献一 · 感知像差",
     "Direct Zernike Coefficient Prediction from Point Spread Functions and "
     "Extended Images using Deep Learning. arXiv:2404.15231 (2024).",
     "https://arxiv.org/abs/2404.15231", VIOLET),
    ("[2]  文献二 · 校正像差",
     "Physics-Informed Graph Neural Networks for Frequency-Aware Optical "
     "Aberration Correction. arXiv:2512.05683 (2025).",
     "https://arxiv.org/abs/2512.05683", TEAL),
    ("[3]  课本知识点参考",
     "M. Born & E. Wolf, Principles of Optics（《光学原理》）, Ch. 9 "
     "“The Diffraction Theory of Aberrations”：波像差函数、泽尼克多项式、"
     "PSF、斯特列尔比与容差理论。",
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
add_para(tf, "注：文献一、二均为预印本(arXiv)，可按上方链接查阅原文与图表；引用时建议核对最新发表版本。",
         11.5, GREY, first=True, space_after=0)


# ----------------------------------------------------------------------------
OUT = "光学原理_像差的衍射理论_文献汇报.pptx"
prs.save(OUT)
print("Saved:", OUT)
print("Total slides:", len(prs.slides._sldIdLst))
