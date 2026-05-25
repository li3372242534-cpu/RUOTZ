# -*- coding: utf-8 -*-
# PPT Generator - Social Factors in Scientific Development
import sys, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

# Colors
C1 = RGBColor(0x1F,0x3A,0x68)  # deep blue
C2 = RGBColor(0xFF,0x8C,0x42)  # orange
C3 = RGBColor(0x2C,0x2C,0x2C)  # dark
C4 = RGBColor(0xFF,0xFF,0xFF)  # white
C5 = RGBColor(0xE8,0xEF,0xF7)  # soft blue
C6 = RGBColor(0x2E,0xA8,0x7E)  # green
C7 = RGBColor(0xD9,0x3B,0x3B)  # red
C8 = RGBColor(0x9C,0xA3,0xAF)  # gray
C9 = RGBColor(0x6B,0x72,0x80)  # mid gray
C10 = RGBColor(0x17,0x2C,0x50) # darker blue
C11 = RGBColor(0xCC,0xD6,0xE6) # light text
C12 = RGBColor(0xFD,0xED,0xED) # light red bg
FN = '\u5fae\u8f6f\u96c5\u9ed1'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BL = prs.slide_layouts[6]


def bg(s, c):
    b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    b.line.fill.background(); b.fill.solid(); b.fill.fore_color.rgb = c
    sp = b._element.getparent(); sp.remove(b._element); sp.insert(2, b._element)

def T(s, l, t, w, h, tx, sz=14, b=False, c=C3, a=PP_ALIGN.LEFT, v=MSO_ANCHOR.TOP, f=None):
    if f is None: f = FN
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = v
    for i, ln in enumerate(tx.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = a; r = p.add_run(); r.text = ln
        r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c; r.font.name = f
        rPr = r._r.get_or_add_rPr()
        rF = rPr.find(qn('a:rFonts'))
        if rF is None: rF = etree.SubElement(rPr, qn('a:rFonts'))
        rF.set('eastAsia', f)
    return tb

def R(s, l, t, w, h, fill=None, lc=None, sh=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(sh, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    else: sp.fill.background()
    if lc: sp.line.color.rgb = lc; sp.line.width = Pt(0.75)
    else: sp.line.fill.background()
    sp.shadow.inherit = False
    return sp

def TB(s, ti, sub=''):
    R(s, 0.4, 0.35, 0.15, 0.5, fill=C2)
    T(s, 0.65, 0.3, 11, 0.55, ti, sz=26, b=True, c=C1, v=MSO_ANCHOR.MIDDLE)
    if sub: T(s, 0.65, 0.88, 11, 0.3, sub, sz=11, c=C9)
    R(s, 0.4, 1.3, 12.5, 0.02, fill=C1)

def FT(s, n):
    T(s, 0.5, 7.05, 9, 0.3, '\u300a\u79d1\u6280\u89c2\u300b\u8bfe\u7a0b\u4f5c\u4e1a', sz=8, c=C8)
    T(s, 11.5, 7.05, 1.3, 0.3, f'{n}/15', sz=8, c=C8, a=PP_ALIGN.RIGHT)


# === P1 Cover ===
s = prs.slides.add_slide(BL)
bg(s, C1); R(s, 0, 0, 4.2, 7.5, fill=C10); R(s, 0, 5.0, 13.333, 0.06, fill=C2)
T(s, 0.7, 1.8, 12, 1.0, '\u5f71\u54cd\u79d1\u5b66\u53d1\u5c55\u7684\u793e\u4f1a\u56e0\u7d20', sz=44, b=True, c=C4)
T(s, 0.7, 2.8, 12, 0.8, '\u53ca\u4fc3\u8fdb\u6211\u56fd\u79d1\u6280\u8fdb\u6b65\u7684\u8def\u5f84\u601d\u8003', sz=36, b=True, c=C2)
T(s, 0.7, 4.0, 12, 0.4, 'Social Factors in Scientific Development & Pathways for China', sz=13, c=C11, f='Calibri')
T(s, 0.7, 5.4, 7, 0.35, '\u300a\u79d1\u6280\u89c2\u300b\u8bfe\u7a0b  |  \u4e2a\u4eba\u4f5c\u4e1a', sz=16, b=True, c=C4)
T(s, 0.7, 5.9, 11, 0.35, '\u59d3\u540d\uff1a____________   \u5b66\u53f7\uff1a____________   \u73ed\u7ea7\uff1a____________', sz=13, c=C11)
T(s, 0.7, 6.3, 11, 0.35, '\u6307\u5bfc\u6559\u5e08\uff1a____________   \u65e5\u671f\uff1a2026\u5e745\u6708', sz=13, c=C11)

# === P2 Intro ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u5f15\u8a00\uff1a\u79d1\u5b66\u53d1\u5c55\u79bb\u4e0d\u5f00\u793e\u4f1a\u571f\u58e4')
cards = [
    ('3.63\u4e07\u4ebf\u5143', '2024\u5e74\u4e2d\u56fdR&D\u7ecf\u8d39\u603b\u6295\u5165\n\u540c\u6bd4\u589e\u957f8.3%\uff0c\u5360GDP 2.68%\n\uff08\u6765\u6e90\uff1a\u56fd\u5bb6\u7edf\u8ba1\u5c40 2025.2\uff09'),
    ('\u5168\u7403\u7b2c11\u4f4d', '2024\u5e74\u5168\u7403\u521b\u65b0\u6307\u6570(GII)\u6392\u540d\n\u4e2d\u56fd10\u5e74\u5185\u4e0a\u5347\u6700\u5feb\u7ecf\u6d4e\u4f53\u4e4b\u4e00\n\uff08\u6765\u6e90\uff1aWIPO GII 2024\uff09'),
    ('\u8fde\u7eed5\u5e74\u7b2c1', '\u4e2d\u56fdPCT\u56fd\u9645\u4e13\u5229\u7533\u8bf7\u91cf\n2023\u5e74\u8fbe69,610\u4ef6\uff0c\u5360\u5168\u740325%\n\uff08\u6765\u6e90\uff1aWIPO 2024.3\uff09'),
]
for i,(n,d) in enumerate(cards):
    x = 0.6 + i*4.15
    R(s, x, 1.7, 3.95, 2.2, fill=C5, lc=C1)
    T(s, x+0.2, 1.85, 3.5, 0.6, n, sz=28, b=True, c=C2, a=PP_ALIGN.CENTER)
    T(s, x+0.2, 2.5, 3.5, 1.3, d, sz=12, c=C3, a=PP_ALIGN.CENTER)
R(s, 0.6, 4.2, 12.1, 1.8, fill=C1)
T(s, 0.9, 4.35, 11.5, 0.5, '\u6838\u5fc3\u95ee\u9898', sz=16, b=True, c=C2)
T(s, 0.9, 4.85, 11.5, 1.0,
  '\u8fd9\u4e9b\u6210\u5c31\u80cc\u540e\uff0c\u54ea\u4e9b\u793e\u4f1a\u56e0\u7d20\u5728\u8d77\u4f5c\u7528\uff1f\u8fd8\u6709\u54ea\u4e9b\u74f6\u9888\u5236\u7ea6\u7740\u4e2d\u56fd\u79d1\u6280\u7684\u8fdb\u4e00\u6b65\u7a81\u7834\uff1f\n'
  '\u672cPPT\u4ece\u653f\u6cbb\u3001\u7ecf\u6d4e\u3001\u6559\u80b2\u3001\u6587\u5316\u3001\u56fd\u9645\u73af\u5883\u4e94\u5927\u7ef4\u5ea6\u5206\u6790\u5f71\u54cd\u56e0\u7d20\uff0c\u5e76\u63d0\u51fa\u5bf9\u7b56\u5efa\u8bae\u3002', sz=14, c=C4)
T(s, 0.6, 6.3, 12.1, 0.5,
  '\u65f6\u95f4\u7ebf\uff1a\u4e24\u5f39\u4e00\u661f(1964)\u2192\u6742\u4ea4\u6c34\u7a3b(1973)\u2192\u8f7d\u4eba\u822a\u5929(2003)\u2192\u5317\u6597\u5168\u7403\u7ec4\u7f51(2020)\u2192\u5ae6\u5a25\u516d\u53f7(2024)\u2192DeepSeek-R1(2025)',
  sz=11, b=True, c=C1)
FT(s, 2)


# === P3 Overview ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u7b2c\u4e00\u90e8\u5206\uff1a\u5f71\u54cd\u79d1\u5b66\u53d1\u5c55\u7684\u4e94\u5927\u793e\u4f1a\u56e0\u7d20')
cx, cy = 6.5, 4.0
core = R(s, cx-0.8, cy-0.8, 1.6, 1.6, fill=C1, sh=MSO_SHAPE.OVAL)
tf = core.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = '\u79d1\u5b66\n\u53d1\u5c55'; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = C4; r.font.name = FN
factors = [('\u653f\u6cbb\u5236\u5ea6', 180), ('\u7ecf\u6d4e\u57fa\u7840', 252), ('\u6559\u80b2\u4f53\u7cfb', 324), ('\u6587\u5316\u4f20\u7edf', 36), ('\u56fd\u9645\u73af\u5883', 108)]
for name, deg in factors:
    rad = math.radians(deg)
    nx, ny = cx + 2.5*math.cos(rad), cy + 2.5*math.sin(rad)
    nd = R(s, nx-0.9, ny-0.45, 1.8, 0.9, fill=C2, sh=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = nd.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = name; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = C4; r.font.name = FN
T(s, 0.5, 6.6, 12.3, 0.4,
  '\u6b64\u5916\u8fd8\u5305\u62ec\uff1a\u6cd5\u5f8b\u5236\u5ea6\uff08\u77e5\u8bc6\u4ea7\u6743\uff09\u3001\u519b\u4e8b\u9700\u6c42\uff08ARPANET\u2192\u4e92\u8054\u7f51\uff09\u3001\u5b97\u6559\u4f26\u7406\u3001\u5a92\u4f53\u4f20\u64ad\u7b49',
  sz=11, c=C9)
FT(s, 3)


# === P4 Politics - Lysenko ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u56e0\u7d20\u2460 \u653f\u6cbb\u5236\u5ea6\u4e0e\u79d1\u6280\u653f\u7b56')
R(s, 0.5, 1.7, 5.8, 2.5, fill=C5)
T(s, 0.7, 1.8, 5.4, 0.35, '\u6b63\u9762\uff1a\u56fd\u5bb6\u6218\u7565\u9a71\u52a8\u91cd\u5927\u7a81\u7834', sz=14, b=True, c=C6)
T(s, 0.7, 2.2, 5.4, 1.8,
  '\u2022 \u4e24\u5f39\u4e00\u661f\uff081964-1970\uff09\uff1a\u4e3e\u56fd\u4f53\u5236\uff0c23\u4f4d\u5143\u52cb\u653b\u514b\u6838\u5f39\u5bfc\u5f39\u536b\u661f\n'
  '\u2022 \u5ae6\u5a25\u516d\u53f7\uff082024.6.25\uff09\uff1a\u4eba\u7c7b\u9996\u6b21\u6708\u7403\u80cc\u9762\u91c7\u6837\uff0c\u53d6\u56de1935.3\u514b\u6708\u58e4\n'
  '\u2022 \u5317\u6597\u4e09\u53f7\uff082020\uff09\uff1a\u5168\u7403\u7ec4\u7f51\uff0c\u6446\u8131\u5bf9GPS\u4f9d\u8d56\n'
  '\u2022 \u653f\u7b56\u5de5\u5177\uff1a\u56fd\u5bb6\u91cd\u70b9\u5b9e\u9a8c\u5ba4\u8d85700\u4e2a\uff0cR&D\u7a0e\u6536\u52a0\u8ba1\u6263\u9664100%\u219220%', sz=12, c=C3)

R(s, 6.5, 1.7, 6.3, 2.5, fill=C12)
T(s, 6.7, 1.8, 5.9, 0.35, '\u53cd\u9762\uff1a\u82cf\u8054\u674e\u68ee\u79d1\u4e8b\u4ef6\uff081935-1964\uff09', sz=14, b=True, c=C7)
T(s, 6.7, 2.2, 5.9, 1.8,
  '\u2022 \u674e\u68ee\u79d1\u5426\u5b9a\u5b5f\u5fb7\u5c14-\u6469\u5c14\u6839\u9057\u4f20\u5b66\uff0c\u9f13\u5439\u300c\u83b7\u5f97\u6027\u9057\u4f20\u300d\n'
  '\u2022 \u65af\u5927\u6797\u652f\u6301\u4e0b\uff0c1948\u5e74\u5168\u82cf\u519c\u79d1\u9662\u5ba3\u5e03\u9057\u4f20\u5b66\u4e3a\u300c\u4f2a\u79d1\u5b66\u300d\n'
  '\u2022 \u9057\u4f20\u5b66\u5bb6\u74e6\u7ef4\u6d1b\u592b\u88ab\u6355\u5165\u72f1\uff0c1943\u5e74\u6b7b\u4e8e\u72f1\u4e2d\n'
  '\u2022 3000\u591a\u540d\u9057\u4f20\u5b66\u5bb6\u88ab\u89e3\u804c\uff0c\u82cf\u8054\u9057\u4f20\u5b66\u5012\u900030\u5e74\n'
  '\u2022 \u6559\u8bad\uff1a\u653f\u6cbb\u6743\u529b\u4e0d\u5f53\u5e72\u9884\u79d1\u5b66\uff0c\u540e\u679c\u662f\u707e\u96be\u6027\u7684', sz=12, c=C3)

R(s, 0.5, 4.5, 12.3, 2.3, fill=C1)
T(s, 0.7, 4.65, 11.9, 0.35, '\u5bf9\u6bd4\u542f\u793a', sz=15, b=True, c=C2)
T(s, 0.7, 5.1, 11.9, 1.5,
  '\u653f\u6cbb\u7a33\u5b9a + \u6b63\u786e\u7684\u79d1\u6280\u6218\u7565 = \u96c6\u4e2d\u529b\u91cf\u529e\u5927\u4e8b\uff08\u4e24\u5f39\u4e00\u661f\u3001\u5ae6\u5a25\u5de5\u7a0b\uff09\n'
  '\u653f\u6cbb\u5e72\u9884\u79d1\u5b66\u81ea\u7531 = \u627c\u6740\u521b\u65b0\uff08\u674e\u68ee\u79d1\u4e8b\u4ef6\u3001\u6587\u9769\u5bf9\u79d1\u7814\u7684\u7834\u574f\uff09\n\n'
  '\u5173\u952e\u539f\u5219\uff1a\u653f\u5e9c\u5e94\u5f53\u7ba1\u65b9\u5411\u3001\u7ba1\u6295\u5165\u3001\u7ba1\u8bc4\u4ef7\uff0c\u4f46\u4e0d\u80fd\u7ba1\u5b66\u672f\u5224\u65ad', sz=13, c=C4)
FT(s, 4)


# === P5 Economy with chart ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u56e0\u7d20\u2461 \u7ecf\u6d4e\u53d1\u5c55\u6c34\u5e73\u4e0e\u7814\u53d1\u6295\u5165')
T(s, 0.5, 1.6, 12, 0.4,
  '2024\u5e74\u4e2d\u56fdR&D\u7ecf\u8d39\u8fbe3.63\u4e07\u4ebf\u5143\uff08\u7ea6$1.03\u4e07\u4ebf PPP\uff09\uff0c\u9996\u6b21\u8d85\u8d8a\u7f8e\u56fd\u6210\u4e3a\u5168\u7403\u6700\u5927\u7814\u53d1\u6295\u5165\u56fd\uff08ITIF, 2026.5\uff09',
  sz=13, b=True, c=C1)
cd = CategoryChartData()
cd.categories = ['\u4ee5\u8272\u5217', '\u97e9\u56fd', '\u7f8e\u56fd', '\u65e5\u672c', '\u5fb7\u56fd', '\u4e2d\u56fd', '\u6cd5\u56fd']
cd.add_series('R&D/GDP %', (5.56, 4.93, 3.40, 3.30, 3.13, 2.68, 2.20))
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(2.1), Inches(6.2), Inches(3.8), cd).chart
ch.has_legend = False
pl = ch.plots[0]; pl.has_data_labels = True
dl = pl.data_labels; dl.font.size = Pt(10); dl.font.bold = True; dl.font.color.rgb = C1
dl.position = XL_LABEL_POSITION.OUTSIDE_END
sr = ch.series[0]; sr.format.fill.solid(); sr.format.fill.fore_color.rgb = C1
# highlight China bar
dPt = etree.fromstring(
    '<c:dPt xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<c:idx val="5"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
    '<c:spPr><a:solidFill><a:srgbClr val="FF8C42"/></a:solidFill></c:spPr></c:dPt>')
cat_el = sr._element.find(qn('c:cat'))
if cat_el is not None: cat_el.addprevious(dPt)
T(s, 0.5, 5.95, 6.2, 0.3, '\u6570\u636e\u6765\u6e90\uff1aOECD MSTI 2024 / ITIF 2026', sz=9, c=C8, a=PP_ALIGN.CENTER)

T(s, 7.0, 2.1, 5.8, 0.4, '\u7ecf\u6d4e\u2192\u79d1\u6280\u7684\u4f20\u5bfc\u673a\u5236', sz=16, b=True, c=C1)
pts = [
    '\u7ecf\u6d4e\u5b9e\u529b\u51b3\u5b9a\u6295\u5165\u80fd\u529b\uff1a\u4e2d\u56fdR&D\u7ecf\u8d3920\u5e74\u589e\u957f\u8d8512\u500d',
    '\u5e02\u573a\u9700\u6c42\u7275\u5f15\u65b9\u5411\uff1a5G\u5546\u7528(2019)\u50ac\u751f\u4e07\u4ebf\u7ea7\u4ea7\u4e1a\u94fe',
    '\u4ea7\u4e1a\u5347\u7ea7\u5012\u903c\u57fa\u7840\u7814\u7a76\uff1a\u82af\u7247\u56f0\u5883\u2192\u56fd\u5bb6\u5927\u57fa\u91d1(3440\u4ebf\u5143)',
    '\u57fa\u7840\u7814\u7a76\u5360\u6bd4\u4ec56.65%\uff0c\u53d1\u8fbe\u56fd\u5bb615-25%\u2014\u2014\u77ed\u677f\u660e\u663e',
    '\u84b8\u6c7d\u673a\u6e90\u4e8e18\u4e16\u7eaa\u82f1\u56fd\u77ff\u5c71\u6392\u6c34\u9700\u6c42\uff0c\u7ecf\u6d4e\u9700\u6c42\u2192\u6280\u672f\u9769\u547d',
]
for i, p in enumerate(pts):
    y = 2.7 + i*0.72
    T(s, 7.0, y, 0.25, 0.3, '\u25b8', sz=13, b=True, c=C2)
    T(s, 7.3, y, 5.5, 0.65, p, sz=12, c=C3)
FT(s, 5)


# === P6 Education ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u56e0\u7d20\u2462 \u6559\u80b2\u4f53\u7cfb\u4e0e\u4eba\u624d\u57f9\u517b')
R(s, 0.5, 1.6, 12.3, 1.5, fill=C5)
T(s, 0.7, 1.7, 11.9, 0.35, '\u4e2d\u56fd\u9ad8\u7b49\u6559\u80b2\uff1a\u89c4\u6a21\u4e16\u754c\u7b2c\u4e00\uff0c\u4f46\u9876\u5c16\u521b\u65b0\u4eba\u624d\u4ecd\u4e0d\u8db3', sz=14, b=True, c=C1)
T(s, 0.7, 2.1, 11.9, 0.9,
  '\u2022 \u9ad8\u7b49\u6559\u80b2\u6bdb\u5165\u5b66\u7387 60.2%\uff082023\uff09\uff0c\u5728\u5b66\u603b\u89c4\u6a21 4763 \u4e07\u4eba\uff0c\u5168\u7403\u6700\u5927\n'
  '\u2022 2024\u5e74STEM\u535a\u58eb\u6bd5\u4e1a\u751f\u7ea68\u4e07\u4eba\uff0c\u6570\u91cf\u8d85\u8fc7\u7f8e\u56fd\n'
  '\u2022 \u4f46\uff1a\u8bfa\u8d1d\u5c14\u81ea\u7136\u79d1\u5b66\u5956\u672c\u571f\u83b7\u5956\u4ec5\u5c60\u5466\u54661\u4eba\uff1b\u83f2\u5c14\u5179\u5956 0\u4eba', sz=12, c=C3)

R(s, 0.5, 3.4, 6.0, 3.0, fill=C1)
T(s, 0.7, 3.5, 5.6, 0.35, '\u94b1\u5b66\u68ee\u4e4b\u95ee\uff082005\u5e74\uff09', sz=15, b=True, c=C2)
T(s, 0.7, 3.95, 5.6, 2.3,
  '2005\u5e74\uff0c\u706b\u7bad\u79d1\u5b66\u5bb6\u94b1\u5b66\u68ee\u5411\u65f6\u4efb\u603b\u7406\u6e29\u5bb6\u5b9d\u63d0\u95ee\uff1a\n\n'
  '\u300c\u4e3a\u4ec0\u4e48\u6211\u4eec\u7684\u5b66\u6821\u603b\u662f\u57f9\u517b\u4e0d\u51fa\u6770\u51fa\u4eba\u624d\uff1f\u300d\n\n'
  '\u94b1\u5b66\u68ee\u8ba4\u4e3a\uff1a\u4e2d\u56fd\u6559\u80b2\u7f3a\u4e4f\u5bf9\u521b\u9020\u529b\u3001\n'
  '\u8de8\u5b66\u79d1\u601d\u7ef4\u548c\u521b\u65b0\u7cbe\u795e\u7684\u57f9\u517b\uff0c\n'
  '\u8fd9\u662f\u963b\u788d\u79d1\u5b66\u8fdb\u6b65\u7684\u6838\u5fc3\u969c\u788d\u3002\n\n'
  '\u8fd9\u4e00\u8ffd\u95ee\u81f3\u4eca\u4ecd\u662f\u4e2d\u56fd\u6559\u80b2\u6539\u9769\u7684\u7075\u9b42\u4e4b\u95ee\u3002', sz=12, c=C4)

T(s, 6.8, 3.5, 6.0, 0.35, '\u6df1\u5c42\u95ee\u9898 vs \u6539\u9769\u65b9\u5411', sz=15, b=True, c=C1)
items = [
    ('\u5e94\u8bd5\u5bfc\u5411', '\u9ad8\u8003\u6307\u6325\u68d2\u4e0b\uff0c\u6807\u51c6\u7b54\u6848\u538b\u5236\u6279\u5224\u6027\u601d\u7ef4', C7),
    ('\u8bba\u6587\u5bfc\u5411', '\u300c\u552f\u8bba\u6587\u300d\u8bc4\u4ef7\u5bfc\u81f4\u8ddf\u98ce\u7814\u7a76\uff0c\u539f\u521b\u4e0d\u8db3', C7),
    ('\u5f3a\u57fa\u8ba1\u5212', '2020\u5e74\u8d7736\u6240\u9ad8\u6821\u9009\u62d4\u57fa\u7840\u5b66\u79d1\u62d4\u5c16\u4eba\u624d', C6),
    ('\u65b0\u5de5\u79d1\u5efa\u8bbe', 'AI\u3001\u96c6\u6210\u7535\u8def\u7b49\u4ea4\u53c9\u5b66\u79d1\uff0c\u4ea7\u6559\u878d\u5408\u57f9\u517b', C6),
]
for i, (h, b, clr) in enumerate(items):
    y = 3.95 + i*0.72
    T(s, 6.8, y, 1.6, 0.3, h, sz=12, b=True, c=clr)
    T(s, 8.5, y, 4.3, 0.6, b, sz=11, c=C3)
FT(s, 6)


# === P7 Culture - Needham ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u56e0\u7d20\u2463 \u6587\u5316\u4f20\u7edf\u4e0e\u793e\u4f1a\u89c2\u5ff5')
R(s, 0.5, 1.6, 6.0, 2.8, fill=C1)
T(s, 0.7, 1.7, 5.6, 0.35, '\u674e\u7ea6\u745f\u96be\u9898 (Needham Question)', sz=14, b=True, c=C2)
T(s, 0.7, 2.1, 5.6, 2.2,
  '\u82f1\u56fd\u5b66\u8005\u674e\u7ea6\u745f\u5728\u300a\u4e2d\u56fd\u79d1\u5b66\u6280\u672f\u53f2\u300b\u4e2d\u63d0\u51fa\uff1a\n\n'
  '\u300c\u4e3a\u4ec0\u4e48\u8fd1\u4ee3\u79d1\u5b66\u6ca1\u6709\u5728\u4e2d\u56fd\u4ea7\u751f\uff0c\n'
  '\u5c3d\u7ba1\u53e4\u4ee3\u4e2d\u56fd\u7684\u6280\u672f\u6c34\u5e73\u957f\u671f\u9886\u5148\u4e16\u754c\uff1f\u300d\n\n'
  '\u53ef\u80fd\u539f\u56e0\uff1a\n'
  '\u2022 \u79d1\u4e3e\u5236\u5ea6\u91cd\u300c\u7ecf\u4e16\u81f4\u7528\u300d\uff0c\u8f7b\u7eaf\u7cb9\u7406\u8bba\u63a2\u7a76\n'
  '\u2022 \u7f3a\u4e4f\u5f62\u5f0f\u903b\u8f91\u4e0e\u6570\u5b66\u5316\u4f20\u7edf\n'
  '\u2022 \u5c01\u5efa\u4e13\u5236\u4e0d\u5229\u4e8e\u81ea\u7531\u5b66\u672f\u8ba8\u8bba', sz=12, c=C4)

T(s, 6.8, 1.6, 6.0, 0.4, '\u5f53\u4ee3\u6587\u5316\u89c2\u5ff5\u5bf9\u79d1\u6280\u7684\u5f71\u54cd', sz=15, b=True, c=C1)
items = [
    ('\u79d1\u5b66\u7cbe\u795e\u4e0d\u8db3', '\u516c\u4f17\u79d1\u5b66\u7d20\u517b\u6bd4\u4f8b\u4ec514.14%\uff082023\u4e2d\u56fd\u79d1\u534f\uff09\n\u53d1\u8fbe\u56fd\u5bb6\u901a\u5e3825-30%\u4ee5\u4e0a', C7),
    ('\u6025\u529f\u8fd1\u5229\u98ce\u6c14', '\u300c\u77ed\u5e73\u5feb\u300d\u7814\u7a76\u504f\u597d\uff0c\u57fa\u7840\u7814\u7a76\u5750\u51b7\u677f\u51f3\u96be\n\u79d1\u7814\u4eba\u5747\u8bba\u6587\u591a\u4f46\u539f\u521b\u6027\u7a81\u7834\u5c11', C7),
    ('\u5bf9\u5931\u8d25\u5bb9\u5fcd\u5ea6\u4f4e', '\u79d1\u7814\u9879\u76ee\u5fc5\u987b\u6210\u529f\u624d\u80fd\u7ed3\u9898\uff0c\u4e0d\u5141\u8bb8\n\u300c\u8bc1\u4f2a\u300d\u4e5f\u662f\u79d1\u5b66\u8d21\u732e\u2014\u2014\u6587\u5316\u8ba4\u77e5\u504f\u5dee', C7),
    ('\u79ef\u6781\u53d8\u5316', '\u79d1\u5b66\u5bb6\u7cbe\u795e\u5ba3\u4f20\uff08\u5357\u4ec1\u4e1c\u3001\u9ec4\u5927\u5e74\uff09\u3001\u79d1\u5e7b\u6587\u5316\n\u5174\u8d77\uff08\u300a\u4e09\u4f53\u300b\u300a\u6d41\u6d6a\u5730\u7403\u300b\uff09\u63d0\u5347\u79d1\u5b66\u5173\u6ce8\u5ea6', C6),
]
for i, (h, b, clr) in enumerate(items):
    y = 2.1 + i*1.15
    T(s, 6.8, y, 2.0, 0.3, h, sz=12, b=True, c=clr)
    T(s, 6.8, y+0.35, 5.8, 0.7, b, sz=11, c=C3)
FT(s, 7)


# === P8 International ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u56e0\u7d20\u2464 \u56fd\u9645\u73af\u5883\u4e0e\u79d1\u6280\u4ea4\u6d41')
R(s, 0.5, 1.6, 12.3, 2.6, fill=C12)
T(s, 0.7, 1.7, 11.9, 0.35, '\u300c\u5361\u8116\u5b50\u300d\u6e05\u5355\u2014\u2014\u6280\u672f\u5c01\u9501\u7684\u73b0\u5b9e\u5f71\u54cd', sz=14, b=True, c=C7)
T(s, 0.7, 2.1, 5.6, 2.0,
  '2018\u5e74\u8d77\u7f8e\u56fd\u5bf9\u534e\u6280\u672f\u5c01\u9501\u5347\u7ea7\uff1a\n'
  '\u2022 \u5149\u523b\u673a\uff1aASML EUV\u7981\u552e\uff0c\u4e2d\u56fd\u53ea\u80fd\u7528DUV\n'
  '\u2022 EDA\u8f6f\u4ef6\uff1aSynopsys/Cadence/Mentor\u65ad\u4f9b\u534e\u4e3a\n'
  '\u2022 AI\u82af\u7247\uff1aNVIDIA A100/H100\u5bf9\u534e\u7981\u552e\n'
  '\u2022 \u64cd\u4f5c\u7cfb\u7edf\uff1aAndroid GMS\u670d\u52a1\u5bf9\u534e\u4e3a\u65ad\u4f9b', sz=12, c=C3)
T(s, 6.5, 2.1, 6.1, 2.0,
  '\u5012\u903c\u4e4b\u4e0b\u7684\u56fd\u4ea7\u7a81\u7834\uff1a\n'
  '\u2022 \u534e\u4e3a\u9e92\u9e9f9000s(2023)\uff1aSMIC 7nm\uff0c\u56de\u5f525G\n'
  '\u2022 \u534e\u4e3aMate70\u642d\u8f7d\u9e92\u9e9f9100\uff1aSMIC 6nm(2024.11)\n'
  '\u2022 HarmonyOS NEXT\uff1a\u5b8c\u5168\u53bb\u5b89\u5353\u5316\u7684\u81ea\u4e3bOS\n'
  '\u2022 DeepSeek-R1(2025.1)\uff1a\u5f00\u6e90AI\u6a21\u578b\uff0c\u6027\u80fd\u6bd4\u80a9\n'
  '  OpenAI o1\uff0c\u6210\u672c\u4ec51/20\uff0c\u9707\u52a8\u7845\u8c37', sz=12, c=C3)

R(s, 0.5, 4.5, 12.3, 2.2, fill=C5)
T(s, 0.7, 4.6, 11.9, 0.35, '\u56fd\u9645\u5408\u4f5c\u4ecd\u7136\u662f\u79d1\u5b66\u53d1\u5c55\u7684\u52a0\u901f\u5668', sz=14, b=True, c=C6)
T(s, 0.7, 5.0, 5.6, 1.5,
  '\u2022 CERN\uff1a\u4e2d\u56fd\u53c2\u4e0e\u5e0c\u683c\u65af\u7c92\u5b50\u53d1\u73b0\u76f8\u5173\u5b9e\u9a8c\n'
  '\u2022 ITER\u56fd\u9645\u70ed\u6838\u805a\u53d8\uff1a\u4e2d\u56fd\u627f\u62c5\u7ea69%\u7ecf\u8d39\u4e0e\u90e8\u4ef6\n'
  '\u2022 \u4eba\u7c7b\u57fa\u56e0\u7ec4\u8ba1\u5212\uff1a\u4e2d\u56fd\u627f\u62c51%\u6d4b\u5e8f\u4efb\u52a1(1999)', sz=12, c=C3)
T(s, 6.5, 5.0, 6.1, 1.5,
  '\u2022 SKA\u5e73\u65b9\u516c\u91cc\u9635\u5217\u5c04\u7535\u671b\u8fdc\u955c\uff1a\u4e2d\u56fd\u662f\u521b\u59cb\u6210\u5458\n'
  '\u2022 \u5ae6\u5a25\u516d\u53f7\u642d\u8f7d\u6cd5\u56fd\u3001\u610f\u5927\u5229\u3001\u5df4\u57fa\u65af\u5766\u7b49\u56fd\u8f7d\u8377\n'
  '\u2022 \u4e2d\u56fd\u7a7a\u95f4\u7ad9\u542717\u56fd\u5f00\u653e\u79d1\u5b66\u5b9e\u9a8c\n'
  '\u2192 \u542f\u793a\uff1a\u5173\u952e\u9886\u57df\u81ea\u4e3b\u53ef\u63a7+\u975e\u654f\u611f\u9886\u57df\u5f00\u653e\u5408\u4f5c', sz=12, c=C3)
FT(s, 8)


# === P9 Other factors ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u5176\u4ed6\u91cd\u8981\u793e\u4f1a\u56e0\u7d20')
headers = ['\u56e0\u7d20', '\u4f5c\u7528\u673a\u5236', '\u771f\u5b9e\u6848\u4f8b']
rows = [
    ('\u6cd5\u5f8b/\u77e5\u8bc6\u4ea7\u6743', '\u4e13\u5229\u4fdd\u62a4\u6fc0\u52b1\u521b\u65b0\u6295\u5165',
     '\u7f8e\u56fd\u300a\u62dc\u675c\u6cd5\u6848\u300b(1980)\u50ac\u751f\u7845\u8c37\u6280\u672f\u8f6c\u8ba9\u7206\u53d1\nPCT\u4e13\u5229\u4e2d\u56fd\u8fde\u7eed5\u5e74\u5168\u7403\u7b2c\u4e00(69,610\u4ef6/\u5e74)'),
    ('\u519b\u4e8b\u9700\u6c42', '\u56fd\u9632\u6218\u7565\u50ac\u751f\u91cd\u5927\u6280\u672f\u7a81\u7834',
     '\u4e92\u8054\u7f51\u524d\u8eabARPANET(1969)\u6e90\u4e8e\u7f8e\u519b\u65b9\u901a\u4fe1\u9700\u6c42\n\u5317\u6597\u7cfb\u7edf(2020\u5168\u7403\u7ec4\u7f51)\u6e90\u4e8e\u519b\u4e8b\u5b9a\u4f4d\u81ea\u4e3b\u9700\u6c42'),
    ('\u5a92\u4f53\u4e0e\u8206\u8bba', '\u5f71\u54cd\u653f\u7b56\u8d70\u5411\u548c\u516c\u4f17\u63a5\u53d7\u5ea6',
     '\u8f6c\u57fa\u56e0\u4e89\u8bae\u5bfc\u81f4\u4e2d\u56fd\u5546\u4e1a\u5316\u79cd\u690d\u63a8\u8fdf10\u5e74\u4ee5\u4e0a\n\u6838\u7535\u90bb\u907f\u6548\u5e94\uff1a\u798f\u5c9b\u540e\u5168\u7403\u6838\u7535\u5ba1\u6279\u8d8b\u4e25'),
    ('\u4eba\u53e3\u4e0e\u4ee3\u9645\u7ed3\u6784', '\u79d1\u7814\u540e\u5907\u529b\u91cf\u3001\u5e02\u573a\u89c4\u6a21',
     '\u65e5\u672c\u5c11\u5b50\u5316\u2192\u79d1\u7814\u4eba\u624d\u65ad\u5c42\uff0c\u9ad8\u6821\u62db\u4e0d\u6ee1\u535a\u58eb\u751f\n\u4e2d\u56fd\u5de5\u7a0b\u5e08\u7ea2\u5229\uff1a\u6bcf\u5e74STEM\u6bd5\u4e1a\u751f\u8d85500\u4e07'),
]
y = 1.65
for i, h in enumerate(headers):
    x = [0.5, 2.6, 6.5][i]; w = [2.0, 3.8, 6.3][i]
    R(s, x, y, w, 0.45, fill=C1)
    T(s, x+0.1, y+0.05, w-0.2, 0.35, h, sz=13, b=True, c=C4, a=PP_ALIGN.CENTER)
y += 0.5
for ri, (c1, c2, c3) in enumerate(rows):
    bkg = C5 if ri%2==0 else C4
    rh = 1.2
    for i, val in enumerate([c1, c2, c3]):
        x = [0.5, 2.6, 6.5][i]; w = [2.0, 3.8, 6.3][i]
        R(s, x, y, w, rh, fill=bkg, lc=RGBColor(0xCC,0xCC,0xCC))
        T(s, x+0.1, y+0.1, w-0.2, rh-0.2, val, sz=11, c=C1 if i==0 else C3, b=(i==0))
    y += rh
FT(s, 9)


# === P10 Part 2 overview ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u7b2c\u4e8c\u90e8\u5206\uff1a\u5982\u4f55\u4fc3\u8fdb\u6211\u56fd\u79d1\u6280\u8fdb\u6b65\uff1f')
R(s, 0.5, 1.6, 12.3, 0.7, fill=C1)
T(s, 0.5, 1.6, 12.3, 0.7,
  '\u603b\u76ee\u6807\uff1a\u5b9e\u73b0\u9ad8\u6c34\u5e73\u79d1\u6280\u81ea\u7acb\u81ea\u5f3a\uff08\u4e8c\u5341\u5927\u62a5\u544a\uff09 | \u56db\u4e2a\u9762\u5411\uff1a\u4e16\u754c\u79d1\u6280\u524d\u6cbf/\u7ecf\u6d4e\u4e3b\u6218\u573a/\u56fd\u5bb6\u91cd\u5927\u9700\u6c42/\u4eba\u6c11\u751f\u547d\u5065\u5eb7',
  sz=13, b=True, c=C4, a=PP_ALIGN.CENTER, v=MSO_ANCHOR.MIDDLE)
measures = [
    ('\u2460', '\u52a0\u5927\u7814\u53d1\u6295\u5165', '\u63d0\u5347\u57fa\u7840\u7814\u7a76\u5360\u6bd4\n\u76ee\u6807\uff1a15%\u4ee5\u4e0a'),
    ('\u2461', '\u6539\u9769\u6559\u80b2\u4f53\u7cfb', '\u5f3a\u57fa\u8ba1\u5212+\u65b0\u5de5\u79d1\n\u57f9\u517b\u539f\u521b\u578b\u4eba\u624d'),
    ('\u2462', '\u6df1\u5316\u4f53\u5236\u6539\u9769', '\u7834\u56db\u552f/\u653e\u7ba1\u670d\n\u79d1\u7814\u8bc4\u4ef7\u53bb\u884c\u653f\u5316'),
    ('\u2463', '\u8425\u9020\u521b\u65b0\u6587\u5316', '\u5bb9\u5fcd\u5931\u8d25/\u9f13\u52b1\u5192\u9669\n\u63d0\u5347\u516c\u4f17\u79d1\u5b66\u7d20\u517b'),
    ('\u2464', '\u653b\u514b\u5361\u8116\u5b50', '\u82af\u7247/OS/\u5de5\u4e1a\u8f6f\u4ef6\n\u65b0\u578b\u4e3e\u56fd\u4f53\u5236'),
    ('\u2465', '\u9ad8\u6c34\u5e73\u5f00\u653e', '\u53c2\u4e0e\u56fd\u9645\u5927\u79d1\u5b66\n\u5438\u5f15\u5168\u7403\u4eba\u624d'),
]
for i, (num, ti, bd) in enumerate(measures):
    col, row = i%3, i//3
    x = 0.5 + col*4.15; y = 2.6 + row*2.1
    R(s, x, y, 3.95, 1.9, fill=C4, lc=C1)
    R(s, x, y, 0.6, 1.9, fill=C2)
    T(s, x, y, 0.6, 1.9, num, sz=22, b=True, c=C4, a=PP_ALIGN.CENTER, v=MSO_ANCHOR.MIDDLE, f='Arial')
    T(s, x+0.75, y+0.15, 3.0, 0.4, ti, sz=15, b=True, c=C1)
    T(s, x+0.75, y+0.6, 3.0, 1.2, bd, sz=12, c=C3)
FT(s, 10)


# === P11 Invest + Talent with line chart ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u5bf9\u7b56\u2460\u2461 \u52a0\u5927\u6295\u5165 + \u57f9\u517b\u4eba\u624d')
R(s, 0.5, 1.6, 6.0, 2.4, fill=C5)
T(s, 0.7, 1.7, 5.6, 0.35, '\u2460 \u6301\u7eed\u52a0\u5927\u7814\u53d1\u6295\u5165', sz=14, b=True, c=C1)
T(s, 0.7, 2.1, 5.6, 1.8,
  '\u73b0\u72b6\uff1a\u57fa\u7840\u7814\u7a76\u7ecf\u8d39\u5360R&D\u603b\u7ecf\u8d396.65%\uff082023\uff09\uff0c\u8fdc\u4f4e\u4e8e\n'
  '\u7f8e\u56fd(16.5%)\u3001\u6cd5\u56fd(23.2%)\u3001\u65e5\u672c(12.5%)\n\n'
  '\u5efa\u8bae\uff1a\n'
  '\u2022 \u8bbe\u7acb\u957f\u5468\u671f\u57fa\u91d1\uff0810-15\u5e74\uff09\uff0c\u5bb9\u5fcd\u77ed\u671f\u65e0\u4ea7\u51fa\n'
  '\u2022 \u56fd\u5bb6\u81ea\u7136\u79d1\u5b66\u57fa\u91d1\u6269\u89c4\u6a21\uff082024\u5e74\u53d7\u740838.4\u4e07\u9879\u7533\u8bf7\uff09\n'
  '\u2022 \u5927\u79d1\u5b66\u88c5\u7f6e\uff1aFAST\u5df2\u53d1\u73b0\u8d85900\u9897\u8109\u51b2\u661f', sz=11, c=C3)

R(s, 6.7, 1.6, 6.1, 2.4, fill=C5)
T(s, 6.9, 1.7, 5.7, 0.35, '\u2461 \u6559\u80b2\u6539\u9769\u4e0e\u4eba\u624d\u8ba1\u5212', sz=14, b=True, c=C1)
T(s, 6.9, 2.1, 5.7, 1.8,
  '\u2022 \u5f3a\u57fa\u8ba1\u5212(2020-)\uff1a\u9996\u627936\u6240\u9ad8\u6821\uff0c\u672c\u7855\u535a\u8d2f\u901a\u57f9\u517b\n'
  '\u2022 \u65b0\u5de5\u79d1\u5efa\u8bbe\uff1aAI\u3001\u96c6\u6210\u7535\u8def\u3001\u91cf\u5b50\u4fe1\u606f\u4e13\u4e1a\u6269\u62db\n'
  '\u2022 \u9752\u5e74\u79d1\u5b66\u5bb6\u652f\u6301\uff1a\u56fd\u81ea\u7136\u4f18\u9752\u540d\u989d\u4ece600\u6269\u81f3800\n'
  '\u2022 \u6d77\u5916\u5f15\u624d\uff1a\u7d2f\u8ba1\u5f15\u8fdb\u9ad8\u5c42\u6b21\u4eba\u624d\u6570\u4e07\u4eba\n'
  '\u2022 \u6539\u9769\u65b9\u5411\uff1a\u7ed935\u5c81\u4ee5\u4e0b\u9752\u5e74\u5b66\u8005\u72ec\u7acbPI\u5c97\u4f4d', sz=11, c=C3)

T(s, 0.5, 4.3, 12, 0.35, '\u4e2d\u56fdR&D\u7ecf\u8d39\u589e\u957f\u8d8b\u52bf\uff08\u4e07\u4ebf\u5143\uff09', sz=13, b=True, c=C1)
cd = CategoryChartData()
cd.categories = ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024']
cd.add_series('R&D', (1.42, 1.57, 1.76, 1.97, 2.21, 2.44, 2.80, 3.09, 3.34, 3.63))
ch = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.5), Inches(4.7), Inches(12.3), Inches(2.2), cd).chart
ch.has_legend = False
pl = ch.plots[0]; pl.has_data_labels = True
dl = pl.data_labels; dl.font.size = Pt(9); dl.font.color.rgb = C1
sr = ch.series[0]; sr.format.line.color.rgb = C2; sr.format.line.width = Pt(2.5)
sr.smooth = True
FT(s, 11)


# === P12 Reform + Culture ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u5bf9\u7b56\u2462\u2463 \u4f53\u5236\u6539\u9769 + \u521b\u65b0\u6587\u5316')
R(s, 0.5, 1.6, 6.0, 5.0, fill=C4, lc=C1)
T(s, 0.7, 1.7, 5.6, 0.4, '\u2462 \u6df1\u5316\u79d1\u6280\u4f53\u5236\u6539\u9769', sz=15, b=True, c=C1)
T(s, 0.7, 2.15, 5.6, 4.2,
  '\u7834\u300c\u56db\u552f\u300d\uff082018\u5e74\u56fd\u52a1\u9662\u53d1\u6587\uff09\uff1a\n'
  '  \u2717 \u552f\u8bba\u6587 \u2192 \u770b\u5b9e\u9645\u8d21\u732e\u548c\u540c\u884c\u8bc4\u8bae\n'
  '  \u2717 \u552f\u804c\u79f0 \u2192 \u770b\u7814\u7a76\u5b9e\u7ee9\n'
  '  \u2717 \u552f\u5b66\u5386 \u2192 \u770b\u5b9e\u9645\u80fd\u529b\n'
  '  \u2717 \u552f\u5956\u9879 \u2192 \u770b\u957f\u671f\u5f71\u54cd\n\n'
  '\u5177\u4f53\u4e3e\u63aa\uff1a\n'
  '\u2022 \u300c\u5305\u5e72\u5236\u300d\u8bd5\u70b9\uff1a\u56fd\u81ea\u7136\u6770\u9752\u9879\u76ee\u7ecf\u8d39\u4e0d\u8bbe\u79d1\u76ee\u9650\u5236\n'
  '\u2022 \u79d1\u7814\u4eba\u5458\u51cf\u8d1f\uff1a\u7cbe\u7b80\u62a5\u8868\u3001\u538b\u7f29\u8bc4\u5ba1\u4f1a\u8bae\n'
  '\u2022 \u4fc3\u8fdb\u8f6c\u5316\uff1a\u5141\u8bb8\u79d1\u7814\u4eba\u5458\u6301\u6709\u6210\u679c\u8f6c\u5316\u80a1\u6743\n'
  '\u2022 \u300c\u63ed\u699c\u6302\u5e05\u300d\u673a\u5236\uff1a\u4e0d\u8bba\u8d44\u5386\uff0c\u8c01\u80fd\u5e72\u8c01\u4e0a', sz=11, c=C3)

R(s, 6.7, 1.6, 6.1, 5.0, fill=C4, lc=C1)
T(s, 6.9, 1.7, 5.7, 0.4, '\u2463 \u8425\u9020\u5c0a\u91cd\u79d1\u5b66\u7684\u793e\u4f1a\u6587\u5316', sz=15, b=True, c=C1)
T(s, 6.9, 2.15, 5.7, 4.2,
  '\u5f18\u626c\u79d1\u5b66\u5bb6\u7cbe\u795e\uff1a\n'
  '\u2022 \u5357\u4ec1\u4e1c\uff1a22\u5e74\u9009\u5740\u5efa\u6210FAST\uff08500\u7c73\u53e3\u5f84\u5c04\u7535\u671b\u8fdc\u955c\uff09\n'
  '  2017\u5e74\u53bb\u4e16\u524d\u4ecd\u5728\u73b0\u573a\u5de5\u4f5c\uff0c2020\u5e74FAST\u6b63\u5f0f\u8fd0\u884c\n'
  '  \u81f3\u4eca\u53d1\u73b0\u8d85900\u9897\u65b0\u8109\u51b2\u661f\n'
  '\u2022 \u9ec4\u5927\u5e74\uff1a\u653e\u5f03\u82f1\u56fd\u9ad8\u85aa\u56de\u56fd\uff0c\u63a8\u8fdb\u5730\u7403\u6df1\u90e8\u63a2\u6d4b\n\n'
  '\u63d0\u5347\u5168\u6c11\u79d1\u5b66\u7d20\u517b\uff1a\n'
  '\u2022 2023\u5e74\u5177\u5907\u79d1\u5b66\u7d20\u8d28\u7684\u516c\u6c11\u6bd4\u4f8b14.14%\uff08\u4e2d\u56fd\u79d1\u534f\uff09\n'
  '\u2022 \u5341\u56db\u4e94\u76ee\u6807\uff1a2025\u5e74\u8fbe\u523015%\n'
  '\u2022 \u5bf9\u6bd4\uff1a\u7f8e\u56fd\u7ea628%\uff0c\u745e\u5178\u7ea635%\u2014\u2014\u5dee\u8ddd\u660e\u663e\n\n'
  '\u5bb9\u5fcd\u5931\u8d25\u6587\u5316\uff1a\n'
  '\u2022 \u57fa\u7840\u7814\u7a76\u5931\u8d25\u7387>90%\u662f\u6b63\u5e38\u7684\n'
  '\u2022 \u9700\u5efa\u7acb\u300c\u5931\u8d25\u4e5f\u662f\u8d21\u732e\u300d\u7684\u8bc4\u4ef7\u5171\u8bc6', sz=11, c=C3)
FT(s, 12)


# === P13 Breakthroughs ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u5bf9\u7b56\u2464\u2465 \u653b\u514b\u5361\u8116\u5b50 + \u9ad8\u6c34\u5e73\u5f00\u653e')
R(s, 0.5, 1.6, 12.3, 2.8, fill=C1)
T(s, 0.7, 1.7, 11.9, 0.35, '\u2464 \u5173\u952e\u6838\u5fc3\u6280\u672f\u653b\u5173\u2014\u2014\u5df2\u6709\u7a81\u7834\u7684\u771f\u5b9e\u6848\u4f8b', sz=14, b=True, c=C2)
cases = [
    ('C919\u5927\u98de\u673a', '2023.5.28\u5546\u4e1a\u9996\u98de\n\u4e0a\u6d77\u2192\u5317\u4eac\n\u6253\u7834\u6ce2\u97f3/\u7a7a\u5ba2\u5784\u65ad\nCOMAC\u5386\u65f615\u5e74'),
    ('\u5ae6\u5a25\u516d\u53f7', '2024.6.25\n\u4eba\u7c7b\u9996\u6b21\u6708\u80cc\u53d6\u6837\n1935.3\u514b\u6708\u58e4\n53\u5929\u4efb\u52a1\u5168\u7a0b\u81ea\u4e3b'),
    ('DeepSeek-R1', '2025.1\u5f00\u6e90\n\u6027\u80fd\u6bd4\u80a9OpenAI o1\n\u8bad\u7ec3\u6210\u672c\u4ec51/20\nMIT\u534f\u8bae\u514d\u8d39\u4f7f\u7528'),
    ('\u58a8\u5b50\u53f7\u536b\u661f', '2016.8\u53d1\u5c04\n\u5168\u7403\u9996\u9897\u91cf\u5b50\u536b\u661f\n\u5b9e\u73b01200km\u91cf\u5b50\n\u6001\u8fdc\u7a0b\u4f20\u8f93(\u6f58\u5efa\u4f1f)'),
]
for i, (ti, bd) in enumerate(cases):
    x = 0.7 + i*3.1
    R(s, x, 2.15, 2.9, 2.1, fill=C10, lc=C2)
    T(s, x+0.1, 2.2, 2.7, 0.35, ti, sz=13, b=True, c=C2, a=PP_ALIGN.CENTER)
    T(s, x+0.1, 2.6, 2.7, 1.5, bd, sz=11, c=C4, a=PP_ALIGN.CENTER)

R(s, 0.5, 4.7, 12.3, 2.0, fill=C5)
T(s, 0.7, 4.8, 11.9, 0.35, '\u2465 \u9ad8\u6c34\u5e73\u5bf9\u5916\u5f00\u653e', sz=14, b=True, c=C1)
T(s, 0.7, 5.2, 5.8, 1.4,
  '\u4e3b\u52a8\u53c2\u4e0e\u56fd\u9645\u5927\u79d1\u5b66\u8ba1\u5212\uff1a\n'
  '\u2022 SKA\u5c04\u7535\u671b\u8fdc\u955c\u9635\u5217\uff08\u521b\u59cb\u6210\u5458\u56fd\uff09\n'
  '\u2022 ITER\u56fd\u9645\u70ed\u6838\u805a\u53d8\u5b9e\u9a8c\u5806\n'
  '\u2022 \u4e2d\u56fd\u7a7a\u95f4\u7ad9\u542717\u56fd\u5f00\u653e\u79d1\u5b66\u5b9e\u9a8c\n'
  '\u2022 \u5ae6\u5a25\u516d\u53f7\u642d\u8f7d\u6cd5/\u610f/\u5df4\u7b49\u56fd\u8f7d\u8377', sz=12, c=C3)
T(s, 6.7, 5.2, 6.0, 1.4,
  '\u4eba\u624d\u653f\u7b56\uff1a\n'
  '\u2022 \u6d77\u5916\u4f18\u9752\u9879\u76ee\uff1a\u5438\u5f1535\u5c81\u4ee5\u4e0b\u5b66\u8005\u56de\u56fd\n'
  '\u2022 \u6765\u534e\u7559\u5b66\u751f\u4e2dSTEM\u5360\u6bd4\u63d0\u5347\u81f340%+\n'
  '\u2022 \u5f00\u653e\u7b56\u7565\uff1a\u975e\u654f\u611f\u9886\u57df\u5408\u4f5c\uff0c\u6838\u5fc3\u9886\u57df\u81ea\u4e3b\n'
  '\u2022 \u4e0d\u641e\u300c\u95ed\u95e8\u9020\u8f66\u300d\uff0c\u5728\u7ade\u4e89\u4e2d\u5b66\u4e60', sz=12, c=C3)
FT(s, 13)


# === P14 Conclusion ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u7ed3\u8bba')
conclusions = [
    ('1', '\u79d1\u5b66\u53d1\u5c55\u662f\u793e\u4f1a\u7cfb\u7edf\u5de5\u7a0b',
     '\u4e0d\u662f\u5355\u9760\u79d1\u5b66\u5bb6\u7684\u5929\u624d\u5c31\u80fd\u5b9e\u73b0\u3002\u653f\u6cbb\u7a33\u5b9a\u3001\u7ecf\u6d4e\u6295\u5165\u3001\u6559\u80b2\u8d28\u91cf\u3001\u6587\u5316\u6c1b\u56f4\u3001\u56fd\u9645\u73af\u5883'
     '\u7f3a\u4e00\u4e0d\u53ef\u3002\u674e\u68ee\u79d1\u4e8b\u4ef6\u3001\u6587\u9769\u5bf9\u79d1\u7814\u7684\u7834\u574f\u90fd\u662f\u60e8\u75db\u6559\u8bad\u3002'),
    ('2', '\u4e2d\u56fd\u6b63\u5904\u4ece\u300c\u8ddf\u8dd1\u300d\u5230\u300c\u5e76\u8dd1/\u9886\u8dd1\u300d\u7684\u5173\u952e\u9636\u6bb5',
     'R&D\u603b\u91cf\u5df2\u5168\u7403\u7b2c\u4e00\uff0cPCT\u4e13\u5229\u8fde\u7eed5\u5e74\u7b2c\u4e00\uff0cGII\u6392\u540d\u5347\u81f3\u7b2c11\uff0c\u4f46\u57fa\u7840\u7814\u7a76\u6295\u5165\u4e0d\u8db3\u3001'
     '\u539f\u521b\u6027\u4e0d\u591f\u3001\u5361\u8116\u5b50\u95ee\u9898\u672a\u6839\u672c\u89e3\u51b3\u3002'),
    ('3', '\u9700\u8981\u957f\u671f\u4e3b\u4e49 + \u7cfb\u7edf\u63a8\u8fdb + \u5f00\u653e\u80f8\u6000',
     '\u52a0\u5927\u57fa\u7840\u7814\u7a76\u6295\u5165\u9700\u8981\u8010\u5fc3\uff1b\u4f53\u5236\u6539\u9769\u9700\u8981\u52c7\u6c14\uff1b\u6587\u5316\u8f6c\u53d8\u9700\u8981\u4ee3\u9645\u79ef\u7d2f\uff1b'
     '\u5bf9\u5916\u5f00\u653e\u9700\u8981\u667a\u6167\u3002\u4e0d\u53ef\u80fd\u6bd5\u5176\u529f\u4e8e\u4e00\u5f79\u3002'),
]
y = 1.6
for num, ti, bd in conclusions:
    R(s, 0.5, y, 0.7, 0.7, fill=C2, sh=MSO_SHAPE.OVAL)
    T(s, 0.5, y, 0.7, 0.7, num, sz=22, b=True, c=C4, a=PP_ALIGN.CENTER, v=MSO_ANCHOR.MIDDLE, f='Arial')
    T(s, 1.4, y+0.05, 11.4, 0.35, ti, sz=16, b=True, c=C1)
    T(s, 1.4, y+0.4, 11.4, 0.7, bd, sz=12, c=C3)
    y += 1.35
R(s, 0.5, 5.7, 12.3, 1.0, fill=C1)
T(s, 0.5, 5.75, 12.3, 0.5,
  '\u300c\u79d1\u6280\u5174\u5219\u6c11\u65cf\u5174\uff0c\u79d1\u6280\u5f3a\u5219\u56fd\u5bb6\u5f3a\u3002\u300d', sz=20, b=True, c=C4, a=PP_ALIGN.CENTER, v=MSO_ANCHOR.MIDDLE)
T(s, 0.5, 6.25, 12.3, 0.35,
  '\u2014\u2014 \u4e60\u8fd1\u5e73  \u5728\u5168\u56fd\u79d1\u6280\u521b\u65b0\u5927\u4f1a\u4e0a\u7684\u8bb2\u8bdd', sz=12, c=C2, a=PP_ALIGN.CENTER)
FT(s, 14)


# === P15 References ===
s = prs.slides.add_slide(BL)
bg(s, C4); TB(s, '\u53c2\u8003\u6587\u732e')
refs = [
    '[1] \u56fd\u5bb6\u7edf\u8ba1\u5c40. 2024\u5e74\u5168\u56fd\u79d1\u6280\u7ecf\u8d39\u6295\u5165\u7edf\u8ba1\u516c\u62a5[R]. stats.gov.cn, 2025.2.',
    '[2] WIPO. Global Innovation Index 2024[R]. wipo.int, 2024.9.',
    '[3] WIPO. PCT Yearly Review 2024: China tops for 5th year[R]. 2024.3.',
    '[4] ITIF. China Overtakes the US in R&D Investment[R]. itif.org, 2026.5.',
    '[5] \u8d1d\u5c14\u7eb3 J.D. \u79d1\u5b66\u7684\u793e\u4f1a\u529f\u80fd[M]. \u5546\u52a1\u5370\u4e66\u9986, 1982.',
    '[6] \u674e\u7ea6\u745f. \u4e2d\u56fd\u79d1\u5b66\u6280\u672f\u53f2[M]. \u79d1\u5b66\u51fa\u7248\u793e.',
    '[7] Borinskaya et al. Lysenkoism Against Genetics[J]. Genetics, 2019.',
    '[8] \u4e2d\u5171\u4e2d\u592e\u56fd\u52a1\u9662. \u5341\u56db\u4e94\u56fd\u5bb6\u79d1\u6280\u521b\u65b0\u89c4\u5212[Z]. 2021.',
    '[9] MIT Tech Review. How DeepSeek ripped up the AI playbook[EB/OL]. 2025.1.',
    '[10] People\'s Daily. C919 completes inaugural commercial flight[N]. 2023.5.',
    '[11] Wikipedia. Chang\'e 6: 1935.3g lunar far side samples[EB/OL]. 2024.',
    '[12] Nature. China\'s cheap open AI model DeepSeek thrills scientists[J]. 2025.1.',
]
y = 1.5
for ref in refs:
    T(s, 0.7, y, 11.9, 0.38, ref, sz=11, c=C3)
    y += 0.38
R(s, 0, 6.3, 13.333, 1.2, fill=C1)
T(s, 0, 6.4, 13.333, 0.5, '\u611f \u8c22 \u8046 \u542c', sz=30, b=True, c=C4, a=PP_ALIGN.CENTER)
T(s, 0, 6.9, 13.333, 0.4, 'Thanks for Your Attention', sz=13, c=C2, a=PP_ALIGN.CENTER, f='Calibri')

# === SAVE ===
OUT = '/projects/sandbox/RUOTZ/\u79d1\u6280\u89c2\u4f5c\u4e1a-\u5f71\u54cd\u79d1\u5b66\u53d1\u5c55\u7684\u793e\u4f1a\u56e0\u7d20.pptx'
prs.save(OUT)
print(f'Done: {OUT}')
print(f'Slides: {len(prs.slides)}')
