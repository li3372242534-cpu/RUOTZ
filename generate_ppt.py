# -*- coding: utf-8 -*-
"""
生成 PPT：悲观主义与乐观主义的科技价值观——基于马克思主义科技价值观的评价
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ============== 主题颜色 ==============
COLOR_PRIMARY      = RGBColor(0xC0, 0x39, 0x2B)   # 深红（马克思主义主题色）
COLOR_SECONDARY    = RGBColor(0x2C, 0x3E, 0x50)   # 深蓝灰
COLOR_ACCENT_OPT   = RGBColor(0xE6, 0x7E, 0x22)   # 橙（乐观主义）
COLOR_ACCENT_PES   = RGBColor(0x34, 0x49, 0x5E)   # 暗蓝（悲观主义）
COLOR_LIGHT_BG     = RGBColor(0xF8, 0xF5, 0xF0)   # 米白
COLOR_TEXT_DARK    = RGBColor(0x1C, 0x1C, 0x1C)
COLOR_TEXT_LIGHT   = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXT_GRAY    = RGBColor(0x6B, 0x6B, 0x6B)
COLOR_LINE         = RGBColor(0xB8, 0x9D, 0x6E)   # 古铜线条

CN_FONT_HEAD = "微软雅黑"
CN_FONT_BODY = "微软雅黑"

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)  # 16:9


# ============== 工具函数 ==============
def set_run_font(run, name=CN_FONT_BODY, size=18, bold=False,
                 color=COLOR_TEXT_DARK):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    # 同时设置东亚字体
    rPr = run._r.get_or_add_rPr()
    eastAsia = rPr.find(qn('a:eastAsia'))
    if eastAsia is None:
        from lxml import etree
        eastAsia = etree.SubElement(rPr, qn('a:ea'))
    rFonts = rPr.find(qn('a:rFonts'))
    if rFonts is None:
        from lxml import etree
        rFonts = etree.SubElement(rPr, qn('a:rFonts'))
    rFonts.set('eastAsia', name)
    rFonts.set('ascii', name)
    rFonts.set('hAnsi', name)


def add_textbox(slide, left, top, width, height, text,
                size=18, bold=False, color=COLOR_TEXT_DARK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                font=CN_FONT_BODY):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, name=font, size=size, bold=bold, color=color)
    return tb


def add_filled_rect(slide, left, top, width, height, fill_color,
                    line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.5)
    shape.shadow.inherit = False
    return shape


def add_line(slide, x1, y1, x2, y2, color=COLOR_LINE, weight=1.5):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    return line


def add_bullets(slide, left, top, width, height, items,
                size=16, color=COLOR_TEXT_DARK,
                bullet_color=COLOR_PRIMARY, line_spacing=1.35,
                bold_first_word=False):
    """items: list[str] 或 list[(标题, 内容)]"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)

    for idx, item in enumerate(items):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        p.space_after = Pt(4)

        # bullet
        run0 = p.add_run()
        run0.text = "■  "
        set_run_font(run0, size=size, bold=True, color=bullet_color)

        if isinstance(item, tuple):
            head, body = item
            run1 = p.add_run()
            run1.text = head
            set_run_font(run1, size=size, bold=True, color=COLOR_SECONDARY)
            run2 = p.add_run()
            run2.text = "  " + body
            set_run_font(run2, size=size, bold=False, color=color)
        else:
            run1 = p.add_run()
            run1.text = item
            set_run_font(run1, size=size, bold=False, color=color)
    return tb


def make_title_bar(slide, title, subtitle=None, page_no=None, total=None):
    """每张内容页统一的标题区"""
    # 顶部细色块
    add_filled_rect(slide, Emu(0), Emu(0), SLIDE_W, Inches(0.18),
                    COLOR_PRIMARY)
    # 主标题
    add_textbox(slide, Inches(0.55), Inches(0.35), Inches(11.5), Inches(0.7),
                title, size=30, bold=True, color=COLOR_SECONDARY,
                font=CN_FONT_HEAD)
    # 副标题
    if subtitle:
        add_textbox(slide, Inches(0.6), Inches(1.05), Inches(11.5),
                    Inches(0.4), subtitle, size=15, bold=False,
                    color=COLOR_TEXT_GRAY)
    # 装饰线
    add_line(slide, Inches(0.55), Inches(1.45),
             Inches(2.0), Inches(1.45), COLOR_PRIMARY, 2.5)
    # 页码
    if page_no is not None and total is not None:
        add_textbox(slide, Inches(11.8), Inches(7.05), Inches(1.3),
                    Inches(0.3), f"{page_no:02d} / {total:02d}",
                    size=10, color=COLOR_TEXT_GRAY, align=PP_ALIGN.RIGHT)
    # 页脚标签
    add_textbox(slide, Inches(0.55), Inches(7.05), Inches(8), Inches(0.3),
                "科技价值观研究  ·  马克思主义视角",
                size=10, color=COLOR_TEXT_GRAY)


# ============== 创建演示文稿 ==============
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
blank = prs.slide_layouts[6]

TOTAL = 16  # 内容页统计


# ---------- 1. 封面 ----------
def slide_cover():
    s = prs.slides.add_slide(blank)
    # 背景大色块
    add_filled_rect(s, Emu(0), Emu(0), SLIDE_W, SLIDE_H, COLOR_LIGHT_BG)
    # 左侧装饰红色块
    add_filled_rect(s, Emu(0), Emu(0), Inches(0.45), SLIDE_H, COLOR_PRIMARY)
    # 右上角小色块
    add_filled_rect(s, Inches(11.5), Inches(0.6), Inches(1.4),
                    Inches(0.12), COLOR_PRIMARY)

    # 主标题
    add_textbox(s, Inches(1.0), Inches(2.1), Inches(11.5), Inches(1.2),
                "悲观主义与乐观主义的科技价值观",
                size=44, bold=True, color=COLOR_SECONDARY,
                font=CN_FONT_HEAD)
    # 副标题
    add_textbox(s, Inches(1.0), Inches(3.3), Inches(11.5), Inches(0.8),
                "代表性思想及其马克思主义评价",
                size=26, bold=False, color=COLOR_PRIMARY,
                font=CN_FONT_HEAD)
    # 分割线
    add_line(s, Inches(1.0), Inches(4.3), Inches(5.0), Inches(4.3),
             COLOR_PRIMARY, 2.5)

    # 课程信息
    add_textbox(s, Inches(1.0), Inches(4.55), Inches(8), Inches(0.5),
                "自然辩证法概论 · 个人作业",
                size=18, color=COLOR_TEXT_GRAY)

    # 署名（占位）
    add_textbox(s, Inches(1.0), Inches(6.0), Inches(6), Inches(0.5),
                "汇  报  人：________________",
                size=16, color=COLOR_TEXT_DARK)
    add_textbox(s, Inches(1.0), Inches(6.45), Inches(6), Inches(0.5),
                "学       号：________________",
                size=16, color=COLOR_TEXT_DARK)
    add_textbox(s, Inches(1.0), Inches(6.9), Inches(6), Inches(0.5),
                "指导教师：________________",
                size=16, color=COLOR_TEXT_DARK)

    # 右下角日期
    add_textbox(s, Inches(9.5), Inches(6.9), Inches(3.5), Inches(0.5),
                "2026 年", size=14, color=COLOR_TEXT_GRAY,
                align=PP_ALIGN.RIGHT)


# ---------- 2. 目录 ----------
def slide_toc():
    s = prs.slides.add_slide(blank)
    add_filled_rect(s, Emu(0), Emu(0), SLIDE_W, SLIDE_H, COLOR_LIGHT_BG)

    # 大字"目录"
    add_textbox(s, Inches(0.7), Inches(0.6), Inches(6), Inches(1.0),
                "目  录", size=48, bold=True, color=COLOR_PRIMARY,
                font=CN_FONT_HEAD)
    add_textbox(s, Inches(0.75), Inches(1.55), Inches(6), Inches(0.5),
                "C O N T E N T S", size=14, color=COLOR_TEXT_GRAY)
    add_line(s, Inches(0.75), Inches(2.0), Inches(2.5), Inches(2.0),
             COLOR_PRIMARY, 2.5)

    items = [
        ("01", "引言：何谓科技价值观"),
        ("02", "科技乐观主义的代表性思想"),
        ("03", "科技悲观主义的代表性思想"),
        ("04", "马克思主义的科技价值观"),
        ("05", "马克思主义对两种思潮的评价"),
        ("06", "现实启示与结论"),
    ]
    # 两列布局
    col_x = [Inches(0.7), Inches(7.0)]
    for i, (num, txt) in enumerate(items):
        row = i % 3
        col = i // 3
        x = col_x[col]
        y = Inches(2.6 + row * 1.4)
        # 编号
        add_textbox(s, x, y, Inches(1.0), Inches(0.7),
                    num, size=40, bold=True, color=COLOR_PRIMARY,
                    font=CN_FONT_HEAD)
        # 文本
        add_textbox(s, x + Inches(1.2), y + Inches(0.05),
                    Inches(5.5), Inches(0.5),
                    txt, size=20, bold=True, color=COLOR_SECONDARY)
        # 短线
        add_line(s, x + Inches(1.2), y + Inches(0.7),
                 x + Inches(2.8), y + Inches(0.7),
                 COLOR_LINE, 1.2)


# ---------- 3. 引言 ----------
def slide_intro():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "一、引言：何谓科技价值观",
                   "Science-Technology Values: An Introduction",
                   page_no=1, total=TOTAL)

    # 左侧概念框
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(2.3),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(0.15),
                    Inches(2.3), COLOR_PRIMARY)
    add_textbox(s, Inches(0.85), Inches(1.85), Inches(5.7), Inches(0.5),
                "概念界定", size=20, bold=True, color=COLOR_PRIMARY)
    add_textbox(s, Inches(0.85), Inches(2.4), Inches(5.7), Inches(1.6),
                "科技价值观，是人们对科学技术的本质、作用、"
                "意义及其与人和社会关系的根本看法和总体评价；"
                "它回答的是「科技对人类究竟意味着什么」这一"
                "根本性问题。",
                size=15, color=COLOR_TEXT_DARK)

    # 右侧：两种对立倾向
    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(2.3),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(0.15),
                    Inches(2.3), COLOR_SECONDARY)
    add_textbox(s, Inches(7.15), Inches(1.85), Inches(5.7), Inches(0.5),
                "两种对立倾向", size=20, bold=True, color=COLOR_SECONDARY)
    add_textbox(s, Inches(7.15), Inches(2.4), Inches(5.7), Inches(1.6),
                "围绕近代以来的科技革命，思想界形成了"
                "「乐观主义」与「悲观主义」两大思潮："
                "前者讴歌科技的解放力量，后者警惕科技的"
                "异化危险。",
                size=15, color=COLOR_TEXT_DARK)

    # 底部要点
    add_textbox(s, Inches(0.55), Inches(4.3), Inches(12.3), Inches(0.5),
                "本课题需要回答的三个问题：", size=18,
                bold=True, color=COLOR_PRIMARY)
    add_bullets(s, Inches(0.85), Inches(4.85), Inches(12), Inches(2.0),
                [
                    ("Q1：", "科技乐观主义的代表性思想是什么？"),
                    ("Q2：", "科技悲观主义的代表性思想是什么？"),
                    ("Q3：", "如何运用马克思主义的科技价值观对两者作出评价？"),
                ], size=16, line_spacing=1.5)


# ---------- 4. 乐观主义概述 ----------
def slide_optimism_overview():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "二、科技乐观主义：概述",
                   "Technological Optimism: Overview",
                   page_no=2, total=TOTAL)

    # 定义
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(12.3), Inches(1.3),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(0.15),
                    Inches(1.3), COLOR_ACCENT_OPT)
    add_textbox(s, Inches(0.85), Inches(1.8), Inches(11.5), Inches(0.5),
                "核心命题", size=18, bold=True, color=COLOR_ACCENT_OPT)
    add_textbox(s, Inches(0.85), Inches(2.25), Inches(11.8), Inches(0.7),
                "科学技术是社会进步的根本动力，能够不断扩大人类的"
                "自由与福祉，最终带来理想社会的到来。",
                size=15, color=COLOR_TEXT_DARK)

    # 历史脉络
    add_textbox(s, Inches(0.55), Inches(3.2), Inches(12), Inches(0.5),
                "历史脉络", size=20, bold=True, color=COLOR_SECONDARY)
    add_line(s, Inches(0.55), Inches(3.7), Inches(2.0), Inches(3.7),
             COLOR_ACCENT_OPT, 2.0)

    # 4 个阶段块
    stages = [
        ("启蒙时期", "16-18 世纪", "培根、笛卡尔等高扬理性与科学",
         COLOR_ACCENT_OPT),
        ("实证主义", "19 世纪",   "圣西门、孔德主张科学治理社会",
         RGBColor(0xD3, 0x6B, 0x1F)),
        ("工业时代", "20 世纪初", "对工业文明与技术进步的礼赞",
         RGBColor(0xB8, 0x57, 0x12)),
        ("信息时代", "20 世纪后期", "贝尔、托夫勒等预言后工业／信息文明",
         RGBColor(0x9E, 0x46, 0x0A)),
    ]
    box_w = Inches(2.95)
    box_h = Inches(2.6)
    gap = Inches(0.15)
    start_x = Inches(0.55)
    for i, (title, era, desc, c) in enumerate(stages):
        x = start_x + (box_w + gap) * i
        y = Inches(3.95)
        add_filled_rect(s, x, y, box_w, box_h,
                        RGBColor(0xFD, 0xF3, 0xE6))
        add_filled_rect(s, x, y, box_w, Inches(0.6), c)
        add_textbox(s, x, y + Inches(0.05), box_w, Inches(0.5),
                    title, size=16, bold=True, color=COLOR_TEXT_LIGHT,
                    align=PP_ALIGN.CENTER)
        add_textbox(s, x, y + Inches(0.7), box_w, Inches(0.4),
                    era, size=12, color=COLOR_TEXT_GRAY,
                    align=PP_ALIGN.CENTER)
        add_textbox(s, x + Inches(0.2), y + Inches(1.15),
                    box_w - Inches(0.4), Inches(1.3),
                    desc, size=13, color=COLOR_TEXT_DARK,
                    align=PP_ALIGN.LEFT)


# ---------- 5. 乐观主义代表人物（一） ----------
def slide_optimism_figures_1():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "三、乐观主义的代表性思想（上）",
                   "Bacon · Saint-Simon · Comte",
                   page_no=3, total=TOTAL)

    figures = [
        ("弗朗西斯 · 培根", "Francis Bacon (1561-1626)",
         "「知识就是力量」",
         "在《新工具》《新大西岛》中提出依靠科学实验"
         "认识自然、利用自然，构想了以科学技术为核心"
         "的理想国 ——「本撒勒姆岛」，开启近代科技乐观"
         "主义的先河。"),
        ("圣西门 / 孔德", "Saint-Simon · Auguste Comte",
         "「实证主义 · 科学主义」",
         "认为社会发展遵循从「神学—形而上学—实证」"
         "的规律；提倡由科学家与实业家治理社会，"
         "把科学方法推广到一切领域，相信科学能彻底"
         "解决社会问题。"),
    ]
    box_w = Inches(6.0)
    box_h = Inches(5.3)
    for i, (name, en, slogan, desc) in enumerate(figures):
        x = Inches(0.55 + i * 6.4)
        y = Inches(1.7)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        add_filled_rect(s, x, y, box_w, Inches(1.1), COLOR_ACCENT_OPT)
        add_textbox(s, x + Inches(0.3), y + Inches(0.15),
                    box_w - Inches(0.6), Inches(0.55),
                    name, size=22, bold=True, color=COLOR_TEXT_LIGHT,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.3), y + Inches(0.7),
                    box_w - Inches(0.6), Inches(0.35),
                    en, size=12, color=COLOR_TEXT_LIGHT)
        # slogan
        add_textbox(s, x + Inches(0.3), y + Inches(1.3),
                    box_w - Inches(0.6), Inches(0.5),
                    slogan, size=18, bold=True, color=COLOR_PRIMARY)
        add_line(s, x + Inches(0.3), y + Inches(1.85),
                 x + Inches(2.5), y + Inches(1.85),
                 COLOR_LINE, 1.5)
        # desc
        add_textbox(s, x + Inches(0.3), y + Inches(2.0),
                    box_w - Inches(0.6), box_h - Inches(2.2),
                    desc, size=14, color=COLOR_TEXT_DARK)


# ---------- 6. 乐观主义代表人物（二） ----------
def slide_optimism_figures_2():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "三、乐观主义的代表性思想（下）",
                   "Bell · Toffler · Naisbitt · Kurzweil",
                   page_no=4, total=TOTAL)

    figures = [
        ("丹尼尔 · 贝尔",
         "《后工业社会的来临》",
         "认为知识与理论科学已成为社会的「中轴原则」，"
         "技术进步将催生一个由专业技术阶级主导的、"
         "以服务业为主的后工业社会。"),
        ("阿尔文 · 托夫勒",
         "《第三次浪潮》",
         "把人类文明划分为农业、工业、信息三次浪潮；"
         "认为以计算机和电信为代表的新技术革命，"
         "将塑造去中心化、个性化的崭新文明。"),
        ("约翰 · 奈斯比特",
         "《大趋势》",
         "预言由工业社会向信息社会的全球转型，"
         "技术进步带来民主化、网络化、全球化等"
         "积极变化。"),
        ("雷 · 库兹韦尔",
         "《奇点临近》",
         "认为技术发展呈指数增长，2045 年前后将到达"
         "「技术奇点」，人工智能与人脑融合，"
         "人类将进入超级智能时代。"),
    ]
    cols, rows = 2, 2
    box_w = Inches(6.0)
    box_h = Inches(2.55)
    for i, (name, work, desc) in enumerate(figures):
        col = i % cols
        row = i // cols
        x = Inches(0.55 + col * 6.4)
        y = Inches(1.7 + row * 2.8)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        add_filled_rect(s, x, y, Inches(0.15), box_h, COLOR_ACCENT_OPT)
        add_textbox(s, x + Inches(0.3), y + Inches(0.1),
                    box_w - Inches(0.6), Inches(0.5),
                    name, size=18, bold=True, color=COLOR_SECONDARY,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.3), y + Inches(0.65),
                    box_w - Inches(0.6), Inches(0.4),
                    work, size=14, bold=True, color=COLOR_ACCENT_OPT)
        add_textbox(s, x + Inches(0.3), y + Inches(1.1),
                    box_w - Inches(0.6), box_h - Inches(1.2),
                    desc, size=13, color=COLOR_TEXT_DARK)


# ---------- 7. 乐观主义核心观点 ----------
def slide_optimism_core():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "三、科技乐观主义的核心观点",
                   "Core Theses",
                   page_no=5, total=TOTAL)

    items = [
        ("技术决定论：",
         "把科学技术视为推动历史发展的最终决定力量。"),
        ("线性进步观：",
         "认为科技与社会进步呈单向、累积、不可逆的发展。"),
        ("理性万能论：",
         "相信科学理性可以解决一切自然与社会问题。"),
        ("乌托邦愿景：",
         "把科技进步与人类幸福、自由、解放直接画等号。"),
        ("增长无限论：",
         "默认资源、生态、人的承受能力是无限的。"),
    ]
    add_bullets(s, Inches(0.6), Inches(1.8), Inches(12.2), Inches(4.5),
                items, size=18, line_spacing=1.7)

    # 底部小结条
    add_filled_rect(s, Inches(0.55), Inches(6.05), Inches(12.3),
                    Inches(0.85), COLOR_ACCENT_OPT)
    add_textbox(s, Inches(0.85), Inches(6.15), Inches(12), Inches(0.65),
                "总评：科技乐观主义高扬人的主体性与创造力，"
                "却往往把科技的社会效应抽象化、绝对化。",
                size=16, bold=True, color=COLOR_TEXT_LIGHT,
                anchor=MSO_ANCHOR.MIDDLE)


# ---------- 8. 悲观主义概述 ----------
def slide_pessimism_overview():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "四、科技悲观主义：概述",
                   "Technological Pessimism: Overview",
                   page_no=6, total=TOTAL)

    # 核心命题
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(12.3), Inches(1.3),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(0.15),
                    Inches(1.3), COLOR_ACCENT_PES)
    add_textbox(s, Inches(0.85), Inches(1.8), Inches(11.5), Inches(0.5),
                "核心命题", size=18, bold=True, color=COLOR_ACCENT_PES)
    add_textbox(s, Inches(0.85), Inches(2.25), Inches(11.8), Inches(0.7),
                "科学技术内含异化与统治的逻辑，将摧毁自然、"
                "瓦解人性、危及人类的生存与意义。",
                size=15, color=COLOR_TEXT_DARK)

    # 历史脉络
    add_textbox(s, Inches(0.55), Inches(3.2), Inches(12), Inches(0.5),
                "历史脉络", size=20, bold=True, color=COLOR_SECONDARY)
    add_line(s, Inches(0.55), Inches(3.7), Inches(2.0), Inches(3.7),
             COLOR_ACCENT_PES, 2.0)

    stages = [
        ("浪漫主义批判", "18 世纪",
         "卢梭批评科学艺术败坏道德",
         RGBColor(0x4A, 0x60, 0x7A)),
        ("现代性反思", "20 世纪上半叶",
         "海德格尔追问技术的本质",
         RGBColor(0x3D, 0x52, 0x6B)),
        ("法兰克福学派", "20 世纪中叶",
         "工具理性与单向度的人",
         RGBColor(0x30, 0x44, 0x5C)),
        ("生态主义", "20 世纪后期",
         "罗马俱乐部呼吁增长极限",
         RGBColor(0x23, 0x36, 0x4D)),
    ]
    box_w = Inches(2.95)
    box_h = Inches(2.6)
    gap = Inches(0.15)
    for i, (title, era, desc, c) in enumerate(stages):
        x = Inches(0.55) + (box_w + gap) * i
        y = Inches(3.95)
        add_filled_rect(s, x, y, box_w, box_h,
                        RGBColor(0xEC, 0xEF, 0xF3))
        add_filled_rect(s, x, y, box_w, Inches(0.6), c)
        add_textbox(s, x, y + Inches(0.05), box_w, Inches(0.5),
                    title, size=15, bold=True, color=COLOR_TEXT_LIGHT,
                    align=PP_ALIGN.CENTER)
        add_textbox(s, x, y + Inches(0.7), box_w, Inches(0.4),
                    era, size=12, color=COLOR_TEXT_GRAY,
                    align=PP_ALIGN.CENTER)
        add_textbox(s, x + Inches(0.2), y + Inches(1.15),
                    box_w - Inches(0.4), Inches(1.3),
                    desc, size=13, color=COLOR_TEXT_DARK,
                    align=PP_ALIGN.LEFT)


# ---------- 9. 悲观主义代表人物（一） ----------
def slide_pessimism_figures_1():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "四、悲观主义的代表性思想（上）",
                   "Rousseau · Heidegger",
                   page_no=7, total=TOTAL)

    figures = [
        ("让-雅克 · 卢梭",
         "Jean-Jacques Rousseau (1712-1778)",
         "「科学与艺术败坏道德」",
         "在《论科学与艺术》中提出：科学技术的进步并未"
         "净化道德，反而使人远离自然本性、加深贫富分化、"
         "强化奢靡与虚荣，是科技悲观主义最早的系统表达。"),
        ("马丁 · 海德格尔",
         "Martin Heidegger (1889-1976)",
         "「技术之本质，乃是集置（Gestell）」",
         "在《技术的追问》中区分「技术」与「技术之本质」："
         "现代技术把万物（包括人）摆置为可被计算和支配的"
         "「持存物」，遮蔽了存在的真理，是最大的危险，"
         "也是「最高的危险」。"),
    ]
    box_w = Inches(6.0)
    box_h = Inches(5.3)
    for i, (name, en, slogan, desc) in enumerate(figures):
        x = Inches(0.55 + i * 6.4)
        y = Inches(1.7)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        add_filled_rect(s, x, y, box_w, Inches(1.1), COLOR_ACCENT_PES)
        add_textbox(s, x + Inches(0.3), y + Inches(0.15),
                    box_w - Inches(0.6), Inches(0.55),
                    name, size=22, bold=True, color=COLOR_TEXT_LIGHT,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.3), y + Inches(0.7),
                    box_w - Inches(0.6), Inches(0.35),
                    en, size=12, color=COLOR_TEXT_LIGHT)
        add_textbox(s, x + Inches(0.3), y + Inches(1.3),
                    box_w - Inches(0.6), Inches(0.55),
                    slogan, size=17, bold=True, color=COLOR_PRIMARY)
        add_line(s, x + Inches(0.3), y + Inches(1.95),
                 x + Inches(2.5), y + Inches(1.95),
                 COLOR_LINE, 1.5)
        add_textbox(s, x + Inches(0.3), y + Inches(2.1),
                    box_w - Inches(0.6), box_h - Inches(2.3),
                    desc, size=14, color=COLOR_TEXT_DARK)


# ---------- 10. 悲观主义代表人物（二） ----------
def slide_pessimism_figures_2():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "四、悲观主义的代表性思想（下）",
                   "Frankfurt School · Ellul · Mumford · Club of Rome",
                   page_no=8, total=TOTAL)

    figures = [
        ("法兰克福学派",
         "霍克海默、阿多诺、马尔库塞",
         "《启蒙辩证法》《单向度的人》揭示：启蒙理性"
         "蜕变为工具理性；技术成为新的意识形态，"
         "把人塑造为缺乏批判维度的「单向度的人」。"),
        ("雅克 · 埃吕尔",
         "Jacques Ellul",
         "《技术社会》提出技术具有自主性，已发展为"
         "支配一切领域的封闭系统；人不再使用技术，"
         "而是被技术所使用。"),
        ("刘易斯 · 芒福德",
         "Lewis Mumford",
         "《机器的神话》区分「多元技术」与「巨型技术」；"
         "现代巨型机器把人变为零件，是对人性的根本"
         "压抑。"),
        ("罗马俱乐部",
         "《增长的极限》(1972)",
         "运用系统动力学模型预警：在资源、人口、"
         "污染等限制下，无限制的工业增长将导致"
         "全球性生态崩溃。"),
    ]
    cols = 2
    box_w = Inches(6.0)
    box_h = Inches(2.55)
    for i, (name, work, desc) in enumerate(figures):
        col = i % cols
        row = i // cols
        x = Inches(0.55 + col * 6.4)
        y = Inches(1.7 + row * 2.8)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        add_filled_rect(s, x, y, Inches(0.15), box_h, COLOR_ACCENT_PES)
        add_textbox(s, x + Inches(0.3), y + Inches(0.1),
                    box_w - Inches(0.6), Inches(0.5),
                    name, size=18, bold=True, color=COLOR_SECONDARY,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.3), y + Inches(0.65),
                    box_w - Inches(0.6), Inches(0.4),
                    work, size=14, bold=True, color=COLOR_ACCENT_PES)
        add_textbox(s, x + Inches(0.3), y + Inches(1.1),
                    box_w - Inches(0.6), box_h - Inches(1.2),
                    desc, size=13, color=COLOR_TEXT_DARK)


# ---------- 11. 悲观主义核心观点 ----------
def slide_pessimism_core():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "四、科技悲观主义的核心观点",
                   "Core Theses",
                   page_no=9, total=TOTAL)

    items = [
        ("反技术决定论的悖论：",
         "在批判技术决定论的同时，又陷入「技术自主」的另一种决定论。"),
        ("工具理性批判：",
         "认为现代科技以效率与控制为旨归，压抑了价值理性。"),
        ("异化与统治论：",
         "科技不仅异化劳动，更异化人本身，成为新的统治形式。"),
        ("生态危机论：",
         "科技驱动的工业增长不可持续，必将引发生态崩溃。"),
        ("意义丧失论：",
         "科技世界遮蔽了人的本真存在，导致价值虚无与精神危机。"),
    ]
    add_bullets(s, Inches(0.6), Inches(1.8), Inches(12.2), Inches(4.5),
                items, size=18, line_spacing=1.7,
                bullet_color=COLOR_ACCENT_PES)

    add_filled_rect(s, Inches(0.55), Inches(6.05), Inches(12.3),
                    Inches(0.85), COLOR_ACCENT_PES)
    add_textbox(s, Inches(0.85), Inches(6.15), Inches(12), Inches(0.65),
                "总评：科技悲观主义揭示了科技的负面后果，"
                "却把这些后果归咎于科技本身，从而走向另一极端。",
                size=16, bold=True, color=COLOR_TEXT_LIGHT,
                anchor=MSO_ANCHOR.MIDDLE)


# ---------- 12. 两者对比 ----------
def slide_comparison():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "五、两种科技价值观的对比",
                   "Optimism vs. Pessimism",
                   page_no=10, total=TOTAL)

    rows = [
        ("比较维度", "科技乐观主义", "科技悲观主义"),
        ("科技本质", "解放力量、进步引擎",
         "异化力量、统治装置"),
        ("社会效应", "自动带来福祉与自由",
         "必然带来奴役与危机"),
        ("人与科技", "人是科技的主人",
         "人被科技所支配"),
        ("发展趋势", "线性、累积、向上",
         "失控、不可持续"),
        ("方法论", "形而上学的肯定",
         "形而上学的否定"),
        ("共同缺陷", "脱离社会制度抽象谈论科技",
         "只见科技、不见生产关系"),
    ]

    table_x = Inches(0.55)
    table_y = Inches(1.75)
    col_w = [Inches(2.7), Inches(4.8), Inches(4.8)]
    row_h = Inches(0.68)

    # 表头
    cur_x = table_x
    for j, header in enumerate(rows[0]):
        bg = COLOR_SECONDARY if j == 0 else (
            COLOR_ACCENT_OPT if j == 1 else COLOR_ACCENT_PES)
        add_filled_rect(s, cur_x, table_y, col_w[j], row_h, bg)
        add_textbox(s, cur_x, table_y, col_w[j], row_h,
                    header, size=16, bold=True, color=COLOR_TEXT_LIGHT,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cur_x += col_w[j]

    # 行
    for i, row in enumerate(rows[1:]):
        cur_y = table_y + row_h * (i + 1)
        bg_row = COLOR_LIGHT_BG if i % 2 == 0 else \
            RGBColor(0xFF, 0xFF, 0xFF)
        cur_x = table_x
        for j, cell in enumerate(row):
            add_filled_rect(s, cur_x, cur_y, col_w[j], row_h, bg_row,
                            line_color=RGBColor(0xDD, 0xD5, 0xC8))
            text_color = COLOR_SECONDARY if j == 0 else COLOR_TEXT_DARK
            bold = (j == 0)
            size = 14 if j == 0 else 13
            align = PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT
            # 内边距
            tb = s.shapes.add_textbox(
                cur_x + (Emu(0) if j == 0 else Inches(0.2)),
                cur_y, col_w[j], row_h)
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Emu(0)
            tf.margin_right = Emu(0)
            p = tf.paragraphs[0]
            p.alignment = align
            run = p.add_run()
            run.text = cell
            set_run_font(run, size=size, bold=bold, color=text_color)
            cur_x += col_w[j]


# ---------- 13. 马克思主义科技价值观 ----------
def slide_marxism_view():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "六、马克思主义的科技价值观",
                   "Marxist View on Science & Technology",
                   page_no=11, total=TOTAL)

    # 顶部金句
    add_filled_rect(s, Inches(0.55), Inches(1.65), Inches(12.3),
                    Inches(0.95), COLOR_PRIMARY)
    add_textbox(s, Inches(0.85), Inches(1.7), Inches(12), Inches(0.45),
                "「科学技术是第一生产力。」  ——  邓小平",
                size=18, bold=True, color=COLOR_TEXT_LIGHT)
    add_textbox(s, Inches(0.85), Inches(2.1), Inches(12), Inches(0.45),
                "「资本主义生产第一次把自然科学置于直接生产过程的"
                "服务之中。」  ——  马克思",
                size=14, color=COLOR_TEXT_LIGHT)

    # 四个要点（2x2 卡片）
    items = [
        ("生产力论",
         "科技是推动生产力发展的革命性力量，是人类"
         "改造自然、解放自身的根本手段。"),
        ("二重性论",
         "在资本主义制度下，科技既是解放力量，"
         "又被资本利用为剥削、异化和统治的工具。"),
        ("社会建构论",
         "科技不是脱离社会的中立力量，其方向与后果"
         "受生产关系、阶级利益、国家制度制约。"),
        ("以人为本",
         "科技进步的根本目的，是促进人的自由全面发展，"
         "而不是相反；技术批判最终是制度批判。"),
    ]
    cols = 2
    box_w = Inches(6.0)
    box_h = Inches(1.95)
    for i, (h, t) in enumerate(items):
        col = i % cols
        row = i // cols
        x = Inches(0.55 + col * 6.4)
        y = Inches(2.85 + row * 2.1)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        add_filled_rect(s, x, y, Inches(0.15), box_h, COLOR_PRIMARY)
        # 编号圆
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                    x + Inches(0.4), y + Inches(0.25),
                                    Inches(0.55), Inches(0.55))
        circle.fill.solid()
        circle.fill.fore_color.rgb = COLOR_PRIMARY
        circle.line.fill.background()
        tf = circle.text_frame
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(i + 1)
        set_run_font(run, size=18, bold=True, color=COLOR_TEXT_LIGHT)

        add_textbox(s, x + Inches(1.1), y + Inches(0.25),
                    box_w - Inches(1.3), Inches(0.5),
                    h, size=18, bold=True, color=COLOR_PRIMARY,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.4), y + Inches(0.95),
                    box_w - Inches(0.6), box_h - Inches(1.0),
                    t, size=13, color=COLOR_TEXT_DARK)


# ---------- 14. 评价乐观主义 ----------
def slide_eval_optimism():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "七、马克思主义对乐观主义的评价",
                   "Critique of Technological Optimism",
                   page_no=12, total=TOTAL)

    # 左列：合理性
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(5.2),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(0.6),
                    RGBColor(0x2E, 0x86, 0x4F))
    add_textbox(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(0.6),
                "✓  合理性 · 可肯定之处",
                size=18, bold=True, color=COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    items_l = [
        "肯定了科学技术作为生产力的革命性作用，"
        "与马克思主义「科技是第一生产力」相通。",
        "凸显了人的主体性、能动性与创造能力。",
        "鼓舞人类积极利用科技解决现实问题，"
        "推动文明的不断进步。",
    ]
    add_bullets(s, Inches(0.85), Inches(2.5), Inches(5.4), Inches(4.3),
                items_l, size=14, bullet_color=RGBColor(0x2E, 0x86, 0x4F),
                line_spacing=1.45)

    # 右列：局限性
    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(5.2),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(0.6),
                    COLOR_PRIMARY)
    add_textbox(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(0.6),
                "✗  局限性 · 应批判之处",
                size=18, bold=True, color=COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    items_r = [
        "陷入技术决定论，把科技抽象为脱离社会关系的"
        "独立力量。",
        "回避资本主义生产关系对科技的支配，掩盖了科技"
        "异化的制度根源。",
        "忽视科技发展的代价：生态破坏、贫富分化、"
        "新型不平等等。",
        "把科技进步等同于人的解放，混淆了「物的发展」"
        "与「人的发展」。",
    ]
    add_bullets(s, Inches(7.15), Inches(2.5), Inches(5.4), Inches(4.3),
                items_r, size=14, bullet_color=COLOR_PRIMARY,
                line_spacing=1.45)

    # 底部结论条
    add_filled_rect(s, Inches(0.55), Inches(7.0), Inches(12.3),
                    Inches(0.05), COLOR_PRIMARY)


# ---------- 15. 评价悲观主义 ----------
def slide_eval_pessimism():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "七、马克思主义对悲观主义的评价",
                   "Critique of Technological Pessimism",
                   page_no=13, total=TOTAL)

    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(5.2),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(0.6),
                    RGBColor(0x2E, 0x86, 0x4F))
    add_textbox(s, Inches(0.55), Inches(1.7), Inches(6.0), Inches(0.6),
                "✓  合理性 · 可肯定之处",
                size=18, bold=True, color=COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    items_l = [
        "深刻揭示了科技在资本主义条件下的异化现象，"
        "与马克思的异化劳动思想存在共鸣。",
        "警示人类反思工具理性的霸权，关注人的价值理性"
        "与精神世界。",
        "促使社会正视科技的生态、伦理与社会风险，"
        "推动可持续发展观念的形成。",
    ]
    add_bullets(s, Inches(0.85), Inches(2.5), Inches(5.4), Inches(4.3),
                items_l, size=14, bullet_color=RGBColor(0x2E, 0x86, 0x4F),
                line_spacing=1.45)

    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(5.2),
                    COLOR_LIGHT_BG)
    add_filled_rect(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(0.6),
                    COLOR_PRIMARY)
    add_textbox(s, Inches(6.85), Inches(1.7), Inches(6.0), Inches(0.6),
                "✗  局限性 · 应批判之处",
                size=18, bold=True, color=COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    items_r = [
        "把科技异化的根源归于科技本身，未抓住资本主义"
        "生产关系这一真正病灶。",
        "倾向于把科技实体化、神秘化，陷入「技术自主」"
        "的形而上学。",
        "对科技进步的全盘否定，易滑向反现代化、"
        "反工业化的历史虚无主义。",
        "缺乏现实可行的解决方案，对劳动人民的解放"
        "缺乏指导意义。",
    ]
    add_bullets(s, Inches(7.15), Inches(2.5), Inches(5.4), Inches(4.3),
                items_r, size=14, bullet_color=COLOR_PRIMARY,
                line_spacing=1.45)

    add_filled_rect(s, Inches(0.55), Inches(7.0), Inches(12.3),
                    Inches(0.05), COLOR_PRIMARY)


# ---------- 16. 启示与结论 ----------
def slide_conclusion():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "八、现实启示与结论",
                   "Implications & Conclusion",
                   page_no=14, total=TOTAL)

    # 上方核心结论
    add_filled_rect(s, Inches(0.55), Inches(1.7), Inches(12.3),
                    Inches(1.4), COLOR_PRIMARY)
    add_textbox(s, Inches(0.85), Inches(1.85), Inches(12), Inches(0.55),
                "核心结论",
                size=18, bold=True, color=COLOR_TEXT_LIGHT)
    add_textbox(s, Inches(0.85), Inches(2.35), Inches(12), Inches(0.7),
                "乐观主义与悲观主义都是对科技的「片面真理」。"
                "唯有站在马克思主义的立场上——把科技放回到"
                "社会生产关系中考察，才能超越二者的对立。",
                size=15, color=COLOR_TEXT_LIGHT)

    # 四点启示
    items = [
        ("辩证看待科技",
         "既肯定科技进步的伟大意义，又警惕其异化风险。"),
        ("制度比技术更根本",
         "解决科技负面问题的根本途径是变革不合理的"
         "社会制度。"),
        ("坚持以人民为中心",
         "科技发展应服务于人的自由全面发展，"
         "而非资本增殖。"),
        ("走可持续发展之路",
         "处理好科技、人、自然三者的辩证关系，"
         "构建人与自然生命共同体。"),
    ]
    cols = 2
    box_w = Inches(6.0)
    box_h = Inches(1.75)
    for i, (h, t) in enumerate(items):
        col = i % cols
        row = i // cols
        x = Inches(0.55 + col * 6.4)
        y = Inches(3.3 + row * 1.85)
        add_filled_rect(s, x, y, box_w, box_h, COLOR_LIGHT_BG)
        # 顶部色条
        add_filled_rect(s, x, y, Inches(0.5), box_h, COLOR_PRIMARY)
        add_textbox(s, x + Inches(0.05), y + Inches(0.4),
                    Inches(0.4), Inches(1.0),
                    f"{i+1:02d}", size=22, bold=True,
                    color=COLOR_TEXT_LIGHT, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, x + Inches(0.7), y + Inches(0.2),
                    box_w - Inches(0.9), Inches(0.5),
                    h, size=17, bold=True, color=COLOR_PRIMARY,
                    font=CN_FONT_HEAD)
        add_textbox(s, x + Inches(0.7), y + Inches(0.8),
                    box_w - Inches(0.9), box_h - Inches(0.9),
                    t, size=13, color=COLOR_TEXT_DARK)


# ---------- 17. 参考文献 ----------
def slide_references():
    s = prs.slides.add_slide(blank)
    make_title_bar(s, "参考文献",
                   "References",
                   page_no=15, total=TOTAL)

    refs = [
        "[1]  马克思.  资本论（第一卷）[M].  北京：人民出版社，2004.",
        "[2]  马克思，恩格斯.  马克思恩格斯文集（1—10卷）[M]."
        "  北京：人民出版社，2009.",
        "[3]  邓小平.  邓小平文选（第三卷）[M].  北京：人民出版社，1993.",
        "[4]  海德格尔.  演讲与论文集[M].  孙周兴 译."
        "  北京：商务印书馆，2018.",
        "[5]  马尔库塞.  单向度的人[M].  刘继 译."
        "  上海：上海译文出版社，2008.",
        "[6]  阿尔文 · 托夫勒.  第三次浪潮[M].  黄明坚 译."
        "  北京：中信出版社，2018.",
        "[7]  丹尼斯 · 米都斯.  增长的极限[M].  李宝恒 译."
        "  长春：吉林人民出版社，1997.",
        "[8]  陈昌曙.  技术哲学引论[M].  北京：科学出版社，2012.",
        "[9]  吴国盛.  科学的历程[M].  长沙：湖南科学技术出版社，2018.",
    ]
    add_bullets(s, Inches(0.7), Inches(1.85), Inches(12.2),
                Inches(5.0), refs, size=14, line_spacing=1.5,
                bullet_color=COLOR_PRIMARY)


# ---------- 18. 谢谢 ----------
def slide_thanks():
    s = prs.slides.add_slide(blank)
    add_filled_rect(s, Emu(0), Emu(0), SLIDE_W, SLIDE_H, COLOR_SECONDARY)
    # 装饰
    add_filled_rect(s, Inches(0), Inches(3.6), SLIDE_W, Inches(0.04),
                    COLOR_PRIMARY)

    add_textbox(s, Inches(0), Inches(2.4), SLIDE_W, Inches(1.2),
                "T H A N K S",
                size=72, bold=True, color=COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER, font=CN_FONT_HEAD)
    add_textbox(s, Inches(0), Inches(3.85), SLIDE_W, Inches(0.7),
                "感谢聆听  ·  敬请批评指正",
                size=24, color=COLOR_LIGHT_BG,
                align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0), Inches(4.7), SLIDE_W, Inches(0.5),
                "Q & A",
                size=18, color=COLOR_LINE, align=PP_ALIGN.CENTER)


# ============== 生成 ==============
slide_cover()
slide_toc()
slide_intro()
slide_optimism_overview()
slide_optimism_figures_1()
slide_optimism_figures_2()
slide_optimism_core()
slide_pessimism_overview()
slide_pessimism_figures_1()
slide_pessimism_figures_2()
slide_pessimism_core()
slide_comparison()
slide_marxism_view()
slide_eval_optimism()
slide_eval_pessimism()
slide_conclusion()
slide_references()
slide_thanks()

OUTPUT = "悲观主义与乐观主义的科技价值观及其马克思主义评价.pptx"
prs.save(OUTPUT)
print(f"✅  生成完毕：{OUTPUT}")
print(f"    幻灯片张数：{len(prs.slides)}")
