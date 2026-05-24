# -*- coding: utf-8 -*-
"""
《科技观》课程作业 PPT 生成脚本
主题：影响科学发展的社会因素及促进我国科技进步的路径思考
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

# ===================== 主题配色 =====================
DEEP_BLUE = RGBColor(0x1F, 0x3A, 0x68)   # 主色：深蓝
TECH_ORANGE = RGBColor(0xFF, 0x8C, 0x42)  # 强调色：科技橙
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)   # 背景灰
DARK_TEXT = RGBColor(0x2C, 0x2C, 0x2C)    # 正文黑
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT_BLUE = RGBColor(0xE8, 0xEF, 0xF7)    # 浅蓝底
ACCENT_GREEN = RGBColor(0x2E, 0xA8, 0x7E)

CN_TITLE_FONT = "微软雅黑"
CN_BODY_FONT = "微软雅黑"

# ===================== 创建演示文稿 =====================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ===================== 通用工具 =====================
def set_slide_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    # 移到最底层
    spTree = bg._element.getparent()
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)
    return bg


def add_text(slide, left, top, width, height, text, *,
             font=CN_BODY_FONT, size=14, bold=False, color=DARK_TEXT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font
        # 中文字体
        rPr = run._r.get_or_add_rPr()
        eastAsia = rPr.find(qn('a:eastAsia'))
        if eastAsia is None:
            eastAsia = etree.SubElement(rPr, qn('a:eastAsia'))
        rFonts = rPr.find(qn('a:rFonts'))
        if rFonts is None:
            rFonts = etree.SubElement(rPr, qn('a:rFonts'))
        rFonts.set('eastAsia', font)
        rFonts.set('ascii', font)
        rFonts.set('hAnsi', font)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return tb


def add_rect(slide, left, top, width, height, fill=None, line=None,
             shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, left, top, width, height)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s


def add_title_bar(slide, title, subtitle=None):
    """页面顶部标题装饰"""
    # 左侧色块
    add_rect(slide, Inches(0.5), Inches(0.45), Inches(0.18), Inches(0.55), fill=TECH_ORANGE)
    # 标题
    add_text(slide, Inches(0.78), Inches(0.4), Inches(11), Inches(0.6),
             title, font=CN_TITLE_FONT, size=28, bold=True, color=DEEP_BLUE,
             anchor=MSO_ANCHOR.MIDDLE)
    # 副标题
    if subtitle:
        add_text(slide, Inches(0.78), Inches(1.0), Inches(11), Inches(0.35),
                 subtitle, size=13, color=RGBColor(0x6B, 0x72, 0x80))
    # 底部分隔线
    line = add_rect(slide, Inches(0.5), Inches(1.45), Inches(12.3), Emu(15000), fill=DEEP_BLUE)
    return


def add_footer(slide, page_num, total=15):
    add_text(slide, Inches(0.5), Inches(7.05), Inches(8), Inches(0.3),
             "《科技观》课程作业  |  影响科学发展的社会因素及促进我国科技进步的路径",
             size=9, color=RGBColor(0x9C, 0xA3, 0xAF))
    add_text(slide, Inches(11.8), Inches(7.05), Inches(1.0), Inches(0.3),
             f"{page_num} / {total}", size=9, color=RGBColor(0x9C, 0xA3, 0xAF),
             align=PP_ALIGN.RIGHT)


# ===================== 第 1 页：封面 =====================
def slide_cover():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, DEEP_BLUE)

    # 装饰几何
    add_rect(s, Inches(0), Inches(0), Inches(4.5), Inches(7.5),
             fill=RGBColor(0x17, 0x2C, 0x50))
    add_rect(s, Inches(0), Inches(5.2), Inches(13.333), Inches(0.08),
             fill=TECH_ORANGE)
    # 装饰圆
    c1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.3), Inches(0.5),
                            Inches(1.6), Inches(1.6))
    c1.fill.solid(); c1.fill.fore_color.rgb = TECH_ORANGE
    c1.line.fill.background()
    c2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(12.3), Inches(1.4),
                            Inches(0.8), Inches(0.8))
    c2.fill.solid(); c2.fill.fore_color.rgb = WHITE
    c2.line.fill.background()

    # 中文主标题
    add_text(s, Inches(0.8), Inches(2.1), Inches(11.5), Inches(1.2),
             "影响科学发展的社会因素",
             size=48, bold=True, color=WHITE, font=CN_TITLE_FONT)
    add_text(s, Inches(0.8), Inches(3.1), Inches(11.5), Inches(1.2),
             "及促进我国科技进步的路径思考",
             size=42, bold=True, color=TECH_ORANGE, font=CN_TITLE_FONT)

    # 英文副标题
    add_text(s, Inches(0.8), Inches(4.4), Inches(11.5), Inches(0.5),
             "Social Factors in Scientific Development & Pathways to Advance China's Sci-Tech Progress",
             size=14, color=RGBColor(0xCC, 0xD6, 0xE6), font="Calibri")

    # 课程信息
    add_text(s, Inches(0.8), Inches(5.6), Inches(7), Inches(0.4),
             "《科技观》课程  |  个人作业",
             size=18, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(6.1), Inches(11), Inches(0.4),
             "姓    名：____________     学    号：____________",
             size=14, color=RGBColor(0xCC, 0xD6, 0xE6))
    add_text(s, Inches(0.8), Inches(6.5), Inches(11), Inches(0.4),
             "指导教师：____________     日    期：2026 年 5 月",
             size=14, color=RGBColor(0xCC, 0xD6, 0xE6))


# ===================== 第 2 页：引言 =====================
def slide_intro():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "引言：为什么研究这个问题？", "Introduction")

    # 引言金句
    quote_box = add_rect(s, Inches(0.7), Inches(1.85), Inches(11.9), Inches(1.0),
                         fill=SOFT_BLUE)
    add_text(s, Inches(1.0), Inches(1.95), Inches(11.5), Inches(0.8),
             '"科学不是孤立的知识体系，而是镶嵌在社会之中的实践活动。"',
             size=20, bold=True, color=DEEP_BLUE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(1.0), Inches(2.45), Inches(11.5), Inches(0.4),
             "—— 贝尔纳《科学的社会功能》",
             size=12, color=RGBColor(0x6B, 0x72, 0x80),
             align=PP_ALIGN.RIGHT)

    # 三个核心要点（卡片式）
    points = [
        ("01", "不只是个人努力", "科学发展不只取决于科学家个人的天才与勤奋，\n还深深嵌入社会结构与历史条件之中。"),
        ("02", "科学是社会建制", "现代科学已成为庞大的社会建制：科研机构、\n资助体系、人才培养、传播评价共同构成生态。"),
        ("03", "现实意义重大", "中美科技博弈、第四次工业革命背景下，\n研究'社会因素 → 科技发展'机制具有迫切意义。"),
    ]
    card_w = Inches(3.9)
    card_h = Inches(2.8)
    gap = Inches(0.18)
    start_x = Inches(0.7)
    top = Inches(3.3)
    for i, (num, head, body) in enumerate(points):
        x = start_x + (card_w + gap) * i
        # 卡片
        add_rect(s, x, top, card_w, card_h, fill=WHITE, line=DEEP_BLUE)
        # 顶部色条
        add_rect(s, x, top, card_w, Inches(0.45), fill=DEEP_BLUE)
        add_text(s, x + Inches(0.2), top + Inches(0.05), Inches(1), Inches(0.4),
                 num, size=18, bold=True, color=TECH_ORANGE, font="Arial Black")
        add_text(s, x + Inches(0.9), top + Inches(0.05), card_w - Inches(1), Inches(0.4),
                 head, size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.25), top + Inches(0.7), card_w - Inches(0.5), card_h - Inches(0.85),
                 body, size=13, color=DARK_TEXT)

    add_text(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
             "▶ 本作业从社会学视角出发，分析多重因素并提出促进我国科技进步的对策建议。",
             size=13, bold=True, color=TECH_ORANGE)
    add_footer(s, 2)


# ===================== 第 3 页：社会因素总览 =====================
def slide_overview():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "第一部分  社会因素总览",
                  "Overview of Social Factors Influencing Science")

    # 中心圆：科学发展
    cx, cy = Inches(6.4), Inches(3.6)
    core = s.shapes.add_shape(MSO_SHAPE.OVAL, cx - Inches(1.0), cy - Inches(1.0),
                              Inches(2.0), Inches(2.0))
    core.fill.solid(); core.fill.fore_color.rgb = DEEP_BLUE
    core.line.color.rgb = TECH_ORANGE; core.line.width = Pt(3)
    tf = core.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "科学发展"
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE
    r.font.name = CN_TITLE_FONT

    # 6 个外围因素
    factors = [
        ("政治制度", "政策导向 / 战略支持", -150),
        ("经济基础", "投入能力 / 市场需求", -90),
        ("教育体系", "人才培养 / 知识传播", -30),
        ("文化传统", "科学精神 / 公众素养", 30),
        ("国际环境", "合作交流 / 技术封锁", 90),
        ("法律军事", "知识产权 / 重大需求", 150),
    ]
    import math
    R = Inches(2.7)
    for name, sub, deg in factors:
        rad = math.radians(deg)
        ex = cx + int(R * math.cos(rad))
        ey = cy + int(R * math.sin(rad))
        # 椭圆节点
        node_w, node_h = Inches(2.2), Inches(0.95)
        node = add_rect(s, ex - node_w//2, ey - node_h//2, node_w, node_h,
                        fill=TECH_ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = node.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = name
        r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = WHITE
        r.font.name = CN_TITLE_FONT
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run(); r2.text = sub
        r2.font.size = Pt(10); r2.font.color.rgb = WHITE; r2.font.name = CN_BODY_FONT
        # 连线
        line = s.shapes.add_connector(1, cx, cy, ex, ey)
        line.line.color.rgb = DEEP_BLUE
        line.line.width = Pt(1.5)
        # 把连线放到节点下方
        spTree = line._element.getparent()
        spTree.remove(line._element)
        spTree.insert(2, line._element)

    add_text(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.4),
             "▶ 科学发展是政治、经济、教育、文化、国际、法律军事等多重社会因素相互作用的结果。",
             size=13, bold=True, color=DEEP_BLUE)
    add_footer(s, 3)


# ===================== 通用单因素页面构造 =====================
def slide_factor(title_zh, title_en, page_num, points, case_text, icon_text):
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, title_zh, title_en)

    # 左侧大数字 / 图标
    add_rect(s, Inches(0.7), Inches(1.85), Inches(3.2), Inches(4.9),
             fill=DEEP_BLUE)
    add_text(s, Inches(0.7), Inches(2.0), Inches(3.2), Inches(1.0),
             icon_text, size=64, bold=True, color=TECH_ORANGE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             font="Arial Black")
    add_text(s, Inches(0.7), Inches(3.2), Inches(3.2), Inches(0.5),
             title_zh.split("｜")[-1] if "｜" in title_zh else title_zh.split("：")[-1] if "：" in title_zh else "",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.85), Inches(3.85), Inches(2.9), Inches(2.7),
             case_text, size=13, color=WHITE, anchor=MSO_ANCHOR.TOP)

    # 右侧要点列表
    add_text(s, Inches(4.2), Inches(1.85), Inches(8.6), Inches(0.5),
             "核心要点", size=18, bold=True, color=DEEP_BLUE)
    add_rect(s, Inches(4.2), Inches(2.4), Inches(0.8), Emu(20000), fill=TECH_ORANGE)
    top = Inches(2.55)
    for i, (h, b) in enumerate(points):
        # 序号圆
        num_c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.25), top,
                                    Inches(0.5), Inches(0.5))
        num_c.fill.solid(); num_c.fill.fore_color.rgb = TECH_ORANGE
        num_c.line.fill.background()
        tf = num_c.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(i+1)
        r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = WHITE
        # 标题
        add_text(s, Inches(4.95), top - Inches(0.02), Inches(7.8), Inches(0.4),
                 h, size=15, bold=True, color=DEEP_BLUE)
        add_text(s, Inches(4.95), top + Inches(0.4), Inches(7.8), Inches(0.6),
                 b, size=12, color=DARK_TEXT)
        top += Inches(1.05)

    add_footer(s, page_num)


# ===================== 第 4 页：政治制度 =====================
def slide_factor_politics():
    slide_factor(
        "① 政治制度与科技政策",
        "Political System & Sci-Tech Policy",
        4,
        [
            ("政治稳定是科学持续发展的前提",
             "战乱、动荡常使科研中断；和平稳定的环境是科学家专注研究的基础。"),
            ("国家战略决定资源流向",
             "曼哈顿计划、阿波罗工程、两弹一星——重大战略集中调动全国资源攻关。"),
            ("具体科技政策塑造创新生态",
             "研发税收优惠、国家重点实验室、人才计划、专项基金等政策工具。"),
            ("反例：政治不当干预科学",
             "苏联李森科事件——政治干预导致遗传学倒退数十年，是深刻教训。"),
        ],
        "典型案例\n\n两弹一星精神：\n热爱祖国、无私奉献、\n自力更生、艰苦奋斗、\n大力协同、勇于登攀。\n\n体现了国家战略\n对科技发展的\n决定性作用。",
        "①"
    )


# ===================== 第 5 页：经济基础（含图表）=====================
def slide_factor_economy():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "② 经济发展水平", "Economic Development Level")

    # 左侧要点
    add_text(s, Inches(0.7), Inches(1.85), Inches(6), Inches(0.5),
             "核心要点", size=18, bold=True, color=DEEP_BLUE)
    add_rect(s, Inches(0.7), Inches(2.4), Inches(0.8), Emu(20000), fill=TECH_ORANGE)
    items = [
        ("经济基础决定科研投入能力",
         "全球 R&D 经费 70% 以上集中于美、中、日、德等经济强国。"),
        ("市场需求牵引技术方向",
         "蒸汽机源于矿山排水需求；互联网兴起于信息共享需求。"),
        ("产业结构反向塑造研究方向",
         "制造业升级 → 倒逼基础研究；金融发达 → 推动量化与算法研究。"),
        ("中国数据：投入持续增长",
         "2023 年我国 R&D 经费 3.3 万亿元，占 GDP 2.64%。"),
    ]
    top = Inches(2.55)
    for i, (h, b) in enumerate(items):
        num_c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.75), top,
                                    Inches(0.45), Inches(0.45))
        num_c.fill.solid(); num_c.fill.fore_color.rgb = TECH_ORANGE
        num_c.line.fill.background()
        tf = num_c.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(i+1)
        r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = WHITE
        add_text(s, Inches(1.4), top - Inches(0.02), Inches(5.3), Inches(0.4),
                 h, size=13, bold=True, color=DEEP_BLUE)
        add_text(s, Inches(1.4), top + Inches(0.35), Inches(5.3), Inches(0.6),
                 b, size=11, color=DARK_TEXT)
        top += Inches(1.05)

    # 右侧：柱状图
    add_text(s, Inches(7.2), Inches(1.85), Inches(5.6), Inches(0.4),
             "主要国家 R&D 经费占 GDP 比重 (%)",
             size=14, bold=True, color=DEEP_BLUE, align=PP_ALIGN.CENTER)
    chart_data = CategoryChartData()
    chart_data.categories = ['以色列', '韩国', '美国', '日本', '德国', '中国', 'OECD均值']
    chart_data.add_series('R&D / GDP %', (5.56, 4.93, 3.46, 3.30, 3.13, 2.64, 2.71))
    chart = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(7.2), Inches(2.3), Inches(5.6), Inches(4.3),
        chart_data
    ).chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(10)
    dl.font.bold = True
    dl.font.color.rgb = DEEP_BLUE
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    # 系列颜色
    series = chart.series[0]
    fill = series.format.fill
    fill.solid()
    fill.fore_color.rgb = DEEP_BLUE
    # 突出"中国"那一根
    from pptx.oxml.ns import nsmap
    # 单独给"中国"那一根上色：通过 dPt
    ser_xml = series._element
    nsmap_local = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                   'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    # 中国 idx=5
    dPt_xml = (
        f'<c:dPt xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<c:idx val="5"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
        f'<c:spPr><a:solidFill><a:srgbClr val="FF8C42"/></a:solidFill></c:spPr>'
        f'</c:dPt>'
    )
    dPt = etree.fromstring(dPt_xml)
    # 插入到 ser 元素中（在 cat 之前）
    cat = ser_xml.find(qn('c:cat'))
    if cat is not None:
        cat.addprevious(dPt)
    else:
        ser_xml.append(dPt)

    # 数据来源
    add_text(s, Inches(7.2), Inches(6.6), Inches(5.6), Inches(0.3),
             "数据来源：OECD Main Science and Technology Indicators (2023)",
             size=9, color=RGBColor(0x9C, 0xA3, 0xAF), align=PP_ALIGN.CENTER)
    add_footer(s, 5)


# ===================== 第 6 页：教育体系 =====================
def slide_factor_education():
    slide_factor(
        "③ 教育体系与人才结构",
        "Education System & Talent Structure",
        6,
        [
            ("教育是科学的'再生产机制'",
             "没有持续的高质量教育，就没有源源不断的科研后备力量。"),
            ("规模与质量同样重要",
             "我国高等教育毛入学率已超 60%，但顶尖创新人才仍显不足。"),
            ("创新型 vs 应试型培养方式",
             "应试教育压抑好奇心；创新教育鼓励质疑、探究与跨学科融合。"),
            ("钱学森之问的警示",
             "'为什么我们的学校总是培养不出杰出人才？'值得深思与改革。"),
        ],
        "钱学森之问\n\n2005 年钱学森向\n温家宝总理发问：\n\n'为什么我们的\n学校总是培养\n不出杰出人才？'\n\n直指中国教育与\n科研体制深层问题。",
        "③"
    )


# ===================== 第 7 页：文化传统 =====================
def slide_factor_culture():
    slide_factor(
        "④ 文化传统与社会观念",
        "Cultural Tradition & Social Values",
        7,
        [
            ("科学精神四要素",
             "怀疑精神、实证态度、理性思维、自由探讨——现代科学的基石。"),
            ("中国传统文化的双重影响",
             "重'经世致用'促进了应用技术，但缺乏纯粹科学的形而上传统。"),
            ("李约瑟难题",
             "为何近代科学没有在中国诞生？文化、制度、经济综合作用的结果。"),
            ("公众科学素养：14.14% → 15%",
             "2023 年我国公民科学素质比例 14.14%，'十四五'目标达 15%。"),
        ],
        "李约瑟难题\n\n英国学者李约瑟\n提出：\n\n古代中国科技长期\n领先世界，\n为什么近代科学\n却诞生于欧洲\n而非中国？\n\n是科技史经典命题。",
        "④"
    )


# ===================== 第 8 页：国际环境 =====================
def slide_factor_international():
    slide_factor(
        "⑤ 国际环境与科技交流",
        "International Environment & Sci-Tech Exchange",
        8,
        [
            ("科学是世界性事业",
             "全球化时代，重大科学突破往往依赖跨国合作与开放交流。"),
            ("国际合作加速发现",
             "CERN 大型强子对撞机、ITER 国际热核聚变、人类基因组计划。"),
            ("技术封锁阻碍发展",
             "光刻机、EDA 软件、AI 芯片等领域的出口管制对我国形成压力。"),
            ("人才国际流动的双刃剑",
             "既是引进顶尖智力的机遇，也面临核心人才外流的挑战。"),
        ],
        "中美科技博弈\n\n2018 年起，美国对\n华为、中兴等企业\n实施技术封锁。\n\n半导体、AI、生物\n科技成为博弈焦点。\n\n倒逼我国走\n'科技自立自强'\n之路。",
        "⑤"
    )


# ===================== 第 9 页：其他因素（表格）=====================
def slide_factor_others():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "⑥ 其他重要社会因素", "Other Important Social Factors")

    # 表格
    headers = ["因素", "作用机制", "典型案例", "影响方向"]
    rows = [
        ("法律制度", "知识产权保护激励创新；反垄断维护公平竞争",
         "美国 1980 年《拜杜法案》", "正向"),
        ("军事需求", "国防战略需求催生重大基础研究与工程突破",
         "互联网起源于美国 ARPANET", "正向"),
        ("宗教伦理", "既可促进（如新教伦理），也可压制（如教会审判）",
         "哥白尼日心说受教会压制", "双向"),
        ("媒体传播", "影响公众认知、舆论与政策走向",
         "转基因、核能争议", "双向"),
        ("人口结构", "人口规模与年龄结构影响科研后备与市场规模",
         "少子化对日本科研的长期影响", "双向"),
    ]
    table_left = Inches(0.7)
    table_top = Inches(1.95)
    table_w = Inches(11.9)
    cols_w = [Inches(1.6), Inches(4.2), Inches(4.0), Inches(2.1)]
    row_h = Inches(0.85)

    # 表头
    x = table_left
    for i, h in enumerate(headers):
        cell = add_rect(s, x, table_top, cols_w[i], Inches(0.55), fill=DEEP_BLUE)
        tf = cell.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = h
        r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = WHITE
        r.font.name = CN_TITLE_FONT
        x += cols_w[i]

    # 数据行
    y = table_top + Inches(0.55)
    for ri, row in enumerate(rows):
        bg = SOFT_BLUE if ri % 2 == 0 else WHITE
        x = table_left
        for ci, val in enumerate(row):
            cell = add_rect(s, x, y, cols_w[ci], row_h, fill=bg,
                            line=RGBColor(0xCC, 0xCC, 0xCC))
            tf = cell.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Emu(60000); tf.margin_right = Emu(60000)
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci in (0, 3) else PP_ALIGN.LEFT
            r = p.add_run(); r.text = val
            r.font.size = Pt(12)
            r.font.bold = (ci == 0)
            color = DEEP_BLUE if ci == 0 else DARK_TEXT
            if ci == 3:
                color = ACCENT_GREEN if val == "正向" else TECH_ORANGE
                r.font.bold = True
            r.font.color.rgb = color
            r.font.name = CN_BODY_FONT
            x += cols_w[ci]
        y += row_h

    add_text(s, Inches(0.7), Inches(6.6), Inches(12), Inches(0.4),
             "▶ 各类社会因素相互交织、共同作用，构成科学发展的复杂社会网络。",
             size=13, bold=True, color=TECH_ORANGE)
    add_footer(s, 9)


# ===================== 第 10 页：第二部分总览 =====================
def slide_part2_overview():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "第二部分  促进我国科技进步的战略路径",
                  "Strategic Pathways for China's Sci-Tech Progress")

    # 顶部红色目标横幅
    add_rect(s, Inches(0.7), Inches(1.85), Inches(11.9), Inches(0.85),
             fill=DEEP_BLUE)
    add_text(s, Inches(0.7), Inches(1.85), Inches(11.9), Inches(0.85),
             "战略目标：实现高水平科技自立自强（党的二十大报告）",
             size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 6 项对策卡片
    measures = [
        ("①", "加大研发投入", "提升基础研究占比\n建设国家实验室"),
        ("②", "改革教育体系", "培养创新型人才\n推动 STEM 教育"),
        ("③", "深化体制改革", "破'四唯'评价\n扩大经费自主权"),
        ("④", "营造创新文化", "弘扬科学家精神\n提升科学素养"),
        ("⑤", "攻克卡脖子技术", "新型举国体制\n企业创新主体"),
        ("⑥", "高水平对外开放", "国际大科学计划\n吸引全球人才"),
    ]
    card_w = Inches(3.85)
    card_h = Inches(1.85)
    gap_x = Inches(0.18)
    gap_y = Inches(0.18)
    start_x = Inches(0.7)
    start_y = Inches(2.95)
    for i, (num, title, body) in enumerate(measures):
        col = i % 3
        row = i // 3
        x = start_x + (card_w + gap_x) * col
        y = start_y + (card_h + gap_y) * row
        # 卡片底
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=DEEP_BLUE)
        # 左色块
        add_rect(s, x, y, Inches(0.7), card_h, fill=TECH_ORANGE)
        add_text(s, x, y, Inches(0.7), card_h,
                 num, size=32, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Arial Black")
        # 标题
        add_text(s, x + Inches(0.85), y + Inches(0.15), card_w - Inches(1), Inches(0.45),
                 title, size=16, bold=True, color=DEEP_BLUE)
        # 内容
        add_text(s, x + Inches(0.85), y + Inches(0.65), card_w - Inches(1), Inches(1.1),
                 body, size=12, color=DARK_TEXT)

    add_text(s, Inches(0.7), Inches(6.85), Inches(12), Inches(0.4),
             "▶ 六大举措形成系统合力：投入 + 人才（输入） → 体制 + 文化（环境） → 攻关 + 开放（突破）",
             size=12, bold=True, color=DEEP_BLUE)
    add_footer(s, 10)


# ===================== 通用对策页面 =====================
def slide_measure_page(title, page_num, items, side_title, side_body):
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, title, "")

    # 左侧 2 大对策
    top = Inches(1.95)
    for i, (num, head, points) in enumerate(items):
        # 序号大圆
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.7), top,
                                     Inches(0.9), Inches(0.9))
        circle.fill.solid(); circle.fill.fore_color.rgb = TECH_ORANGE
        circle.line.fill.background()
        tf = circle.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Emu(0); tf.margin_right = Emu(0)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = num
        r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = WHITE
        r.font.name = "Arial Black"
        # 标题
        add_text(s, Inches(1.75), top + Inches(0.05), Inches(7), Inches(0.45),
                 head, size=18, bold=True, color=DEEP_BLUE)
        add_rect(s, Inches(1.75), top + Inches(0.55), Inches(7), Emu(15000),
                 fill=DEEP_BLUE)
        # 要点
        py = top + Inches(0.7)
        for pt in points:
            add_text(s, Inches(1.85), py, Inches(0.3), Inches(0.3),
                     "▸", size=14, bold=True, color=TECH_ORANGE)
            add_text(s, Inches(2.15), py, Inches(6.7), Inches(0.7),
                     pt, size=12, color=DARK_TEXT)
            py += Inches(0.55)
        top += Inches(2.55)

    # 右侧亮点框
    add_rect(s, Inches(9.2), Inches(1.95), Inches(3.6), Inches(4.9),
             fill=DEEP_BLUE)
    add_text(s, Inches(9.2), Inches(2.05), Inches(3.6), Inches(0.5),
             "重点关注", size=15, bold=True, color=TECH_ORANGE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(10.6), Inches(2.55), Inches(0.8), Emu(15000),
             fill=TECH_ORANGE)
    add_text(s, Inches(9.2), Inches(2.7), Inches(3.6), Inches(0.55),
             side_title, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(9.4), Inches(3.4), Inches(3.2), Inches(3.3),
             side_body, size=12, color=WHITE)

    add_footer(s, page_num)


# ===================== 第 11 页：投入 + 人才 =====================
def slide_measure_input():
    slide_measure_page(
        "对策一  投入 + 人才（输入端）",
        11,
        [
            ("1", "持续加大研发投入",
             ["提高基础研究经费占比（目前约 6.65%，发达国家 15-20%）",
              "引导社会资本投入硬科技、长周期领域",
              "建设国家实验室、大科学装置（FAST、托卡马克）"]),
            ("2", "改革教育、培养创新人才",
             ["中小学：强化数理基础与科学兴趣",
              "高校：推动 STEM 教育、新工科、强基计划",
              "完善博士后制度，给青年学者独立研究空间"]),
        ],
        "强基计划",
        "2020 年起教育部\n实施'强基计划'：\n\n选拔有志于服务\n国家重大战略需求\n的拔尖人才。\n\n聚焦数学、物理、\n化学、生物、历史、\n哲学等基础学科。\n\n是培养顶尖创新\n人才的重要尝试。"
    )


# ===================== 第 12 页：体制 + 文化 =====================
def slide_measure_env():
    slide_measure_page(
        "对策二  体制 + 文化（环境端）",
        12,
        [
            ("3", "深化科技体制改革",
             ["改革科研评价：破'四唯'（论文/职称/学历/奖项）",
              "赋予科学家更大经费使用自主权",
              "推动产学研深度融合，畅通成果转化"]),
            ("4", "营造尊重科学的社会文化",
             ["弘扬科学家精神，树立时代偶像（黄大年、南仁东）",
              "提升全民科学素养，加强科普教育",
              "容忍失败、鼓励原创、严惩学术不端"]),
        ],
        "破'四唯'",
        "2018 年起，国务院\n推动破除科研评价\n中的'四唯'：\n\n  ✗ 唯论文\n  ✗ 唯职称\n  ✗ 唯学历\n  ✗ 唯奖项\n\n转向以创新质量、\n学术贡献和实际\n绩效为核心的\n多元评价体系。"
    )


# ===================== 第 13 页：攻关 + 开放 =====================
def slide_measure_break():
    slide_measure_page(
        "对策三  攻关 + 开放（突破端）",
        13,
        [
            ("5", "关键核心技术攻关",
             ["重点：集成电路、工业软件、高端装备、生物医药",
              "新型举国体制：政府主导 + 市场配置资源",
              "强化企业创新主体地位（华为、比亚迪、宁德时代）"]),
            ("6", "高水平对外开放",
             ["主动参与 / 牵头国际大科学计划",
              "吸引全球顶尖人才'为我所用'",
              "受限领域加快国产替代，但不搞封闭排外"]),
        ],
        "新型举国体制",
        "新型举国体制：\n\n党的领导 + 政府\n主导 + 市场配置 +\n企业主体 +\n社会参与。\n\n既发挥社会主义\n集中力量办大事\n的优势，又激发\n市场和企业活力。\n\n是攻克'卡脖子'\n技术的中国方案。"
    )


# ===================== 第 14 页：结论 =====================
def slide_conclusion():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "结论与展望", "Conclusion & Outlook")

    # 三个核心观点
    add_text(s, Inches(0.7), Inches(1.95), Inches(12), Inches(0.5),
             "核心观点", size=22, bold=True, color=DEEP_BLUE)

    points = [
        ("01",
         "科学发展是社会系统工程",
         "受政治、经济、教育、文化、国际环境、法律军事等多重社会因素共同作用，\n任何单一维度的努力都难以独自推动整体进步。"),
        ("02",
         "坚持'四个面向'战略导向",
         "面向世界科技前沿、经济主战场、国家重大需求、人民生命健康，\n是我国科技工作的基本遵循。"),
        ("03",
         "需要长期主义 + 系统思维 + 开放胸怀",
         "实现高水平科技自立自强既要久久为功，又要协同推进，\n更要在开放合作中融入全球创新网络。"),
    ]
    top = Inches(2.6)
    for num, head, body in points:
        # 大圆数字
        circle = add_rect(s, Inches(0.7), top, Inches(0.95), Inches(0.95),
                          fill=DEEP_BLUE, shape=MSO_SHAPE.OVAL)
        add_text(s, Inches(0.7), top, Inches(0.95), Inches(0.95),
                 num, size=22, bold=True, color=TECH_ORANGE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font="Arial Black")
        add_text(s, Inches(1.85), top + Inches(0.05), Inches(10.5), Inches(0.4),
                 head, size=17, bold=True, color=DEEP_BLUE)
        add_text(s, Inches(1.85), top + Inches(0.5), Inches(10.5), Inches(0.7),
                 body, size=12, color=DARK_TEXT)
        top += Inches(1.15)

    # 底部金句
    add_rect(s, Inches(0.7), Inches(6.3), Inches(11.9), Inches(0.85),
             fill=TECH_ORANGE)
    add_text(s, Inches(0.7), Inches(6.3), Inches(11.9), Inches(0.85),
             '"科技兴则民族兴，科技强则国家强。"  —— 习近平',
             size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_footer(s, 14)


# ===================== 第 15 页：参考文献 + 致谢 =====================
def slide_refs():
    s = prs.slides.add_slide(BLANK)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "参考文献", "References")

    refs = [
        "[1] 贝尔纳 J.D. 《科学的社会功能》[M]. 商务印书馆, 1982.",
        "[2] 库恩 T. 《科学革命的结构》[M]. 北京大学出版社, 2003.",
        "[3] 李约瑟. 《中国科学技术史》[M]. 科学出版社.",
        "[4] 国家统计局. 2023 年全国科技经费投入统计公报[R]. 2024.",
        "[5] 中共中央国务院. 《'十四五'国家科技创新规划》[Z]. 2021.",
        "[6] 习近平. 在全国科技创新大会上的讲话[N]. 人民日报, 2024.",
        "[7] OECD. Main Science and Technology Indicators[R]. 2023.",
        "[8] 中国科协. 中国公民科学素质基准[Z]. 2023.",
        "[9] 路甬祥. 论中国近现代科技发展之路[J]. 中国科学院院刊, 2009.",
    ]
    top = Inches(2.0)
    for ref in refs:
        add_text(s, Inches(0.9), top, Inches(11.5), Inches(0.45),
                 ref, size=13, color=DARK_TEXT)
        top += Inches(0.42)

    # 致谢横幅
    add_rect(s, 0, Inches(6.4), SW, Inches(1.1), fill=DEEP_BLUE)
    add_text(s, Inches(0), Inches(6.45), Inches(13.333), Inches(0.5),
             "感 谢 聆 听",
             size=32, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, font=CN_TITLE_FONT)
    add_text(s, Inches(0), Inches(6.95), Inches(13.333), Inches(0.4),
             "Thanks for Your Attention",
             size=14, color=TECH_ORANGE,
             align=PP_ALIGN.CENTER, font="Calibri")


# ===================== 生成 =====================
slide_cover()
slide_intro()
slide_overview()
slide_factor_politics()
slide_factor_economy()
slide_factor_education()
slide_factor_culture()
slide_factor_international()
slide_factor_others()
slide_part2_overview()
slide_measure_input()
slide_measure_env()
slide_measure_break()
slide_conclusion()
slide_refs()

OUT = "/projects/sandbox/RUOTZ/科技观作业-影响科学发展的社会因素.pptx"
prs.save(OUT)
print(f"✅ 已生成：{OUT}")
print(f"📊 共 {len(prs.slides)} 页")
