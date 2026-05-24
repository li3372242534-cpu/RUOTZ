# -*- coding: utf-8 -*-
"""
生成自然辩证法课程论文 .docx 文档
- 正文：小四号（12pt）、宋体
- 题目下提供关键词
- 不含封面与摘要
- 设置 A4 双面打印（镜像页边距）
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION


def set_run_font(run, size_pt=12, bold=False, font_cn="宋体", font_en="Times New Roman"):
    run.font.name = font_en
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_en)
    rFonts.set(qn("w:hAnsi"), font_en)
    rFonts.set(qn("w:eastAsia"), font_cn)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    # 题目用三号黑体
    set_run_font(run, size_pt=16, bold=True, font_cn="黑体", font_en="SimHei")


def add_subtitle(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, size_pt=14, bold=True, font_cn="黑体", font_en="SimHei")


def add_keywords(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.first_line_indent = Cm(0.74)
    run_label = p.add_run("关键词：")
    set_run_font(run_label, size_pt=12, bold=True)
    run_kw = p.add_run(text)
    set_run_font(run_kw, size_pt=12, bold=False)


def add_heading(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run_font(run, size_pt=14, bold=True, font_cn="黑体", font_en="SimHei")


def add_subheading(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0.74)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    set_run_font(run, size_pt=12, bold=True, font_cn="黑体", font_en="SimHei")


def add_body(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(0.74)  # 首行缩进2字符
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, size_pt=12)  # 小四号 = 12pt


def add_ref(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.first_line_indent = Cm(-0.74)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, size_pt=12)


# ========= 创建文档 =========
doc = Document()

# 页面设置：A4，双面打印（镜像页边距）
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width = Cm(21.0)
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(3.18)
section.right_margin = Cm(3.18)

# 启用双面打印 / 镜像页边距
sectPr = section._sectPr
from docx.oxml import OxmlElement
mirror = OxmlElement("w:mirrorMargins")
sectPr.insert(0, mirror)

# 默认正文样式（小四宋体）
style_normal = doc.styles["Normal"]
style_normal.font.name = "Times New Roman"
style_normal.font.size = Pt(12)
style_normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

# ========= 论文内容 =========
add_title(doc, "薄名利以养德，敏于事以求真")
add_subtitle(doc, "——从老孔养生思想谈科研人员的道德自觉")

add_keywords(doc, "自然辩证法；学术不端；道德自觉；老子养生观；孔子仁学")

# 一
add_heading(doc, "一、问题的提出")
add_body(doc,
    "近年来，论文造假、数据篡改、抄袭剽窃等学术不端事件在全球范围内频繁发生。"
    "从韩国黄禹锡干细胞造假，到国内\u201c汉芯一号\u201d骗局，再到高校论文撤稿"
    "事件频发，科技界的诚信危机已不容忽视。Retraction Watch 数据库统计显示，"
    "全球被撤回的学术论文数量逐年攀升，其中相当比例源于人为造假[6]。这一现象"
    "的深层根源，在于科研人员在个人名利前途与道德养生之间发生了严重失衡。")
add_body(doc,
    "自然辩证法作为科学技术的哲学反思，不仅关注科技发展的客观规律，更关注科研"
    "主体的精神世界与价值取向[4]。重读老子\u201c善摄生者先除六害\u201d与孔子"
    "\u201c仁者寿\u201d的养生思想，结合当代自然辩证法的自然观、科技观与科技方"
    "法论，对于反思科研人员的道德自觉、提升其生活品质，具有深刻的现实意义。")

# 二
add_heading(doc, "二、老孔养生思想的核心要义")
add_body(doc,
    "老子在《摄生论》中言：\u201c且夫善摄生者，要当先除六害，然后可以保性命，"
    "延驻百年。\u201d所谓六害，即\u201c一者薄名利，二者禁声色，三者廉货财，四"
    "者损滋味，五者除佞妄，六者去妒忌\u201d[1]。其中\u201c薄名利\u201d列居六"
    "害之首，足见老子认为名利之累乃养生之最大障碍。若名利不除，纵然\u201c心希"
    "妙理，口念真经\u201d，亦\u201c不能补其短促\u201d。这一思想揭示了一个朴"
    "素真理：养生之本在于养心，而养心之要在于淡泊。")
add_body(doc,
    "孔子的养生观则与\u201c仁\u201d紧密相连。《论语·雍也》曰：\u201c知者乐"
    "水，仁者乐山；知者动，仁者静；知者乐，仁者寿。\u201d[2] 仁者之所以寿，"
    "在于其内心宽厚平和，能够\u201c恭、宽、信、敏、惠\u201d五德兼备。《论语·"
    "述而》记载孔子自况：\u201c饭疏食饮水，曲肱而枕之，乐亦在其中矣……发愤忘"
    "食，乐以忘忧，不知老之将至云尔。\u201d这种安贫乐道、专注于学的精神状态，"
    "恰是当代部分科研人员所缺失的境界。")
add_body(doc,
    "老孔养生思想虽分属道、儒两家，却殊途同归地指向同一核心命题：真正的健康"
    "长寿不在于物质的丰盈，而在于德性的涵养。过度追求外在的奉养，结果往往是"
    "\u201c损于其本而妄求其末\u201d；唯有养心、养德，方能真正延年益寿。")

# 三
add_heading(doc, "三、自然辩证法视域下的科技伦理审视")

add_subheading(doc, "（一）自然观的启示：人与自然的和谐共生")
add_body(doc,
    "马克思主义自然辩证法认为，自然界并非被动的客体，而是与人类相互作用、相互"
    "生成的有机整体。恩格斯在《自然辩证法》中警示道：\u201c我们不要过分陶醉于"
    "我们对自然界的胜利。对于每一次这样的胜利，自然界都对我们进行报复。\u201d"
    "[3] 这一警示与老子\u201c道法自然\u201d的思想遥相呼应。科研人员作为认识"
    "自然、改造自然的主体，其价值取向直接影响科技发展的方向。若科研活动被名利"
    "驱使，就会偏离\u201c求真\u201d的本性，导致数据造假、论文灌水，乃至研发"
    "出违背自然规律的伪成果。这种对客观规律的违背，本质上是对自然的背叛，最终"
    "也会反噬科研人员自身。")

    
add_subheading(doc, "（二）科技观的反思：双刃剑下的价值抉择")
add_body(doc,
    "科技是一把双刃剑，既可造福人类，亦可贻害无穷。自然辩证法的科技观强调，"
    "科技的善恶不取决于技术本身，而取决于使用技术的人[7]。爱因斯坦曾言："
    "\u201c关心人本身和人的命运，必须始终成为一切技术努力的主要目标。\u201d[5] "
    "当今学术不端问题的根源，正在于部分科研人员将科技异化为获取名利的工具，"
    "丧失了\u201c为天地立心，为生民立命\u201d的初心。老子所说的\u201c损滋"
    "味\u201d\u201c廉货财\u201d，并非要求科研人员清苦度日，而是警示其勿为物"
    "欲所役。当科研被简化为论文数量、影响因子、项目经费的单一角逐时，科技的"
    "人文价值便被消解殆尽。")

add_subheading(doc, "（三）科技方法论的要求：求真务实的治学精神")
add_body(doc,
    "自然辩证法的科技方法论强调归纳与演绎、分析与综合、历史与逻辑的辩证统一，"
    "其根本在于\u201c实事求是\u201d。然而，学术不端行为恰恰违背了这一基本原"
    "则。无论是数据捏造、图片篡改，还是同行评议舞弊，本质上都是用主观臆造取"
    "代客观规律，用短期利益置换长期真理。孔子曰：\u201c食无求饱，居无求安，"
    "敏于事而慎于言。\u201d[2] 这正是科研方法论的古典表达——在治学态度上敏锐"
    "严谨，在生活欲望上知足节制。\u201c敏则有功\u201d，专注与勤勉方能催生真"
    "知；\u201c慎于言\u201d则提醒科研人员对未经验证的结论保持审慎，不可为博"
    "眼球而妄下断言。")

# 四
add_heading(doc, "四、科研人员道德自觉的实现路径")

add_subheading(doc, "（一）薄名利：树立淡泊明志的价值观")
add_body(doc,
    "老子\u201c薄名利\u201d的智慧，对当代科研人员具有直接的警醒意义。所谓"
    "\u201c薄\u201d，并非否定名利的存在，而是不为名利所累。诸葛亮\u201c非淡泊"
    "无以明志，非宁静无以致远\u201d，与老子之意一脉相承。科研人员当将名利视作"
    "研究的副产品，而非追求的目标。当心无旁骛地专注于科学问题本身时，真正的成"
    "就反而水到渠成。屠呦呦数十年潜心于青蒿素研究，钟扬十六载守望雪域高原采集"
    "种子，他们的事迹证明：唯有薄名利，才能成大器。")

add_subheading(doc, "（二）去妒忌：构建合作共赢的科研生态")
add_body(doc,
    "老子\u201c六害\u201d中的\u201c除佞妄\u201d\u201c去妒忌\u201d，揭示了"
    "科研团队内部的伦理要求。学术不端的另一重土壤，是同行之间的恶性竞争与相互"
    "倾轧。剽窃他人成果、阻挠同行发表、滥用同行评议权力等行为，皆源于妒忌之心"
    "[6]。孔子\u201c宽则得众\u201d\u201c信则人任焉\u201d的教诲，启示我们以"
    "宽厚之心对待同行，以信任之态合作创新，方能形成良性学术共同体。")

add_subheading(doc, "（三）仁者寿：以德性涵养支撑科研生涯")
add_body(doc,
    "科研是一项长跑而非短跑。许多科研人员英年早逝、心力交瘁，固然有工作强度"
    "的因素，但更深层的原因在于心态失衡。孔子\u201c仁者乐山，仁者静，仁者"
    "寿\u201d的洞见，揭示了道德境界与身心健康的内在统一[2]。心怀仁爱、安于求"
    "道的科研人员，更能在漫长学术道路上保持身心平衡，实现可持续创造。现代心身"
    "医学也表明：长期处于焦虑、嫉妒、急功近利状态下的个体，免疫功能显著下降，"
    "慢性病发病率明显升高。可见，养德即养生，立德即立命。")

add_subheading(doc, "（四）制度与自觉：他律与自律的双重保障")
add_body(doc,
    "当然，仅靠道德自觉不足以根除学术不端，还需建立健全的科研诚信制度。一是"
    "完善学术评价体系，破除\u201c唯论文\u201d\u201c唯帽子\u201d的导向，回归"
    "科研本真价值；二是强化学术伦理教育，将自然辩证法的科技伦理纳入研究生必"
    "修课程；三是建立学术不端的严厉惩戒机制，使违规者付出代价。制度是底线，"
    "自觉是高线，二者结合方能形成清朗的科研环境[7]。")

# 五
add_heading(doc, "五、结语")
add_body(doc,
    "老子曰：\u201c五色令人目盲，五音令人耳聋，五味令人口爽。\u201d过度的名利"
    "追逐，最终损害的是科研人员自身的道德生命与学术生命。孔子\u201c发愤忘食，"
    "乐以忘忧，不知老之将至\u201d的境界，则展示了一种以学问为乐、以道德为本"
    "的理想人格。在自然辩证法的视域下，科技的进步与个人的修养从来不是对立的"
    "两极，而是辩证统一的整体。科研人员唯有\u201c薄名利\u201d以养其德，"
    "\u201c敏于事\u201d以求其真，\u201c乐山水\u201d以养其心，方能在个人前途"
    "与道德养生之间实现真正的平衡。这不仅是个人生活品质的提升，更是科技事业"
    "行稳致远的根本保障。当每一位科研人员都能以淡泊之心做学问、以仁厚之心待"
    "同行、以敬畏之心对自然，中国科技强国之梦才能真正照进现实。")

# 参考文献
add_heading(doc, "参考文献")
add_ref(doc, "[1] 老子. 道德经[M]. 陈鼓应, 注译. 北京: 中华书局, 2009.")
add_ref(doc, "[2] 杨伯峻. 论语译注[M]. 北京: 中华书局, 2017.")
add_ref(doc, "[3] 恩格斯. 自然辩证法[M]. 中共中央马克思恩格斯列宁斯大林著作编译局, 译. 北京: 人民出版社, 2018.")
add_ref(doc, "[4] 郭贵春. 自然辩证法概论[M]. 北京: 高等教育出版社, 2019.")
add_ref(doc, "[5] 爱因斯坦. 爱因斯坦文集: 第三卷[M]. 许良英, 等编译. 北京: 商务印书馆, 2010.")
add_ref(doc, "[6] 樊春良. 学术不端行为的特点、原因与治理[J]. 科学与社会, 2020, 10(3): 22-35.")
add_ref(doc, "[7] 王前. 中国科技伦理史纲[M]. 北京: 人民出版社, 2006.")

out = "/projects/sandbox/RUOTZ/科研人员道德自觉论文.docx"
doc.save(out)
print("Saved:", out)

# 统计字数
import re
all_text = []
for p in doc.paragraphs:
    all_text.append(p.text)
joined = "".join(all_text)
cn_chars = re.findall(r"[\u4e00-\u9fff]", joined)
print("中文字符数:", len(cn_chars))
print("总字符数:", len(joined))
