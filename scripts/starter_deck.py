# -*- coding: utf-8 -*-
"""
starter_deck.py — 设计经验的"手法演示"（不是版式模板，也不是起点模板）

⚠️ 本文件把 references/design-experience.md 里的经验写成可运行示例，用途是
   **查阅手法**（怎么用原语实现"编号块""结论条""传导链""对比栏""比例条"），
   以及跑一遍看渲染效果。**新 deck 请按内容重新设计版式**，按需摘用其中手法，
   不要整份复制换文案。

   **页面文本一律只放内容**（真实内容或明确占位），写作指导与版式说明写在注释里，
   不得写进幻灯片——见 SKILL.md 硬性规范 7。

   本文件展示的手法 ↔ 经验对应：
     封面 / 目录 / 结尾 —— 三种特殊页由 kit 内置（kit.cover / kit.toc / kit.end），
                            视觉识别元素与内容页同源，按需调用，不必自己重画
     P1  编号驱动 + 右侧数据卡 —— E2 信息分层碎片化 / E3 不对称分栏
     P2  结论置顶 + 表格       —— E1 结论先行 / E6 表格是稀缺手段
     P3  chevron 传导链         —— E4 深色只给结构（链条由深到浅、末端用橙）
     P4  结论置顶 + 对比 + VS 圆 —— E1 结论先行 / E3 交替底色划分主次
     P5  结论置顶 + 比例条 + 指标 —— E1 结论先行 / E4 橙色只给负面
   全部内容页共同遵循 E1（结论式标题）与 E5（统一网格）。
   凡有「全页结论」的页（P2/P4/P5），结论条一律紧贴标题下方、证据之上（结论先行）。

用法:
    python starter_deck.py <输出.pptx>      仅用于查看手法效果

字号规范（唯一标准）: 正文类不得小于 12pt——默认 14pt(S_BODY)，信息密集页 12pt(S_DENSE)；
                 注释类（口径说明、单位、来源行）用 9pt(S_AUX)。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx.dml.color import RGBColor  # noqa: E402
from esg_deck_kit import (  # noqa: E402
    EsgDeck, PRIMARY, NAVY, DEEP, BRIGHT, BODY, GRAY, LIGHT, BORDER,
    MUTED, ORANGE, WHITE, F_HEAVY, F_MED, F_LIGHT,
    S_H2, S_BODY, S_DENSE, S_AUX, S_NUM, S_NUM_L,
    MSO_SHAPE, PP_ALIGN, MSO_ANCHOR,
)

TRACK = RGBColor.from_string("E3EAF6")   # 比例条底槽

OUT = sys.argv[1] if len(sys.argv) > 1 else "lhesg_starter_deck.pptx"
SRC = "资料来源：此处填写可回溯的原始出处"

d = EsgDeck(source_line=SRC)

# ══════════════════════ 封面（自带样式，按需使用）══════════════════════
# 封面无副标题槽：需要补充信息就写进标题次行（\n 换行）
d.cover("报告主标题\n一句话结论",
        presenter="汇报人姓名", date="2026年X月X日")

# ══════════════════════ 目录（自带样式，按需使用）══════════════════════
# 条目可写 "篇章名"，需要时写成 ("篇章名", "右侧注释") 挂一条 9pt 右对齐注释
# 卡片浅蓝底与白底自动交替，分隔靠底色不靠线条
d.toc([("第一部分　章节标题"),
       ("第二部分　章节标题"),
       ("第三部分　章节标题"),
       ("第四部分　章节标题")])

# ══════════════════════ P1 编号驱动 + 右侧数据卡 ══════════════════════
# E1：标题写判断不写主题。E2：左侧拆成 编号块 + 小标题 + 正文 三级
# 三条论据同构、信息量相当，便于横向比较
s = d.page("三重要素共同推动，议题进入主流视野")

items = [
    ("01", "要素一", "此处填写第一层论据。"),
    ("02", "要素二", "此处填写第二层论据。"),
    ("03", "要素三", "此处填写第三层论据。"),
]
y = 1.42
for num, lab, txt in items:
    chip = d.rect(s, 0.67, y, 0.62, 0.62, fill=PRIMARY)          # 编号方块
    d.shape_text(chip, [[(num, S_H2, WHITE, F_MED, False)]])
    d.text(s, 1.48, y - 0.02, 5.35, 0.35, [[(lab, S_H2, PRIMARY, F_MED, True)]])
    d.text(s, 1.48, y + 0.36, 5.35, 1.05,
           [[(txt, S_BODY, BODY, F_LIGHT)]], line_spacing=1.22)
    y += 1.72

# 右侧数据卡：一个大数字 + 一条结构条 + 一句结论（E3 不对称分栏 7:11）
d.rect(s, 7.25, 1.42, 5.41, 5.16, fill=LIGHT, rounded=True, radius=0.045)
d.icon(s, "image31.png", 7.58, 1.72, 0.52, circle_fill=PRIMARY)
d.text(s, 8.28, 1.76, 4.2, 0.45, [[("数据卡标题", S_H2, PRIMARY, F_MED, True)]],
       anchor=MSO_ANCHOR.MIDDLE)
d.text(s, 7.58, 2.48, 4.75, 1.0, [
    [("00.00", S_NUM_L, PRIMARY, F_HEAVY, False), (" 单位", S_DENSE, PRIMARY, F_MED, True)],
    [("数字口径说明（统计期间 / 范围）", S_AUX, GRAY, F_LIGHT)],
], line_spacing=1.1)
d.text(s, 7.58, 3.66, 4.75, 0.3,
       [[("占比构成", S_DENSE, BODY, F_MED, True)]])
barw = 4.75
d.rect(s, 7.58, 4.06, barw, 0.34, fill=DEEP)                     # 底槽（示意比例）
d.rect(s, 7.58, 4.06, barw * 0.10, 0.34, fill=BRIGHT)
d.text(s, 7.58, 4.50, 2.0, 0.25, [[("次要项 00%", S_DENSE, BRIGHT, F_MED, True)]])
d.text(s, 10.0, 4.50, 2.33, 0.25, [[("主要项 00%", S_DENSE, DEEP, F_MED, True)]],
       align=PP_ALIGN.RIGHT)
d.rect(s, 7.58, 4.94, 4.75, 0.012, fill=BORDER)                  # 细分隔线
d.text(s, 7.58, 5.10, 4.75, 1.35, [
    [("此处填写对该数据的判断。", S_H2, ORANGE, F_HEAVY, True)],
], line_spacing=1.25)

# ══════════════════════ P2 表格：多主体横向对照 ══════════════════════
# E6：全 deck 一张表即可，只用于多主体横向对照
# E1：全页结论紧贴标题下方，表格作为证据排在其后（结论先行）
s = d.page("三大主体走出三条不同的技术路径")
d.conclusion(s, [("此处填写全页结论。", S_BODY, PRIMARY, F_MED, True)])
data = [
    ["主体", "做法", "关键特征"],
    ["主体 A", "做法一", "此处填写关键特征。"],
    ["主体 B", "做法二", "此处填写关键特征。"],
    ["主体 C", "做法三", "此处填写关键特征。"],
]
d.table(s, 0.67, 2.54, 11.99, 3.9, data,
        col_w=[2.1, 3.2, 6.69], row_h=1.09, header_h=0.62)

# ══════════════════════ P3 chevron 传导链 ══════════════════════
# E4：深色只给结构；链条由深到浅，末端用橙表示落点/警示
# 链条下方配"原文佐证"与"补充观察"两块，形成 事实 → 佐证 → 旁证 的层次
s = d.page("风险沿经营链条逐级传导至最终结果")
chain = ["起点", "环节一", "环节二", "环节三", "结果"]
fills = [NAVY, DEEP, PRIMARY, BRIGHT, ORANGE]      # 由深到浅再到橙（结果/警示）
x, w, overlap = 0.67, 2.62, 0.28
for i, (txt, fl) in enumerate(zip(chain, fills)):
    kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
    shp = d.shape(s, kind, x, 1.75, w, 0.95, fl)
    d.shape_text(shp, [[(txt, S_H2, WHITE, F_MED, True)]])
    x += w - overlap
notes = ["起点说明", "环节机制", "环节机制", "环节机制", "影响落点"]
x = 0.67
for i, txt in enumerate(notes):
    d.text(s, x + (0.22 if i else 0.05), 2.92, w - 0.5, 0.8,
           [[(txt, S_DENSE, GRAY, F_LIGHT)]], align=PP_ALIGN.CENTER, line_spacing=1.15)
    x += w - overlap
d.rect(s, 0.67, 4.05, 11.99, 1.35, fill=LIGHT, rounded=True, radius=0.10)
d.text(s, 1.02, 4.28, 11.3, 0.95, [[
    ("此处填写佐证该链条的原文摘录，并注明出处。", S_BODY, BODY, F_LIGHT),
]], line_spacing=1.35)
d.rect(s, 0.67, 5.68, 11.99, 0.82, fill=WHITE, line=BORDER, rounded=True, radius=0.16)
d.vline(s, 0.67, 5.68, 0.82, color=ORANGE, w=0.06)
d.text(s, 1.0, 5.68, 11.4, 0.82, [[
    ("此处填写一条旁证或趋势。", S_BODY, BODY, F_LIGHT),
]], anchor=MSO_ANCHOR.MIDDLE)

# ══════════════════════ P4 左右对比 + VS 圆 ══════════════════════
# E1：全页结论置顶（标题下、对比证据上）
# E3：左右 6:5，浅蓝底 / 白底+描边 交替划分主次；VS 圆居中压缝
s = d.page("二者差距：一方量化全覆盖，一方仍以定性为主")
d.conclusion(s, [("此处填写全页结论。", S_BODY, PRIMARY, F_MED, True)])
def compare_block(x, label, chip_fill, rows, fill, line):
    d.rect(s, x, 2.54, 5.6, 3.9, fill=fill, line=line, rounded=True, radius=0.05)
    chip = d.rect(s, x + 0.35, 2.86, 1.9, 0.55, fill=chip_fill,
                  rounded=True, radius=0.5)
    d.shape_text(chip, [[(label, S_BODY, WHITE, F_MED, True)]])
    paras = []
    for head, sub in rows:
        paras.append([(f"● {head}", S_BODY, BODY, F_MED, True)])
        paras.append([(sub, S_DENSE, GRAY, F_LIGHT)])
    d.text(s, x + 0.35, 3.66, 4.9, 2.62, paras, line_spacing=1.16, space_after=10)

compare_block(0.67, "对象一", PRIMARY,
              [("特征一", "此处填写说明"), ("特征二", "此处填写说明"), ("特征三", "此处填写说明")],
              LIGHT, None)
compare_block(7.06, "对象二", MUTED,
              [("特征一", "此处填写说明"), ("特征二", "此处填写说明"), ("特征三", "此处填写说明")],
              WHITE, BORDER)
vs = d.shape(s, MSO_SHAPE.OVAL, 6.32, 4.13, 0.72, 0.72, NAVY)
d.shape_text(vs, [[("VS", S_H2, WHITE, F_HEAVY, False)]])

# ══════════════════════ P5 比例条 + 定量指标 ══════════════════════
# E1：全页结论置顶；E4：橙色每页 ≤1 处且只指向负面事实（此处给第一行）
# E3：卡片底色交替；比例条长度仅示意形态，实际填写时按真实数值取 frac
s = d.page("调整行动以负面为主，方向与体系认知一致")
d.conclusion(s, [("此处填写全页结论。", S_BODY, PRIMARY, F_MED, True)])
stats = [
    ("主体 A", "00%", 0.84, "此处填写口径（统计期间 / 样本量）。"),
    ("主体 B", "00%", 0.18, "此处填写口径（统计期间 / 样本量）。"),
    ("主体 C", "00%", 0.20, "此处填写口径（统计期间 / 样本量）。"),
]
y = 2.54
for i, (org, num, frac, desc) in enumerate(stats):
    d.rect(s, 0.67, y, 11.99, 1.30,
           fill=WHITE if i % 2 else LIGHT, rounded=True, radius=0.10,
           line=BORDER if i % 2 else None)
    d.text(s, 1.02, y + 0.22, 2.2, 0.80, [[(org, S_BODY, PRIMARY, F_MED, True)]],
           anchor=MSO_ANCHOR.MIDDLE)
    d.text(s, 3.2, y + 0.06, 2.35, 1.10,
           [[(num, S_NUM, ORANGE if i == 0 else PRIMARY, F_HEAVY, False)]],
           anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    bx, bw = 5.85, 6.2
    d.rect(s, bx, y + 0.44, bw, 0.36, fill=TRACK)
    d.rect(s, bx, y + 0.44, bw * frac, 0.36, fill=ORANGE if i == 0 else DEEP)
    d.text(s, 1.02, y + 0.88, 10.9, 0.40, [[(desc, S_DENSE, BODY, F_LIGHT)]])
    y += 1.44

# ══════════════════════ 结尾页（复刻基底原版样式，按需使用）═════════════════════
d.end("谢谢关注！",
      lines=["地址：XX市XX区XX路XX号XX大厦XX层",
             "邮编：XXXXXX　电话：+86-XXX-XXXXXXXX",
             "网址：https://www.example.com"],
      contact=["联系人姓名", "138XXXXXXXX"])

d.save(OUT)
