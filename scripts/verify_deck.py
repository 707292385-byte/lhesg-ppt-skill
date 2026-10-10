# -*- coding: utf-8 -*-
"""
verify_deck.py — 检查 deck 是否真正继承母版识别元素（结构断言）

检查项：
  1. 每页挂的版式是否为「标题和内容」内容页版式（属于带右上角 logo 的母版）
  2. 该母版是否含：右上角 logo 图、底部 logo 图、三段式底部装饰条
  3. 版式是否含标题蓝块 + 灰块 + 标题占位符（24pt / 044AC0 / 思源黑体 Heavy）
  4. 每页是否保留了标题占位符与页码占位符
  5. 是否残留分节定义（p14:sectionLst，应已由 strip_sections 清除）

用法: python verify_deck.py <deck.pptx>   全部通过退出码 0。
"""
import os
import sys
import zipfile

from pptx import Presentation
from pptx.util import Emu

EXPECT = {
    "master_logo_top": (9.56, 0.12, 3.41, 0.97),   # image1.png 跨越标语
    "master_logo_bottom": (0.61, 7.26, 2.21, 0.21),  # image2.png 底部 logo
    "layout_blue_block": (0.0, 0.50, 0.39, 0.44),
    "layout_gray_block": (0.46, 0.50, 0.14, 0.44),
}


def inches(v):
    return round(v / 914400, 2) if v is not None else None


def find_pic(shapes, box, tol=0.03):
    for sh in shapes:
        try:
            got = (inches(sh.left), inches(sh.top), inches(sh.width), inches(sh.height))
        except Exception:
            continue
        if got and all(abs(a - b) <= tol for a, b in zip(got, box)):
            return True
    return False


def main(path):
    prs = Presentation(path)
    errors, notes = [], []

    # 1) 找到内容页版式（含蓝块/灰块的 layout），确定其母版
    layout = None
    for m in prs.slide_masters:
        for lay in m.slide_layouts:
            names = [sh.name for sh in lay.shapes]
            if any("矩形" in n for n in names) and lay.name == "标题和内容":
                layout = lay
                break
    if layout is None:
        errors.append("未找到「标题和内容」内容页版式")
        layout = prs.slides[0].slide_layout
    master = layout.slide_master

    # 2) 母版识别元素
    if not find_pic(master.shapes, EXPECT["master_logo_top"]):
        errors.append("母版缺少右上角 logo 图（期望位于 9.56,0.12,3.41×0.97）")
    if not find_pic(master.shapes, EXPECT["master_logo_bottom"]):
        errors.append("母版缺少底部 logo 图（期望位于 0.61,7.26,2.21×0.21）")
    bars = [sh for sh in master.shapes
            if inches(sh.top) and 7.0 <= inches(sh.top) <= 7.3 and inches(sh.height) == 0.3]
    if len(bars) < 3:
        errors.append(f"母版底部装饰条不足三段（当前 {len(bars)} 段）")

    # 3) 版式上的蓝块/灰块/标题样式
    if not find_pic(layout.shapes, EXPECT["layout_blue_block"]) and \
       not any(find_pic([sh], EXPECT["layout_blue_block"]) for sh in layout.shapes):
        blue = [sh for sh in layout.shapes
                if inches(sh.left) == 0.0 and inches(sh.width) and 0.3 <= inches(sh.width) <= 0.45]
        if not blue:
            errors.append("版式缺少标题左侧蓝块")
    if not any(inches(sh.left) == 0.46 for sh in layout.shapes
               if inches(sh.left) is not None):
        errors.append("版式缺少标题旁灰块")
    title_ph = None
    for ph in layout.placeholders:
        if ph.placeholder_format.type == 13 or ph.placeholder_format.idx == 0:
            pass
        if ph.placeholder_format.idx == 0 and "title" in str(ph.placeholder_format.type).lower():
            title_ph = ph
    if title_ph is None:
        title_ph = layout.placeholders.title if hasattr(layout.placeholders, "title") else None
    if title_ph is not None:
        xml = title_ph._element.xml
        for kw, why in [("044AC0", "标题主题色 044AC0"),
                        ("2400", "标题字号 24pt"),
                        ("思源黑体 Heavy", "标题字体 思源黑体 Heavy")]:
            if kw not in xml:
                errors.append(f"标题占位符缺少 {why}")

    # 4) 每页结构（特殊页不要求标题占位符：封面「标题幻灯片」、
    #    结尾页等挂在其它母版上的页面、以及名为「空白」的版式）
    content_master = layout.slide_master
    for i, s in enumerate(prs.slides, 1):
        lay_name = s.slide_layout.name
        other_master = s.slide_layout.slide_master is not content_master
        if s.slide_layout is not layout:
            notes.append(f"P{i} 版式为 {lay_name}（非内容页版式，识别元素可能不完整）")
        if lay_name in {"标题幻灯片", "空白"} or other_master:
            continue
        has_title = has_num = False
        for ph in s.placeholders:
            t = str(ph.placeholder_format.type)
            if "TITLE" in t:
                has_title = True
            if "SLIDE_NUMBER" in t:
                has_num = True
        if not has_title:
            errors.append(f"P{i} 缺少标题占位符")
        if not has_num:
            notes.append(f"P{i} 无页码占位符（如需页码请保留 sldNum）")

    # 5) 分节残留（否则分节视图出现残留节名，可能触发修复提示）
    with zipfile.ZipFile(path) as z:
        pres = z.read("ppt/presentation.xml").decode("utf-8")
    if "sectionLst" in pres:
        n = pres.count("<p14:section ")
        errors.append(f"残留分节定义 {n} 个（用 scripts/strip_sections.py 清除）")

    print(f"deck    : {path}")
    print(f"slides  : {len(prs.slides)} | layout: {layout.name} | master ok")
    for n in notes:
        print("NOTE   ", n)
    if errors:
        for e in errors:
            print("FAIL   ", e)
        sys.exit(1)
    print("PASS    母版识别元素继承结构完整")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    main(sys.argv[1])
