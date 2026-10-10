# -*- coding: utf-8 -*-
"""
font_lint.py — deck 字号合规检查（lhesg 风格唯一标准）

规范:
    正文类（观点/论据/表格内容/卡片正文/模块标题/关键数字/数据标签）: ≥ 12pt
        默认 14pt(S_BODY)，信息密集页 12pt(S_DENSE)
    注释类（口径说明、图注、单位、页脚来源行）: 9pt(S_AUX)
    合法档位: 9 / 12 / 14 / 15 / 24 / 30 / 34

判定:
    ❌ 违规  字号 < 9pt，或落在 10–11pt（既非正文也非注释的无主档位）
    ❌ 违规  字号 9pt 但 run 文本 > 30 字，且不在页脚来源行（疑似正文用了注释字号）
    ⚠️ 提示  字号不在规范档位（如 13 / 16pt）——建议归入 9/12/14/15，杜绝随手数

用法:
    python font_lint.py <deck.pptx>              # 逐页列字号 + 违规项
    python font_lint.py <deck.pptx> --quiet      # 只看问题
退出码: 0 合规；1 有违规。
"""
import sys
import re
from collections import defaultdict

from pptx import Presentation

SCALE = (9, 12, 14, 15, 16, 24, 28, 30, 34, 44)   # 规范档位（16/28/44 为封面/结尾页基底原生固定档）
BODY_FLOOR = 12                        # 正文类下限
AUX = 9                                # 注释类字号
AUX_MAX_CHARS = 30                     # 注释 run 的长度上限（超出疑似正文）
SRC_Y = 6.75                           # 页脚来源行判定（英寸）
EMU = 914400.0

AUX_HEAD = re.compile(r"^(资料来源|数据来源|来源|注[:：]|口径|单位)")


def _iter_shapes(shapes):
    for sh in shapes:
        if str(sh.shape_type).startswith("GROUP"):
            yield from _iter_shapes(sh.shapes)
        else:
            yield sh


def collect(deck):
    """返回 [(slide_idx, size, text, top_inch, box_name), ...]"""
    prs = Presentation(deck)
    rows = []
    for si, slide in enumerate(prs.slides, 1):
        for sh in _iter_shapes(slide.shapes):
            top = sh.top / EMU if sh.top is not None else None
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            sz = r.font.size.pt if r.font.size else None
                            rows.append((si, sz, r.text.strip(), top, sh.name))
            if getattr(sh, "has_table", False) and sh.has_table:
                for ri, row in enumerate(sh.table.rows):
                    for ci, cell in enumerate(row.cells):
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if r.text.strip():
                                    sz = r.font.size.pt if r.font.size else None
                                    rows.append((si, sz, r.text.strip(), top,
                                                 f"{sh.name}[r{ri}c{ci}]"))
    return rows


def classify(size, text, top):
    """返回 (是否违规, 是否提示, 理由)"""
    if size is None:
        return False, True, "无显式字号（继承默认，风险高）"
    if size < AUX:
        return True, False, f"低于注释下限 {AUX}pt"
    if AUX < size < BODY_FLOOR:
        return True, False, f"落在 {AUX}–{BODY_FLOOR}pt 无主档位（要么正文 ≥12，要么注释 =9）"
    if size == AUX:
        is_src = top is not None and top >= SRC_Y
        if not is_src and len(text) > AUX_MAX_CHARS and not AUX_HEAD.match(text):
            return True, False, f"{AUX}pt 但文本 {len(text)} 字（疑似正文用了注释字号）"
        return False, False, "注释/来源"
    if size not in SCALE:
        return False, True, f"非规范档位（归入 {SCALE}）"
    return False, False, "正文类"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    quiet = "--quiet" in sys.argv
    if not args:
        print(__doc__)
        return 2
    deck = args[0]
    rows = collect(deck)

    by_slide = defaultdict(list)
    for si, sz, txt, top, name in rows:
        by_slide[si].append((sz, txt, top, name))

    bad, warn = [], []
    for si, sz, txt, top, name in rows:
        v, w, why = classify(sz, txt, top)
        if v:
            bad.append((si, sz, txt, top, why))
        elif w:
            warn.append((si, sz, txt, top, why))

    if not quiet:
        print(f"=== 逐页字号（{deck}）===")
        for si in sorted(by_slide):
            sizes = sorted({s for s, *_ in by_slide[si] if s is not None})
            print(f"  P{si:<3} 档位 {sizes}")
            for sz, txt, top, name in by_slide[si]:
                y = f"y={top:.2f}" if top is not None else "y=?"
                mark = "  ← 违规" if any(b[0] == si and b[1] == sz and b[2] == txt for b in bad) else ""
                print(f"        {str(sz):>5}pt  {y:<9} {txt!r}{mark}")

    print(f"\n=== 结果（正文下限 {BODY_FLOOR}pt / 注释 {AUX}pt）===")
    print(f"文本 run 总数: {len(rows)}")
    if bad:
        print(f"❌ 违规 {len(bad)} 处:")
        for si, sz, txt, top, why in bad:
            print(f"   P{si} {sz}pt {txt!r} —— {why}")
    else:
        print("✅ 无违规：所有字号均符合正文 ≥12pt / 注释 =9pt")
    if warn:
        print(f"⚠️  提示 {len(warn)} 处:")
        for si, sz, txt, top, why in warn:
            print(f"   P{si} {sz}pt {txt!r} —— {why}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
