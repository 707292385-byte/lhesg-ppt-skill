# -*- coding: utf-8 -*-
"""probe_template.py — 探查基底的母版/版式结构与主题色（自包含，读 assets/base.pptx）"""
import zipfile, re, os
BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "assets", "base.pptx")
z = zipfile.ZipFile(BASE)

print("=== 主题色 clrScheme ===")
th = z.read("ppt/theme/theme1.xml").decode("utf-8")
m = re.search(r"<a:clrScheme.*?</a:clrScheme>", th, re.S)
blk = re.sub(r">\s*<", "><", m.group(0))
for slot in re.findall(r"<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>(.*?)</a:\1>", blk, re.S):
    name, body = slot
    c = re.search(r'val="([0-9A-Fa-f]{6})"', body)
    print(f"  {name:8s} {c.group(1) if c else body}")

print("\n=== layout2 标题占位符样式 ===")
x = z.read("ppt/slideLayouts/slideLayout2.xml").decode("utf-8")
i = x.find("Title Placeholder")
seg = re.sub(r">\s*<", "><", x[max(0, i - 300): i + 1600])
print(seg)

print("\n=== layout2 矩形6(灰块) ===")
j = x.find("矩形 6")
print(re.sub(r">\s*<", "><", x[j: j + 900]))

print("\n=== master1 装饰条与 logo 的 rId 映射 ===")
rel = z.read("ppt/slideMasters/_rels/slideMaster1.xml.rels").decode("utf-8")
for r in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rel):
    if "media" in r[1]:
        print("  ", r[0], "->", r[1])
