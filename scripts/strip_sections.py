# -*- coding: utf-8 -*-
"""
strip_sections.py — 清除 pptx 中残留的分节定义（p14:sectionLst）

背景：presentation.xml 的 p14:sectionLst 独立于页码存在，删页不会删它；分节定义里的
slideId 指向已删页时，打开文件分节视图会出现残留节名，并可能触发修复提示。
（内置基底已预清分节；本脚本用于外部拿来的既有 deck。）

用法:
    python strip_sections.py <deck.pptx> [更多.pptx ...]        # 原地清除（先备份到 /tmp）
    python strip_sections.py <deck.pptx> -o <输出.pptx>          # 输出到新文件，不改原件
    python strip_sections.py --check <deck.pptx>                 # 只检查，不改动

说明：只改 ppt/presentation.xml 一个部件，其余条目按原字节复制，不影响母版/版式/图片。
"""
import os
import shutil
import sys
import zipfile
from datetime import datetime

from lxml import etree
from pptx.oxml.ns import qn

# 分节扩展的固定 URI（PowerPoint 2010+ sectionLst）
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SEC_PATH = ".//{%s}section" % P14_NS
PRESENTATION_PART = "ppt/presentation.xml"


def strip_from_xml(xml_bytes):
    """移除 presentation.xml 中的分节定义，返回 (新字节, 移除的节数)。"""
    root = etree.fromstring(xml_bytes)
    n_sec = 0
    for ext_lst in root.findall(qn("p:extLst")):
        for ext in list(ext_lst.findall(qn("p:ext"))):
            secs = ext.findall(SEC_PATH)
            if ext.get("uri") == SECTION_EXT_URI or secs:
                n_sec += len(secs)
                ext_lst.remove(ext)
        if len(ext_lst) == 0:            # extLst 空了就整体删掉，避免留空壳
            ext_lst.getparent().remove(ext_lst)
    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out, n_sec


def count_sections(path):
    """返回 (节数, 节名列表)。"""
    with zipfile.ZipFile(path) as z:
        x = z.read(PRESENTATION_PART)
    root = etree.fromstring(x)
    secs = root.findall(SEC_PATH)
    return len(secs), [s.get("name") or "" for s in secs]


def strip_file(path, out=None, backup=True):
    """清除单个 deck 的分节。原地修改时默认备份到 /tmp/esg_ppt_bak/。"""
    n_sec, names = count_sections(path)
    if n_sec == 0:
        print(f"SKIP   {os.path.basename(path)}：无分节定义")
        return 0, names

    if out is None and backup:
        bdir = "/tmp/esg_ppt_bak"
        os.makedirs(bdir, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        bak = os.path.join(bdir, f"{os.path.splitext(os.path.basename(path))[0]}.{stamp}.bak.pptx")
        shutil.copy2(path, bak)
        print(f"BACKUP {bak}")

    src, dst = path, out or path
    tmp = dst + ".tmp-strip"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == PRESENTATION_PART:
                data, removed = strip_from_xml(data)
                assert removed == n_sec, f"节数不一致 {removed} != {n_sec}"
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    os.replace(tmp, dst)
    print(f"OK     {os.path.basename(dst)}：移除 {n_sec} 个分节 {names}")
    return n_sec, names


def main(argv):
    args = [a for a in argv if not a.startswith("-")]
    check_only = "--check" in argv
    out = None
    if "-o" in argv:
        i = argv.index("-o")
        out = argv[i + 1]
        args = [a for a in args if a != out]

    if not args:
        print(__doc__)
        return 2

    if check_only:
        bad = 0
        for p in args:
            n, names = count_sections(p)
            print(f"{'FAIL' if n else 'PASS'}   {os.path.basename(p)}：{n} 个分节 {names}")
            bad += 1 if n else 0
        return 1 if bad else 0

    if out and len(args) > 1:
        print("ERROR: -o 只能配合单个输入文件使用")
        return 2

    for p in args:
        strip_file(p, out=out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
