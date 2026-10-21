# lhesg-ppt · lhesg 风格 PPT 生成 Skill

制作 lhesg 风格演示文稿。**skill 自包含**：内置基底
`assets/base.pptx`（母版 + 版式 + 主题 + 封面/结尾页原型）与品牌/图标资产，构建时
不依赖任何外部模板文件。视觉识别元素（标题蓝块灰块、双 logo、三段装饰条、主题色
`#044AC0`）由母版原生继承，正文版式按内容自由设计。

固化的对象是**判断与规范**，不是某一版成品：设计经验见 `references/design-experience.md`（E1–E7），表达手法见 `references/layout-library.md`。内容页版式每次按当次内容重新设计，不套用任何既有 deck 的结构。

## 依赖

- Python 3.9+，`python-pptx`、`Pillow`
- （无外部文件依赖；如需替换基底，设环境变量 `ESG_BASE` 指向自定义 base.pptx）

```bash
pip install -r requirements.txt          # 等价于 pip install python-pptx pillow
```

## 目录

```
SKILL.md                      技能入口：流程、硬约束、工具用法
requirements.txt              依赖清单（python-pptx / Pillow）
references/
  design-experience.md        E1–E7 设计经验（判断依据，不含某版成品参数）
  layout-library.md           十二种表达手法与选择、去重规则
  visual-spec.md              品牌识别、色板、字号、安全区
scripts/
  esg_deck_kit.py             构建器：EsgDeck（page / cover / toc / end / table / shape …）
  pptx_render_preview.py      版式渲染预览（替代 qlmanage，支持 p:bg 与垂直锚点）
  font_lint.py                字号合规检查
  verify_deck.py              结构断言（母版继承元素、分节残留）
  strip_sections.py           清除既有 deck 的分节残留
  probe_base.py               探查基底的母版/版式结构
assets/
  base.pptx                   内置基底（母版/版式/主题 + 封面与结尾页原型）
  brand/                      品牌素材：地球背景、箭头条、双 logo、公众号二维码
  icons/                      图标 100 个（模板原生 7 个 + 同系列 Material 实心 93 个，含 INDEX.txt）
```

## 快速开始

```python
import sys
sys.path.insert(0, "<repo>/scripts")
from esg_deck_kit import EsgDeck

d = EsgDeck(source_line="资料来源：（逐条可回溯到原件，禁止臆造）")
d.cover("此处写报告主标题：\n一句话点明主张", presenter="姓名", date="2026年X月X日")
d.toc([("第一章", None), ("第二章", "右侧 9pt 注释")])
s = d.page("结论式标题：一句话说清本页观点")
d.text(s, 0.67, 1.42, 11.99, 0.4, [[("正文", 14, None, None)]])
d.end("谢谢关注！", lines=["地址：…"], contact=["联系人：…", "联系电话：…"])
d.save("输出.pptx")
```

## 硬约束

1. **一页一论点**：标题必须是结论式陈述句，不复述事实。
2. **字号（唯一标准）**：正文类 ≥ 12pt（默认 14pt，信息密集页 12pt）；注释类（口径说明、单位、来源行）9pt。内容页合法档位 9/12/14/15/24/30/34，10–11pt 禁止使用；封面 16/28pt、结尾 44pt 属固定原型页字号。交付前必跑 `font_lint.py`。
3. **版式是开放集**：只约束识别元素、字号、安全区、色板；版式按当次内容重新设计，判断标准是「观点讲清楚」。
4. **版式多样性**：出方案时每页标注表达手法；相邻页不同构；同一手法整份最多 2 次；不沿用旧 deck 结构。
5. **特殊页固定**：封面 / 结尾页沿用基底原型，仅替换文字；封面**不设副标题**。
6. **不留分节残留**：`presentation.xml` 的 `p14:sectionLst` 独立于页码，删页不会随之删除（基底已预清，kit 建 deck 时再清一次）。
7. **渲染验证必做**：`pptx_render_preview.py` 逐页渲染后目检（文字溢出 / 压线 / 留白失衡 / 相邻页同构）。
8. **全页结论置顶**：有全页结论的页，结论条紧贴标题下方、排在图表证据之上（顺序：标题 → 全页结论 → 证据 → 来源行），用 `d.conclusion(...)` 实现。

## 版权说明

基底与品牌素材（`assets/`）版权归原权利人所有，仅限授权主体内部使用。
