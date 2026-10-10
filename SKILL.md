---
name: lhesg-ppt
description: 构建 lhesg 风格 PPT。skill 自包含：内置基底（母版识别元素 + 封面/结尾页原型）与图标/品牌资产（100 个同风格图标），不依赖外部模板或既有 deck；固化设计经验（一页一论点、信息分层碎片化、不对称分栏、克制用色）与硬性规范（字号/色板/安全区），版式每次按内容重新设计。当用户要求"做 PPT / 做 deck / 做素材页"且要求"lhesg 风格"时使用。
---

# lhesg 风格 PPT

## 固化的是什么（务必先读）

本 skill **自包含**：资产、经验、规范都在仓内，不依赖任何外部模板或既有 deck。

| 层 | 内容 | 载体 |
|---|---|---|
| 资产：基底与素材 | 母版/版式/主题（标题蓝块灰块、双 logo、三段装饰条、主题色板）+ 封面/结尾页原型 + 图标 | `assets/base.pptx`、`assets/icons/`、`assets/brand/`，母版原生继承，自动生效 |
| 软法：设计经验 | 一页一论点、信息分层碎片化、不对称分栏+交替底色、克制用色、统一网格、稀缺手段 | `references/design-experience.md`，**每次按内容重新运用** |
| 规范：硬性要求 | 字号两类两档、色板、安全区、数据可溯源 | 下文「硬性规范」，`font_lint.py` / `verify_deck.py` 强制 |

> 三条已明确的边界（勿再摇摆）：
> 1. **只要求视觉识别统一，不要求套用固定版式**；版式是开放集，没有白名单。
> 2. `starter_deck.py` 是**手法演示**（把经验写成可运行示例），可拆解借用，不是必须使用的版式集。
> 3. **字号只有一条标准**：正文 14pt、密集页 12pt、下限 12pt，不参考任何其他文件的字号。

## 设计经验（核心）

完整版见 `references/design-experience.md`（含逐页自检、禁忌）。要点：

- **E1 标题即结论**：标题写判断不写主题，双段式（要害判断主蓝加粗 + 机理说明常规字重），≤25 字，禁"XX 概述 / XX 介绍"。**有全页结论时结论先行**——结论条紧贴标题下方、排在图表等证据之上（顺序：标题 → 全页结论 → 证据 → 来源行），`d.conclusion()` 已内置该条样式与位置。
- **E2 信息分层碎片化**：一页拆 13–23 个文本框（至少 8 个），四级层级 模块标题 15pt → 关键数字 30–34pt → 正文 14pt → 注脚/口径 9pt；每个语义单元都拆成"标题行 + 内容行"。
- **E3 不对称分栏 + 交替底色**：左右约 6:5 或 7:4，不用等分三列；并列卡片用浅蓝底 `EEF3FC` 与白底+描边 `D6E2F5` 交替，而非加分隔线。
- **E4 克制用色**：深色只给结论条/表头/VS 圆等结构块（≤2 处/页）；橙色 `C55A11` 每页 ≤1 处且必须指向负面事实。
- **E5 统一网格 + 相邻页不同构**：锚点 左 0.67 / 宽 11.99 / 内容起 1.42 / 来源行 6.88；版式形态逐页变化。
- **E6 稀缺手段**：表格只用于多主体横向对照（全 deck 1 张即可）；图标 0.5–3 个/页，白图标垫蓝圆底、蓝图标直放。

反面警示（已被否掉）：三列并列排满、连续两页同构、一页只有 3–5 个大文本框、内容精简到只剩短语。

## 硬性规范

1. **字号两类两档**（全篇只有这一条标准，不参考任何既有文件的字号）：
   - **正文类 ≥12pt**：正文默认 14pt，信息密集页 12pt；模块标题 15pt；关键数字 30–34pt；表格内容、卡片正文、编号/标签同属此类。**严禁低于 12pt**。
   - **注释类 =9pt**：口径说明、图注、单位、页脚来源行。
   - 10–11pt 为无主档位，**禁止使用**；合法档位只有 9 / 12 / 14 / 15 / 24 / 30 / 34。页标题 24pt 由母版固定。
   - 落地工具：`font_lint.py` 逐页审计；10–11pt 与 <9pt 直接判违规。
2. **主题色**：主蓝 044AC0 / 深蓝 123685 / 次深 0F4A9E / 亮蓝 376FFF / 浅蓝底 EEF3FC / 描边 D6E2F5 / 弱化蓝 8FAADC；橙色 C55A11 仅负面。
3. **数据真实性**：所有数字可回溯到原始材料；页脚统一来源行（`EsgDeck(source_line=...)` 改一处）；没有出处就不写。
4. **版式安全区**：正文 y ∈ [1.42, 6.80]，x ∈ [0.67, 12.66]；右上角 x>9.5 且 y<1.2 被 logo 占用；y>7.2 是装饰条。
5. **标题写结论不写主题**（见 E1）。
6. **不留分节残留**：`presentation.xml` 的 `p14:sectionLst` 独立于页码，删页不会删它。基底已预清，kit 建 deck 时再清一次；既有 deck 用 `scripts/strip_sections.py` 补清，否则分节视图出现残留节名。
7. **页面文本只放内容**：幻灯片上只出现「真实内容」或「明确占位」（如 *此处填写…*、`00.00`、`XX%`）。写作指导（"一句话讲清机制""不要堆形容词"）、版式说明（"卡片底色自动交替""条目右侧可挂 9pt 注释"）、对使用者的交代一律**不得写进页面**；结论条**直接写结论**，不加"本页结论：""原文示例："这类栏目标签。交付前把成品文本导出扫一遍，出现祈使句/第二人称/栏目标签即为违规。

## 工作流

1. **自检**：`python scripts/esg_deck_kit.py --check` 确认内置基底与图标资产在位（skill 自包含，无需外部文件）。
2. **内容先行**：先理清论点清单（每页一句话结论 + 支撑证据），再谈版式。证据不足的页先补材料，不靠排版遮掩。
3. **出方案待确认**：给页面清单（每页标题 + 该页打算用的表达手法 + 为什么这样表达），用户确认后再动手。
4. **设计并构建**：**按 E1–E6 为每页单独设计版式**，用 kit 原语实现。需要借鉴手法时查阅 `starter_deck.py` 并改写成适合本页的表达，**不要整份复制换文案**。需要封面/目录/结尾页时直接调 `d.cover / d.toc / d.end`（样式已内置）。几何单位英寸。
5. **字号合规检查**（必做）：`python scripts/font_lint.py <deck>`，必须无违规项（正文 ≥12pt / 注释 =9pt）。
6. **渲染验证**（必做）：`python scripts/pptx_render_preview.py <deck> <png目录>` 逐页渲染 → Pillow 拼图 → Read 目检：文字溢出/截断、压线、留白失衡、相邻页是否同构。
7. **结构校验**（必做）：`python scripts/verify_deck.py <deck>` 确认识别元素继承完整、占位符保留正确、**无分节残留**（既有 deck 先跑 `strip_sections.py`）。
8. **经验自检**：按 `design-experience.md` 的六问逐页过一遍。
9. **交付**：文件名 `主题-用途-lhesg风格-YYYYMMDD.pptx`。

## 工具与用法

```python
import sys; sys.path.insert(0, "<skill目录>/scripts")
from esg_deck_kit import EsgDeck, PRIMARY, LIGHT, S_H2, S_BODY, S_DENSE, F_MED, F_LIGHT

d = EsgDeck(source_line="资料来源：……")     # 加载内置基底：清空原型页、清分节、备好 layout2
d.cover("主标题", presenter="姓名", date="2026年X月X日")  # 封面，按需；无副标题
d.toc(["第一章", ("第二章", "右侧9pt注释"), ...])                            # 目录，按需
s = d.page("结论式标题")                     # 新页：删正文占位符，保留标题+页码，写标题与来源行
d.conclusion(s, "全页结论：紧贴标题下方，证据排在其下")   # 结论先行（按需，返回底边 y）
d.rect(s, x, y, w, h, fill=LIGHT, rounded=True, radius=0.10)   # 卡片底
d.vline(s, x, y, h, color=PRIMARY, w=0.06)                     # 卡片左竖强调线
d.text(s, x, y, w, h, [[("模块标题", S_H2, PRIMARY, F_MED, True)],
                       [("正文", S_BODY, BODY, F_LIGHT)]])      # 段落 = run 列表
d.shape_text(chip, [[("01", S_H2, WHITE, F_MED)]])             # 形状内文字（编号/标签用 15pt 档）
d.icon(s, "image31.png", x, y, 0.52, circle_fill=PRIMARY)      # 白图标垫蓝圆底
d.table(s, x, y, w, h, data, col_w=[...], row_h=1.09)          # 表头深蓝，正文自动左对齐
d.end("谢谢关注！", lines=["地址：…", "网址：…"], contact=["联系人：…", "联系电话：…"])  # 结尾页，按需
d.save("输出.pptx")
```

三种特殊页（`cover` / `toc` / `end`）由 kit 内置、复刻基底原型样式与资产（地球背景、箭头条、
公司 logo、公众号二维码），调用即可，不必自己重画；不需要时整页跳过（如素材页 deck）。
封面与结尾页用 `_clone_template_slide()` **整页克隆基底原型页**（slide1 / slide28）——版式、
p:bg 背景、全部形状（图片/组合/freeform）、字号颜色原样保留，仅替换文字；
封面把箭头条加高到 1.89in（中心不变）并把标题框改为 1.70,1.70,5.90×1.60 垂直居中，容纳两行
28pt 标题（title 含 \n 即两行）。目录用 layout2 与内容页同款标题样式。
特殊页同样接受 font_lint 与 verify_deck 校验（verify 对特殊页只记 NOTE 不判缺标题）。

原语可任意组合，也可直接用 python-pptx 原生 API 加柱图/时间轴/矩阵。图标共 **100 个**，随 skill 分发（`assets/icons/`，完整清单见 `assets/icons/INDEX.txt`）：
- **模板原生 7 个**（image21–image33，颜色固定）：蓝系 `image21` 人像 / `image23` 闪电 / `image25` 开箱直接放；白系 `image27` 图表 / `image29` 用户 / `image31` 组件 / `image33` 负面须 `circle_fill=` 垫蓝圆底。
- **同系列补充 93 个**（`icon_<语义名>.png`，实心 glyph，#4472C4，与原生蓝系同风格）：覆盖 E/S/G 与商业咨询常用语义（`icon_eco` 生态、`icon_recycling` 循环、`icon_water_drop` 水滴、`icon_co2` 碳排放、`icon_factory` 工厂、`icon_handshake` 合作、`icon_account_balance` 银行、`icon_trending_up` 趋势升、`icon_bar_chart` 柱图、`icon_rocket_launch` 启动……）。蓝系直接放；要垫蓝圆底时传 `tint=WHITE` 现场染白：
  `d.icon(s, "icon_eco.png", x, y, 0.52, circle_fill=PRIMARY, tint=WHITE)`

命令行工具（`scripts/`）：

| 脚本 | 用途 |
|---|---|
| `esg_deck_kit.py --check` | 自检基底与图标资产 |
| `starter_deck.py <out.pptx>` | 生成 8 页手法演示样板（封面+目录+5 内容页+结尾页，**仅参考**） |
| `font_lint.py <deck>` | 字号合规审计（必须无违规） |
| `pptx_render_preview.py <deck> <png目录>` | 逐页渲染，供目检 |
| `verify_deck.py <deck>` | 结构断言（识别元素继承 + 无分节残留） |
| `strip_sections.py <deck> [--check]` | 清除 deck 分节残留（新建已自动，仅既有 deck 需补清） |
| `analyze_deck.py <deck>` | 反解被认可 deck 的经验参数 |
| `probe_base.py` | 探查基底母版/版式结构 |

## 已知坑（都踩过）

- **python-pptx 解释器**：用 `~/.workbuddy/binaries/python/envs/default/bin/python`（本机已装 pptx+PIL）；托管 3.13 纯净版无 pptx。**换机器 / 同事首次使用**：先 `~/.workbuddy/binaries/python/envs/default/bin/pip install -r <skill目录>/requirements.txt`（python-pptx + Pillow），否则构建与渲染脚本会报缺模块。
- **中文字体**：必须 latin/ea/cs 三个 typeface 都设（`set_font` 已处理）；只设 `font.name` 只写 latin。
- **清空基底页**：必须 `drop_rel` + `_sldIdLst.remove`，只 remove 会留悬空关系。
- **分节会跟着留下**：`presentation.xml` 的 `p14:sectionLst` 独立于页码存在，删页不会删它，其 slideId 指向已删页 → 打开文件分节视图出现残留节名，可能触发修复提示（基底已预清，此坑主要出现在外部拿来的既有 deck）。kit 的 `_clear_slides()` 已调用 `strip_sections()` 自动清除；对既有 deck 用 `strip_sections.py`（只改 `presentation.xml`，其余条目字节不变，默认备份到 `/tmp/esg_ppt_bak/`）。
- **表格 CJK 对齐**：默认表格样式会把中文右对齐/居中，正文必须显式 `PP_ALIGN.LEFT`（`d.table` 已处理）。
- **qlmanage 在沙箱内失效**（sandbox initialization failed）、Keynote 导出超时 → 用 `pptx_render_preview.py`（Pillow 几何重绘，表格按真实行列宽/字号/颜色换行绘制）。渲染器不画母版元素，logo/装饰条不在预览里属正常，由 `verify_deck.py` 结构断言兜底。
- **Edit 工具偶发假成功**：对同一脚本的多处修改改用 Python 整体重写或串行单处修改，改完 grep 验证。
- **图标内缩**：白色图标垫圆底时图标内缩 22%（`d.icon` 已处理），不要手工再缩。
- **封面/结尾页的窄值框**：基底「汇报人」「联系人」值框原宽仅 1.13in（同级「时间」「电话」框 2.24in），5 个汉字加边距就折行。`cover()` / `end()` 已调 `_fit_text_width()` 放宽并设 `wrap="none"`；组内文本框按 group 的 `chExt→ext` 比例换算子坐标宽度，自绘同类位置时照做。
- **卡片高度要匹配字号**：14pt 正文下卡片净高建议 ≥1.5in，卡内每段正文 ≤2 行；装不下时靠分层（拆文本框、精简措辞）解决，不许缩字号。
- **借用手法时字号必须重设**：参考任何外部 deck 的手法时，字号一律按规范档位（9/12/14/15/24/30/34）重设，**不得照抄原文件字号**。kit 的 `set_font` 会拦截 <9pt；`font_lint.py` 拦其余档位问题。
- **表格行高是「最小值」**：PPTX 渲染时行高只会随内容放大、不会缩小——压缩行高（按 fit_to 等比缩小设定值）不能防止溢出，只会掩盖真实底边。装不下时唯一正解是拆页或上移起点；构建期应按**自然行高**（估算换行数 × 行距 + padding）告警越界，而非压缩后告警。
- **页脚来源行长度**：注释行（9pt）一旦折行，第二行会被母版底条裁切。来源行按单行设计：材料显示名用可识别短名（如「港交所《气候信息披露实施指引》(2024)」），全名放链接指向的原文；多材料并引时每条只保留页码不加括注。
- **富文本标记必须全量解析**：内容稿里的 `**强调**` 可能出现在段首、段中、整段——写 run 解析器时按「** 交替切分，奇数段粗体」实现，不要只匹配段首粗头，否则中段强调会以字面 `**` 渲染出来（交付前 grep 成品文本扫残留）。
