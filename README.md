# lhesg-ppt · lhesg 风格 PPT 生成 Skill

制作 lhesg 风格演示文稿。**skill 自包含**：内置基底
`assets/base.pptx`（母版 + 版式 + 主题 + 封面/结尾页原型）与品牌/图标资产，构建时
不依赖任何外部模板文件。视觉识别元素（标题蓝块灰块、双 logo、三段装饰条、主题色
`#044AC0`）由母版原生继承，正文版式按内容自由设计。

固化的对象是**设计经验**（`references/design-experience.md` E1–E6），不是某一版成品文件；`scripts/starter_deck.py` 只是手法演示，不是必须沿用的版式白名单。

## 依赖

- Python 3.9+，`python-pptx`、`Pillow`
- （无外部文件依赖；如需替换基底，设环境变量 `ESG_BASE` 指向自定义 base.pptx）

```bash
pip install -r requirements.txt          # 等价于 pip install python-pptx pillow
```

WorkBuddy 自带的托管解释器：
`~/.workbuddy/binaries/python/envs/default/bin/python`（未装依赖时先按上面的命令装进这个环境）。

## 安装（分发给同事）

**方式一 · 上传技能包（推荐，官方入口）**
把本目录（含 `SKILL.md`）打包成 zip 交给同事 → 同事打开 WorkBuddy 左侧「技能」面板 →
「添加技能」→「上传技能」→ 拖入或选择该 zip（有的版本需先解压、选中 `lhesg-ppt` 文件夹）→
系统自动完成配置，无需手动放路径。

**方式二 · 放进技能目录（兜底，必定生效）**
把 `lhesg-ppt` 文件夹放到：
- 用户级（本人所有项目可用）：`~/.workbuddy/skills/`
- 项目级（同一项目目录下的团队成员共享）：`<项目根>/.workbuddy/skills/`

Windows 对应用 `%USERPROFILE%\.workbuddy\skills\`。放好后新开一个对话即被识别。

**方式三 · Git 仓库导入（便于后续更新）**
```bash
git clone https://github.com/707292385-byte/lhesg-ppt-skill.git ~/.workbuddy/skills/lhesg-ppt
```
仓库默认为私有，需要先给同事开放访问（加 collaborator 或改公开）。

> 装完验证：在对话框说「用 lhesg-ppt 做一页封面」，能跑通即安装成功。
> 若提示缺少模块，按上面「依赖」一节把 python-pptx / Pillow 装进对应解释器。


## 目录

```
SKILL.md                      技能入口：流程、硬约束、工具用法
requirements.txt              依赖清单（python-pptx / Pillow）
references/
  design-experience.md        E1–E6 设计经验（本 skill 的核心资产）
  visual-spec.md              几何/字号/配色规格与特殊页实现说明
scripts/
  esg_deck_kit.py             构建器：EsgDeck（page / cover / toc / end / table / shape …）
  starter_deck.py             8 页手法演示（封面+目录+5 内容页+结尾）
  pptx_render_preview.py      版式渲染预览（替代 qlmanage，支持 p:bg 与垂直锚点）
  font_lint.py                字号合规检查
  verify_deck.py              结构断言（母版继承元素、分节残留）
  strip_sections.py           清除既有 deck 的分节残留
  analyze_deck.py             从被认可的 deck 反解设计参数
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
2. **字号（唯一标准）**：正文类 ≥ 12pt（默认 14pt，信息密集页 12pt）；注释类（口径说明、单位、来源行）9pt。合法档位 9/12/14/15/16/24/28/30/34/44，10–11pt 禁止使用。交付前必跑 `font_lint.py`。
3. **版式是开放集**：只约束识别元素、字号、安全区、色板；版式按当次内容重新设计，判断标准是「观点讲清楚」。
4. **特殊页只克隆不自绘**：封面 / 结尾页整页克隆基底原型页 slide1 / slide28，仅替换文字。封面**不设副标题**。
5. **不留分节残留**：`presentation.xml` 的 `p14:sectionLst` 独立于页码，删页不会随之删除（基底已预清，kit 建 deck 时再清一次）。
6. **渲染验证必做**：`pptx_render_preview.py` 逐页渲染后目检（文字溢出 / 压线 / 留白失衡 / 相邻页同构）。
7. **全页结论置顶**：有全页结论的页，结论条紧贴标题下方、排在图表证据之上（顺序：标题 → 全页结论 → 证据 → 来源行），用 `d.conclusion(...)` 实现。

## 版权说明

基底与品牌素材（`assets/`）版权归原权利人所有，仅限授权主体内部使用。
