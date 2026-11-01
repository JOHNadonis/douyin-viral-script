# douyin-viral-script

抖音口播脚本生成器 —— WorkBuddy / Claude Skill。

三条产线，按「要不要开拍」选：

| 产线 | 输出 | 什么时候用 |
|---|---|---|
| **① 45秒三段式** | **文案 + 画面逐行配对** | **要开拍**（有剪辑师、要配图） |
| ② 带货口播稿 | 纯口播文字 | 只念，或混剪配音 |
| ③ 费曼知识稿 | 纯口播文字 | 讲概念、AI 科普 |

## 这套东西从哪来的

不是凭空编的，三个来源：

1. **722 条真实成交文案**（376 + 349，两个抖音带货号的全量转写稿）→ 提炼出 36 个句式原型，覆盖开口/痛点/卖点/背书/价格/收尾六个位置，每个都挂真实原文例句。
2. **一套经过验证的 45 秒框架** → 三段式（开头留人 0-5s / 中段说服 5-30s / 结尾转化 30-45s）、4 类开头钩子、7 种购买理由、2 步收尾，每句文案都必须配画面指令。
3. **费曼学习法** → 知识类口播的讲解规范。核心是：不要讲「是什么」，要讲「为什么会这样」。结论留不住人，因果链能留。

## 三条产线怎么配合

可以叠。典型用法：

```
45秒框架管骨架（保证画面齐、节奏对）
  +
费曼手法管讲解质量（中段加类比、翻术语、纠误解）
  +
品牌话术管味道（开头人情味、结尾「喜欢的话，左下角小黄车」）
```

## 安装

### 方式一：手动复制（最省事）

把整个目录复制到你的 skill 目录下：

```bash
# Windows
xcopy /E /I douyin-viral-script "%USERPROFILE%\.workbuddy\skills\douyin-viral-script"

# macOS / Linux
cp -r douyin-viral-script ~/.workbuddy/skills/
```

如果是 Claude Code 等其他支持 Anthropic Skills 规范的客户端，放到对应的 skills 目录即可。

### 方式二：git clone

```bash
git clone https://github.com/JOHNadonis/douyin-viral-script.git ~/.workbuddy/skills/douyin-viral-script
```

之后更新：`cd ~/.workbuddy/skills/douyin-viral-script && git pull`

### 依赖

只有 Python 标准库 + `openpyxl`（装配 Excel 用）：

```bash
pip install openpyxl
```

## 快速开始

```bash
cd ~/.workbuddy/skills/douyin-viral-script

# 看全部素材
python scripts/dvt.py list

# 产线①：45秒拍摄脚本（要开拍时用这条）
python scripts/dvt.py short45 --product "产品名" --hook H1 --reasons 4 --out plan.json
python scripts/dvt.py check45 --script filled.json --range 180-260 --require-cta --strict
python scripts/dvt.py assemble45 --filled filled.json --out 45秒脚本.xlsx

# 产线②：带货口播稿
python scripts/dvt.py plan --skeleton K1 --product "产品名" --seed 42 --out plan.json
python scripts/dvt.py check --scripts filled.json --range 280-330 --require-cta --strict
python scripts/dvt.py assemble --filled filled.json --out 带货脚本.xlsx

# 产线③：费曼知识稿
python scripts/dvt.py feynman --topic "主题" --platform 抖音 --hook 反常识 --out plan.json
python scripts/dvt.py check-feynman --scripts filled.json --range 260-330 --strict
```

## 目录结构

```
douyin-viral-script/
├── SKILL.md                        # 主文档（技能入口）
├── assets/
│   ├── structures.json             # 全部结构化数据
│   └── structures.json.bak         # v1.0 备份
├── references/
│   ├── short45-guide.md            # 45秒拍摄脚本规范
│   ├── feynman-guide.md            # 费曼口播规范 + 费曼×带货融合
│   ├── section-guide.md            # 带货六位置写法
│   └── corpus-findings.md          # 722条语料统计结论
└── scripts/
    ├── dvt.py                      # 主工具（10 个子命令）
    ├── build_structures.py         # 从语料重建带货句式库
    ├── build_feynman.py            # 注入费曼模块
    └── build_short45.py            # 注入45秒框架
```

## 子命令一览

| 命令 | 作用 |
|---|---|
| `list` | 列出全部素材（句式/钩子/理由/骨架） |
| `show --id O1` | 看单个素材详情 |
| `short45` | 导出 45 秒拍摄脚本模板（逐行文案+画面） |
| `check45` | 校验 45 秒脚本（多商品自动分组） |
| `assemble45` | 装配 45 秒脚本 Excel（六列） |
| `plan` | 导出带货口播骨架 |
| `check` | 校验带货稿字数/促单句 |
| `assemble` | 装配带货稿 Excel |
| `feynman` | 导出费曼知识稿大纲 |
| `check-feynman` | 校验费曼稿五要素 |

## 几条踩过的坑（写在这里省得你再踩）

**1. 45 秒脚本必须逐行配画面。** 只有文案没有画面，那还是口播稿，剪辑师拿到会问「这句放什么画面」。`check45` 会报 `NO_VISUAL`。

**2. 多商品校验要给每行加 `product` 字段。** 否则会把所有商品字数合并算，四篇一起算必然误报超长。

**3. 从 Excel 读回 JSON 要小心合并单元格。** 「结构」「时间」列纵向合并后只有首格有值，其余是 `None`。`check45` 已内置回填，但你自己写脚本读的时候要注意。

**4. 卖货稿不要出现元叙述。** 「这期讲完」「本期」「今天讲」这类话一说，观众立刻觉得「他在给我上课」，广告感暴露。`check45` 和 `check-feynman` 都会拦。

**5. 数字后面必须跟一句「这意味着什么」。** 原文案的通病是纯数字罗列（含氟0.1%、520mg、馅含量80%），观众听完只记住一串数。类比 + 机制翻译才是转化的关键。

**6. 品牌话术该重复就重复。** 「喜欢的话，左下角小黄车」这类固定句式不要为了「原创度」去改，观众认的就是这个调子。要避免的是整段照搬，不是句式复用。

## 版本

- **v1.2.1** — 修 check45 多商品统计（按商品分组，逐条独立算字数）+ 合并单元格回填
- v1.2.0 — 并入 45 秒三段式拍摄脚本产线
- v1.1.1 — 去掉元叙述，恢复品牌话术，新增「费曼 × 带货融合」
- v1.1.0 — 新增费曼学习法知识口播产线
- v1.0.0 — 基于 722 条真实成交文案的带货脚本生成

## License

MIT
