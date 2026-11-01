---
name: douyin-viral-script
description: 抖音口播脚本生成器，三条产线：①45秒三段式拍摄脚本（文案+画面逐行配对，直接能开拍）②带货口播稿（722条真实成交文案提炼）③知识/AI类费曼口播稿（费曼学习法）。当用户要求「写抖音带货文案」「生成口播脚本」「做拍摄脚本」「按爆款结构写文案」「费曼学习法」「口播稿」时使用。
agent_created: true
version: 1.3.0
---

# 抖音口播脚本生成器

按要不要开拍选产线：

| 产线 | 输出 | 什么时候用 | 详细规范 |
|---|---|---|---|
| **① 45秒三段式** | **文案+画面逐行配对** | **要开拍** | `references/short45-guide.md` |
| ② 带货口播稿 | 纯口播文字 | 只念，或混剪配音 | `references/commerce-pipeline.md` |
| ③ 费曼知识稿 | 纯口播文字 | 讲概念、AI 科普 | `references/feynman-guide.md` |

**⚠️ 写知识类 / AI 类口播稿，必须用费曼学习法，这是硬要求。**

命令在 `<skill目录>` 下跑；`dvt.py list` 查素材，`show --id S2` 查句式。

---

## ① 45秒三段式拍摄脚本

**每一句文案都必须配一个画面指令**——只有口播没有画面，那还是稿子，不是拍摄脚本。

结构：开头 2 秒内出钩、5 秒完成；中段默认 3 个购买理由（重要产品 4 个），只放一个就是单薄；结尾两步（先点人群，再指路）。

```bash
python scripts/dvt.py short45 --product "产品名" --hook H1 --reasons 4 --out /path/s45.json   # hook 取 H1-H4
python scripts/dvt.py check45 --script /path/s45.json --range 180-260 --require-cta --strict
python scripts/dvt.py assemble45 --filled /path/s45.json --out /path/45秒脚本.xlsx
```

逐行填三样：**文案 / 画面 / 作用+目的**。画面要具体到能执行——✅「切开丸子，露出粉质横截面特写」❌「展示产品」。

**`NO_VISUAL` 报错就说明你写成了口播稿，补画面。** 输出六列：结构 / 时间 / 文案 / 画面 / 作用 / 目的。

---

## ② 费曼知识口播

**不要讲「是什么」，要讲「为什么会这样」。结论留不住人，因果链能留。**

观众理解得越快越容易划走：「听懂了」和「想继续听」是两件事。
五段顺序不可调：钩子（前3秒，约15字）→ 痛点共鸣 → 本期承诺 → 正文·费曼讲解 → 结尾 CTA。

正文必做四件事，少一件不合格：**生活类比**（必须同构，不是听起来像）/ **术语翻译**（专业词当场翻）/ **主动纠误解**（最容易漏：观众听到「很多人以为…其实不是」的顿悟，就是继续看的钩子）/ **节奏**（长句讲机制，短句砸结论）。

```bash
python scripts/dvt.py feynman --topic "大模型为什么会产生幻觉" --platform 抖音 --hook 反常识
python scripts/dvt.py check-feynman --scripts /path/feynman.json --range 260-330 --strict
python scripts/dvt.py assemble --filled /path/feynman.json --out /path/口播稿.xlsx
```

- `--platform` 抖音 260-330 字/60-90 秒 · 视频号 300-380 字（用家庭职场类比）· 小红书 220-280 字（结尾引导收藏）
- `--hook` FH1-FH4 或 反问/反常识/数字/冲突；不给则随机

**`NO_ANALOGY` / `NO_TRANSLATE` / `NO_MISCONCEPTION` 报错，说明这段不叫费曼稿，重写。** 先想清楚类比再动笔——类比想不出来说明你没吃透机制，硬写会变成罗列定义。

---

## ③ 带货口播稿

句式库来自 722 条真实转写提炼的 36 个句式原型，覆盖开口/痛点/卖点/背书/价格/收尾六个位置，**六个位置全填满**，缺一个脚本就松散。

**先确认品类再动笔**：食品日用＝家庭采购向（「孩子/我们」+ 口感质检数字 + 小黄车收尾）；女装美妆＝女性个人向（「姐妹」+ 设计上身体验 + 人群限定收尾）。换错画像整篇跑偏，对照表见 `references/corpus-findings.md` 第七节。

```bash
python scripts/dvt.py plan --skeleton K1 --product "产品名" --seed 42 --out /path/plan.json   # K1/K2/K3
python scripts/dvt.py check --scripts /path/filled.json --range 280-330 --require-cta --strict
python scripts/dvt.py assemble --filled /path/filled.json --out /path/带货脚本.xlsx --product-card /path/card.json
```

骨架：`K1` 60 秒 280-330 字（默认）· `K2` 30 秒 140-180 字（单卖点强打）· `K3` 90 秒 420-480 字（多卖点多背书）。

**卖点必须有数字或具体物**——「很好用」是废话，「通过 180 项质检」才是卖点。

完整流程（产品卡、槽位、批量 seed、装配列）见 `references/commerce-pipeline.md`；六位置写法 + 自检八问见 `references/section-guide.md`。

---

## 交付与文件

生成后用 `present_files` 交付。

- `assets/structures.json` — 结构库（45秒框架 / 带货 36 句式 / 费曼五段）
- `scripts/dvt.py` — 主工具，十个子命令；`build_*.py` 幂等注入，语料更新后重跑
