# PPT Skill Forge

[English](README.md) | 简体中文

把一家企业现有的 PPT 资产——官方模板、种子页、页型库、一堆历史稿——变成一份企业专属 PPT 技能：可执行的品牌底座、声明制的字号白名单，以及能自动检查每一份新稿的交付闸门。

## 为什么做这个

大多数企业*有*自己的汇报风格，但几乎没有一家把它沉淀成 AI Agent（或新人）能执行、能验证的形态。风格被困在模板文件和「上季度那份稿是谁做的」的习惯里，而文件里声明的主题往往还是 Office 默认值。本技能从**实际使用**中提取真实标准，把真正有歧义的发现汇成一张批量决策单交给人拍板，再生成一份自带检查器、能自我执行的技能包。

它是 [company-knowledge-to-skill](https://github.com/ryangucheng2022-sudo/company-knowledge-to-skill)（术语与业务口径）与 [business-agent-skill-builder](https://github.com/ryangucheng2022-sudo/business-agent-skill-builder)（业务流程）在演示文稿域的姊妹篇：口径 → PPT → 流程，正好成套。

## 工作方法

| 阶段 | 做什么 | 人工检查点 |
|---|---|---|
| 盘点 | 资产分类、语料按页型分层 | 先标注语料缺口，再谈推导 |
| 指纹 | 从实际使用提取：字号直方图、字体、填充色、角标位置聚类；主题仅作对照记录 | — |
| 推导 | 三层规则产出白名单候选＋决策单 | **拍板决策单**——下限以下字号、孤例字号、近似品牌色 |
| 生成 | 组装企业技能包：SKILL.md、通用检查器、`font_sizes.json`、指纹与决策参考、资产 | 确认品牌底座一节 |
| 验收 | 生成的检查器跑企业自有资产＋构造缺陷稿 | 判定须一致、缺陷须检出 |

设计原理：

- **使用优先于声明**：真实品牌活在直接格式里，不在主题里。演练示例中，模板主题是 Office 默认值，真正的品牌（字体、两个近似品牌蓝、一个角标）只存在于形状上。
- **推导值不等于标准**：孤例与下限以下字号永不静默入白名单。植入的角标规格被聚类分毫不差地重新发现——但它依然要过拍板才成为规格。
- **自动缩字是一类现象、不是错误**：贴在登记值正下方的非阶梯字号是自动缩字产物，系统报告并容忍，不撑大白名单。
- **检查器由配置驱动**：企业专属性只活在一个配置文件里，检查逻辑原样复制进生成的技能包。

## 包里有什么

- `SKILL.md`——方法本体：五阶段、六条可信规则、能力边界。
- `references/method.md`——五阶段的详细说明。
- `scripts/extract_fingerprint.py`——从实际使用提取风格指纹：字号、字体、填充色、文字颜色、角标位置聚类；主题仅作对照。
- `scripts/derive_whitelist.py`——三层规则推导白名单，并生成人工决策单。
- `scripts/forge_skill.py`——组装企业专属技能包。
- `scripts/check_consistency.py`——推导确定性、生成冒烟、检查逻辑完整性、示例语料可复现性。
- `templates/generated-skill.md`——生成包的 SKILL.md 模板。
- `assets/checkers/`——复制进生成包的通用几何/内容检查器（逻辑不变，专属性只在配置里）。
- `examples/meridian-ppt/`——虚构公司完整示例，可复现。
- `ci/`——备好的 CI 工具包，复制即可启用。

## 谁适合用

- 有模板、但没有可执行标准的品牌与演示负责人；
- 要统一周期性经营汇报口径的运营/PMO 团队；
- 需要快速为客户立起一套可辩护的 PPT 标准的顾问；
- 正在为企业搭建专属技能的 Agent 开发者（与知识技能配套）。

典型场景：让 AI Agent 上手一家企业的汇报风格；换标前审计存量模板；汇报旺季前冻结字号白名单。

## 快速开始

```bash
pip install python-pptx
python3 scripts/extract_fingerprint.py fingerprint.json path/to/assets/*.pptx
python3 scripts/derive_whitelist.py fingerprint.json --out font_sizes.json --decisions decisions.md
# → 审阅并拍板 decisions.md，按结论调整 font_sizes.json
python3 scripts/forge_skill.py --company "Acme" --fingerprint fingerprint.json \
    --whitelist font_sizes.json --decisions decisions.md \
    --assets-dir path/to/assets --out ./acme-ppt-skill
```

## 完整示例（虚构公司，可复现）

`examples/meridian-ppt/` 内置一家虚构公司的语料（由脚本生成，任何人可复现），并故意植入真实语料的典型现象：原始推导出 6 个白名单字号、4 个展示型孤例字号被扣为决策项；拍板前模板恰好只挂在这四项上，拍板后两份文件全过，构造缺陷仍被检出。完整记录见 `examples/meridian-ppt/regression.md`。本管线在开发期另以一家生产级企业技能的资产做过验证，拍板后判定等价。

## 验证结果（2026-11-05 实跑）

```bash
python3 scripts/check_consistency.py
# {"root": "...", "errors": []}——推导确定性、生成冒烟、检查逻辑完整性、孤儿文件检查全过
```

## 持续集成

`ci/` 里备有可直接启用的工具包：把 `ci/ci.yml` 复制到 `.github/workflows/ci.yml`，每次 push 与 PR 都会跑一致性检查——推导确定性、生成冒烟、检查逻辑完整性、示例语料可复现性。见 `ci/README.md`。

## 能力边界（诚实说明）

forge 完整建成的是**检查层**；生成技能的**生成层**起步只是骨架（通用页型＋种子克隆）。内容口径——标题规则、术语——无法从文件提取，应挖掘为候选后用 company-knowledge-to-skill 的机制拍板定版。

## 版本

当前版本：**v0.1.0**（2026-11-05）——见 [CHANGELOG.md](CHANGELOG.md)。

## 许可证

MIT
