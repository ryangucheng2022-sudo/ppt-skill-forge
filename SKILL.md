---
name: ppt-skill-forge
description: Turn a company's existing PPT assets into a company-specific PPT skill. Use when a team has templates, seed decks and historical presentations but no enforceable standard: inventory the assets, extract a style fingerprint from observed usage, derive a font-size whitelist with a human decision list, generate a self-contained skill package with delivery gates, and verify it against the company's own assets. Do not use for one-off deck production (use the generated company skill) or for inventing a brand style from nothing.
---

# PPT Skill Forge

Turn scattered presentation assets into an enforceable, self-checking PPT
skill for one company. The forge is method + three scripts; the company data
never leaves the generated package.

## Working flow

| Stage | Input | Output | Human checkpoint |
|---|---|---|---|
| 1. Inventory | A folder of pptx files | Asset list with slide counts and canvas | Corpus stratified by page type (covers, chapters, body, closing all present?) |
| 2. Fingerprint | The corpus | `fingerprint.json`: size histogram, fonts, fills, corner furniture, theme | — |
| 3. Derive | Fingerprint | Whitelist candidate + decision list | **Sign off the decision list** (sub-floor sizes, singletons, near-brand colors) |
| 4. Generate | Fingerprint + signed-off whitelist + assets | Company skill package (SKILL.md, checkers, config, references, assets) | Confirm the generated SKILL.md's brand-base section |
| 5. Verify | Generated package + the corpus | Verdict table per asset + a constructed-defect test | Generated verdicts must match expectations; defects must be caught |

Scripts: `scripts/extract_fingerprint.py`, `scripts/derive_whitelist.py`,
`scripts/forge_skill.py`. Method detail: [references/method.md](references/method.md).
Generated-package template: [templates/generated-skill.md](templates/generated-skill.md).
Worked example (fictional company, fully reproducible): [examples/meridian-ppt/regression.md](examples/meridian-ppt/regression.md).

## Rules that make the output trustworthy

1. **Usage outranks declarations.** A company's real style lives in direct
   formatting, not in the Office theme (which is usually the default).
   Fingerprint statistics come first; theme data is recorded for contrast.
2. **Derived ≠ standard.** The whitelist is a candidate until a human signs
   off the decision list. Singletons and sub-floor sizes are never silently
   whitelisted; near-duplicate brand colors are decision items, never
   silently merged.
3. **Auto-fit is a class, not an error.** Non-quarter font sizes hugging the
   underside of a declared size are auto-shrink artifacts — reported,
   tolerated, never registered.
4. **Coverage before confidence.** Sizes absent from the corpus cannot be
   derived. If the corpus lacks covers or chapter pages, say so in the
   deliverable and source display sizes from the templates/style guide.
5. **Checkers are configuration-driven.** All company specificity lives in
   `font_sizes.json`; the checking logic is copied unchanged (only program
   names are rewritten). A generated package whose checker logic differs
   from `assets/checkers/` is a failed generation.
6. **Verify against the owner's assets.** A generated skill is accepted
   only when it passes the company's own baseline assets (asset mode) and
   still catches a constructed defect. Both results are recorded.

## Honest boundaries

- The forge builds the *checking* layer completely. The *generation* layer
  of the produced skill starts as a skeleton: generic page types and
  seed-cloning; a company's distinctive page types remain design work.
- Content conventions (title rules, terminology, banned phrasing) cannot
  be extracted from files. Mine them from deck text as candidates and put
  them through a decision process — see the sibling skill
  company-knowledge-to-skill.
- Chart-internal text is not traversed by the fingerprint extractor yet.
