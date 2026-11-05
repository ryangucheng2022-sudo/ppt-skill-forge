# Method — the five stages in detail

## Stage 1 · Inventory

Collect every pptx the company can offer: the official template, seed decks,
page libraries, and a sample of real delivered decks. Record per file: slide
count, canvas size, and its probable role (template / seed / library /
delivered). Then stratify by **page type**: covers, tables of contents,
chapter dividers, body pages, closings. A corpus missing a page type will
produce a standard missing that type's display sizes — flag the gap now,
not after derivation.

## Stage 2 · Fingerprint

`extract_fingerprint.py` walks every slide's shape tree (including groups
and tables) and records:

- **font-size histogram** over explicit run sizes — the raw material of the
  whitelist;
- **fonts** (latin + east-asian) by run share;
- **solid fills** by use count — where the brand colors actually live;
- **corner furniture**: small shapes clustered by position in the top-left
  zone. Recurring clusters are layout furniture (logos, corner marks) and
  often encode hard specs — in the worked example a planted corner spec
  was rediscovered exactly, position and size, from clustering alone;
- **theme** colors/fonts, recorded only for contrast (rule: usage outranks
  declarations).

## Stage 3 · Derive + decide

`derive_whitelist.py` applies the three-layer rule (quarter-grid regulars →
whitelist; non-quarter → auto-fit class; sub-floor + singletons → decision
items) and writes a decision list. The human checkpoint is batched: review
the list once, sign off or reject each group, and record the outcome in the
whitelist comment. Also review **near-brand colors** from the fingerprint
(e.g. two reds in use, neither equal to the official brand red): pick the
standard, list the variants as legacy.

## Stage 4 · Generate

`forge_skill.py` assembles the company package: SKILL.md from the template
(brand base filled from the fingerprint), the signed-off `font_sizes.json`,
the generic checkers (logic byte-identical; only program names rewritten),
the fingerprint and decision list as references, and the source assets.
The package is self-contained and belongs to the company.

## Stage 5 · Verify

Run the generated checker over the company's own baseline assets with
`--asset-mode` (relaxed floor for template furniture): all must PASS.
Then build a defect deck (a shape partially covering text) and confirm the
checker FAILs it. Record both in a regression note inside the generated
package's references. If a baseline asset fails, the failure list should
consist only of items already on the decision list — anything else means
the fingerprint or derivation missed something; go back a stage.
