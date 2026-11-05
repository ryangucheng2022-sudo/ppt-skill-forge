# Changelog

## v0.1.0 (2026-11-05)

Initial release.

- Five-stage method (inventory → fingerprint → derive + decide → generate →
  verify) documented in SKILL.md and references/method.md.
- scripts/extract_fingerprint.py: style fingerprint from observed usage —
  font-size histogram, fonts, solid fills, corner-furniture clusters, theme
  for contrast.
- scripts/derive_whitelist.py: three-layer whitelist derivation
  (quarter-grid regulars / auto-fit class / decision items) emitting a
  font_sizes.json candidate plus a human decision list.
- scripts/forge_skill.py: generates a self-contained company PPT skill
  (SKILL.md, generic checkers with config-driven specificity, fingerprint
  and decision references, assets).
- Worked example: fictional, fully reproducible Meridian corpus in
  examples/meridian-ppt/ — raw derivation, decision items, sign-off and
  the negative control are all recorded. (The pipeline was also validated
  against a production enterprise skill during development.)
- scripts/check_consistency.py: derivation determinism, forge smoke test,
  checker-logic integrity, orphan-file check.
- CI kit in ci/ (ready to enable; see ci/README.md).
