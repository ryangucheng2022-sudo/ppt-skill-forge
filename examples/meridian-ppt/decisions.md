# Whitelist derivation — decision list

Corpus: 9 slides, 59 sized runs, 13 distinct sizes.
Derived whitelist: 6 sizes, directly covering 47 runs (79.7%); auto-fit values account for 5 runs and are tolerated by design.

## Items needing a human decision

- **Sub-floor sizes** (below 8.0pt): 7.0pt ×3. If these are template furniture (footnotes, page numbers), set an asset-mode floor at or below 7.0pt for baseline checks instead of whitelisting them.
- **Singleton sizes** (seen once, at/above floor): 20.0pt, 24.0pt, 28.0pt, 54.0pt. One-off usage — confirm whether each is a deliberate display size or a mistake before whitelisting.
- **Auto-fit clusters** (no action expected): 10.0pt ← 9.85×2; 12.0pt ← 11.41×3. Values hugging the underside of a declared size are auto-shrink artifacts.

## Corpus coverage warning

Sizes never observed cannot be derived. Verify the corpus contains covers, chapter dividers and closing pages; display sizes for those page types must be added from the style guide or the templates themselves before this candidate becomes the standard.
