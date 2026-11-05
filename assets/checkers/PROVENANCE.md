# Provenance

`check_geometry.py` and `check_content.py` are generic checkers taken from a
production enterprise PPT skill, copied verbatim except program-name strings.
All company-specific values live in the `font_sizes.json` config they load
(`--config`), not in the code — that separation is what makes this forge
possible. When `forge_skill.py` copies them into a generated package, it only
rewrites the program-name strings to the target company; the checking logic
is never modified.
