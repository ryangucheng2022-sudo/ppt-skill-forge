#!/usr/bin/env python3
"""check_consistency.py — self-checks for the ppt-skill-forge package.

  errors (exit 1):
    - re-running derive_whitelist.py on examples/meridian-ppt/fingerprint.json does
      not reproduce examples/meridian-ppt/derived-font-sizes.json byte-for-byte
      (derivation must be deterministic);
    - forge_skill.py smoke run misses a required output file, or leaves a
      {{PLACEHOLDER}} unsubstituted in the generated SKILL.md;
    - the generated checker differs from assets/checkers/ by more than the
      program-name strings (checking logic must never drift);
    - a shipped references/ or templates/ file is not mentioned in SKILL.md.

Usage: python3 scripts/check_consistency.py  (needs python-pptx for the smoke run)
"""
import json, os, subprocess, sys, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def main():
    errors = []
    # 1. deterministic derivation
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, 'w.json')
        r = subprocess.run([sys.executable, os.path.join(ROOT, 'scripts', 'derive_whitelist.py'),
                            os.path.join(ROOT, 'examples', 'meridian-ppt', 'fingerprint.json'), '--out', out],
                           capture_output=True, text=True)
        if r.returncode != 0:
            errors.append(f'derive failed: {r.stderr[:200]}')
        elif open(out, encoding='utf-8').read() != open(os.path.join(ROOT, 'examples', 'meridian-ppt', 'derived-font-sizes.json'), encoding='utf-8').read():
            errors.append('derived whitelist is not reproducible from the bundled fingerprint')
        # 2. forge smoke
        pkg = os.path.join(td, 'pkg')
        r = subprocess.run([sys.executable, os.path.join(ROOT, 'scripts', 'forge_skill.py'),
                            '--company', 'Smoke Co', '--fingerprint', os.path.join(ROOT, 'examples', 'meridian-ppt', 'fingerprint.json'),
                            '--whitelist', out, '--out', pkg], capture_output=True, text=True)
        if r.returncode != 0:
            errors.append(f'forge smoke failed: {r.stderr[:200]}')
        else:
            for rel in ('SKILL.md', 'scripts/font_sizes.json', 'scripts/check_geometry.py',
                        'scripts/check_content.py', 'references/style-fingerprint.md'):
                if not os.path.isfile(os.path.join(pkg, rel)):
                    errors.append(f'forge smoke missing {rel}')
            skill = open(os.path.join(pkg, 'SKILL.md'), encoding='utf-8').read()
            if '{{' in skill:
                errors.append('generated SKILL.md has unsubstituted placeholders')
            if 'Smoke Co' not in skill:
                errors.append('generated SKILL.md lacks the company name')
            # 3. checker logic integrity: only name strings may differ
            gen = open(os.path.join(pkg, 'scripts', 'check_geometry.py'), encoding='utf-8').read()
            orig = open(os.path.join(ROOT, 'assets', 'checkers', 'check_geometry.py'), encoding='utf-8').read()
            gen_norm = gen.replace('Smoke Co PPT geometry checker', 'Enterprise PPT geometry checker').replace('Smoke Co PPT 版面几何校验器', 'Enterprise PPT 版面几何校验器')
            if gen_norm != orig:
                errors.append('generated checker logic differs from assets/checkers/')
    # 3b. worked-example corpus reproducibility: regenerate assets, re-extract, compare
    try:
        with tempfile.TemporaryDirectory() as td:
            r = subprocess.run([sys.executable, os.path.join(ROOT, 'examples', 'meridian-ppt', 'make_assets.py'), td],
                               capture_output=True, text=True)
            if r.returncode != 0:
                errors.append(f'make_assets failed: {r.stderr[:200]}')
            else:
                import glob as _g
                fp_out = os.path.join(td, 'fp.json')
                r = subprocess.run([sys.executable, os.path.join(ROOT, 'scripts', 'extract_fingerprint.py'),
                                    fp_out] + sorted(_g.glob(os.path.join(td, '*.pptx'))),
                                   capture_output=True, text=True)
                if r.returncode != 0:
                    errors.append(f're-extract failed: {r.stderr[:200]}')
                else:
                    new = json.load(open(fp_out, encoding='utf-8'))
                    old = json.load(open(os.path.join(ROOT, 'examples', 'meridian-ppt', 'fingerprint.json'), encoding='utf-8'))
                    for key in ('font_sizes', 'text_colors', 'corner_furniture', 'total_slides'):
                        if new.get(key) != old.get(key):
                            errors.append(f'worked-example fingerprint drift on {key}: corpus or extractor changed without refreshing the example')
    except Exception as e:
        errors.append(f'corpus reproducibility check errored: {e}')
    # 4. no orphan shipped files
    skill_text = open(os.path.join(ROOT, 'SKILL.md'), encoding='utf-8').read()
    for sub in ('references', 'templates'):
        d = os.path.join(ROOT, sub)
        if os.path.isdir(d):
            for name in sorted(os.listdir(d)):
                if name not in skill_text and os.path.splitext(name)[0] not in skill_text:
                    errors.append(f'orphan file not mentioned in SKILL.md: {sub}/{name}')
    print(json.dumps({'root': ROOT, 'errors': errors}, ensure_ascii=False, indent=1))
    return 1 if errors else 0

if __name__ == '__main__':
    sys.exit(main())
