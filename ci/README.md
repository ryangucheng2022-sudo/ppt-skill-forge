# CI kit

Ready-to-enable GitHub Actions setup, kept outside `.github/workflows/`
so it can be reviewed and versioned like any other file.

**Enable it:** copy `ci.yml` to `.github/workflows/ci.yml` and push.

The job installs python-pptx and runs the package consistency check:
derivation determinism against the bundled example fingerprint, a forge
smoke test (required files, substituted placeholders), and the
checker-logic integrity check.
