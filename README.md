# Diet Plan

This repo hosts a public mobile-friendly diet dashboard and PDF backup.

## Public website

Once GitHub Pages is enabled for the repo, the site URL should be:

`https://balamm27.github.io/Diet-Plan/`

That is the link you can share with anyone. They do not need GitHub, a login, or an app.

## What is in this repo

- `index.html`: the public site
- `site-data.json`: the structured meal, recipe, and shopping data used by the site
- `output/weekly-diet-dashboard.pdf`: printable backup version
- `scripts/build_dashboard.py`: rebuilds the site from the workbook
- `scripts/export_pdf.mjs`: rebuilds the PDF from the generated site

The weekly buy quantities shown on the site are estimated for one person following this exact weekly plan.

## Simple update steps

1. Put the latest workbook at the repo root as `Diet Plan.xlsx`.
2. Run:

```bash
/Users/bala/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_dashboard.py
/Users/bala/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node scripts/export_pdf.mjs
```

3. Commit the changed files.
4. Push to GitHub.
5. GitHub Pages republishes the same public link.

## Notes

- The workbook file is ignored by Git so it does not get published automatically.
- Recipes stay inside the workbook ingredient list plus pantry basics:
  - salt
  - black pepper
  - water
  - neutral oil or olive oil
