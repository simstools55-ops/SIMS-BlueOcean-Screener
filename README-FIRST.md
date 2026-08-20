# SIMS Blue Ocean Screener v0.1.2

## Distribution type

Single-Code distribution for easier SIMS maintenance and updates.

Apps Script source is consolidated into **one `Code.gs`**. `DrivePicker.html` remains separate because the Google Apps Script HTML service requires it as an HTML file.

## Files to add to Apps Script

- `Code.gs` — all Apps Script runtime code
- `DrivePicker.html` — Google Drive keyword-file picker UI
- `appsscript.json` — manifest (only when manifest editing/import is used)

`README-FIRST.md` and `CHANGELOG.md` are repository/distribution documentation and do not need to be added to the Apps Script editor.

## v0.1.2 runtime scope

- SIMS-style numbered menu
- Google Drive file picker with folder navigation and "one level up"
- CSV/TSV import
- UTF-8 / UTF-16LE / UTF-16BE BOM handling
- Rakko Keyword UTF-16 tab-separated export support
- automatic 3-word / 4-word counting when no word-count column exists
- normalization and preliminary intent clustering
- preliminary screening
- Candidates sheet generation
- separate `Pre Score` and unconfirmed `Blue Ocean Score`
- Writer referral generation only after a candidate is confirmed GREEN
- processing-state persistence skeleton

## Intentionally not yet connected

- live SERP provider and final Blue Ocean Score
- automatic SIMS Evidence ZIP import
- full existing-article cannibalization diagnosis
- external demand validation of generated four-word keywords
- output-folder picker

No direct Google Search scraping is implemented.

## First runtime regression test

Use:

`rakkokeyword_suggestKeywords_iPhone17_2026-08-20_20-36-03.csv`

Expected import counts:

- Total: **993**
- 3-word: **281**
- 4-word: **41**

## Update procedure

For v0.1.2, ordinary source updates should require replacing **`Code.gs` only** unless the Drive picker UI itself changed. Replace `DrivePicker.html` only when explicitly instructed.

## Release rule

If source code changes after this v0.1.2 artifact is issued, increase PATCH or higher. Do not reissue modified code under the same formal version.
