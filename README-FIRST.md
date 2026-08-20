# SIMS Blue Ocean Screener v0.1.3

Complete Single-Code distribution for the first runtime test.

## Important change from v0.1.2

The Google Drive file picker UI is now embedded directly in `Code.gs`.
**Do not create `DrivePicker.html`.** Only `Code.gs` is required for the application source.

This fixes the runtime error:

`Exception: 「DrivePicker」というHTMLファイルは見つかりませんでした。`

## Files

- `Code.gs` — complete application source including the Drive picker UI
- `appsscript.json` — optional manifest reference
- `README-FIRST.md`
- `CHANGELOG.md`

## Apps Script replacement instructions

If v0.1.2 is already installed:

1. Open Apps Script.
2. Replace the entire contents of `Code.gs` with the v0.1.3 `Code.gs`.
3. If a `DrivePicker.html` file was created previously, it may be deleted; v0.1.3 does not use it.
4. Save the project.
5. Return to the spreadsheet and reload the browser tab.
6. Confirm the `SIMS Blue Ocean Screener` menu appears.
7. Run `1. キーワードファイルを読み込む`.

## First runtime test target

Using the uploaded Rakko Keyword iPhone 17 dataset, the expected counts are:

- Total: **993**
- 3-word: **281**
- 4-word: **41**

## Runtime scope

The current runtime scope covers Drive file selection, CSV/TSV import including UTF-16 tab-separated exports, normalization, word-count extraction, preliminary intent clustering, preliminary screening, Candidates display, and the Writer referral skeleton.

Actual SERP evaluation, SIMS Evidence automatic import, full cannibalization evaluation, and verified 4-word expansion remain later implementation stages. Candidates remain `PENDING` until final evidence checks are connected.

## Update policy

From v0.1.3 onward, ordinary application updates should normally require replacing **`Code.gs` only**.
