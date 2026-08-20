# CHANGELOG

## v0.1.2 - 2026-08-20

- Consolidate all Apps Script `.gs` modules into a single `Code.gs` for SIMS-style maintainability.
- Keep `DrivePicker.html` separate as required by HtmlService.
- Fix Candidates output range width: 13-column rows are now written to a 13-column range.
- Synchronize user-facing prototype version strings with v0.1.2.
- Retain v0.1.1 first-runtime-test behavior and regression expectations.

## v0.1.1 - 2026-08-20

- Fix 3-word / 4-word counts when source files do not contain a word-count column.
- Separate preliminary Pre Score from final Blue Ocean Score.
- Add parent-folder navigation to the Google Drive picker.
- Persist processing state to both Document Properties and the hidden state sheet.
- Keep Home product/version display synchronized with the code version.
- Prepare the package for the first spreadsheet runtime test.

## v0.1.0 - 2026-08-20

- Initial prototype.
