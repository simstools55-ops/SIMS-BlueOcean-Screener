# SIMS Blue Ocean Screener v0.1.6

## Purpose

v0.1.6 adds the return path from ChatGPT SERP review to SIMS Blue Ocean Screener.
The product keeps Apps Script runtime logic consolidated in `Code.gs`; `DrivePicker.html` remains the dedicated Drive selection UI.

## Replace for v0.1.5 users

- REPLACE: `Code.gs`
- REPLACE: `DrivePicker.html`
- NO CHANGE REQUIRED: `appsscript.json`

## New workflow

1. Import keyword file.
2. Run Blue Ocean screening.
3. Create SERP review package.
4. Upload the package to ChatGPT and obtain `SIMS_BOS_SERP_REVIEW_RESULT_V1` JSON.
5. Put the returned JSON in Google Drive.
6. Run `4. SERP精査結果を登録する` and select the JSON through the Drive picker.
7. Review Candidates.

## v0.1.6 behavior

- Validates `format = SIMS_BOS_SERP_REVIEW_RESULT_V1`.
- Validates the returned `request_id` against the package request ID stored in the spreadsheet.
- Writes SERP decision, Blue Ocean Score and Evidence Summary back to Candidates.
- Stores the complete returned result in hidden `_SerpReview` for audit/reference.
- A SERP `GREEN` does NOT become final GREEN. It is shown as `CANNIBAL_PENDING` until cannibalization review is completed.
- YELLOW and BLOCK are reflected immediately.
- Strengthens Intent clustering for equivalent expressions such as `0` / `0パーセント` in charging context and `nfc 設定 どこ` / `nfc どこ`.
- Duplicate Intent candidates are blocked from becoming separate article candidates.

## Important

Writer referral remains available only for final `GREEN` candidates. v0.1.6 deliberately prevents SERP-stage GREEN candidates from being sent to Writer before cannibalization review.
