# Changelog

## v0.1.6 - 2026-08-21

### Added
- Google Drive selection flow for ChatGPT SERP review result JSON.
- `4. SERP精査結果を登録する` menu action.
- Import/validation for `SIMS_BOS_SERP_REVIEW_RESULT_V1`.
- Automatic reflection of SERP decision, Blue Ocean Score and Evidence Summary into Candidates.
- Hidden `_SerpReview` archive sheet preserving the complete SERP review result.
- `CANNIBAL_PENDING` guard state for SERP GREEN candidates.

### Improved
- Intent normalization now treats charging-context standalone `0` and `0パーセント` as the same zero-battery intent.
- `nfc 設定 どこ` and `nfc どこ` are normalized to the same location intent.
- Post-review duplicate Intent candidates are consolidated and blocked from separate article creation.
- Candidate status coloring distinguishes cannibalization-pending, YELLOW and BLOCK states.

### Validation
- JavaScript syntax check passed for `Code.gs`.
- Regression check confirmed the two duplicate Intent pairs found in the first iPhone 17 SERP review are now clustered together.
