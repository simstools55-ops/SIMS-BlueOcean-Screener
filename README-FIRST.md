# SIMS Blue Ocean Screener v0.1.5

## Apps Scriptへ入れるファイル

- `Code.gs` — Apps Script本体。処理ロジックはこの1本に統合しています。
- `DrivePicker.html` — Google Driveのファイル／保存先フォルダー選択UIです。
- `appsscript.json` — マニフェスト。

## v0.1.4からの更新

- `Code.gs` を丸ごと置換してください。
- `DrivePicker.html` も丸ごと置換してください。
- `appsscript.json` は変更ありません。

## 新機能: SERP精査依頼Package

一次選抜後、メニュー `3. SERP精査依頼Packageを作成する` から、PENDING候補をChatGPTでWeb精査するためのZIPを作成できます。

ZIP内容:

- `README-FIRST.md`
- `SERP-REVIEW-REQUEST.md`
- `SERP_REVIEW_REQUEST_V1.json`

Package作成前に対象ブログとGoogle Drive保存先を設定します。保存先は `追加の操作 → 保存先を設定する` からWindows風Drive選択ダイアログで指定できます。

## 実用試験の次の手順

1. 一次選抜20件がある状態で `追加の操作 → 保存先を設定する` を実行。
2. 保存先フォルダーを選択。
3. `3. SERP精査依頼Packageを作成する` を実行。
4. 完了ダイアログに表示されたZIPをChatGPTへそのままアップロード。
5. ChatGPTにSERP精査を依頼。

SERP段階のGREENは新記事作成の最終GREENではありません。後段のカニバリ検査を通過して最終確定します。
