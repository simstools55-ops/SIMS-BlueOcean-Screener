# SIMS Blue Ocean Screener v0.1.4

## Apps Scriptへ入れるファイル

- `Code.gs` — Apps Script本体。SIMSの処理コードはこの1本に統合しています。
- `DrivePicker.html` — Google Driveファイル選択ダイアログ専用の画面部品です。
- `appsscript.json` — マニフェスト。

## v0.1.3からの更新

1. `Code.gs` をv0.1.4へ丸ごと置換してください。
2. Apps Scriptの「＋」→「HTML」で `DrivePicker` を作成し、`DrivePicker.html` の内容を丸ごと貼り付けてください。
3. 保存後、スプレッドシートを再読み込みしてください。

`DrivePicker.html` は処理ロジックではなくUI部品です。通常のSIMSロジック更新は引き続き `Code.gs` 1本の置換を基本とします。

## 第1実機試験

SIMS Blue Ocean Screener → `1. キーワードファイルを読み込む` を実行し、Google Driveのフォルダー・CSVファイル一覧が表示されることを確認してください。

今回のラッコキーワードCSVを選択した場合の期待値:

- 総件数: 993
- 3語: 281
- 4語: 41
