# SIMS Blue Ocean Screener v0.1.9

v0.1.8のカニバリ精査Evidence選択後に「エラー」とだけ表示される問題を修正。

## 変更
- HTMLダイアログから呼ばれるPackage生成処理内のSpreadsheet UI alertを廃止
- Package生成処理を段階別に例外捕捉
- エラー時に「処理段階」と「詳細」をDrivePicker画面とalertへ表示
- Evidence ZIPをそのままカニバリ精査Packageへ同梱

## v0.1.8からの置換
- Code.gs: 置換
- DrivePicker.html: 置換
- appsscript.json: 変更なし
