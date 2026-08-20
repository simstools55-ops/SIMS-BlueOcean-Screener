# SIMS Blue Ocean Screener v0.1.10

## 修正内容
- カニバリ精査Package生成時の `sbosSafeFilePart_ is not defined` を修正。
- ZIPファイル名用の安全化関数 `sbosSafeFilePart_()` を追加。
- `sbos...` 系の関数参照を静的点検し、未定義関数が他にないことを確認。

## v0.1.9からの置換
- Code.gs: 置換
- DrivePicker.html: 変更なし
- appsscript.json: 変更なし

## 再試験
1. Code.gs を丸ごと置換
2. 保存してスプレッドシートを再読み込み
3. `5. カニバリ精査Packageを作成する`
4. `SIMS-Evidence-スマホ生活ナビ-20260815-1545.zip` を選択
