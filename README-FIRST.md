# SIMS Blue Ocean Screener v0.5.1

## Hotfix内容
v0.5.0で、キーワード読込後の［2. ブルーオーシャン候補を探索］から実行すると、
一次選抜完了時に旧Apps Script標準alertが発火し、［3. SERP精査Packageを作成］ボタンが表示されない実装漏れを修正しました。

## 修正
- `sbosRunScreening_()` 内の旧「一次選抜・4語深掘り完了」標準alertを削除
- 一次選抜処理は結果メタ情報を返すだけに統一
- キーワード読込ダイアログからの遷移では、同じSIMS UI内で
  `結果概要 → [3. SERP精査Packageを作成] → [閉じる]`
  を表示
- 「処理を再開する」から一次選抜へ戻った場合も同じ完了UIを使用
- SERP精査Package作成成功時も旧標準alertを廃止し、SIMS共通結果ダイアログへ統一
- 「処理状態」表示も共通結果ダイアログへ統一

## UI方針
利用者向けの正常完了画面は原則、
1. 処理名
2. 結果概要
3. 次に行う操作の青ボタン（アプリ内で直接遷移可能な場合）
4. 閉じる
の順に統一します。

確認・エラー・入力不足など、利用者判断が必要な警告ダイアログは標準UIを残す場合があります。

## Apps Scriptで置換するファイル
- `Code.gs` : 置換
- `DrivePicker.html` : 置換

## 推奨コミットメッセージ
`fix: release SIMS Blue Ocean Screener v0.5.1 unified screening completion flow`
