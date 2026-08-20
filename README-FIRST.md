# SIMS Blue Ocean Screener v0.5.0

## 変更点
- SIMS共通UI方針に合わせ、DrivePickerの見た目・操作を統一
- 保存先フォルダーを実際のPackage保存先へ確実に反映
- 前回のキーワード読込フォルダー／設定フォルダーからファイル選択を開始
- キーワード読込、一次選抜、SERP結果登録、カニバリ結果登録後に「結果概要＋次工程ボタン＋閉じる」を表示
- 閉じても処理済み状態を保持
- カニバリEvidence選択画面に「誰が作った何のファイルか」の説明を追加
- 時間のかかるDrive読込・解析・ZIP生成にスピナー表示
- 生成ZIP名の先頭を `SIMS-BOS-` に短縮
- CandidatesにSBM連携列を追加し、メニュー「9. SBM登録結果を記録する」を追加
- SBM Article ID / 公開URL / MONITORING状態をBOS候補へ紐付け可能

## Apps Scriptで置換するファイル
- `Code.gs` : 置換
- `DrivePicker.html` : 置換

## 変更なし
- スプレッドシート本体は既存のものをそのまま利用できます。
- 既存Candidatesデータは先頭13列を維持し、右側にSBM連携4列を追加します。

## 推奨コミットメッセージ
`feat: release SIMS Blue Ocean Screener v0.5.0 unified workflow UI and SBM feedback tracking`
