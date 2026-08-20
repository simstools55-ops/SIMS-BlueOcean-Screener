# SIMS Blue Ocean Screener v0.5.2

## UI PATCH
実運用試験で確認した軽微なUI不具合を修正しました。

### 修正内容
1. カニバリ精査Package作成完了画面
   - 「閉じる（ChatGPTへアップロード）」と「閉じる」の二重表示を解消
   - 外部作業が必要なため、ボタンは「閉じる」1個に統一
   - 次工程を「ChatGPTで精査 → 返却JSONをDrive保存 → 6.結果登録」と明記

2. エラー時のスピナー
   - request_id不一致、JSON契約エラー等のFailure時にbusy overlayを確実に解除
   - alertを閉じた後もスピナーが残らないよう二重解除を実施
   - 選択状態を維持し、再試行可能にする

3. 完了ダイアログのスクロール
   - 短い結果表示で不要なスクロールバーが出にくいよう寸法・overflowを調整
   - 共通完了ダイアログも少し拡張

## 変更ファイル
- Code.gs : 置換
- DrivePicker.html : 置換

## 変更なし
- 既存スプレッドシートデータ
- SERP判定ロジック
- カニバリ判定ロジック
- Creator/SBM連携データ
- 既に登録済みのBOS結果

## 推奨コミットメッセージ
fix: release SIMS Blue Ocean Screener v0.5.2 UI polish after end-to-end validation
