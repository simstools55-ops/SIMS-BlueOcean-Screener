# SIMS Blue Ocean Screener v0.5.3

## UI PATCH
Creator依頼文ダイアログの操作性を改善しました。

### 修正内容
- 「全文をコピー」ボタンを追加
- 「閉じる」ボタンを追加
- コピー成功時に「コピーしました」と表示
- Clipboard APIが使えない場合は従来方式へフォールバック
- コピー失敗時は全文選択状態にして Ctrl+C でコピーできるよう案内
- SIMS共通UIに合わせて青い主操作ボタン＋灰色の閉じるボタンへ統一

## Apps Scriptで置換するファイル
- Code.gs : 置換

## 変更なし
- DrivePicker.html
- 既存スプレッドシートデータ
- SERP / カニバリ判定ロジック
- Creator依頼文の内容
- SBM/BOS連携データ

## 推奨コミットメッセージ
fix: release SIMS Blue Ocean Screener v0.5.3 creator referral copy and close controls
