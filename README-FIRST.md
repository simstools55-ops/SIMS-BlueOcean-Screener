# SIMS Blue Ocean Screener v0.6.0

## 新機能：複数ブログの保存・切替・復元

v0.5.xまでは、Keywords / Candidates / 状態が1ブログ分しかなく、
別ブログのキーワードファイルを読み込むと前ブログの作業状態を上書きする設計でした。

v0.6.0ではブログ単位のセッション保存を追加しました。

### 保存する内容
- Keywords
- Candidates（GREEN/YELLOW/BLOCK）
- Creator Status
- SBM Article ID
- 公開URL
- BOS Outcome / MONITORING
- Settings
- _State（SERP/Cannibal request_id等）
- SERP精査アーカイブ
- 入力ファイル名・処理状態

### 操作
追加の操作 → 「対象ブログを切り替える・再開する」

- 保存済みブログを選択して復元
- 新しいブログを開始
- 切替直前に現在ブログを自動保存
- 同じURLの保存済みブログを新規入力しても、上書きせず復元

### 自動保存タイミング
- キーワード読込後
- 一次選抜・4語深掘り後
- SERP精査結果登録後
- カニバリ精査結果登録後
- Creator依頼文作成後
- SBM登録結果をBOSへ記録後
- ブログ切替直前

### 新しい隠しシート
- _BlogSessions
- _SessionKeywords
- _SessionCandidates
- _SessionSettings
- _SessionState
- _SessionSerpReview

利用者が直接操作する必要はありません。

## v0.5.3からの置換
- Code.gs : 置換
- DrivePicker.html : 変更なし

既存のスマホ生活知識メモの現在データは、
最初に「対象ブログを切り替える・再開する」を開いた時点で自動保存されます。

## 実運用試験
1. v0.6.0 Code.gsを適用
2. 追加の操作 → 対象ブログを切り替える・再開する
3. 現在のスマホ生活知識メモが保存済み一覧に出ることを確認
4. 新しいブログを開始してキーワード探索
5. 再び切替画面からスマホ生活知識メモを選択
6. GREEN候補・Creator/SBM処理済み状態・未処理候補が復元されることを確認

## 推奨コミットメッセージ
feat: release SIMS Blue Ocean Screener v0.6.0 multi-blog session save and restore
