# SIMS Blue Ocean Screener v0.6.2

## 修正内容

### 1. カスタムメニューが消える問題への対策
- onOpen() で最初に SIMS Blue Ocean Screener メニューを生成
- シート整形処理でエラーが起きてもメニュー自体は表示される設計へ変更
- onInstall() も追加

### 2. Keywordsシートの視認性改善
通常表示：
- No
- キーワード
- 語数
- SEO難易度
- 月間検索数
- CPC
- 競合性
- 出現時期

非表示（内部保持）：
- Source
- SourceWordCount
- Normalized Keyword
- Intent Key
- Primary Candidate

### 3. Candidatesシートの視認性改善
通常表示：
- Rank
- 状態
- メインキーワード
- 語数
- Blue Ocean Score
- カニバリ
- 検索意図
- 判定根拠
- Creator状態
- SBM Article ID
- 公開URL

非表示（内部保持）：
- Pre Score
- Source
- Intent Key
- SERP Status
- BOS Outcome
- SBM Linked At

- 長文列は折返し
- 列幅を利用者向けに調整
- フィルタを追加
- SBM登録完了行のみ行全体をグレーアウト

### 4. 利用者向けタブ
表示：
- Home
- Keywords
- Candidates

非表示：
- Settings
- _State
- _ExistingArticles
- _SerpReview
- _BlogSessions
- 各Session内部シート

## 置換
- Code.gs : 置換
- DrivePicker.html : 変更なし

## 推奨コミットメッセージ
fix: release SIMS Blue Ocean Screener v0.6.2 restore menu and simplify user sheets
