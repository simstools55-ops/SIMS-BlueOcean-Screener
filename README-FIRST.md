# SIMS Blue Ocean Screener v0.1.8

## 今回の主な変更
- 新規記事連携先を Writer から SIMS Article Creator へ修正
- Candidates の `Writer Status` を `Creator Status` へ変更
- `Creator依頼文を作成する` に変更
- SERP GREEN候補向け「カニバリ精査Package」生成を追加
- Google DriveからSIMS Evidence ZIP/CSV/TSV/JSONを選択可能
- CLUSTERED候補は総合Statusも `CLUSTERED` と表示

## v0.1.7からの置換
- Code.gs: 置換
- DrivePicker.html: 置換
- appsscript.json: 変更なし

## 実機試験
1. Code.gs / DrivePicker.html を置換して保存
2. スプレッドシートを再読み込み
3. `5. カニバリ精査Packageを作成する`
4. 対象ブログのSIMS Evidenceを選択
5. 生成されたZIPをChatGPTへそのままアップロード
