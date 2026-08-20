# SIMS Blue Ocean Screener v0.1.11

## 主な変更
- `6. カニバリ精査結果を登録する` を追加
- `SIMS_BOS_CANNIBAL_REVIEW_RESULT_V1` JSONをDriveから選択してCandidatesへ反映
- request_id照合を実施
- 最終GREEN/YELLOW/BLOCKとCannibalization LOW/MEDIUM/HIGHを登録
- カニバリ精査の article_scope / existing_article_boundary / matched_articles / internal_link_candidates を保持
- Creator依頼文へ上記のカニバリ防止情報を自動挿入
- Creator依頼文は最終GREEN候補だけ生成可能

## v0.1.10からの置換
- Code.gs: 置換
- DrivePicker.html: 置換
- appsscript.json: 変更なし

## 実機試験
1. Code.gs / DrivePicker.html を置換
2. スプレッドシートを再読み込み
3. `6. カニバリ精査結果を登録する`
4. `SIMS-BOS-Cannibal-Review-Result-SBOS-CANNIBAL-20260821-041645.json` を選択
5. GREEN 2件とCannibalization LOW/MEDIUMの反映を確認
6. GREEN行を選択して `8. Creator依頼文を作成する`
