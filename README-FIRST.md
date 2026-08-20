# SIMS Blue Ocean Screener v0.1.7

## 今回の更新

v0.1.6の実機試験で確認された、同一Intentの別キーワード間でSERP結果が上書きされる問題を修正しました。

SERP結果登録時は `main_keyword` を基本完全一致で照合し、その後にIntent Clusterを使って代表候補とSupporting Queryを整理します。

また、「ない」を一律に同じ意味へ正規化する処理を廃止し、否定表現を文脈別に扱います。

## v0.1.6からの更新方法

Apps Scriptでは次のファイルだけを置換してください。

- `Code.gs` : 置換

次のファイルは変更ありません。

- `DrivePicker.html`
- `appsscript.json`

## 実機再確認

v0.1.7反映後、前回と同じ `SIMS_BOS_SERP_REVIEW_RESULT_V1` JSONをもう一度「4. SERP精査結果を登録する」から登録してください。

期待結果:

- `iphone17 0 充電できない` : YELLOW / Blue Ocean Score 69
- `iphone17 0パーセント 充電できない` : BLOCK / SERP Status CLUSTERED / Supporting Query扱い
- `iphone17 セージ 人気ない` : CANNIBAL_PENDING / SERP GREEN / 82
- `iphone17 ゲーム 音が出ない` : CANNIBAL_PENDING / SERP GREEN / 80

SERP GREENはまだ新記事確定ではありません。次工程のカニバリ検査を通過して最終GREENになります。
