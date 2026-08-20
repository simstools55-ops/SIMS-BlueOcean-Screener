# CHANGELOG

## v0.1.7 - 2026-08-21

- SERP精査結果の登録照合を `main_keyword` の基本完全一致優先へ変更。
- Intent正規化済みキーをSERP結果照合に使わないよう修正し、同一Intentの別表現による結果上書きを防止。
- `iphone17 0 充電できない` と `iphone17 0パーセント 充電できない` のような行を個別に正しく復元した後、代表候補/CLUSTEREDへ整理する仕様に修正。
- 否定表現の一律 `NOT_SHOWN` 化を廃止。
- `表示されない / 反応しない / 読み込めない / 読み取れない / 音が出ない / 在庫ない / 人気ない` を文脈別Intentトークンへ分離。
- 一般的な「ない」は意味を潰さず保持。

## v0.1.6

- ChatGPT SERP精査結果JSONの登録機能を追加。
- SERP GREENをカニバリ検査待ち `CANNIBAL_PENDING` として保持。
- Intent Cluster重複候補の統合処理を追加。
