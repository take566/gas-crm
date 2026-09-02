| `oauthScopes` を増やしたら `terraform/variables.tf` の `enabled_services` も見直す | GCP プロジェクト側で対応する API が無効だと、同意画面を通っても実行時に権限エラーで落ちる |
| GCP の設定変更はコンソールで直接行わず `terraform/` に書いて `apply` する | コンソールで直接変えると次の `terraform apply` で巻き戻る。Terraform で扱えない OAuth 同意画面と Apps Script の紐付けだけは例外で、手順は `terraform/README.md` にある |
# エージェント向けメモ

このリポジトリで作業するときの約束事。人間が読んでも困らない程度に短くしてある。

## 知識管理

ByteRover MCP は**廃止**。`brv` CLI を使う（ワークスペース共通規約に準拠）。

```bash
brv query "自然言語の質問"           # 取得（LLM あり）
brv search "キーワード" --limit 10   # 取得（BM25 のみ、コストなし）
brv curate "学んだこと: 判断・ファイルパス・理由" [-f path/to/file]
```

セットアップ: `npm i -g byterover-cli` → リポジトリルートで `brv status`。

- 新しいタスクに入る前、アーキテクチャを決める前、不慣れな箇所を触る前に `brv query`
- パターン・エラーの解法・設計判断を得たとき、まとまった作業を終えたときに `brv curate`

## このリポジトリ固有の注意

| 守ること | 理由 |
|---|---|
| GAS のソースは `src/` 配下のみが正。リポジトリ直下に `.gs` を置かない | 直下と `src/` の両方が push されると、GAS は全ファイルを単一のグローバルスコープに展開するため同名関数が二重定義され、後勝ちでサイレントに壊れる |
| 列を増減するときは `src/Utils.gs` の `getCompanySpec` / `getCustomerSpec` を直す | ヘッダー生成・行の組み立て・読み取りはすべてこの定義から導出される |
| シートから読んだ値は `normalizeCell` を通す | 会社側と顧客側で正規化が食い違うと、会社ID による結合がエラーにならないまま静かに外れる |
| UI や新しい Google サービスを使い始めたら `src/appsscript.json` の `oauthScopes` を更新する | `oauthScopes` を明示すると自動スコープ検出が上書きされる |
| `oauthScopes` を増やしたら `terraform/variables.tf` の `enabled_services` も見直す | GCP プロジェクト側で対応する API が無効だと、同意画面を通っても実行時に権限エラーで落ちる |
| GCP の設定変更はコンソールで直接せず `terraform/` に書いて `apply` する | コンソールで直接変えると次の `terraform apply` で巻き戻る。Terraform で扱えない OAuth 同意画面と Apps Script の紐付けだけが例外で、手順は `terraform/README.md` にある |
| 採番を伴う書き込みは `withDocumentLock` で囲む | 「最大値を読む → 追加」の read-modify-write なので、同時実行で ID が重複する |
| 変更したら `npm test` | `tests/local-harness.mjs` が GAS API をスタブしてロジックを検証する |
| `clasp push` の前に `npx clasp status` | 送信されるファイル一覧を確認する |

詳細は [gas-crm-implementation-guide.md](gas-crm-implementation-guide.md) を参照。
