# GAS CRM System

Google Apps Script（GAS）を使用したシンプルなCRMシステムです。

## 概要

Google スプレッドシートと Google Apps Script を使用して、会社情報と顧客情報を管理するシステムです。

### 機能

- 会社情報の管理（会社ID自動採番: C001, C002...）
- 顧客情報の管理（顧客ID自動採番: P001, P002...）
- 会社と顧客の紐付け
- 顧客検索機能
- 未採番行への ID 割り当て（既存 ID の最大値の次から連番。**欠番は埋めません**）
- Gmail送信履歴からの顧客候補取り込み（会社名は送信先ドメインからの推測。要確認・編集）

## プロジェクト構成

```
crm/
├── src/                         # clasp の rootDir。この配下だけが GAS に push される
│   ├── Utils.gs                 # 設定・ヘルパー関数、ID自動採番
│   ├── Company.gs               # 会社関連の関数
│   ├── Customer.gs              # 顧客関連の関数
│   ├── Gmail.gs                 # Gmail送信履歴からの顧客候補取り込み
│   ├── Menu.gs                  # カスタムメニューとダイアログ表示
│   ├── Test.gs                  # テスト用関数
│   ├── appsscript.json          # GAS設定（タイムゾーン・OAuthスコープ）
│   └── html/                    # HTMLダイアログ
│       ├── AddCompanyDialog.html
│       ├── AddCustomerDialog.html
│       └── SearchDialog.html
├── tests/                       # push 対象外
│   └── local-harness.mjs        # GAS API をスタブしたロジック検証（npm test）
├── docs/                        # ドキュメント（push 対象外）
│   ├── gas-crm-implementation-guide.md
│   ├── AGENTS.md
│   └── CLAUDE.md
├── terraform/                   # CRM 用 GCP プロジェクトの IaC（push 対象外）
│   ├── main.tf                  # プロジェクト・API 有効化・ログ保持・IAM
│   ├── variables.tf
│   ├── outputs.tf
│   └── README.md                # 適用手順と、Terraform で扱えない手動設定
├── .clasp.json                  # Clasp設定（rootDir: src、gitignore 対象）
├── .claspignore
└── package.json                 # Node.js設定
```

`clasp push` では `rootDir`（= `src`）の部分が取り除かれるため、Apps Script プロジェクト側のファイル名は次のようになります。

```
Utils / Company / Customer / Gmail / Menu / Test
html/AddCompanyDialog / html/AddCustomerDialog / html/SearchDialog
```

`Menu.gs` の `HtmlService.createHtmlOutputFromFile('html/AddCompanyDialog')` はこの名前を参照しています。**GAS のソースは `src/` 配下のみが正**です。リポジトリ直下に `.gs` を置くと同名関数が二重に push され、GAS は全ファイルを単一のグローバルスコープに展開するため後勝ちでサイレントに壊れます。

送信されるファイル一覧は `npx clasp status` で確認できます。

## セットアップ

### 1. 必要なツール

- Node.js
- [clasp](https://github.com/google/clasp) — `npm install` で devDependency として入ります

### 2. インストール

```bash
npm install
```

### 3. 認証とプロジェクトのリンク

```bash
npx clasp login
```

`.clasp.json` は `scriptId` が環境ごとに異なるため **git 管理外**です（`.gitignore`）。
clone 直後は存在しないので、次のどちらかで用意します。

**A. 既存の Apps Script プロジェクトに繋ぐ**

スクリプトエディタの URL `https://script.google.com/.../projects/<scriptId>/edit` から
`scriptId` を控えて、リポジトリ直下に `.clasp.json` を作成します。

```json
{
  "scriptId": "ここに scriptId",
  "rootDir": "src"
}
```

**B. 新規に作る**

```bash
# CRM 用スプレッドシートを開き、拡張機能 → Apps Script でプロジェクトを作ってから A の手順
# もしくはコンテナバインドで新規作成
npx clasp create --type sheets --rootDir src
```

`rootDir` は必ず `src` にしてください（`.` にすると直下のファイルまで push されます）。

### 4. 動作確認とアップロード

```bash
npm test            # シートに触らないローカル検証
npx clasp status    # 送信されるファイル一覧を確認
npx clasp push
```

## 使用方法

### コードのアップロード

```bash
npx clasp push      # ローカル → GAS
npx clasp pull      # GAS → ローカル
```

### テスト

```bash
npm test
```

`tests/local-harness.mjs` が GAS API をスタブして `src/*.gs` を Node で実行し、
列マッピング・セル値の正規化・会社ID による結合・採番を検証します。
実 GAS 環境の代替ではありません（UI・権限・同時実行は再現しません）。

### スプレッドシートでの操作

1. スプレッドシートを開く
2. メニューバーに「🗂️ CRM」が表示される
3. メニューから各種操作を実行：
   - 会社を追加
   - 顧客を追加
   - 顧客を検索
   - 会社IDを割り当て
   - 顧客IDを割り当て
   - シートを初期化

## GCP（Terraform）

Apps Script は既定で「非表示の GCP プロジェクト」に自動で紐づく。この状態では有効な API も
OAuth 同意画面もログの保持期間もコードに残らないため、CRM 専用の GCP プロジェクトを
`terraform/` で定義している。

```bash
cd terraform
cp terraform.example.tfvars terraform.tfvars   # project_id を一意な値に変える
terraform init
terraform plan
terraform apply
terraform output project_number                # Apps Script への紐付けに使う
```

OAuth 同意画面の設定と Apps Script への紐付けは Terraform では扱えないため手動で行う。
手順は [terraform/README.md](terraform/README.md) を参照。

## 詳細ドキュメント

詳細な実装手順については、[gas-crm-implementation-guide.md](docs/gas-crm-implementation-guide.md)を参照してください。

## ライセンス

MIT

