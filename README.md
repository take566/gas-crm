# GAS CRM System

Google Apps Script（GAS）を使用したシンプルなCRMシステムです。

## 概要

Google スプレッドシートと Google Apps Script を使用して、会社情報と顧客情報を管理するシステムです。

### 機能

- 会社情報の管理（会社ID自動採番: C001, C002...）
- 顧客情報の管理（顧客ID自動採番: P001, P002...）
- 会社と顧客の紐付け
- 顧客検索機能
- ID割り当て機能（欠番の補完）

## プロジェクト構成

```
crm/
├── src/                         # clasp の rootDir。この配下だけが GAS に push される
│   ├── Utils.gs                 # 設定・ヘルパー関数、ID自動採番
│   ├── Company.gs               # 会社関連の関数
│   ├── Customer.gs              # 顧客関連の関数
│   ├── Menu.gs                  # カスタムメニューとダイアログ表示
│   ├── Test.gs                  # テスト用関数
│   ├── appsscript.json          # GAS設定（タイムゾーン・OAuthスコープ）
│   └── html/                    # HTMLダイアログ
│       ├── AddCompanyDialog.html
│       ├── AddCustomerDialog.html
│       └── SearchDialog.html
├── docs/                        # ドキュメント（push 対象外）
│   ├── gas-crm-implementation-guide.md
│   ├── AGENTS.md
│   └── CLAUDE.md
├── .clasp.json                  # Clasp設定（rootDir: src）
├── .claspignore
└── package.json                 # Node.js設定
```

`clasp push` では `rootDir`（= `src`）の部分が取り除かれるため、Apps Script プロジェクト側のファイル名は次のようになります。

```
Utils / Company / Customer / Menu / Test
html/AddCompanyDialog / html/AddCustomerDialog / html/SearchDialog
```

`Menu.gs` の `HtmlService.createHtmlOutputFromFile('html/AddCompanyDialog')` はこの名前を参照しています。**GAS のソースは `src/` 配下のみが正**です。リポジトリ直下に `.gs` を置くと同名関数が二重に push され、GAS は全ファイルを単一のグローバルスコープに展開するため後勝ちでサイレントに壊れます。

送信されるファイル一覧は `npx clasp status` で確認できます。

## セットアップ

### 1. 必要なツール

- [clasp](https://github.com/google/clasp) - Google Apps Script のコマンドラインツール
- Node.js (オプション)

### 2. インストール

```bash
# claspのインストール
npm install -g @google/clasp

# 依存関係のインストール（オプション）
npm install
```

### 3. 認証とプロジェクトのリンク

```bash
# Googleアカウントでログイン
clasp login

# 既存のプロジェクトにリンク（.clasp.jsonにscriptIdが設定されている場合）
clasp push
```

## 使用方法

### コードのアップロード

```bash
# ローカルのコードをGASプロジェクトにアップロード
clasp push

# GASプロジェクトからコードをダウンロード
clasp pull
```

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

## 詳細ドキュメント

詳細な実装手順については、[gas-crm-implementation-guide.md](docs/gas-crm-implementation-guide.md)を参照してください。

## ライセンス

MIT

