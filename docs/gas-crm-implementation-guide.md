# Google Apps Script CRM 実装手順書

Google スプレッドシートと Google Apps Script（GAS）を使ったシンプルな CRM の構築・運用手順。

### 機能

- 会社情報の管理（会社ID 自動採番: C001, C002...）
- 顧客情報の管理（顧客ID 自動採番: P001, P002...）
- 会社と顧客の紐付け
- 顧客検索
- 未採番行への ID 割り当て

> **コードはこの文書に転記しません。**
> 以前はソース全文をこの手順書にコピーしていましたが、実ファイルと乖離して
> どちらが正か分からなくなりました。正は常に `src/` 配下の実ファイルです。
> ここでは「どのファイルが何を担うか」と「なぜそうなっているか」だけを書きます。

---

## 1. スプレッドシートの作成

### 1.1 新規スプレッドシートを作成

1. [Google Drive](https://drive.google.com) にアクセス
2. 「新規」→「Google スプレッドシート」→「空白のスプレッドシート」
3. ファイル名を「CRM管理」などに変更

### 1.2 シートの準備

シートは手で作らなくて構いません。メニューの「シートを初期化」を実行すると、
定義どおりのヘッダー付きで自動作成されます（[4.2](#42-シートの初期化初回のみ)）。

作られるシートは次の 2 つです。

**「会社」シート**

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| 会社ID | 会社名 | 住所 | 電話番号 | 備考 | 作成日時 |

**「顧客」シート**

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| 顧客ID | 名前 | 会社ID | メールアドレス | 電話番号 | 備考 | 作成日時 |

この列構成の**正は `src/Utils.gs` の `getCompanySpec()` / `getCustomerSpec()`** です。
上の表はその写しなので、列を変えるときは定義のほうを直してください（[6.2](#62-フィールドの追加)）。

---

## 2. コードの配置

### 2.1 clasp で push する（推奨）

```bash
npm install                 # @google/clasp を取得
npx clasp login
npx clasp status            # 送信されるファイルを確認
npx clasp push
```

`.clasp.json` の `rootDir` は `src` です。`src/` の外（`docs/`, `tests/`, `package.json` など）は
push されません。`clasp push` では `rootDir` の部分が取り除かれるため、Apps Script
プロジェクト側のファイル名は次のようになります。

```
Utils / Company / Customer / Menu / Test
html/AddCompanyDialog / html/AddCustomerDialog / html/SearchDialog
appsscript.json
```

> ⚠️ **リポジトリ直下に `.gs` を置かないでください。**
> GAS はプロジェクト内の全 `.gs` を**単一のグローバルスコープ**に展開します。
> 直下と `src/` の両方が push されると同名関数が二重定義され、どちらが有効かは
> ファイルの評価順に依存します。中身が同じうちは動いてしまうので、片方だけを
> 編集した瞬間にサイレントに壊れます。

`.clasp.json` は `.gitignore` されています（`scriptId` は環境ごとに違うため）。
clone 直後のセットアップは [README](../README.md#3-認証とプロジェクトのリンク) を参照。

### 2.2 各ファイルの役割

| ファイル | 役割 | 主な関数 |
|---|---|---|
| [`src/Utils.gs`](../src/Utils.gs) | シート定義、正規化、シートアクセス、排他制御、採番 | `getCompanySpec` / `getCustomerSpec` / `normalizeCell` / `readRows` / `buildRow` / `withDocumentLock` / `getNextId` / `assignMissingIds` |
| [`src/Company.gs`](../src/Company.gs) | 会社の追加・取得 | `addCompany` / `getCompanies` / `getCompanyById` / `getCompanyIdByName` / `assignMissingCompanyIds` |
| [`src/Customer.gs`](../src/Customer.gs) | 顧客の追加・取得・検索 | `addCustomer` / `getCustomers` / `searchCustomers` / `getCustomersByCompanyId` / `assignMissingCustomerIds` |
| [`src/Gmail.gs`](../src/Gmail.gs) | Gmail 送信履歴からの顧客候補抽出・取り込み | `getGmailImportCandidates` / `importSelectedGmailContacts` / `guessCompanyFromEmail` / `parseRecipientList` |
| [`src/Menu.gs`](../src/Menu.gs) | カスタムメニューとダイアログ表示 | `onOpen` / `initializeSheets` / `show*Dialog` / `formatAssignResult` |
| [`src/Test.gs`](../src/Test.gs) | 手動スモークテスト（**実シートに書き込む**） | `DEV_testAddCompany` / `DEV_testAddCustomer` |
| [`src/html/`](../src/html) | 入力・検索ダイアログ | — |

### 2.3 設計上の約束

コードを読む前に把握しておくと迷いません。

**シート定義が単一の正**

列・ヘッダー・ID プレフィックス・必須キーは `getCompanySpec()` / `getCustomerSpec()` に
まとまっています。ヘッダー生成も `appendRow` 用の行組み立ても読み取りも、すべて
この定義から導出されるので、列インデックス（`row[2]` のような添字）をコードに
散らさないでください。

**読み書きは必ず `normalizeCell` を通す**

スプレッドシートのセルは数値・日付・真偽値にもなります。会社側と顧客側で正規化が
食い違うと、会社ID による突き合わせ（厳密等価）が**エラーにならないまま静かに外れ**、
`companyName` が空になります。`getCompanies` / `getCustomers` は例外時に `[]` を返す
設計なので、この破綻はログにもユーザーにも出ません。だから入口で揃えます。

**採番は `withDocumentLock` で囲む**

採番は「既存 ID の最大値を読む → 行を追加する」という read-modify-write です。
囲まないと、同じスプレッドシートを 2 人が同時に操作したときに双方が同じ ID を得ます。

**採番は ID 列の全行を見る**

`readRows` は必須キーが空の行を落とします。会社の必須キーは `id` と `name` なので、
「会社名が空の C005」は `getCompanies()` から消えます。これを採番に使うと次が `C002`
になり既存の `C005` と衝突するため、採番だけは `readIdColumn()` で ID 列の全行を見ます。

**Gmail の「会社名候補」は会社名ではなくドメイン**

メールヘッダーには会社名フィールドが存在しません。`guessCompanyFromEmail` が返すのは
送信先メールアドレスのドメイン（`example.co.jp` など）で、実際の会社名とは限りません。
`FREE_EMAIL_DOMAINS`（Gmail/Yahoo!等）はそもそも候補を出しません。取り込みダイアログで
必ずユーザーに確認・編集させ、`addCompany` をそのまま呼ぶことで既存の正規化・採番・
ロックのロジックに乗せています（Gmail 用に別ロジックを作らない）。

### 2.4 必要な OAuth スコープ

`src/appsscript.json` でスコープを**明示宣言**しています。

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.container.ui",
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/userinfo.email"
]
```

`oauthScopes` を書くと Apps Script の自動スコープ検出は**上書きされ、宣言したものだけ**が
承認時に要求されます。したがってコードが使う API とこの一覧は手で同期させる必要が
あります。対応は次のとおり。

| スコープ | 必要とする API 呼び出し | 該当箇所 |
|---|---|---|
| `spreadsheets.currentonly` | `SpreadsheetApp.getActiveSpreadsheet()` / `getSheetByName` / `insertSheet` / `appendRow` / `getDataRange` / `getRange().setValue()` | `Utils.gs`, `Company.gs`, `Customer.gs` |
| `script.container.ui` | `SpreadsheetApp.getUi()`、`ui.createMenu().addToUi()`、`ui.alert()`、`ui.showModalDialog()`、`HtmlService.createHtmlOutputFromFile()` | `Menu.gs` 全体 |
| `gmail.readonly` | `GmailApp.search()` / `GmailApp.getMessagesForThreads()` | `Gmail.gs` |
| `userinfo.email` | `Session.getActiveUser().getEmail()` | `Gmail.gs`（自分が送信したメッセージかどうかの判定に使用） |

`script.container.ui` はコンテナバインドスクリプトの UI（カスタムメニュー・モーダル
ダイアログ・サイドバー）に必要です。これが無いとメニュー表示やダイアログ起動が
認可エラーになります。`gmail.readonly` は Gmail 送信履歴を**読み取るだけ**の最小権限で、
送信・削除・ラベル変更はできません（そもそも本機能はそれらを行いません）。

`Session.getActiveUser()` は一見スコープ不要に見えますが、実際には呼び出しに
`userinfo.email` を要求されます（実機検証で発覚。`Session.getScriptTimeZone()` など
他の `Session` メソッドは追加スコープ無しで動くため紛らわしい点です）。

> **UI を伴う機能を追加したとき**（サイドバー、`ui.prompt` など）や、**新しい Google
> サービスを使い始めたとき**（`MailApp`、`DriveApp` など）は、この表と `oauthScopes` を
> 必ず更新してください。既に承認済みのアカウントでは古いトークンが残っていて
> 気付けないことがあるため、**別アカウントで承認し直して確認**するのが確実です。

---

## 3. HTMLダイアログ

`src/html/` に 3 つあります。サーバ側関数は `google.script.run` 経由で呼びます。

| ファイル | 呼ぶサーバ関数 |
|---|---|
| [`AddCompanyDialog.html`](../src/html/AddCompanyDialog.html) | `addCompany(name, address, phone, note)` |
| [`AddCustomerDialog.html`](../src/html/AddCustomerDialog.html) | `getCompanies()`（会社プルダウン）、`addCustomer(name, companyId, email, phone, note)` |
| [`SearchDialog.html`](../src/html/SearchDialog.html) | `searchCustomers(keyword)` |
| [`ImportFromGmailDialog.html`](../src/html/ImportFromGmailDialog.html) | `getGmailImportCandidates(days, maxThreads)`、`importSelectedGmailContacts(rows)` |

`Menu.gs` からは `HtmlService.createHtmlOutputFromFile('html/AddCompanyDialog')` の
ように、**push 後のファイル名**（`src/` が取り除かれた名前）で参照します。

### ダイアログを追加・変更するときの約束

- **`withSuccessHandler` と `withFailureHandler` を必ず両方付ける。**
  failureHandler が無いとサーバ側の例外時にハンドラが一切呼ばれず、ダイアログが
  無反応になります。ユーザーには「ボタンが効かない」としか見えません
- **サーバから受け取った値を `innerHTML` に文字列連結しない。**
  スプレッドシートは誰でも直接編集できるので、備考や会社名に `<` や `&` が入ると
  表示が崩れます。`textContent` と DOM API で組み立ててください
- サーバ側の関数は「ログに残してからクライアントへ投げ返す」方針で揃えています。
  `[]` を返して握り潰すと「0 件」と「失敗」が区別できなくなります

---

## 4. 動作確認

### 4.1 カスタムメニューの表示

1. スプレッドシートに戻る（または再読み込み）
2. メニューバーに「🗂️ CRM」が表示されることを確認

> 初回は表示されるまで数秒かかる場合があります

### 4.2 シートの初期化（初回のみ）

1. 「🗂️ CRM」→「シートを初期化」
2. 初回実行時は承認が必要
   - 「続行」→ Google アカウントを選択
   - 「詳細」→「CRM（安全でないページ）に移動」
   - 「許可」をクリック
3. 「シートの初期化が完了しました。」と表示されることを確認
4. 「会社」「顧客」シートがヘッダー付きで作成されることを確認

### 4.3 会社の追加

1. 「🗂️ CRM」→「会社を追加」
2. フォームに入力して「追加」
3. 「会社」シートに行が追加され、会社ID が採番されることを確認

### 4.4 顧客の追加

1. 「🗂️ CRM」→「顧客を追加」
2. 会社プルダウンに 4.3 で追加した会社が出ることを確認
3. 会社ID を直接入力するか、プルダウンから選択
4. 入力して「追加」

> プルダウンは**ダイアログを開いた時点で 1 度だけ**読み込みます。開いたまま別途
> 会社を追加した場合は、いったん閉じて開き直してください。

### 4.5 検索

1. 「🗂️ CRM」→「顧客を検索」
2. 名前・会社名・メールの部分一致で検索できることを確認
3. 会社名の列が埋まっていることを確認（空なら会社ID の突き合わせが外れています）

### 4.6 Gmail送信履歴からの取り込み

1. 「🗂️ CRM」→「Gmail送信履歴から取り込み」
2. 初回は Gmail 読み取りの追加承認が必要（[2.4](#24-必要な-oauth-スコープ)）
3. 直近30日の送信履歴から、未登録の宛先が一覧表示されることを確認
4. 会社名候補（ドメイン）が編集可能なこと、フリーメール宛には候補が出ないことを確認
5. チェックを付けて「選択した行を登録」→ 会社・顧客が追加されることを確認
6. 再度開くと、登録済みの相手が候補から消えていることを確認

### 4.7 ロジックのローカル検証

シートに触らずに列マッピング・正規化・結合・採番を確認できます。

```bash
npm test        # tests/local-harness.mjs
```

GAS API をスタブして `src/*.gs` を Node で実行します。**実 GAS 環境の代替では
ありません**（UI・権限・同時実行は再現しません）。`tests/` は `rootDir` の外なので
push 対象にもなりません。

`src/Test.gs` の `DEV_*` 関数は**実シートに書き込む**スモークテストです。GAS エディタの
関数プルダウンから誤実行しないよう `DEV_` を付けてあります。

---

## 5. ファイル構成

**リポジトリ**

```
crm/
├── src/                         # clasp の rootDir。この配下だけが push される
│   ├── Utils.gs
│   ├── Company.gs
│   ├── Customer.gs
│   ├── Menu.gs
│   ├── Test.gs
│   ├── Gmail.gs
│   ├── appsscript.json
│   └── html/
│       ├── AddCompanyDialog.html
│       ├── AddCustomerDialog.html
│       ├── SearchDialog.html
│       └── ImportFromGmailDialog.html
├── tests/
│   └── local-harness.mjs        # npm test
├── docs/
│   ├── gas-crm-implementation-guide.md
│   ├── AGENTS.md
│   └── CLAUDE.md
├── .clasp.json                  # rootDir: src（gitignore 対象）
├── .claspignore
└── package.json
```

**Apps Script プロジェクト（push 後）**

```
CRM
├── appsscript.json
├── Utils.gs
├── Company.gs
├── Customer.gs
├── Menu.gs
├── Test.gs
├── Gmail.gs
└── html/
    ├── AddCompanyDialog.html
    ├── AddCustomerDialog.html
    ├── SearchDialog.html
    └── ImportFromGmailDialog.html
```

---

## 6. カスタマイズ

### 6.1 ID 形式の変更

`src/Utils.gs` のシート定義の `idDigits` / `idPrefix` を変えるだけです。

```js
idPrefix: 'C',
idDigits: 4,     // → C0001 形式
```

採番（`getNextId`）も割り当て（`assignMissingIds`）も同じ定義を見るので、直す箇所は
1 つです。

### 6.2 フィールドの追加

1. `src/Utils.gs` の `getCompanySpec()` / `getCustomerSpec()` の `columns` に
   `{ key: 'xxx', header: '表示名' }` を追加する
2. `addCompany` / `addCustomer` の引数に追加し、`buildRow` に渡すオブジェクトへ足す
3. HTML フォームに入力欄を追加する

ヘッダー・列順・読み取りは定義から導出されるので、既存シートには**列を追加するだけ**で
順序を合わせる必要はありません（定義の順序どおりに追記してください）。

---

## トラブルシューティング

| 問題 | 対処法 |
|------|--------|
| メニューが表示されない | スプレッドシートを再読み込み。出なければ `onOpen` を手動実行 |
| 承認画面が出ない | 一度 `onOpen` 関数を手動実行 |
| メニューやダイアログで認可エラー | `src/appsscript.json` の `oauthScopes` に `script.container.ui` があるか確認（[2.4](#24-必要な-oauth-スコープ)） |
| シートが見つからないエラー | 「🗂️ CRM」→「シートを初期化」を実行 |
| ID が正しく採番されない | ヘッダー行が 1 行目にあるか確認 |
| ID が空欄の行がある | 「🗂️ CRM」→「会社IDを割り当て」または「顧客IDを割り当て」。**空欄の行にのみ**採番します |
| ID 列に想定外の値が入っている | 割り当て実行時に「上書きしていません」と件数が出ます。手入力値を消さないための仕様です。行番号は実行ログ（表示 → 実行数）に記録されます |
| 検索結果の「会社」列が空 | 顧客シートの会社ID が会社シートに存在するか確認。`npm test` が通るなら前後空白や数値セルは吸収されています |
| 検索しても何も起きない | 現在はエラーがダイアログに表示されます。表示が出ない場合は実行ログを確認 |
| HTML ファイルが見つからない | `npx clasp status` で `src/html/*.html` が送信対象か確認。`Menu.gs` のパスが `html/AddCompanyDialog`（`src/` 無し）になっているか確認 |
| ID が重複した | 同時操作で発生した可能性。現在は `withDocumentLock` で防いでいます |
| Gmail取り込みで認可エラー | `src/appsscript.json` の `oauthScopes` に `gmail.readonly` と `userinfo.email` があるか確認（[2.4](#24-必要な-oauth-スコープ)） |
| Gmail取り込みが `Session.getActiveUser を呼び出すことができません` で失敗 | `userinfo.email` スコープが不足しています。追加後は再 push し、初回のみ再承認が必要です |
| Gmail取り込みの候補が0件 | 直近30日に送信履歴が無いか、宛先が全員すでに顧客登録済みの可能性があります |
| Gmail取り込みの会社名がおかしい | 候補は送信先メールアドレスの**ドメイン**であって会社名ではありません。登録前に必ず編集してください |

### ID の「欠番」について

`assignMissing*Ids` は**欠番を埋めません**。既存 ID の最大値の次から連番を振ります。

C001 と C003 がある状態で空欄行に割り当てると、C002 ではなく **C004** が振られます。
欠番を再利用すると、削除済みの ID を指している既存データ（顧客シートの会社ID など）が
別の会社を指してしまうためです。

---

## 次のステップ（拡張案）

- 編集・削除機能の追加
- 対応履歴シートの追加
- メール送信機能（`MailApp` のスコープ追加が必要）
- CSV 出力機能
- Web アプリとして公開
