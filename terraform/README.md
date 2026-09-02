# terraform — CRM 用 GCP プロジェクト

GAS CRM が使う **標準 GCP プロジェクト**を Terraform で定義する。

Apps Script は既定で「非表示の GCP プロジェクト」を自動生成してそこに紐づく。この状態だと
有効な API も OAuth 同意画面もログの保持期間もコードに残らないため、専用プロジェクトを立てて
Apps Script 側をそこに向ける。

## 管理するもの

| リソース | 内容 |
|---|---|
| `google_project` | CRM 専用の GCP プロジェクト |
| `google_project_service` | Apps Script / Sheets / Gmail / Logging / Error Reporting などの API 有効化 |
| `google_logging_project_bucket_config` | `_Default` ログバケットの保持日数 |
| `google_project_iam_member` | `roles/logging.viewer` の付与（任意） |

有効化する API の既定値は `../src/appsscript.json` の `oauthScopes` と
`exceptionLogging: "STACKDRIVER"` に対応する最小構成。スコープを増やしたら
`variables.tf` の `enabled_services` にも足す。

## 管理しないもの（Terraform で扱えない）

| 項目 | 理由 | 対処 |
|---|---|---|
| OAuth 同意画面（External） | Google プロバイダに対応リソースが無い（`google_iap_brand` は組織内部向け） | 後述の手動手順 |
| Apps Script ↔ GCP プロジェクトの紐付け | Apps Script エディタからの操作のみ | 後述の手動手順 |
| Apps Script のコード・デプロイ | `clasp push` の担当 | リポジトリ直下の README 参照 |

## 前提

- Terraform >= 1.5
- gcloud CLI で認証済み、かつ Application Default Credentials（ADC）が有効なこと

```bash
gcloud auth login tmf566@gmail.com
gcloud auth application-default login
```

ADC が生きているかは次で確認できる（トークンは表示せず終了コードだけ見る）。

```bash
gcloud auth application-default print-access-token > /dev/null && echo OK
```

## 適用手順

```bash
cd terraform
cp terraform.example.tfvars terraform.tfvars
```

`terraform.tfvars` の `project_id` を**全 GCP で一意な値**に変える（作成後は変更できない）。
`*.tfvars` は `.gitignore` 対象なのでコミットされない。

```bash
terraform init
terraform plan
terraform apply
```

適用後、Apps Script への紐付けに使うプロジェクト番号を控える。

```bash
terraform output project_number
```

## 適用後の手動手順

この 3 手順は 2026-09-02 に実施済み。以下は再構築するときの手順書であり、
実際に踏んだ落とし穴も残してある。

### 1. OAuth 同意画面

```bash
terraform output oauth_consent_url   # 設定ページの URL が出る
```

- User Type: **外部**（個人 Gmail アカウントには組織が無いため内部は選べない）
- 公開ステータス: **テスト**のまま。テストユーザーに自分のアカウントを追加する
- スコープに `../src/appsscript.json` の `oauthScopes` を登録する
  - `.../auth/spreadsheets.currentonly`
  - `.../auth/script.container.ui`
  - `.../auth/gmail.readonly`（機密スコープ。テスト公開のままなら審査は不要）

3 つとも「スコープを追加または削除」パネルの一覧には出ないことがある。その場合は
パネル下部の **「スコープの手動追加」** にカンマ区切りで貼り付けて［テーブルに追加］
→［更新］する。

設定を終えても［対象］ページに
**「アプリの OAuth 構成が完了していません」** の警告が残るが、これはアプリを一般公開
するときにだけ必要な任意項目（ホームページ / プライバシーポリシー / 利用規約の URL）が
空であることを指す。テストステータスのままなら実害はなく、［ブランディング］の
確認ステータスも「テストステータスであるため、検証は必要ありません」と表示される。

### 2. Apps Script プロジェクトを紐付ける

1. スプレッドシート → 拡張機能 → Apps Script でエディタを開く
2. 左メニュー「プロジェクトの設定」→「Google Cloud Platform（GCP）プロジェクト」→「プロジェクトを変更」
3. `terraform output project_number` の値を入力して設定

**この操作は取り消せない。** Google 側の警告のとおり、

- 旧（デフォルト）プロジェクトでのユーザー認証がすべて取り消される
- Apps Script 管理のデフォルトプロジェクトには**戻せない**

**スクリプトがゴミ箱にあると設定画面が操作できない。**「プロジェクトがゴミ箱にあります」
と出たら先に［ゴミ箱の外に移動する］で復元する。本リポジトリのスクリプトはコンテナ
バインドなので、復元するとコンテナのスプレッドシートも一緒にゴミ箱から戻る。

紐付け後、GCP の表示が「デフォルト」から「標準」に変わる。GAS の例外ログは次で確認できる。

```bash
terraform output logs_url
```

### 3. 動作確認

紐付けでスコープの同意がリセットされるため、スプレッドシートのメニューから
一度実行して承認し直す。**再承認するまでメニューは動かない。**

## 変更を反映する

`variables.tf` の既定値ではなく `terraform.tfvars` を編集して `terraform apply` する。

- API を追加する → `enabled_services` に足す
- 他の人にログを見せる → `log_viewer_members = ["user:someone@example.com"]`

## 注意

- **`deletion_policy` は既定で `PREVENT`。** CRM の実データに紐づくプロジェクトのため、
  `terraform destroy` は拒否される。本当に消すときだけ `terraform.tfvars` で
  `deletion_policy = "DELETE"` にしてから `apply` → `destroy` の順で実行する。
- **`auto_create_network` は `true`（既定）のまま。** `false` にするとプロバイダが default
  ネットワークを削除するために Compute Engine API を有効化しにいき、請求先アカウントが
  未設定のプロジェクトでは `Billing must be enabled for activation of service(s)
  'compute.googleapis.com'` で `apply` が失敗する。新規プロジェクトでは Compute Engine API
  自体が無効なため、`true` でも default VPC は作られない。
- **`_Default` ログバケットは plan 上「作成」と表示される。** 実際にはプロジェクト作成時に
  自動生成されたバケットを取得して更新する（プロバイダの acquire-or-create 動作）。
- **`log_retention_days` の無料枠は 30 日。** 伸ばすと課金対象になる。請求先アカウントを
  紐付けていない状態（`billing_account = ""`）では 30 のままにする。
- **state はローカル保管**（`backend "local"`）。`terraform.tfstate` は `.gitignore` 対象。
  共有が必要になったら GCS バケットを作って `backend "gcs"` に差し替え、
  `terraform init -migrate-state` で移行する。
