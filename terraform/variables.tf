variable "project_id" {
  description = "CRM 用 GCP プロジェクト ID。全 GCP で一意・6〜30 文字・小文字英数字とハイフン。作成後は変更できない"
  type        = string

  validation {
    condition     = can(regex("^[a-z][a-z0-9-]{4,28}[a-z0-9]$", var.project_id))
    error_message = "project_id は小文字英字で始まり、小文字英数字とハイフンのみ、6〜30 文字である必要があります。"
  }
}

variable "project_name" {
  description = "GCP コンソールに表示されるプロジェクト名"
  type        = string
  default     = "GAS CRM"
}

variable "billing_account" {
  description = "紐付ける請求先アカウント ID（例: 0X0X0X-0X0X0X-0X0X0X）。空の場合は請求先を紐付けない。本構成の API はいずれも無料枠で足りるため既定は空"
  type        = string
  default     = ""
}

variable "org_id" {
  description = "所属する組織 ID。個人アカウント（組織なし）では空のままにする。folder_id と同時に指定はできない"
  type        = string
  default     = ""
}

variable "folder_id" {
  description = "所属するフォルダ ID。組織配下に置く場合のみ指定する。org_id と同時に指定はできない"
  type        = string
  default     = ""
}

variable "region" {
  description = "既定リージョン。本構成ではリージョン依存リソースを作らないが、provider の既定として設定する"
  type        = string
  default     = "asia-northeast1"
}

variable "enabled_services" {
  description = <<-EOT
    有効化する GCP API。
    既定値は src/appsscript.json の oauthScopes（spreadsheets.currentonly / script.container.ui /
    gmail.readonly）と exceptionLogging: STACKDRIVER に対応する最小構成。
    Drive API などを使い始めたらここに追加する。
  EOT
  type        = list(string)
  default = [
    # Terraform 自身がプロジェクト・API を操作するために必要
    "cloudresourcemanager.googleapis.com",
    "serviceusage.googleapis.com",

    # Apps Script 本体（clasp / Apps Script API）
    "script.googleapis.com",

    # appsscript.json の oauthScopes に対応
    "sheets.googleapis.com",
    "gmail.googleapis.com",

    # exceptionLogging: STACKDRIVER の出力先
    "logging.googleapis.com",
    "clouderrorreporting.googleapis.com",
  ]
}

variable "log_viewer_members" {
  description = "Cloud Logging の閲覧権限（roles/logging.viewer）を付与するメンバー。IAM 形式で指定する（例: [\"user:tmf566@gmail.com\"]）。空の場合は付与しない（プロジェクトオーナーは既定で閲覧できる）"
  type        = list(string)
  default     = []

  validation {
    condition     = alltrue([for m in var.log_viewer_members : can(regex("^(user|group|serviceAccount|domain):", m))])
    error_message = "log_viewer_members は user: / group: / serviceAccount: / domain: のいずれかの接頭辞付きで指定してください。"
  }
}

variable "log_retention_days" {
  description = "Cloud Logging の _Default バケットの保持日数。無料枠は 30 日。伸ばすと課金対象になるため、請求先アカウント未設定なら 30 のままにする"
  type        = number
  default     = 30

  validation {
    condition     = var.log_retention_days >= 1 && var.log_retention_days <= 3650
    error_message = "log_retention_days は 1〜3650 の範囲で指定してください。"
  }
}

variable "deletion_policy" {
  description = "terraform destroy 時のプロジェクトの扱い。PREVENT=削除を拒否（既定）、DELETE=削除する、ABANDON=state から外すだけ。CRM の実データを載せるプロジェクトのため既定は PREVENT"
  type        = string
  default     = "PREVENT"

  validation {
    condition     = contains(["PREVENT", "DELETE", "ABANDON"], var.deletion_policy)
    error_message = "deletion_policy は PREVENT / DELETE / ABANDON のいずれかです。"
  }
}

variable "labels" {
  description = "プロジェクトに付ける追加ラベル。managed-by / app は本構成が常に付与する"
  type        = map(string)
  default     = {}
}
