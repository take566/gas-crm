provider "google" {
  region = var.region
}

locals {
  # org_id / folder_id は google_project 側で同時指定できないため、空文字は null に落とす
  org_id    = var.org_id != "" ? var.org_id : null
  folder_id = var.folder_id != "" ? var.folder_id : null

  labels = merge(
    {
      managed-by = "terraform"
      app        = "gas-crm"
    },
    var.labels
  )
}

# CRM 用の標準 GCP プロジェクト。
# Apps Script が自動生成する非表示プロジェクトの代わりにこれを紐付けることで、
# 有効な API・ログの保持期間・権限をコードで管理できるようにする。
resource "google_project" "crm" {
  name       = var.project_name
  project_id = var.project_id

  org_id    = local.org_id
  folder_id = local.folder_id

  # 空の場合は請求先を紐付けない（本構成の API は無料枠で足りる）
  billing_account = var.billing_account != "" ? var.billing_account : null

  # 既定は PREVENT。CRM の実データに紐づくプロジェクトを destroy で消さないため
  deletion_policy = var.deletion_policy

  labels = local.labels

  # プロジェクト作成 API 自体を叩くための API。既定プロジェクトで有効化済みである前提
  auto_create_network = false
}

# 有効化する API。
# disable_on_destroy = false: destroy 時に API を無効化すると、同じプロジェクトを
# 使う他のリソース（Apps Script 実行時のログ出力など）が巻き添えで止まるため。
resource "google_project_service" "enabled" {
  for_each = toset(var.enabled_services)

  project = google_project.crm.project_id
  service = each.value

  disable_on_destroy         = false
  disable_dependent_services = false
}

# Cloud Logging の _Default バケット。
# プロジェクト作成時に自動生成されるため、ここでは保持期間の更新だけを行う。
resource "google_logging_project_bucket_config" "default" {
  project        = google_project.crm.project_id
  location       = "global"
  bucket_id      = "_Default"
  retention_days = var.log_retention_days

  depends_on = [google_project_service.enabled]
}

# GAS の例外ログ（exceptionLogging: STACKDRIVER）を読むための権限。
# プロジェクトオーナーは既定で閲覧できるため、オーナー以外に見せる場合のみ指定する。
resource "google_project_iam_member" "log_viewer" {
  for_each = toset(var.log_viewer_members)

  project = google_project.crm.project_id
  role    = "roles/logging.viewer"
  member  = each.value

  depends_on = [google_project_service.enabled]
}
