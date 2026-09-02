output "project_id" {
  description = "作成した GCP プロジェクトの ID"
  value       = google_project.crm.project_id
}

output "project_number" {
  description = "GCP プロジェクト番号。Apps Script エディタの「プロジェクトの設定 → Google Cloud Platform（GCP）プロジェクト」に入力する値"
  value       = google_project.crm.number
}

output "enabled_services" {
  description = "有効化した API の一覧"
  value       = sort([for s in google_project_service.enabled : s.service])
}

output "log_retention_days" {
  description = "Cloud Logging _Default バケットの保持日数"
  value       = google_logging_project_bucket_config.default.retention_days
}

output "oauth_consent_url" {
  description = "OAuth 同意画面の設定ページ（Terraform では扱えないため手動設定する。README 参照）"
  value       = "https://console.cloud.google.com/apis/credentials/consent?project=${google_project.crm.project_id}"
}

output "logs_url" {
  description = "GAS の例外ログを確認する Cloud Logging のページ"
  value       = "https://console.cloud.google.com/logs/query?project=${google_project.crm.project_id}"
}
