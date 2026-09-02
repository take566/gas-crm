# state はローカル保管。
# 個人利用かつリソース数が少ないため共有バックエンドを置かない。
# チームで共有する必要が出たら、state 用の GCS バケットを別途作ってから
# backend "gcs" に差し替える（tfstate は terraform init -migrate-state で移行できる）。
terraform {
  backend "local" {
    path = "terraform.tfstate"
  }
}
