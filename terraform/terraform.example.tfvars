# terraform.tfvars としてコピーして使う（*.tfvars は .gitignore 対象）。
#   cp terraform.example.tfvars terraform.tfvars

# 全 GCP で一意にする必要があるため、末尾に日付やランダム文字を足す
project_id   = "gas-crm-20260902"
project_name = "GAS CRM"

# 課金アカウントを持っていない場合は空のまま。本構成の API は無料枠で足りる
billing_account = ""

# 組織なしの個人アカウントでは両方空のまま
org_id    = ""
folder_id = ""

# オーナー以外に GAS の例外ログを見せる場合のみ指定する
log_viewer_members = []

# 無料枠は 30 日。伸ばすと課金対象
log_retention_days = 30

# CRM の実データに紐づくため、既定では destroy を拒否する
deletion_policy = "PREVENT"
