# ====== 設定値 ======
# このスクリプトは テナントA（SPO 側）で実行し、テナントB のアプリに Sites.Selected 権限を付与します
$Tenant  = "<tenant-a-id>"       # テナントA のテナント ID（SPO が存在するテナント）
$ClientId = "<tenant-b-app-client-id>"  # テナントB に登録したアプリの Client ID
$spHost   = "<tenant-a>.sharepoint.com" # テナントA の SharePoint ホスト名
$sitePath = "/sites/<your-site>"

# ====== モジュール確認 ======
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# ====== Tenant に接続 ======
Connect-MgGraph -TenantId $Tenant -Scopes "Sites.FullControl.All"

# ====== サイト取得 ======
$site = Get-MgSite -SiteId "${spHost}:${sitePath}"

Write-Host "Site ID:" $site.Id

# ====== 権限付与 ======
$params = @{
  roles = @("read")
  grantedToIdentities = @(
    @{
      application = @{
        id = $ClientId
        displayName = "sp-graph-ingest-cross-tenant-prod"
      }
    }
  )
}

New-MgSitePermission -SiteId $site.Id -BodyParameter $params

Write-Host "Permission granted successfully."