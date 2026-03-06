# ====== 設定値 ======
$Tenant  = "<your-tenant-id>"
$ClientId = "<your-client-id>"
$spHost   = "<your-tenant>.sharepoint.com"
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