# ====== 設定値 ======
$TenantB  = "<Tenant-B-ID>"
$ClientId = "<Client-ID>"
$spHost   = "<tenantB>.sharepoint.com"
$sitePath = "/sites/<sitename>"

# ====== モジュール確認 ======
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# ====== Tenant B に接続 ======
Connect-MgGraph -TenantId $TenantB -Scopes "Sites.FullControl.All"

# ====== サイト取得 ======
$site = Get-MgSite -SiteId "$spHost:$sitePath"

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