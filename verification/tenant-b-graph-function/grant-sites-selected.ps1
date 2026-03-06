# ============================================================
# テナントA 側: SharePoint サイト権限付与スクリプト
# ============================================================
# テナントB のマルチテナントアプリに対して、
# テナントA の SharePoint サイトへの Sites.Selected 権限を付与します。
#
# 前提:
#   - テナントA 管理者で実行
#   - Admin Consent が完了済み (Enterprise Application が作成済み)
#   - Microsoft.Graph モジュールがインストール済み
# ============================================================

# ====== 設定値 ======
# テナントA のテナント ID（SPO が存在するテナント）
$TenantA  = "<tenant-a-id>"

# テナントB で作成したマルチテナントアプリの Client ID
$ClientId = "<tenant-b-app-client-id>"

# テナントB アプリの表示名 (任意)
$AppDisplayName = "<app-display-name>"

# テナントA の SharePoint ホスト名・サイトパス
$spHost   = "<tenant-a>.sharepoint.com"
$sitePath = "/sites/<site-name>"

# 権限レベル: read / write
$permissionRole = "read"

# ====== モジュール確認 ======
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# ====== テナントA に接続 ======
Connect-MgGraph -TenantId $TenantA -Scopes "Sites.FullControl.All"

# ====== サイト取得 ======
$site = Get-MgSite -SiteId "${spHost}:${sitePath}"

Write-Host "Site ID:" $site.Id

# ====== 権限付与 ======
$params = @{
  roles = @($permissionRole)
  grantedToIdentities = @(
    @{
      application = @{
        id = $ClientId
        displayName = $AppDisplayName
      }
    }
  )
}

New-MgSitePermission -SiteId $site.Id -BodyParameter $params

Write-Host "Permission granted successfully."

# ====== 確認 ======
Write-Host ""
Write-Host "=== Current Site Permissions ==="
Get-MgSitePermission -SiteId $site.Id | Format-List
