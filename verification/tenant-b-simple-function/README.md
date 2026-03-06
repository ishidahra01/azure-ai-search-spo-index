# テナントB シンプル認証 Graph 検証 Function

このフォルダは、`docs/entra-graph-security-and-batch-runbook.md` の [5.3.4（パターンB: テナントB 側アプリ登録方式）](../../docs/entra-graph-security-and-batch-runbook.md)を検証するための Azure Functions (Python) です。

> **テナント定義**
> - **テナントA**: 取得元の SharePoint Online (SPO) が存在するテナント
> - **テナントB**: SPO のデータを取得するアプリの実行基盤があるテナント（このFunctionのデプロイ先）

> **⚠️ Managed Identity によるクロステナント認証について**
> テナントB の Managed Identity をそのまま使ってテナントA の Graph API を呼ぶ構成は**成立しません**。
> テナントB の MI が取得するトークンはテナントB の Entra ID が発行するものであり、テナントA のリソースへのアクセス権を持ちません。
> このサンプルでは **`client_secret` 方式**（テナントA エンドポイント向け）を標準としています。

## 何をするか

- テナントB（実行基盤）から テナントA（SPO）の Graph API へ **Client Secret** でアクセス（クロステナント）
- `SP_SITE_PATHS` のサイト情報を取得
- 各サイトのドキュメント ライブラリ ルート配下のアイテムを取得
- ファイル URL、タイトル、更新日時、サイズなどのメタ情報を返却/ログ出力

## 構成

- HTTP Trigger: `GET /api/graph/tenant-scan`
- Timer Trigger: `%GRAPH_SCAN_SCHEDULE%` の cron で定期実行

---

## セットアップ手順

### 1. テナントB でアプリ登録

1. Azure Portal > Microsoft Entra ID（テナントB） > アプリの登録 > 新規登録
2. 設定:

| 項目 | 設定値 |
|------|--------|
| 名前 | 任意（例: `sp-graph-ingest-cross-tenant-prod`） |
| サポートされるアカウント種類 | **任意の組織ディレクトリ内のアカウント（マルチテナント）** |
| リダイレクト URI | 不要（App-only） |

3. 登録後、以下を記録:

| 項目 | 用途 |
|------|------|
| Application (client) ID | `GRAPH_CLIENT_ID` で使用 |
| Directory (tenant) ID | テナントB の ID（参照用） |

### 2. テナントB でアプリ設定（Graph API 権限 + Client Secret）

1. API のアクセス許可 > アクセス許可の追加
2. Microsoft Graph > Application permissions > `Sites.Selected` を追加
3. 「管理者の同意を与えます」は**ここでは実行しない**
   - このアプリはテナントA のリソースにアクセスするため、テナントA 側の管理者が同意を行います（手順3）
4. 証明書とシークレット > 新しいクライアント シークレット > 追加
   - 生成されたシークレット値を記録（`GRAPH_CLIENT_SECRET` で使用）
   - **本番環境では Key Vault 参照を使用すること**

### 3. テナントA で Admin Consent 実行

テナントA の全体管理者またはアプリケーション管理者が以下のいずれかで同意を実施:

**方法A: 同意 URL を使用（推奨）**

```
https://login.microsoftonline.com/{テナントA-ID}/adminconsent?client_id={Client-ID}
```

- `{テナントA-ID}`: テナントA のテナント ID
- `{Client-ID}`: 手順1で取得したアプリの Client ID
- 同意画面で「この組織の代理として同意する」をチェックして承認
- 成功するとテナントA に **エンタープライズアプリケーション** が作成される

**方法B: テナントA の Entra ID ポータルから同意**

1. テナントA のポータルで Microsoft Entra ID > エンタープライズアプリケーション
2. 手順1で作成したアプリを検索（Client ID で検索）
3. アクセス許可 > 「管理者の同意を与えます」を実行

### 4. テナントA で Sites.Selected 割当

`Sites.Selected` は付与するだけではアクセス不可です。対象サイトへの明示的な権限割当が必要です。

```powershell
# PowerShell + Microsoft.Graph モジュール
Connect-MgGraph -TenantId {テナントA-ID} -Scopes "Sites.FullControl.All"

# サイトID取得
$siteUrl = "https://{tenant-a}.sharepoint.com/sites/{sitename}"
$site = Get-MgSite -Search $siteUrl

# アプリに read 権限を付与
$params = @{
    roles = @("read")
    grantedToIdentities = @(
        @{
            application = @{
                id = "{Client-ID}"          # テナントB のアプリの Client ID
                displayName = "sp-graph-ingest-cross-tenant-prod"
            }
        }
    )
}
New-MgSitePermission -SiteId $site.Id -BodyParameter $params
```

> 複数サイトへの一括割当スクリプトは `../grant-sites-selected.ps1` を参照してください。

### 5. テナントB の Function App 設定

アプリケーション設定（環境変数）:

```properties
GRAPH_AUTH_MODE=client_secret        # クロステナントでは client_secret を使用
GRAPH_TENANT_ID={テナントA-ID}       # Graph データが存在するテナント（テナントA）
GRAPH_CLIENT_ID={Client-ID}          # テナントB に登録したアプリの Client ID
GRAPH_CLIENT_SECRET={Client-Secret}  # 本番環境は Key Vault 参照推奨
SP_HOSTNAME={tenant-a}.sharepoint.com
SP_SITE_PATHS=/sites/hr,/sites/legal
GRAPH_SCAN_SCHEDULE=0 */30 * * * *
```

> **本番環境での Secret 管理**: `GRAPH_CLIENT_SECRET` は Key Vault 参照を使用してください。
> 例: `@Microsoft.KeyVault(SecretUri=https://<vault-name>.vault.azure.net/secrets/<secret-name>)`

---

## ローカル実行

```powershell
cd verification/tenant-b-simple-function
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item local.settings.sample.json local.settings.json
# local.settings.json を編集して実際の値を設定
func start
```

`local.settings.json` の設定例:

```json
{
  "Values": {
    "GRAPH_AUTH_MODE": "client_secret",
    "GRAPH_TENANT_ID": "<テナントA-ID>",
    "GRAPH_CLIENT_ID": "<テナントB-アプリ-client-id>",
    "GRAPH_CLIENT_SECRET": "<テナントB-アプリ-client-secret>",
    "SP_HOSTNAME": "<tenant-a>.sharepoint.com",
    "SP_SITE_PATHS": "/sites/hr,/sites/legal",
    "GRAPH_SCAN_SCHEDULE": "0 */30 * * * *"
  }
}
```

HTTP テスト:

```powershell
curl "http://localhost:7071/api/graph/tenant-scan"
```

期待結果:
- 200 応答で、`siteId` / `siteTitle` / `items[].url` / `items[].title` / `items[].lastModifiedDateTime` が返る

負のテスト（403 確認）:
- `SP_SITE_PATHS` に未割当サイト（例: `/sites/not-granted`）を追加して再実行
- Graph 呼び出しが 403 になることを確認

## Azure へのデプロイ（テナントB）

```powershell
cd verification/tenant-b-simple-function
func azure functionapp publish <function-app-name>
```

---

## アプリ設定（Function App）一覧

| 変数名 | 必須 | 説明 |
|--------|------|------|
| `GRAPH_AUTH_MODE` | ○ | `client_secret` を設定（クロステナントでは `managed_identity` 不可） |
| `GRAPH_TENANT_ID` | ○ | テナントA の ID |
| `GRAPH_CLIENT_ID` | ○ | テナントB に登録したアプリの Client ID |
| `GRAPH_CLIENT_SECRET` | ○ | テナントB のアプリのシークレット（本番は Key Vault 参照推奨） |
| `SP_HOSTNAME` | ○ | テナントA の SharePoint ホスト名 |
| `SP_SITE_PATHS` | ○ | カンマ区切りのサイトパス |
| `GRAPH_SCAN_SCHEDULE` | ○ | Timer Trigger の cron 式 |

`managed_identity` モードについて:

- このモードはコード上サポートされていますが、**クロステナントシナリオ（テナントB → テナントA）では動作しません**
- テナントB の MI トークンはテナントB 発行であり、テナントA の Graph リソースへのアクセス権がないため認証が失敗します
- 同一テナント内での検証目的にのみ使用してください

---

## 注意

クロステナントアクセスでは `client_secret` 方式（`GRAPH_AUTH_MODE=client_secret`）を使用してください。
`managed_identity` モードはテナントB の Entra ID が発行するトークンを使うため、テナントA の Graph API へのアクセスは成立しません。
本番環境での Secret 運用は Key Vault 参照を利用し、Secret の直接設定を避けることを推奨します。
例: `@Microsoft.KeyVault(SecretUri=https://<vault-name>.vault.azure.net/secrets/<secret-name>)`
より高度なシークレットレス構成（FIC/OIDC）については `../tenant-b-graph-function/` を参照してください。
