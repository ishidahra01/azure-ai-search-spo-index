# テナントB Graph Function — クロステナント Graph アクセス検証（Workload Identity Federation 方式）

このフォルダは、`docs/entra-graph-security-and-batch-runbook.md` の `5.3.3`（テナントA 側アプリ登録方式）を Workload Identity Federation で実装した検証用 Azure Functions です。

> **テナント定義**
> - **テナントA**: 取得元の SharePoint Online (SPO) が存在するテナント
> - **テナントB**: SPO のデータを取得するアプリの実行基盤があるテナント（このFunctionのデプロイ先）

## 構成概要

```
テナントB（実行基盤）
 ├ Azure Functions
 │   └ User Assigned Managed Identity (UAMI)
 │
 └ App Registration (Multitenant)
      └ Federated Credential → UAMI

              ↓ admin consent

テナントA（SPO）
 └ Enterprise Application
      └ Microsoft Graph: Sites.Selected
      └ SharePoint Site Permission
```

**特徴:**
- Client Secret 不要
- 証明書不要
- Managed Identity ベース認証 (Workload Identity Federation)

## 認証フロー

1. Azure Functions が UAMI で `api://AzureADTokenExchange` トークンを取得
2. そのトークンを Client Assertion として テナントA 向け Graph トークンを取得
3. Graph API で テナントA の SharePoint サイトへアクセス

---

## セットアップ手順

### 1. テナントB 側：Entra ID アプリ登録

1. Azure Portal → Microsoft Entra ID → アプリの登録 → 新規登録
2. 設定:

| 項目 | 設定値 |
|------|--------|
| 名前 | 任意（例：AppGalileo00001） |
| サポートされているアカウント | 任意の組織ディレクトリ (マルチテナント) |
| リダイレクト URI | 未設定 |

3. 作成後、以下を控える:

| 項目 | 用途 |
|------|------|
| アプリケーション (Client) ID | `APP_CLIENT_ID` で使用 |
| ディレクトリ (Tenant) ID | 認証時に使用 |

### 2. テナントB 側：User Assigned Managed Identity 作成

1. Azure Portal → Managed Identity → User Assigned を作成
2. 取得する情報:

| 項目 | 用途 |
|------|------|
| Client ID | `MANAGED_IDENTITY_CLIENT_ID` で使用 |
| Object ID | Federated Credential で使用 |

### 3. テナントB 側：Functions に UAMI を割り当て

1. Azure Functions → ID → ユーザー割り当て → 追加
2. 作成した UAMI を選択

### 4. テナントB 側：Federated Credential 作成

1. Entra ID → アプリの登録 → 作成したアプリ
2. 証明書とシークレット → Federated credentials → Add credential

| 項目 | 設定 |
|------|------|
| Credential scenario | Managed Identity |
| Subscription | 対象サブスクリプション |
| Managed Identity | 作成した UAMI |
| Name | 任意 |

自動生成される値:

| 項目 | 値 |
|------|------|
| Issuer | `https://login.microsoftonline.com/<テナントB-ID>/v2.0` |
| Audience | `api://AzureADTokenExchange` |
| Subject | Managed Identity Object ID |

### 5. テナントB 側：Graph API 権限追加

1. アプリの登録 → API のアクセス許可 → アクセス許可の追加
2. 設定:

| 項目 | 設定 |
|------|------|
| API | Microsoft Graph |
| 権限タイプ | Application |
| 権限 | Sites.Selected |

> ※ この時点では Admin Consent は不要

### 6. テナントA 側：Admin Consent 実行

テナントA 管理者に以下 URL を実行してもらいます:

```
https://login.microsoftonline.com/<テナントA-ID>/adminconsent
  ?client_id=<テナントB-App-ClientID>
  &redirect_uri=http://localhost
```

成功すると テナントA に **Enterprise Application** が作成されます。

### 7. テナントA 側：SharePoint サイト権限付与

`grant-sites-selected.ps1` を編集して実行します。

```powershell
# 設定値を編集
$TenantA  = "<tenant-a-id>"
$ClientId = "<tenant-b-app-client-id>"
$AppDisplayName = "<app-display-name>"
$spHost   = "<tenant-a>.sharepoint.com"
$sitePath = "/sites/<site-name>"
$permissionRole = "read"   # read または write

# 実行
cd verification/tenant-b-graph-function
.\grant-sites-selected.ps1
```

---

## Functions アプリ

### エンドポイント

| 種類 | ルート / スケジュール |
|------|----------------------|
| HTTP Trigger | `GET /api/graph/cross-tenant-scan` |
| Timer Trigger | `%GRAPH_SCAN_SCHEDULE%` cron |

### 環境変数

| 変数名 | 必須 | 説明 |
|--------|------|------|
| `TENANT_A_ID` | ○ | アクセス先テナントA のテナント ID |
| `APP_CLIENT_ID` | ○ | テナントB で作成した Multitenant アプリの Client ID |
| `MANAGED_IDENTITY_CLIENT_ID` | △ | UAMI の Client ID（System Assigned の場合は省略可） |
| `SP_HOSTNAME` | ○ | テナントA の SharePoint ホスト名 |
| `SP_SITE_PATHS` | ○ | カンマ区切りのサイトパス |
| `MAX_ITEMS_PER_SITE` | - | サイトあたりの取得件数 (デフォルト: 20) |
| `GRAPH_SCAN_SCHEDULE` | ○ | Timer Trigger の cron 式 |

### ローカル実行

```powershell
cd verification/tenant-b-graph-function
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item local.settings.sample.json local.settings.json
# local.settings.json を編集して実際の値を設定
func start
```

> ⚠️ ローカルでは Managed Identity が使えないため、Azure 上にデプロイして検証してください。

HTTP テスト:

```powershell
curl "http://localhost:7071/api/graph/cross-tenant-scan"
```

### Azure へのデプロイ（テナントB）

```powershell
cd verification/tenant-b-graph-function
func azure functionapp publish <function-app-name>
```

---

## テナントB 側から テナントA に共有する情報

| 項目 | 用途 |
|------|------|
| App Client ID | Admin Consent URL |
| Tenant ID | 認証 |
| App Name | SharePoint 権限設定 |

## まとめ

| ポイント | 内容 |
|----------|------|
| 認証方式 | Managed Identity + Workload Identity Federation |
| Secret | 不要 |
| 証明書 | 不要 |
| テナント跨ぎ | Multitenant App + Admin Consent |
