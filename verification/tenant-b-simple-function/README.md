# テナントB シンプル認証 Graph 検証 Function

このフォルダは、`docs/entra-graph-security-and-batch-runbook.md` の `5.3.5`（テナントB 側アプリ登録方式・Client Secret によるクロステナント認証）を検証するための Azure Functions (Python) です。

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

## 前提

- テナントA 側で admin consent 済み（テナントB のアプリに対して）
- テナントA 側で `Sites.Selected` のサイト割当済み
- その割当対象のアプリ ID が、実際にトークン発行される主体と一致していること
- **`client_secret` 方式**: テナントB に登録したアプリの `GRAPH_CLIENT_ID` / `GRAPH_CLIENT_SECRET` が設定済みであること

## ローカル実行

```powershell
cd verification/tenant-b-simple-function
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item local.settings.sample.json local.settings.json
func start
```

HTTP テスト:

```powershell
curl "http://localhost:7071/api/graph/tenant-scan"
```

## Azure へのデプロイ（テナントB）

```powershell
cd verification/tenant-b-simple-function
func azure functionapp publish <function-app-name>
```

## アプリ設定（Function App）

必須:

- `GRAPH_AUTH_MODE` (`client_secret` を推奨。クロステナントでは `managed_identity` は使用不可)
- `GRAPH_TENANT_ID`（テナントA の ID）
- `SP_HOSTNAME`（テナントA の SharePoint ホスト名）
- `SP_SITE_PATHS`
- `GRAPH_SCAN_SCHEDULE`

`client_secret` 利用時のみ必須（クロステナント検証では常に必要）:

- `GRAPH_CLIENT_ID`（テナントB に登録したアプリの Client ID）
- `GRAPH_CLIENT_SECRET`

`managed_identity` について:

- このモードはコード上サポートされていますが、**クロステナントシナリオ（テナントB → テナントA）では動作しません**
- テナントB の MI トークンはテナントB 発行であり、テナントA の Graph リソースへのアクセス権がないため認証が失敗します
- 同一テナント内での検証目的にのみ使用してください

## 注意

クロステナントアクセスでは `client_secret` 方式（`GRAPH_AUTH_MODE=client_secret`）を使用してください。
`managed_identity` モードはテナントB の Entra ID が発行するトークンを使うため、テナントA の Graph API へのアクセスは成立しません。
本番環境での Secret 運用は Key Vault 参照を利用し、Secret の直接設定を避けることを推奨します。
例: `@Microsoft.KeyVault(SecretUri=https://<vault-name>.vault.azure.net/secrets/<secret-name>)`
より高度なシークレットレス構成（FIC/OIDC）については `verification/tenant-b-graph-function` を参照してください。
