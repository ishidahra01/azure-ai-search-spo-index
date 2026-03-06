# Tenant A Graph Verification Function

このフォルダは、`docs/entra-graph-security-and-batch-runbook.md` の `5.3.5` (Tenant A 側アプリ登録方式) を検証するための Azure Functions (Python) です。

## 何をするか

- Graph で `SP_SITE_PATHS` のサイト情報を取得
- 各サイトのドキュメント ライブラリ ルート配下のアイテムを取得
- ファイル URL、タイトル、更新日時、サイズなどのメタ情報を返却/ログ出力

## 構成

- HTTP Trigger: `GET /api/graph/tenant-scan`
- Timer Trigger: `%GRAPH_SCAN_SCHEDULE%` の cron で定期実行

## 前提

- Tenant B 側で admin consent 済み
- Tenant B 側で `Sites.Selected` のサイト割当済み
- その割当対象のアプリ ID が、実際にトークン発行される主体と一致していること

## ローカル実行

```powershell
cd verification/tenant-a-graph-function
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

## Azure へのデプロイ（Tenant A）

```powershell
cd verification/tenant-a-graph-function
func azure functionapp publish <function-app-name>
```

## アプリ設定（Function App）

必須:

- `GRAPH_AUTH_MODE` (`managed_identity` または `client_secret`)
- `SP_HOSTNAME`
- `SP_SITE_PATHS`
- `GRAPH_SCAN_SCHEDULE`

`client_secret` 利用時のみ必須:

- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`

`managed_identity` 利用時:

- User-assigned MI の場合のみ `MANAGED_IDENTITY_CLIENT_ID` を設定

## 注意

`managed_identity` でクロステナント検証する場合、テナント設計によっては MI 直接トークンで成立しないケースがあります。その場合は検証目的として `client_secret` で先に Graph 疎通を確認し、後続で Federation/証明書方式へ移行してください。
