# 検証用サンプル（マルチテナント構成）

このディレクトリには、`docs/entra-graph-security-and-batch-runbook.md` に記載のマルチテナント構成を実際に検証するためのサンプル Azure Functions が含まれています。

## テナント定義

本リポジトリのサンプル全体で以下の定義に統一しています:

| 名称 | 役割 |
|------|------|
| **テナントA** | 取得元の SharePoint Online (SPO) が存在するテナント |
| **テナントB** | SPO のデータを取得するアプリの実行基盤があるテナント |

## ディレクトリ構成

### [`tenant-b-graph-function/`](./tenant-b-graph-function/)

- **対応セクション**: Runbook [5.3.3](../docs/entra-graph-security-and-batch-runbook.md)（パターンA: テナントB 側アプリ登録 + Workload Identity Federation）
- **認証方式**: Workload Identity Federation (UAMI)
- **概要**: Secret 不要の推奨方式でクロステナントアクセスを検証。

## 前提条件

サンプルで、以下の作業がテナントA 側で必要です:

1. テナントB のアプリに対する **Admin Consent** の実行
   ```
   https://login.microsoftonline.com/<テナントA-ID>/adminconsent?client_id=<テナントB-App-ClientID>
   ```
2. `Sites.Selected` によるシェアポイントサイトへの権限付与
   - `tenant-b-graph-function/grant-sites-selected.ps1` を使用

詳細は `docs/entra-graph-security-and-batch-runbook.md` の該当セクションを参照してください。
