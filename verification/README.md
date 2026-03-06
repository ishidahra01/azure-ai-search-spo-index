# 検証用サンプル（マルチテナント構成）

このディレクトリには、`docs/entra-graph-security-and-batch-runbook.md` に記載のマルチテナント構成を実際に検証するためのサンプル Azure Functions が含まれています。

## テナント定義

本リポジトリのサンプル全体で以下の定義に統一しています:

| 名称 | 役割 |
|------|------|
| **テナントA** | 取得元の SharePoint Online (SPO) が存在するテナント |
| **テナントB** | SPO のデータを取得するアプリの実行基盤があるテナント |

## ディレクトリ構成

### [`tenant-b-simple-function/`](./tenant-b-simple-function/)

- **対応セクション**: Runbook [5.3.5](../docs/entra-graph-security-and-batch-runbook.md)（テナントB 側アプリ登録）
- **認証方式**: Managed Identity または Client Secret
- **概要**: シンプルな認証でクロステナントアクセスを検証。PoC や初期検証に最適。

### [`tenant-b-graph-function/`](./tenant-b-graph-function/)

- **対応セクション**: Runbook [5.3.3](../docs/entra-graph-security-and-batch-runbook.md)（テナントA 側アプリ登録）
- **認証方式**: Workload Identity Federation (UAMI)
- **概要**: Secret 不要の本番向け認証方式でクロステナントアクセスを検証。

## どちらを使うか

```
まず試したい / PoC
  └→ tenant-b-simple-function
        ・Managed Identity または Client Secret で手軽に動作確認
        ・テナントA で admin consent と Sites.Selected 割当が必要
        ・クロステナントで動かない場合は client_secret で先に疎通確認

本番環境 / Secret 不要構成
  └→ tenant-b-graph-function
        ・Workload Identity Federation (UAMI) を使用
        ・Client Secret も証明書も不要
        ・テナントA で admin consent と Sites.Selected 割当が必要
```

## 共通の前提条件

両方のサンプルで、以下の作業がテナントA 側で必要です:

1. テナントB のアプリに対する **Admin Consent** の実行
   ```
   https://login.microsoftonline.com/<テナントA-ID>/adminconsent?client_id=<テナントB-App-ClientID>
   ```
2. `Sites.Selected` によるシェアポイントサイトへの権限付与
   - `tenant-b-graph-function/grant-sites-selected.ps1` を使用

詳細は `docs/entra-graph-security-and-batch-runbook.md` の該当セクションを参照してください。
