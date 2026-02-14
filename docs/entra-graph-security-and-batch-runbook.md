# Entra ID アプリ登録（Graph API 実行用）とバッチ処理（Azure 基盤）Runbook

> 対象: SharePoint Online データを Microsoft Graph API 経由で取得し、Azure AI Search へ連携する運用担当者

## 1. 先に押さえる設計原則（最重要）

- **原則1: 「何を読めるか」と「誰が叩けるか」を分離して設計する**
  - 何を読めるか: `Sites.Selected`（SharePoint の対象サイト限定）
  - 誰が叩けるか: 認証方式（Delegated / Managed Identity / OIDC Federation）
- **原則2: Secret を極力使わない**
  - 第一選択は **Managed Identity**（Azure 実行基盤）
  - Azure 外/CI は **Workload Identity Federation (OIDC)**
- **原則3: 最小権限 + 運用監査**
  - 必要最小限の Graph 権限
  - PIM/JIT・監査ログ・アラートを前提にする

---

## 2. 方式の全体像（README 貼り付け用）

```text
実務的なおすすめ（多い順）

① 人が操作するツール
   → Delegated + Conditional Access(MFA) + Sites.Selected

② バッチ / 自動処理
   → App-only + Sites.Selected + Secret廃止
      ├─ ②-A Azure実行基盤：Managed Identity
      ├─ ②-B GitHub Actions：OIDC
      └─ ②-C その他実行基盤：OIDC（Other issuer）
```

### ① 人が操作するツール

```text
[ User ]
   │  (サインイン + MFA)
   ▼
[ Entra ID ]
   │  - Delegated permission
   │  - Conditional Access
   │      ・MFA 必須
   │      ・特定ユーザー / グループ
   ▼
[ App / Tool ]
   │  (Delegated Access Token)
   ▼
[ Microsoft Graph ]
   ▼
[ SharePoint Online ]
   └─ Sites.Selected
       └─ 許可された特定サイトのみ
```

### ②-A バッチ（Azure 実行基盤）

```text
[ Azure Function / VM / App Service ]
   │  (Managed Identity)
   ▼
[ Entra ID ]
   │  - App Registration
   │  - Application permission
   │      ・Sites.Selected
   ▼
[ Microsoft Graph ]
   ▼
[ SharePoint Online ]
   └─ Sites.Selected
       └─ 許可された特定サイトのみ
```

### ②-B バッチ（GitHub Actions）

```text
[ GitHub Actions ]
   │  (OIDC Token)
   │   ・org / repo
   │   ・branch
   │   ・environment
   ▼
[ Entra ID ]
   │  - Federated Credential (GitHub)
   │  - App Registration
   │      ・Sites.Selected
   │  - Conditional Access (Workload ID)
   │      ・特定IPのみ 等
   ▼
[ Microsoft Graph ]
   ▼
[ SharePoint Online ]
   └─ Sites.Selected
       └─ 許可された特定サイトのみ
```

### ②-C バッチ（その他の基盤）

```text
[ External Runtime ]
   │  (OIDC Token)
   │   ・issuer
   │   ・subject
   │   ・audience
   ▼
[ Entra ID ]
   │  - Federated Credential (Other issuer)
   │  - App Registration
   │      ・Sites.Selected
   ▼
[ Microsoft Graph ]
   ▼
[ SharePoint Online ]
   └─ Sites.Selected
       └─ 許可された特定サイトのみ
```

---

## 3. Entra ID アプリ登録（Graph API 実行用）詳細手順

以下は **App-only（バッチ前提）** の推奨手順です。Delegated が必要な場合は 3.7 を参照してください。

### 3.1 事前準備（チェックリスト）

- [ ] Graph データが存在するテナント（SharePoint テナント）を特定
- [ ] 実行基盤（Azure Functions / App Service / Container Apps / VM）を確定
- [ ] 対象 SharePoint サイト一覧（URL・用途・データ分類）を確定
- [ ] 運用責任者（アプリ所有者/権限付与担当/監査担当）を確定

### 3.2 アプリ登録

1. Azure Portal > **Microsoft Entra ID** > **アプリの登録** > **新規登録**
2. 入力値
   - 名前: `sp-graph-ingest-prod` など環境識別子つき
   - サポートされるアカウント種類:
     - 同一テナントのみなら **シングルテナント** 推奨
     - 別テナント連携なら **マルチテナント**（後述の 5 章）
   - リダイレクト URI: App-only のみなら不要
3. 登録後、以下を記録
   - Application (client) ID
   - Directory (tenant) ID
   - Object ID（運用時参照用）

### 3.3 Graph API 権限の設計（推奨最小）

- 基本方針
  - SharePoint ファイル取得中心: `Sites.Selected`
  - 追加で必要な場合のみ拡張（例: ユーザー解決が必要なとき `User.Read.All`）
- 非推奨
  - いきなり `Sites.Read.All` / `Files.Read.All` を本番で恒久利用

### 3.4 Graph API 権限の付与

1. アプリ > **API のアクセス許可** > **アクセス許可の追加**
2. `Microsoft Graph` > `Application permissions`
3. `Sites.Selected` を追加
4. **管理者の同意を与えます** を実行（該当テナント管理者）

### 3.5 Sites.Selected の実効化（サイト単位割当）

`Sites.Selected` は「付与しただけではアクセス不可」で、**対象サイトへの明示割当**が必要です。

- 代表的な割当フロー
  1. 対象サイトの `siteId` を取得
  2. Graph API で対象アプリに site permission を付与（`read` or `write`）
  3. 付与結果を一覧取得して監査証跡に保存

> 実運用では、初期導入時に対象サイトを IaC か自動スクリプトで一括割当し、変更管理チケットで追加/削除する運用にしてください。

### 3.6 認証情報（Credential）の選択

優先順位:
1. **Managed Identity**（Azure 実行）
2. **Federated Credential (OIDC)**（Azure 外/CI）
3. **証明書**（やむを得ない場合）
4. **Client Secret**（非推奨）

### 3.7 Delegated（人が操作するツール）を使う場合

- API 権限は Delegated を選択
- Conditional Access で以下を強制
  - MFA 必須
  - 許可ユーザー/グループ限定
  - 必要に応じて準拠デバイス/場所制限
- SharePoint データ範囲をさらに絞る必要がある場合は、アプリ設計とサイト権限設計を分離して実装

---

## 4. セキュリティ設計（推奨対策カタログ）

## 4.1 ID・認証

- Secret を廃止し、Managed Identity/OIDC に移行
- Entra アプリ所有者は最小人数（2〜4 名程度）
- 高権限ロールは PIM（JIT）で時間制約付き昇格

## 4.2 認可（Authorization）

- Graph 権限は最小化（`Sites.Selected` 優先）
- サイト割当は「業務データ分類」単位で分離
- 検索インデックス側 ACL も必須（取得できても見せない設計）

## 4.3 実行基盤

- Azure リソースは Private Endpoint / VNet Integration を優先
- egress 制御（Firewall / NAT / NSG）を設計
- 実行環境の構成変更は IaC（Bicep/Terraform）で追跡可能に

## 4.4 秘密情報・鍵管理

- どうしても証明書/Secret が必要な場合は Key Vault 管理
- 期限・ローテーション・失効手順を Runbook 化
- アプリ設定への平文埋め込み禁止

## 4.5 監査・検知

- 監査対象
  - アプリ権限追加/同意
  - Sites.Selected 割当変更
  - サインイン失敗急増・異常トークン要求
- ログ連携
  - Entra サインインログ / 監査ログ
  - Azure Monitor / Log Analytics / Sentinel（必要に応じて）

## 4.6 運用ガバナンス

- 四半期ごとのアクセスレビュー
- 不要サイト割当の棚卸し
- 障害時のフェイルセーフ（取り込み停止・再開手順）

---

## 5. バッチ処理（Azure 基盤）超詳細手順

本章は、**「この手順どおりで作業ミスなく構築できる」**ことを狙った実装手順です。

## 5.0 先に決める実行モデル

- Azure Functions（Timer Trigger）: 定期バッチ向き
- App Service / Container Apps（WebJob・Cron）: 既存アプリ統合向き
- VM: 既存資産活用向き（運用コストは高め）

以下は Azure Functions を例に説明します（他基盤でも同じ考え方）。

## 5.1 共通準備

1. リソースグループ作成
2. Function App 作成
3. Function App の **Managed Identity** を有効化
   - 同一テナント: System-assigned で開始推奨
   - 別テナント: User-assigned + Federation を使うケースが多い
4. Application Insights / Log Analytics 接続
5. Key Vault（必要なら）接続

---

## 5.2 パターンA: Azure テナントと Graph データテナントが同一

### 5.2.1 構成

```text
[Azure Function (MI)] -> [同一テナント Entra App] -> [Graph] -> [SharePoint 同一テナント]
```

### 5.2.2 手順

1. **Entra アプリ登録**（3章）
   - シングルテナント
   - Graph Application Permission: `Sites.Selected`
   - 管理者同意
2. **対象サイト割当（Sites.Selected）**
   - 対象サイトを read 権限で明示付与
3. **Function の Managed Identity 有効化**
4. **Managed Identity とアプリの紐付け設計**
   - 実装で MI トークンを取得し Graph にアクセス
   - 必要に応じて中継 API（自社 API）を挟み、アクセス境界を分離
5. **アプリ設定（環境変数）**
   - `GRAPH_TENANT_ID`
   - `GRAPH_CLIENT_ID`
   - `SP_HOSTNAME`
   - `SP_SITE_PATHS`
   - `AZURE_SEARCH_ENDPOINT` など
6. **初回ドライラン**
   - 1 サイト / 10 ファイル程度で実行
   - 429, 403, 404 の扱いを確認
7. **本番切替**
   - バッチ頻度（例: 1h, 4h, 日次）
   - 並列度と再試行回数を調整
8. **運用監視設定**
   - 実行失敗アラート
   - 取り込み件数急減アラート
   - 認証エラー増加アラート

### 5.2.3 この構成の利点/注意

- 利点: Secret 不要、構成が最もシンプル
- 注意: Workload ID 向け CA で制御できる範囲には制約があるため、ネットワーク・実行基盤側制御を必ず併用

---

## 5.3 パターンB: Azure テナントと Graph データテナントが別（マルチテナント）

### 5.3.1 典型構成

```text
[Azure Function (Tenant A)] -> [Entra App (Tenant B, multi-tenant)] -> [Graph(Tenant B)] -> [SharePoint(Tenant B)]
```

- Tenant A: 実行基盤を持つテナント
- Tenant B: SharePoint/Graph データを持つテナント

### 5.3.2 重要ポイント

- **権限同意（admin consent）は Tenant B 側で必要**
- `Sites.Selected` のサイト割当も Tenant B 側で実施
- 実行主体の真正性は、Federation/OIDC または厳格な credential 運用で担保

### 5.3.3 手順（推奨フロー）

1. **Tenant B でアプリ登録**
   - マルチテナント設定
   - Graph Application Permission に `Sites.Selected`
2. **Tenant B 管理者で admin consent**
3. **Tenant B で Sites.Selected 割当**
   - 対象 SharePoint サイトごとに read を付与
4. **Tenant A の実行基盤（Function）準備**
   - Managed Identity を有効化
5. **認証方式を決定**
   - 推奨: OIDC Federation（workload identity）
   - 代替: 証明書（Key Vault 保管）
   - 非推奨: Client Secret
6. **Federated Credential 設定（Tenant B のアプリ側）**
   - issuer / subject / audience を実行基盤に合わせて正確に設定
   - subject は可能な限り限定的に（環境・ワークロード単位）
7. **接続テスト**
   - トークン取得
   - Graph `/sites` 参照
   - 許可外サイトで 403 になることを確認（負のテスト）
8. **本番運用設定**
   - 監査ログを Tenant A/B 双方で保全
   - 片側障害時の再実行方針（RPO/RTO）を決定

### 5.3.4 マルチテナントでの失敗パターン

- 同意テナントを誤る（A で同意して B で未同意）
- Sites.Selected を付与しただけで割当していない
- issuer/subject/audience の不一致
- テナントID固定ミス（アプリ設定値の取り違え）

---

## 5.4 実装時の推奨設定値（例）

```properties
# Graph / Entra
GRAPH_TENANT_ID=<graph-data-tenant-id>
GRAPH_CLIENT_ID=<entra-app-client-id>
GRAPH_AUTH_MODE=managed_identity   # or federated_oidc

# SharePoint
SP_HOSTNAME=<tenant>.sharepoint.com
SP_SITE_PATHS=/sites/hr,/sites/legal
SP_LIBRARY_NAMES=Documents

# Search
AZURE_SEARCH_ENDPOINT=https://<name>.search.windows.net
AZURE_SEARCH_INDEX_NAME=sp-docs
```

---

## 5.5 受け入れテスト（UAT）観点

- 正常系
  - 指定サイトのファイルを取得できる
  - 差分同期で更新分のみ処理できる
- 異常系
  - 非許可サイトが 403 で拒否される
  - 認証失敗時に再試行後エラー終了する
  - 429 で指数バックオフが動作する
- 運用系
  - アラートが期待どおりに通知される
  - 監査ログから権限変更履歴を追跡できる

---

## 6. 推奨構成（結論）

- 人が操作するツール: **Delegated + Conditional Access (MFA) + Sites.Selected**
- 自動処理（Azure 基盤）: **App-only + Sites.Selected + Managed Identity（同一テナント）**
- 自動処理（別テナント/CI 含む）: **App-only + Sites.Selected + OIDC Federation**

> 要点: `Sites.Selected` は「どのデータを読めるか」、認証方式は「誰が実行できるか」を制御する。
