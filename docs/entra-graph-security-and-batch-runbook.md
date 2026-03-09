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

## 5.2 同一テナント構成: Azure テナントと Graph データテナントが同一

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

## 5.3 マルチテナント構成: Azure テナントと Graph データテナントが別

### 5.3.1 なぜこの構成が必要か（設計背景と機能制約）

#### テナント分離が発生する背景

Azure の実行基盤（Functions など）と SharePoint Online の組織が **異なる Entra ID テナント** に属するケースがあります。

典型例:
- 社内の IT 部門テナント（テナントB）から、事業部門テナント（テナントA）の SharePoint にアクセス
- SaaS プロバイダー（テナントB）が顧客テナント（テナントA）の SharePoint にアクセス
- グループ企業（テナントB）が親会社テナント（テナントA）の SharePoint にアクセス

#### 本ドキュメントのテナント定義

| 名称 | 役割 |
|------|------|
| **テナントA** | 取得元の SharePoint Online (SPO) が存在するテナント |
| **テナントB** | SPO のデータを取得するアプリの実行基盤があるテナント |

#### Managed Identity のクロステナント制約（重要）

マルチテナント設計を理解する上で最も重要な技術制約が **Managed Identity（MI）のクロステナント制限**です。

```text
【失敗する構成（なぜ動かないのか）】

テナントB の Azure Functions
  └ Managed Identity (MI)
       └ get_token() → テナントB の Entra ID がトークンを発行
                             ↓
                       テナントA の Graph API へ渡す
                             ↓
                       ❌ 認証失敗
                          （テナントA はこのトークンを受け入れない）
```

- テナントB の MI が取得するアクセストークンは、**テナントB の Entra ID が発行したもの**
- テナントA の Graph API（`https://graph.microsoft.com`）を呼び出すには、**テナントA が発行したトークン**が必要
- テナントB 発行のトークンをそのまま Graph に渡しても、テナントA のリソースへの権限がないため認証は失敗する

この制約を回避するために、以下のいずれかの認証方式が必要です:

| 方式 | 概要 | 推奨度 |
|------|------|--------|
| Workload Identity Federation (FIC) | テナントB の UAMI トークンを Client Assertion として使い、テナントA 向けトークンを取得 | ★★★ 推奨（Secret 不要） |
| Client Secret / 証明書 | テナントA エンドポイント向けにクライアント認証でトークンを直接取得 | ★★ 代替（Secret 管理が必要） |

---

### 5.3.2 全体設計とアプリ登録先の選択

マルチテナント構成では、まず **「アプリ登録をどちらのテナントに置くか」** を決め、次に **「認証方式」** を選択します。

```text
┌──────────────────────────────────────────────────────────────────────┐
│                       マルチテナント構成                                │
│                                                                      │
│  【第1の選択】アプリ登録をどちらのテナントに置くか？                         │
│                                                                      │
│  ┌─────────────────────────────────┐  ┌────────────────────────────┐ │
│  │ パターンA: テナントA にアプリ登録   │  │ パターンB: テナントBにアプリ登録│ │
│  │ (データ側が管理)                  │  │ (実行基盤側が管理)           │ │
│  │                                 │  │                            │ │
│  │ 【第2の選択】認証方式             │  │ 【第2の選択】認証方式         │ │
│  │                                 │  │                            │ │
│  │  FIC (OIDC Federation):         │  │  FIC (Managed Identity):   │ │
│  │   ・GitHub Actions OIDC         │  │   ・UAMI + FIC  ★推奨      │ │
│  │   ・外部 OIDC issuer            │  │   ※ MI がテナントB にある    │ │
│  │   ※ MI は使用不可               │  │     ため FIC 紐付け可能     │ │
│  │    (MI はテナントB 発行の         │  │                            │ │
│  │     トークンのため               │  │  FIC (OIDC Federation):    │ │
│  │     テナントA の FIC に           │  │   ・GitHub Actions OIDC    │ │
│  │     紐付けできない)              │  │   ・外部 OIDC issuer       │ │
│  │                                 │  │                            │ │
│  │  Client Secret / 証明書:        │  │  Client Secret / 証明書:   │ │
│  │   ・非推奨                      │  │   ・非推奨                  │ │
│  └─────────────────────────────────┘  └────────────────────────────┘ │
│                                                                      │
│  【共通の必須作業】                                                     │
│  ・テナントA での admin consent                                        │
│  ・テナントA での Sites.Selected サイト割当                               │
└──────────────────────────────────────────────────────────────────────┘
```

| | パターンA | パターンB |
|---|---|---|
| **アプリ登録場所** | テナントA（データ側） | テナントB（実行側） |
| **FIC で使える認証元** | OIDC のみ（GitHub Actions / 外部 OIDC issuer） | **Managed Identity（UAMI）** + OIDC |
| **FIC で MI が使えるか** | ❌（MI はテナントB 発行のため紐付け不可） | ✅（アプリと UAMI が同一テナント） |
| **Secret 不要（Azure 実行基盤）** | ❌（MI 不可のため OIDC or Secret） | ✅（UAMI + FIC） |
| **データガバナンス** | テナントA がアプリ管理 | テナントB がアプリ管理 |
| **複数テナント接続** | テナントごとにアプリ登録が必要 | 1つのアプリで複数テナント対応可 |
| **推奨度** | ガバナンス要件次第 | ★★★ 推奨（Azure 実行基盤の場合） |
| **どちらも必須** | テナントA での admin consent + Sites.Selected 割当 ||

> **Azure Functions 等の Azure 実行基盤で Secret 不要にしたい場合は、パターンB（テナントB 登録 + UAMI + FIC）が唯一の選択肢**です。パターンA では MI をFIC に紐付けできないため、Azure 実行基盤でも OIDC Federation か Client Secret が必要になります。

---

### 5.3.3 パターンA: テナントA 側にアプリ登録

テナントA（SPO データ側）にアプリを登録するパターンです。データガバナンスの観点から、SharePoint データの所有テナントがアプリのライフサイクルを管理したい場合に採用します。

#### 認証フロー（OIDC Federation の場合）

```text
[外部実行基盤 (GitHub Actions / 外部 OIDC issuer)]
      │
      │ 1. OIDC トークン取得（issuer / subject / audience）
      │ 2. そのトークンを Client Assertion として使用
      ▼
[テナントA: Entra App (Multitenant) + Federated Credential → 外部 OIDC]
      │  Sites.Selected 割当済み
      ▼
[テナントA: Graph API → SharePoint Online]
```

#### 認証フロー（Client Secret の場合 — 非推奨）

```text
[テナントB: Azure Functions / 任意の実行基盤]
      │
      │ 1. テナントA エンドポイント向けに client_secret でトークン取得
      │    (https://login.microsoftonline.com/{テナントA-ID}/oauth2/v2.0/token)
      ▼
[テナントA: Entra App + Client Secret]
      │  Sites.Selected 割当済み
      ▼
[テナントA: Graph API → SharePoint Online]
```

**設計のポイント:**
- テナントA 側でアプリのライフサイクルを管理するため、**データガバナンスが明確**
- FIC を使う場合の認証元は **OIDC Federation のみ**（GitHub Actions / 外部 OIDC issuer）
- **Managed Identity は FIC に紐付けできない**（MI はテナントB の Entra ID が発行するトークンであり、テナントA のアプリの Federated Credential として設定できない）
- Azure Functions 等の Azure 実行基盤で MI ベースの Secret 不要構成にしたい場合は、パターンB を選択すること
- Client Secret を使う場合は Key Vault 管理必須（[原則2](#1-先に押さえる設計原則最重要) に反するため非推奨）

**適用シナリオ:**
- GitHub Actions / 外部 CI から SharePoint にアクセスする
- テナントA の管理者がアプリの権限管理を完全に掌握したい
- 複数の実行基盤からアクセスする場合、テナントA ごとにアプリを管理する方針

---

### 5.3.4 パターンB: テナントB 側にアプリ登録（Workload Identity Federation）

テナントB（実行基盤側）にアプリを登録するパターンです。**Azure 実行基盤で MI ベースの Secret 不要構成が可能**なため、Azure Functions / App Service / Container Apps からのバッチ実行ではこのパターンを推奨します。

```text
[テナントB: Azure Functions + UAMI]
      │
      │ 1. UAMI で api://AzureADTokenExchange トークン取得
      │ 2. そのトークンを Client Assertion として使用（Workload Identity Federation）
      ▼
[テナントB: Entra App (Multitenant) + Federated Credential → UAMI]
      │
      │ テナントA エンドポイント向けにトークン取得
      ▼
[テナントA: Enterprise Application (admin consent 済み)]
      │  Sites.Selected 割当済み
      ▼
[テナントA: Graph API → SharePoint Online]
```

**設計のポイント:**
- テナントB 側でアプリのライフサイクルを管理（Client Secret / 証明書不要）
- Workload Identity Federation（FIC）により、**UAMI のトークンを Client Assertion として使用**
- アプリと UAMI が同一テナント（テナントB）にあるため、Federated Credential の紐付けが可能
- Federated Credential で テナントB の UAMI と紐付け（issuer/subject/audience の正確な設定が必要）
- 複数の SharePoint テナント（テナントA, テナントA', ...）にアクセスする場合、1つのアプリで対応可能（各テナントで admin consent を取得）

**詳細手順**: [`verification/tenant-b-graph-function/`](../verification/tenant-b-graph-function/)

---

### 5.3.5 共通の必須作業（どちらのパターンでも必要）

テナントA 側で必ず実施する:

1. **Admin Consent の実行**
   ```
   https://login.microsoftonline.com/{テナントA-ID}/adminconsent?client_id={Client-ID}
   ```
   - テナントA の全体管理者またはアプリケーション管理者が実行
   - 同意画面で「この組織の代理として同意する」をチェックして承認

2. **Sites.Selected のサイト割当**（付与だけでは不可。対象サイトへの明示割当が必要）
   ```powershell
   # PowerShell + Microsoft.Graph モジュール を使った割当例
   # 詳細は verification/tenant-b-graph-function/grant-sites-selected.ps1 を参照
   Connect-MgGraph -TenantId {テナントA-ID} -Scopes "Sites.FullControl.All"
   $site = Get-MgSite -Search "https://{tenant}.sharepoint.com/sites/{sitename}"
   $params = @{
       roles = @("read")
       grantedToIdentities = @(@{ application = @{ id = "{Client-ID}"; displayName = "{AppName}" } })
   }
   New-MgSitePermission -SiteId $site.Id -BodyParameter $params
   ```

3. **監査ログの保全設定**（テナントA / テナントB 双方）

---

### 5.3.6 パターンの選択基準

**パターンA（テナントA 登録）を選択すべきケース:**

- データガバナンス優先: SharePoint データの所有者（テナントA）がアプリのライフサイクルも管理したい
- GitHub Actions / 外部 CI からのアクセスが主体（OIDC Federation を活用）
- データ漏洩時の責任所在をデータ側テナントに明確化したい
- ただし、Azure 実行基盤で MI ベースの Secret 不要構成は**実現できない**点に注意

**パターンB（テナントB 登録 + WIF）を選択すべきケース（推奨）:**

- Azure Functions 等の Azure 実行基盤で **Secret 不要構成**を実現したい（UAMI + FIC）
- 実行基盤側で認証管理を集中したい（アプリのライフサイクル管理を実行側で統一）
- 複数テナント接続が必要: テナントB の実行基盤から複数の SharePoint テナントにアクセス
- テナントA 側の管理負荷を軽減したい（テナントA は consent と Sites.Selected のみ）

**推奨**: Azure 実行基盤の場合は **パターンB（テナントB 登録 + UAMI + WIF）** で構築する。GitHub Actions 等の外部 CI が主体の場合は パターンA も選択肢に入る。

> 重要: どちらのパターンでも、テナントA での admin consent と Sites.Selected 割当は必須です。

---

### 5.3.7 マルチテナントでの失敗パターン

- **同意テナントを誤る**: テナントB で同意して テナントA で未同意（consent は必ずテナントA 側で実行）
- **Sites.Selected の割当漏れ**: 権限を付与しただけで、対象サイトへの明示割当を忘れる
- **issuer/subject/audience の不一致**: Federated Credential の設定値と実行基盤の値が合っていない
- **テナントID の取り違え**: `GRAPH_TENANT_ID` / `TENANT_A_ID` にテナントB の ID を誤設定
- **MI をそのまま使う**: テナントB の MI トークンをテナントA の Graph にそのまま渡して認証失敗（[5.3.1 参照](#531-なぜこの構成が必要か設計背景と機能制約)）

---

## 5.4 実装時の推奨設定値（例）

```properties
# Graph / Entra
GRAPH_TENANT_ID=<graph-data-tenant-id>
GRAPH_CLIENT_ID=<entra-app-client-id>
GRAPH_AUTH_MODE=client_secret   # クロステナントの場合は client_secret（managed_identity はクロステナント不可）

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
