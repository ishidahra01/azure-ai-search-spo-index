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

### 5.3.5 パターンB 別方式: Tenant A 側にアプリ登録（マルチテナント・クロステナント認証）

#### 5.3.5.1 構成

```text
[Azure Function (Tenant A)] -> [Entra App (Tenant A, multi-tenant)] -> [Graph(Tenant B)] -> [SharePoint(Tenant B)]
```

- Tenant A: 実行基盤を持つテナント（アプリも Tenant A に登録）
- Tenant B: SharePoint/Graph データを持つテナント

この方式では、**アプリ登録を実行基盤側（Tenant A）に配置**し、Tenant B のデータにクロステナントアクセスします。

#### 5.3.5.2 重要ポイント

- **アプリ登録は Tenant A 側**（実行基盤と同じテナント）
- **マルチテナント設定必須**（他テナントのリソースにアクセスするため）
- **権限同意（admin consent）は Tenant B 側で必要**
- `Sites.Selected` のサイト割当も Tenant B 側で実施
- Tenant A の Managed Identity を利用可能（同一テナント内のアプリのため）

#### 5.3.5.3 詳細手順

1. **Tenant A でアプリ登録**
   - Azure Portal > Microsoft Entra ID（Tenant A） > アプリの登録 > 新規登録
   - 名前: `sp-graph-ingest-cross-tenant-prod`
   - サポートされるアカウント種類: **任意の組織ディレクトリ内のアカウント（任意の Microsoft Entra ID テナント - マルチテナント）**
   - リダイレクト URI: 不要（App-only）
   - 登録後、以下を記録:
     - Application (client) ID
     - Directory (tenant) ID（Tenant A の ID）

2. **Tenant A でアプリ設定**
   - API のアクセス許可 > アクセス許可の追加
   - Microsoft Graph > Application permissions
   - `Sites.Selected` を追加
   - 「管理者の同意を与えます」は**ここでは実行しない**（このアプリは Tenant B のリソースにアクセスするため、Tenant A のリソースに対する権限は不要。権限同意は後述の手順3で Tenant B 側の管理者により実施）

3. **Tenant B で管理者同意（Admin Consent）を取得**
   
   方法A: 同意 URL を使用（推奨）
   ```
   https://login.microsoftonline.com/{Tenant-B-ID}/adminconsent?client_id={Client-ID}
   ```
   - `{Tenant-B-ID}`: Tenant B のテナント ID
   - `{Client-ID}`: 手順1で取得したアプリの Client ID
   - Tenant B の全体管理者またはアプリケーション管理者でこの URL にアクセス
   - 同意画面で「この組織の代理として同意する」をチェックして承認
   
   方法B: Tenant B の Entra ID ポータルから同意
   - Tenant B のポータルで Microsoft Entra ID > エンタープライズアプリケーション
   - 手順1で作成したアプリを検索（Client ID で検索）
   - アクセス許可 > 「管理者の同意を与えます」を実行

4. **Tenant B で Sites.Selected 割当**
   - Graph API を使用して対象サイトに権限を付与
   - PowerShell / Graph Explorer などを利用
   ```powershell
   # 例: PowerShell + Microsoft.Graph モジュール
   Connect-MgGraph -TenantId {Tenant-B-ID} -Scopes "Sites.FullControl.All"
   
   # サイトID取得
   $siteUrl = "https://{tenant}.sharepoint.com/sites/{sitename}"
   $site = Get-MgSite -Search $siteUrl
   
   # アプリに read 権限を付与
   $params = @{
       roles = @("read")
       grantedToIdentities = @(
           @{
               application = @{
                   id = "{Client-ID}"
                   displayName = "sp-graph-ingest-cross-tenant-prod"
               }
           }
       )
   }
   New-MgSitePermission -SiteId $site.Id -BodyParameter $params
   ```

5. **Tenant A の Function App 設定**
   - Azure Function の Managed Identity を有効化（System-assigned または User-assigned）
   - この MI は Tenant A のアプリにアクセスするために使用
   - アプリケーション設定（環境変数）:
     ```properties
     GRAPH_TENANT_ID={Tenant-B-ID}        # Graph データが存在するテナント
     GRAPH_CLIENT_ID={Client-ID}          # Tenant A に登録したアプリの Client ID
     GRAPH_AUTH_MODE=managed_identity     # Managed Identity を使用
     SP_HOSTNAME={tenant}.sharepoint.com  # Tenant B の SharePoint
     SP_SITE_PATHS=/sites/hr,/sites/legal
     ```

6. **認証フロー設定**
   - Function コード内で Managed Identity を使用してトークンを取得
   - トークン取得時のリソース/スコープ: `https://graph.microsoft.com/.default`
   - 取得したトークンは Tenant A のアプリとして認証される
   - このトークンで Tenant B の Graph API にアクセス（マルチテナント同意済みのため可能）

7. **接続テスト**
   - 小規模データで実行テスト
   - トークン取得の確認
   - Graph API `/sites` で Tenant B のサイトにアクセス
   - 許可されたサイトのみアクセス可能、未割当サイトで 403 を確認

8. **本番運用設定**
   - Tenant A と Tenant B 双方で監査ログを有効化
   - アラート設定（認証失敗、権限変更）
   - 定期的な権限レビュー（四半期ごと推奨）

9. **このリポジトリでの検証実装（Tenant A デプロイ用）**
   - 検証用 Function: `verification/tenant-a-graph-function`
   - 実装内容:
     - HTTP Trigger: `GET /api/graph/tenant-scan`
     - Timer Trigger: `%GRAPH_SCAN_SCHEDULE%`
     - 取得データ: ファイル URL、タイトル、サイズ、作成/更新日時、親パスなど
   - セットアップ:
     ```powershell
     cd verification/tenant-a-graph-function
     python -m venv .venv
     . .\.venv\Scripts\Activate.ps1
     pip install -r requirements.txt
     Copy-Item local.settings.sample.json local.settings.json
     ```
   - `local.settings.json` の主な値:
     ```json
     {
       "Values": {
         "GRAPH_AUTH_MODE": "managed_identity",
         "SP_HOSTNAME": "<tenant-b>.sharepoint.com",
         "SP_SITE_PATHS": "/sites/hr,/sites/legal",
         "GRAPH_SCAN_SCHEDULE": "0 */30 * * * *"
       }
     }
     ```
   - ローカル実行:
     ```powershell
     func start
     curl "http://localhost:7071/api/graph/tenant-scan"
     ```
   - 期待結果:
     - 200 応答で、`siteId` / `siteTitle` / `items[].url` / `items[].title` / `items[].lastModifiedDateTime` が返る
   - 負のテスト（403 確認）:
     - `SP_SITE_PATHS` に未割当サイト（例: `/sites/not-granted`）を追加して再実行
     - Graph 呼び出しが 403 になることを確認
   - Azure へデプロイ（Tenant A）:
     ```powershell
     cd verification/tenant-a-graph-function
     func azure functionapp publish <function-app-name>
     ```

#### 5.3.5.4 この構成の利点と注意点

**利点:**
- 実行基盤（Tenant A）とアプリ登録が同一テナントのため、Managed Identity が使いやすい
- Tenant A 側で完結する管理（アプリライフサイクル、認証情報）
- Tenant A のセキュリティポリシー（Conditional Access for Workload Identity など）を直接適用可能
- 複数の Graph データテナントに接続する場合、アプリは1つで管理可能（各テナントで admin consent を取得）

**注意点:**
- マルチテナント設定が必須（設定ミスで動作しないリスク）
- Tenant B での admin consent プロセスが追加で必要
- Tenant B 側の管理者が同意内容を理解している必要がある
- Tenant B 側からはエンタープライズアプリケーションとして表示される（外部アプリ扱い）

#### 5.3.5.5 推奨される使用場面

- 実行基盤（Azure Functions など）が特定のテナント（Tenant A）に存在
- 複数の外部 SharePoint テナント（Tenant B, C, D...）にアクセスする必要がある
- 認証情報管理を実行基盤側のテナント（Tenant A）で集中管理したい
- Tenant A 側の Managed Identity や Workload Identity の制御を活用したい

---

### 5.3.6 パターン比較: Tenant A 登録 vs Tenant B 登録

#### 5.3.6.1 比較表

| 項目 | Tenant B 登録（5.3.3） | Tenant A 登録（5.3.5） |
|------|------------------------|------------------------|
| **アプリ登録場所** | Graph データテナント（Tenant B） | 実行基盤テナント（Tenant A） |
| **マルチテナント設定** | 必須 | 必須 |
| **Admin Consent 場所** | Tenant B（登録と同じ） | Tenant B（別テナント） |
| **Managed Identity 利用** | 別テナントのアプリのため<br/>Federated Credential 必要※1 | 同一テナントのため容易<br/>（直接利用可能） |
| **認証情報管理** | Tenant B で管理<br/>（データ側で集中） | Tenant A で管理<br/>（実行側で集中） |
| **複数テナント接続** | テナントごとにアプリ登録が必要 | 1つのアプリで複数テナント対応可能<br/>（各テナントで consent） |
| **セキュリティ境界** | データテナント側で制御しやすい | 実行テナント側で制御しやすい |
| **Tenant B の管理負荷** | アプリ管理も含めて高い | consent と Sites.Selected のみ |
| **Tenant A の管理負荷** | 低い（認証のみ） | アプリ管理を含めて高い |

※1: Tenant B のアプリに Federated Credential（Workload Identity）を設定し、Tenant A の Managed Identity と紐付ける必要があります（5.3.3 手順6参照）。

> 注記: どちらの方式も、クロステナントでのトークン取得を実現するためにマルチテナント設定が必要です。

#### 5.3.6.2 使い分けの判断基準

**Tenant B 登録（5.3.3）を選択すべきケース:**

1. **データガバナンスを優先する場合**
   - SharePoint データの所有者（Tenant B）がアプリのライフサイクルも管理したい
   - データ側のセキュリティポリシーを厳格に適用したい
   - 外部の実行基盤からのアクセスを明確に「外部アクセス」として扱いたい

2. **単一テナント連携の場合**
   - Tenant A から Tenant B への接続のみ
   - 複数テナント対応の予定がない

3. **データテナント側の責任範囲を明確にする場合**
   - データ漏洩時の責任所在をデータ側テナントに明確化
   - コンプライアンス要件でデータ側の管理が求められる

**Tenant A 登録（5.3.5）を選択すべきケース:**

1. **複数テナント接続が必要な場合**
   - Tenant A の実行基盤から複数の SharePoint テナント（B, C, D...）にアクセス
   - 統合データレイク・統合検索などのシナリオ

2. **実行基盤側で認証管理を集中したい場合**
   - アプリのライフサイクル管理を実行側で統一
   - Managed Identity などの実行基盤の認証機能を最大限活用

3. **Tenant B 側の管理負荷を軽減したい場合**
   - データ側テナント（Tenant B）はデータとサイト権限の管理に専念
   - アプリ登録・認証情報管理は実行側に委譲

4. **SaaS・マルチテナント SaaS として提供する場合**
   - サービス提供者（Tenant A）が複数の顧客テナントに接続
   - 顧客側（Tenant B）にはアプリ管理を要求しない設計

#### 5.3.6.3 ハイブリッド構成の検討

実務では以下のような組み合わせも検討されます:

- **Phase 1**: Tenant B 登録でスタート（PoC・小規模）
- **Phase 2**: Tenant A 登録に移行（本番・拡張時）
- **併用**: 重要度の高いテナントは Tenant B 登録、その他は Tenant A 登録

いずれの方式でも、**Sites.Selected の適切な割当**と**監査ログの保全**が最重要です。

#### 5.3.6.4 推奨構成（結論）

- **データガバナンス優先・単一テナント接続**: Tenant B 登録（5.3.3）
- **マルチテナント接続・実行基盤統合管理**: Tenant A 登録（5.3.5）
- **迷った場合**: まず Tenant B 登録で開始し、必要に応じて Tenant A 登録に移行

> 重要: どちらの方式でも、Tenant B での admin consent と Sites.Selected 割当は必須です。

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
