# SharePoint → Azure AI Search ハンズオン

SharePoint Online のドキュメントを Azure AI Search に取り込み、高度な検索・RAG を実現するための実践的なハンズオン資料です。

## 概要

このリポジトリでは、SharePoint から Azure AI Search へのデータ取り込みを**2つのパターン**で学習できます:

### Pattern A: Graph API パターン
- Microsoft Graph API でファイルとメタデータを直接取得
- Python でカスタマイズ可能な取り込みパイプライン
- 本文抽出・チャンク分割・埋め込み生成を完全制御
- Graph delta API による差分同期

### Pattern B: SharePoint Indexer パターン
- Azure AI Search の組み込み SharePoint インデクサを使用
- 設定だけで自動取り込み
- スケジュール実行による自動差分同期
- 運用が簡単で保守性が高い

## アーキテクチャ

```
┌─────────────────────┐
│  SharePoint Online  │
│   (Documents)       │
└──────────┬──────────┘
           │
           ├─────────────────────┬───────────────────────┐
           │                     │                       │
    ┌──────▼────────┐   ┌───────▼────────┐   ┌─────────▼─────────┐
    │  Pattern A    │   │   Pattern B    │   │  Azure Document   │
    │  Graph API    │   │  SP Indexer    │   │  Intelligence     │
    │  + Python     │   │  (Native)      │   │  (Optional)       │
    └──────┬────────┘   └───────┬────────┘   └─────────┬─────────┘
           │                     │                       │
           └─────────────────────┴───────────────────────┘
                                 │
                        ┌────────▼──────────┐
                        │  Azure AI Search  │
                        │  (Index + Vector) │
                        └───────────────────┘
```

## 主な機能

- ✅ **本文抽出**: PDF, Word, Excel, PowerPoint からテキスト抽出
- ✅ **メタデータ**: ファイル情報、作成者、更新日時などを保持
- ✅ **チャンク分割**: 長文を適切なサイズに分割
- ✅ **ベクター検索**: Azure OpenAI で埋め込み生成 (オプション)
- ✅ **ACL (アクセス制御)**: ユーザー/グループ権限のフィルタリング
- ✅ **差分同期**: 追加・更新・削除の検知と反映
- ✅ **セマンティック検索**: 意味ベースの高度な検索
- ✅ **ハイブリッド検索**: キーワード + ベクター検索

## 前提条件

### Azure リソース
- **Azure AI Search** (Basic 以上推奨)
- **Azure OpenAI** (ベクター検索用、オプション)
- **Azure Document Intelligence** (高精度抽出用、オプション)

### SharePoint
- SharePoint Online サイトとドキュメントライブラリ
- テスト用サンプルファイル

### Azure AD アプリ登録 (Pattern A)
- アプリケーション権限: `Sites.Read.All`, `Files.Read.All`
- 管理者の同意

### 開発環境
- Python 3.9 以上
- Jupyter Notebook

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/azure-ai-search-spo-index.git
cd azure-ai-search-spo-index
```

### 2. 仮想環境の作成と依存関係のインストール

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Microsoft Graph API の認証情報取得 (Pattern A 用)

Pattern A を使用する場合、Azure AD アプリを登録して認証情報を取得する必要があります。

#### 3.1 Azure AD アプリの登録

1. **Azure Portal** にアクセス: [https://portal.azure.com](https://portal.azure.com)
2. **Microsoft Entra ID** (旧 Azure Active Directory) に移動
3. 左メニューから **アプリの登録** を選択
4. **新規登録** をクリック
   - **名前**: `SharePoint-AI-Search-App` (任意の名前)
   - **サポートされているアカウントの種類**: `この組織ディレクトリのみに含まれるアカウント`
   - **リダイレクト URI**: 空欄でOK
5. **登録** をクリック

#### 3.2 認証情報の確認・取得

登録後、アプリの **概要** ページで以下の情報を確認:

- **`GRAPH_TENANT_ID`** (テナント ID / ディレクトリ ID)
  - 概要ページの「ディレクトリ (テナント) ID」をコピー
  - 例: `12345678-1234-1234-1234-123456789abc`

- **`GRAPH_CLIENT_ID`** (アプリケーション (クライアント) ID)
  - 概要ページの「アプリケーション (クライアント) ID」をコピー
  - 例: `abcdefgh-1234-5678-90ab-cdefghijklmn`

#### 3.3 クライアントシークレットの作成

- **`GRAPH_CLIENT_SECRET`** (クライアント シークレット)
  1. 左メニューから **証明書とシークレット** を選択
  2. **クライアント シークレット** タブで **新しいクライアント シークレット** をクリック
  3. **説明**: `sharepoint-access` (任意)
  4. **有効期限**: `180 日 (6か月)` または `24か月` を選択
  5. **追加** をクリック
  6. ⚠️ **重要**: 作成された **値** をすぐにコピー (後で確認できません)
  7. 例: `abC8Q~xYz123456789abcdefghijklmnopqrstuv`

#### 3.4 API アクセス許可の設定

1. 左メニューから **API のアクセス許可** を選択
2. **アクセス許可の追加** をクリック
3. **Microsoft Graph** を選択
4. **アプリケーションの許可** を選択
5. 以下の権限を検索して追加:
   - ✅ `Sites.Read.All` - SharePoint サイトの読み取り
   - ✅ `Files.Read.All` - すべてのファイルの読み取り
   - (オプション) `User.Read.All` - ユーザー情報の取得 (ACL用)
6. **アクセス許可の追加** をクリック
7. ⚠️ **重要**: **{組織名} に管理者の同意を与えます** をクリック
   - グローバル管理者権限が必要
   - 同意後、状態が「✅ {組織名} に付与されました」になることを確認

#### 3.5 環境変数の設定

`.env.example` をコピーして `.env` を作成し、取得した情報を記入:

```bash
cp .env.example .env
```

`.env` ファイルを編集:

```properties
# Azure AI Search
AZURE_SEARCH_ENDPOINT=https://your-search.search.windows.net
AZURE_SEARCH_API_KEY=your-api-key
AZURE_SEARCH_INDEX_NAME=sp-docs

# Microsoft Graph (Pattern A)
GRAPH_TENANT_ID=12345678-1234-1234-1234-123456789abc  # 上記で取得したテナントID
GRAPH_CLIENT_ID=abcdefgh-1234-5678-90ab-cdefghijklmn  # 上記で取得したクライアントID
GRAPH_CLIENT_SECRET=abC8Q~xYz123456789abcdefghijk...  # 上記で取得したクライアントシークレット

# SharePoint
SP_HOSTNAME=yourtenant.sharepoint.com
SP_SITE_PATHS=/sites/your-site
SP_LIBRARY_NAMES=Documents

# Azure OpenAI (Optional)
AZURE_OPENAI_ENDPOINT=https://your-openai.openai.azure.com
AZURE_OPENAI_API_KEY=your-api-key
AZURE_OPENAI_EMBEDDING_DEPLOYMENT=text-embedding-ada-002
```

> **💡 ヒント**:
>
> - テナントIDは Azure Portal の右上のアカウント情報からも確認可能
> - クライアントシークレットの有効期限に注意(期限切れ前に更新が必要)
> - Pattern B (SharePoint Indexer) を使用する場合、Graph API の設定は不要です

### 4. Jupyter Notebook の起動

```bash
jupyter notebook notebooks/
```

## ノートブック構成

### 共通
- **`00_prereqs_and_env.ipynb`**: 環境準備と設定確認

### Pattern A: Graph API 取り込み
- **`A1_graph_ingest_setup.ipynb`**: 認証とサイト探索
- **`A2_graph_ingest_run.ipynb`**: 実取り込み処理
- **`A3_graph_delta_sync.ipynb`**: 差分同期
- **`A4_query_and_acl_demo.ipynb`**: 検索とACLデモ

### Pattern B: SharePoint Indexer
- **`B1_sp_indexer_setup.ipynb`**: インデクサ構成と初回実行
- **`B2_indexer_monitor_troubleshoot.ipynb`**: 監視とトラブルシューティング
- **`B3_query_and_acl_demo.ipynb`**: 検索とACLデモ

### クリーンアップ
- **`C_cleanup.ipynb`**: リソース削除

## 学習の進め方

### 初めての方
1. `00_prereqs_and_env.ipynb` で環境を確認
2. Pattern A または Pattern B のどちらかを選択
3. 順番にノートブックを実行

### Pattern A (カスタマイズ重視)
- カスタマイズ可能な取り込みロジック
- 柔軟なデータ変換
- 複雑な要件に対応可能

### Pattern B (運用重視)
- 設定だけで動作
- 自動スケジュール実行
- 保守が容易

## トラブルシューティング

### 権限エラー
- Azure AD アプリの権限を確認
- 管理者の同意が完了しているか確認

### サイトが見つからない
- `SP_SITE_PATHS` の値を確認
- SharePoint 上でサイトが存在するか確認

### API スロットリング (429 エラー)
- リトライロジックが自動で動作
- 処理対象ファイル数を減らす

## コスト管理

- **Azure AI Search**: インデックスサイズとクエリ数
- **Azure OpenAI**: 埋め込み生成のトークン数
- **Document Intelligence**: ページ数と API 呼び出し

最小構成で検証後、本番環境に展開することを推奨します。

## セキュリティ

⚠️ **重要な注意事項**:
- `.env` ファイルは Git にコミットしない
- 本番環境では Azure Key Vault を使用
- ログに機密情報を出力しない
- ACL フィルタを適切に実装

## 貢献

Issue や Pull Request を歓迎します。

## ライセンス

MIT License

## リソース

- [Azure AI Search ドキュメント](https://learn.microsoft.com/azure/search/)
- [Microsoft Graph API](https://learn.microsoft.com/graph/)
- [Azure OpenAI](https://learn.microsoft.com/azure/ai-services/openai/)

## サポート

質問や問題がある場合は、GitHub Issues でお知らせください。