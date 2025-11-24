# 実装ロードマップ

**プロジェクト**: Slide AI Tool - サブスクリプション機能
**目標**: 月額$7のProプラン実装
**期間**: 2週間（MVP）

---

## Week 1: MVP実装

### Day 1-3: 認証・使用量トラッキング

#### Day 1: Firebase セットアップ
- [ ] Firebase プロジェクト作成
  - プロジェクト名: `slide-ai-tool`
  - リージョン: `us-central1`
- [ ] Firebase Authentication 有効化
  - Google プロバイダー設定
  - 認証ドメイン設定
- [ ] Firestore Database 作成
  - ネイティブモード
  - セキュリティルール設定
- [ ] Firebase SDK を Google Apps Script に統合
  - FirebaseApp ライブラリ追加
  - 設定情報の追加

**成果物**:
- Firebase プロジェクト
- 認証設定完了

---

#### Day 2: 認証UI実装
- [ ] Code.gs に認証関数追加
  ```javascript
  function signInWithGoogle() {
    // Firebase Auth実装
  }

  function signOut() {
    // ログアウト実装
  }

  function getCurrentUser() {
    // 現在のユーザー取得
  }
  ```
- [ ] Slidebar_fix.html にログインUI追加
  - ログインボタン
  - ユーザー情報表示
  - ログアウトボタン
- [ ] ログイン状態の保持
  - Session Storage 使用

**成果物**:
- ログイン/ログアウト機能
- ユーザー状態管理

---

#### Day 3: 使用量トラッキング実装
- [ ] Firestore にユーザードキュメント作成
  ```javascript
  // 初回ログイン時
  users/{userId} = {
    email: string,
    planType: "free",
    usage: {
      aiExtend: { count: 0, limit: 3, resetDate: timestamp },
      aiUpscale: { count: 0, limit: 2, resetDate: timestamp }
    },
    createdAt: timestamp
  }
  ```
- [ ] backend/main.py に使用量チェック追加
  ```python
  def check_quota(user_id, mode):
      # Firestoreから使用量確認
      # 制限チェック
      # 使用量インクリメント
  ```
- [ ] エラーハンドリング
  - 制限超過時のレスポンス
  - フロントエンドでのエラー表示

**成果物**:
- 使用量トラッキングシステム
- 制限チェック機能

---

### Day 4-5: Stripe決済統合

#### Day 4: Stripe セットアップ
- [ ] Stripe アカウント作成
  - テストモード有効化
- [ ] Product & Price 作成
  - Product名: "Slide AI Tool Pro"
  - Price: $7/月（recurring）
  - 通貨: USD
- [ ] Stripe Checkout Session 作成
  ```python
  # backend に追加
  @functions_framework.http
  def create_checkout_session(request):
      stripe.checkout.Session.create(
          customer_email=email,
          line_items=[{
              'price': 'price_xxx',
              'quantity': 1,
          }],
          mode='subscription',
          success_url=success_url,
          cancel_url=cancel_url,
      )
  ```
- [ ] Code.gs にアップグレード関数追加
  ```javascript
  function upgradeToProPlan() {
    // Checkout Session作成
    // URLを開く
  }
  ```

**成果物**:
- Stripe Checkout フロー
- アップグレードボタン

---

#### Day 5: Webhook実装
- [ ] Webhook エンドポイント作成
  ```python
  @functions_framework.http
  def stripe_webhook(request):
      # 署名検証
      event = stripe.Webhook.construct_event(
          payload, sig_header, webhook_secret
      )

      # イベント処理
      if event.type == 'customer.subscription.created':
          # Firestoreを更新
          update_user_plan(user_id, 'pro')
  ```
- [ ] Webhook URLを Stripe に登録
- [ ] テスト
  - Stripe CLI でローカルテスト
  - 実際の決済フロー確認

**成果物**:
- Webhook処理
- サブスク同期機能

---

### Day 6-7: UI実装・テスト

#### Day 6: ダッシュボード実装
- [ ] 使用状況表示UI
  ```html
  <div class="usage-dashboard">
    <div class="usage-item">
      <span>AI画像拡張</span>
      <span id="extendUsage">2 / 3回</span>
    </div>
    <div class="usage-item">
      <span>AI高画質化</span>
      <span id="upscaleUsage">1 / 2回</span>
    </div>
  </div>
  ```
- [ ] プラン表示
  - 現在のプラン（Free/Pro）
  - アップグレードボタン（Freeのみ）
- [ ] アップグレード促進UI
  - 制限超過時のモーダル
  - プラン比較表示

**成果物**:
- ユーザーダッシュボード
- アップグレードフロー完成

---

#### Day 7: 統合テスト
- [ ] E2Eテストシナリオ
  1. 新規ユーザー登録
  2. Free プランで使用
  3. 制限到達
  4. アップグレード
  5. Pro機能使用
  6. サブスクキャンセル
- [ ] バグ修正
- [ ] パフォーマンステスト
- [ ] セキュリティチェック

**成果物**:
- テスト完了
- MVP完成

---

## Week 2: リリース準備

### Day 8-10: テスト・修正

#### Day 8: ユーザーテスト
- [ ] ベータテスター募集（5-10人）
- [ ] フィードバック収集
- [ ] 改善点リストアップ

#### Day 9: バグ修正
- [ ] 優先度高いバグ修正
- [ ] UI/UX改善

#### Day 10: セキュリティ監査
- [ ] Firebase セキュリティルール確認
- [ ] API認証チェック
- [ ] 脆弱性スキャン

---

### Day 11-12: ドキュメント・マーケティング

#### Day 11: ドキュメント作成
- [ ] ユーザーガイド
  - 使い方
  - FAQ
  - トラブルシューティング
- [ ] 利用規約
- [ ] プライバシーポリシー

#### Day 12: マーケティング準備
- [ ] ランディングページ作成
- [ ] プロモーション動画
- [ ] SNS投稿準備

---

### Day 13-14: リリース

#### Day 13: ソフトローンチ
- [ ] Google Workspace Marketplace 申請準備
- [ ] 限定公開リンク配布
- [ ] モニタリング開始

#### Day 14: パブリックリリース
- [ ] 正式公開
- [ ] Product Hunt 投稿
- [ ] SNS告知
- [ ] プレスリリース

---

## Week 3以降: 追加機能開発

### Phase 4: バッチ処理（1週間）

**目標**: 複数画像の一括処理

- [ ] UI: 複数画像選択
- [ ] バックエンド: キュー処理
- [ ] 進捗表示
- [ ] エラーハンドリング

---

### Phase 5: 履歴保存（1週間）

**目標**: 処理前後の画像を30日間保存

- [ ] Cloud Storage 設定
- [ ] 画像アップロード機能
- [ ] 履歴一覧UI
- [ ] 復元機能
- [ ] 自動削除（30日後）

---

### Phase 6: 優先処理（3日）

**目標**: Proユーザーの処理を優先

- [ ] プラン別キュー
- [ ] 優先度ルーティング
- [ ] レート制限

---

## 技術スタック詳細

### フロントエンド
```
Google Apps Script
├── Code.gs (サーバーサイド処理)
│   ├── 認証関数
│   ├── Firestore連携
│   └── Stripe連携
└── Slidebar_fix.html (UI)
    ├── ログインUI
    ├── ダッシュボード
    └── アップグレードフロー
```

### バックエンド
```
Cloud Run (Python)
├── main.py
│   ├── generate_image (既存)
│   ├── check_quota (新規)
│   ├── create_checkout_session (新規)
│   └── stripe_webhook (新規)
├── requirements.txt
│   ├── firebase-admin
│   ├── stripe
│   └── 既存パッケージ
└── Procfile
```

### データベース
```
Firestore
├── users/{userId}
│   ├── email
│   ├── planType
│   ├── subscriptionId
│   └── usage
├── subscriptionEvents/{eventId}
│   └── 監査ログ
└── history/{userId}/images/{imageId}
    └── Pro限定履歴
```

---

## チェックリスト

### 開発環境
- [ ] Firebase CLI インストール
- [ ] Stripe CLI インストール
- [ ] Python 3.11 環境
- [ ] Google Cloud SDK

### アカウント・サービス
- [ ] Firebase プロジェクト
- [ ] Stripe アカウント
- [ ] Google Cloud Project
- [ ] （オプション）Sendgrid

### デプロイ
- [ ] Cloud Run デプロイ
- [ ] Firebase デプロイ
- [ ] 環境変数設定
- [ ] Webhook URL設定

### テスト
- [ ] ローカルテスト
- [ ] ステージングテスト
- [ ] 本番テスト
- [ ] パフォーマンステスト

### ドキュメント
- [ ] README更新
- [ ] API仕様書
- [ ] ユーザーガイド
- [ ] FAQ

### マーケティング
- [ ] ランディングページ
- [ ] プロモーション動画
- [ ] SNS投稿
- [ ] Product Hunt

---

## リスク管理

### 技術的リスク

| リスク | 影響度 | 対策 |
|-------|-------|------|
| Firebase 障害 | 高 | フォールバック実装 |
| Stripe 障害 | 中 | 手動処理の準備 |
| API制限超過 | 中 | キャッシュ実装 |
| 不正利用 | 高 | レート制限、監視 |

### ビジネスリスク

| リスク | 影響度 | 対策 |
|-------|-------|------|
| 転換率が低い | 高 | 価格見直し、機能追加 |
| チャーン率が高い | 高 | ユーザーフィードバック収集 |
| コスト増加 | 中 | 使用量監視、最適化 |

---

## KPI追跡

### デイリー
- [ ] 新規登録数
- [ ] アクティブユーザー数
- [ ] エラー率

### ウィークリー
- [ ] 有料転換率
- [ ] 使用量平均
- [ ] サポート問い合わせ数

### マンスリー
- [ ] MRR（月次経常収益）
- [ ] チャーンレート
- [ ] NPS
- [ ] 機能利用率

---

## サポート体制

### ユーザーサポート
- **チャネル**: メール
- **対応時間**: 24時間以内
- **FAQ**: オンラインドキュメント

### 開発者サポート
- **ドキュメント**: GitHub Wiki
- **問い合わせ**: GitHub Issues

---

## 次のマイルストーン

### 3ヶ月後
- [ ] 500ユーザー達成
- [ ] 有料転換率 10%
- [ ] バッチ処理実装

### 6ヶ月後
- [ ] 2,000ユーザー達成
- [ ] 有料転換率 15%
- [ ] 履歴保存実装

### 1年後
- [ ] 10,000ユーザー達成
- [ ] 有料転換率 20%
- [ ] API公開

---

**End of Roadmap**
