# Slide AI Tool - ドキュメント

このディレクトリには、Slide AI Toolのサブスクリプション機能に関する設計・実装ドキュメントが含まれています。

---

## 📚 ドキュメント一覧

### 1. [サブスクリプション設計書](./subscription-design.md)
**概要**: 月額$7のサブスクリプションモデルの詳細設計

**内容**:
- プラン設計（Free vs Pro）
- 価格設定の根拠と競合比較
- システムアーキテクチャ
- データベース設計（Firestore）
- 収益シミュレーション
- セキュリティ対策
- FAQ

**対象読者**: プロダクトマネージャー、開発者、ステークホルダー

---

### 2. [実装ロードマップ](./implementation-roadmap.md)
**概要**: 2週間でMVPを実装するための詳細スケジュール

**内容**:
- 日次タスク分解
- 技術スタック詳細
- チェックリスト
- リスク管理
- KPI追跡

**対象読者**: 開発者、プロジェクトマネージャー

---

## 🎯 プロジェクト概要

### ビジョン
Google Slidesで使える最高のAI画像編集ツールを、手頃な価格で提供する

### ミッション
プレゼン資料作成の時間を90%削減し、誰でもプロフェッショナルな資料を作れるようにする

---

## 💰 ビジネスモデル

### プラン構成

| プラン | 月額 | 主な機能 |
|-------|------|---------|
| Free | $0 | AI拡張3回、AI高画質化2回 |
| Pro | $7 | AI拡張80回、AI高画質化50回 + 追加機能 |

### 目標数値（6ヶ月後）
- ユーザー数: 2,000人
- 有料転換率: 15%
- 月次経常収益: $2,100

---

## 🏗️ アーキテクチャ概要

```
Google Slides Add-on (Frontend)
        ↓
Firebase (Auth + Database)
        ↓
Cloud Run (Backend API)
        ↓
Stripe (Payment)
```

### 主要技術スタック
- **Frontend**: Google Apps Script
- **Backend**: Python 3.11 + Cloud Run
- **Database**: Firestore
- **Authentication**: Firebase Auth
- **Payment**: Stripe
- **AI**: Google Vertex AI (Imagen)

---

## 📅 開発スケジュール

### Phase 1: MVP（Week 1-2）
- 認証・使用量トラッキング
- Stripe決済統合
- UI実装

### Phase 2: 追加機能（Week 3以降）
- バッチ処理
- 履歴保存
- 優先処理

---

## 🚀 クイックスタート

### 開発者向け

1. **リポジトリクローン**
```bash
git clone [repo-url]
cd slide-ai-tool
```

2. **環境変数設定**
```bash
cp .env.example .env.local
# .env.local を編集
```

3. **依存関係インストール**
```bash
cd backend
pip install -r requirements.txt
```

4. **Firebase セットアップ**
```bash
firebase login
firebase init
```

5. **ローカル実行**
```bash
# Backend
cd backend
functions-framework --target=generate_image --port=8080

# Frontend
clasp push
```

---

## 📖 関連ドキュメント

### 外部リンク
- [Firebase ドキュメント](https://firebase.google.com/docs)
- [Stripe ドキュメント](https://stripe.com/docs)
- [Google Apps Script ガイド](https://developers.google.com/apps-script)
- [Vertex AI Imagen](https://cloud.google.com/vertex-ai/docs/generative-ai/image/overview)

### 内部リンク
- [既存の設計書](../README.md)
- [バックエンドコード](../backend/)
- [Google Apps Script](../Google%20Script/)

---

## 🤝 コントリビューション

### 開発フロー
1. Issueを作成
2. ブランチを切る（`feature/xxx`）
3. 実装・テスト
4. Pull Request作成
5. レビュー・マージ

### コーディング規約
- Python: PEP 8
- JavaScript: Google Style Guide
- コミットメッセージ: Conventional Commits

---

## 📞 サポート・問い合わせ

### 技術的な質問
- GitHub Issues

### ビジネスに関する問い合わせ
- Email: [your-email]

### バグ報告
- GitHub Issues（テンプレート使用）

---

## 📝 ライセンス

[ライセンス情報を記載]

---

## 🙏 謝辞

- Google Cloud Platform
- Firebase
- Stripe
- オープンソースコミュニティ

---

**最終更新**: 2025-11-24
**バージョン**: 1.0.0
