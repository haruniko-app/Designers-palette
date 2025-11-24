# Slide AI Tool

[日本語](#japanese) | [English](#english)

---

<a name="japanese"></a>
## 🇯🇵 日本語

### 概要

Slide AI Toolは、Google Slidesで使用できるAI搭載の画像編集アドオンです。Vertex AI Imagenを活用し、画像の拡張、高画質化、各種エフェクトの適用を簡単に行うことができます。

### 主な機能

#### 🎨 AI画像拡張
- 選択した画像を上下左右に自動的に拡張
- AIが周囲のコンテキストを理解し、自然な延長を生成
- シームレスなブレンディングで違和感のない仕上がり

#### 🔍 AI高画質化
- 画像の解像度を2倍にアップスケール
- AIによる画質向上で鮮明な画像に

#### 🎭 豊富なエフェクト
- **色調整**: 明るさ、コントラスト、彩度、色相、RGB調整
- **効果**: ぼかし、モザイク、グレースケール、セピア、ビネット
- リアルタイムプレビューで確認しながら調整可能

#### 💾 ダウンロード機能
- 編集した画像をPNG形式でダウンロード
- タイムスタンプ付きのファイル名で自動保存

### 使用方法

1. Google Slidesでプレゼンテーションを開く
2. 編集したい画像を選択
3. メニューから「アドオン」→「Slide AI Tool」を選択
4. サイドバーで各種編集機能を使用
5. 「スライドに適用」で画像を反映、または「ダウンロード」で保存

### 技術スタック

- **フロントエンド**: Google Apps Script + HTML/CSS/JavaScript
- **バックエンド**: Python 3.11 + Flask
- **AI**: Google Cloud Vertex AI (Imagen)
- **インフラ**: Google Cloud Run

### リポジトリ構成

```
slide-ai-tool/
├── Google Script/     # Apps Script コード
│   ├── Code.gs       # サーバーサイド処理
│   ├── Sidebar.html  # UI
│   └── appsscript.json
├── backend/          # Python バックエンド
│   └── main.py       # Cloud Run API
├── store-listing/    # Marketplaceアセット
└── doc/             # ドキュメント
```

### サポート

問題が発生した場合や機能のリクエストがある場合は、[GitHubのIssues](https://github.com/h-abe222/slide-ai-tool/issues)で報告してください。

---

## 利用規約 (Terms of Service)

最終更新日: 2024年11月24日

### 1. サービスの提供

Slide AI Tool（以下「本サービス」）は、Google Slides用の画像編集アドオンとして提供されます。本サービスを使用することで、以下の利用規約に同意したものとみなされます。

### 2. 利用条件

- 本サービスは、個人または法人のユーザーが無償で利用できます
- Google アカウントを持つユーザーのみが利用できます
- 不正な目的での使用は禁止されています

### 3. 禁止事項

本サービスの利用において、以下の行為を禁止します：

- 違法または不適切なコンテンツの生成・処理
- 本サービスに対する不正アクセスや攻撃
- 他のユーザーの迷惑となる行為
- 著作権や知的財産権を侵害する画像の処理
- 商業目的での大量利用（API制限を超える使用）

### 4. 免責事項

- 本サービスは「現状のまま」提供され、いかなる保証も行いません
- サービスの中断、エラー、データ損失について、開発者は責任を負いません
- AI生成結果の品質や正確性について保証しません
- ユーザーが生成したコンテンツの責任はユーザー自身にあります

### 5. サービスの変更・終了

開発者は、事前の通知なく、本サービスの内容を変更または終了する権利を有します。

### 6. 準拠法

本規約は日本法に準拠し、解釈されるものとします。

---

## プライバシーポリシー (Privacy Policy)

最終更新日: 2024年11月24日

### 1. 収集する情報

Slide AI Toolは、サービス提供のために以下の情報を収集します：

#### 1.1 自動的に収集される情報
- **Googleアカウント情報**: メールアドレス、プロフィール情報（OAuth認証時）
- **使用ログ**: 機能の使用状況、エラーログ（Stackdriver経由）

#### 1.2 ユーザーが提供する情報
- **画像データ**: 編集・処理のためにアップロードされた画像
- **編集パラメータ**: 適用したエフェクトや調整値

### 2. 情報の使用目的

収集した情報は以下の目的でのみ使用されます：

- サービスの提供と機能の実行（画像処理、AI生成）
- サービスの改善と不具合の修正
- 技術的なサポートの提供

### 3. 情報の保存と保護

#### 3.1 画像データ
- アップロードされた画像は一時的に処理され、**処理完了後すぐに削除されます**
- 画像データは永続的に保存されません
- 通信はHTTPS経由で暗号化されます

#### 3.2 ログデータ
- エラーログはGoogle Cloud Stackdriverに記録されます
- ログは技術的な問題解決のためにのみ使用され、個人を特定する情報は含まれません

### 4. 第三者との情報共有

以下の場合を除き、ユーザー情報を第三者と共有することはありません：

- **Google Cloud Platform**: サービスのインフラストラクチャとして使用
- **Vertex AI**: 画像処理のために使用
- 法的要請がある場合

### 5. Cookieと追跡技術

本サービスは、Googleアカウント認証のためにOAuth 2.0トークンを使用しますが、マーケティングや追跡目的のCookieは使用しません。

### 6. OAuth権限スコープ

本サービスは以下のGoogle APIスコープを要求します：

- `https://www.googleapis.com/auth/presentations.currentonly` - 現在のプレゼンテーションへのアクセス
- `https://www.googleapis.com/auth/presentations` - プレゼンテーションの編集

これらの権限は画像の取得と更新にのみ使用され、他の用途には使用されません。

### 7. データの削除

ユーザーは以下の方法でデータを削除できます：

- アドオンをアンインストールすることで、アクセス権限を取り消せます
- [Googleアカウントの権限設定](https://myaccount.google.com/permissions)から、本サービスへのアクセスを削除できます

### 8. 子供のプライバシー

本サービスは13歳未満の子供を対象としていません。13歳未満の子供の個人情報を故意に収集することはありません。

### 9. プライバシーポリシーの変更

本ポリシーは予告なく変更される場合があります。重要な変更がある場合は、本ページで通知します。

### 10. お問い合わせ

プライバシーに関する質問や懸念がある場合は、[GitHubのIssues](https://github.com/h-abe222/slide-ai-tool/issues)でお問い合わせください。

---

<a name="english"></a>
## 🇺🇸 English

### Overview

Slide AI Tool is an AI-powered image editing add-on for Google Slides. Leveraging Vertex AI Imagen, it enables easy image extension, upscaling, and various effect applications.

### Key Features

#### 🎨 AI Image Extension
- Automatically extend selected images in all directions (top, bottom, left, right)
- AI understands surrounding context to generate natural extensions
- Seamless blending for natural-looking results

#### 🔍 AI Upscaling
- Upscale image resolution by 2x
- AI-powered enhancement for sharper images

#### 🎭 Rich Effects
- **Color Adjustments**: Brightness, Contrast, Saturation, Hue, RGB adjustment
- **Effects**: Blur, Pixelate, Grayscale, Sepia, Vignette
- Real-time preview while adjusting

#### 💾 Download Feature
- Download edited images as PNG
- Auto-save with timestamped filenames

### How to Use

1. Open a presentation in Google Slides
2. Select an image you want to edit
3. Choose "Add-ons" → "Slide AI Tool" from the menu
4. Use various editing features in the sidebar
5. Click "Apply to Slide" to update the image, or "Download" to save

### Tech Stack

- **Frontend**: Google Apps Script + HTML/CSS/JavaScript
- **Backend**: Python 3.11 + Flask
- **AI**: Google Cloud Vertex AI (Imagen)
- **Infrastructure**: Google Cloud Run

### Repository Structure

```
slide-ai-tool/
├── Google Script/     # Apps Script code
│   ├── Code.gs       # Server-side logic
│   ├── Sidebar.html  # User interface
│   └── appsscript.json
├── backend/          # Python backend
│   └── main.py       # Cloud Run API
├── store-listing/    # Marketplace assets
└── doc/             # Documentation
```

### Support

If you encounter issues or have feature requests, please report them on [GitHub Issues](https://github.com/h-abe222/slide-ai-tool/issues).

---

## Terms of Service

Last Updated: November 24, 2024

### 1. Service Provision

Slide AI Tool (hereinafter "the Service") is provided as an image editing add-on for Google Slides. By using the Service, you agree to these Terms of Service.

### 2. Terms of Use

- The Service is available free of charge to individual or corporate users
- Only users with a Google Account can use the Service
- Use for unlawful purposes is prohibited

### 3. Prohibited Actions

The following actions are prohibited when using the Service:

- Generating or processing illegal or inappropriate content
- Unauthorized access or attacks against the Service
- Actions that disturb other users
- Processing images that infringe copyright or intellectual property rights
- High-volume commercial use exceeding API limits

### 4. Disclaimer

- The Service is provided "as is" without any warranties
- The developer is not liable for service interruptions, errors, or data loss
- The quality or accuracy of AI-generated results is not guaranteed
- Users are responsible for content they generate

### 5. Service Modifications and Termination

The developer reserves the right to modify or terminate the Service without prior notice.

### 6. Governing Law

These terms shall be governed by and construed in accordance with the laws of Japan.

---

## Privacy Policy

Last Updated: November 24, 2024

### 1. Information We Collect

Slide AI Tool collects the following information to provide the Service:

#### 1.1 Automatically Collected Information
- **Google Account Information**: Email address, profile information (during OAuth authentication)
- **Usage Logs**: Feature usage, error logs (via Stackdriver)

#### 1.2 User-Provided Information
- **Image Data**: Images uploaded for editing and processing
- **Edit Parameters**: Applied effects and adjustment values

### 2. Purpose of Information Use

Collected information is used solely for:

- Providing the Service and executing features (image processing, AI generation)
- Service improvement and bug fixes
- Providing technical support

### 3. Information Storage and Protection

#### 3.1 Image Data
- Uploaded images are processed temporarily and **deleted immediately after processing**
- Image data is not stored persistently
- Communications are encrypted via HTTPS

#### 3.2 Log Data
- Error logs are recorded in Google Cloud Stackdriver
- Logs are used only for technical troubleshooting and do not contain personally identifiable information

### 4. Information Sharing with Third Parties

We do not share user information with third parties except in the following cases:

- **Google Cloud Platform**: Used as service infrastructure
- **Vertex AI**: Used for image processing
- When legally required

### 5. Cookies and Tracking Technologies

The Service uses OAuth 2.0 tokens for Google Account authentication but does not use cookies for marketing or tracking purposes.

### 6. OAuth Permission Scopes

The Service requests the following Google API scopes:

- `https://www.googleapis.com/auth/presentations.currentonly` - Access to current presentation
- `https://www.googleapis.com/auth/presentations` - Edit presentations

These permissions are used only for retrieving and updating images, not for other purposes.

### 7. Data Deletion

Users can delete their data by:

- Uninstalling the add-on to revoke access permissions
- Removing Service access from [Google Account Permissions](https://myaccount.google.com/permissions)

### 8. Children's Privacy

The Service is not intended for children under 13. We do not knowingly collect personal information from children under 13.

### 9. Privacy Policy Changes

This policy may be updated without notice. Significant changes will be announced on this page.

### 10. Contact

For privacy-related questions or concerns, please contact us via [GitHub Issues](https://github.com/h-abe222/slide-ai-tool/issues).

---

## License

This project is provided for personal and educational use. Commercial use requires prior permission.

## Developer

Developed by h-abe222
- GitHub: [@h-abe222](https://github.com/h-abe222)
- Email: h-abe@haruniko.co.jp

---

© 2024 Slide AI Tool. All rights reserved.
