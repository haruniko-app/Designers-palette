# Google Workspace Marketplace 審査リジェクト対応

**日付**: 2025-11-28
**ステータス**: 対応中
**プロジェクト番号**: 1087458701260

---

## 審査結果サマリー

3つの主要な問題が指摘されました：

1. **アプリ名の不一致** - Marketplace と OAuth 同意画面で名前が異なる
2. **OAuth スコープの不一致** - Apps Script、Marketplace SDK、OAuth 同意画面でスコープが一致していない
3. **OAuth 検証が未完了** - sensitive/restricted スコープの検証が必要

---

## タスク一覧

### タスク1: アプリ名の統一 ⚠️ 優先度: 高

**問題**: アプリ名が2つの画面で異なっている
- 参照画像: https://photos.app.goo.gl/fbYEc1UW3f2QBwmF8

**対応手順**:
1. 以下の3箇所でアプリ名を完全に一致させる：
   - Marketplace SDK のアプリ名
   - OAuth 同意画面のアプリ名
   - Google Slides アドオン内のアプリ名

2. 統一する名前: `Designer's Palette for Google Slides™`

**確認場所**:
- Google Cloud Console → APIs & Services → OAuth consent screen
- Google Cloud Console → APIs & Services → Google Workspace Marketplace SDK → App Configuration

---

### タスク2: OAuth スコープの同期 ⚠️ 優先度: 高

**問題**: Apps Script、Marketplace SDK、OAuth 同意画面でスコープが一致していない
- 参照画像: https://photos.app.goo.gl/ALVserg4LErHW6Rb7

**必要なスコープ一覧** (Apps Script プロジェクトから確認):
```
https://www.googleapis.com/auth/userinfo.email (デフォルト)
https://www.googleapis.com/auth/userinfo.profile (デフォルト)
https://www.googleapis.com/auth/presentations.currentonly
https://www.googleapis.com/auth/presentations
https://www.googleapis.com/auth/script.container.ui
```

**対応手順**:

#### ステップ A: Apps Script のスコープを確認
1. https://script.google.com/ でプロジェクトを開く
2. 左メニュー「Overview」をクリック
3. 下にスクロールして「OAuth Scopes」セクションを確認
4. すべてのスコープをメモする

#### ステップ B: Marketplace SDK にスコープを追加
1. Google Cloud Console → APIs & Services → Google Workspace Marketplace SDK
2. 「App Configuration」メニューを開く
3. 「OAuth Scopes」セクションまでスクロール
4. Apps Script で確認したすべてのスコープを追加
5. 「Save」をクリック

#### ステップ C: OAuth 同意画面にスコープを追加
1. Google Cloud Console → APIs & Services → OAuth consent screen
2. 「Edit App」をクリック
3. 必須フィールドを入力して「Save and continue」
4. 「Add or Remove Scopes」をクリック
5. 右サイドバーの「Manually add scopes」セクションまでスクロール
6. 各スコープを1つずつ追加：
   - スコープURLを入力
   - 「Add to table」をクリック
   - すべてのスコープで繰り返す
7. 「Update」をクリック
8. 「Save and continue」をクリック
9. 最後まで進める

---

### タスク3: OAuth 検証の申請 ⚠️ 優先度: 高

**問題**: sensitive/restricted スコープの OAuth 検証が未完了
- 参照画像: https://photos.app.goo.gl/HcwhfdLfnVoCu3Fo6
- 「Google has not verified this app」の画面が表示される

**対応手順**:

1. **スコープの同期が完了してから** OAuth 検証を申請する

2. OAuth 検証の申請:
   - Google Cloud Console → APIs & Services → OAuth consent screen
   - 「Submit for Verification」ボタンをクリック
   - 必要な情報を入力して送信

3. 検証に必要な情報の準備:
   - アプリの説明
   - アプリがスコープを使用する理由の説明
   - プライバシーポリシーURL: https://haruniko-and-design.github.io/designers-palette/privacy-policy.html
   - ホームページURL: https://haruniko-and-design.github.io/designers-palette/

4. **重要**: ウェブサイトに Marketplace への未公開リンクがある場合は削除する
   - 「Coming soon」テキストや無効なリンクがあると OAuth 審査が遅延する可能性がある

5. 検証メールの確認:
   - `api-oauth-dev-verification-reply+[project-id]@google.com` からのメールを確認
   - 追加情報が必要な場合はそのメールに返信
   - 問い合わせ: oauth-feedback@google.com

---

## 対応フロー

```
┌─────────────────────────────────────┐
│ 1. アプリ名を統一                     │
│    (Marketplace SDK + OAuth画面)     │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 2. Apps Script のスコープを確認       │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 3. Marketplace SDK にスコープ追加    │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 4. OAuth同意画面にスコープ追加        │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 5. OAuth検証を申請                   │
│    「Submit for Verification」       │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 6. OAuth検証の承認を待つ             │
│    (Trust & Safety チーム)           │
└──────────────┬──────────────────────┘
               ↓
┌─────────────────────────────────────┐
│ 7. Marketplace に再申請              │
└─────────────────────────────────────┘
```

---

## 参考リンク

- [Marketplace Review Guidelines](https://developers.google.com/workspace/marketplace/about-app-review#areas_of_review)
- [OAuth Scopes Configuration](https://developers.google.com/apps-script/concepts/scopes)
- [Marketplace SDK OAuth Scopes](https://developers.google.com/workspace/marketplace/enable-configure-sdk#oauth_scopes)
- [OAuth Consent Screen Configuration](https://support.google.com/cloud/answer/10311615?hl=en&ref_topic=3473162)
- [Request OAuth Verification](https://developers.google.com/apps-script/guides/client-verification#requesting_verification)
- [OAuth Verification Process](https://support.google.com/cloud/answer/9110914)
- [Preparing for OAuth Verification (Blog)](https://developers.googleblog.com/2019/09/get-smart-about-preparing-your-app-for-OAuth-verfication.html)

---

## 注意事項

1. **OAuth 検証は別チーム** (Trust & Safety) が担当するため、Marketplace 審査とは別プロセス
2. OAuth 検証が完了するまで Marketplace への再申請は保留
3. ウェブサイトに Marketplace への無効なリンクがあると OAuth 審査が遅延する可能性がある
4. デフォルトスコープ (email, profile) は削除しないこと

---

## 問い合わせ先

- Marketplace Review: gwm-review@google.com
- OAuth Verification: oauth-feedback@google.com
