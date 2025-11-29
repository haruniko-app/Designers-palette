# バージョン2 アップグレード計画

## 概要

- **現行バージョン**: v1.0.0（Google審査中）
- **目標**: 便利ツールセクションの大幅強化
- **コンセプト**: Adobe Illustratorのパネル風UIで作業性向上

---

## 現在の実装済み機能

### 画像編集
- トリミング確定
- AI画像拡張
- AI高画質化
- 色調調整
- エフェクト

### 便利ツール
- 整列（左/中央/右/上/中央/下）
- 均等配置（水平/垂直）

---

## 実装フェーズ

### フェーズ1: カラーパネル（色・塗り系）

#### 1-1. 塗り/線カラー
- [x] Illustrator風 塗り/線切り替えUI
- [x] スウォッチパレット（基本色グリッド）
- [x] 最近使用したカラー（LocalStorage）※適用時のみ履歴追加
- [x] カラースペクトラム（Canvas実装）
- [x] RGBスライダー（0-255）
- [x] HEX入力（#RRGGBB）
- [x] 「なし」選択（赤斜線表示）

#### 1-2. 枠線スタイル
- [x] 線の太さ（pt指定）※即時適用
- [x] 線のスタイル（実線/破線/点線/一点鎖線）※即時適用

#### 1-3. スポイト
- [x] 選択要素から色を取得
- [x] 取得した色を他の要素に適用

---

### フェーズ2: レイアウト強化

#### 2-1. サイズ変更
- [x] 幅・高さ数値入力
- [x] 縦横比維持オプション
- [x] 複数要素を同サイズに揃え

#### 2-2. 間隔調整
- [x] 要素間の余白を数値指定

#### 2-3. 複製
- [x] 指定方向・個数・間隔で連続複製

---

### フェーズ3: 変形・順序

#### 3-1. 回転・反転
- [x] 90°/180°回転
- [x] 水平/垂直反転

#### 3-2. 順序変更
- [x] 最前面/最背面
- [x] 1つ前/1つ後ろ

#### 3-3. グループ化
- [x] グループ化/解除

---

### フェーズ4: テキスト

#### 4-1. テキスト色
- [x] カラーパネルと連携

#### 4-2. フォント
- [x] フォント名
- [x] サイズ
- [x] 太字/斜体

#### 4-3. テキスト背景・装飾
- [x] 蛍光ペン効果（文字背景色）
- [x] 下線
- [x] 取消線

---

## フェーズ1 詳細タスク

| # | タスク | ステータス | 備考 |
|---|--------|:----------:|------|
| 1 | 塗り/線切り替えUIのHTML/CSS作成 | ✅ 完了 | プロトタイプ作成 |
| 2 | スウォッチパレット実装 | ✅ 完了 | カラーグリッド |
| 3 | カラースペクトラム（Canvas）実装 | ✅ 完了 | 色選択機能 |
| 4 | RGBスライダー + HEX入力実装 | ✅ 完了 | 数値入力機能 |
| 5 | 最近使用したカラー（LocalStorage） | ✅ 完了 | 適用時のみ履歴追加 |
| 6 | GAS側: 図形の塗り/線色取得・設定 | ✅ 完了 | Code.gs関数 |
| 7 | 枠線スタイル（太さ・種類）実装 | ✅ 完了 | 即時適用 |
| 8 | スポイト機能実装 | ✅ 完了 | 色取得・適用 |
| 9 | Sidebar.htmlへ統合 | ✅ 完了 | フェーズ1-4全セクション統合 |
| 10 | テスト・調整 | ⬜ 未着手 | リリース準備 |

### ステータス凡例
- ⬜ 未着手
- 🔄 進行中
- ✅ 完了
- ⏸️ 保留

---

## 技術仕様

### 使用API（Google Apps Script）

#### 塗りつぶし
```javascript
// 色の設定
shape.getFill().setSolidFill('#RRGGBB');
shape.getFill().setSolidFill(r, g, b);
shape.getFill().setTransparent();

// 色の取得
shape.getFill().getSolidFill().getColor();
```

#### 枠線
```javascript
// 色の設定
shape.getBorder().getLineFill().setSolidFill('#RRGGBB');

// 太さの設定
shape.getBorder().setWeight(points);

// スタイルの設定
shape.getBorder().setDashStyle(SlidesApp.DashStyle.SOLID);
```

#### テキスト
```javascript
// 文字色
textRange.getTextStyle().setForegroundColor('#RRGGBB');

// フォント
textRange.getTextStyle().setFontFamily('Arial');
textRange.getTextStyle().setFontSize(12);
textRange.getTextStyle().setBold(true);
```

### UI実装

#### カラースペクトラム
- Canvas 2D APIで実装
- HSV/HSLカラーモデル使用
- マウスドラッグで色選択

#### スウォッチ
- CSS Gridでタイル配置
- 基本色 + グレースケール + カスタム色

#### 最近使用したカラー
- LocalStorageに最大8色保存
- 新しい色を使用時に自動追加

---

## 参考資料

- [TextStyle Class](https://developers.google.com/apps-script/reference/slides/text-style)
- [Fill Class](https://developers.google.com/apps-script/reference/slides/fill)
- [Border Class](https://developers.google.com/apps-script/reference/slides/border)
- [Shape Class](https://developers.google.com/apps-script/reference/slides/shape)

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2025-11-25 | 計画策定・ドキュメント作成 |
| 2025-11-25 | フェーズ1 プロトタイプ完成（dev/color-panel-prototype.html） |

---

## 成果物

### プロトタイプ
- `dev/color-panel-prototype.html` - カラーパネル・枠線セクションのUI

### GAS関数（Code.gs追加分）
- `getSelectedElementColors()` - 選択要素の色情報を取得
- `setFillColor(hexColor)` - 塗りつぶし色を設定
- `setStrokeColor(hexColor)` - 枠線色を設定
- `setStrokeStyle(weight, dashStyle)` - 枠線スタイルを設定
- `checkElementForColor()` - 要素タイプ確認（軽量）
- `rgbColorToHex(rgbColor)` - RgbColor→HEX変換
