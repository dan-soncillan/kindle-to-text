# kindle-to-text

Kindle の本を自動ページめくり + スクリーンショット + OCR でテキスト化する macOS ツール。

Kindle アプリ、Kindle Cloud Reader（ブラウザ）、その他の電子書籍リーダーに対応。

## 仕組み

1. 指定したウィンドウを `screencapture -l` でキャプチャ（フォーカス不要）
2. AppleScript で対象アプリに右矢印キーを送信（自動ページめくり）
3. macOS Vision framework で日本語/英語 OCR
4. 全ページのテキストを1つのファイルに結合

## セットアップ

```bash
# uv がなければインストール
curl -LsSf https://astral.sh/uv/install.sh | sh

# 依存関係インストール
cd kindle-to-text
uv sync
```

### macOS 権限設定

初回実行時に以下の権限を許可する必要があります：

- **アクセシビリティ** (システム設定 → プライバシーとセキュリティ → アクセシビリティ): ターミナルアプリを追加
- **画面収録** (システム設定 → プライバシーとセキュリティ → 画面収録とシステム音声の録音): ターミナルアプリを追加

## 使い方

### GUI アプリ（推奨）

```bash
uv run app.py
```

1. ドロップダウンからキャプチャ対象のウィンドウを選択
2. ページ数・ディレイ等を設定
3. 「▶ 開始」をクリック
4. 自動でスクショ → ページめくり → OCR → テキスト出力

フォーカスを変えても動作します。

### CLI

```bash
# 自動停止モード
uv run kindle_to_text.py

# ページ数指定
uv run kindle_to_text.py --pages 100

# スクリーンショットだけ取得（OCRは後で）
uv run kindle_to_text.py --skip-ocr

# 既存スクリーンショットからOCRのみ
uv run kindle_to_text.py --ocr-only
```

## Tips

- **ページめくりが速すぎる場合**: ディレイを 2.0 秒以上に設定
- **OCR精度が低い場合**: ウィンドウからツールバー等を非表示にして本文領域を最大化
- **大きな本**: 「OCRのみ」ボタンで後からOCRし直せる

## 要件

- macOS (screencapture + Vision framework 使用)
- Python 3.11+
- アクセシビリティ権限 + 画面収録権限

## License

MIT
