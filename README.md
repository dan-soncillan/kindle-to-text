# kindle-to-text

Kindle Cloud Reader の本を自動ページめくり + スクリーンショット + OCR でテキスト化するmacOSツール。

## 仕組み

1. ブラウザで開いた Kindle Cloud Reader のページを `screencapture` でキャプチャ
2. 右矢印キーで自動ページめくり
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

### 基本（全自動）

```bash
# 1. ブラウザで Kindle Cloud Reader を開き、本の最初のページを表示
# 2. 実行（ページ末尾を自動検出して停止）
uv run kindle_to_text.py
```

カウントダウン中（5秒）にブラウザウィンドウをクリックしてフォーカスを合わせてください。

### オプション

```bash
# ページ数を指定
uv run kindle_to_text.py --pages 100

# 特定の画面領域だけキャプチャ（ブラウザのUIを除外）
uv run kindle_to_text.py --region 200,100,1200,800

# ページめくり速度を調整（デフォルト: 1.5秒）
uv run kindle_to_text.py --delay 2.0

# 出力ファイル名を指定
uv run kindle_to_text.py -o my_book.txt

# スクリーンショットだけ取得（OCRは後で）
uv run kindle_to_text.py --skip-ocr

# 既存のスクリーンショットからOCRだけ実行
uv run kindle_to_text.py --ocr-only

# 英語の本
uv run kindle_to_text.py --lang en
```

### 全オプション

| オプション | デフォルト | 説明 |
|-----------|-----------|------|
| `--pages` | 自動停止 | キャプチャするページ数 |
| `--delay` | 1.5 | ページめくり間の待機秒数 |
| `-o, --output` | output.txt | 出力ファイル |
| `--region` | 全画面 | キャプチャ領域 (x,y,width,height) |
| `--lang` | ja,en | OCR言語 |
| `--countdown` | 5 | 開始前カウントダウン秒数 |
| `--screenshots-dir` | screenshots/ | スクリーンショット保存先 |
| `--skip-ocr` | false | OCRをスキップ |
| `--ocr-only` | false | 既存スクショからOCRのみ |

## Tips

- **`--region` の決め方**: まず `screencapture -i test.png` で領域選択スクリーンショットを撮り、プレビューで座標を確認
- **ページめくりが速すぎる場合**: `--delay 2.0` 以上に設定
- **OCR精度が低い場合**: `--region` でテキスト領域だけをキャプチャすると精度が上がる
- **大きな本**: `--skip-ocr` でまずスクショだけ取得し、`--ocr-only` で後からOCRすると安全

## 要件

- macOS (screencapture + Vision framework を使用)
- Python 3.11+
- アクセシビリティ権限 + 画面収録権限

## License

MIT
