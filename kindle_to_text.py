#!/usr/bin/env python3
"""
kindle-to-text: Kindle Cloud Reader 自動ページめくり + スクリーンショット + OCR

使い方:
    1. ブラウザで Kindle Cloud Reader を開き、読みたい本の最初のページを表示
    2. このスクリプトを実行
    3. カウントダウン中にブラウザウィンドウに切り替え
    4. 自動でスクリーンショット → ページめくり → OCR → テキスト出力
"""

import argparse
import hashlib
import subprocess
import sys
import time
from pathlib import Path


def capture_screenshot(output_path: str, region: tuple[int, int, int, int] | None = None) -> None:
    """macOS screencapture でスクリーンショットを取得"""
    cmd = ["screencapture", "-x"]  # -x: 無音
    if region:
        x, y, w, h = region
        cmd.extend(["-R", f"{x},{y},{w},{h}"])
    cmd.append(str(output_path))
    subprocess.run(cmd, check=True)


def ocr_image(image_path: str, languages: list[str] | None = None) -> str:
    """macOS Vision framework で OCR"""
    import Vision
    from Foundation import NSURL
    from Quartz import CGImageSourceCreateWithURL, CGImageSourceCreateImageAtIndex

    if languages is None:
        languages = ["ja", "en"]

    url = NSURL.fileURLWithPath_(str(image_path))
    source = CGImageSourceCreateWithURL(url, None)
    if not source:
        raise FileNotFoundError(f"画像を読み込めません: {image_path}")

    cg_image = CGImageSourceCreateImageAtIndex(source, 0, None)
    if not cg_image:
        raise ValueError(f"画像の作成に失敗: {image_path}")

    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLanguages_(languages)
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)

    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cg_image, None)
    success, error = handler.performRequests_error_([request], None)
    if not success:
        raise RuntimeError(f"OCR失敗: {error}")

    lines = []
    for obs in request.results():
        candidates = obs.topCandidates_(1)
        if candidates:
            lines.append(candidates[0].string())
    return "\n".join(lines)


def images_match(path1: Path, path2: Path) -> bool:
    """2つの画像ファイルが同一かハッシュで比較"""
    return (
        hashlib.md5(path1.read_bytes()).hexdigest()
        == hashlib.md5(path2.read_bytes()).hexdigest()
    )


def turn_page(direction: str = "left") -> None:
    """矢印キーでページめくり（日本語縦書き=left, 英語横書き=right）"""
    import pyautogui

    pyautogui.press(direction)


def countdown(seconds: int) -> None:
    """カウントダウン表示"""
    print(f"\n⏳ {seconds}秒後に開始します。ブラウザに切り替えてください。")
    for i in range(seconds, 0, -1):
        print(f"  {i}...", flush=True)
        time.sleep(1)
    print("  開始！\n")


def run_capture(args: argparse.Namespace) -> int:
    """スクリーンショット取得 + ページめくり"""
    import pyautogui  # noqa: F401 — アクセシビリティ権限の早期チェック

    region = None
    if args.region:
        parts = args.region.split(",")
        if len(parts) != 4:
            print("Error: --region は x,y,width,height 形式で指定してください")
            return 1
        region = tuple(int(x.strip()) for x in parts)

    screenshots_dir = Path(args.screenshots_dir)
    screenshots_dir.mkdir(parents=True, exist_ok=True)

    countdown(args.countdown)

    max_pages = args.pages or 9999
    label = f"最大 {max_pages} ページ" if args.pages else "自動停止モード"
    print(f"📸 ページキャプチャ中... ({label})")

    captured = 0
    for i in range(max_pages):
        page_path = screenshots_dir / f"page_{i:04d}.png"
        capture_screenshot(str(page_path), region)
        captured += 1

        # 前ページと比較 → 同一なら最終ページ
        if i > 0:
            prev_path = screenshots_dir / f"page_{i - 1:04d}.png"
            if images_match(prev_path, page_path):
                page_path.unlink()
                captured -= 1
                print(f"  🏁 最終ページ検出 (page {captured})。停止します。")
                break

        print(f"  [{captured}/{max_pages if args.pages else '?'}] {page_path.name}", flush=True)

        if i < max_pages - 1:
            turn_page(args.direction)
            time.sleep(args.delay)

    print(f"\n✅ スクリーンショット完了: {screenshots_dir}/ ({captured} ページ)")
    return captured


def run_ocr(screenshots_dir: Path, languages: list[str], output_file: str) -> None:
    """スクリーンショットを OCR してテキスト出力"""
    image_files = sorted(screenshots_dir.glob("page_*.png"))
    if not image_files:
        print(f"Error: {screenshots_dir}/page_*.png が見つかりません")
        sys.exit(1)

    print(f"\n🔍 OCR処理中... ({len(image_files)} ページ)")
    all_text = []
    for i, path in enumerate(image_files):
        print(f"  [{i + 1}/{len(image_files)}] {path.name}", flush=True)
        text = ocr_image(str(path), languages)
        all_text.append(text)

    output = Path(output_file)
    output.write_text("\n\n---\n\n".join(all_text), encoding="utf-8")
    total_chars = sum(len(t) for t in all_text)
    print(f"\n🎉 完了！ {output} ({len(all_text)} ページ, {total_chars:,} 文字)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Kindle Cloud Reader → テキスト変換ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 本全体をキャプチャ + OCR（自動停止）
  uv run kindle_to_text.py

  # 100ページ分キャプチャ + OCR
  uv run kindle_to_text.py --pages 100

  # スクリーンショットだけ取得（OCRは後で）
  uv run kindle_to_text.py --pages 50 --skip-ocr

  # 既存スクリーンショットから OCR のみ実行
  uv run kindle_to_text.py --ocr-only

  # 画面の特定領域だけキャプチャ
  uv run kindle_to_text.py --region 200,100,1200,800
        """,
    )
    parser.add_argument("--pages", type=int, default=None, help="キャプチャするページ数 (未指定で自動停止)")
    parser.add_argument("--delay", type=float, default=1.5, help="ページめくり間の待機秒数 (default: 1.5)")
    parser.add_argument("--output", "-o", default="output.txt", help="出力テキストファイル (default: output.txt)")
    parser.add_argument("--region", type=str, default=None, help="キャプチャ領域 x,y,width,height (例: 200,100,1200,800)")
    parser.add_argument("--lang", default="ja,en", help="OCR言語 (default: ja,en)")
    parser.add_argument("--countdown", type=int, default=5, help="開始前カウントダウン秒数 (default: 5)")
    parser.add_argument("--screenshots-dir", default="screenshots", help="スクリーンショット保存先 (default: screenshots/)")
    parser.add_argument("--direction", default="left", choices=["left", "right"], help="ページめくり方向: left=日本語縦書き, right=英語横書き (default: left)")
    parser.add_argument("--skip-ocr", action="store_true", help="スクリーンショットのみ取得 (OCRスキップ)")
    parser.add_argument("--ocr-only", action="store_true", help="既存スクリーンショットからOCRのみ実行")
    args = parser.parse_args()

    languages = [lang.strip() for lang in args.lang.split(",")]
    screenshots_dir = Path(args.screenshots_dir)

    if args.ocr_only:
        run_ocr(screenshots_dir, languages, args.output)
        return

    captured = run_capture(args)
    if captured == 0:
        print("キャプチャされたページがありません。")
        return

    if not args.skip_ocr:
        run_ocr(screenshots_dir, languages, args.output)


if __name__ == "__main__":
    main()
