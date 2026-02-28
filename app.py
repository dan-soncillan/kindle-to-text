#!/usr/bin/env python3
"""kindle-to-text GUI App: ウィンドウ指定キャプチャ + 自動ページめくり + OCR"""

import hashlib
import subprocess
import threading
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from pathlib import Path

from Quartz import (
    CGWindowListCopyWindowInfo,
    kCGWindowListOptionOnScreenOnly,
    kCGNullWindowID,
    CGEventCreateMouseEvent,
    CGEventCreateKeyboardEvent,
    CGEventPost,
    kCGEventLeftMouseDown,
    kCGEventLeftMouseUp,
    kCGHIDEventTap,
    CGImageSourceCreateWithURL,
    CGImageSourceCreateImageAtIndex,
    CGImageCreateWithImageInRect,
    CGImageGetWidth,
    CGImageGetHeight,
    CGRectMake,
    CGImageDestinationCreateWithURL,
    CGImageDestinationAddImage,
    CGImageDestinationFinalize,
)


# ---------------------------------------------------------------------------
# Core functions
# ---------------------------------------------------------------------------

def _get_app_window_titles(app_name: str) -> list[str]:
    """AppleScript でアプリのウィンドウタイトルを取得"""
    script = f'tell application "{app_name}" to return title of every window'
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode == 0 and result.stdout.strip():
        return [t.strip() for t in result.stdout.strip().split(", ")]
    return []


def get_visible_windows() -> list[dict]:
    """画面上の可視ウィンドウ一覧を取得"""
    windows = CGWindowListCopyWindowInfo(
        kCGWindowListOptionOnScreenOnly, kCGNullWindowID
    )
    result = []
    seen = set()
    owner_index: dict[str, int] = {}  # タイトルなしウィンドウのインデックス追跡
    title_cache: dict[str, list[str]] = {}  # AppleScript で取得したタイトルキャッシュ
    skip_owners = {"Window Server", "Dock", "SystemUIServer", "Control Center", "Spotlight"}

    for w in windows:
        name = w.get("kCGWindowName", "") or ""
        owner = w.get("kCGWindowOwnerName", "") or ""
        wid = w.get("kCGWindowNumber", 0)
        layer = w.get("kCGWindowLayer", 0)
        bounds = w.get("kCGWindowBounds", {})
        width = int(bounds.get("Width", 0))
        height = int(bounds.get("Height", 0))

        if layer == 0 and owner and owner not in skip_owners and width > 100 and height > 100 and wid not in seen:
            seen.add(wid)

            # タイトルがない場合、AppleScript で取得を試みる
            if not name:
                if owner not in title_cache:
                    title_cache[owner] = _get_app_window_titles(owner)
                idx = owner_index.get(owner, 0)
                titles = title_cache[owner]
                if idx < len(titles):
                    name = titles[idx]
                owner_index[owner] = idx + 1

            label = f"{owner} — {name}" if name else owner
            result.append({
                "id": wid,
                "name": name,
                "owner": owner,
                "bounds": bounds,
                "label": label,
            })
    return result


def capture_window(window_id: int, output_path: str) -> None:
    """特定ウィンドウをキャプチャ（フォーカス不要）"""
    subprocess.run(
        ["screencapture", "-x", "-o", "-l", str(window_id), output_path],
        check=True,
    )


def click_at(x: float, y: float) -> None:
    """Quartz CGEvent でクリック（pyautogui 不要）"""
    point = (x, y)
    event_down = CGEventCreateMouseEvent(None, kCGEventLeftMouseDown, point, 0)
    event_up = CGEventCreateMouseEvent(None, kCGEventLeftMouseUp, point, 0)
    CGEventPost(kCGHIDEventTap, event_down)
    time.sleep(0.05)
    CGEventPost(kCGHIDEventTap, event_up)


def _is_browser(app_name: str) -> bool:
    """Chrome 系ブラウザかどうか判定"""
    return any(k in app_name.lower() for k in ("chrome", "brave", "edge", "chromium"))


def activate_window(app_name: str, window_title: str = "") -> None:
    """アプリを前面に持ってくる。Chrome 系はタブタイトルで検索してアクティブにする。"""
    if window_title and _is_browser(app_name):
        # Chrome 系: タブタイトルで検索して正しいタブ＋ウィンドウをアクティベート
        escaped_title = window_title.replace('"', '\\"')
        script = f'''
        tell application "{app_name}"
            activate
            repeat with w in windows
                set tabList to tabs of w
                repeat with i from 1 to count of tabList
                    if title of item i of tabList contains "{escaped_title}" then
                        set active tab index of w to i
                        set index of w to 1
                        return
                    end if
                end repeat
            end repeat
        end tell
        '''
    else:
        script = f'tell application "System Events" to set frontmost of process "{app_name}" to true'
    subprocess.run(["osascript", "-e", script], capture_output=True)
    time.sleep(0.5)


def turn_page_by_key(direction: str = "left") -> None:
    """CGEvent で矢印キーを送信してページめくり（オーバーレイに影響されない）"""
    # Left arrow: keycode 123, Right arrow: keycode 124
    keycode = 123 if direction == "left" else 124
    event_down = CGEventCreateKeyboardEvent(None, keycode, True)
    event_up = CGEventCreateKeyboardEvent(None, keycode, False)
    CGEventPost(kCGHIDEventTap, event_down)
    time.sleep(0.05)
    CGEventPost(kCGHIDEventTap, event_up)


def ocr_image(image_path: str, languages: list[str] | None = None, invert: bool = False) -> str:
    """macOS Vision framework で OCR（ダークモード自動対応）"""
    import Vision
    from Foundation import NSURL
    from Quartz import (
        CGImageSourceCreateWithURL,
        CGImageSourceCreateImageAtIndex,
        CIImage,
        CIFilter,
        CIContext,
    )

    if languages is None:
        languages = ["ja", "en"]

    url = NSURL.fileURLWithPath_(str(image_path))
    source = CGImageSourceCreateWithURL(url, None)
    if not source:
        raise FileNotFoundError(f"画像を読み込めません: {image_path}")

    cg_image = CGImageSourceCreateImageAtIndex(source, 0, None)
    if not cg_image:
        raise ValueError(f"画像の作成に失敗: {image_path}")

    def _run_ocr(img):
        """Vision OCR を実行して文字列を返す"""
        if img is None:
            return ""
        request = Vision.VNRecognizeTextRequest.alloc().init()
        request.setRecognitionLanguages_(languages)
        request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
        request.setUsesLanguageCorrection_(True)
        try:
            request.setRevision_(3)
        except Exception:
            pass
        handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(
            img, None
        )
        success, error = handler.performRequests_error_([request], None)
        if not success:
            return ""
        lines = []
        for obs in request.results():
            candidates = obs.topCandidates_(1)
            if candidates:
                lines.append(candidates[0].string())
        return "\n".join(lines)

    def _preprocess_dark(img):
        """ダークモード前処理: グレースケール → 反転 → コントラスト強化"""
        try:
            ci = CIImage.imageWithCGImage_(img)
            if not ci:
                return None

            # グレースケール化（色味を除去して反転精度を上げる）
            mono = CIFilter.filterWithName_("CIPhotoEffectMono")
            if mono:
                mono.setDefaults()
                mono.setValue_forKey_(ci, "inputImage")
                ci = mono.outputImage() or ci

            # 色反転
            inv = CIFilter.filterWithName_("CIColorInvert")
            if inv:
                inv.setDefaults()
                inv.setValue_forKey_(ci, "inputImage")
                ci = inv.outputImage() or ci

            # コントラスト強化（背景を白く、文字を黒くする）
            ctrl = CIFilter.filterWithName_("CIColorControls")
            if ctrl:
                ctrl.setDefaults()
                ctrl.setValue_forKey_(ci, "inputImage")
                ctrl.setValue_forKey_(2.0, "inputContrast")
                ctrl.setValue_forKey_(0.1, "inputBrightness")
                ci = ctrl.outputImage() or ci

            context = CIContext.contextWithOptions_(None)
            result = context.createCGImage_fromRect_(ci, ci.extent())
            return result
        except Exception:
            return None

    if invert:
        # 元画像と前処理画像の両方でOCRし、長い方を採用（白/ダーク混在対応）
        text_normal = _run_ocr(cg_image)
        processed = _preprocess_dark(cg_image)
        text_dark = _run_ocr(processed)
        return text_dark if len(text_dark) > len(text_normal) else text_normal

    return _run_ocr(cg_image)


def crop_image(image_path: str, crop_top: int = 0, crop_bottom: int = 0) -> None:
    """画像の上下をクロップ（ブラウザUI除去用）"""
    if crop_top <= 0 and crop_bottom <= 0:
        return
    from Foundation import NSURL

    url = NSURL.fileURLWithPath_(image_path)
    source = CGImageSourceCreateWithURL(url, None)
    if not source:
        return
    cg_image = CGImageSourceCreateImageAtIndex(source, 0, None)
    if not cg_image:
        return

    w = CGImageGetWidth(cg_image)
    h = CGImageGetHeight(cg_image)
    new_h = h - crop_top - crop_bottom
    if new_h <= 0:
        return

    rect = CGRectMake(0, crop_top, w, new_h)
    cropped = CGImageCreateWithImageInRect(cg_image, rect)

    dest = CGImageDestinationCreateWithURL(url, "public.png", 1, None)
    if dest:
        CGImageDestinationAddImage(dest, cropped, None)
        CGImageDestinationFinalize(dest)


def images_match(path1: Path, path2: Path) -> bool:
    """2つの画像が同一かハッシュで比較"""
    return (
        hashlib.md5(path1.read_bytes()).hexdigest()
        == hashlib.md5(path2.read_bytes()).hexdigest()
    )


# ---------------------------------------------------------------------------
# GUI App
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Kindle to Text")
        self.root.geometry("680x630")

        self.running = False
        self.stop_event = threading.Event()
        self.windows: list[dict] = []

        self._build_ui()
        self.refresh_windows()

    def _build_ui(self):
        # --- Settings ---
        settings = ttk.LabelFrame(self.root, text="設定", padding=10)
        settings.pack(fill="x", padx=10, pady=(10, 5))
        settings.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(settings, text="ウィンドウ:").grid(row=row, column=0, sticky="w", pady=3)
        self.window_var = tk.StringVar()
        self.window_combo = ttk.Combobox(
            settings, textvariable=self.window_var, state="readonly"
        )
        self.window_combo.grid(row=row, column=1, sticky="ew", padx=5, pady=3)
        self.window_combo.bind("<<ComboboxSelected>>", self._on_window_selected)
        ttk.Button(settings, text="更新", width=6, command=self.refresh_windows).grid(
            row=row, column=2, pady=3
        )

        row += 1
        ttk.Label(settings, text="ページ数:").grid(row=row, column=0, sticky="w", pady=3)
        pages_frame = ttk.Frame(settings)
        pages_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5, pady=3)
        self.pages_var = tk.StringVar(value="")
        ttk.Entry(pages_frame, textvariable=self.pages_var, width=10).pack(side="left")
        ttk.Label(pages_frame, text="空欄 = 自動停止", foreground="gray").pack(
            side="left", padx=10
        )

        row += 1
        ttk.Label(settings, text="めくり方向:").grid(row=row, column=0, sticky="w", pady=3)
        self.direction_var = tk.StringVar(value="← 左 (日本語/縦書き)")
        direction_combo = ttk.Combobox(
            settings, textvariable=self.direction_var, state="readonly", width=25,
            values=["← 左 (日本語/縦書き)", "→ 右 (英語/横書き)"],
        )
        direction_combo.grid(row=row, column=1, sticky="w", padx=5, pady=3)

        row += 1
        ttk.Label(settings, text="クロップ(px):").grid(row=row, column=0, sticky="w", pady=3)
        crop_frame = ttk.Frame(settings)
        crop_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5, pady=3)
        ttk.Label(crop_frame, text="上:").pack(side="left")
        self.crop_top_var = tk.StringVar(value="0")
        ttk.Entry(crop_frame, textvariable=self.crop_top_var, width=6).pack(side="left")
        ttk.Label(crop_frame, text="  下:").pack(side="left")
        self.crop_bottom_var = tk.StringVar(value="0")
        ttk.Entry(crop_frame, textvariable=self.crop_bottom_var, width=6).pack(side="left")
        ttk.Label(crop_frame, text="ブラウザUI除去用", foreground="gray").pack(
            side="left", padx=10
        )

        row += 1
        ttk.Label(settings, text="ディレイ(秒):").grid(row=row, column=0, sticky="w", pady=3)
        self.delay_var = tk.StringVar(value="1.5")
        ttk.Entry(settings, textvariable=self.delay_var, width=10).grid(
            row=row, column=1, sticky="w", padx=5, pady=3
        )

        row += 1
        ttk.Label(settings, text="OCR言語:").grid(row=row, column=0, sticky="w", pady=3)
        ocr_frame = ttk.Frame(settings)
        ocr_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=5, pady=3)
        self.lang_var = tk.StringVar(value="ja,en")
        ttk.Entry(ocr_frame, textvariable=self.lang_var, width=12).pack(side="left")
        self.invert_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(ocr_frame, text="ダークモード自動対応", variable=self.invert_var).pack(
            side="left", padx=15
        )

        row += 1
        ttk.Label(settings, text="出力先:").grid(row=row, column=0, sticky="w", pady=3)
        self.output_var = tk.StringVar(value="output.txt")
        ttk.Entry(settings, textvariable=self.output_var).grid(
            row=row, column=1, sticky="ew", padx=5, pady=3
        )
        ttk.Button(settings, text="選択", width=6, command=self._browse_output).grid(
            row=row, column=2, pady=3
        )

        # --- Buttons ---
        buttons = ttk.Frame(self.root)
        buttons.pack(fill="x", padx=10, pady=5)

        self.start_btn = ttk.Button(buttons, text="▶ 開始", command=self.start_capture)
        self.start_btn.pack(side="left", padx=(0, 5))

        self.stop_btn = ttk.Button(
            buttons, text="■ 停止", command=self.stop_capture, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=5)

        self.ocr_btn = ttk.Button(buttons, text="OCRのみ", command=self.ocr_only)
        self.ocr_btn.pack(side="left", padx=5)

        # --- Progress ---
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.root, variable=self.progress_var, maximum=100
        )
        self.progress.pack(fill="x", padx=10, pady=(5, 0))

        self.status_var = tk.StringVar(value="待機中")
        ttk.Label(self.root, textvariable=self.status_var).pack(anchor="w", padx=10)

        # --- Log ---
        self.log = scrolledtext.ScrolledText(
            self.root, height=12, state="disabled", font=("Menlo", 11)
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All", "*.*")],
        )
        if path:
            self.output_var.set(path)

    # --- Logging ---

    def log_msg(self, msg: str):
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def set_status(self, msg: str):
        self.root.after(0, self.status_var.set, msg)

    def set_progress(self, value: float):
        self.root.after(0, self.progress_var.set, value)

    # --- Window list ---

    def refresh_windows(self):
        self.windows = get_visible_windows()
        labels = [w["label"] for w in self.windows]
        self.window_combo["values"] = labels
        if labels:
            # Kindle / Safari / Chrome を優先選択
            for i, w in enumerate(self.windows):
                if any(k in w["owner"].lower() for k in ("kindle", "safari", "chrome", "firefox", "arc")):
                    self.window_combo.current(i)
                    break
            else:
                self.window_combo.current(0)
            self._on_window_selected()

    def _on_window_selected(self, event=None):
        window = self._get_selected_window()
        if window:
            owner = window["owner"].lower()
            if any(b in owner for b in ("chrome", "safari", "firefox", "arc", "brave", "edge")):
                self.crop_top_var.set("280")
                self.crop_bottom_var.set("140")
            else:
                self.crop_top_var.set("0")
                self.crop_bottom_var.set("0")

    def _get_selected_window(self) -> dict | None:
        idx = self.window_combo.current()
        if 0 <= idx < len(self.windows):
            return self.windows[idx]
        return None

    # --- UI state ---

    def _set_running(self, running: bool):
        self.running = running
        state_normal = "normal" if not running else "disabled"
        state_stop = "normal" if running else "disabled"
        self.root.after(0, lambda: self.start_btn.config(state=state_normal))
        self.root.after(0, lambda: self.ocr_btn.config(state=state_normal))
        self.root.after(0, lambda: self.stop_btn.config(state=state_stop))

    # --- Capture ---

    def start_capture(self):
        window = self._get_selected_window()
        if not window:
            self.log_msg("ウィンドウを選択してください。")
            return

        self.stop_event.clear()
        self._set_running(True)
        self.set_progress(0)

        thread = threading.Thread(target=self._capture_worker, args=(window,), daemon=True)
        thread.start()

    def stop_capture(self):
        self.stop_event.set()
        self.log_msg("停止リクエスト送信...")

    def _capture_worker(self, window: dict):
        try:
            screenshots_dir = Path("screenshots")
            # 古いスクリーンショットを削除
            if screenshots_dir.exists():
                for old in screenshots_dir.glob("page_*.png"):
                    old.unlink()
            screenshots_dir.mkdir(exist_ok=True)

            pages_text = self.pages_var.get().strip()
            max_pages = int(pages_text) if pages_text else 9999
            delay = float(self.delay_var.get())
            direction = "left" if "左" in self.direction_var.get() else "right"

            crop_top = int(self.crop_top_var.get() or 0)
            crop_bottom = int(self.crop_bottom_var.get() or 0)

            self.log_msg(f"対象: {window['label']}")
            self.log_msg(f"Window ID: {window['id']}")
            if crop_top > 0 or crop_bottom > 0:
                self.log_msg(f"クロップ: 上{crop_top}px / 下{crop_bottom}px")
            self.log_msg("")

            # 正しいウィンドウ/タブをアクティベート（Chrome 系はタブ検索付き）
            activate_window(window["owner"], window["name"])

            captured = 0
            for i in range(max_pages):
                if self.stop_event.is_set():
                    self.log_msg("\n手動停止。")
                    break

                page_path = screenshots_dir / f"page_{i:04d}.png"
                capture_window(window["id"], str(page_path))
                crop_image(str(page_path), crop_top, crop_bottom)
                captured += 1

                # 前ページと比較 → 自動停止
                if i > 0:
                    prev_path = screenshots_dir / f"page_{i - 1:04d}.png"
                    if images_match(prev_path, page_path):
                        page_path.unlink()
                        captured -= 1
                        self.log_msg(f"\n最終ページ検出 (page {captured})。")
                        break

                self.log_msg(f"[{captured}] {page_path.name}")
                self.set_status(f"キャプチャ中... {captured} ページ")

                if max_pages < 9999:
                    self.set_progress((captured / max_pages) * 50)

                if i < max_pages - 1:
                    turn_page_by_key(direction)
                    time.sleep(delay)

            self.log_msg(f"\nスクリーンショット完了: {captured} ページ")

            if captured == 0 or self.stop_event.is_set():
                return

            # OCR
            languages = [lang.strip() for lang in self.lang_var.get().split(",")]
            self._run_ocr(screenshots_dir, languages)

        except Exception as e:
            self.log_msg(f"\nError: {e}")
        finally:
            self._set_running(False)

    # --- OCR ---

    def ocr_only(self):
        screenshots_dir = Path("screenshots")
        if not screenshots_dir.exists() or not list(screenshots_dir.glob("page_*.png")):
            self.log_msg("screenshots/ にファイルがありません。")
            return

        self.stop_event.clear()
        self._set_running(True)
        self.set_progress(0)

        languages = [lang.strip() for lang in self.lang_var.get().split(",")]
        thread = threading.Thread(
            target=self._ocr_worker, args=(screenshots_dir, languages), daemon=True
        )
        thread.start()

    def _ocr_worker(self, screenshots_dir: Path, languages: list[str]):
        try:
            self._run_ocr(screenshots_dir, languages)
        except Exception as e:
            self.log_msg(f"\nError: {e}")
        finally:
            self._set_running(False)

    def _run_ocr(self, screenshots_dir: Path, languages: list[str]):
        image_files = sorted(screenshots_dir.glob("page_*.png"))
        if not image_files:
            self.log_msg("OCR対象のファイルがありません。")
            return

        invert = self.invert_var.get()
        total = len(image_files)
        self.log_msg(f"\nOCR処理中... ({total} ページ{', ダークモード自動対応' if invert else ''})")
        self.set_status("OCR処理中...")

        all_text = []
        for i, path in enumerate(image_files):
            if self.stop_event.is_set():
                self.log_msg("\nOCR停止。")
                break

            text = ocr_image(str(path), languages, invert=invert)
            chars = len(text)
            self.log_msg(f"  [{i + 1}/{total}] {path.name} → {chars} 文字")
            all_text.append(text)
            self.set_progress(50 + ((i + 1) / total) * 50)

        if not all_text:
            return

        output = Path(self.output_var.get())
        output.write_text("\n\n---\n\n".join(all_text), encoding="utf-8")
        total_chars = sum(len(t) for t in all_text)
        self.log_msg(f"\n完了！ {output} ({len(all_text)} ページ, {total_chars:,} 文字)")
        self.set_status(f"完了！ {len(all_text)} ページ, {total_chars:,} 文字")
        self.set_progress(100)


def main():
    root = tk.Tk()

    # macOS: tkinter が NSApp を作った後にアクティベートして前面に出す
    try:
        from AppKit import NSApp
        NSApp.setActivationPolicy_(0)  # NSApplicationActivationPolicyRegular
        NSApp.activateIgnoringOtherApps_(True)
    except ImportError:
        pass

    root.lift()
    root.attributes("-topmost", True)
    root.after(200, lambda: root.attributes("-topmost", False))
    root.focus_force()

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
