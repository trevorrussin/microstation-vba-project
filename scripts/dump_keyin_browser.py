"""
Dump OpenRoads Key-in Browser column-1 keywords (Tables=All) and diff
against Data/command-registry.tsv.

Open path: Ctrl+D (ribbon search) → type Key-in → Enter.
(F9 may be remapped by AECOM Barmenu.)

The Key-in window is owner-drawn (no UIA list children), so this script
BitBlt-captures scrolled column-1 pages and OCRs them with Tesseract.

Usage (ORD must be open):
  python scripts/dump_keyin_browser.py
"""
from __future__ import annotations

import csv
import re
import sys
import time
from pathlib import Path

import uiautomation as auto
import win32api
import win32con
import win32gui
import win32ui
from PIL import Image, ImageEnhance, ImageOps
import pytesseract

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "keyin-browser-dump"
PAGES = OUT / "col1_pages"
REGISTRY = ROOT / "Data" / "command-registry.tsv"

TESS = Path(r"C:\Users\RussinT\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = str(TESS)

# Mid crop: enough left pad for first glyph; stop before col2 bleed
COL1_BOX = (0.02, 0.17, 0.27, 0.72)


def find_keyin_hwnd():
    found = []

    def enum_tl(h, _):
        try:
            if (win32gui.GetWindowText(h) or "") == "Key-in":
                found.append(h)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(enum_tl, None)
    return found[0] if found else None


def ensure_keyin():
    hwnd = find_keyin_hwnd()
    if hwnd and win32gui.IsWindowVisible(hwnd):
        return hwnd
    for w in auto.GetRootControl().GetChildren():
        n = w.Name or ""
        if n.startswith("pw:") or "OpenRoads Designer" in n:
            w.SetActive()
            break
    time.sleep(0.4)
    auto.SendKeys("{ESC}")
    time.sleep(0.15)
    auto.SendKeys("^d")
    time.sleep(0.9)
    auto.SendKeys("^a")
    time.sleep(0.05)
    auto.SendKeys("Key-in")
    time.sleep(1.0)
    auto.SendKeys("{ENTER}")
    time.sleep(1.2)
    hwnd = find_keyin_hwnd()
    if not hwnd:
        auto.SendKeys("{DOWN}{ENTER}")
        time.sleep(1.0)
        hwnd = find_keyin_hwnd()
    if not hwnd:
        raise SystemExit("Key-in failed to open (Ctrl+D then Key-in)")
    return hwnd


def bitblt(hwnd) -> Image.Image:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = max(1, right - left), max(1, bottom - top)
    desk = win32gui.GetDesktopWindow()
    hdc = win32gui.GetWindowDC(desk)
    mfc = win32ui.CreateDCFromHandle(hdc)
    save = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc, w, h)
    save.SelectObject(bmp)
    save.BitBlt((0, 0), (w, h), mfc, (left, top), win32con.SRCCOPY)
    info = bmp.GetInfo()
    bits = bmp.GetBitmapBits(True)
    im = Image.frombuffer(
        "RGB", (info["bmWidth"], info["bmHeight"]), bits, "raw", "BGRX", 0, 1
    )
    win32gui.DeleteObject(bmp.GetHandle())
    save.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(desk, hdc)
    return im


def click_client(hwnd, fx, fy):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    win32api.SetCursorPos((int(l + (r - l) * fx), int(t + (b - t) * fy)))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.04)


def col1_crop(im: Image.Image) -> Image.Image:
    w, h = im.size
    l, t, r, b = COL1_BOX
    return im.crop((int(l * w), int(t * h), int(r * w), int(b * h)))


def ocr_col1(im: Image.Image) -> list[str]:
    crop = col1_crop(im)
    pre = crop.resize((crop.width * 4, crop.height * 4), Image.Resampling.LANCZOS)
    pre = ImageOps.grayscale(pre)
    pre = ImageEnhance.Contrast(ImageOps.autocontrast(pre)).enhance(2.0)
    raw = pytesseract.image_to_string(pre, config="--psm 6")
    lines = []
    for ln in raw.splitlines():
        # drop trailing col2 bleed like "change |"
        s = ln.strip().lower()
        s = s.split("|")[0].split("(")[0].strip()
        s = re.sub(r"[^a-z0-9!$%+\-.]+", "", s)
        if len(s) >= 1:
            lines.append(s)
    return lines


def registry_tokens():
    first, full = set(), set()
    rows = 0
    with REGISTRY.open(encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f, delimiter="\t"):
            rows += 1
            for part in (row.get("recipeLines") or "").split("|"):
                part = part.strip()
                u = part.upper()
                body = None
                if u.startswith("KEYIN:"):
                    body = part[6:].strip()
                elif u.startswith("COMMAND:"):
                    body = part[8:].strip()
                if not body:
                    continue
                body = re.sub(r"\{[^}]+\}", "", body).strip()
                body = re.sub(r"\s+", " ", body)
                if not body:
                    continue
                full.add(body.upper())
                first.add(body.split()[0].upper())
    return rows, first, full


def capture_pages(hwnd, max_pages=150) -> int:
    if PAGES.exists():
        for p in PAGES.glob("*.jpg"):
            p.unlink()
    else:
        PAGES.mkdir(parents=True, exist_ok=True)

    click_client(hwnd, 0.12, 0.35)
    auto.SendKeys("{HOME}")
    time.sleep(0.2)

    first_sig = None
    count = 0
    for page in range(max_pages):
        im = bitblt(hwnd)
        crop = col1_crop(im)
        # save LANCZOS-prep-ready crop (2x for archival)
        save = crop.resize((crop.width * 2, crop.height * 2), Image.Resampling.LANCZOS)
        path = PAGES / f"p{page:03d}.jpg"
        save.save(path, quality=95)
        sig = save.resize((12, 12)).tobytes()
        print(f"saved {path.name}", flush=True)
        count += 1
        if page == 0:
            first_sig = sig
        elif page > 8 and sig == first_sig:
            print("wrapped at", page, flush=True)
            break
        click_client(hwnd, 0.12, 0.35)
        auto.SendKeys("{PAGEDOWN}")
        time.sleep(0.11)
    return count


def ocr_all_pages() -> list[str]:
    items: list[str] = []
    seen = set()
    pages = sorted(PAGES.glob("p*.jpg"))
    for i, p in enumerate(pages):
        # pages already cropped+2x; OCR path expects full window — re-open as crop
        im = Image.open(p)
        # treat whole image as col1 crop
        pre = im.resize((im.width * 2, im.height * 2), Image.Resampling.LANCZOS)
        pre = ImageOps.grayscale(pre)
        pre = ImageEnhance.Contrast(ImageOps.autocontrast(pre)).enhance(2.0)
        raw = pytesseract.image_to_string(pre, config="--psm 6")
        lines = []
        for ln in raw.splitlines():
            s = ln.strip().lower()
            s = s.split("|")[0].split("(")[0].strip()
            s = re.sub(r"[^a-z0-9!$%+\-.]+", "", s)
            if len(s) >= 1:
                lines.append(s)
                if s not in seen:
                    seen.add(s)
                    items.append(s)
        if i < 4 or i % 20 == 0 or i >= len(pages) - 2:
            print(i, p.name, lines, flush=True)
    # drop 1-char junk except symbols
    return [x for x in items if len(x) > 1 or x in {"!", "$", "%"}]


def main():
    if not TESS.exists():
        raise SystemExit(f"tesseract missing: {TESS}")
    OUT.mkdir(parents=True, exist_ok=True)

    hwnd = ensure_keyin()
    print("keyin", hex(hwnd), win32gui.GetWindowRect(hwnd), flush=True)
    n = capture_pages(hwnd)
    print("captured", n, "pages", flush=True)

    items = ocr_all_pages()
    tsv = OUT / "keyin_browser_col1.tsv"
    tsv.write_text("keyword\n" + "\n".join(items) + "\n", encoding="utf-8")
    print("col1", len(items), "->", tsv, flush=True)

    rows, reg_first, reg_full = registry_tokens()
    col1 = {x.upper() for x in items}
    matched = sorted(col1 & reg_first)
    only_b = sorted(col1 - reg_first)
    only_r = sorted(reg_first - col1)

    report = OUT / "keyin_registry_diff.txt"
    report.write_text(
        "\n".join(
            [
                f"col1_keywords={len(col1)}",
                f"registry_rows={rows}",
                f"registry_first_tokens={len(reg_first)}",
                f"registry_full_keyins={len(reg_full)}",
                f"col1_matched_registry_first={len(matched)}",
                f"col1_not_in_registry={len(only_b)}",
                f"registry_first_missing_from_col1={len(only_r)}",
                "",
                "SCOPE: Key-in Browser column-1 only (Tables=All).",
                "Open: Ctrl+D then Key-in.",
                "Source: BitBlt + Tesseract OCR. Not full leaf paths (col2/col3).",
                "",
                "=== matched ===",
                *matched,
                "",
                "=== col1 not in registry ===",
                *only_b,
                "",
                "=== registry first-token missing from col1 ===",
                *only_r,
            ]
        ),
        encoding="utf-8",
    )
    # ascii-safe console summary
    print(
        f"SUMMARY col1={len(col1)} matched={len(matched)} "
        f"only_browser={len(only_b)} only_registry={len(only_r)} "
        f"reg_rows={rows} reg_full={len(reg_full)}",
        flush=True,
    )
    print("wrote", report, flush=True)


if __name__ == "__main__":
    main()
