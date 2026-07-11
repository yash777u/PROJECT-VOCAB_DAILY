"""
pptx_loader.py — Slide image preparation module.

Converts German.pptx → optimised WebP images served as static URLs.
Returns lightweight URL strings instead of base64 blobs so the browser
can load and cache each slide on demand (dramatically faster on mobile).

LOCAL vs DEPLOYMENT detection
───────────────────────────────
A `.local_dev` marker file in the project root signals a local machine.
When that file is present every startup wipes both caches and regenerates
all slides fresh — so dropping a new PPTX is always reflected immediately.
In deployment (no `.local_dev`) the existing WebP files are reused until
they are older than the PPTX, keeping startup fast.

Static serving must be enabled in .streamlit/config.toml:
    [server]
    enableStaticServing = true

Slides are served at: /app/static/slides/slide-XX.webp
"""
import os
import subprocess
import streamlit as st

# ── Paths ────────────────────────────────────────────────────────────────────
STATIC_DIR   = os.path.join("static", "slides")   # served as /app/static/slides/
PPTX_PATH    = os.path.join("data", "German.pptx")
CACHE_DIR    = os.path.join("data", "slides_cache")  # intermediate PNGs from pdftoppm
LOCAL_MARKER = ".local_dev"                           # gitignored; exists only on laptop

# ── Quality settings ─────────────────────────────────────────────────────────
WEBP_QUALITY = 78   # ~170 KB/slide — good quality, loads fast on mobile
WEBP_MAX_W   = 1280 # px — enough for any phone / tablet


# ── Helpers ──────────────────────────────────────────────────────────────────

def _is_local() -> bool:
    """Return True when running on the developer's local machine."""
    return os.path.exists(LOCAL_MARKER)


def _get_num(filename: str) -> int:
    """Extract slide number from names like 'slide-01.webp' / 'slide-01.png'."""
    try:
        return int(filename.split("-")[1].split(".")[0])
    except Exception:
        return 99999


def _wipe_dir(directory: str) -> None:
    """Delete every file inside *directory* (keeps the dir itself)."""
    if not os.path.exists(directory):
        return
    for f in os.listdir(directory):
        try:
            os.remove(os.path.join(directory, f))
        except Exception:
            pass


def _convert_pptx_to_pngs() -> list:
    """
    Convert German.pptx → per-slide PNGs in CACHE_DIR via LibreOffice + pdftoppm.
    Returns a sorted list of absolute PNG file paths.
    """
    os.makedirs(CACHE_DIR, exist_ok=True)

    # Step 1: LibreOffice pptx → pdf
    pdf_path = os.path.join(CACHE_DIR, "German.pdf")
    for lo_bin in ["/snap/bin/libreoffice", "libreoffice"]:
        try:
            subprocess.run(
                [lo_bin, "--headless", "--convert-to", "pdf",
                 "--outdir", CACHE_DIR, PPTX_PATH],
                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            break
        except Exception:
            continue

    if not os.path.exists(pdf_path):
        print("[pptx_loader] LibreOffice conversion failed — no PDF produced.")
        return []

    # Step 2: pdftoppm pdf → per-page PNGs at 150 dpi
    for ppb in ["/usr/bin/pdftoppm", "pdftoppm"]:
        try:
            subprocess.run(
                [ppb, "-png", "-r", "150",
                 pdf_path, os.path.join(CACHE_DIR, "slide")],
                check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            break
        except Exception:
            continue

    # Clean up intermediate PDF
    try:
        os.remove(pdf_path)
    except Exception:
        pass

    pngs = sorted(
        [f for f in os.listdir(CACHE_DIR) if f.startswith("slide-") and f.endswith(".png")],
        key=_get_num,
    )
    return [os.path.join(CACHE_DIR, p) for p in pngs]


def _pngs_to_webp(png_paths: list) -> list:
    """
    Convert a list of PNG paths → WebP files in STATIC_DIR.
    Skips files that already exist and are newer than their source PNG.
    Returns a sorted list of output WebP paths.
    """
    try:
        from PIL import Image
    except ImportError:
        print("[pptx_loader] Pillow not installed — cannot convert to WebP.")
        return []

    os.makedirs(STATIC_DIR, exist_ok=True)
    webp_paths = []

    for png_path in png_paths:
        name = os.path.basename(png_path).replace(".png", ".webp")
        out  = os.path.join(STATIC_DIR, name)

        # Regenerate if missing or source PNG is newer
        if not os.path.exists(out) or os.path.getmtime(out) < os.path.getmtime(png_path):
            try:
                img = Image.open(png_path)
                w, h = img.size
                if w > WEBP_MAX_W:
                    h = int(h * WEBP_MAX_W / w)
                    img = img.resize((WEBP_MAX_W, h), Image.LANCZOS)
                img.save(out, "WEBP", quality=WEBP_QUALITY, method=6)
            except Exception as e:
                print(f"[pptx_loader] WebP conversion failed for {png_path}: {e}")
                continue

        webp_paths.append(out)

    return webp_paths


# ── Public API ────────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def get_slide_images() -> list:
    """
    Prepare optimised WebP slides and return a list of static URL strings.

    LOCAL mode  (.local_dev present)
    ─────────────────────────────────
    • Always wipes both caches (data/slides_cache + static/slides) on startup.
    • Re-converts the PPTX from scratch so a freshly placed file is always
      reflected without any manual cache clearing.

    DEPLOYMENT mode (.local_dev absent)
    ─────────────────────────────────────
    • Uses existing WebP files if they are newer than the PPTX.
    • Only re-converts if the PPTX has been updated or WebPs are missing.

    Returns URLs like /app/static/slides/slide-01.webp (one per slide).
    Cached for the lifetime of the server process.
    """
    if not os.path.exists(PPTX_PATH):
        return []

    os.makedirs(STATIC_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)

    local = _is_local()

    if local:
        # ── LOCAL: always regenerate everything ──────────────────────────
        print("[pptx_loader] LOCAL mode — wiping caches and regenerating slides…")
        _wipe_dir(CACHE_DIR)
        _wipe_dir(STATIC_DIR)
        png_paths = _convert_pptx_to_pngs()
        webp_paths = _pngs_to_webp(png_paths)

    else:
        # ── DEPLOYMENT: use cached WebPs when still valid ────────────────
        pptx_mtime = os.path.getmtime(PPTX_PATH)

        existing_webp = sorted(
            [f for f in os.listdir(STATIC_DIR)
             if f.startswith("slide-") and f.endswith(".webp")],
            key=_get_num,
        )
        cache_valid = (
            len(existing_webp) > 0
            and all(
                os.path.getmtime(os.path.join(STATIC_DIR, f)) >= pptx_mtime
                and os.path.getsize(os.path.join(STATIC_DIR, f)) > 0
                for f in existing_webp
            )
        )

        if cache_valid:
            return [f"/app/static/slides/{f}" for f in existing_webp]

        # WebPs stale / missing — check PNG cache before calling LibreOffice
        existing_pngs = sorted(
            [f for f in os.listdir(CACHE_DIR)
             if f.startswith("slide-") and f.endswith(".png")],
            key=_get_num,
        ) if os.path.exists(CACHE_DIR) else []

        png_cache_valid = (
            len(existing_pngs) > 0
            and all(
                os.path.getmtime(os.path.join(CACHE_DIR, f)) >= pptx_mtime
                and os.path.getsize(os.path.join(CACHE_DIR, f)) > 0
                for f in existing_pngs
            )
        )

        if png_cache_valid:
            png_paths = [os.path.join(CACHE_DIR, f) for f in existing_pngs]
        else:
            _wipe_dir(CACHE_DIR)
            png_paths = _convert_pptx_to_pngs()

        _wipe_dir(STATIC_DIR)
        webp_paths = _pngs_to_webp(png_paths)

    if not webp_paths:
        return []

    return [
        f"/app/static/slides/{os.path.basename(p)}"
        for p in sorted(webp_paths, key=lambda p: _get_num(os.path.basename(p)))
    ]
