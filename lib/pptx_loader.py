"""
pptx_loader.py — Slide image preparation module.

Converts German.pptx → optimised WebP images served as static URLs.
Returns lightweight URL strings instead of base64 blobs so the browser
can load and cache each slide on demand (dramatically faster on mobile).

Static serving must be enabled in .streamlit/config.toml:
    [server]
    enableStaticServing = true

Slides are served at: app/static/slides/slide-XX.webp
(relative URL — works on local dev and Streamlit Cloud under any subpath)
"""
import os
import subprocess
import streamlit as st

# ── Paths ────────────────────────────────────────────────────────────────────
STATIC_DIR   = os.path.join("static", "slides")   # served as app/static/slides/
PPTX_PATH    = os.path.join("data", "German.pptx")
CACHE_DIR    = os.path.join("data", "slides_cache")  # intermediate PNGs from pdftoppm
LOCAL_MARKER = ".local_dev"                           # gitignored; exists only on laptop

# ── Quality settings ─────────────────────────────────────────────────────────
WEBP_QUALITY = 78   # ~170 KB/slide — good quality, loads fast on mobile
WEBP_MAX_W   = 1280 # px — enough for any phone / tablet


def _static_url(filename: str) -> str:
    """
    Build the correct static file URL for Streamlit's built-in file server.

    Streamlit serves ./static/ at <baseUrlPath>/app/static/.
    Using NO leading slash (relative URL) makes it work on:
      • Local dev:         http://localhost:8501/app/static/slides/slide-01.webp
      • Streamlit Cloud:   https://xxx.streamlit.app/app/static/slides/slide-01.webp
    regardless of any deploy subpath.
    """
    try:
        base = (st.get_option("server.baseUrlPath") or "").strip("/")
    except Exception:
        base = ""
    prefix = f"{base}/app/static/slides" if base else "app/static/slides"
    return f"{prefix}/{filename}"


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
    Return a list of static URL strings for the optimised WebP slides.

    Strategy:
    ─────────
    1. If WebP files already exist in static/slides/ → return their URLs
       immediately. No conversion, no wait, lightning-fast every time.

    2. If static/slides/ is empty → run the full pipeline once:
         PPTX → PDF → PNG (intermediate) → WebP → delete PNGs
       After that, step 1 applies on every subsequent startup.

    To force a fresh rebuild: delete the static/slides/ folder and restart.

    URLs are relative (no leading slash) so they work on both local dev
    and Streamlit Cloud under any deploy subpath.
    """
    if not os.path.exists(PPTX_PATH):
        return []

    os.makedirs(STATIC_DIR, exist_ok=True)

    # ── Step 1: check for existing WebPs ─────────────────────────────────────
    existing_webp = sorted(
        [f for f in os.listdir(STATIC_DIR)
         if f.startswith("slide-") and f.endswith(".webp")
         and os.path.getsize(os.path.join(STATIC_DIR, f)) > 0],
        key=_get_num,
    )

    if existing_webp:
        if _is_local():
            print(f"[pptx_loader] {len(existing_webp)} WebP slides cached — loaded instantly.")
            print("[pptx_loader] To refresh: delete the  static/slides/  folder and restart.")
        return [_static_url(f) for f in existing_webp]

    # ── Step 2: first run / empty folder — generate everything ───────────────
    print("[pptx_loader] static/slides/ is empty — generating WebP slides from PPTX…")
    os.makedirs(CACHE_DIR, exist_ok=True)
    _wipe_dir(CACHE_DIR)

    png_paths  = _convert_pptx_to_pngs()   # PPTX → PDF → PNG  (intermediate)
    webp_paths = _pngs_to_webp(png_paths)   # PNG  → WebP       (final served files)
    _wipe_dir(CACHE_DIR)                    # delete intermediate PNGs

    if not webp_paths:
        return []

    print(f"[pptx_loader] Done — {len(webp_paths)} WebP slides ready.")
    return [
        _static_url(os.path.basename(p))
        for p in sorted(webp_paths, key=lambda p: _get_num(os.path.basename(p)))
    ]
