# Deployment Guide

## Recommended Personal Deployment

Use Streamlit Community Cloud for the easiest personal project deployment.

1. Create a GitHub repo containing the files from `DeutschFlash_Personal_Deploy/`.
2. Make sure these files are at the repo root:
   - `app.py`
   - `requirements.txt`
   - `packages.txt`
   - `data/`
   - `lib/`
3. Connect the repo in Streamlit Cloud.
4. Choose `app.py` as the entry file.

## Local Smoke Test

```bash
cd DeutschFlash_Personal_Deploy
pip install -r requirements.txt
python scripts/check_project.py
streamlit run app.py
```

## Notes

- First-time TTS generation can be slower because it calls `edge-tts`. Cached audio is much faster.
- Image lookup uses DuckDuckGo and can be rate-limited. The app fails gracefully when no image is returned.
- Visual slides use the existing `data/slides_cache` first. If the PPT changes, the app tries to regenerate slides using LibreOffice and Poppler.
