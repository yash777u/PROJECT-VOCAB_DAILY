# Migration Notes

## What Changed

This is not a risky rewrite. The project was packaged into one clean deployment folder while keeping the working Streamlit application, Python helpers, Excel files, PPTX file, slide cache, and audio cache together.

## Easiest Parts

- Copying the existing app structure was straightforward because most functionality was already separated into `lib/`.
- Keeping Excel files unchanged was easy because `lib/excel_manager.py` already discovers `data/*_vocab.xlsx` dynamically.
- Preserving audio cache behavior was easy because `lib/tts_service.py` stores files under `data/audio_cache` relative to the project folder.
- Mobile behavior was mostly already handled by the app's CSS media queries and Streamlit column overrides.

## Hardest Parts

- The Visuals/PPT path was the main deployment risk. The old loader used `data/German.pptx` relative to whichever directory started Streamlit. I changed the copied `lib/pptx_loader.py` to resolve `data/` relative to the deployed project folder.
- PPT conversion needs operating-system tools, not only Python packages. `packages.txt` now lists `libreoffice` and `poppler-utils` for hosts like Streamlit Cloud.
- The app depends on external services for new audio, images, and translations. Existing cached audio/slides work offline, but new TTS/image/dict.cc lookups still need internet access.

## How I Tackled It

- I first removed the aborted migration folder so there was no mixed stack left behind.
- I copied the existing working files into `DeutschFlash_Personal_Deploy/`.
- I kept all data inside the new folder so deployment does not depend on files outside it.
- I patched only the copied PPT loader for robust folder-relative paths.
- I added deployment metadata and docs inside the same folder.

## Verification Checklist

- `python -m compileall app.py lib`
- `python scripts/check_project.py`
- `streamlit run app.py`
- Test these tabs manually on phone width and laptop width:
  - Practice
  - Search
  - Random Test
  - Visuals
