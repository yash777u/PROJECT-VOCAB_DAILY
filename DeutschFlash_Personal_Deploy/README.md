# DeutschFlash Personal Deploy

Self-contained Streamlit app for German vocabulary practice.

Everything needed for the personal deployment lives in this folder:

- `app.py` - Streamlit UI
- `lib/` - Excel, image, audio, translation, and PPT helpers
- `data/` - copied Excel workbooks, `German.pptx`, slide cache, and audio cache
- `docs/` - migration and maintenance notes
- `packages.txt` - system packages for Streamlit Cloud PPT conversion

## Run Locally

```bash
cd DeutschFlash_Personal_Deploy
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Or:

```bash
cd DeutschFlash_Personal_Deploy
bash start.sh
```

## Deploy

### Streamlit Community Cloud

1. Put this folder in a GitHub repository.
2. In Streamlit Cloud, create a new app.
3. Set the app file to `app.py`.
4. Keep `packages.txt` in the repository root so slide conversion can use LibreOffice and Poppler.

### Render / Railway / VPS

Use this command:

```bash
bash start.sh
```

The script respects the platform `PORT` variable.

## Features Preserved

- Practice by level and sheet
- Marathon practice mode
- MCQ answer flow with score, accuracy, and streak
- Search across all Excel workbooks
- Random timed tests
- Dict.cc-style hover translations
- German TTS with disk cache
- DuckDuckGo image lookup by `keyword`
- Visual slide viewer from `data/German.pptx` and `data/slides_cache`

## Data

The Excel files and PowerPoint file are copied into `data/` and are read directly by the app. To add words, edit the workbook in this folder, then restart the Streamlit app if cached data does not refresh immediately.
