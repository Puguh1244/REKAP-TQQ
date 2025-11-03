# Deploy Guide for `app.py` (Streamlit)

## 1) Quickest: Streamlit Community Cloud (free)
1. Push `app.py`, `requirements.txt`, and the folder `.streamlit/` to a GitHub repo.
2. Go to https://streamlit.io/cloud → **New app** → connect your repo.
3. Set **Main file path** = `app.py`.
4. (Optional) Add secrets if needed; not required for this app.
5. Click **Deploy**.

> Upload cap is set via `.streamlit/config.toml` (`maxUploadSize = 1000` MB). Community Cloud may still enforce platform limits.

## 2) Local run
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## 3) Docker (Render/Railway/any VPS)
```bash
docker build -t rekap-app .
docker run -p 8501:8501 -e PORT=8501 rekap-app
```
Then open http://localhost:8501

## Notes
- Required Python packages are pinned in `requirements.txt`.
- Theme and server settings live in `.streamlit/config.toml`.
- For Heroku-like platforms, `Procfile` is included.
