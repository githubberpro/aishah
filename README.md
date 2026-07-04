# Aishah

A Streamlit app, deployed on [Streamlit Community Cloud](https://share.streamlit.io).

## Run locally

```bash
pip install -r requirements-dev.txt
streamlit run streamlit_app.py
```

The app opens at http://localhost:8501.

## Test

```bash
pip install -r requirements-dev.txt
pytest -q
```

Tests also run automatically in GitHub Actions on every push and pull request
(see `.github/workflows/ci.yml`).

## Deploy to share.streamlit.io

Streamlit Community Cloud deploys directly from this GitHub repo — no build
scripts or servers to manage. One-time setup:

1. Push this repo to GitHub (already done on your branch).
2. Go to <https://share.streamlit.io> and sign in with GitHub.
3. Click **Create app** → **Deploy a public app from GitHub**.
4. Fill in:
   - **Repository:** `githubberpro/aishah`
   - **Branch:** the branch you want to deploy (e.g. `main`)
   - **Main file path:** `streamlit_app.py`
5. Click **Deploy**.

Streamlit installs the deps in `requirements.txt`, pins Python via
`runtime.txt`, and serves the app at a `*.streamlit.app` URL.

### Updates & secrets

- **Continuous deploy:** every push to the deployed branch redeploys automatically.
- **Secrets:** add API keys etc. in the app's **Settings → Secrets** on
  Streamlit Cloud (TOML format). Read them in code via `st.secrets`. Never
  commit `.streamlit/secrets.toml` — it is git-ignored.

## Files

| File | Purpose |
| --- | --- |
| `streamlit_app.py` | App entry point (Streamlit Cloud's default filename) |
| `requirements.txt` | Runtime deps installed by Streamlit Cloud |
| `requirements-dev.txt` | Dev/CI deps (adds pytest) |
| `runtime.txt` | Pins the Python version on Streamlit Cloud |
| `.streamlit/config.toml` | Streamlit config |
| `tests/` | Unit tests run by pytest / CI |
| `.github/workflows/ci.yml` | Runs tests on every push and PR |
