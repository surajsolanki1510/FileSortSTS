# FileSort Cleaner (Streamlit)

Python Streamlit app to clean race-registration files with:

- Upload CSV/XLSX/XLS
- Post-scan manual column mapping
- Manual category mapping
- Rule-based cleaning
- Row highlighting (yellow/red)
- Download cleaned Excel

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Regression checks (recommended before deploy)

```bash
python regression_checks.py
```

If this prints `All regression checks passed.`, core cleaning behavior is intact.

## Deploy on Streamlit Community Cloud

1. Push this folder to GitHub.
2. Open [https://share.streamlit.io](https://share.streamlit.io).
3. Choose your repo and deploy `app.py`.

