# FileSort Cleaner (Streamlit)

Streamlit app that cleans and merges race-registration files for Sports Timing Solutions.

- Upload one or more CSV / XLSX / XLS files (multi-sheet Excel supported)
- Keep or ignore each sheet
- Map columns per sheet, then rename category and T-shirt values
- Rule-based cleaning (names, gender, phone, DOB, age, blood group, country)
- Flagged cells are colour-coded; download one merged Excel file

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Regression checks (run before deploying)

```bash
python regression_checks.py
```

`All regression checks passed.` means the core cleaning behaviour is intact.

## Deploy on Streamlit Community Cloud

Push to GitHub, then deploy `app.py` from [share.streamlit.io](https://share.streamlit.io).
