# Voelker Quote Extractor v8

Updates in v8:
- Download output as one ZIP containing the completed Excel and renamed quote PDFs.
- Renamed PDF format uses the PDF column value, for example `VOELKER_133795.pdf`.
- Privacy-focused processing: uploaded MSG files, generated Excel, and PDFs are processed in memory only. Temporary MSG files used by the parser are deleted immediately. No output files are saved in the app folder.
- Download button clears generated rows/ZIP from Streamlit session memory on click so the page returns to upload mode on the next rerun.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud secrets
Keep real ReferralManager emails in Streamlit Cloud Secrets or local `.streamlit/secrets.toml`. Do not commit `secrets.toml` to GitHub.
