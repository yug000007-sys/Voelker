# Voelker Quote Extractor v9

Updates:
- Upload only `.msg` files; no `Volkr.xlsx` upload needed.
- Built-in Voelker output headers are used automatically.
- Download output as ZIP containing Excel + renamed quote PDFs.
- Added **Clear uploaded files / reset** button.
- Uploaded MSG files, extracted Excel, and renamed PDFs are processed in memory only.
- Temporary MSG parser files are deleted immediately.

Run:
```bash
pip install -r requirements.txt
streamlit run app.py
```

For Streamlit Cloud privacy, keep referral email mapping in Streamlit Secrets where possible.
