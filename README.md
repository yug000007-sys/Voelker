# Voelker Quote Extractor

Streamlit app for extracting Voelker quote data from Outlook `.msg` emails.

## What it does

- Upload one or more Voelker `.msg` files.
- Upload `Volkr.xlsx` as the template.
- The app reads the PDF quote attached inside each MSG.
- It extracts quote header, ship-to details, item, quantity, unit price, total, salesperson, quoted by, customer number, quote date, and expiration.
- Download the completed Excel file.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Create a GitHub repository.
2. Upload these files:
   - `app.py`
   - `requirements.txt`
   - `README.md`
3. In Streamlit Cloud, select the repo and set main file path to:

```text
app.py
```

## Notes

This version extracts data from the PDF attachment inside the `.msg`, not only the email body. The earlier blank output happened because the useful quote details are in the attached PDF.
