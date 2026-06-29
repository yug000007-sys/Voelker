# Voelker Quote Extractor - Streamlit

Upload Voelker `.msg` quote emails and the `Volkr.xlsx` template, then download a completed Excel file.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and optionally your blank `Volkr.xlsx` template.
3. Go to Streamlit Community Cloud and deploy the repo.
4. Main file path: `app.py`.

## Notes

- No API key is required.
- The app uses `extract-msg` to read Outlook `.msg` files.
- Voelker mapping rules included:
  - Ship To block is used for customer location.
  - Salesperson maps to `ReferralManager`.
  - Quoted By maps to `Created_By`.
  - Cust # / Customer # maps to `CustomerNumber`.
  - Quote line rows map to item fields.
