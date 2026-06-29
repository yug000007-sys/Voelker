# Voelker Quote Extractor

Streamlit app for Voelker `.msg` quote emails.

## What it does

- Upload one or more Voelker Outlook `.msg` files
- Reads the attached quote PDF inside each email
- Extracts quote-level fields into the Volkr template
- Default output matches the manual/yellow row style: **one row per quote**
- Optional line-item mode is available from the sidebar

## Main fixes in this version

- Keeps full quote number format like `01/129486`
- Creates PDF filename as `VOELKER_129486.pdf`
- Uses quote-level `Quote Total` in `TotalSales`
- Does not create separate line-item rows unless you select line-item mode
- Clears item columns in summary mode
- Formats company/address/city in readable title case
- Fixes ReferralManager duplicate issue by separating Authorization from Salesperson
- Populates Country as USA from Sold To block
- Handles multi-line Sold To addresses like street + suite
- Ignores supporting PDF drawings/specs and processes only the quote PDF
- Handles revised quote numbers like 01/133800-R01 while outputting VOELKER_133800.pdf

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Upload these files to GitHub
2. In Streamlit Cloud, create a new app from the repo
3. Main file path: `app.py`


Notes:
- Uses Voelker PDF attachment inside each MSG.
- Address is taken from Sold To, not Ship To.
- Brand is populated as Voelker Controls.


## v5 fix
- Cleans Sold To address prefixes glued by PDF extraction, e.g. `div of Pinnpack 1151 Pacific Ave.` -> `1151 Pacific Ave.` and `a Lincoln Electric Company 407 South Main St` -> `407 South Main St`.


## v6 fix
- Fixed Created_By when Voelker prints a PO/reference number between QuoteNumber and Quoted By, e.g. `01/133795 2401334280 Jessica Halter ...`.
