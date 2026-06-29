import io
import os
import re
import tempfile
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup

try:
    import extract_msg
except Exception:
    extract_msg = None

st.set_page_config(page_title="Voelker Quote Extractor", layout="wide")
st.title("Voelker Quote Extractor")
st.caption("Upload Voelker .msg quote emails + Volkr.xlsx template. No API required.")

TEMPLATE_COLUMNS = [
    "ReferralManager", "ReferralEmail", "QuoteNumber", "QuoteDate", "Company",
    "FirstName", "LastName", "ContactEmail", "ContactPhone", "Address", "County",
    "City", "State", "ZipCode", "Country", "manufacturer_Name", "item_id",
    "item_desc", "Quantity", "TotalSales", "PDF", "Brand", "QuoteExpiration",
    "CustomerNumber", "UnitSales", "Unit_Cost", "sales_cost", "cust_type",
    "QuoteComment", "Created_By", "quote_line_no", "DemoQuote"
]

STATE_RE = r"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY"


def clean_text(value: str) -> str:
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def money_to_float(value: str):
    if not value:
        return ""
    value = re.sub(r"[^0-9.\-]", "", str(value))
    if value in ("", ".", "-", "-."):
        return ""
    try:
        return float(value)
    except Exception:
        return ""


def qty_to_number(value: str):
    if not value:
        return ""
    value = re.sub(r"[^0-9.\-]", "", str(value))
    if value in ("", ".", "-", "-."):
        return ""
    try:
        num = float(value)
        return int(num) if num.is_integer() else num
    except Exception:
        return ""


def read_msg(uploaded_file) -> Tuple[str, str, str, List[str]]:
    """Return subject, body text, html, attachment filenames."""
    if extract_msg is None:
        raise RuntimeError("extract-msg is not installed. Check requirements.txt on Streamlit Cloud.")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
        tmp.write(uploaded_file.getvalue())
        msg_path = tmp.name

    try:
        msg = extract_msg.Message(msg_path)
        subject = clean_text(getattr(msg, "subject", "") or uploaded_file.name)
        body = clean_text(getattr(msg, "body", "") or "")
        html = getattr(msg, "htmlBody", None) or getattr(msg, "html", None) or ""
        if isinstance(html, bytes):
            html = html.decode("utf-8", errors="ignore")
        attachments = []
        for att in getattr(msg, "attachments", []) or []:
            name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or ""
            if name:
                attachments.append(name)
        return subject, body, html, attachments
    finally:
        try:
            os.remove(msg_path)
        except Exception:
            pass


def html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    for br in soup.find_all(["br", "p", "tr"]):
        br.append("\n")
    return clean_text(soup.get_text(" "))


def get_regex(pattern: str, text: str, flags=re.I) -> str:
    m = re.search(pattern, text, flags)
    return clean_text(m.group(1)) if m else ""


def split_name(name: str) -> Tuple[str, str]:
    name = clean_text(name)
    if not name:
        return "", ""
    parts = name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def parse_city_state_zip(line: str) -> Tuple[str, str, str, str]:
    line = clean_text(line).replace(",", " ")
    m = re.search(rf"(.+?)\s+({STATE_RE})\s+(\d{{5}})(?:-\d{{4}})?\b", line, re.I)
    if m:
        return clean_text(m.group(1)), m.group(2).upper(), m.group(3), "USA"
    return "", "", "", ""


def parse_address_block(text: str) -> Dict[str, str]:
    """Prefer Ship To. Fallback to Sold To/Bill To/Customer blocks."""
    out = {"Company": "", "Address": "", "City": "", "State": "", "ZipCode": "", "Country": "", "County": ""}
    patterns = [
        r"Ship\s*To\s*:?\s*(.*?)(?:\n\s*(?:Bill\s*To|Sold\s*To|Quote\s*Details|Line|Item|Description|Subtotal|Total)\b)",
        r"SHIP\s*TO\s*:?\s*(.*?)(?:\n\s*(?:BILL\s*TO|SOLD\s*TO|QUOTE|LINE|ITEM|DESCRIPTION|SUBTOTAL|TOTAL)\b)",
        r"Customer\s*:?\s*(.*?)(?:\n\s*(?:Quote|Line|Item|Description|Subtotal|Total)\b)",
    ]
    block = ""
    for pat in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            block = clean_text(m.group(1))
            break
    if not block:
        return out

    lines = [clean_text(x) for x in block.splitlines() if clean_text(x)]
    lines = [x for x in lines if not re.search(r"^(ship to|bill to|sold to|attn|phone|fax|email)\b", x, re.I)]

    csz_idx = None
    for i, line in enumerate(lines):
        city, state, zip_code, country = parse_city_state_zip(line)
        if city:
            csz_idx = i
            out.update({"City": city, "State": state, "ZipCode": zip_code, "Country": country})
            break

    if csz_idx is not None:
        if csz_idx >= 2:
            out["Company"] = lines[0]
            out["Address"] = lines[1]
        elif csz_idx == 1:
            out["Address"] = lines[0]
        if not out["Country"] and csz_idx + 1 < len(lines):
            if re.search(r"united states|usa|u\.s\.a", lines[csz_idx + 1], re.I):
                out["Country"] = "USA"
    return out


def parse_header(subject: str, text: str, filename: str) -> Dict[str, str]:
    d = {}
    all_text = subject + "\n" + text + "\n" + filename

    d["QuoteNumber"] = get_regex(r"Quote\s*#?\s*[:\-]?\s*(\d+)", all_text) or get_regex(r"#\s*(\d{5,})", all_text)
    d["Company"] = get_regex(r"Customer\s*[:\-]\s*([^\n\r]+)", all_text) or get_regex(r"Customer_\s*([^\.]+)", filename)
    d["QuoteDate"] = get_regex(r"Quote\s*Date\s*[:\-]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", text)
    d["QuoteExpiration"] = get_regex(r"(?:Expiration|Expires|Valid\s*Until|Quote\s*Expiration)\s*[:\-]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", text)
    d["CustomerNumber"] = get_regex(r"(?:Cust(?:omer)?\s*#|Customer\s*No\.?|Cust\s*No\.?)\s*[:\-]?\s*([A-Z0-9\-]+)", text)
    d["ReferralManager"] = get_regex(r"Salesperson\s*[:\-]?\s*([^\n\r]+)", text) or get_regex(r"Sales\s*Person\s*[:\-]?\s*([^\n\r]+)", text)
    d["Created_By"] = get_regex(r"Quoted\s*By\s*[:\-]?\s*([^\n\r]+)", text) or get_regex(r"Quote\s*Prepared\s*By\s*[:\-]?\s*([^\n\r]+)", text)
    d["ContactEmail"] = get_regex(r"([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", text)
    d["ContactPhone"] = get_regex(r"(?:Phone|Tel)\s*[:\-]?\s*(\(?\d{3}\)?[\s\.-]?\d{3}[\s\.-]?\d{4})", text)

    contact = get_regex(r"(?:Contact|Attention|Attn)\s*[:\-]?\s*([^\n\r]+)", text)
    d["FirstName"], d["LastName"] = split_name(contact)
    d["PDF"] = f"Voelker_Quote_{d['QuoteNumber']}.pdf" if d.get("QuoteNumber") else ""
    return {k: clean_text(v) for k, v in d.items()}


def extract_html_table_rows(html: str) -> List[List[str]]:
    rows = []
    if not html:
        return rows
    soup = BeautifulSoup(html, "lxml")
    for tr in soup.find_all("tr"):
        cells = [clean_text(c.get_text(" ")) for c in tr.find_all(["td", "th"])]
        cells = [c for c in cells if c]
        if cells:
            rows.append(cells)
    return rows


def parse_line_items(text: str, html: str) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = []

    # HTML table path: find rows containing qty and money values.
    for cells in extract_html_table_rows(html):
        joined = " | ".join(cells)
        if not re.search(r"\$|\d+\.\d{2}", joined):
            continue
        if re.search(r"subtotal|tax|freight|shipping|grand total|total due", joined, re.I):
            continue
        money_vals = [c for c in cells if re.search(r"\$?\s*\d[\d,]*\.\d{2}", c)]
        qty_vals = [c for c in cells if re.fullmatch(r"\d+(?:\.\d+)?", c)]
        if len(money_vals) >= 1 and qty_vals:
            line_no = cells[0] if re.fullmatch(r"\d+", cells[0]) else ""
            qty = qty_vals[0]
            total = money_vals[-1]
            unit = money_vals[-2] if len(money_vals) >= 2 else ""
            item_id = ""
            desc_parts = []
            for c in cells:
                if c in money_vals or c in qty_vals or c == line_no:
                    continue
                if not item_id and re.search(r"[A-Z0-9][A-Z0-9\-_/]{2,}", c, re.I):
                    item_id = c
                else:
                    desc_parts.append(c)
            if item_id or desc_parts:
                items.append({
                    "quote_line_no": line_no,
                    "item_id": item_id,
                    "item_desc": clean_text(" ".join(desc_parts)),
                    "Quantity": qty_to_number(qty),
                    "UnitSales": money_to_float(unit),
                    "TotalSales": money_to_float(total),
                })

    if items:
        return items

    # Plain text fallback.
    for line in text.splitlines():
        l = clean_text(line)
        if len(l) < 15 or re.search(r"subtotal|tax|freight|shipping|grand total|total due", l, re.I):
            continue
        m = re.match(r"^(\d+)\s+([A-Z0-9][A-Z0-9\-_/\.]+)\s+(.+?)\s+(\d+(?:\.\d+)?)\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})$", l, re.I)
        if m:
            items.append({
                "quote_line_no": m.group(1),
                "item_id": m.group(2),
                "item_desc": clean_text(m.group(3)),
                "Quantity": qty_to_number(m.group(4)),
                "UnitSales": money_to_float(m.group(5)),
                "TotalSales": money_to_float(m.group(6)),
            })
    return items


def parse_one_msg(uploaded_file) -> List[Dict[str, str]]:
    subject, body, html, attachments = read_msg(uploaded_file)
    text = clean_text(body + "\n" + html_to_text(html))
    header = parse_header(subject, text, uploaded_file.name)
    address = parse_address_block(text)
    for k, v in address.items():
        if v:
            header[k] = v
    if not header.get("Company"):
        header["Company"] = address.get("Company", "")

    rows = []
    line_items = parse_line_items(text, html)
    if not line_items:
        line_items = [{"QuoteComment": "No line items detected - review manually."}]

    for item in line_items:
        row = {col: "" for col in TEMPLATE_COLUMNS}
        row.update(header)
        row.update(item)
        row["DemoQuote"] = "No"
        row["cust_type"] = ""
        rows.append(row)
    return rows


def write_output(template_file, rows: List[Dict[str, str]]) -> bytes:
    df_new = pd.DataFrame(rows)
    for col in TEMPLATE_COLUMNS:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[TEMPLATE_COLUMNS]

    if template_file is not None:
        template_bytes = template_file.getvalue()
        try:
            base = pd.read_excel(io.BytesIO(template_bytes), dtype=str)
            columns = list(base.columns) if len(base.columns) else TEMPLATE_COLUMNS
        except Exception:
            columns = TEMPLATE_COLUMNS
    else:
        columns = TEMPLATE_COLUMNS

    for col in columns:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[columns]

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_new.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.book["Sheet1"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        for col_cells in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col_cells[:100])
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 10), 38)
    return out.getvalue()


with st.sidebar:
    st.header("Upload")
    msg_files = st.file_uploader("Voelker .msg files", type=["msg"], accept_multiple_files=True)
    template_file = st.file_uploader("Volkr.xlsx template", type=["xlsx"])

if extract_msg is None:
    st.error("Missing dependency: extract-msg. Add it to requirements.txt and redeploy.")

if msg_files:
    if st.button("Extract Voelker Quotes", type="primary"):
        all_rows = []
        errors = []
        progress = st.progress(0)
        for i, f in enumerate(msg_files, start=1):
            try:
                all_rows.extend(parse_one_msg(f))
            except Exception as e:
                errors.append(f"{f.name}: {e}")
            progress.progress(i / len(msg_files))

        if all_rows:
            df_preview = pd.DataFrame(all_rows)
            st.success(f"Extracted {len(all_rows)} quote line rows from {len(msg_files)} email(s).")
            st.dataframe(df_preview[TEMPLATE_COLUMNS], use_container_width=True)
            output = write_output(template_file, all_rows)
            st.download_button(
                "Download completed Volkr output",
                data=output,
                file_name=f"voelker_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        if errors:
            st.warning("Some files need manual review:")
            for err in errors:
                st.write("- " + err)
else:
    st.info("Upload one or more Voelker .msg files to start.")
