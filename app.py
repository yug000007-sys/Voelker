import io
import os
import re
import struct
import tempfile
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import extract_msg
except Exception:
    extract_msg = None

st.set_page_config(page_title="Voelker Quote Extractor", layout="wide")
st.title("Voelker Quote Extractor")
st.caption("Upload Voelker .msg quote emails + Volkr.xlsx template. Extracts quote data from the attached PDF inside each MSG. Default output is one row per quote to match the manual Voelker template. No API required.")

TEMPLATE_COLUMNS = [
    "ReferralManager", "ReferralEmail", "QuoteNumber", "QuoteDate", "Company",
    "FirstName", "LastName", "ContactEmail", "ContactPhone", "Address", "County",
    "City", "State", "ZipCode", "Country", "manufacturer_Name", "item_id",
    "item_desc", "Quantity", "TotalSales", "PDF", "Brand", "QuoteExpiration",
    "CustomerNumber", "UnitSales", "Unit_Cost", "sales_cost", "cust_type",
    "QuoteComment", "Created_By", "quote_line_no", "DemoQuote"
]

STATE_RE = r"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY"

# ------------------------- Basic helpers -------------------------

def clean_text(value: str) -> str:
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ").replace("\r", "\n")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n[ \t]+", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def money_to_float(value: str):
    if value is None or value == "":
        return ""
    value = re.sub(r"[^0-9.\-]", "", str(value))
    if value in ("", ".", "-", "-."):
        return ""
    try:
        return float(value)
    except Exception:
        return ""


def qty_to_number(value: str):
    if value is None or value == "":
        return ""
    value = re.sub(r"[^0-9.\-]", "", str(value))
    if value in ("", ".", "-", "-."):
        return ""
    try:
        num = float(value)
        return int(num) if num.is_integer() else num
    except Exception:
        return ""


def split_name(name: str) -> Tuple[str, str]:
    name = clean_text(name)
    if not name:
        return "", ""
    parts = name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def normalize_us_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return clean_text(value)


def get_regex(pattern: str, text: str, flags=re.I) -> str:
    m = re.search(pattern, text, flags)
    return clean_text(m.group(1)) if m else ""

# ------------------------- Pure Python MSG/OLE fallback -------------------------
# This extracts MSG body streams and attachment binary streams even if extract-msg is unavailable.

FREE = 0xFFFFFFFF
END = 0xFFFFFFFE

class CFB:
    def __init__(self, data: bytes):
        self.data = data
        h = data[:512]
        if h[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("Not an Outlook MSG/OLE file")
        self.sector_size = 1 << struct.unpack_from("<H", h, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", h, 32)[0]
        self.first_dir = struct.unpack_from("<I", h, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", h, 56)[0]
        self.first_mini_fat = struct.unpack_from("<I", h, 60)[0]
        self.num_mini_fat = struct.unpack_from("<I", h, 64)[0]
        difat = list(struct.unpack_from("<109I", h, 76))
        self.fat = []
        for sid in difat:
            if sid in (FREE, END):
                continue
            sec = self.sector(sid)
            self.fat.extend(struct.unpack("<%dI" % (len(sec) // 4), sec))
        self.dirs = []
        self._load_dirs()
        root = self.dirs[0]
        self.mini_stream = self._read_big_stream(root["start"], root["size"])
        self.minifat = []
        if self.first_mini_fat not in (FREE, END):
            b = self._read_big_stream(self.first_mini_fat, self.num_mini_fat * self.sector_size)
            self.minifat = list(struct.unpack("<%dI" % (len(b) // 4), b)) if b else []

    def sector(self, sid: int) -> bytes:
        off = 512 + sid * self.sector_size
        return self.data[off:off + self.sector_size]

    def chain(self, start: int, fat=None) -> List[int]:
        fat = fat or self.fat
        out, seen, sid = [], set(), start
        while sid not in (FREE, END) and sid < len(fat) and sid not in seen:
            seen.add(sid)
            out.append(sid)
            sid = fat[sid]
        return out

    def _read_big_stream(self, start: int, size: int = None) -> bytes:
        b = b"".join(self.sector(s) for s in self.chain(start))
        return b[:size] if size is not None else b

    def _read_mini_stream(self, start: int, size: int) -> bytes:
        parts = []
        for sid in self.chain(start, self.minifat):
            off = sid * self.mini_sector_size
            parts.append(self.mini_stream[off:off + self.mini_sector_size])
        return b"".join(parts)[:size]

    def _load_dirs(self):
        b = self._read_big_stream(self.first_dir, None)
        for i in range(len(b) // 128):
            ent = b[i * 128:(i + 1) * 128]
            name_len = struct.unpack_from("<H", ent, 64)[0]
            raw = ent[:max(0, name_len - 2)]
            name = raw.decode("utf-16le", "ignore") if raw else ""
            typ = ent[66]
            left, right, child = struct.unpack_from("<III", ent, 68)
            start = struct.unpack_from("<I", ent, 116)[0]
            size = struct.unpack_from("<Q", ent, 120)[0]
            self.dirs.append({"idx": i, "name": name, "type": typ, "left": left, "right": right, "child": child, "start": start, "size": size})

    def walk(self, idx=0, prefix=""):
        def rec(i, p):
            if i == FREE or i >= len(self.dirs):
                return
            e = self.dirs[i]
            yield from rec(e["left"], p)
            path = p + "/" + e["name"] if p else e["name"]
            yield i, path, e
            if e["child"] != FREE:
                yield from rec(e["child"], path)
            yield from rec(e["right"], p)
        if self.dirs[idx]["child"] != FREE:
            yield from rec(self.dirs[idx]["child"], prefix or self.dirs[idx]["name"])

    def read_stream(self, idx: int) -> bytes:
        e = self.dirs[idx]
        if e["type"] == 2 and e["size"] < self.mini_cutoff:
            return self._read_mini_stream(e["start"], e["size"])
        return self._read_big_stream(e["start"], e["size"])


def decode_msg_string(b: bytes) -> str:
    if not b:
        return ""
    for enc in ("utf-16le", "utf-8", "latin1"):
        try:
            return clean_text(b.decode(enc, "ignore").replace("\x00", ""))
        except Exception:
            pass
    return ""


def read_msg_bytes(data: bytes, filename: str) -> Tuple[str, str, List[Tuple[str, bytes]]]:
    subject, body = "", ""
    attachments: List[Tuple[str, bytes]] = []

    # Preferred library path when available.
    if extract_msg is not None:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
                tmp.write(data)
                msg_path = tmp.name
            try:
                msg = extract_msg.Message(msg_path)
                subject = clean_text(getattr(msg, "subject", "") or filename)
                body = clean_text(getattr(msg, "body", "") or "")
                for att in getattr(msg, "attachments", []) or []:
                    name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "attachment"
                    att_data = getattr(att, "data", None)
                    if callable(att_data):
                        att_data = att_data()
                    if isinstance(att_data, bytes):
                        attachments.append((name, att_data))
            finally:
                try:
                    os.remove(msg_path)
                except Exception:
                    pass
        except Exception:
            pass

    # Robust fallback: parse OLE streams directly.
    if not subject or not attachments:
        cfb = CFB(data)
        names: Dict[str, str] = {}
        bins: Dict[str, bytes] = {}
        for idx, path, e in cfb.walk():
            if e["type"] != 2:
                continue
            sname = path.split("/")[-1]
            raw = cfb.read_stream(idx)
            if sname == "__substg1.0_0037001F":
                subject = subject or decode_msg_string(raw)
            elif sname in ("__substg1.0_1000001F", "__substg1.0_1000001E"):
                body = body or decode_msg_string(raw)
            elif "/__attach_version1.0_" in path:
                base = path.split("/__substg1.0_")[0]
                if sname in ("__substg1.0_3707001F", "__substg1.0_3704001F"):
                    names[base] = decode_msg_string(raw) or names.get(base, "attachment")
                elif sname == "__substg1.0_37010102":
                    bins[base] = raw
        for base, b in bins.items():
            attachments.append((names.get(base, "attachment"), b))

    return subject or filename, body, attachments


def pdf_bytes_to_text(pdf_bytes: bytes) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        pages = []
        for page in reader.pages:
            pages.append(page.extract_text() or "")
        return clean_text("\n".join(pages))
    except Exception:
        return ""

# ------------------------- Voelker PDF parser -------------------------

def parse_city_state_zip(line: str) -> Tuple[str, str, str, str]:
    line = clean_text(line).replace(",", " ")
    m = re.search(rf"(.+?)\s+({STATE_RE})\s+(\d{{5}})(?:-\d{{4}})?\b", line, re.I)
    if m:
        return clean_text(m.group(1)), m.group(2).upper(), m.group(3), "USA"
    return "", "", "", ""


def find_after_label(lines: List[str], label: str) -> str:
    for i, line in enumerate(lines):
        if re.fullmatch(label, line, re.I):
            for j in range(i + 1, min(i + 8, len(lines))):
                val = clean_text(lines[j])
                if val and not re.fullmatch(r"Page|Salesperson|Cust #|Terms|Quantity|Quoted By|Ship Via|Ppd/Col|Shipped From", val, re.I):
                    return val
    return ""


def parse_pdf_text(pdf_text: str, fallback_subject: str = "") -> Dict[str, str]:
    text = clean_text(pdf_text)
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]
    joined = "\n".join(lines)

    d = {col: "" for col in TEMPLATE_COLUMNS}
    d["ReferralEmail"] = ""
    d["DemoQuote"] = "No"
    d["Country"] = "USA"

    d["QuoteNumber"] = get_regex(r"\b\d{2}/(\d{5,})\b", joined) or get_regex(r"Quote\s*#\s*(\d{5,})", fallback_subject)

    date_candidates = re.findall(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", joined)
    if date_candidates:
        d["QuoteDate"] = date_candidates[0]
    if len(date_candidates) > 1:
        d["QuoteExpiration"] = date_candidates[1]

    subj_company = get_regex(r"Customer[:_]\s*(.+?)(?:\.msg)?$", fallback_subject) or get_regex(r"Customer_\s*([^\.]+)", fallback_subject)
    if subj_company:
        d["Company"] = subj_company

    # pypdf often puts the header in two compact lines:
    # 6/23/26 7/23/26 AUTHORIZATION Salesperson Cust# Terms
    # 01/133832 QuotedBy ShipVia Ppd/Col ShippedFrom
    # Header line example:
    # 6/23/26 7/23/26 Rob McCullough 9877 SPECIAL SEE BELOW
    # Older parser required terms like NET/CASH and missed SPECIAL terms.
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+(.+?)\s+(\d{3,})\s+([^\n]+)", joined, re.I)
    if m:
        d["QuoteDate"] = m.group(1)
        d["QuoteExpiration"] = m.group(2)
        d["ReferralManager"] = clean_text(m.group(3))
        d["CustomerNumber"] = m.group(4)

    m = re.search(r"\b\d{2}/\d{5,}\s+([A-Za-z]+\s+[A-Za-z]+)", joined)
    if m:
        d["Created_By"] = clean_text(m.group(1))

    # Fallbacks for alternate text extraction order.
    d["ReferralManager"] = d["ReferralManager"] or find_after_label(lines, r"Salesperson")
    d["CustomerNumber"] = d["CustomerNumber"] or find_after_label(lines, r"Cust #")
    d["Created_By"] = d["Created_By"] or find_after_label(lines, r"Quoted By")

    # Sold To: Voelker PDFs show two customer blocks: Sold To first, Ship To second.
    # The manual Voelker sheet uses Sold To, so choose the FIRST matching customer address block.
    company = re.escape(d["Company"]) if d["Company"] else r"[A-Z0-9 &\-.,]+"
    addr_matches = list(re.finditer(rf"({company})\s*\n([^\n]+)\s*\n([^\n]*\b(?:{STATE_RE})\s+\d{{5}}(?:-\d{{4}})?\b)", joined, re.I))
    if addr_matches:
        m = addr_matches[0]
        d["Company"] = clean_text(m.group(1))
        d["Address"] = clean_text(m.group(2))
        city, state, z, country = parse_city_state_zip(m.group(3))
        d["City"], d["State"], d["ZipCode"] = city, state, z
        # Country is not printed in the Voelker Sold To block; leave blank to match manual fetch.
        d["Country"] = ""

    # Fallback: find the first company occurrence after the quote header, then the city/state/zip line after it.
    if d["Company"] and not d["Address"]:
        idxs = [i for i, line in enumerate(lines) if line.upper() == d["Company"].upper()]
        if idxs:
            start_i = idxs[0]
            for j in range(start_i + 1, min(start_i + 8, len(lines))):
                city, state, z, country = parse_city_state_zip(lines[j])
                if city:
                    d["City"], d["State"], d["ZipCode"] = city, state, z
                    d["Country"] = ""
                    prev = lines[j - 1] if j - 1 > start_i else ""
                    if re.search(r"ATTN|ACCOUNTS PAYABLE|PLANT", prev, re.I) and j - 2 > start_i:
                        prev = lines[j - 2]
                    d["Address"] = prev
                    break

    return {k: clean_text(v) for k, v in d.items()}





def normalize_date(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ""
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%m/%-d/%y", "%-m/%d/%y"):
        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", value)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/20{m.group(3)}"
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", value)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    return value

def title_case_company(name: str) -> str:
    """Make ALL CAPS customer names look like manual entry while preserving punctuation."""
    name = clean_text(name)
    if not name:
        return ""
    # Keep common short business words readable.
    return " ".join(w.capitalize() if w.isupper() else w for w in name.split())


def extract_quote_total(pdf_text: str) -> float:
    """Return the final Quote Total/last total amount from the Voelker quote PDF."""
    text = clean_text(pdf_text)
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]

    # Best case: label followed by value in the next few lines.
    for i, line in enumerate(lines):
        if re.fullmatch(r"Quote Total", line, re.I):
            for nxt in lines[i + 1:i + 8]:
                if re.fullmatch(r"[\d,]+\.\d{2}", nxt):
                    return money_to_float(nxt)

    # pypdf often extracts the bottom totals as several money-only lines. Quote Total is usually
    # the last money-only amount before optional "Total Tariffs" text.
    money_lines = []
    for line in lines:
        if re.fullmatch(r"[\d,]+\.\d{2}", line):
            money_lines.append(line)
        elif re.match(r"^(Total Tariffs|Goods in this quote/order|Page|QUOTE)$", line, re.I) and money_lines:
            # do not reset; totals normally appear immediately before this.
            pass
    return money_to_float(money_lines[-1]) if money_lines else ""


def make_summary_row(header: Dict[str, str], pdf_text: str) -> Dict[str, str]:
    """Create one manual-style row per quote, without item details."""
    row = {col: "" for col in TEMPLATE_COLUMNS}
    row.update(header)

    # Manual file keeps the full quote number like 01/129486.
    full_q = get_regex(r"\b(\d{2}/\d{5,})\b", pdf_text)
    if full_q:
        row["QuoteNumber"] = full_q

    plain_q = row.get("QuoteNumber", "")
    if "/" in plain_q:
        plain_q = plain_q.split("/", 1)[1]
    row["PDF"] = f"VOELKER_{plain_q}.pdf" if plain_q else row.get("PDF", "")

    row["QuoteDate"] = normalize_date(row.get("QuoteDate", ""))
    row["Company"] = title_case_company(row.get("Company", ""))
    row["Address"] = title_case_company(row.get("Address", ""))
    row["City"] = title_case_company(row.get("City", ""))
    row["TotalSales"] = extract_quote_total(pdf_text)

    # Match manual yellow row: quote-level record only.
    for col in ["item_id", "item_desc", "Quantity", "UnitSales", "quote_line_no", "QuoteExpiration"]:
        row[col] = ""
    row["Brand"] = "Voelker Controls"
    row["DemoQuote"] = ""
    return row

def parse_items_from_pdf(pdf_text: str) -> List[Dict[str, str]]:
    text = clean_text(pdf_text)
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]
    joined = "\n".join(lines)
    items: List[Dict[str, str]] = []

    csz_indexes = [i for i, line in enumerate(lines) if parse_city_state_zip(line)[0]]
    # Pick the ship-to city/state/zip that is actually followed by line-item details.
    start = (csz_indexes[-1] + 1) if csz_indexes else 0
    for idx in csz_indexes:
        # The actual ship-to block is followed immediately by quantity or compact qty+item.
        nxt = lines[idx + 1] if idx + 1 < len(lines) else ""
        nxt2 = lines[idx + 2] if idx + 2 < len(lines) else ""
        if re.match(r"^(\d+(?:\.\d+)?)(?:\s+[A-Z0-9])?", nxt, re.I):
            start = idx + 1
            break
        if nxt.upper() not in ("", "PAGE", "QUOTE") and re.match(r"^\d+(?:\.\d+)?\s+[A-Z0-9][A-Z0-9_./-]*-", nxt2, re.I):
            start = idx + 2
            break

    detail_lines = []
    for line in lines[start:]:
        if re.match(r"^(DELIVERY:|DELIVERIS|Goods in this quote/order|Unit Price|SubTotal|Freight|Sales Tax|Quote Total|Total Tariffs|\*\*Continued\*\*)", line, re.I):
            break
        if re.match(r"^\d[\d,]*\.\d{2,4}(?:\s+EA\s+\d[\d,]*\.\d{2})?$", line, re.I):
            break
        detail_lines.append(line)
    detail_lines = [x for x in detail_lines if not re.fullmatch(r"Page|QUOTE|Voelker Controls Company", x, re.I)]

    # Price extraction independent of quantity detection.
    after_detail = lines[start + len(detail_lines):]
    unit_prices: List[str] = []
    ext_prices: List[str] = []
    nums_before_ea: List[str] = []
    nums_after_ea: List[str] = []
    seen_ea = False
    for line in after_detail:
        if re.fullmatch(r"EA", line, re.I):
            seen_ea = True
            continue
        if re.match(r"^(SubTotal|Freight|Sales Tax|Quote Total|Total Tariffs|\*\*Continued\*\*|Page|QUOTE)", line, re.I) and nums_after_ea:
            break
        if re.fullmatch(r"\d[\d,]*\.\d{2,4}", line):
            if not seen_ea:
                nums_before_ea.append(line)
            elif len(nums_after_ea) < len(nums_before_ea):
                nums_after_ea.append(line)
    unit_prices = nums_before_ea
    ext_prices = nums_after_ea

    mp_all = re.findall(r"\b([\d,]+\.\d{2,4})\s+EA\s+([\d,]+\.\d{2})\b", joined, re.I)
    if mp_all and (not unit_prices or not ext_prices):
        unit_prices = [a for a, _ in mp_all]
        ext_prices = [b for _, b in mp_all]

    item_count_hint = max(len(unit_prices), len(ext_prices), len(mp_all))

    # Quantity prefix is limited by item count hint so numeric-only item codes are not mistaken for quantities.
    qtys = []
    pos = 0
    while pos < len(detail_lines) and re.fullmatch(r"\d+(?:\.\d+)?", detail_lines[pos]):
        if item_count_hint and len(qtys) >= item_count_hint:
            break
        qtys.append(detail_lines[pos])
        pos += 1
    product_lines = detail_lines[pos:]
    product_lines = [x for x in product_lines if not re.match(r"QC\s*QUOTE#", x, re.I)]

    def looks_like_code(s: str) -> bool:
        s = clean_text(s)
        if not s or len(s) > 40 or '"' in s:
            return False
        if re.search(r"DELIVERY|FREIGHT|QUOTE DOES NOT|Incoming|Outgoing|Pallet|BLACK DIMPLE", s, re.I):
            return False
        if re.fullmatch(r"\d{3,6}", s):
            return True
        if "-" in s:
            return True
        if re.match(r"^1515\s+X\s+", s, re.I):
            return True
        return False

    parsed_products: List[Tuple[str, str]] = []

    # Compact first product line: "2 AE4-..." or "1 K-00027".
    compact_qtys = []
    compact_lines = []
    for line in product_lines:
        m = re.match(r"^(\d+(?:\.\d+)?)\s+([A-Z0-9][A-Z0-9\-_/\.]{2,})$", line, re.I)
        if m:
            compact_qtys.append(m.group(1))
            compact_lines.append(m.group(2))
        else:
            compact_lines.append(line)
    if compact_qtys and not qtys:
        qtys = compact_qtys
        product_lines = compact_lines

    # Use simple pairs only when the even lines look like item codes and odd lines look like descriptions.
    n_hint = item_count_hint or len(qtys)
    if n_hint and len(product_lines) >= n_hint * 2:
        can_pair = True
        for i in range(n_hint):
            if not looks_like_code(product_lines[2 * i]):
                can_pair = False
                break
        if can_pair:
            for i in range(n_hint):
                parsed_products.append((product_lines[2 * i], product_lines[2 * i + 1]))

    if not parsed_products:
        current_code = ""
        current_desc: List[str] = []
        for line in product_lines:
            if looks_like_code(line) and (not current_code or len(parsed_products) + 1 < max(n_hint, 1)):
                if current_code:
                    parsed_products.append((current_code, " ".join(current_desc)))
                current_code = line
                current_desc = []
            else:
                current_desc.append(line)
        if current_code:
            parsed_products.append((current_code, " ".join(current_desc)))

    n = max(len(parsed_products), len(qtys), len(unit_prices), len(ext_prices))
    for i in range(n):
        code = parsed_products[i][0] if i < len(parsed_products) else ""
        desc = parsed_products[i][1] if i < len(parsed_products) else ""
        items.append({
            "quote_line_no": i + 1,
            "item_id": code,
            "item_desc": clean_text(desc),
            "Quantity": qty_to_number(qtys[i]) if i < len(qtys) else "",
            "UnitSales": money_to_float(unit_prices[i]) if i < len(unit_prices) else "",
            "TotalSales": money_to_float(ext_prices[i]) if i < len(ext_prices) else "",
        })

    if not items:
        items.append({"quote_line_no": 1, "QuoteComment": "Line item not detected - review PDF manually."})
    return items

def parse_one_uploaded_msg(uploaded_file, output_mode: str = "summary") -> List[Dict[str, str]]:
    data = uploaded_file.getvalue()
    subject, body, attachments = read_msg_bytes(data, uploaded_file.name)

    pdfs = [(name, b) for name, b in attachments if b[:5] == b"%PDF-" or name.lower().endswith(".pdf")]
    if not pdfs:
        raise RuntimeError("No PDF attachment found inside MSG.")

    all_rows = []
    for pdf_name, pdf_bytes in pdfs:
        pdf_text = pdf_bytes_to_text(pdf_bytes)
        if not pdf_text:
            raise RuntimeError(f"Could not read PDF text from {pdf_name}. Make sure pypdf is installed.")
        header = parse_pdf_text(pdf_text, subject)
        header["PDF"] = pdf_name if pdf_name.lower().endswith(".pdf") else f"Quote_{header.get('QuoteNumber','')}.pdf"

        # Extract contact email/phone from email body if present.
        if body:
            header["ContactEmail"] = get_regex(r"([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", body)
            phone = get_regex(r"(?:Phone|Tel|Cell|Mobile)\s*[:\-]?\s*(\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4})", body)
            header["ContactPhone"] = normalize_us_phone(phone)

        if output_mode == "summary":
            all_rows.append(make_summary_row(header, pdf_text))
        else:
            for item in parse_items_from_pdf(pdf_text):
                row = {col: "" for col in TEMPLATE_COLUMNS}
                row.update(header)
                row.update(item)
                row["DemoQuote"] = "No"
                all_rows.append(row)
    return all_rows


def write_output(template_file, rows: List[Dict[str, str]]) -> bytes:
    df_new = pd.DataFrame(rows)
    for col in TEMPLATE_COLUMNS:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[TEMPLATE_COLUMNS]

    if template_file is not None:
        try:
            base = pd.read_excel(io.BytesIO(template_file.getvalue()), dtype=str)
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
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 10), 42)
    return out.getvalue()

# ------------------------- UI -------------------------
with st.sidebar:
    st.header("Upload")
    msg_files = st.file_uploader("Voelker .msg files", type=["msg"], accept_multiple_files=True)
    template_file = st.file_uploader("Volkr.xlsx template", type=["xlsx"])
    output_mode = st.radio(
        "Output style",
        ["summary", "line_items"],
        index=0,
        format_func=lambda x: "One row per quote (matches manual yellow)" if x == "summary" else "Line item rows",
    )

if PdfReader is None:
    st.error("Missing dependency: pypdf. Add pypdf to requirements.txt and redeploy.")

if msg_files:
    if st.button("Extract Voelker Quotes", type="primary"):
        all_rows = []
        errors = []
        progress = st.progress(0)
        for i, f in enumerate(msg_files, start=1):
            try:
                all_rows.extend(parse_one_uploaded_msg(f, output_mode=output_mode))
            except Exception as e:
                errors.append(f"{f.name}: {e}")
            progress.progress(i / len(msg_files))

        if all_rows:
            df_preview = pd.DataFrame(all_rows)
            st.success(f"Extracted {len(all_rows)} quote line row(s) from {len(msg_files)} email(s).")
            st.dataframe(df_preview[[c for c in TEMPLATE_COLUMNS if c in df_preview.columns]], use_container_width=True)
            output = write_output(template_file, all_rows)
            st.download_button(
                "Download completed Volkr output",
                data=output,
                file_name=f"voelker_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("No rows extracted. See errors below.")
        if errors:
            st.warning("Files needing review:")
            for err in errors:
                st.write("- " + err)
else:
    st.info("Upload one or more Voelker .msg files to start.")
