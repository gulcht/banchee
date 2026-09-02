from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Optional, List, Dict, Any
import pdfplumber
import pandas as pd


STATEMENT_COLUMNS = [
    "Date/Time",
    "Channel",
    "Code",
    "Cheque No",
    "Withdrawal (Debit)",
    "Deposit (Credit)",
    "Balance",
    "Description",
    "Note",
]

# Regex pattern to match dates like 01/01/2026 01:31 or 01/01/26 01:31
DATE_TIME_PATTERN = re.compile(r"^(\d{1,2}/\d{1,2}/\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)")

# Regex pattern for amounts with commas and decimals, e.g. 6,109.00 or -1,500.00
AMOUNT_PATTERN = re.compile(r"^-?[\d,]+\.\d{2}$")


def parse_statement_line(line: str) -> Optional[Dict[str, str]]:
    """Parse a single transaction line into structured columns."""
    line = line.strip()
    if not line:
        return None

    # Check if line starts with Date/Time
    dt_match = DATE_TIME_PATTERN.match(line)
    date_time = ""
    rest = line

    if dt_match:
        date_time = dt_match.group(1)
        rest = line[len(date_time) :].strip()

    tokens = rest.split()
    if not tokens:
        if date_time:
            return {"Date/Time": date_time, "is_date_only": True}
        return None

    # Find amount tokens (Withdrawal/Deposit, Balance)
    amount_indices = [i for i, tok in enumerate(tokens) if AMOUNT_PATTERN.match(tok)]

    channel = ""
    code = ""
    cheque_no = ""
    withdrawal = ""
    deposit = ""
    balance = ""
    desc_parts = []
    note_parts = []

    if amount_indices:
        first_amt_idx = amount_indices[0]
        # Text before the first amount is typically Channel / Code / Cheque No
        pre_tokens = tokens[:first_amt_idx]
        if len(pre_tokens) >= 1:
            channel = pre_tokens[0]
        if len(pre_tokens) >= 2:
            code = pre_tokens[1]
        if len(pre_tokens) >= 3:
            cheque_no = " ".join(pre_tokens[2:])

        # Extract amounts
        amounts = [tokens[idx] for idx in amount_indices]
        if len(amounts) == 1:
            # Only balance or one amount
            balance = amounts[0]
        elif len(amounts) == 2:
            # Amount and Balance
            amt, balance = amounts[0], amounts[1]
            if amt.startswith("-"):
                withdrawal = amt.replace("-", "")
            else:
                deposit = amt
        elif len(amounts) >= 3:
            withdrawal = amounts[0]
            deposit = amounts[1]
            balance = amounts[2]

        # Everything after the last amount token is Description / Note
        last_amt_idx = amount_indices[-1]
        post_tokens = tokens[last_amt_idx + 1 :]

        # Look for Note prefix (e.g. Ref1:, Ref:, Note:, etc.)
        note_start_idx = None
        for i, tok in enumerate(post_tokens):
            if any(tok.startswith(prefix) for prefix in ["Ref", "Ref1:", "Ref2:", "Ref:", "Note:", "Memo:"]):
                note_start_idx = i
                break

        if note_start_idx is not None:
            desc_parts = post_tokens[:note_start_idx]
            note_parts = post_tokens[note_start_idx:]
        else:
            desc_parts = post_tokens

    else:
        # No numerical amounts found - skip header or non-transaction line
        return None

    return {
        "Date/Time": date_time,
        "Channel": channel,
        "Code": code,
        "Cheque No": cheque_no,
        "Withdrawal (Debit)": withdrawal,
        "Deposit (Credit)": deposit,
        "Balance": balance,
        "Description": " ".join(desc_parts),
        "Note": " ".join(note_parts),
    }


def extract_statement_data(pdf_path: Path) -> List[Dict[str, str]]:
    """Extract and parse statement transactions across all pages."""
    transactions: List[Dict[str, str]] = []
    pending_date_times: List[str] = []
    transaction_lines_without_date: List[Dict[str, str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(layout=False) or ""
            lines = text.split("\n")

            for raw_line in lines:
                cleaned = raw_line.strip()
                if not cleaned:
                    continue

                # Ignore known header boilerplate text
                if any(
                    hdr in cleaned
                    for hdr in [
                        "วัน/เวลา",
                        "Date/Time",
                        "จำนวนเงินที่หักบัญชี",
                        "Withdrawal",
                        "Deposit",
                        "Balance",
                        "เลขที่บัญชี",
                        "Account No",
                    ]
                ):
                    continue

                # Check if this line is strictly a Date/Time
                dt_match = DATE_TIME_PATTERN.match(cleaned)
                if dt_match and len(cleaned.split()) <= 2:
                    pending_date_times.append(dt_match.group(1))
                    continue

                parsed = parse_statement_line(cleaned)
                if parsed:
                    if parsed.get("is_date_only"):
                        pending_date_times.append(parsed["Date/Time"])
                    elif parsed["Date/Time"]:
                        transactions.append(parsed)
                    else:
                        transaction_lines_without_date.append(parsed)

    # If dates were listed separately from transaction records, zip them together
    if pending_date_times and transaction_lines_without_date:
        for dt, tx in zip(pending_date_times, transaction_lines_without_date):
            tx["Date/Time"] = dt
            transactions.append(tx)
    elif transaction_lines_without_date:
        transactions.extend(transaction_lines_without_date)

    return transactions


def pdf_to_excel(pdf_path: str, output_excel_path: Optional[str] = None) -> Path:
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if output_excel_path is None:
        output_excel_path = pdf_file.with_suffix(".xlsx")
    else:
        output_excel_path = Path(output_excel_path)

    print(f"Parsing statement from {pdf_file.name}...")
    records = extract_statement_data(pdf_file)

    if not records:
        print("No transactions matched statement pattern. Trying generic table extraction...")
        # Fallback to generic table extraction
        all_rows = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table or []:
                        cleaned_row = [" ".join(str(c).split()) if c else "" for c in row]
                        if any(cleaned_row):
                            all_rows.append(cleaned_row)
        if all_rows:
            df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
        else:
            raise ValueError("No table data could be extracted.")
    else:
        df = pd.DataFrame(records, columns=STATEMENT_COLUMNS)

    df.to_excel(output_excel_path, index=False, engine="openpyxl")
    print(f"Successfully exported {len(df)} transactions to: {output_excel_path}")
    return output_excel_path


def main():
    parser = argparse.ArgumentParser(
        description="Convert bank PDF statement to Excel (.xlsx) with clean columns"
    )
    parser.add_argument("pdf_path", nargs="?", help="Path to PDF statement")
    parser.add_argument("-o", "--output", dest="output_path", help="Output Excel file path")

    args = parser.parse_args()

    if not args.pdf_path:
        pdf_files = list(Path(".").glob("*.pdf"))
        if pdf_files:
            pdf_to_excel(str(pdf_files[0]), args.output_path)
        else:
            print("Usage: python main.py <statement.pdf> [-o output.xlsx]")
    else:
        pdf_to_excel(args.pdf_path, args.output_path)


if __name__ == "__main__":
    main()
