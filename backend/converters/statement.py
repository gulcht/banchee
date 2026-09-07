#!/usr/bin/env python3
"""
Bank Statement (ประเภทบัญชี: เดินสะพัด / Current Account) PDF Data Extractor

Extracts structured transaction records from bank statement PDFs (ประเภทบัญชีเดินสะพัด):
- Date/Time (วัน/เวลา)
- Channel (ช่องทาง)
- Code (รายการ)
- Cheque No (เลขที่เช็ค)
- Withdrawal (Debit) (จำนวนเงินที่หักบัญชี)
- Deposit (Credit) (จำนวนเงินนำเข้าบัญชี)
- Balance (ยอดเงินคงเหลือ)
- Description (รายละเอียด)
- Note (บันทึกช่วยจำ)

Account Metadata (ข้อมูลบัญชี):
- ประเภทบัญชี: เดินสะพัด (Current Account)
- ชื่อบัญชี (Account Name)
- เลขที่บัญชี (Account No.)
- สาขา (Branch Name / Code)
- สกุลเงิน (Currency)
"""

from __future__ import annotations

import argparse
import io
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd
import pdfplumber

# Account type designation for this statement converter
ACCOUNT_TYPE = "เดินสะพัด"

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

# Known withdrawal channels/codes across Thai banks (e.g. SCB, KBank)
WITHDRAWAL_CHANNELS = {"XW", "CW", "DW", "PW", "SW", "TW", "DD", "FE", "IT", "EDC"}
DEPOSIT_CHANNELS = {"XP", "XD", "CD", "PD", "SD", "TD", "CR"}


def parse_statement_line(line: str) -> Optional[Dict[str, str]]:
    """Parse a single transaction line into structured columns."""
    line = line.strip()
    if not line:
        return None

    # Skip footer / summary lines
    if "รายการ" in line or "Total Credit" in line or "Total Debit" in line or "Items" in line:
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

        # Skip if channel or code is a summary indicator
        if channel in ["รายการ", "Total"] or "รายการ" in code:
            return None

        # Extract amounts
        amounts = [tokens[idx] for idx in amount_indices]
        if len(amounts) == 1:
            # Only balance
            balance = amounts[0]
        elif len(amounts) == 2:
            # Transaction Amount and Ending Balance
            amt, balance = amounts[0], amounts[1]
            amt_clean = amt.replace("-", "")

            # Classify Debit vs Credit based on Channel/Sign
            if (
                amt.startswith("-")
                or channel.upper() in WITHDRAWAL_CHANNELS
                or channel.upper().endswith("W")
            ):
                withdrawal = amt_clean
            elif (
                channel.upper() in DEPOSIT_CHANNELS
                or channel.upper().endswith("D")
                or channel.upper().endswith("P")
            ):
                deposit = amt_clean
            else:
                deposit = amt_clean
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
        # No numerical amounts found
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


def extract_statement_metadata(pdf_source: Union[Path, str, io.BytesIO]) -> Dict[str, str]:
    """Extract bank statement header metadata (Account Type: เดินสะพัด, Account Name, No, etc.)."""
    metadata: Dict[str, str] = {
        "account_type": ACCOUNT_TYPE,
    }

    try:
        with pdfplumber.open(pdf_source) as pdf:
            if not pdf.pages:
                return metadata
            # Header is typically on the first page
            first_page_text = pdf.pages[0].extract_text(layout=False) or ""

            # Check for account type (defaults to เดินสะพัด)
            type_match = re.search(r"ประเภทบัญชี:\s*([^\s\n\r]+)", first_page_text)
            if type_match:
                metadata["account_type"] = type_match.group(1).strip()

            # Account Name
            name_match = re.search(r"ชื่อบัญชี:\s*([^\n\r]+?)(?:\s+ประเภทบัญชี:|\n|$)", first_page_text)
            if name_match:
                metadata["account_name"] = name_match.group(1).strip()

            # Account No
            acc_no_match = re.search(r"เลขที่บัญชี:\s*([^\s\n\r]+)", first_page_text)
            if acc_no_match:
                metadata["account_no"] = acc_no_match.group(1).strip()

            # Branch
            branch_match = re.search(r"ชื่อสาขา/รหัส:\s*([^\n\r]+?)(?:\s+สกุลเงิน:|\n|$)", first_page_text)
            if branch_match:
                metadata["branch"] = branch_match.group(1).strip()

            # Currency
            curr_match = re.search(r"สกุลเงิน:\s*([A-Z]{3})", first_page_text)
            if curr_match:
                metadata["currency"] = curr_match.group(1).strip()

            # Start and End dates
            start_match = re.search(r"เริ่ม\s+(\d{1,2}/\d{1,2}/\d{4})", first_page_text)
            if start_match:
                metadata["start_date"] = start_match.group(1).strip()
            end_match = re.search(r"สิ้นสุด\s+(\d{1,2}/\d{1,2}/\d{4})", first_page_text)
            if end_match:
                metadata["end_date"] = end_match.group(1).strip()
    except Exception:
        pass

    return metadata


def extract_statement_data(pdf_path: Union[Path, str, io.BytesIO]) -> List[Dict[str, str]]:
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

                # Ignore known header and summary boilerplate text
                if any(
                    hdr in cleaned
                    for hdr in [
                        "วัน/เวลา",
                        "Date/Time",
                        "จำนวนเงินที่หักบัญชี",
                        "จำนวนเงินนำเข้าบัญชี",
                        "Total Credit Amount",
                        "Total Debit Amount",
                        "Withdrawal",
                        "Deposit",
                        "Balance",
                        "เลขที่บัญชี",
                        "Account No",
                        "ยอดเงินคงเหลือ",
                        "รวมทั้งสิ้น",
                        "รายการ (Items)",
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


def format_and_clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Format Date/Time column to datetime and amount columns to numeric types."""
    if "Date/Time" in df.columns:
        df["Date/Time"] = pd.to_datetime(
            df["Date/Time"], dayfirst=True, errors="coerce"
        )

    numeric_columns = ["Withdrawal (Debit)", "Deposit (Credit)", "Balance"]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def extract_statement_dataframe(
    pdf_source: Union[Path, str, io.BytesIO]
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Extract transactions into a cleaned DataFrame along with statement metadata."""
    metadata = extract_statement_metadata(pdf_source)
    records = extract_statement_data(pdf_source)

    if not records:
        all_rows = []
        with pdfplumber.open(pdf_source) as pdf:
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
            df = pd.DataFrame(columns=STATEMENT_COLUMNS)
    else:
        df = pd.DataFrame(records, columns=STATEMENT_COLUMNS)

    df = format_and_clean_dataframe(df)
    return df, metadata


def save_to_excel(df: pd.DataFrame, target: Any) -> None:
    """Export statement DataFrame to Excel with standard column formatting."""
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Statement")
        worksheet = writer.sheets["Statement"]

        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            max_len = max(
                len(str(col_name)),
                max((len(str(val or "")) for val in df[col_name]), default=0),
            )
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

            if col_name == "Date/Time":
                for cell in worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        if c.value is not None:
                            c.number_format = "DD/MM/YYYY HH:MM"
            elif col_name in ["Withdrawal (Debit)", "Deposit (Credit)", "Balance"]:
                for cell in worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        if c.value is not None:
                            c.number_format = "#,##0.00"


def save_to_csv(df: pd.DataFrame, target: Any) -> None:
    """Export statement DataFrame to CSV UTF-8 with BOM."""
    df.to_csv(target, index=False, encoding="utf-8-sig")


def parse_statement_month_year(
    df: Optional[pd.DataFrame] = None, metadata: Optional[Dict[str, str]] = None
) -> Tuple[Optional[str], Optional[str]]:
    """Extract month (MM) and year (YYYY) from statement metadata or transaction dates."""
    if metadata and metadata.get("start_date"):
        parts = metadata["start_date"].split("/")
        if len(parts) == 3:
            return parts[1], parts[2]
    if df is not None and not df.empty and "Date/Time" in df.columns:
        valid_dates = df["Date/Time"].dropna()
        if not valid_dates.empty:
            first_dt = valid_dates.iloc[0]
            if hasattr(first_dt, "strftime"):
                return first_dt.strftime("%m"), first_dt.strftime("%Y")
    return None, None


def format_statement_export_name(
    metadata: Optional[Dict[str, str]] = None,
    ext: str = "xlsx",
    df: Optional[pd.DataFrame] = None,
) -> str:
    """Generate export filename e.g. statement-current-01-2026.xlsx."""
    clean_ext = ext.lstrip(".")
    month, year = parse_statement_month_year(df=df, metadata=metadata)
    if month and year:
        return f"statement-current-{month}-{year}.{clean_ext}"
    return f"statement-current.{clean_ext}"


def pdf_to_excel(pdf_path: str, output_excel_path: Optional[str] = None) -> Path:
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if output_excel_path is None:
        output_excel_path = pdf_file.with_suffix(".xlsx")
    else:
        output_excel_path = Path(output_excel_path)

    metadata = extract_statement_metadata(pdf_file)
    acc_type = metadata.get("account_type", ACCOUNT_TYPE)
    print(f"Parsing statement (ประเภทบัญชี: {acc_type}) from {pdf_file.name}...")
    df, _ = extract_statement_dataframe(pdf_file)

    if df.empty:
        raise ValueError("No table data could be extracted.")

    save_to_excel(df, output_excel_path)
    print(f"Successfully exported {len(df)} transactions to: {output_excel_path}")
    return output_excel_path


def main():
    parser = argparse.ArgumentParser(
        description="Convert bank PDF statement (ประเภทบัญชี: เดินสะพัด / Current Account) to Excel (.xlsx) with clean columns"
    )
    parser.add_argument("pdf_path", nargs="?", help="Path to PDF statement")
    parser.add_argument("-o", "--output", dest="output_path", help="Output Excel file path")

    args = parser.parse_args()

    if not args.pdf_path:
        pdf_files = list(Path(".").rglob("*.PDF")) + list(Path(".").rglob("*.pdf"))
        if pdf_files:
            pdf_to_excel(str(pdf_files[0]), args.output_path)
        else:
            print("Usage: python -m utils.convert_statement <statement.pdf> [-o output.xlsx]")
    else:
        pdf_to_excel(args.pdf_path, args.output_path)


if __name__ == "__main__":
    main()
