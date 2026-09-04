#!/usr/bin/env python3
"""
PND3 (ภ.ง.ด.3 ใบแนบ) PDF Data Extractor

Extracts structured withholding tax records (บุคคลธรรมดา) from PND3 attachment PDFs:
- ลำดับ (No.)
- เลขประจำตัวผู้เสียภาษีอากร (Tax ID)
- ชื่อผู้มีเงินได้ (Payee Name)
- ที่อยู่ (Address)
- วัน เดือน ปี ที่จ่าย (Payment Date)
- ประเภทเงินได้ (Income Type)
- อัตราภาษีร้อยละ (Tax Rate %)
- จำนวนเงินที่จ่ายในครั้งนี้ (Payment Amount)
- จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้ (Withholding Tax Amount)
- เงื่อนไข (Condition)
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber


DEFAULT_PDF_PATH = "data/P030017302507_20260213_104401_attach.pdf"

COLUMNS = [
    "ลำดับ",
    "เลขประจำตัวผู้เสียภาษีอากร",
    "ชื่อผู้มีเงินได้",
    "ที่อยู่",
    "วัน เดือน ปี ที่จ่าย",
    "ประเภทเงินได้",
    "อัตราภาษีร้อยละ",
    "จำนวนเงินที่จ่ายในครั้งนี้",
    "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้",
    "เงื่อนไข",
]


def clean_tax_id_name_address(cell_text: str) -> Tuple[str, str, str]:
    """Parse cell containing Tax ID, Name, and Address into clean strings."""
    if not cell_text:
        return "", "", ""

    lines = [line.strip() for line in cell_text.split("\n") if line.strip()]
    tax_id, raw_name, address = "", "", ""

    if len(lines) >= 1:
        digits = re.sub(r"\D", "", lines[0])
        if len(digits) == 13:
            tax_id = digits
        else:
            tax_id = lines[0]

    if len(lines) >= 2:
        raw_name = lines[1]
    if len(lines) >= 3:
        address = " ".join(lines[2:])

    # Clean Thai name prefix and labels
    # e.g., 'ชื่อนายถาวร ชื่อสกุลทับเนียม' -> 'นายถาวร ทับเนียม'
    # e.g., 'ชื่อคุณธงชัย ชื่อสกุลเจริญรัชเดช' -> 'คุณธงชัย เจริญรัชเดช'
    name = re.sub(r"^ชื่อ\s*", "", raw_name)
    name = re.sub(r"\s*ชื่อสกุล\s*", " ", name).strip()
    name = re.sub(r"\s+", " ", name)

    # Clean address prefix
    address = re.sub(r"^ที่อยู่\s*", "", address).strip()

    return tax_id, name, address


def parse_amount(val1: Any, val2: Any = None) -> float:
    """Parse amount from one or two cell values (handling space/comma formatted numbers)."""
    if not val1 and not val2:
        return 0.0
    s = f"{val1 or ''} {val2 or ''}".strip()
    parts = s.split()
    if len(parts) == 2 and parts[1].isdigit() and len(parts[1]) <= 2:
        s = f"{parts[0]}.{parts[1]}"
    elif len(parts) > 1:
        s = "".join(parts)

    clean_str = s.replace(",", "")
    try:
        return float(clean_str)
    except ValueError:
        return 0.0


def extract_pnd3_data(pdf_path: str | Path) -> pd.DataFrame:
    """Extract PND3 attachment table from PDF into a pandas DataFrame."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    records: List[Dict[str, Any]] = []
    current_person = {
        "ลำดับ": 0,
        "เลขประจำตัวผู้เสียภาษีอากร": "",
        "ชื่อผู้มีเงินได้": "",
        "ที่อยู่": "",
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # Ensure table belongs to the PND3 details table
                table_text = " ".join([str(c) for r in table for c in r if c])
                if (
                    "รายละเอียดเกี่ยวกับการจ่ายเงิน" not in table_text
                    and "เลขประจำตัวผู้เสียภาษีอากร" not in table_text
                ):
                    continue

                for row in table:
                    if not row or len(row) < 9:
                        continue

                    # Skip total and footer rows
                    first_cell = str(row[0] or "")
                    second_cell = str(row[1] or "")
                    if "รวมยอดเงินได้" in first_cell or "รวมยอดเงินได้" in second_cell:
                        continue

                    idx_str = (row[0] or "").strip()
                    pay_date = (row[2] or "").strip()

                    if idx_str.isdigit():
                        tax_id, name, address = clean_tax_id_name_address(row[1] or "")
                        current_person = {
                            "ลำดับ": int(idx_str),
                            "เลขประจำตัวผู้เสียภาษีอากร": tax_id,
                            "ชื่อผู้มีเงินได้": name,
                            "ที่อยู่": address,
                        }
                    elif not pay_date or not re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", pay_date):
                        # Not a valid new row or continuation row
                        continue

                    income_type = (row[3] or "").strip()
                    tax_rate = parse_amount(row[4])
                    payment_amount = parse_amount(row[5], row[6])
                    tax_amount = parse_amount(row[7], row[8])
                    condition = (row[9] or "").strip() if len(row) > 9 else ""

                    records.append(
                        {
                            "ลำดับ": current_person["ลำดับ"],
                            "เลขประจำตัวผู้เสียภาษีอากร": current_person["เลขประจำตัวผู้เสียภาษีอากร"],
                            "ชื่อผู้มีเงินได้": current_person["ชื่อผู้มีเงินได้"],
                            "ที่อยู่": current_person["ที่อยู่"],
                            "วัน เดือน ปี ที่จ่าย": pay_date,
                            "ประเภทเงินได้": income_type,
                            "อัตราภาษีร้อยละ": tax_rate,
                            "จำนวนเงินที่จ่ายในครั้งนี้": payment_amount,
                            "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้": tax_amount,
                            "เงื่อนไข": condition,
                        }
                    )

    df = pd.DataFrame(records, columns=COLUMNS)
    return df


def main():
    parser = argparse.ArgumentParser(
        description="Extract withholding tax data from PND3 (ภ.ง.ด.3) attachment PDF."
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="file",
        help="Path to PND3 attachment PDF file.",
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default=None,
        help=f"Path to PND3 attachment PDF file (default: {DEFAULT_PDF_PATH})",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Path to export file (.xlsx, .csv, or .json). If omitted, prints table to terminal.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print output in JSON format.",
    )

    args = parser.parse_args()
    pdf_path = args.file or args.pdf_path or DEFAULT_PDF_PATH

    try:
        df = extract_pnd3_data(pdf_path)
    except Exception as e:
        print(f"Error processing PDF: {e}", file=sys.stderr)
        sys.exit(1)

    if df.empty:
        print("No records found in PDF.", file=sys.stderr)
        sys.exit(0)

    # Output handling
    if args.output:
        out_path = Path(args.output)
        if out_path.suffix.lower() == ".xlsx":
            df.to_excel(out_path, index=False)
            print(f"Exported {len(df)} rows to Excel: {out_path}")
        elif out_path.suffix.lower() == ".csv":
            df.to_csv(out_path, index=False, encoding="utf-8-sig")
            print(f"Exported {len(df)} rows to CSV: {out_path}")
        elif out_path.suffix.lower() == ".json":
            df.to_json(out_path, orient="records", force_ascii=False, indent=2)
            print(f"Exported {len(df)} rows to JSON: {out_path}")
        else:
            df.to_excel(out_path, index=False)
            print(f"Exported {len(df)} rows to: {out_path}")
    elif args.json:
        print(df.to_json(orient="records", force_ascii=False, indent=2))
    else:
        pd.set_option("display.max_columns", None)
        pd.set_option("display.max_rows", None)
        pd.set_option("display.width", 1000)
        pd.set_option("display.unicode.east_asian_width", True)
        print(f"\n--- Extracted Data ({len(df)} records) from {pdf_path} ---")
        print(df.to_string(index=False))

        total_income = df["จำนวนเงินที่จ่ายในครั้งนี้"].sum()
        total_tax = df["จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้"].sum()
        print("-" * 100)
        print(f"ยอดรวมเงินได้: {total_income:,.2f} บาท | ยอดรวมภาษีหัก ณ ที่จ่าย: {total_tax:,.2f} บาท\n")


if __name__ == "__main__":
    main()
