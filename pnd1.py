#!/usr/bin/env python3
"""
PND1 (ภ.ง.ด.1 ใบแนบ) PDF Data Extractor

Extracts structured employee withholding tax records from PND1 attachment PDFs:
- ลำดับ (No.)
- เลขประจำตัวผู้เสียภาษีอากร (Tax ID)
- ชื่อผู้มีเงินได้ (Employee Name)
- วัน เดือน ปี ที่จ่าย (Payment Date)
- จำนวนเงินได้ที่จ่ายในครั้งนี้ (Income Amount)
- จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้ (Withholding Tax Amount)
- เงื่อนไข (Condition)
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import pdfplumber


DEFAULT_PDF_PATH = "/Users/gu/Workspace/personal/banchee/data/sp-lakeside/01/ภงด.1-01/P010012603563_20260203_180147_attach.pdf"

COLUMNS = [
    "ลำดับ",
    "เลขประจำตัวผู้เสียภาษีอากร",
    "ชื่อผู้มีเงินได้",
    "วัน เดือน ปี ที่จ่าย",
    "จำนวนเงินได้ที่จ่ายในครั้งนี้",
    "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้",
    "เงื่อนไข",
]


def clean_tax_id_and_name(cell_text: str) -> tuple[str, str]:
    """Parse cell containing Tax ID and Name into clean strings."""
    if not cell_text:
        return "", ""

    lines = [line.strip() for line in cell_text.split("\n") if line.strip()]
    tax_id = ""
    raw_name = ""

    if len(lines) >= 2:
        tax_id = re.sub(r"\s+", "", lines[0])
        raw_name = " ".join(lines[1:])
    elif len(lines) == 1:
        # Check if digits are present
        digits = re.sub(r"\D", "", lines[0])
        if len(digits) == 13:
            tax_id = digits
            raw_name = re.sub(r"[\d\s]+", " ", lines[0]).strip()
        else:
            raw_name = lines[0]

    # Clean Thai name prefix and labels
    # e.g., 'ชื่อนาง Su Su ชื่อสกุล Thein' -> 'นาง Su Su Thein'
    # e.g., 'ชื่อนาย สมรักษ์ ชื่อสกุล นันตะนุ' -> 'นาย สมรักษ์ นันตะนุ'
    name = re.sub(r"^ชื่อ\s*", "", raw_name)
    name = re.sub(r"\s*ชื่อสกุล\s*", " ", name).strip()
    name = re.sub(r"\s+", " ", name)

    return tax_id, name


def parse_amount(main_val: Any, dec_val: Any = None) -> float:
    """Parse integer and decimal parts into a float amount."""
    if main_val is None:
        return 0.0
    main_str = str(main_val).strip().replace(",", "")
    if dec_val is not None:
        dec_str = str(dec_val).strip().replace(",", "")
        if dec_str and not main_str.endswith("." + dec_str):
            main_str = f"{main_str}.{dec_str}"
    try:
        return float(main_str)
    except ValueError:
        return 0.0


def extract_pnd1_data(pdf_path: str | Path) -> pd.DataFrame:
    """Extract PND1 attachment table from PDF into a pandas DataFrame."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    records: List[Dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # Ensure table belongs to the employee details table
                table_text = " ".join([str(c) for r in table for c in r if c])
                if (
                    "รายละเอียดเกี่ยวกับการจ่ายเงิน" not in table_text
                    and "เลขประจำตัวผู้เสียภาษีอากร" not in table_text
                ):
                    continue

                for row in table:
                    if not row or len(row) < 8:
                        continue

                    idx_str = (row[0] or "").strip()
                    if not idx_str.isdigit():
                        continue

                    # Validate date format DD/MM/YYYY
                    pay_date = (row[2] or "").strip()
                    if not re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", pay_date):
                        continue

                    tax_id, name = clean_tax_id_and_name(row[1] or "")
                    income = parse_amount(row[3], row[4])
                    tax = parse_amount(row[5], row[6])
                    condition = (row[7] or "").strip()

                    records.append(
                        {
                            "ลำดับ": int(idx_str),
                            "เลขประจำตัวผู้เสียภาษีอากร": tax_id,
                            "ชื่อผู้มีเงินได้": name,
                            "วัน เดือน ปี ที่จ่าย": pay_date,
                            "จำนวนเงินได้ที่จ่ายในครั้งนี้": income,
                            "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้": tax,
                            "เงื่อนไข": condition,
                        }
                    )

    df = pd.DataFrame(records, columns=COLUMNS)
    return df


def main():
    parser = argparse.ArgumentParser(
        description="Extract employee tax data from PND1 (ภ.ง.ด.1) attachment PDF."
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="file",
        help="Path to PND1 attachment PDF file.",
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default=None,
        help=f"Path to PND1 attachment PDF file (default: {DEFAULT_PDF_PATH})",
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
        df = extract_pnd1_data(pdf_path)
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
        # Format terminal display
        pd.set_option("display.max_columns", None)
        pd.set_option("display.max_rows", None)
        pd.set_option("display.width", 1000)
        pd.set_option("display.unicode.east_asian_width", True)
        print(f"\n--- Extracted Data ({len(df)} records) from {pdf_path} ---")
        print(df.to_string(index=False))

        total_income = df["จำนวนเงินได้ที่จ่ายในครั้งนี้"].sum()
        total_tax = df["จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้"].sum()
        print("-" * 80)
        print(f"ยอดรวมเงินได้: {total_income:,.2f} บาท | ยอดรวมภาษีหัก ณ ที่จ่าย: {total_tax:,.2f} บาท\n")


if __name__ == "__main__":
    main()
