#!/usr/bin/env python3
"""
PND53 (ภ.ง.ด.53 ใบแนบ) PDF Data Extractor

Extracts structured withholding tax records (นิติบุคคล) from PND53 attachment PDFs:
- ลำดับ (No.)
- เลขประจำตัวผู้เสียภาษีอากร (Tax ID)
- ชื่อผู้มีเงินได้ (Company / Payee Name)
- สาขาที่ (Branch No.)
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


COLUMNS = [
    "ลำดับ",
    "เลขประจำตัวผู้เสียภาษีอากร",
    "ชื่อผู้มีเงินได้",
    "สาขาที่",
    "ที่อยู่",
    "วัน เดือน ปี ที่จ่าย",
    "ประเภทเงินได้",
    "อัตราภาษีร้อยละ",
    "จำนวนเงินที่จ่ายในครั้งนี้",
    "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้",
    "เงื่อนไข",
]


def clean_tax_id_name_address(cell_text: str) -> Tuple[str, str, str]:
    """Parse cell containing Tax ID, Company Name, and Address into clean strings."""
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

    # Clean prefix labels
    name = re.sub(r"^ชื่อ\s*", "", raw_name)
    name = re.sub(r"\s*ชื่อสกุล\s*", " ", name).strip()
    name = re.sub(r"\s+", " ", name)

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


def extract_pnd53_data(pdf_path: str | Path) -> pd.DataFrame:
    """Extract PND53 attachment table from PDF into a pandas DataFrame."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    records: List[Dict[str, Any]] = []
    current_company = {
        "ลำดับ": 0,
        "เลขประจำตัวผู้เสียภาษีอากร": "",
        "ชื่อผู้มีเงินได้": "",
        "สาขาที่": "00000",
        "ที่อยู่": "",
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # Ensure table belongs to the PND53 details table
                table_text = " ".join([str(c) for r in table for c in r if c])
                if (
                    "รายละเอียดเกี่ยวกับการจ่ายเงิน" not in table_text
                    and "เลขประจำตัวผู้เสียภาษีอากร" not in table_text
                    and "ภ.ง.ด.53" not in table_text
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

                    # Find date column index
                    date_idx = -1
                    for i, cell in enumerate(row):
                        if cell and re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", str(cell).strip()):
                            date_idx = i
                            break
                    if date_idx == -1:
                        continue

                    idx_str = (row[0] or "").strip()
                    if idx_str.isdigit():
                        tax_id, name, address = clean_tax_id_name_address(row[1] or "")
                        branch = (
                            re.sub(r"\s+", "", str(row[2] or ""))
                            if len(row) > 2 and row[2]
                            else "00000"
                        )
                        current_company = {
                            "ลำดับ": int(idx_str),
                            "เลขประจำตัวผู้เสียภาษีอากร": tax_id,
                            "ชื่อผู้มีเงินได้": name,
                            "สาขาที่": branch,
                            "ที่อยู่": address,
                        }

                    pay_date = str(row[date_idx]).strip()
                    income_type = str(row[date_idx + 1] or "").strip()

                    # Remaining columns after payment date and income type
                    rem = [c for c in row[date_idx + 2:] if c is not None and str(c).strip() != ""]

                    tax_rate = parse_amount(rem[0]) if len(rem) >= 1 else 0.0

                    if len(rem) == 6:
                        pay_amount = parse_amount(rem[1], rem[2])
                        tax_amount = parse_amount(rem[3], rem[4])
                        cond = str(rem[5]).strip()
                    elif len(rem) == 5:
                        pay_amount = parse_amount(rem[1], rem[2])
                        tax_amount = parse_amount(rem[3])
                        cond = str(rem[4]).strip()
                    elif len(rem) == 4:
                        pay_amount = parse_amount(rem[1])
                        tax_amount = parse_amount(rem[2])
                        cond = str(rem[3]).strip()
                    else:
                        pay_amount = parse_amount(rem[1]) if len(rem) > 1 else 0.0
                        tax_amount = parse_amount(rem[2]) if len(rem) > 2 else 0.0
                        cond = str(rem[-1]).strip() if len(rem) > 3 else ""

                    records.append(
                        {
                            "ลำดับ": current_company["ลำดับ"],
                            "เลขประจำตัวผู้เสียภาษีอากร": current_company["เลขประจำตัวผู้เสียภาษีอากร"],
                            "ชื่อผู้มีเงินได้": current_company["ชื่อผู้มีเงินได้"],
                            "สาขาที่": current_company["สาขาที่"],
                            "ที่อยู่": current_company["ที่อยู่"],
                            "วัน เดือน ปี ที่จ่าย": pay_date,
                            "ประเภทเงินได้": income_type,
                            "อัตราภาษีร้อยละ": tax_rate,
                            "จำนวนเงินที่จ่ายในครั้งนี้": pay_amount,
                            "จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้": tax_amount,
                            "เงื่อนไข": cond,
                        }
                    )

    df = pd.DataFrame(records, columns=COLUMNS)
    return df


def extract_month_year_from_df(df: pd.DataFrame) -> Optional[str]:
    """Extract MM-YYYY string (e.g., '01-2569', '01-2026') from payment dates in DataFrame."""
    if df.empty or "วัน เดือน ปี ที่จ่าย" not in df.columns:
        return None
    for date_val in df["วัน เดือน ปี ที่จ่าย"].dropna():
        match = re.search(r"^\d{1,2}/(\d{1,2})/(\d{2,4})", str(date_val).strip())
        if match:
            month_num = int(match.group(1))
            year_val = match.group(2)
            return f"{month_num:02d}-{year_val}"
    return None


def main():
    parser = argparse.ArgumentParser(
        description="Extract withholding tax data from PND53 (ภ.ง.ด.53) attachment PDF."
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="file",
        help="Path to PND53 attachment PDF file.",
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default=None,
        help="Path to PND53 attachment PDF file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Path to export file (.xlsx, .csv, or .json). If omitted, exports to <MM-YYYY>.xlsx.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print output in JSON format.",
    )

    args = parser.parse_args()
    pdf_path = args.file or args.pdf_path
    if not pdf_path:
        parser.print_help(sys.stderr)
        print(
            "\nError: PDF file is required. Please specify with -f / --file <path> or as an argument.",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        df = extract_pnd53_data(pdf_path)
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
        month_year = extract_month_year_from_df(df)
        default_filename = f"{month_year}.xlsx" if month_year else f"{Path(pdf_path).stem}.xlsx"
        out_path = Path(default_filename)
        df.to_excel(out_path, index=False)
        print(f"Exported {len(df)} rows to Excel: {out_path}")

        total_income = df["จำนวนเงินที่จ่ายในครั้งนี้"].sum()
        total_tax = df["จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้"].sum()
        print("-" * 100)
        print(f"ยอดรวมเงินได้: {total_income:,.2f} บาท | ยอดรวมภาษีหัก ณ ที่จ่าย: {total_tax:,.2f} บาท\n")


if __name__ == "__main__":
    main()
