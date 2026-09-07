#!/usr/bin/env python3
"""
SSO (สปส. 1-10 ส่วนที่ 2) PDF Data Extractor

Extracts structured employee social security contribution records from SSO 1-10 (ส่วนที่ 2) PDFs:
- ลำดับที่ (No.)
- เลขประจำตัวประชาชน (ID Card / SSN)
- ชื่อ-ชื่อสกุล (Employee Full Name)
- ค่าจ้าง (Wages)
- เงินสมทบ (Contribution)
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
    "ลำดับที่",
    "เลขประจำตัวประชาชน",
    "ชื่อ-ชื่อสกุล",
    "ค่าจ้าง",
    "เงินสมทบ",
]

THAI_MONTHS: Dict[str, str] = {
    "มกราคม": "01",
    "กุมภาพันธ์": "02",
    "มีนาคม": "03",
    "เมษายน": "04",
    "พฤษภาคม": "05",
    "มิถุนายน": "06",
    "กรกฎาคม": "07",
    "สิงหาคม": "08",
    "กันยายน": "09",
    "ตุลาคม": "10",
    "พฤศจิกายน": "11",
    "ธันวาคม": "12",
    "ม.ค.": "01",
    "ก.พ.": "02",
    "มี.ค.": "03",
    "เม.ย.": "04",
    "พ.ค.": "05",
    "มิ.ย.": "06",
    "ก.ค.": "07",
    "ส.ค.": "08",
    "ก.ย.": "09",
    "ต.ค.": "10",
    "พ.ย.": "11",
    "ธ.ค.": "12",
}


def parse_period(period_str: str) -> Tuple[Optional[str], Optional[str]]:
    """Extract month (MM) and Gregorian year (YYYY) from period string."""
    if not period_str:
        return None, None

    month: Optional[str] = None
    for m_name, m_num in THAI_MONTHS.items():
        if m_name in period_str:
            month = m_num
            break

    # Look for 4-digit year (Buddhist or Gregorian)
    year_match = re.search(r"(?:พ\.ศ\.\s*)?(\d{4})", period_str)
    year: Optional[str] = None
    if year_match:
        y_val = int(year_match.group(1))
        if y_val > 2400:  # Buddhist year
            y_val -= 543
        year = str(y_val)

    return month, year


def format_sso_export_name(metadata: Dict[str, str], ext: str = "xlsx") -> str:
    """Generate export filename in format sso-MM-YYYY.<ext>."""
    period_str = metadata.get("period", "") if metadata else ""
    month, year = parse_period(period_str)
    clean_ext = ext.lstrip(".")

    if month and year:
        return f"sso-{month}-{year}.{clean_ext}"
    return f"sso.{clean_ext}"



def clean_name(raw_name: str) -> str:
    """Clean extra spaces and trailing dashes in employee names."""
    if not raw_name:
        return ""
    name = re.sub(r"\s+", " ", raw_name).strip()
    if name.endswith(" -"):
        name = name[:-2].strip()
    return name


def parse_sso_line(line: str) -> Optional[Dict[str, Any]]:
    """Parse a single line from SSO 1-10 part 2 text."""
    line = line.strip()
    if not line:
        return None

    # Skip header, total and footer lines
    if any(
        skip_word in line
        for skip_word in [
            "แบบรายการแสดงการส่งเงินสมทบ",
            "แผ่นที่",
            "การนำส่งเงินสมทบ",
            "ชื่อสถานประกอบการ",
            "ลำดับที่สาขา",
            "ลำดับที่ เลขประจำตัวประชาชน",
            "คำนำหน้านาม",
            "ยอดรวม",
            "รวม",
            "คำชี้แจง",
            "คำเตือน",
            "หมายเลขธุรกรรม",
            "สาํ นกั งานประกนั สงัคม",
        ]
    ):
        return None

    # Pattern 1: Decimal format (e.g. 1 0020071214798 นางYAN YEW KYEIN - 18,500.00 875.00)
    m1 = re.match(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$",
        line,
    )
    if m1:
        seq, id_card, name, wage_str, contrib_str = m1.groups()
        wage = float(wage_str.replace(",", ""))
        contrib = float(contrib_str.replace(",", ""))
        return {
            "ลำดับที่": int(seq),
            "เลขประจำตัวประชาชน": str(id_card),
            "ชื่อ-ชื่อสกุล": clean_name(name),
            "ค่าจ้าง": wage,
            "เงินสมทบ": contrib,
        }

    # Pattern 2: Space separated Baht and Satang (e.g. 1 0020721056059 นางSU SU THEIN 12,000 00 600 00)
    m2 = re.match(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+(\d{2})\s+([\d,]+)\s+(\d{2})$",
        line,
    )
    if m2:
        seq, id_card, name, w_b, w_s, c_b, c_s = m2.groups()
        wage = float(f"{w_b.replace(',', '')}.{w_s}")
        contrib = float(f"{c_b.replace(',', '')}.{c_s}")
        return {
            "ลำดับที่": int(seq),
            "เลขประจำตัวประชาชน": str(id_card),
            "ชื่อ-ชื่อสกุล": clean_name(name),
            "ค่าจ้าง": wage,
            "เงินสมทบ": contrib,
        }

    return None


def extract_sso_data(pdf_source: str | Path | Any) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Extract SSO 1-10 part 2 records and metadata from PDF into a pandas DataFrame."""
    if isinstance(pdf_source, (str, Path)):
        pdf_path = Path(pdf_source)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        open_target = pdf_path
    else:
        open_target = pdf_source

    records: List[Dict[str, Any]] = []
    metadata: Dict[str, str] = {
        "company": "",
        "account_no": "",
        "branch": "",
        "period": "",
    }

    with pdfplumber.open(open_target) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.split("\n")

            for line in lines:
                cleaned = line.strip()

                # Extract metadata if available
                if "ชื่อสถานประกอบการ" in cleaned and not metadata["company"]:
                    comp_match = re.search(r"ชื่อสถานประกอบการ\s*[\.:]?\s*(.+?)(?:เลขที่บัญชี|$)", cleaned)
                    if comp_match:
                        metadata["company"] = re.sub(r"[\.]{2,}", "", comp_match.group(1)).strip()
                if "เลขที่บัญชี" in cleaned and not metadata["account_no"]:
                    acc_match = re.search(r"เลขที่บัญชี\s*[\.:]?\s*([\d\s\.]+)", cleaned)
                    if acc_match:
                        metadata["account_no"] = re.sub(r"\D", "", acc_match.group(1)).strip()
                if "การนำส่งเงินสมทบสำหรับค่าจ้างเดือน" in cleaned or "สำหรับค่าจ้างเดือน" in cleaned:
                    if not metadata["period"]:
                        metadata["period"] = cleaned

                record = parse_sso_line(cleaned)
                if record:
                    records.append(record)

    df = pd.DataFrame(records, columns=COLUMNS)
    return df, metadata


def save_to_excel(df: pd.DataFrame, output_target: str | Path | Any) -> None:
    """Save DataFrame to Excel with proper formatting."""
    with pd.ExcelWriter(output_target, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SSO")
        worksheet = writer.sheets["SSO"]

        # Adjust column widths and apply formatting
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            max_len = max(
                len(str(col_name)),
                max((len(str(val or "")) for val in df[col_name]), default=0),
            )
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

            # Format ID card as text to keep leading zeros
            if col_name == "เลขประจำตัวประชาชน":
                for cell in worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        if c.value is not None:
                            c.number_format = "@"
            # Format numbers as currency
            elif col_name in ["ค่าจ้าง", "เงินสมทบ"]:
                for cell in worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        if c.value is not None:
                            c.number_format = "#,##0.00"


def save_to_csv(df: pd.DataFrame, output_target: str | Path | Any) -> None:
    """Save DataFrame to CSV with UTF-8 BOM encoding."""
    csv_df = df.copy()
    csv_df["เลขประจำตัวประชาชน"] = csv_df["เลขประจำตัวประชาชน"].astype(str)
    csv_df.to_csv(output_target, index=False, encoding="utf-8-sig")


def main():
    parser = argparse.ArgumentParser(
        description="Extract employee data from SSO (สปส. 1-10 ส่วนที่ 2) PDF to Excel/CSV."
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="file",
        help="Path to SSO PDF file.",
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default=None,
        help="Path to SSO PDF file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Custom output file path (.xlsx or .csv).",
    )
    parser.add_argument(
        "--excel",
        action="store_true",
        help="Force export to Excel (.xlsx).",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="Force export to CSV (.csv).",
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
        df, metadata = extract_sso_data(pdf_path)
    except Exception as e:
        print(f"Error processing PDF: {e}", file=sys.stderr)
        sys.exit(1)

    if df.empty:
        print("No records found in PDF.", file=sys.stderr)
        sys.exit(0)

    input_stem = Path(pdf_path).stem

    if args.json:
        print(df.to_json(orient="records", force_ascii=False, indent=2))
        return

    # Default export filename sso-MM-YYYY
    default_excel_name = format_sso_export_name(metadata, ext="xlsx")
    default_csv_name = format_sso_export_name(metadata, ext="csv")

    # Handle explicit single output
    if args.output:
        out_path = Path(args.output)
        if out_path.suffix.lower() == ".xlsx":
            save_to_excel(df, out_path)
            print(f"Exported {len(df)} rows to Excel: {out_path}")
        elif out_path.suffix.lower() == ".csv":
            save_to_csv(df, out_path)
            print(f"Exported {len(df)} rows to CSV: {out_path}")
        elif out_path.suffix.lower() == ".json":
            df.to_json(out_path, orient="records", force_ascii=False, indent=2)
            print(f"Exported {len(df)} rows to JSON: {out_path}")
        else:
            save_to_excel(df, out_path)
            print(f"Exported {len(df)} rows to: {out_path}")
    elif args.excel and not args.csv:
        excel_path = Path(default_excel_name)
        save_to_excel(df, excel_path)
        print(f"Exported {len(df)} rows to Excel: {excel_path}")
    elif args.csv and not args.excel:
        csv_path = Path(default_csv_name)
        save_to_csv(df, csv_path)
        print(f"Exported {len(df)} rows to CSV: {csv_path}")
    else:
        # Default: Export to both Excel and CSV
        excel_path = Path(default_excel_name)
        csv_path = Path(default_csv_name)
        save_to_excel(df, excel_path)
        save_to_csv(df, csv_path)
        print(f"Exported {len(df)} rows to Excel: {excel_path}")
        print(f"Exported {len(df)} rows to CSV: {csv_path}")

    total_wages = df["ค่าจ้าง"].sum()
    total_contrib = df["เงินสมทบ"].sum()
    print("-" * 80)
    print(f"จำนวนรายการทั้งหมด: {len(df)} รายการ")
    print(f"ยอดรวมค่าจ้าง: {total_wages:,.2f} บาท | ยอดรวมเงินสมทบ: {total_contrib:,.2f} บาท")
    print("-" * 80)


if __name__ == "__main__":
    main()
