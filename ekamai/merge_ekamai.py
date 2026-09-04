import argparse
import glob
import os
import re
import pandas as pd


def parse_date_ddmmyyyy(filename):
    """
    Extract DD/MM/YYYY from filename (e.g., '01012026.xlsx' -> '01/01/2026').
    """
    basename = os.path.basename(filename)
    match = re.match(r"^(\d{2})(\d{2})(\d{4})", basename)
    if match:
        day, month, year = match.groups()
        return f"{day}/{month}/{year}"
    return None


def merge_excel_files(input_dir="data/01-31012026", output_file="merged_sales_data.xlsx", short=False):
    files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
    if not files:
        print(f"No Excel files found in {input_dir}")
        return

    print(f"Found {len(files)} files to merge from '{input_dir}' (short_mode={short})...")

    merged_rows = []

    for file_path in files:
        filename = os.path.basename(file_path)
        date_str = parse_date_ddmmyyyy(filename)

        # Read data starting at header row (row index 2 / 3rd row in Excel)
        df = pd.read_excel(file_path, skiprows=2)

        if df.empty:
            continue

        bill_col = df.columns[0]  # Usually 'เลขที่บิล'

        # Filter: keep bill item rows (excluding summary and shift header rows)
        valid_mask = (
            df[bill_col].notna()
            & ~df[bill_col].astype(str).str.contains("รวมยอด|รวมทั้งสิ้น", na=False)
            & ~df[bill_col].astype(str).str.contains("ปิดรอบ", na=False)
        )
        bill_df = df[valid_mask].copy()

        # Add date column at the beginning
        bill_df.insert(0, "date", date_str)

        # Short mode filter
        if short:
            target_cols = ["date", "เลขที่บิล", "Cash", "Credit Card", "VISA", "Tranfer"]
            # Ensure required columns exist, fill with 0 / '-' if missing in some daily sheets
            for col in target_cols:
                if col not in bill_df.columns:
                    bill_df[col] = "-"
            bill_df = bill_df[target_cols]

        merged_rows.append(bill_df)
        print(f"  Processed {filename}: {len(bill_df)} rows (date: {date_str})")

    # Combine all DataFrames
    all_data = pd.concat(merged_rows, ignore_index=True)

    # Save to Excel
    all_data.to_excel(output_file, index=False)
    print(f"\n[DONE] Successfully merged {len(merged_rows)} files ({len(all_data)} rows) into: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Merge daily sales Excel files into one.")
    parser.add_argument(
        "-i", "--input", default="data/01-31012026", help="Input directory containing Excel files"
    )
    parser.add_argument(
        "-o", "--output", default="merged_sales_data.xlsx", help="Output Excel filename"
    )
    parser.add_argument(
        "-s", "--short", action="store_true", help="Output only date, เลขที่บิล, Cash, Credit Card, VISA, Tranfer"
    )
    args = parser.parse_args()

    merge_excel_files(input_dir=args.input, output_file=args.output, short=args.short)


if __name__ == "__main__":
    main()
