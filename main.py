from datetime import datetime
from enum import Enum
import io
from typing import Any, Callable, Dict, List, Optional, Tuple
import urllib.parse
import zipfile

from fastapi import FastAPI, File, HTTPException, Query, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
from pydantic import BaseModel, Field

from utils.convert_pnd1 import (
    extract_pnd1_data,
    format_pnd1_export_name,
    parse_month_year as parse_pnd1_month_year,
    save_to_csv as save_pnd1_to_csv,
    save_to_excel as save_pnd1_to_excel,
)
from utils.convert_pnd3 import (
    extract_pnd3_data,
    format_pnd3_export_name,
    parse_month_year as parse_pnd3_month_year,
    save_to_csv as save_pnd3_to_csv,
    save_to_excel as save_pnd3_to_excel,
)
from utils.convert_pnd53 import (
    extract_pnd53_data,
    format_pnd53_export_name,
    parse_month_year as parse_pnd53_month_year,
    save_to_csv as save_pnd53_to_csv,
    save_to_excel as save_pnd53_to_excel,
)
from utils.convert_sso import (
    extract_sso_data,
    format_sso_export_name,
    parse_period,
    save_to_csv as save_sso_to_csv,
    save_to_excel as save_sso_to_excel,
)
from utils.convert_statement import (
    extract_statement_dataframe,
    format_statement_export_name,
    parse_statement_month_year,
    save_to_csv as save_statement_to_csv,
    save_to_excel as save_statement_to_excel,
)


class OutputFormat(str, Enum):
    XLSX = "xlsx"
    CSV = "csv"


class ErrorResponse(BaseModel):
    detail: str = Field(..., example="Invalid file type. Please upload a PDF file.")


tags_metadata = [
    {
        "name": "Social Security (SSO)",
        "description": "Extract employee contributions from SSO Form 1-10 Part 2.",
    },
    {
        "name": "Withholding Tax (P.N.D.)",
        "description": "Extract withholding tax attachments: P.N.D. 1 (payroll), P.N.D. 3 (individuals), P.N.D. 53 (companies).",
    },
    {
        "name": "Bank Statement (รายการเดินบัญชี)",
        "description": "Extract transaction records from bank statement PDFs (ประเภทบัญชีเดินสะพัด / Current Account).",
    },
    {
        "name": "System",
        "description": "Health check and operational endpoints.",
    },
]

app = FastAPI(
    title="Banchee Document Conversion API",
    description="""
## Overview
High-performance REST API service designed to extract and convert Thai accounting and tax PDF forms into structured data formats (**Excel .xlsx** and **CSV .csv**).

### Supported Documents
1. **SSO 1-10 Part 2 (สปส. 1-10 ส่วนที่ 2)**: Employee Social Security contribution records.
2. **P.N.D. 1 (ภ.ง.ด. 1 ใบแนบ)**: Employee withholding tax records on monthly salaries and wages.
3. **P.N.D. 3 (ภ.ง.ด. 3 ใบแนบ)**: Withholding tax records paid to natural persons (individuals).
4. **P.N.D. 53 (ภ.ง.ด. 53 ใบแนบ)**: Withholding tax records paid to corporate entities.
5. **Bank Statement (รายการเดินบัญชี ประเภทบัญชีเดินสะพัด)**: Current Account bank statement records (SCB, etc.).

### Multi-file & Packaging Behavior
* **Single file uploaded**: Returns the converted file directly (`{prefix}-MM-YYYY.xlsx` or `{prefix}-MM-YYYY.csv`).
* **Multiple files uploaded (> 1 file)**: Bundles all converted files into a **ZIP archive** named **`{prefix}-YYYY.zip`** (where `YYYY` is extracted from the document periods).
    """,
    version="1.0.0",
    openapi_tags=tags_metadata,
    contact={
        "name": "Banchee Engineering Support",
    },
)

# Enable CORS for frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get(
    "/",
    tags=["System"],
    summary="Service Health Check",
    description="Returns the status and health check of the Banchee API service.",
    response_model=Dict[str, str],
)
def root():
    return {"message": "Banchee API is running"}


async def _process_multi_or_single_files(
    files: List[UploadFile],
    format: OutputFormat,
    doc_prefix: str,
    extractor: Callable[[io.BytesIO], Tuple[pd.DataFrame, Any]],
    filename_generator: Callable[[pd.DataFrame, Any, str], str],
    year_extractor: Callable[[pd.DataFrame, Any], Tuple[Optional[str], Optional[str]]],
    save_excel_func: Callable[[pd.DataFrame, io.BytesIO], None],
    save_csv_func: Callable[[pd.DataFrame, io.BytesIO], None],
) -> StreamingResponse:
    if not files:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No files provided.",
        )

    export_ext = format.value
    converted_files: List[Tuple[str, bytes]] = []
    detected_years: List[str] = []

    for upload_file in files:
        if not upload_file.filename.lower().endswith(".pdf"):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid file type for '{upload_file.filename}'. Please upload PDF files only.",
            )

        content = await upload_file.read()
        if not content:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Uploaded file '{upload_file.filename}' is empty.",
            )

        try:
            df, extra = extractor(io.BytesIO(content))
        except Exception as e:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail=f"Failed to process PDF '{upload_file.filename}': {str(e)}",
            )

        if df.empty:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"No valid records found in PDF '{upload_file.filename}'.",
            )

        _, year = year_extractor(df, extra)
        if year:
            detected_years.append(year)

        export_filename = filename_generator(df, extra, export_ext)

        file_buffer = io.BytesIO()
        if export_ext == "xlsx":
            save_excel_func(df, file_buffer)
        else:
            save_csv_func(df, file_buffer)

        converted_files.append((export_filename, file_buffer.getvalue()))

    # Case 1: Single file upload
    if len(converted_files) == 1:
        export_filename, file_bytes = converted_files[0]
        media_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if export_ext == "xlsx"
            else "text/csv; charset=utf-8"
        )
        encoded_filename = urllib.parse.quote(export_filename)
        headers = {
            "Content-Disposition": f'attachment; filename="{export_filename}"; filename*=UTF-8\'\'{encoded_filename}',
            "Access-Control-Expose-Headers": "Content-Disposition",
        }
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type=media_type,
            headers=headers,
        )

    # Case 2: Multiple files uploaded (> 1) -> ZIP package named {doc_prefix}-YYYY.zip
    zip_year = detected_years[0] if detected_years else datetime.now().strftime("%Y")
    zip_filename = f"{doc_prefix}-{zip_year}.zip"

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        used_names: set[str] = set()
        for fname, fbytes in converted_files:
            final_name = fname
            idx = 1
            name_stem = fname.rsplit(".", 1)[0]
            name_ext = fname.rsplit(".", 1)[1]
            while final_name in used_names:
                final_name = f"{name_stem}_{idx}.{name_ext}"
                idx += 1
            used_names.add(final_name)
            zip_file.writestr(final_name, fbytes)

    zip_buffer.seek(0)
    encoded_zip_filename = urllib.parse.quote(zip_filename)
    headers = {
        "Content-Disposition": f'attachment; filename="{zip_filename}"; filename*=UTF-8\'\'{encoded_zip_filename}',
        "Access-Control-Expose-Headers": "Content-Disposition",
    }
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers=headers,
    )


# -----------------------------------------------------------------------------
# 1. SSO
# -----------------------------------------------------------------------------
@app.post(
    "/convert/sso",
    tags=["Social Security (SSO)"],
    summary="Convert SSO (สปส. 1-10 ส่วนที่ 2) PDF(s) to Excel / CSV / ZIP",
    description="""
Upload one or more **SSO 1-10 Part 2 (สปส. 1-10 ส่วนที่ 2)** PDF documents.

### Output Behavior:
- **Single file**: Returns `sso-MM-YYYY.xlsx` or `sso-MM-YYYY.csv`.
- **Multiple files (> 1)**: Returns a ZIP archive named **`sso-YYYY.zip`**.
    """,
    responses={
        status.HTTP_200_OK: {
            "description": "Successful conversion. Returns the converted file stream (.xlsx, .csv) or .zip archive.",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "text/csv": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/zip": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        status.HTTP_400_BAD_REQUEST: {"model": ErrorResponse},
        status.HTTP_422_UNPROCESSABLE_ENTITY: {"model": ErrorResponse},
    },
)
async def convert_sso(
    files: List[UploadFile] = File(
        ...,
        description="One or more SSO 1-10 Part 2 PDF document files.",
    ),
    format: OutputFormat = Query(
        OutputFormat.XLSX,
        description="Export format: `xlsx` or `csv` (Default: `xlsx`).",
    ),
):
    return await _process_multi_or_single_files(
        files=files,
        format=format,
        doc_prefix="sso",
        extractor=lambda stream: extract_sso_data(stream),
        filename_generator=lambda df, meta, ext: format_sso_export_name(meta, ext),
        year_extractor=lambda df, meta: parse_period(meta.get("period", "") if meta else ""),
        save_excel_func=save_sso_to_excel,
        save_csv_func=save_sso_to_csv,
    )


# -----------------------------------------------------------------------------
# 2. P.N.D. 1 (ภ.ง.ด. 1)
# -----------------------------------------------------------------------------
@app.post(
    "/convert/pnd1",
    tags=["Withholding Tax (P.N.D.)"],
    summary="Convert P.N.D. 1 (ภ.ง.ด. 1 ใบแนบ) PDF(s) to Excel / CSV / ZIP",
    description="""
Upload one or more **P.N.D. 1 Attachment (ภ.ง.ด. 1 ใบแนบ)** PDF documents.

### Output Behavior:
- **Single file**: Returns `pnd1-MM-YYYY.xlsx` or `pnd1-MM-YYYY.csv`.
- **Multiple files (> 1)**: Returns a ZIP archive named **`pnd1-YYYY.zip`**.

### What it extracts:
- **ลำดับ** (No.)
- **เลขประจำตัวผู้เสียภาษีอากร** (Tax ID, preserving leading zeros)
- **ชื่อผู้มีเงินได้** (Employee Name)
- **วัน เดือน ปี ที่จ่าย** (Payment Date)
- **จำนวนเงินได้ที่จ่ายในครั้งนี้** (Income Amount)
- **จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้** (Withholding Tax Amount)
- **เงื่อนไข** (Condition)
    """,
    responses={
        status.HTTP_200_OK: {
            "description": "Successful conversion. Returns the converted file stream (.xlsx, .csv) or .zip archive.",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "text/csv": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/zip": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        status.HTTP_400_BAD_REQUEST: {"model": ErrorResponse},
        status.HTTP_422_UNPROCESSABLE_ENTITY: {"model": ErrorResponse},
    },
)
async def convert_pnd1(
    files: List[UploadFile] = File(
        ...,
        description="One or more P.N.D. 1 attachment PDF document files.",
    ),
    format: OutputFormat = Query(
        OutputFormat.XLSX,
        description="Export format: `xlsx` or `csv` (Default: `xlsx`).",
    ),
):
    return await _process_multi_or_single_files(
        files=files,
        format=format,
        doc_prefix="pnd1",
        extractor=lambda stream: (extract_pnd1_data(stream), None),
        filename_generator=lambda df, extra, ext: format_pnd1_export_name(df, ext),
        year_extractor=lambda df, extra: parse_pnd1_month_year(df),
        save_excel_func=save_pnd1_to_excel,
        save_csv_func=save_pnd1_to_csv,
    )


# -----------------------------------------------------------------------------
# 3. P.N.D. 3 (ภ.ง.ด. 3)
# -----------------------------------------------------------------------------
@app.post(
    "/convert/pnd3",
    tags=["Withholding Tax (P.N.D.)"],
    summary="Convert P.N.D. 3 (ภ.ง.ด. 3 ใบแนบ) PDF(s) to Excel / CSV / ZIP",
    description="""
Upload one or more **P.N.D. 3 Attachment (ภ.ง.ด. 3 ใบแนบ)** PDF documents (Individuals).

### Output Behavior:
- **Single file**: Returns `pnd3-MM-YYYY.xlsx` or `pnd3-MM-YYYY.csv`.
- **Multiple files (> 1)**: Returns a ZIP archive named **`pnd3-YYYY.zip`**.

### What it extracts:
- **ลำดับ** (No.)
- **เลขประจำตัวผู้เสียภาษีอากร** (Tax ID, preserving leading zeros)
- **ชื่อผู้มีเงินได้** (Payee Name)
- **ที่อยู่** (Address)
- **วัน เดือน ปี ที่จ่าย** (Payment Date)
- **ประเภทเงินได้** (Income Type)
- **อัตราภาษีร้อยละ** (Tax Rate %)
- **จำนวนเงินที่จ่ายในครั้งนี้** (Payment Amount)
- **จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้** (Withholding Tax Amount)
- **เงื่อนไข** (Condition)
    """,
    responses={
        status.HTTP_200_OK: {
            "description": "Successful conversion. Returns the converted file stream (.xlsx, .csv) or .zip archive.",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "text/csv": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/zip": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        status.HTTP_400_BAD_REQUEST: {"model": ErrorResponse},
        status.HTTP_422_UNPROCESSABLE_ENTITY: {"model": ErrorResponse},
    },
)
async def convert_pnd3(
    files: List[UploadFile] = File(
        ...,
        description="One or more P.N.D. 3 attachment PDF document files.",
    ),
    format: OutputFormat = Query(
        OutputFormat.XLSX,
        description="Export format: `xlsx` or `csv` (Default: `xlsx`).",
    ),
):
    return await _process_multi_or_single_files(
        files=files,
        format=format,
        doc_prefix="pnd3",
        extractor=lambda stream: (extract_pnd3_data(stream), None),
        filename_generator=lambda df, extra, ext: format_pnd3_export_name(df, ext),
        year_extractor=lambda df, extra: parse_pnd3_month_year(df),
        save_excel_func=save_pnd3_to_excel,
        save_csv_func=save_pnd3_to_csv,
    )


# -----------------------------------------------------------------------------
# 4. P.N.D. 53 (ภ.ง.ด. 53)
# -----------------------------------------------------------------------------
@app.post(
    "/convert/pnd53",
    tags=["Withholding Tax (P.N.D.)"],
    summary="Convert P.N.D. 53 (ภ.ง.ด. 53 ใบแนบ) PDF(s) to Excel / CSV / ZIP",
    description="""
Upload one or more **P.N.D. 53 Attachment (ภ.ง.ด. 53 ใบแนบ)** PDF documents (Corporates).

### Output Behavior:
- **Single file**: Returns `pnd53-MM-YYYY.xlsx` or `pnd53-MM-YYYY.csv`.
- **Multiple files (> 1)**: Returns a ZIP archive named **`pnd53-YYYY.zip`**.

### What it extracts:
- **ลำดับ** (No.)
- **เลขประจำตัวผู้เสียภาษีอากร** (Tax ID, preserving leading zeros)
- **ชื่อผู้มีเงินได้** (Company / Payee Name)
- **สาขาที่** (Branch No., preserving leading zeros)
- **ที่อยู่** (Address)
- **วัน เดือน ปี ที่จ่าย** (Payment Date)
- **ประเภทเงินได้** (Income Type)
- **อัตราภาษีร้อยละ** (Tax Rate %)
- **จำนวนเงินที่จ่ายในครั้งนี้** (Payment Amount)
- **จำนวนเงินภาษีที่หัก และนำส่งในครั้งนี้** (Withholding Tax Amount)
- **เงื่อนไข** (Condition)
    """,
    responses={
        status.HTTP_200_OK: {
            "description": "Successful conversion. Returns the converted file stream (.xlsx, .csv) or .zip archive.",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "text/csv": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/zip": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        status.HTTP_400_BAD_REQUEST: {"model": ErrorResponse},
        status.HTTP_422_UNPROCESSABLE_ENTITY: {"model": ErrorResponse},
    },
)
async def convert_pnd53(
    files: List[UploadFile] = File(
        ...,
        description="One or more P.N.D. 53 attachment PDF document files.",
    ),
    format: OutputFormat = Query(
        OutputFormat.XLSX,
        description="Export format: `xlsx` or `csv` (Default: `xlsx`).",
    ),
):
    return await _process_multi_or_single_files(
        files=files,
        format=format,
        doc_prefix="pnd53",
        extractor=lambda stream: (extract_pnd53_data(stream), None),
        filename_generator=lambda df, extra, ext: format_pnd53_export_name(df, ext),
        year_extractor=lambda df, extra: parse_pnd53_month_year(df),
        save_excel_func=save_pnd53_to_excel,
        save_csv_func=save_pnd53_to_csv,
    )


# -----------------------------------------------------------------------------
# 5. Bank Statement (รายการเดินบัญชี ประเภทบัญชีเดินสะพัด)
# -----------------------------------------------------------------------------
@app.post(
    "/convert/statement",
    tags=["Bank Statement (รายการเดินบัญชี)"],
    summary="Convert Bank Statement (ประเภทบัญชี: เดินสะพัด) PDF(s) to Excel / CSV / ZIP",
    description="""
Upload one or more **Bank Statement (ประเภทบัญชี: เดินสะพัด / Current Account)** PDF documents.

### Output Behavior:
- **Single file**: Returns `statement-current-MM-YYYY.xlsx` or `statement-current-MM-YYYY.csv`.
- **Multiple files (> 1)**: Returns a ZIP archive named **`statement-current-YYYY.zip`**.

### What it extracts:
- **Date/Time** (วัน/เวลา)
- **Channel** (ช่องทาง)
- **Code** (รายการ)
- **Cheque No** (เลขที่เช็ค)
- **Withdrawal (Debit)** (จำนวนเงินที่หักบัญชี)
- **Deposit (Credit)** (จำนวนเงินนำเข้าบัญชี)
- **Balance** (ยอดเงินคงเหลือ)
- **Description** (รายละเอียด)
- **Note** (บันทึกช่วยจำ)
    """,
    responses={
        status.HTTP_200_OK: {
            "description": "Successful conversion. Returns the converted file stream (.xlsx, .csv) or .zip archive.",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "text/csv": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/zip": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        status.HTTP_400_BAD_REQUEST: {"model": ErrorResponse},
        status.HTTP_422_UNPROCESSABLE_ENTITY: {"model": ErrorResponse},
    },
)
async def convert_statement(
    files: List[UploadFile] = File(
        ...,
        description="One or more Bank Statement PDF documents (ประเภทบัญชีเดินสะพัด).",
    ),
    format: OutputFormat = Query(
        OutputFormat.XLSX,
        description="Export format: `xlsx` or `csv` (Default: `xlsx`).",
    ),
):
    return await _process_multi_or_single_files(
        files=files,
        format=format,
        doc_prefix="statement-current",
        extractor=lambda stream: extract_statement_dataframe(stream),
        filename_generator=lambda df, meta, ext: format_statement_export_name(meta, ext, df),
        year_extractor=lambda df, meta: parse_statement_month_year(df, meta),
        save_excel_func=save_statement_to_excel,
        save_csv_func=save_statement_to_csv,
    )

