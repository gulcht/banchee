# Banchee Document Conversion REST API

FastAPI-based REST service for extracting and converting Thai accounting and tax PDF documents (**SSO**, **P.N.D. 1**, **P.N.D. 3**, **P.N.D. 53**, and **Bank Statement**) into structured Excel (`.xlsx`), CSV (`.csv`), or ZIP (`.zip`) files.

---

## 🚀 Quick Start

### 1. Requirements & Dependencies

Ensure Python 3.9+ is installed, then install the dependencies:

```bash
pip install -r requirements.txt
```

### 2. Running the Server

Start the API with Uvicorn:

```bash
uvicorn main:app --reload --port 8000
```

- **Interactive Swagger UI**: [http://localhost:8000/docs](http://localhost:8000/docs)
- **ReDoc Documentation**: [http://localhost:8000/redoc](http://localhost:8000/redoc)

---

## 📌 Summary of Endpoints

All conversion endpoints share identical query parameters and packaging behaviors:

| Endpoint | Document Type | Description |
| :--- | :--- | :--- |
| `POST /convert/sso` | SSO 1-10 Part 2 (สปส. 1-10 ส่วนที่ 2) | Social Security contribution records |
| `POST /convert/pnd1` | P.N.D. 1 (ภ.ง.ด. 1 ใบแนบ) | Withholding tax on payroll / salary |
| `POST /convert/pnd3` | P.N.D. 3 (ภ.ง.ด. 3 ใบแนบ) | Withholding tax for natural persons (individuals) |
| `POST /convert/pnd53` | P.N.D. 53 (ภ.ง.ด. 53 ใบแนบ) | Withholding tax for corporate entities |
| `POST /convert/statement` | Bank Statement (ประเภทบัญชี: เดินสะพัด) | Current account bank statement records (SCB, etc.) |

---

## ⚙️ Parameters & Output Behavior

### Query Parameters

| Parameter | Type | Required | Default | Values | Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `format` | `string` | No | `xlsx` | `xlsx`, `csv` | Target output format. Defaults to `xlsx`. |

### Form Data

| Field | Type | Description |
| :--- | :--- | :--- |
| `files` | `binary` (Array of PDFs) | One or more PDF document files. |

---

### 📦 Naming Conventions & Packaging

- **Single File Upload (`len(files) == 1`)**:
  Returns the converted file stream directly:
  $$\mathbf{\{prefix\}\text{-}MM\text{-}YYYY.\{xlsx|csv\}}$$
  *Examples*: `sso-01-2026.xlsx`, `pnd1-01-2026.xlsx`, `pnd3-01-2026.csv`, `pnd53-01-2026.xlsx`

- **Multiple Files Upload (`len(files) > 1`)**:
  Packages all converted files into a **ZIP archive**:
  $$\mathbf{\{prefix\}\text{-}YYYY.zip}$$
  *Examples*: `sso-2026.zip`, `pnd1-2026.zip`, `pnd3-2026.zip`, `pnd53-2026.zip`

*(Month and Year are dynamically extracted from the payment dates or declaration period in the PDF).*

---

## 💻 Usage Examples (cURL)

### 1. P.N.D. 1 (ภ.ง.ด. 1)

```bash
# Single file -> pnd1-MM-YYYY.xlsx
curl -X POST "http://localhost:8000/convert/pnd1" \
  -F "files=@pnd1_jan.pdf" \
  -OJ

# Multiple files -> pnd1-YYYY.zip
curl -X POST "http://localhost:8000/convert/pnd1?format=xlsx" \
  -F "files=@pnd1_jan.pdf" \
  -F "files=@pnd1_feb.pdf" \
  -OJ
```

### 2. P.N.D. 3 (ภ.ง.ด. 3)

```bash
# Multiple files as CSV -> pnd3-YYYY.zip
curl -X POST "http://localhost:8000/convert/pnd3?format=csv" \
  -F "files=@pnd3_jan.pdf" \
  -F "files=@pnd3_feb.pdf" \
  -OJ
```

### 3. P.N.D. 53 (ภ.ง.ด. 53)

```bash
# Single file as CSV -> pnd53-MM-YYYY.csv
curl -X POST "http://localhost:8000/convert/pnd53?format=csv" \
  -F "files=@pnd53_jan.pdf" \
  -OJ

# Multiple files as Excel -> pnd53-YYYY.zip
curl -X POST "http://localhost:8000/convert/pnd53?format=xlsx" \
  -F "files=@pnd53_jan.pdf" \
  -F "files=@pnd53_feb.pdf" \
  -OJ
```

### 4. SSO 1-10 (สปส. 1-10 ส่วนที่ 2)

```bash
# Multiple files -> sso-YYYY.zip
curl -X POST "http://localhost:8000/convert/sso" \
  -F "files=@sso_jan.pdf" \
  -F "files=@sso_feb.pdf" \
  -OJ
```

### 5. Bank Statement (ประเภทบัญชี: เดินสะพัด)

```bash
# Single file -> statement-current-MM-YYYY.xlsx
curl -X POST "http://localhost:8000/convert/statement" \
  -F "files=@SCB.PDF" \
  -OJ

# Multiple files as Excel -> statement-current-YYYY.zip
curl -X POST "http://localhost:8000/convert/statement?format=xlsx" \
  -F "files=@scb_jan.pdf" \
  -F "files=@scb_feb.pdf" \
  -OJ
```

---

## 🌐 Swagger UI

Run the server and visit:
- **Interactive Swagger Docs**: [http://localhost:8000/docs](http://localhost:8000/docs)
- **ReDoc UI**: [http://localhost:8000/redoc](http://localhost:8000/redoc)
