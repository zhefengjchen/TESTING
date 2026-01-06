# Invoice Management System

A local Windows desktop application for managing Excel-based invoice data from multiple testing labs (BV, ITS, TUV, SGS).

## Features

- **Excel Invoice Upload**: Parse `TESTING+SERVICES` tab from Excel files
- **Lab Name Detection**: Auto-detect lab names from filenames
- **Data Validation**: Automatic validation with abnormal items tracking
- **Full CRUD Operations**: View, add, edit, and delete invoice records
- **Filter & Search**: Filter by lab name, date range, or search text
- **Settings Management**: Manage lab names and test/service types
- **Data Export**: Export data to Excel format

## Requirements

- Python 3.8 or higher
- Windows OS (recommended for full GUI support)

## Installation

1. Clone or download this repository:
   ```bash
   git clone <repository-url>
   cd invoice-management-system
   ```

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   # or
   source venv/bin/activate  # Linux/Mac
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the application:
```bash
python main.py
```

### Uploading Invoices

1. Click **Upload Invoice** button
2. Select an Excel file (`.xlsx`) containing invoice data
3. Enter the invoice date when prompted (optional)
4. The system will:
   - Parse the `TESTING+SERVICES` tab
   - Auto-detect lab name from filename
   - Validate all records
   - Import valid records to the main table
   - Move invalid records to "Abnormal Items"

### Excel File Requirements

Your Excel files should have a sheet named `TESTING+SERVICES` with the following columns:

| Column Name | Description |
|------------|-------------|
| Tajan Bidding Tracking Number | Tracking identifier |
| Amazon Test Request # | Test request number |
| Amazon Tracker Number | Tracker number |
| ASIN# | Amazon product ID |
| Development Center (SEA, LUX/EU, JP) | Development center code |
| Category (New HCC) | Product category |
| Product Brand | Brand name |
| Product Description | Product description |
| Testing SLA (Working days) | SLA in days |
| Test / Service Type | Type of test/service |
| Quotation/Order Number | Order number |
| Request Date | Request date |
| Test Start Date | Start date |
| Report Delivered Date | Delivery date |
| Report number | Report ID |
| Test / inspection Location | Location |
| Product line | Product line |
| AMAZON QUALITY MANAGER | QA manager |
| AMAZON SOURCING MANAGER | Sourcing manager |
| Invoice # | Invoice number |
| Lab contact | Lab contact info |
| COMMENT (IF ANY) | Comments |
| AMOUNT (Currency=USD) | Amount in USD |

### Lab Name Detection

The system detects lab names from filenames. Include one of these in your filename:
- BV
- ITS
- TUV
- SGS

Example: `BV_Invoice_2024Q1.xlsx` will be detected as lab "BV"

### Validation Rules

Records are moved to "Abnormal Items" if they fail these checks:

1. **Missing Tracking Number**: `Tajan Bidding Tracking Number` is empty
2. **Invalid Test Type**: `Test / Service Type` not in allowed list

### Allowed Test/Service Types (Default)

- Benchmark Test
- Benchmark Sample Purchase
- Usability Test
- Usability Protocol
- Comparison
- DDC Design Document Collection
- Delorean DDC
- DP Review
- PRD & Protocol Upgrade
- Performance Protocol
- PVT, PT, PPT, PT-Initial, PT-Final
- ET, VT
- Delorean PT
- Packaging
- Failure & Defect Analysis
- Declaration of Conformity
- ECR

You can add/remove test types in the Settings panel.

## Project Structure

```
invoice-management-system/
├── main.py              # Application entry point
├── requirements.txt     # Python dependencies
├── README.md           # This file
├── src/
│   ├── __init__.py     # Package init
│   ├── database.py     # SQLite database module
│   ├── excel_parser.py # Excel parsing & validation
│   └── gui.py          # CustomTkinter GUI
├── data/
│   └── invoices.db     # SQLite database (created on first run)
└── tests/
    └── (test files)
```

## Data Storage

All data is stored in a local SQLite database (`data/invoices.db`). The database is created automatically on first run.

## Exporting Data

Click the **Export** button to export all data to an Excel file with two sheets:
- Valid Invoices
- Abnormal Items

## License

This project is for internal use.
