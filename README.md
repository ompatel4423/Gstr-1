# GSTR-1 Automation & Portal JSON Compiler

A streamlined Google Sheets and Google Apps Script solution to automate outward supply segregation and generate GST portal-ready JSON files for GSTR-1 filing.

## Key Features

- **Automated Table Segregation:**
  - Table 4: B2B (Registered outward supplies)
  - Table 5: B2C Large (Inter-state unregistered supplies > ₹1 Lakh)
  - Table 7: B2C Small (Intra-state & remaining inter-state unregistered supplies)
  - Table 9B: CDNR (Credit/Debit Notes to registered persons)
  - Table 12: HSN Summary (Automatic calculation of taxable amounts, tax splits, and gross values)
  - Table 13: Document Summary (Invoice serial ranges, cancellations, and net counts)
- **Portal Compliance:**
  - Compliant with direct web upload standards (`gst.gov.in`).
  - Includes mandatory `val` (Gross Value) field for HSN summary lines.
  - Supports both Bifurcated (`hsn_b2b` / `hsn_b2c`) and Unified (`data`) HSN structures.
  - Validates GSTIN format and Modulus-36 checksums.
- **Built-in Reconciliation:**
  - Cross-checks raw invoice totals against segregated tables.
  - Reconciles outward gross sales against Table 12 HSN totals.
- **Client Reset Utility:**
  - Resets invoice tables while preserving headers and configuration for the next tax period.

## Directory Structure

```text
├── src/
│   ├── Code.js              # Google Apps Script automation logic
│   └── appsscript.json      # Apps Script project manifest
├── sample_data/
│   ├── Invoices_Raw_Template.csv   # Template format for raw ledger entries
│   └── sample_gstr1_output.json    # Example portal-ready JSON payload
├── .gitignore
├── LICENSE
└── README.md
```

## Quick Start Guide

### 1. Spreadsheet Setup
Create a Google Spreadsheet with the following sheet tabs:
- `Config`: Store supplier GSTIN, tax period (`MMYYYY`), state code, and B2C threshold.
- `Invoices_Raw`: Master invoice register.
- `B2B_Summary`, `B2C_Large`, `B2C_Small`, `HSN_Summary`, `CDNR_Summary`, `Docs_Summary`.

### 2. Install Apps Script
1. In Google Sheets, navigate to **Extensions > Apps Script**.
2. Replace any boilerplate code with `src/Code.js`.
3. Save the project and refresh your spreadsheet.

### 3. Usage Workflow
1. **Enter Data:** Populate `Invoices_Raw` with invoice details.
2. **Segregate:** Click **⚡ GSTR-1 Automation > 🚀 1. Process & Segregate All Tables**.
3. **Audit:** Click **⚡ GSTR-1 Automation > 🔍 2. Run Audit & Cross-Reconciliation**.
4. **Export:** Click **⚡ GSTR-1 Automation > 📦 3. Export GST Portal Upload Ready JSON** and download the `.json` file.
5. **Upload:** Go to **gst.gov.in > Returns Dashboard > GSTR-1 > Prepare Offline > Upload**.

## License
MIT License. See [LICENSE](LICENSE) for details.
