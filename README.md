# Apisomi Invoice Studio & Automated Ledger

An automated corporate invoicing system designed for **Apisomi**, supporting both **Non-GST** and **GST Tax Invoices**, live vector PDF export, and synchronized dual logging into **Excel (.xlsx)** and a permanent **Google Sheet**.

---

## Key Features

1. **Exact Apisomi Numbering & Template**:
   - Invoice Numbering: `APS-INV-[PROJECT]-[NN]` (e.g., `APS-INV-SSV003-01`).
   - Standard PDF filename: `Apisomi_[Project]_Invoice_[NN].pdf`.
   - 3-Column Parties comparison table (`Billed By` vs `Billed To`).
   - Project metadata strip with SAC codes (e.g. `998314`) and units.

2. **Smart Tax Switch (Non-GST & GST)**:
   - **Non-GST Mode**: Generates a clean Bill of Supply with zeroed tax rows and threshold disclosures.
   - **GST Mode**: Toggle on when GSTIN arrives.
     - Auto detects Place of Supply:
       - Client in **Madhya Pradesh (23)** &rarr; **CGST 9% + SGST 9%** (Intra-state).
       - Client outside MP &rarr; **IGST 18%** (Inter-state).
   - Indian Numbering to Words conversion (`Indian Rupees ... Only`).

3. **Installment Milestones & Txn Refs**:
   - Payment schedule tranches with scheduled due dates.
   - Reserved empty / populated transaction slots (`Txn Ref / UTR #`).
   - Real-time balance due calculations.

4. **Permanent Google Sheets Live Sync**:
   - Linked to [invoice_tracking](https://docs.google.com/spreadsheets/d/1fubQ-v4GZixmfQOOV2VnrJtoUDNjsSM-wdmwwXwkrQk/edit?usp=sharing).
   - Automatically appends a structured row on every export or when clicking **"Update Sheet"**.

5. **Client-Side Excel Tracker (.xlsx)**:
   - Maintains a local cumulative ledger with instant `.xlsx` export featuring automated totals.

---

## Deployment on Vercel

This application is ready to deploy directly on Vercel as a zero-config static site:

```bash
# Push to GitHub
git push -u origin main

# Import repository in Vercel dashboard:
# Framework Preset: Other
# Build Command: (leave empty)
# Output Directory: ./
```

---

## Google Apps Script Setup (for Google Sheet Sync)

1. Open your sheet &rarr; **Extensions &rarr; Apps Script**.
2. Paste the provided Webhook script into `Code.gs`.
3. Click **Deploy &rarr; New deployment** &rarr; Select type: **Web app**.
4. Set *Execute as:* **Me** and *Who has access:* **Anyone**.
5. Copy the Web app URL and paste it into the **Google Sheet** settings modal in the dashboard.
