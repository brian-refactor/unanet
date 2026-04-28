# QuickBooks Desktop — Data Export Instructions
### Fusion AE — Unanet Migration

**To:** Dallas & Orlando QuickBooks Administrators  
**Purpose:** We are migrating financial and project data into Unanet. Please run the exports below and return all files. No data will be changed in QuickBooks — these are read-only exports.

---

## Before You Begin

- You will need **QuickBooks Desktop open** and logged in to your company file
- Each export will produce an Excel file (`.xlsx`)
- **Name each file exactly as shown** so we can process them correctly
- When exporting to Excel, always choose **"Create New Worksheet"** unless noted otherwise
- Set all date ranges to **"All"** unless a specific range is noted

---

## Export Checklist

- [ ] 1. Chart of Accounts
- [ ] 2. Customer List
- [ ] 3. Vendor List
- [ ] 4. Employee List
- [ ] 5. Employee Pay Rates
- [ ] 6. Items / Services List

---

## Export 1 — Chart of Accounts

**File name:** `[OFFICE]_COA.xlsx` (e.g., `Dallas_COA.xlsx`)

**Steps:**
1. From the top menu, go to **Reports → Accountant & Taxes → Account Listing**
2. Click **Customize Report** (top left of the report window)
3. Under the **Display** tab, make sure these columns are checked:
   - Account
   - Type
   - Description
   - Balance Total
   - Include inactive (check this box if present)
4. Click **OK**
5. Click the **Excel** button (top of report) → **Create New Worksheet**
6. Save as `[OFFICE]_COA.xlsx`

---

## Export 2 — Customer List

**File name:** `[OFFICE]_Customers.xlsx`

**Steps:**
1. From the top menu, go to **Reports → Customers & Receivables → Customer Contact List**
2. Click **Customize Report**
3. Under **Display**, make sure these columns are checked:
   - Customer
   - Company Name
   - Bill to Address (Street 1, Street 2, City, State, Zip)
   - Main Phone
   - Main Email
   - Terms
   - Type
   - Notes (if available)
4. Set the date range to **All**
5. Click **OK**
6. Click **Excel → Create New Worksheet**
7. Save as `[OFFICE]_Customers.xlsx`

> **Also do this:** Go to **Customers → Customer Center**. Click **Excel** (top toolbar) → **Export Customer List**. Save as `[OFFICE]_Customers_Full.xlsx`. This captures any fields the report may have missed.

---

## Export 3 — Vendor List

**File name:** `[OFFICE]_Vendors.xlsx`

**Steps:**
1. From the top menu, go to **Reports → Vendors & Payables → Vendor Contact List**
2. Click **Customize Report**
3. Under **Display**, make sure these columns are checked:
   - Vendor
   - Company Name
   - Bill from Address (Street 1, Street 2, City, State, Zip)
   - Main Phone
   - Main Email
   - Terms
   - Type
   - 1099 (check this — important for Unanet)
   - Account No.
   - Notes (if available)
4. Set the date range to **All**
5. Click **OK**
6. Click **Excel → Create New Worksheet**
7. Save as `[OFFICE]_Vendors.xlsx`

> **Also do this:** Go to **Vendors → Vendor Center**. Click **Excel** (top toolbar) → **Export Vendor List**. Save as `[OFFICE]_Vendors_Full.xlsx`.

---

## Export 4 — Employee List

**File name:** `[OFFICE]_Employees.xlsx`

**Steps:**
1. From the top menu, go to **Reports → Employees & Payroll → Employee Contact List**
2. Click **Customize Report**
3. Under **Display**, make sure these columns are checked:
   - Employee
   - SS No. (Social Security — this helps match records; handle securely)
   - Address
   - Main Phone
   - Mobile
   - Main Email
   - Hire Date
   - Release Date (if applicable)
   - Type (Regular, Officer, Statutory)
4. Check **Include Inactive Employees** if that option is available
5. Click **OK**
6. Click **Excel → Create New Worksheet**
7. Save as `[OFFICE]_Employees.xlsx`

---

## Export 5 — Employee Pay Rates

**File name:** `[OFFICE]_PayRates.xlsx`

This export requires two separate reports.

### Part A — Earnings Summary
1. Go to **Reports → Employees & Payroll → Employee Earnings Summary**
2. Set date range: **From:** `01/01/2020` **To:** today's date (captures all pay history)
3. Click **Customize Report → Display** and make sure these are checked:
   - Employee
   - Regular Pay
   - Overtime Pay
   - Salary (if applicable)
4. Click **OK**
5. Click **Excel → Create New Worksheet**
6. Save as `[OFFICE]_PayRates_Summary.xlsx`

### Part B — Payroll Item Detail (for actual rates)
1. Go to **Reports → Employees & Payroll → Payroll Item Detail**
2. Set date range: **All**
3. Click **Excel → Create New Worksheet**
4. Save as `[OFFICE]_PayRates_Detail.xlsx`

> **Note:** If prompted about payroll subscription or rate visibility, just export what is visible. Do not skip this step — we need whatever pay rate history is accessible.

---

## Export 6 — Items / Services / Expense Codes

**File name:** `[OFFICE]_Items.xlsx`

**Steps:**
1. From the top menu, go to **Reports → Lists → Item Price List**
2. Click **Customize Report**
3. Under **Display**, make sure these columns are checked:
   - Item
   - Description
   - Type
   - Account
   - Price
   - Cost
   - Is Inactive
4. Click **OK**
5. Click **Excel → Create New Worksheet**
6. Save as `[OFFICE]_Items.xlsx`

> **Also do this:** Go to **Lists → Item List**. Click **Excel** (or right-click in the list → **Export**). Save as `[OFFICE]_Items_Full.xlsx`. This may capture sub-items and additional detail.

---

## Sending the Files

Please send the completed files to **[INSERT CONTACT / SHARED FOLDER PATH]**.

Expected files per office (12 total):
```
[OFFICE]_COA.xlsx
[OFFICE]_Customers.xlsx
[OFFICE]_Customers_Full.xlsx
[OFFICE]_Vendors.xlsx
[OFFICE]_Vendors_Full.xlsx
[OFFICE]_Employees.xlsx
[OFFICE]_PayRates_Summary.xlsx
[OFFICE]_PayRates_Detail.xlsx
[OFFICE]_Items.xlsx
[OFFICE]_Items_Full.xlsx
```

---

## Questions?

If a menu option is missing or the export produces an error, take a screenshot and send it along. Do not skip an export — even a partial file is useful.

**QuickBooks version note:** These instructions apply to QuickBooks Desktop Pro, Premier, and Enterprise 2019 or newer. If your version looks different, the same reports exist but may be in slightly different menu locations.
