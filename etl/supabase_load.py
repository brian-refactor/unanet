"""
Supabase loader — reads all output CSVs and upserts into Supabase.

Strategy per table:
  - coa, clients, vendors, expense_codes: upsert on natural unique key
  - client_contacts, vendor_contacts, employees: delete-by-office then re-insert

Usage:
    python etl/supabase_load.py                    # all offices, all entities
    python etl/supabase_load.py --office dallas    # one office
    python etl/supabase_load.py --entity clients   # one entity, all offices

Credentials: SUPABASE_URL and SUPABASE_KEY in etl/.env
"""

import argparse
import csv
import os
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client, Client

load_dotenv(Path(__file__).parent / ".env")

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]

OFFICES = ["minnesota", "cincinnati", "dallas", "orlando"]
ENTITIES = ["coa", "clients", "client_contacts", "vendors", "vendor_contacts", "employees", "expense_codes", "projects"]

OUTPUT_DIR = Path(__file__).parent.parent / "output"

# ---------------------------------------------------------------------------
# Column mappings: CSV column name → DB column name (None = skip)
# ---------------------------------------------------------------------------

COA_MAP = {
    "BaseCode": "base_code",
    "BaseName": "base_name",
    "Description": "description",
    "IsActive": "is_active",
    "Is1099": "is_1099",
    "IsSubcontractor": "is_subcontractor",
    "FinancialType": "financial_type",
    "SubledgerType": "subledger_type",
    "MetricType": "metric_type",
    "CostType": "cost_type",
    "PMType": "pm_type",
    "LaborRevenueType": "labor_revenue_type",
    "ExpenseRevenueType": "expense_revenue_type",
}

CLIENTS_MAP = {
    "FirmCode": "firm_code",
    "_source_id": "source_id",
    "FirmName": "firm_name",
    "IsActive": "is_active",
    "Website": "website",
    "ClientType": "client_type",
    "Specialty": "specialty",
    "Note": "note",
    "PayDays": "pay_days",
    "MainEmail": "main_email",
    "BillToAddress_Phone": "bill_to_phone",
    "BillToAddress_Street1": "bill_to_street1",
    "BillToAddress_Street2": "bill_to_street2",
    "BillToAddress_Street3": "bill_to_street3",
    "BillToAddress_Street4": "bill_to_street4",
    "BillToAddress_City": "bill_to_city",
    "BillToAddress_State": "bill_to_state",
    "BillToAddress_Zip": "bill_to_zip",
    "BillToAddress_Country": "bill_to_country",
    "MainContact_Prefix": "main_contact_prefix",
    "MainContact_Suffix": "main_contact_suffix",
    "MainContact_Title": "main_contact_title",
    "MainContact_FirstName": "main_contact_first_name",
    "MainContact_LastName": "main_contact_last_name",
    "MainContact_WorkPhone": "main_contact_work_phone",
    "MainContact_CellPhone": "main_contact_cell_phone",
    "MainContact_WorkEmail": "main_contact_work_email",
    "MainContact_HomeEmail": "main_contact_home_email",
}

CONTACTS_MAP = {
    "FirmCode": "firm_code",
    "FirmRelationship": "firm_relationship",
    "Prefix": "prefix",
    "Suffix": "suffix",
    "Title": "title",
    "FirstName": "first_name",
    "LastName": "last_name",
    "WorkPhone": "work_phone",
    "CellPhone": "cell_phone",
    "WorkEmail": "work_email",
    "HomeEmail": "home_email",
    "WorkAddress1": "work_address1",
    "WorkAddress2": "work_address2",
    "WorkAddress3": "work_address3",
    "WorkAddress4": "work_address4",
    "WorkCity": "work_city",
    "WorkState": "work_state",
    "WorkZip": "work_zip",
    "WorkCountry": "work_country",
    "HomeAddress1": "home_address1",
    "HomeAddress2": "home_address2",
    "HomeAddress3": "home_address3",
    "HomeAddress4": "home_address4",
    "HomeCity": "home_city",
    "HomeState": "home_state",
    "HomeZip": "home_zip",
    "HomeCountry": "home_country",
}

VENDORS_MAP = {
    "FirmCode": "firm_code",
    "_source_id": "source_id",
    "FirmName": "firm_name",
    "IsActive": "is_active",
    "IsConsultant": "is_consultant",
    "ConsultantType": "consultant_type",
    "Is1099": "is_1099",
    "Website": "website",
    "VendorType": "vendor_type",
    "Note": "note",
    "NetDays": "net_days",
    "EIN": "ein",
    "PayToAddress_Phone": "pay_to_phone",
    "PayToAddress_Street1": "pay_to_street1",
    "PayToAddress_Street2": "pay_to_street2",
    "PayToAddress_Street3": "pay_to_street3",
    "PayToAddress_Street4": "pay_to_street4",
    "PayToAddress_City": "pay_to_city",
    "PayToAddress_State": "pay_to_state",
    "PayToAddress_Zip": "pay_to_zip",
    "PayToAddress_Country": "pay_to_country",
    "MainContact_Prefix": "main_contact_prefix",
    "MainContact_Suffix": "main_contact_suffix",
    "MainContact_Title": "main_contact_title",
    "MainContact_FirstName": "main_contact_first_name",
    "MainContact_LastName": "main_contact_last_name",
    "MainContact_WorkPhone": "main_contact_work_phone",
    "MainContact_CellPhone": "main_contact_cell_phone",
    "MainContact_WorkEmail": "main_contact_work_email",
    "MainContact_HomeEmail": "main_contact_home_email",
    "EnableEFT": "enable_eft",
    "CompanyID": "company_id",
    "CompanyName": "company_name",
    "ABA/Routing": "aba_routing",
    "Account#": "account_number",
    "Savings": "savings",
    "EFType(SEC)": "ef_type_sec",
}

EMPLOYEES_MAP = {
    "EmployeeCode": "employee_code",
    "_source_id": "source_id",
    "EmployeeName": "employee_name",
    "PayRate": "pay_rate",
    "salaryperpayperiod": "salary_per_pay_period",
    "PayRateStartDate": "pay_rate_start_date",
    "PayRateEndDate": "pay_rate_end_date",
    "IsHourly": "is_hourly",
    "OTRate": "ot_rate",
    "OTMU": "otmu",
}

EXPENSE_CODES_MAP = {
    "ECCode": "ec_code",
    "ECName": "ec_name",
    "ShowInES": "show_in_es",
    "IsUnit": "is_unit",
    "UnitTypename": "unit_type_name",
    "ECTypename": "ec_type_name",
    "ExpMarkupTypename": "exp_markup_type_name",
    "Markup": "markup",
    "UnitRate": "unit_rate",
    "BillStatusname": "bill_status_name",
    "DirectBaseCode": "direct_base_code",
    "DirectBasename": "direct_base_name",
    "OHBaseCode": "oh_base_code",
    "OHBasename": "oh_base_name",
    "BilledDirectBaseCode": "billed_direct_base_code",
    "BilledDirectBasename": "billed_direct_base_name",
    "BilledMarkupBaseCode": "billed_markup_base_code",
    "BilledMarkupBasename": "billed_markup_base_name",
    "UnBilledBaseCode": "unbilled_base_code",
    "UnBilledBasename": "unbilled_base_name",
    "CurrencyCode": "currency_code",
    "PMCmtRequired": "pm_cmt_required",
    "IntCmtRequired": "int_cmt_required",
    "IsNonReim": "is_non_reim",
    "_IsActive": "is_active",
}

PROJECTS_MAP = {
    "office":            "office",
    "project_code":      "project_code",
    "project_name":      "project_name",
    "client_firm_code":  "client_firm_code",
    "owning_org":        "owning_org",
    "charge_type":       "charge_type",
    "start_date":        "start_date",
    "end_date":          "end_date",
    "contract_type":     "contract_type",
    "project_note":      "project_note",
    "po_number":         "po_number",
    "pm_emp_code":       "pm_emp_code",
    "pic_emp_code":      "pic_emp_code",
    "pa_emp_code":       "pa_emp_code",
    "billing_term_type": "billing_term_type",
    "net_days":          "net_days",
    "invoice_email":     "invoice_email",
    "use_client_bill_to":"use_client_bill_to",
    "is_active":         "is_active",
    "_source_id":        "source_id",
}

ENTITY_CONFIG = {
    "coa":             ("coa",             COA_MAP,           "upsert", ["office", "base_code"]),
    "clients":         ("clients",         CLIENTS_MAP,       "upsert", ["firm_code"]),
    "client_contacts": ("client_contacts", CONTACTS_MAP,      "replace", None),
    "vendors":         ("vendors",         VENDORS_MAP,       "upsert", ["firm_code"]),
    "vendor_contacts": ("vendor_contacts", CONTACTS_MAP,      "replace", None),
    "employees":       ("employees",       EMPLOYEES_MAP,     "replace", None),
    "expense_codes":   ("expense_codes",   EXPENSE_CODES_MAP, "upsert", ["ec_code"]),
    "projects":        ("projects",        PROJECTS_MAP,      "upsert", ["office", "project_code"]),
}

CSV_NAMES = {
    "coa":             "COA",
    "clients":         "Clients",
    "client_contacts": "ClientContacts",
    "vendors":         "Vendors",
    "vendor_contacts": "VendorContacts",
    "employees":       "Employees",
    "expense_codes":   "ExpenseCodes",
    "projects":        "Projects",
}

BOOLEAN_COLS = {
    "is_active", "is_1099", "is_subcontractor", "is_consultant", "is_hourly",
    "show_in_es", "is_unit", "pm_cmt_required", "int_cmt_required", "is_non_reim",
    "enable_eft", "savings", "use_client_bill_to",
}
INTEGER_COLS = {"pay_days", "net_days"}
NUMERIC_COLS = {
    "pay_rate", "salary_per_pay_period", "ot_rate", "otmu", "markup", "unit_rate",
}
DATE_COLS = {"pay_rate_start_date", "pay_rate_end_date"}

DUPLICATE_CSV_COLS = {"MainContact_HomeEmail"}  # Vendors CSV has this column twice


def coerce(col: str, val: str):
    if val == "" or val is None:
        return None
    if col in BOOLEAN_COLS:
        return val.upper() == "TRUE"
    if col in INTEGER_COLS:
        try:
            return int(float(val))
        except ValueError:
            return None
    if col in NUMERIC_COLS:
        try:
            return float(val)
        except ValueError:
            return None
    if col in DATE_COLS:
        return val if val else None
    return val


def load_csv(path: Path, col_map: dict, office: str) -> list[dict]:
    rows = []
    seen_csv_cols: set[str] = set()
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            record = {"office": office}
            seen_csv_cols.clear()
            for csv_col, db_col in col_map.items():
                if csv_col in DUPLICATE_CSV_COLS and csv_col in seen_csv_cols:
                    continue
                seen_csv_cols.add(csv_col)
                val = raw.get(csv_col, "")
                record[db_col] = coerce(db_col, val)
            rows.append(record)
    return rows


def chunk(lst: list, size: int):
    for i in range(0, len(lst), size):
        yield lst[i : i + size]


def load_entity(sb: Client, entity: str, office: str) -> None:
    table_name, col_map, strategy, conflict_cols = ENTITY_CONFIG[entity]
    csv_file = OUTPUT_DIR / office / f"{office}_{CSV_NAMES[entity]}.csv"

    if not csv_file.exists():
        print(f"  SKIP  {csv_file.name} (not found)")
        return

    rows = load_csv(csv_file, col_map, office)
    print(f"  {len(rows):>5} rows  <- {csv_file.name}")

    if not rows:
        return

    if strategy == "replace":
        sb.table(table_name).delete().eq("office", office).execute()

    # Insert/upsert in batches of 500
    inserted = 0
    for batch in chunk(rows, 500):
        if strategy == "upsert":
            sb.table(table_name).upsert(batch, on_conflict=",".join(conflict_cols)).execute()
        else:
            sb.table(table_name).insert(batch).execute()
        inserted += len(batch)

    print(f"          -> {inserted} rows loaded into {table_name}")


def main():
    parser = argparse.ArgumentParser(description="Load ETL output CSVs into Supabase")
    parser.add_argument("--office", choices=OFFICES, help="Load one office only")
    parser.add_argument("--entity", choices=ENTITIES, help="Load one entity type only")
    args = parser.parse_args()

    sb: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

    offices = [args.office] if args.office else OFFICES
    entities = [args.entity] if args.entity else ENTITIES

    for office in offices:
        print(f"\n=== {office.upper()} ===")
        for entity in entities:
            load_entity(sb, entity, office)

    print("\nDone.")


if __name__ == "__main__":
    main()
