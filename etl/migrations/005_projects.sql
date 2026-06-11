-- Projects and project phases tables for Unanet migration
-- office + project_code is the natural unique key across all four offices

CREATE TABLE IF NOT EXISTS projects (
    id              BIGSERIAL PRIMARY KEY,
    office          TEXT        NOT NULL,
    source_id       TEXT,                   -- QBO customer ID or QB sub-job ID
    project_code    TEXT        NOT NULL,
    project_name    TEXT,
    client_firm_code TEXT,                  -- FK to clients.firm_code (nullable until verified)
    owning_org      TEXT,                   -- populated after org-unit lookup file arrives
    charge_type     TEXT        NOT NULL DEFAULT 'Billable',
    start_date      TEXT,                   -- mm/dd/yyyy or blank
    end_date        TEXT,
    contract_type   TEXT,
    project_note    TEXT,
    po_number       TEXT,
    pm_emp_code     TEXT,                   -- populated after employee code file arrives
    pic_emp_code    TEXT,
    pa_emp_code     TEXT,                   -- populated after employee code file arrives
    billing_term_type TEXT,
    net_days        INTEGER,
    invoice_email   TEXT,
    location_street1 TEXT,
    location_street2 TEXT,
    location_city   TEXT,
    location_state  TEXT,
    location_zip    TEXT,
    location_country TEXT,
    use_client_bill_to BOOLEAN  NOT NULL DEFAULT TRUE,
    bill_to_street1 TEXT,
    bill_to_street2 TEXT,
    bill_to_city    TEXT,
    bill_to_state   TEXT,
    bill_to_zip     TEXT,
    bill_to_country TEXT,
    is_active       BOOLEAN     NOT NULL DEFAULT TRUE,
    UNIQUE (office, project_code)
);

-- WBS phases and tasks (Phases & Tasks tab in the Unanet template)
-- Each row = one L2+L3 combination; L3 columns NULL when L2 is the leaf node
CREATE TABLE IF NOT EXISTS project_phases (
    id              BIGSERIAL PRIMARY KEY,
    office          TEXT        NOT NULL,
    project_code    TEXT        NOT NULL,   -- FK to projects(office, project_code)
    contract_type   TEXT,
    level2_name     TEXT        NOT NULL,
    level2_code     TEXT        NOT NULL,
    level3_name     TEXT,
    level3_code     TEXT,
    start_date      TEXT,
    end_date        TEXT,
    org_path        TEXT,
    fixed_fee       NUMERIC,
    labor_contract_cap  NUMERIC,
    odc_contract_cap    NUMERIC,
    occ_contract_cap    NUMERIC,
    icc_fixed_fee       NUMERIC,
    labor_budget        NUMERIC,
    odc_budget          NUMERIC,
    occ_budget          NUMERIC,
    icc_budget          NUMERIC,
    hours_budget        NUMERIC,
    UNIQUE (office, project_code, level2_code, level3_code)
);
