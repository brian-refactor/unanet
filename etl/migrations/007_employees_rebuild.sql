-- Rebuild employees table with full schema from Unanet HR load file.
-- Table was empty so we drop and recreate cleanly.

DROP TABLE IF EXISTS employees CASCADE;

CREATE TABLE employees (
    id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    office           TEXT NOT NULL,
    record_key       TEXT NOT NULL,          -- employee_code, e.g. UOC000068
    employee_code    TEXT NOT NULL UNIQUE,   -- same value, explicit name
    first_name       TEXT NOT NULL,
    last_name        TEXT NOT NULL,
    middle_name      TEXT,
    full_name        TEXT NOT NULL,          -- computed: first [mid] last
    email            TEXT,
    org_path         TEXT NOT NULL,          -- e.g. FUS-MSP-A01
    hire_date        DATE,
    termination_date DATE,
    is_active        BOOLEAN NOT NULL DEFAULT TRUE,
    timesheet_group  TEXT,                   -- Weekly / Inactive
    pay_group        TEXT,                   -- Salary / Hourly
    job_type         TEXT,                   -- Principal-In-Charge / Project Manager / Project Accountant
    job_title_code   TEXT,
    job_title_name   TEXT,
    is_subcontractor BOOLEAN NOT NULL DEFAULT FALSE,
    pay_rate         NUMERIC,
    billing_rate     NUMERIC,
    target_pct       NUMERIC,
    created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX employees_office_idx ON employees (office);
CREATE INDEX employees_org_path_idx ON employees (org_path);
