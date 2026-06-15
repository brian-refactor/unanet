CREATE TABLE IF NOT EXISTS org_units (
    id            BIGSERIAL PRIMARY KEY,
    org_code      TEXT NOT NULL,
    org_name      TEXT NOT NULL,
    parent_org_code TEXT,
    is_active     BOOLEAN NOT NULL DEFAULT TRUE,
    org_path      TEXT NOT NULL UNIQUE
);
