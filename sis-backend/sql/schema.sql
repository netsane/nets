-- Unity SIS — PostgreSQL schema (Finance + HR + Material)
-- Run: psql -U postgres -d unity_sis -f sql/schema.sql

CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ---------------------------------------------------------------------------
-- Auth & RBAC
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS roles (
  id            TEXT PRIMARY KEY,
  name          TEXT NOT NULL UNIQUE,
  description   TEXT,
  permissions   JSONB NOT NULL DEFAULT '[]'::jsonb,
  modules       JSONB NOT NULL DEFAULT '["finance","hr","material"]'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS users (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  username      TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  display_name  TEXT,
  role_id       TEXT REFERENCES roles(id) ON DELETE SET NULL,
  employee_id   TEXT,
  is_active     BOOLEAN NOT NULL DEFAULT TRUE,
  modules       JSONB DEFAULT NULL,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_users_role ON users(role_id);

-- ---------------------------------------------------------------------------
-- Shared / org
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS departments (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  code          TEXT UNIQUE,
  name          TEXT NOT NULL,
  head_name     TEXT,
  status        TEXT DEFAULT 'Active',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ---------------------------------------------------------------------------
-- HR
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS hr_employees (
  id                    TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  id_number             TEXT UNIQUE,
  name                  TEXT NOT NULL,
  gender                TEXT,
  dob_gc                DATE,
  dob_ec                TEXT,
  age                   INT,
  marital_status        TEXT,
  seniority_date_gc     DATE,
  seniority_date_ec     TEXT,
  work_location         TEXT,
  department            TEXT,
  job_title             TEXT,
  reporting_to          TEXT,
  employment_category   TEXT,
  employment_type       TEXT,
  educational_status    TEXT,
  cumulative_gpa        TEXT,
  auth_status           TEXT,
  field_of_study        TEXT,
  gross_salary          NUMERIC(18,2) DEFAULT 0,
  family_contact        TEXT,
  pension_id            TEXT,
  retirement_date_ec    TEXT,
  national_id           TEXT,
  tin_no                TEXT,
  telephone_number      TEXT,
  office_tel            TEXT,
  ext_no                TEXT,
  emergency_contact_name TEXT,
  emergency_contact_telephone TEXT,
  emergency_contact_address TEXT,
  remarks               TEXT,
  system_role           TEXT DEFAULT 'employee',
  status                TEXT DEFAULT 'Active',
  payload               JSONB DEFAULT '{}'::jsonb,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_hr_emp_dept ON hr_employees(department);
CREATE INDEX IF NOT EXISTS idx_hr_emp_name ON hr_employees(name);

CREATE TABLE IF NOT EXISTS hr_leave (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  id_number     TEXT,
  leave_type    TEXT,
  start_date    DATE,
  end_date      DATE,
  days          NUMERIC(8,2),
  status        TEXT DEFAULT 'Pending',
  department    TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_recruitment (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  vacancy_title TEXT,
  department    TEXT,
  candidate     TEXT,
  status        TEXT DEFAULT 'Open',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_vacancies (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  title         TEXT,
  department    TEXT,
  status        TEXT DEFAULT 'Open',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_training (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  title         TEXT,
  employee_name TEXT,
  status        TEXT DEFAULT 'Planned',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_appraisal (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  score         NUMERIC(8,2),
  period        TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_insurance (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  amount        NUMERIC(18,2),
  department    TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_attrition (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  reason        TEXT,
  effective_date DATE,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_awards (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  category      TEXT,
  amount        NUMERIC(18,2) DEFAULT 0,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS hr_penalties (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT,
  reason        TEXT,
  amount        NUMERIC(18,2) DEFAULT 0,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Generic HR bag for other CRUD screens
CREATE TABLE IF NOT EXISTS hr_records (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  collection    TEXT NOT NULL,
  payload       JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_hr_records_col ON hr_records(collection);

-- ---------------------------------------------------------------------------
-- Finance
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS fin_coa (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  account_no    TEXT UNIQUE,
  account_code  TEXT,
  name          TEXT,
  account_type  TEXT,
  parent_code   TEXT,
  is_active     BOOLEAN DEFAULT TRUE,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS fin_partners (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  partner_name  TEXT NOT NULL,
  partner_type  TEXT,
  default_ar_account TEXT,
  default_ap_account TEXT,
  default_bank_account TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS fin_transactions (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  date          DATE,
  partner_name  TEXT,
  account_id    TEXT,
  account_code  TEXT,
  account_name  TEXT,
  type          TEXT,
  amount        NUMERIC(18,2) DEFAULT 0,
  debit         NUMERIC(18,2) DEFAULT 0,
  credit        NUMERIC(18,2) DEFAULT 0,
  balancing_account TEXT,
  narration     TEXT,
  description   TEXT,
  status        TEXT DEFAULT 'Posted',
  module        TEXT DEFAULT 'FINANCE',
  source        TEXT,
  material_ref  TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_fin_tx_date ON fin_transactions(date);
CREATE INDEX IF NOT EXISTS idx_fin_tx_module ON fin_transactions(module);
CREATE INDEX IF NOT EXISTS idx_fin_tx_source ON fin_transactions(source);

CREATE TABLE IF NOT EXISTS fin_day_journals (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  date          DATE,
  account_id    TEXT,
  account_code  TEXT,
  type          TEXT,
  amount        NUMERIC(18,2) DEFAULT 0,
  debit         NUMERIC(18,2) DEFAULT 0,
  credit        NUMERIC(18,2) DEFAULT 0,
  narration     TEXT,
  module        TEXT DEFAULT 'FINANCE',
  source        TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pay_master (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  employee_name TEXT NOT NULL,
  id_number     TEXT,
  position      TEXT,
  department    TEXT,
  cost_center   TEXT,
  location_code TEXT,
  basic_salary  NUMERIC(18,2) DEFAULT 0,
  doctoral_incentive NUMERIC(18,2) DEFAULT 0,
  pos_allowance NUMERIC(18,2) DEFAULT 0,
  cash_indemnity NUMERIC(18,2) DEFAULT 0,
  tin_no        TEXT,
  pension_id    TEXT,
  saving_acc    TEXT,
  status        TEXT DEFAULT 'Active',
  from_hr       BOOLEAN DEFAULT FALSE,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pay_process (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  payroll_process_date DATE,
  period        TEXT,
  employee_name TEXT,
  id_number     TEXT,
  position      TEXT,
  tin_no        TEXT,
  pension_id    TEXT,
  saving_acc    TEXT,
  cost_center   TEXT,
  location_code TEXT,
  total_basic   NUMERIC(18,2) DEFAULT 0,
  doctoral_incentive NUMERIC(18,2) DEFAULT 0,
  pos_allowance NUMERIC(18,2) DEFAULT 0,
  cash_indemnity NUMERIC(18,2) DEFAULT 0,
  total_ot_pay  NUMERIC(18,2) DEFAULT 0,
  working_days  NUMERIC(8,2),
  absent_days   NUMERIC(8,2),
  absent_amt    NUMERIC(18,2) DEFAULT 0,
  gross_earnings NUMERIC(18,2) DEFAULT 0,
  total_tax     NUMERIC(18,2) DEFAULT 0,
  total_pension NUMERIC(18,2) DEFAULT 0,
  employer_pension NUMERIC(18,2) DEFAULT 0,
  cost_sharing  NUMERIC(18,2) DEFAULT 0,
  credit_association NUMERIC(18,2) DEFAULT 0,
  sundry_debtors NUMERIC(18,2) DEFAULT 0,
  total_deductions NUMERIC(18,2) DEFAULT 0,
  total_net     NUMERIC(18,2) DEFAULT 0,
  status        TEXT DEFAULT 'Draft',
  posted_to_ledger BOOLEAN DEFAULT FALSE,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pay_overtime (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  period        TEXT,
  employee_name TEXT,
  id_number     TEXT,
  position      TEXT,
  basic_salary  NUMERIC(18,2) DEFAULT 0,
  hourly_rate   NUMERIC(18,4) DEFAULT 0,
  day_hrs       NUMERIC(10,2) DEFAULT 0,
  night_hrs     NUMERIC(10,2) DEFAULT 0,
  weekend_hrs   NUMERIC(10,2) DEFAULT 0,
  holiday_hrs   NUMERIC(10,2) DEFAULT 0,
  total_ot_pay  NUMERIC(18,2) DEFAULT 0,
  status        TEXT DEFAULT 'Draft',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pay_ad (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  period        TEXT,
  employee_name TEXT,
  id_number     TEXT,
  name          TEXT,
  type          TEXT,
  taxable       TEXT,
  amount        NUMERIC(18,2) DEFAULT 0,
  maps_to_process TEXT,
  status        TEXT DEFAULT 'Active',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pay_factors (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  pension_employee NUMERIC(8,4) DEFAULT 0.07,
  pension_employer NUMERIC(8,4) DEFAULT 0.11,
  working_hours NUMERIC(10,2) DEFAULT 160,
  ot_day        NUMERIC(8,4) DEFAULT 1.25,
  ot_night      NUMERIC(8,4) DEFAULT 1.50,
  ot_weekend    NUMERIC(8,4) DEFAULT 2.00,
  ot_holiday    NUMERIC(8,4) DEFAULT 2.50,
  tax_json      JSONB,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS fin_records (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  collection    TEXT NOT NULL,
  payload       JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_fin_records_col ON fin_records(collection);

-- ---------------------------------------------------------------------------
-- Material
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS mat_inventory (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  sku           TEXT,
  item_name     TEXT NOT NULL,
  category      TEXT,
  uom           TEXT,
  qty_on_hand   NUMERIC(18,4) DEFAULT 0,
  unit_cost     NUMERIC(18,4) DEFAULT 0,
  location      TEXT,
  coa_account   TEXT,
  status        TEXT DEFAULT 'Active',
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_grv (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  grv_no        TEXT,
  date          DATE,
  supplier      TEXT,
  item_name     TEXT,
  qty_received  NUMERIC(18,4) DEFAULT 0,
  unit_cost     NUMERIC(18,4) DEFAULT 0,
  line_value    NUMERIC(18,2) DEFAULT 0,
  coa_account   TEXT,
  posted_to_finance BOOLEAN DEFAULT FALSE,
  posted_at     DATE,
  posted_ref    TEXT,
  finance_status TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_giv (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  giv_no        TEXT,
  date          DATE,
  department    TEXT,
  item_name     TEXT,
  quantity      NUMERIC(18,4) DEFAULT 0,
  unit_cost     NUMERIC(18,4) DEFAULT 0,
  line_value    NUMERIC(18,2) DEFAULT 0,
  coa_account   TEXT,
  coa_expense   TEXT,
  posted_to_finance BOOLEAN DEFAULT FALSE,
  posted_at     DATE,
  posted_ref    TEXT,
  finance_status TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_po_local (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  lpo_no        TEXT,
  date          DATE,
  supplier      TEXT,
  item_name     TEXT,
  total_value   NUMERIC(18,2) DEFAULT 0,
  coa_account   TEXT,
  posted_to_finance BOOLEAN DEFAULT FALSE,
  posted_at     DATE,
  posted_ref    TEXT,
  finance_status TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_po_foreign (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  import_ref    TEXT,
  date          DATE,
  supplier      TEXT,
  item_name     TEXT,
  fob_value     NUMERIC(18,2) DEFAULT 0,
  freight_cost  NUMERIC(18,2) DEFAULT 0,
  insurance_cost NUMERIC(18,2) DEFAULT 0,
  customs_duty  NUMERIC(18,2) DEFAULT 0,
  shipping_cost NUMERIC(18,2) DEFAULT 0,
  total_landed  NUMERIC(18,2) DEFAULT 0,
  coa_account   TEXT,
  posted_to_finance BOOLEAN DEFAULT FALSE,
  posted_at     DATE,
  posted_ref    TEXT,
  finance_status TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_robbing (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  robbing_no    TEXT,
  date          DATE,
  part_taken    TEXT,
  part_value    NUMERIC(18,2) DEFAULT 0,
  line_value    NUMERIC(18,2) DEFAULT 0,
  posted_to_finance BOOLEAN DEFAULT FALSE,
  posted_at     DATE,
  posted_ref    TEXT,
  finance_status TEXT,
  payload       JSONB DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS mat_records (
  id            TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  collection    TEXT NOT NULL,
  payload       JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_mat_records_col ON mat_records(collection);

-- ---------------------------------------------------------------------------
-- Audit
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS audit_log (
  id            BIGSERIAL PRIMARY KEY,
  actor         TEXT,
  action        TEXT,
  module        TEXT,
  entity        TEXT,
  entity_id     TEXT,
  detail        JSONB,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
