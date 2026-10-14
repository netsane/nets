/**
 * Full seed for Finance + HR + Material — all data lives in PostgreSQL.
 * Run after schema:  npm run db:seed
 */
require('dotenv').config();
const bcrypt = require('bcryptjs');
const { pool } = require('../config/db');
const { randomUUID } = require('crypto');

const uid = () => randomUUID();
const today = () => new Date().toISOString().slice(0, 10);

async function clearSeedTables() {
  // Order: children first
  const tables = [
    'audit_log',
    'fin_day_journals', 'fin_transactions', 'fin_records',
    'pay_process', 'pay_overtime', 'pay_ad', 'pay_master', 'pay_factors',
    'fin_partners', 'fin_coa',
    'mat_robbing', 'mat_po_foreign', 'mat_po_local', 'mat_giv', 'mat_grv', 'mat_inventory', 'mat_records',
    'hr_penalties', 'hr_awards', 'hr_attrition', 'hr_insurance', 'hr_appraisal',
    'hr_training', 'hr_vacancies', 'hr_recruitment', 'hr_leave', 'hr_employees', 'hr_records',
    'departments',
    'users', 'roles',
  ];
  for (const t of tables) {
    try {
      await pool.query(`TRUNCATE TABLE ${t} CASCADE`);
    } catch (e) {
      /* table may not exist */
    }
  }
}

async function main() {
  console.log('Seeding Unity SIS (full)…');
  await clearSeedTables();

  const hash = await bcrypt.hash('password', 10);

  // ── Roles ──────────────────────────────────────────────────────────
  await pool.query(`
    INSERT INTO roles (id, name, description, permissions, modules) VALUES
      ('admin', 'admin', 'Full system access', '[]'::jsonb, '["finance","hr","material"]'::jsonb),
      ('finance_officer', 'finance_officer', 'Finance module', '[]'::jsonb, '["finance"]'::jsonb),
      ('hr_manager', 'hr_manager', 'HR + payroll view', '[]'::jsonb, '["hr","finance"]'::jsonb),
      ('hr_staff', 'hr_staff', 'HR operations', '[]'::jsonb, '["hr"]'::jsonb),
      ('store_manager', 'store_manager', 'Material stores', '[]'::jsonb, '["material"]'::jsonb),
      ('employee', 'employee', 'Self service', '[]'::jsonb, '["hr"]'::jsonb)
    ON CONFLICT (id) DO NOTHING
  `);

  // ── Users ──────────────────────────────────────────────────────────
  const users = [
    ['admin', 'System Administrator', 'admin', '["finance","hr","material"]'],
    ['hr_manager', 'HR Manager', 'hr_manager', '["hr","finance"]'],
    ['hr_staff', 'HR Staff', 'hr_staff', '["hr"]'],
    ['finance_officer', 'Finance Officer', 'finance_officer', '["finance"]'],
    ['store_manager', 'Store Manager', 'store_manager', '["material"]'],
    ['employee', 'Demo Employee', 'employee', '["hr"]'],
  ];
  for (const [username, name, role, modules] of users) {
    await pool.query(
      `INSERT INTO users (id, username, password_hash, display_name, role_id, modules)
       VALUES ($1,$2,$3,$4,$5,$6::jsonb)`,
      [uid(), username, hash, name, role, modules]
    );
  }

  // ── Departments ────────────────────────────────────────────────────
  const depts = [
    ['ENG', 'Engineering'],
    ['FIN', 'Finance'],
    ['HR', 'Human Resource'],
    ['MAT', 'Material Management'],
    ['PRE', 'President Office'],
    ['CEO', 'CEO Office'],
    ['REG', 'Registrar'],
    ['IT', 'IT Department'],
    ['MKT', 'Marketing'],
    ['OPS', 'Operations'],
    ['ACD', 'Academic Department'],
  ];
  for (const [code, name] of depts) {
    await pool.query(
      `INSERT INTO departments (id, code, name, status) VALUES ($1,$2,$3,'Active')`,
      [uid(), code, name]
    );
  }

  // ── HR Employees (matches frontend roster) ─────────────────────────
  const employees = [
    ['UU-001', 'Netsanet Abebe', 'Female', '1977-04-21', 49, 'Married', '2021-10-18', 'Burayu', 'Engineering', 'Officer 2', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 23293, 'PEN-1001', 'TIN-3888', '+251995483330', 'employee'],
    ['UU-002', 'Damte Tesfaye', 'Male', '1993-09-02', 33, 'Divorced', '2024-07-02', 'Dessie', 'Human Resource', 'Lecturer 3', 'Academic', 'Permanent', 'Masters Degree', 'Advanced Academics', 26446, 'PEN-1002', 'TIN-41210', '+251989234696', 'employee'],
    ['UU-003', 'Keni Hailu', 'Female', '1965-10-27', 61, 'Widowed', '2021-03-03', 'Adama', 'President Office', 'Officer 1', 'Non-Academic', 'Non-Permanent', 'Bachelors Degree', 'Operational Management', 20192, 'PEN-1003', 'TIN-63223', '+251920144142', 'employee'],
    ['UU-004', 'Kalab Mekonnen', 'Male', '1981-10-04', 45, 'Single', '2021-10-04', 'Gerji', 'CEO Office', 'Lecturer 2', 'Academic', 'Permanent', 'Masters Degree', 'Advanced Academics', 23369, 'PEN-1004', 'TIN-55112', '+251911223344', 'employee'],
    ['UU-005', 'Hanna Bekele', 'Female', '1988-03-15', 38, 'Married', '2019-08-01', 'Bole', 'Registrar', 'Officer 2', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 21800, 'PEN-1005', 'TIN-66101', '+251922334455', 'employee'],
    ['UU-006', 'Yonas Tadesse', 'Male', '1990-11-20', 35, 'Married', '2020-01-10', 'Torhailoch', 'Finance', 'Accountant', 'Non-Academic', 'Permanent', 'Masters Degree', 'Operational Management', 27500, 'PEN-1006', 'TIN-77220', '+251933445566', 'finance_officer'],
    ['UU-007', 'Mulu Getachew', 'Female', '1985-06-08', 41, 'Single', '2018-05-12', 'Addis Ababa', 'Library', 'Librarian', 'Non-Academic', 'Permanent', 'Masters Degree', 'Operational Management', 20900, 'PEN-1007', 'TIN-88330', '+251944556677', 'employee'],
    ['UU-008', 'Abebe Kebede', 'Male', '1979-02-14', 47, 'Married', '2015-09-01', 'Burayu', 'Engineering', 'Lecturer 3', 'Academic', 'Permanent', 'PhD', 'Advanced Academics', 32000, 'PEN-1008', 'TIN-99440', '+251955667788', 'employee'],
    ['UU-009', 'Sara Hailu', 'Female', '1992-07-22', 34, 'Married', '2022-03-01', 'Gerji', 'Human Resource', 'HR Officer', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 19500, 'PEN-1009', 'TIN-10550', '+251966778899', 'hr_staff'],
    ['UU-010', 'Tesfaye Alemu', 'Male', '1983-12-01', 42, 'Divorced', '2017-11-15', 'Dessie', 'IT Department', 'IT Officer', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 24100, 'PEN-1010', 'TIN-11660', '+251977889900', 'employee'],
    ['UU-011', 'Birhane Wolde', 'Female', '1975-05-30', 51, 'Widowed', '2010-01-20', 'Adama', 'Marketing', 'Officer 1', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 18800, 'PEN-1011', 'TIN-12770', '+251988990011', 'employee'],
    ['UU-012', 'Gebre Michael', 'Male', '1987-09-09', 38, 'Married', '2019-04-01', 'Torhailoch', 'Registrar', 'Registrar', 'Non-Academic', 'Permanent', 'Masters Degree', 'Operational Management', 25600, 'PEN-1012', 'TIN-13880', '+251900112233', 'employee'],
    ['UU-013', 'Fikirte Asfa', 'Female', '1995-01-18', 31, 'Single', '2023-06-01', 'Bole', 'Finance', 'Cashier', 'Non-Academic', 'Contract', 'Diploma', 'Operational Management', 15200, 'PEN-1013', 'TIN-14990', '+251911223355', 'employee'],
    ['UU-014', 'Mehari Tekle', 'Male', '1980-08-25', 45, 'Married', '2016-02-01', 'Gerji', 'Administration', 'Administrator', 'Non-Academic', 'Permanent', 'Masters Degree', 'Operational Management', 28900, 'PEN-1014', 'TIN-15100', '+251922334466', 'admin'],
    ['UU-015', 'Elsabet Taye', 'Female', '1991-04-12', 35, 'Married', '2021-01-15', 'Burayu', 'Health Services', 'Nurse', 'Non-Academic', 'Permanent', 'Bachelors Degree', 'Operational Management', 21000, 'PEN-1015', 'TIN-16210', '+251933445577', 'employee'],
  ];

  for (const e of employees) {
    const [id_number, name, gender, dob_gc, age, marital_status, seniority_date_gc, work_location, department, job_title, employment_category, employment_type, educational_status, field_of_study, gross_salary, pension_id, tin_no, telephone_number, system_role] = e;
    await pool.query(
      `INSERT INTO hr_employees (
        id, id_number, name, gender, dob_gc, age, marital_status, seniority_date_gc,
        work_location, department, job_title, reporting_to, employment_category, employment_type,
        educational_status, auth_status, field_of_study, gross_salary, pension_id, tin_no,
        telephone_number, system_role, status
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,'Department Head',$12,$13,$14,'Approved',$15,$16,$17,$18,$19,$20,'Active')`,
      [
        uid(), id_number, name, gender, dob_gc, age, marital_status, seniority_date_gc,
        work_location, department, job_title, employment_category, employment_type,
        educational_status, field_of_study, gross_salary, pension_id, tin_no,
        telephone_number, system_role,
      ]
    );
  }

  // ── Leave ──────────────────────────────────────────────────────────
  const leaveRows = [
    ['Netsanet Unity', 'UU-001', 'Annual', '2026-08-01', '2026-08-05', 5, 'Approved', 'Engineering'],
    ['Damte Unity', 'UU-002', 'Sick', '2026-08-10', '2026-08-12', 2, 'Pending', 'Human Resource'],
    ['Kenu Unity', 'UU-003', 'Annual', '2026-09-01', '2026-09-10', 8, 'Submitted', 'President Office'],
    ['Selam Bekele', 'UU-011', 'Maternity', '2026-07-01', '2026-10-01', 90, 'Approved', 'Finance'],
    ['Yonas Haile', 'UU-012', 'Annual', '2026-08-15', '2026-08-20', 5, 'Pending', 'Engineering'],
  ];
  for (const [employee_name, id_number, leave_type, start_date, end_date, days, status, department] of leaveRows) {
    await pool.query(
      `INSERT INTO hr_leave (id, employee_name, id_number, leave_type, start_date, end_date, days, status, department)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
      [uid(), employee_name, id_number, leave_type, start_date, end_date, days, status, department]
    );
  }

  // ── Vacancies / Recruitment / Training / Appraisal ─────────────────
  await pool.query(`
    INSERT INTO hr_vacancies (id, title, department, status) VALUES
      ('${uid()}', 'Lecturer – Computer Science', 'Engineering', 'Open'),
      ('${uid()}', 'Accountant II', 'Finance', 'Open'),
      ('${uid()}', 'Store Assistant', 'Material Management', 'Published')
  `);
  await pool.query(`
    INSERT INTO hr_recruitment (id, vacancy_title, department, candidate, status) VALUES
      ('${uid()}', 'Lecturer – Computer Science', 'Engineering', 'Abebe Kebele', 'Interview'),
      ('${uid()}', 'Accountant II', 'Finance', 'Sara Lemma', 'Shortlisted'),
      ('${uid()}', 'Store Assistant', 'Material Management', 'Tigist Hailu', 'Open')
  `);
  await pool.query(`
    INSERT INTO hr_training (id, title, employee_name, status) VALUES
      ('${uid()}', 'IFRS Update 2026', 'Selam Bekele', 'Planned'),
      ('${uid()}', 'Pedagogy Workshop', 'Damte Unity', 'In Progress'),
      ('${uid()}', 'Inventory Control', 'Biruk Alemu', 'Completed')
  `);
  await pool.query(`
    INSERT INTO hr_appraisal (id, employee_name, score, period) VALUES
      ('${uid()}', 'Netsanet Unity', 4.2, '2025/26'),
      ('${uid()}', 'Damte Unity', 4.5, '2025/26'),
      ('${uid()}', 'Yonas Haile', 3.8, '2025/26')
  `);
  await pool.query(`
    INSERT INTO hr_insurance (id, employee_name, amount, department) VALUES
      ('${uid()}', 'Netsanet Unity', 8500, 'Engineering'),
      ('${uid()}', 'Damte Unity', 9200, 'Human Resource'),
      ('${uid()}', 'Selam Bekele', 7800, 'Finance')
  `);
  await pool.query(`
    INSERT INTO hr_awards (id, employee_name, category, amount) VALUES
      ('${uid()}', 'Damte Unity', 'Best Lecturer', 5000),
      ('${uid()}', 'Biruk Alemu', 'Service Excellence', 2000)
  `);
  await pool.query(`
    INSERT INTO hr_penalties (id, employee_name, reason, amount) VALUES
      ('${uid()}', 'employee', 'Late attendance (warning)', 0)
  `);
  await pool.query(`
    INSERT INTO hr_attrition (id, employee_name, reason, effective_date) VALUES
      ('${uid()}', 'Former Staff Demo', 'Resignation', '2026-01-15')
  `);

  // ── Finance COA ────────────────────────────────────────────────────
  const coa = [
    ['1110-01-10', 'Cash on Hand', 'Asset'],
    ['1141-01-10', 'Cash at Bank – CBE', 'Asset'],
    ['1210-01-10', 'Accounts Receivable', 'Asset'],
    ['1474-01-10', 'Inventory – Stores', 'Asset'],
    ['1490-01-10', 'Goods in Transit', 'Asset'],
    ['2110-01-10', 'Trade Creditors', 'Liability'],
    ['2115-01-10', 'GRNI', 'Liability'],
    ['2240-01-10', 'Salary Payable', 'Liability'],
    ['2270-01-10', 'Income Tax Payable', 'Liability'],
    ['2273-01-10', 'Pension Payable', 'Liability'],
    ['3100-01-10', 'Capital', 'Equity'],
    ['4100-01-10', 'Tuition Revenue', 'Revenue'],
    ['6110-01-10', 'Salary Expense', 'Expense'],
    ['6310-01-10', 'Material Expense', 'Expense'],
    ['6330-01-10', 'Maintenance Expense', 'Expense'],
    ['6500-01-10', 'Admin Expense', 'Expense'],
  ];
  for (const [account_no, name, account_type] of coa) {
    await pool.query(
      `INSERT INTO fin_coa (id, account_no, account_code, name, account_type, is_active)
       VALUES ($1,$2,$2,$3,$4,TRUE)`,
      [uid(), account_no, name, account_type]
    );
  }

  // ── Partners ───────────────────────────────────────────────────────
  const partners = [
    ['CBE Main Branch', 'Bank', '1210-01-10', '2110-01-10', '1141-01-10'],
    ['Office Supplies Plc', 'Supplier', null, '2110-01-10', null],
    ['Import Trade ET', 'Supplier', null, '2110-01-10', null],
    ['Unity Student Receivables', 'Customer', '1210-01-10', null, null],
  ];
  for (const [partner_name, partner_type, ar, ap, bank] of partners) {
    await pool.query(
      `INSERT INTO fin_partners (id, partner_name, partner_type, default_ar_account, default_ap_account, default_bank_account)
       VALUES ($1,$2,$3,$4,$5,$6)`,
      [uid(), partner_name, partner_type, ar, ap, bank]
    );
  }

  // ── Pay factors ────────────────────────────────────────────────────
  await pool.query(
    `INSERT INTO pay_factors (id, pension_employee, pension_employer, working_hours, ot_day, ot_night, ot_weekend, ot_holiday)
     VALUES ($1, 0.07, 0.11, 160, 1.25, 1.50, 2.00, 2.50)`,
    [uid()]
  );

  // ── Pay master (from HR salaries) ──────────────────────────────────
  for (const e of employees) {
    const [id_number, name, , , , , , , department, job_title, , , , , gross_salary, pension_id, tin_no] = e;
    await pool.query(
      `INSERT INTO pay_master (
        id, employee_name, id_number, position, department, basic_salary,
        doctoral_incentive, pos_allowance, cash_indemnity, tin_no, pension_id, from_hr, status
      ) VALUES ($1,$2,$3,$4,$5,$6,0,0,0,$7,$8,TRUE,'Active')`,
      [uid(), name, id_number, job_title, department, gross_salary, tin_no, pension_id]
    );
  }

  // ── Sample payroll process (current period) ────────────────────────
  const period = new Date().toISOString().slice(0, 7);
  for (const e of employees.slice(0, 5)) {
    const [id_number, name, , , , , , , , job_title, , , , , gross_salary, pension_id, tin_no] = e;
    const basic = Number(gross_salary);
    const pension = Math.round(basic * 0.07 * 100) / 100;
    const tax = Math.round(basic * 0.1 * 100) / 100;
 // simplified
    const net = Math.round((basic - pension - tax) * 100) / 100;
    await pool.query(
      `INSERT INTO pay_process (
        id, payroll_process_date, period, employee_name, id_number, position,
        tin_no, pension_id, total_basic, gross_earnings, total_tax, total_pension,
        total_deductions, total_net, status, posted_to_ledger
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$9,$10,$11,$12,$13,'Draft',FALSE)`,
      [uid(), today(), period, name, id_number, job_title, tin_no, pension_id, basic, tax, pension, pension + tax, net]
    );
  }

  // ── Sample OT ──────────────────────────────────────────────────────
  await pool.query(
    `INSERT INTO pay_overtime (id, period, employee_name, id_number, position, basic_salary, hourly_rate, day_hrs, night_hrs, total_ot_pay, status)
     VALUES ($1,$2,'Netsanet Unity','UU-001','Officer 2',23293,23293/160,10,4, round((23293/160.0)*1.25*10 + (23293/160.0)*1.5*4, 2),'Draft')`,
    [uid(), period]
  );

  // ── Sample A/D ─────────────────────────────────────────────────────
  await pool.query(
    `INSERT INTO pay_ad (id, period, employee_name, id_number, name, type, taxable, amount, maps_to_process, status)
     VALUES
       ($1,$2,'Damte Unity','UU-002','Transport Allowance','Allowance','No',1500,'pos_allowance','Active'),
       ($3,$2,'Netsanet Unity','UU-001','Credit Association','Deduction','No',500,'credit_association','Active')`,
    [uid(), period, uid()]
  );

  // ── Sample finance transactions ────────────────────────────────────
  const txDate = today();
  await pool.query(
    `INSERT INTO fin_transactions (id, date, partner_name, account_id, account_code, account_name, type, amount, debit, credit, narration, status, module, source)
     VALUES
       ($1,$2,'CBE Main Branch','1141-01-10','1141-01-10','Cash at Bank – CBE','Debit',500000,500000,0,'Opening bank balance','Posted','FINANCE','SEED'),
       ($3,$2,'CBE Main Branch','3100-01-10','3100-01-10','Capital','Credit',500000,0,500000,'Opening capital','Posted','FINANCE','SEED')`,
    [uid(), txDate, uid()]
  );

  // ── Material inventory ─────────────────────────────────────────────
  const items = [
    ['SKU-A4-80', 'A4 Paper Ream 80gsm', 'Stationery', 'REAM', 120, 280, 'Main Store', '1474-01-10'],
    ['SKU-PEN-BL', 'Ballpoint Pen Blue', 'Stationery', 'BOX', 45, 150, 'Main Store', '1474-01-10'],
    ['SKU-TONER', 'Laser Toner HP 26A', 'IT Consumable', 'PCS', 8, 4200, 'IT Store', '1474-01-10'],
    ['SKU-CHAIR', 'Office Chair Standard', 'Furniture', 'PCS', 15, 3500, 'Main Store', '1474-01-10'],
    ['SKU-LAB-GL', 'Lab Gloves (box)', 'Lab', 'BOX', 30, 450, 'Lab Store', '1474-01-10'],
    ['SKU-CLEAN', 'Floor Cleaner 5L', 'Cleaning', 'PCS', 40, 320, 'Main Store', '1474-01-10'],
  ];
  for (const [sku, item_name, category, uom, qty_on_hand, unit_cost, location, coa_account] of items) {
    await pool.query(
      `INSERT INTO mat_inventory (id, sku, item_name, category, uom, qty_on_hand, unit_cost, location, coa_account, status)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,'Active')`,
      [uid(), sku, item_name, category, uom, qty_on_hand, unit_cost, location, coa_account]
    );
  }

  // ── GRV ────────────────────────────────────────────────────────────
  const grvId = uid();
  await pool.query(
    `INSERT INTO mat_grv (id, grv_no, date, supplier, item_name, qty_received, unit_cost, line_value, coa_account, posted_to_finance, finance_status)
     VALUES
       ($1,'GRV-2026-001',$2,'Office Supplies Plc','A4 Paper Ream 80gsm',50,280,14000,'1474-01-10',FALSE,'Pending'),
       ($3,'GRV-2026-002',$2,'Office Supplies Plc','Ballpoint Pen Blue',20,150,3000,'1474-01-10',FALSE,'Pending'),
       ($4,'GRV-2026-003',$2,'Import Trade ET','Laser Toner HP 26A',5,4200,21000,'1474-01-10',FALSE,'Pending')`,
    [grvId, txDate, uid(), uid()]
  );

  // ── GIV ────────────────────────────────────────────────────────────
  await pool.query(
    `INSERT INTO mat_giv (id, giv_no, date, department, item_name, quantity, unit_cost, line_value, coa_account, coa_expense, posted_to_finance, finance_status)
     VALUES
       ($1,'GIV-2026-001',$2,'Engineering','A4 Paper Ream 80gsm',10,280,2800,'1474-01-10','6310-01-10',FALSE,'Pending'),
       ($3,'GIV-2026-002',$2,'Finance','Ballpoint Pen Blue',2,150,300,'1474-01-10','6310-01-10',FALSE,'Pending')`,
    [uid(), txDate, uid()]
  );

  // ── Local / Foreign PO ─────────────────────────────────────────────
  await pool.query(
    `INSERT INTO mat_po_local (id, lpo_no, date, supplier, item_name, total_value, coa_account, posted_to_finance, finance_status)
     VALUES ($1,'LPO-2026-010',$2,'Office Supplies Plc','Office Chair Standard',17500,'1474-01-10',FALSE,'Pending')`,
    [uid(), txDate]
  );
  await pool.query(
    `INSERT INTO mat_po_foreign (id, import_ref, date, supplier, item_name, fob_value, freight_cost, insurance_cost, customs_duty, shipping_cost, total_landed, coa_account, posted_to_finance, finance_status)
     VALUES ($1,'IMP-2026-003',$2,'Import Trade ET','Lab Equipment Kit',80000,5000,1200,8000,3000,97200,'1490-01-10',FALSE,'Pending')`,
    [uid(), txDate]
  );

  // ── Robbing ────────────────────────────────────────────────────────
  await pool.query(
    `INSERT INTO mat_robbing (id, robbing_no, date, part_taken, part_value, line_value, posted_to_finance, finance_status)
     VALUES ($1,'ROB-2026-001',$2,'Printer drum from idle unit',1500,1500,FALSE,'Pending')`,
    [uid(), txDate]
  );

  // Counts
  const counts = await pool.query(`
    SELECT 'roles' AS t, count(*)::int AS c FROM roles
    UNION ALL SELECT 'users', count(*) FROM users
    UNION ALL SELECT 'departments', count(*) FROM departments
    UNION ALL SELECT 'hr_employees', count(*) FROM hr_employees
    UNION ALL SELECT 'hr_leave', count(*) FROM hr_leave
    UNION ALL SELECT 'fin_coa', count(*) FROM fin_coa
    UNION ALL SELECT 'pay_master', count(*) FROM pay_master
    UNION ALL SELECT 'pay_process', count(*) FROM pay_process
    UNION ALL SELECT 'mat_inventory', count(*) FROM mat_inventory
    UNION ALL SELECT 'mat_grv', count(*) FROM mat_grv
    UNION ALL SELECT 'mat_giv', count(*) FROM mat_giv
  `);
  console.log('Seed counts:');
  counts.rows.forEach((r) => console.log(' ', r.t, r.c));
  console.log('\nUsers (password = "password"): admin, hr_manager, hr_staff, finance_officer, store_manager, employee');
  console.log('All module master data is in PostgreSQL — not local mock generation.');
  console.log('Next (npm run db:seed runs these): seedCoa + seedFinanceDemo + seedFinanceExtra for COA, TB, partners, payroll.');
  await pool.end();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
