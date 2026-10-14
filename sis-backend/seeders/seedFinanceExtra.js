/**
 * Partners + payroll demo rows (process, OT, A/D, master from HR).
 *   node seeders/seedFinanceExtra.js
 */
require('dotenv').config();
const { pool } = require('../config/db');
const { randomUUID } = require('crypto');

async function main() {
  console.log('Seeding partners + payroll demo…');

  // ── Partners (rich) ──────────────────────────────────────────────
  await pool.query('TRUNCATE fin_partners CASCADE');
  const partners = [
    ['CBE Main Branch', 'Bank', '1210-01-10', '2110-01-10', '1141-01-10'],
    ['Awash Bank', 'Bank', '1210-01-10', '2110-01-10', '1142-01-10'],
    ['Office Supplies Plc', 'Supplier', null, '2110-01-10', null],
    ['Import Trade ET', 'Supplier', null, '2110-01-10', null],
    ['Campus Maintenance Co', 'Supplier', null, '2110-01-10', null],
    ['Ethio Telecom', 'Supplier', null, '2110-01-10', null],
    ['Students Receivable Pool', 'Customer', '1210-01-10', null, null],
    ['Sponsor – Ministry', 'Customer', '1210-01-11', null, null],
    ['Research Grant Partner', 'Customer', '1210-01-10', null, null],
    ['Unity Staff Credit Assoc', 'Other', null, '2280-01-10', null],
    ['Pension Trust Fund', 'Other', null, '2273-01-10', null],
    ['Tax Authority (ERCA)', 'Government', null, '2270-01-10', null],
  ];
  for (const [name, type, ar, ap, bank] of partners) {
    await pool.query(
      `INSERT INTO fin_partners (id, partner_name, partner_type, default_ar_account, default_ap_account, default_bank_account, payload)
       VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb)`,
      [randomUUID(), name, type, ar, ap, bank, JSON.stringify({ status: 'Active', country: 'ET' })]
    );
  }

  // ── Pay master from HR employees ─────────────────────────────────
  await pool.query('TRUNCATE pay_master CASCADE');
  const emp = await pool.query(
    `SELECT id_number, name, job_title, department, gross_salary, tin_no, pension_id FROM hr_employees ORDER BY id_number`
  );
  for (const e of emp.rows) {
    await pool.query(
      `INSERT INTO pay_master (
        id, employee_name, id_number, position, department, basic_salary,
        doctoral_incentive, pos_allowance, cash_indemnity, tin_no, pension_id, from_hr, status
      ) VALUES ($1,$2,$3,$4,$5,$6,0,0,0,$7,$8,TRUE,'Active')`,
      [randomUUID(), e.name, e.id_number, e.job_title, e.department, e.gross_salary || 0, e.tin_no, e.pension_id]
    );
  }

  // ── Pay process (current period, first 8 staff) ──────────────────
  await pool.query('TRUNCATE pay_process CASCADE');
  const period = new Date().toISOString().slice(0, 7);
  const today = new Date().toISOString().slice(0, 10);
  let n = 0;
  for (const e of emp.rows.slice(0, 8)) {
    const basic = Number(e.gross_salary || 0);
    const pension = Math.round(basic * 0.07 * 100) / 100;
    const tax = Math.round(basic * 0.1 * 100) / 100;
    const net = Math.round((basic - pension - tax) * 100) / 100;
    await pool.query(
      `INSERT INTO pay_process (
        id, payroll_process_date, period, employee_name, id_number, position,
        tin_no, pension_id, total_basic, gross_earnings, total_tax, total_pension,
        total_deductions, total_net, status, posted_to_ledger
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$9,$10,$11,$12,$13,'Draft',FALSE)`,
      [randomUUID(), today, period, e.name, e.id_number, e.job_title, e.tin_no, e.pension_id, basic, tax, pension, pension + tax, net]
    );
    n++;
  }

  // ── Overtime ─────────────────────────────────────────────────────
  await pool.query('TRUNCATE pay_overtime CASCADE');
  for (const e of emp.rows.slice(0, 5)) {
    const basic = Number(e.gross_salary || 20000);
    const hourly = basic / 160;
    const dayH = 8;
    const nightH = 2;
    const total = Math.round((hourly * 1.25 * dayH + hourly * 1.5 * nightH) * 100) / 100;
    await pool.query(
      `INSERT INTO pay_overtime (
        id, period, employee_name, id_number, position, basic_salary, hourly_rate,
        day_hrs, night_hrs, total_ot_pay, status
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,'Draft')`,
      [randomUUID(), period, e.name, e.id_number, e.job_title, basic, Math.round(hourly * 100) / 100, dayH, nightH, total]
    );
  }

  // ── Allowance / Deduction ────────────────────────────────────────
  await pool.query('TRUNCATE pay_ad CASCADE');
  const ads = [
    [emp.rows[0], 'Transport Allowance', 'Allowance', 'No', 1500, 'pos_allowance'],
    [emp.rows[1], 'Housing Allowance', 'Allowance', 'No', 2000, 'pos_allowance'],
    [emp.rows[2], 'Credit Association', 'Deduction', 'No', 500, 'credit_association'],
    [emp.rows[3], 'Staff Loan Recovery', 'Deduction', 'No', 800, 'other_deduction'],
    [emp.rows[0], 'Position Allowance', 'Allowance', 'Yes', 1200, 'pos_allowance'],
  ];
  for (const [e, name, type, taxable, amount, maps] of ads) {
    if (!e) continue;
    await pool.query(
      `INSERT INTO pay_ad (id, period, employee_name, id_number, name, type, taxable, amount, maps_to_process, status)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,'Active')`,
      [randomUUID(), period, e.name, e.id_number, name, type, taxable, amount, maps]
    );
  }

  const counts = await pool.query(`
    SELECT 'partners' t, count(*)::int c FROM fin_partners
    UNION ALL SELECT 'pay_master', count(*) FROM pay_master
    UNION ALL SELECT 'pay_process', count(*) FROM pay_process
    UNION ALL SELECT 'pay_overtime', count(*) FROM pay_overtime
    UNION ALL SELECT 'pay_ad', count(*) FROM pay_ad
  `);
  counts.rows.forEach((r) => console.log(r.t, r.c));
  await pool.end();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
