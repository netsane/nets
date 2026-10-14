/**
 * Rich finance DEMO data in PostgreSQL — transactions + day journals.
 * Run AFTER seedCoa.js so COA codes exist.
 *
 *   node seeders/seedFinanceDemo.js
 */
require('dotenv').config();
const { pool } = require('../config/db');
const { randomUUID } = require('crypto');

function leg(date, code, name, type, amount, narration, ref, extra = {}) {
  const isDr = String(type).toLowerCase().startsWith('d');
  const amt = Number(Number(amount).toFixed(2));
  return {
    id: randomUUID(),
    date,
    account_id: code,
    account_code: code,
    account_name: name,
    type: isDr ? 'Debit' : 'Credit',
    debit: isDr ? amt : 0,
    credit: isDr ? 0 : amt,
    amount: amt,
    narration,
    description: narration,
    partner_name: extra.partner || 'Unity University',
    status: 'Posted',
    module: extra.module || 'FINANCE',
    source: extra.source || 'DEMO_SEED',
    reference: ref,
  };
}

function buildAll() {
  const rows = [];
  const add = (...xs) => rows.push(...xs);

  add(
    leg('2026-01-01', '1141-01-10', 'Cash at Bank CBE, -6512', 'Debit', 2500000, 'Opening capital deposit — CBE main', 'OPEN-CAP-01'),
    leg('2026-01-01', '3100-01-10', 'Capital / Fund Balance', 'Credit', 2500000, 'Opening capital', 'OPEN-CAP-01')
  );
  add(
    leg('2026-01-02', '1110-01-10', 'Cash on hand-Main Office,West-', 'Debit', 50000, 'Till float from bank', 'OPEN-CASH-01'),
    leg('2026-01-02', '1141-01-10', 'Cash at Bank CBE, -6512', 'Credit', 50000, 'Transfer to cash on hand', 'OPEN-CASH-01')
  );
  add(
    leg('2026-02-15', '1210-01-10', 'Trade Debtors', 'Debit', 850000, 'Tuition & fees billed', 'OP-REV-01', { partner: 'Students' }),
    leg('2026-02-15', '4110-01-10', 'Tuition Fees – Undergraduate', 'Credit', 850000, 'Tuition revenue recognition', 'OP-REV-01', { partner: 'Students' })
  );
  add(
    leg('2026-03-10', '1141-01-10', 'Cash at Bank CBE, -6512', 'Debit', 620000, 'Student fee collections', 'OP-COL-01', { partner: 'Students' }),
    leg('2026-03-10', '1210-01-10', 'Trade Debtors', 'Credit', 620000, 'AR settlement', 'OP-COL-01', { partner: 'Students' })
  );
  add(
    leg('2026-03-20', '6310-01-10', 'Material / Consumables', 'Debit', 45000, 'Academic / admin supplies', 'OP-AP-01', { partner: 'Office Supplies Plc' }),
    leg('2026-03-20', '2110-01-10', 'Trade Creditors', 'Credit', 45000, 'Payable to suppliers', 'OP-AP-01', { partner: 'Office Supplies Plc' })
  );
  add(
    leg('2026-04-05', '1141-01-10', 'Cash at Bank CBE, -6512', 'Debit', 120000, 'Research grant received', 'OP-GRANT-01'),
    leg('2026-04-05', '4210-01-10', 'Research Grants', 'Credit', 120000, 'Grant income', 'OP-GRANT-01')
  );
  add(
    leg('2026-05-12', '6330-01-10', 'Maintenance & Repairs', 'Debit', 18000, 'Campus maintenance', 'OP-MNT-01'),
    leg('2026-05-12', '1141-01-10', 'Cash at Bank CBE, -6512', 'Credit', 18000, 'Paid maintenance', 'OP-MNT-01')
  );
  add(
    leg('2026-06-01', '6500-01-10', 'General Admin', 'Debit', 22000, 'Admin expenses', 'OP-ADM-01'),
    leg('2026-06-01', '1141-01-10', 'Cash at Bank CBE, -6512', 'Credit', 22000, 'Paid admin', 'OP-ADM-01')
  );

  const gross = 42500;
  const eePen = Math.round(gross * 0.07 * 100) / 100;
  const erPen = Math.round(gross * 0.11 * 100) / 100;
  const tax = 0;
  const net = Math.round((gross - eePen - tax) * 100) / 100;
  add(
    leg('2026-07-31', '6110-01-10', 'Basic Salary', 'Debit', gross, 'Payroll gross', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' }),
    leg('2026-07-31', '6120-01-10', 'Employer Pension', 'Debit', erPen, 'ER pension 11%', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' }),
    leg('2026-07-31', '2240-01-10', 'Salary Payable', 'Credit', net, 'Net pay', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' }),
    leg('2026-07-31', '2270-01-10', 'Income Tax Payable', 'Credit', tax, 'PAYE', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' }),
    leg('2026-07-31', '2273-01-10', 'Pension Payable – Employee', 'Credit', eePen, 'EE pension 7%', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' }),
    leg('2026-07-31', '2274-01-10', 'Pension Payable – Employer', 'Credit', erPen, 'ER pension payable', 'PAY-2026-07', { module: 'PAYROLL', source: 'PAYROLL' })
  );

  return rows;
}

async function main() {
  console.log('Finance DEMO seed → fin_transactions + fin_day_journals');

  await pool.query(`DELETE FROM fin_transactions WHERE source IN ('DEMO_SEED', 'SEED', 'PAYROLL') OR narration ILIKE 'Opening%'`);
  await pool.query(`DELETE FROM fin_day_journals WHERE source IN ('DEMO_SEED', 'SEED')`);

  const txs = buildAll();
  let dr = 0;
  let cr = 0;

  for (const t of txs) {
    dr += Number(t.debit);
    cr += Number(t.credit);
    await pool.query(
      `INSERT INTO fin_transactions (
        id, date, partner_name, account_id, account_code, account_name,
        type, amount, debit, credit, narration, description, status, module, source
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15)`,
      [
        t.id, t.date, t.partner_name, t.account_id, t.account_code, t.account_name,
        t.type, t.amount, t.debit, t.credit, t.narration, t.description,
        t.status, t.module, t.source,
      ]
    );
    // Mirror into day journals for GL/journal screens
    await pool.query(
      `INSERT INTO fin_day_journals (
        id, date, account_id, account_code, type, amount, debit, credit, narration, module, source
      ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
      [
        randomUUID(), t.date, t.account_id, t.account_code, t.type,
        t.amount, t.debit, t.credit, t.narration, t.module, 'DEMO_SEED',
      ]
    );
  }

  const c1 = await pool.query('SELECT count(*)::int AS c FROM fin_transactions');
  const c2 = await pool.query('SELECT count(*)::int AS c FROM fin_day_journals');
  const bal = await pool.query(
    `SELECT COALESCE(SUM(debit),0)::float AS dr, COALESCE(SUM(credit),0)::float AS cr FROM fin_transactions`
  );

  console.log('fin_transactions:', c1.rows[0].c);
  console.log('fin_day_journals:', c2.rows[0].c);
  console.log('Σ Dr', bal.rows[0].dr, 'Σ Cr', bal.rows[0].cr);
  console.log('Balanced:', Math.abs(bal.rows[0].dr - bal.rows[0].cr) < 0.05);
  if (Math.abs(dr - cr) >= 0.05) {
    console.warn('Builder imbalance Dr', dr, 'Cr', cr);
  }
  await pool.end();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
