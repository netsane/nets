/**
 * Full demo seed + verify counts.
 *   node seeders/seedAll.js
 */
require('dotenv').config();
const { spawnSync } = require('child_process');
const path = require('path');
const { pool } = require('../config/db');

function run(script) {
  console.log('\n>>>', script);
  const r = spawnSync(process.execPath, [path.join(__dirname, script)], {
    stdio: 'inherit',
    cwd: path.join(__dirname, '..'),
    env: process.env,
  });
  if (r.status !== 0) {
    console.error('FAILED', script, 'exit', r.status);
    process.exit(r.status || 1);
  }
}

async function verify() {
  const q = async (sql) => (await pool.query(sql)).rows;
  const rows = await q(`
    SELECT 'users' t, count(*)::int c FROM users
    UNION ALL SELECT 'hr_employees', count(*) FROM hr_employees
    UNION ALL SELECT 'fin_coa', count(*) FROM fin_coa
    UNION ALL SELECT 'fin_transactions', count(*) FROM fin_transactions
    UNION ALL SELECT 'fin_day_journals', count(*) FROM fin_day_journals
    UNION ALL SELECT 'fin_partners', count(*) FROM fin_partners
    UNION ALL SELECT 'pay_process', count(*) FROM pay_process
    UNION ALL SELECT 'pay_overtime', count(*) FROM pay_overtime
    UNION ALL SELECT 'pay_ad', count(*) FROM pay_ad
    UNION ALL SELECT 'pay_factors', count(*) FROM pay_factors
    UNION ALL SELECT 'pay_master', count(*) FROM pay_master
    ORDER BY 1
  `);
  console.log('\n=== VERIFY ===');
  rows.forEach((r) => console.log(String(r.t).padEnd(22), r.c));
  const map = Object.fromEntries(rows.map((r) => [r.t, r.c]));
  const problems = [];
  if (!map.users) problems.push('users=0 — login will fail');
  if (!map.fin_coa) problems.push('fin_coa=0 — run seedCoa');
  if (!map.fin_transactions) problems.push('fin_transactions=0 — run seedFinanceDemo');
  if (!map.hr_employees) problems.push('hr_employees=0');
  if (problems.length) {
    console.log('\nPROBLEMS:');
    problems.forEach((p) => console.log(' -', p));
    process.exit(1);
  }
  console.log('\nOK — login admin / password then open the app.');
  await pool.end();
}

(async () => {
  run('seed.js');
  run('seedCoa.js');
  run('seedFinanceDemo.js');
  run('seedFinanceExtra.js');
  await verify();
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
