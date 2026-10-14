/**
 * Seed Chart of Accounts from Unity institutional list (exact FE registry names).
 * Source: original UnityMasterFinanceRegistry COA (coa_unity.json)
 *
 *   node seeders/seedCoa.js
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { pool } = require('../config/db');
const { randomUUID } = require('crypto');

async function main() {
  const file = path.join(__dirname, 'coa_unity.json');
  if (!fs.existsSync(file)) {
    console.error('Missing seeders/coa_unity.json');
    process.exit(1);
  }
  const list = JSON.parse(fs.readFileSync(file, 'utf8'));
  console.log('Loading Unity COA rows:', list.length);

  await pool.query('TRUNCATE TABLE fin_coa CASCADE');

  let n = 0;
  for (const r of list) {
    const account_no = String(r.account_no || r.id || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(account_no)) continue;
    const name = r.name || r.desc || r.short_desc || account_no;
    const account_type = r.account_type || r.type || 'Asset';
    const payload = {
      base: r.base,
      cc: r.cc,
      loc: r.loc,
      unit: r.unit,
      campus: r.campus,
      short_desc: r.short_desc || name,
    };
    await pool.query(
      `INSERT INTO fin_coa (id, account_no, account_code, name, account_type, parent_code, is_active, payload)
       VALUES ($1,$2,$2,$3,$4,NULL,TRUE,$5::jsonb)
       ON CONFLICT (account_no) DO UPDATE SET
         name = EXCLUDED.name,
         account_type = EXCLUDED.account_type,
         account_code = EXCLUDED.account_code,
         payload = EXCLUDED.payload,
         is_active = TRUE`,
      [randomUUID(), account_no, name, account_type, JSON.stringify(payload)]
    );
    n++;
    if (n % 200 === 0) console.log('  …', n);
  }

  const c = await pool.query('SELECT count(*)::int AS c FROM fin_coa');
  const s = await pool.query(
    `SELECT account_no, name, account_type FROM fin_coa
     WHERE account_no IN ('1110-01-10','1141-01-10','6110-01-10')
     ORDER BY account_no`
  );
  console.log('fin_coa count =', c.rows[0].c);
  console.log('samples:');
  s.rows.forEach((r) => console.log(' ', r.account_no, r.account_type, r.name));
  await pool.end();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
