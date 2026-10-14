const { query } = require('../config/db');
const { money } = require('../utils/money');
const { randomUUID } = require('crypto');

const MAT_GL = {
  inventory: '1474-01-10',
  goodsInTransit: '1490-01-10',
  grni: '2115-01-10',
  tradeCreditors: '2110-01-10',
  materialExpense: '6310-01-10',
  maintenanceExpense: '6330-01-10',
};

const SOURCE_TABLE = {
  GRV: 'mat_grv',
  GIV: 'mat_giv',
  LOCAL_PO: 'mat_po_local',
  FOREIGN_PO: 'mat_po_foreign',
  ROBBING: 'mat_robbing',
};

function pushPair(legs, { date, partner, ref, uid, module, source, amt, drCode, drName, crCode, crName, narration }) {
  if (!amt || amt <= 0) return;
  const base = {
    date,
    partner_name: partner,
    module,
    source,
    material_ref: ref,
    narration,
    status: 'Posted',
  };
  legs.push({
    id: `${uid}-DR`,
    ...base,
    account_id: drCode,
    account_code: drCode,
    account_name: drName,
    type: 'Debit',
    amount: amt,
    debit: amt,
    credit: 0,
  });
  legs.push({
    id: `${uid}-CR`,
    ...base,
    account_id: crCode,
    account_code: crCode,
    account_name: crName,
    type: 'Credit',
    amount: amt,
    debit: 0,
    credit: amt,
  });
}

function buildMaterialJournalLegs(source, record) {
  if (!record) return [];
  const date = String(record.date || record.arrival_date || new Date().toISOString().slice(0, 10));
  const partner = record.supplier || record.department || 'Unity Stores';
  const uid = String(record.id || randomUUID());
  const legs = [];
  const desc = record.item_name || record.name || record.part_taken || 'Material';

  if (source === 'GRV') {
    const amt = money(record.line_value) || money(record.qty_received || record.quantity) * money(record.unit_cost);
    const inv = record.coa_account || MAT_GL.inventory;
    pushPair(legs, {
      date, partner, ref: record.grv_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: inv, drName: `Inventory · ${inv}`, crCode: MAT_GL.grni, crName: 'GRNI',
      narration: `GRV ${record.grv_no || ''} · ${desc}`,
    });
  } else if (source === 'GIV') {
    const amt = money(record.line_value) || money(record.quantity) * money(record.unit_cost);
    const exp = record.coa_expense || record.coa_account || MAT_GL.materialExpense;
    const inv = record.coa_account || MAT_GL.inventory;
    pushPair(legs, {
      date, partner, ref: record.giv_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: exp, drName: `Expense · ${exp}`, crCode: inv, crName: `Inventory · ${inv}`,
      narration: `GIV ${record.giv_no || ''} · ${desc} → ${record.department || ''}`,
    });
  } else if (source === 'LOCAL_PO') {
    const amt = money(record.total_value);
    const inv = record.coa_account || MAT_GL.inventory;
    pushPair(legs, {
      date, partner, ref: record.lpo_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: inv, drName: `Inventory · ${inv}`, crCode: MAT_GL.tradeCreditors, crName: 'Trade Creditors',
      narration: `Local PO ${record.lpo_no || ''} · ${desc}`,
    });
  } else if (source === 'FOREIGN_PO') {
    const amt =
      money(record.total_landed) ||
      money(record.fob_value) + money(record.freight_cost) + money(record.insurance_cost) +
      money(record.customs_duty) + money(record.shipping_cost);
    const git = record.coa_account || MAT_GL.goodsInTransit;
    pushPair(legs, {
      date, partner, ref: record.import_ref || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: git, drName: `GIT · ${git}`, crCode: MAT_GL.tradeCreditors, crName: 'Trade Creditors',
      narration: `Foreign import ${record.import_ref || ''} · ${desc}`,
    });
  } else if (source === 'ROBBING') {
    const amt = money(record.part_value || record.line_value);
    if (amt > 0) {
      pushPair(legs, {
        date, partner, ref: record.robbing_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
        drCode: MAT_GL.maintenanceExpense, drName: 'Maintenance', crCode: MAT_GL.inventory, crName: 'Inventory',
        narration: `Robbing ${record.robbing_no || ''} · ${record.part_taken || ''}`,
      });
    }
  }
  return legs;
}

async function insertLegs(legs) {
  for (const leg of legs) {
    await query(
      `INSERT INTO fin_transactions
        (id, date, partner_name, account_id, account_code, account_name, type, amount, debit, credit,
         narration, status, module, source, material_ref, payload)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16::jsonb)
       ON CONFLICT (id) DO NOTHING`,
      [
        leg.id,
        leg.date || null,
        leg.partner_name || null,
        leg.account_id || null,
        leg.account_code || null,
        leg.account_name || null,
        leg.type || null,
        leg.amount || 0,
        leg.debit || 0,
        leg.credit || 0,
        leg.narration || null,
        leg.status || 'Posted',
        leg.module || 'MATERIAL',
        leg.source || null,
        leg.material_ref || null,
        JSON.stringify(leg),
      ]
    );
    await query(
      `INSERT INTO fin_day_journals
        (id, date, account_id, account_code, type, amount, debit, credit, narration, module, source, payload)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12::jsonb)
       ON CONFLICT (id) DO NOTHING`,
      [
        leg.id,
        leg.date || null,
        leg.account_id || null,
        leg.account_code || null,
        leg.type || null,
        leg.amount || 0,
        leg.debit || 0,
        leg.credit || 0,
        leg.narration || null,
        leg.module || 'MATERIAL',
        leg.source || null,
        JSON.stringify(leg),
      ]
    );
  }
}

async function postMaterialToFinance(source, recordId) {
  const table = SOURCE_TABLE[source];
  if (!table) throw new Error(`Unknown material source: ${source}`);
  const r = await query(`SELECT * FROM ${table} WHERE id = $1`, [recordId]);
  const record = r.rows[0];
  if (!record) throw new Error('Material record not found');
  if (record.posted_to_finance) throw new Error(`Already posted (${record.posted_ref || 'see ledger'})`);

  const legs = buildMaterialJournalLegs(source, record);
  if (!legs.length && source !== 'ROBBING') {
    throw new Error('No amount to post — check qty × cost / line value');
  }
  await insertLegs(legs);
  const postedAt = new Date().toISOString().slice(0, 10);
  const ref = legs[0]?.material_ref || `${source}-${recordId}`;
  await query(
    `UPDATE ${table}
     SET posted_to_finance = TRUE, posted_at = $2, posted_ref = $3, finance_status = 'Posted', updated_at = NOW()
     WHERE id = $1`,
    [recordId, postedAt, ref]
  );
  return { legs, ref, count: legs.length };
}

/** Sync HR employees → pay_master standing data */
async function syncHrToPayMaster() {
  const emp = await query(`SELECT * FROM hr_employees ORDER BY id_number`);
  let n = 0;
  for (const e of emp.rows) {
    const name = e.name;
    if (!name) continue;
    const existing = await query(`SELECT id FROM pay_master WHERE employee_name = $1 LIMIT 1`, [name]);
    if (existing.rows[0]) {
      await query(
        `UPDATE pay_master SET
          id_number = COALESCE($2, id_number),
          position = COALESCE($3, position),
          department = COALESCE($4, department),
          basic_salary = CASE WHEN basic_salary = 0 OR basic_salary IS NULL THEN $5 ELSE basic_salary END,
          tin_no = COALESCE($6, tin_no),
          pension_id = COALESCE($7, pension_id),
          from_hr = TRUE,
          updated_at = NOW()
         WHERE id = $1`,
        [
          existing.rows[0].id,
          e.id_number,
          e.job_title,
          e.department,
          e.gross_salary || 0,
          e.tin_no,
          e.pension_id,
        ]
      );
    } else {
      await query(
        `INSERT INTO pay_master
          (id, employee_name, id_number, position, department, basic_salary, tin_no, pension_id, from_hr, status)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,TRUE,'Active')`,
        [
          randomUUID(),
          name,
          e.id_number,
          e.job_title,
          e.department,
          e.gross_salary || 0,
          e.tin_no,
          e.pension_id,
        ]
      );
    }
    n += 1;
  }
  return { synced: n };
}

/** Post approved payroll process rows into ledger */
async function postPayrollProcess(ids = []) {
  let rows;
  if (ids.length) {
    rows = await query(
      `SELECT * FROM pay_process WHERE id = ANY($1::text[]) AND COALESCE(posted_to_ledger, FALSE) = FALSE`,
      [ids]
    );
  } else {
    rows = await query(
      `SELECT * FROM pay_process WHERE status IN ('Approved','Posted','Final') AND COALESCE(posted_to_ledger, FALSE) = FALSE`
    );
  }
  let posted = 0;
  for (const p of rows.rows) {
    const gross = money(p.gross_earnings);
    const tax = money(p.total_tax);
    const pen = money(p.total_pension);
    const net = money(p.total_net);
    if (gross <= 0) continue;
    const uid = `PAY-${p.id}`;
    const date = p.payroll_process_date || new Date().toISOString().slice(0, 10);
    const legs = [];
    // simplified: Dr salary expense, Cr tax/pension/net payable
    legs.push({
      id: `${uid}-DR`,
      date,
      account_id: '6110-01-10',
      account_code: '6110-01-10',
      account_name: 'Salary Expense',
      type: 'Debit',
      amount: gross,
      debit: gross,
      credit: 0,
      narration: `Payroll ${p.period} · ${p.employee_name}`,
      module: 'PAYROLL',
      source: 'PAY_PROCESS',
      partner_name: p.employee_name,
      status: 'Posted',
    });
    if (tax > 0) {
      legs.push({
        id: `${uid}-TAX`,
        date,
        account_id: '2270-01-10',
        account_code: '2270-01-10',
        account_name: 'Income Tax Payable',
        type: 'Credit',
        amount: tax,
        debit: 0,
        credit: tax,
        narration: `Payroll tax ${p.period} · ${p.employee_name}`,
        module: 'PAYROLL',
        source: 'PAY_PROCESS',
        status: 'Posted',
      });
    }
    if (pen > 0) {
      legs.push({
        id: `${uid}-PEN`,
        date,
        account_id: '2273-01-10',
        account_code: '2273-01-10',
        account_name: 'Pension Payable',
        type: 'Credit',
        amount: pen,
        debit: 0,
        credit: pen,
        narration: `Payroll pension ${p.period} · ${p.employee_name}`,
        module: 'PAYROLL',
        source: 'PAY_PROCESS',
        status: 'Posted',
      });
    }
    if (net > 0) {
      legs.push({
        id: `${uid}-NET`,
        date,
        account_id: '2240-01-10',
        account_code: '2240-01-10',
        account_name: 'Salary Payable',
        type: 'Credit',
        amount: net,
        debit: 0,
        credit: net,
        narration: `Net pay ${p.period} · ${p.employee_name}`,
        module: 'PAYROLL',
        source: 'PAY_PROCESS',
        status: 'Posted',
      });
    }
    await insertLegs(legs);
    await query(`UPDATE pay_process SET posted_to_ledger = TRUE, status = 'Posted', updated_at = NOW() WHERE id = $1`, [
      p.id,
    ]);
    posted += 1;
  }
  return { posted };
}

module.exports = {
  buildMaterialJournalLegs,
  postMaterialToFinance,
  syncHrToPayMaster,
  postPayrollProcess,
};
