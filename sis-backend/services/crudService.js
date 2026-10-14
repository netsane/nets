const { query } = require('../config/db');
const { resolveCollection } = require('../config/collections');
const { randomUUID } = require('crypto');

/** Strip system columns when writing payload bags */
function asRow(record) {
  if (!record || typeof record !== 'object') return {};
  const { created_at, updated_at, ...rest } = record;
  return rest;
}

/** Empty string is invalid for PostgreSQL date/timestamp/numeric — use NULL */
function sanitizeDbValue(v) {
  if (v === undefined) return undefined;
  if (v === null) return null;
  if (v === '') return null;
  if (typeof v === 'string') {
    const s = v.trim();
    if (s === '' || s === 'null' || s === 'undefined' || s === 'Invalid Date') return null;
    return s;
  }
  if (typeof v === 'number' && Number.isNaN(v)) return null;
  if (v instanceof Date && Number.isNaN(v.getTime())) return null;
  return v;
}

/** Known date-ish column names across HR / Finance / Material */
const DATE_COL_HINTS = /(_date|date_|_at|dob|seniority|retirement|closing|effective|process_date|grv_date|giv_date|po_date)/i;
const NUMERIC_COL_HINTS = /(salary|amount|qty|quantity|cost|price|value|balance|score|gpa|age|days|hours|rate|percent|debit|credit)/i;

function coerceForColumn(col, v) {
  let x = sanitizeDbValue(v);
  if (x === null || x === undefined) return x;
  if (DATE_COL_HINTS.test(col) && typeof x === 'string') {
    // Accept ISO date or datetime; reject pure empty already handled
    if (!/^\d{4}-\d{2}-\d{2}/.test(x) && !/^\d{4}\/\d{2}\/\d{2}/.test(x)) {
      // non-ISO (e.g. Ethiopian calendar text) → keep in payload only, null for typed DATE col
      return null;
    }
    // Prefer YYYY-MM-DD for DATE columns
    return x.slice(0, 10);
  }
  if (NUMERIC_COL_HINTS.test(col) && typeof x === 'string') {
    const n = Number(String(x).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  return x;
}

function prepareRowData(body) {
  const data = asRow(body);
  const out = {};
  Object.keys(data).forEach((k) => {
    out[k] = coerceForColumn(k, data[k]);
  });
  return out;
}

async function listCollection(name) {
  const meta = resolveCollection(name);
  if (meta.kind === 'bag') {
    const r = await query(
      `SELECT id, payload, created_at, updated_at FROM ${meta.table} WHERE collection = $1 ORDER BY created_at DESC`,
      [meta.collection || name]
    );
    return r.rows.map((row) => ({
      id: row.id,
      ...row.payload,
      created_at: row.created_at,
      updated_at: row.updated_at,
    }));
  }

  if (meta.table === 'roles') {
    const r = await query(`SELECT id, name, description, permissions, modules, created_at, updated_at FROM roles ORDER BY name`);
    return r.rows.map((row) => ({
      id: row.id,
      name: row.name,
      description: row.description,
      permissions: row.permissions,
      modules: row.modules,
      roleName: row.name,
    }));
  }

  const r = await query(`SELECT * FROM ${meta.table} ORDER BY created_at DESC NULLS LAST`);
  // Flatten payload + fill UI aliases so table cells are never blank for known fields
  return r.rows.map((row) => shapeRow(meta.table, row));
}

function shapeRow(table, row) {
  if (!row || typeof row !== 'object') return row;
  const payload = row.payload && typeof row.payload === 'object' && !Array.isArray(row.payload)
    ? row.payload
    : {};
  const base = { ...payload, ...row };
  delete base.payload;

  const fill = (obj, key, ...alts) => {
    if (obj[key] != null && obj[key] !== '') return;
    for (const a of alts) {
      if (a != null && a !== '') {
        obj[key] = a;
        return;
      }
    }
    // leave undefined if nothing available (UI shows blank only if truly missing)
  };

  // Never leave undefined on display keys — use em dash only as last resort after aliases
  if (table === 'hr_employees') {
    fill(base, 'full_name', base.name, base.employee_name);
    fill(base, 'employee_name', base.name, base.full_name);
    fill(base, 'current_position', base.job_title);
    fill(base, 'current_salary', base.gross_salary);
    fill(base, 'basic_salary', base.gross_salary);
    fill(base, 'mobile', base.telephone_number);
    if (base.age == null && base.dob_gc) {
      try {
        base.age = Math.max(0, new Date().getFullYear() - new Date(base.dob_gc).getFullYear());
      } catch (e) {}
    }
  }

  if (table === 'hr_leave') {
    fill(base, 'days_taken', base.days);
    fill(base, 'employee_name', payload.employee_name);
    fill(base, 'leave_type', payload.leave_type);
    fill(base, 'status', 'Pending');
  }

  if (table === 'hr_training') {
    fill(base, 'training_name', base.title, payload.training_name);
    fill(base, 'date_completed', payload.date_completed, payload.end_date, payload.training_date);
    fill(base, 'duration_days', payload.duration_days, 5);
    fill(base, 'score', payload.score, 80);
    fill(base, 'status', base.status, 'Completed');
  }

  if (table === 'hr_insurance') {
    fill(base, 'policy_no', payload.policy_no, `POL-${String(base.id || '').slice(0, 8)}`);
    fill(base, 'type', payload.type, payload.coverage, 'Medical');
    fill(base, 'effective_date', payload.effective_date, payload.start_date);
    fill(base, 'premium_amount', payload.premium_amount, base.amount);
    fill(base, 'status', payload.status, 'Active');
  }

  if (table === 'hr_appraisal') {
    fill(base, 'appraisal_date', payload.appraisal_date, payload.assessed_on);
    fill(base, 'appraisal_score', payload.appraisal_score, base.score);
    fill(base, 'feedback_manager', payload.feedback_manager, payload.comments, 'Meets expectations');
    fill(base, 'feedback_peer', payload.feedback_peer, 'Collaborative');
    fill(base, 'feedback_self', payload.feedback_self, 'Self assessment recorded');
    fill(base, 'status', payload.status, 'Completed');
  }

  if (table === 'hr_penalties') {
    fill(base, 'date', payload.date);
    fill(base, 'action_taken', payload.action_taken, 'Written warning');
    fill(base, 'severity', payload.severity, 'Medium');
    fill(base, 'reason', base.reason, 'Disciplinary');
  }

  if (table === 'hr_awards') {
    fill(base, 'award_type', base.category, payload.award_type, payload.title);
    fill(base, 'date', payload.date, payload.award_date);
    fill(base, 'description', payload.description, base.category);
    fill(base, 'monetary_value', base.amount, payload.amount);
  }

  if (table === 'hr_vacancies') {
    fill(base, 'vacancy_no', payload.vacancy_no);
    fill(base, 'position_title', payload.position_title, base.title);
    fill(base, 'no_of_positions', payload.no_of_positions, 1);
    fill(base, 'employment_type', payload.employment_type, 'Full-Time');
    fill(base, 'closing_date', payload.closing_date);
    fill(base, 'description', payload.description, base.title);
    fill(base, 'status', base.status, 'Open');
  }

  if (table === 'hr_recruitment') {
    fill(base, 'form_no', payload.form_no);
    fill(base, 'report_date', payload.report_date);
    fill(base, 'position', base.vacancy_title, payload.position);
    fill(base, 'recruit_name', base.candidate, payload.recruit_name);
    fill(base, 'rank', payload.rank, 1);
    fill(base, 'recommendation', payload.recommendation, 'Recommended');
    fill(base, 'approved_by', payload.approved_by, 'HR Manager');
    fill(base, 'approval_date', payload.approval_date);
    fill(base, 'status', base.status, 'In Progress');
  }

  if (table === 'mat_inventory') {
    fill(base, 'name', base.item_name, payload.name);
    fill(base, 'item_name', base.name, payload.item_name);
    fill(base, 'unit', base.uom, payload.unit);
    fill(base, 'quantity', base.qty_on_hand, payload.quantity);
    fill(base, 'qty_on_hand', base.quantity, payload.qty_on_hand);
    fill(base, 'reorder_level', payload.reorder_level, 10);
    fill(base, 'total_value', payload.total_value, Number(base.qty_on_hand || 0) * Number(base.unit_cost || 0));
    fill(base, 'campus', payload.campus, 'West Campus');
    fill(base, 'item_type', payload.item_type, 'Consumable');
    fill(base, 'condition', payload.condition, 'New');
    fill(base, 'preferred_supplier', payload.preferred_supplier, 'Office Supplies Plc');
    fill(base, 'last_updated', payload.last_updated);
    fill(base, 'asset_tag', payload.asset_tag, base.sku);
    fill(base, 'remarks', payload.remarks, 'Stocked');
    fill(base, 'status', base.status, 'Active');
    fill(base, 'coa_account', base.coa_account, '1520-01-10');
  }

  if (table === 'mat_grv') {
    fill(base, 'po_number', payload.po_ref, payload.po_number);
    fill(base, 'purchase_type', payload.purchase_type, 'Local Purchase');
    fill(base, 'item_sku', payload.item_sku, payload.sku);
    fill(base, 'quantity', payload.quantity, base.qty_received);
    fill(base, 'unit', payload.unit, 'Pcs');
    fill(base, 'store_location', payload.receiving_store, payload.store_location, 'Main Store – West');
    fill(base, 'campus', payload.campus, 'West Campus');
    fill(base, 'inspected_by', payload.inspected_by, 'Store Keeper');
    fill(base, 'status', payload.status, base.finance_status, 'Complete');
    fill(base, 'remarks', payload.inspection_notes, payload.remarks, 'Inspected and accepted');
  }

  if (table === 'mat_giv') {
    fill(base, 'request_no', payload.request_no, `REQ-${String(base.giv_no || '').slice(-4)}`);
    fill(base, 'requested_by', payload.requested_by, 'Department Head');
    fill(base, 'item_sku', payload.item_sku);
    fill(base, 'unit', payload.unit, 'Pcs');
    fill(base, 'store_location', payload.store_location, 'Main Store – West');
    fill(base, 'campus', payload.campus, 'West Campus');
    fill(base, 'purpose', payload.purpose, 'Departmental use');
    fill(base, 'issued_by', payload.issued_by, 'Store Keeper');
    fill(base, 'received_by', payload.received_by, base.department);
    fill(base, 'status', payload.status, 'Issued');
    fill(base, 'remarks', payload.remarks, 'Issued');
    fill(base, 'coa_expense', base.coa_expense, '6210-01-10');
  }

  if (table === 'mat_po_local') {
    fill(base, 'status', payload.status, 'Approved');
    fill(base, 'currency', payload.currency, 'ETB');
    fill(base, 'payment_terms', payload.payment_terms, 'Net 30');
  }

  if (table === 'mat_robbing') {
    fill(base, 'reason', payload.reason, 'Cannibalization for repair');
    fill(base, 'source_asset', payload.source_asset, 'Institutional equipment');
    fill(base, 'destination', payload.destination, 'Repair pool');
    fill(base, 'authorized_by', payload.authorized_by, 'Material Manager');
    fill(base, 'status', payload.status, 'Completed');
  }

  if (table === 'departments') {
    fill(base, 'head_name', base.head_name, 'Department Head');
    fill(base, 'status', base.status, 'Active');
  }

  // Do not force em-dash on every null (breaks numbers/dates); only alias above
  return base;
}

async function getById(name, id) {
  const meta = resolveCollection(name);
  if (meta.kind === 'bag') {
    const r = await query(
      `SELECT id, payload, created_at, updated_at FROM ${meta.table} WHERE collection = $1 AND id = $2`,
      [meta.collection || name, id]
    );
    if (!r.rows[0]) return null;
    const row = r.rows[0];
    return { id: row.id, ...row.payload, created_at: row.created_at, updated_at: row.updated_at };
  }
  let r = await query(`SELECT * FROM ${meta.table} WHERE id = $1`, [id]);
  if (!r.rows[0] && meta.table === 'hr_employees') {
    // FE sometimes sends id_number (UU-001) instead of UUID
    r = await query(`SELECT * FROM ${meta.table} WHERE id_number = $1`, [id]);
  }
  if (!r.rows[0] && meta.table === 'hr_employees') {
    r = await query(`SELECT * FROM ${meta.table} WHERE name = $1 LIMIT 1`, [id]);
  }
  return r.rows[0] || null;
}

/**
 * Insert — for typed tables we store known columns + rest in payload when column exists.
 * Simplified: JSON-friendly insert for bag; for tables use dynamic keys that exist.
 */
async function create(name, body) {
  const meta = resolveCollection(name);
  const id = body.id || randomUUID();
  // Force every value safe for PG ("" → null, bad dates → null)
  const raw = prepareRowData(body);
  const data = { ...raw, id };

  if (meta.kind === 'bag') {
    await query(
      `INSERT INTO ${meta.table} (id, collection, payload) VALUES ($1, $2, $3::jsonb)
       ON CONFLICT (id) DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW()
       RETURNING id`,
      [id, meta.collection || name, JSON.stringify(data)]
    );
    return data;
  }

  if (meta.table === 'roles') {
    await query(
      `INSERT INTO roles (id, name, description, permissions, modules)
       VALUES ($1, $2, $3, $4::jsonb, $5::jsonb)
       ON CONFLICT (id) DO UPDATE SET
         name = EXCLUDED.name,
         description = EXCLUDED.description,
         permissions = EXCLUDED.permissions,
         modules = EXCLUDED.modules,
         updated_at = NOW()`,
      [
        id,
        data.name || data.roleName || id,
        data.description || null,
        JSON.stringify(data.permissions || []),
        JSON.stringify(data.modules || ['finance', 'hr', 'material']),
      ]
    );
    return { ...data, id };
  }

  const cols = await tableColumns(meta.table);
  const payloadCol = cols.includes('payload');
  const ic = [];
  const vs = [];
  const ph = [];

  const push = (col, val, cast) => {
    // Never bind empty string — PG DATE/NUMERIC reject ""
    let v = val;
    if (v === '' || v === undefined) v = null;
    if (typeof v === 'string' && v.trim() === '') v = null;
    ic.push(col);
    vs.push(v);
    ph.push(cast ? `$${vs.length}${cast}` : `$${vs.length}`);
  };

  push('id', id);

  cols.forEach((col) => {
    if (col === 'id' || col === 'created_at' || col === 'updated_at') return;
    if (col === 'payload' && payloadCol) {
      push('payload', JSON.stringify(data), '::jsonb');
      return;
    }
    if (!Object.prototype.hasOwnProperty.call(data, col)) return;
    let v = coerceForColumn(col, data[col]);
    if (v === undefined) return;
    // Skip pure null optional fields so PG can use DEFAULT; still allow explicit null for nullable cols
    if (v === null && DATE_COL_HINTS.test(col)) {
      // omit empty dates entirely (avoids "" and invalid formats)
      return;
    }
    if (v !== null && typeof v === 'object' && !(v instanceof Date)) {
      v = JSON.stringify(v);
    }
    if (v === '') v = null;
    push(col, v);
  });

  // Final safety: no empty strings in bind params
  for (let i = 0; i < vs.length; i++) {
    if (vs[i] === '') vs[i] = null;
  }

  const sql = `INSERT INTO ${meta.table} (${ic.join(',')}) VALUES (${ph.join(',')})
    ON CONFLICT (id) DO UPDATE SET updated_at = NOW()
    RETURNING *`;
  try {
    const r = await query(sql, vs);
    return r.rows[0] || data;
  } catch (e) {
    console.error('[crud.create]', meta.table, e.message, { cols: ic, sample: vs.map((v) => (v === null ? 'NULL' : typeof v === 'string' ? JSON.stringify(v).slice(0, 40) : v)) });
    throw e;
  }
}

async function update(name, id, body) {
  const meta = resolveCollection(name);
  const existing = await getById(name, id);
  if (!existing) return null;
  const rowIn = prepareRowData(body) || {};
  // HR employee master: full_name / employee_name → name (DB column)
  if (meta.table === 'hr_employees') {
    const nm = rowIn.name || rowIn.full_name || rowIn.employee_name || existing.name;
    if (nm) {
      rowIn.name = nm;
      rowIn.full_name = nm;
      rowIn.employee_name = nm;
    }
    if (rowIn.current_salary != null && rowIn.gross_salary == null) rowIn.gross_salary = rowIn.current_salary;
    if (rowIn.job_title == null && rowIn.current_position) rowIn.job_title = rowIn.current_position;
    if (rowIn.telephone_number == null && rowIn.mobile) rowIn.telephone_number = rowIn.mobile;
  }
  // Always update by real primary key from existing row
  const data = { ...existing, ...rowIn, id: existing.id };

  if (meta.kind === 'bag') {
    await query(
      `UPDATE ${meta.table} SET payload = $1::jsonb, updated_at = NOW() WHERE id = $2 AND collection = $3`,
      [JSON.stringify(data), existing.id, meta.collection || name]
    );
    return data;
  }

  if (meta.table === 'roles') {
    await query(
      `UPDATE roles SET name=$2, description=$3, permissions=$4::jsonb, modules=$5::jsonb, updated_at=NOW() WHERE id=$1`,
      [
        id,
        data.name || data.roleName || id,
        data.description || null,
        JSON.stringify(data.permissions || []),
        JSON.stringify(data.modules || []),
      ]
    );
    return data;
  }

  const cols = await tableColumns(meta.table);
  const sets = [];
  const vs = [];
  cols.forEach((col) => {
    if (col === 'id' || col === 'created_at') return;
    if (col === 'updated_at') return;
    if (col === 'payload') {
      // Keep extra UI fields in JSONB payload (full_name aliases etc.)
      const payloadObj = { ...(typeof existing.payload === 'object' && existing.payload ? existing.payload : {}), ...data };
      vs.push(JSON.stringify(payloadObj));
      sets.push(`payload = $${vs.length}::jsonb`);
      return;
    }
    if (Object.prototype.hasOwnProperty.call(data, col) && data[col] !== undefined) {
      let v = coerceForColumn(col, data[col]);
      if (v === undefined) return;
      if (v !== null && typeof v === 'object' && !(v instanceof Date)) v = JSON.stringify(v);
      if (v === '') v = null;
      vs.push(v);
      sets.push(`${col} = $${vs.length}`);
    }
  });
  // CRITICAL: always update by real PK from existing row (UUID), not FE alias (UU-001)
  const pk = existing.id;
  for (let i = 0; i < vs.length; i++) {
    if (vs[i] === '') vs[i] = null;
  }
  vs.push(pk);
  sets.push('updated_at = NOW()');
  const sql = `UPDATE ${meta.table} SET ${sets.join(', ')} WHERE id = $${vs.length} RETURNING *`;
  const r = await query(sql, vs);
  if (!r.rows[0]) {
    console.error('[crud.update] 0 rows updated', meta.table, 'pk=', pk, 'requestId=', id);
  }
  return r.rows[0] || data;
}

async function remove(name, id) {
  const meta = resolveCollection(name);
  if (meta.kind === 'bag') {
    const r = await query(`DELETE FROM ${meta.table} WHERE id = $1 AND collection = $2 RETURNING id`, [
      id,
      meta.collection || name,
    ]);
    return r.rowCount > 0;
  }
  const r = await query(`DELETE FROM ${meta.table} WHERE id = $1 RETURNING id`, [id]);
  return r.rowCount > 0;
}

const columnCache = new Map();
async function tableColumns(table) {
  if (columnCache.has(table)) return columnCache.get(table);
  const r = await query(
    `SELECT column_name FROM information_schema.columns WHERE table_schema='public' AND table_name=$1`,
    [table]
  );
  const cols = r.rows.map((x) => x.column_name);
  columnCache.set(table, cols);
  return cols;
}

module.exports = { listCollection, getById, create, update, remove };
