/**
 * Unity SIS — cross-module data bridge (Finance ↔ HR ↔ Material)
 * No extra "module" UI — pure integration for existing apps.
 *
 * Data flows:
 *  1) HR employees  →  Finance PAY_MASTER / PAY_PROCESS employee select
 *  2) Material GRV/GIV/PO/Robbing  →  Finance transactions (dual legs)
 *  3) Finance COA  →  Material COA picklists / validation
 *  4) HR leave / exit  →  Finance payroll eligibility signals
 *  5) Material inventory low-stock  →  optional HR/ops signals
 *  6) Shared roles / session via SharedCredentials (already)
 *
 * Persistence: localStorage + CustomEvent so separately mounted apps sync.
 * Prefer SharedDataProvider when all three share one React tree.
 */

const BUS_KEY = 'uu_sis_integration_v1';
const EMP_KEY = 'uu_shared_hr_employees_v1';
const COA_KEY = 'uu_shared_finance_coa_v1';
const MAT_LEDGER_KEY = 'uu_shared_material_ledger_v1';
const EVT = 'uu-sis-sync';

const money = (v) => {
  const n = Number(v);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
};

const safeParse = (raw, fb) => {
  try {
    if (!raw) return fb;
    return JSON.parse(raw);
  } catch {
    return fb;
  }
};

const emit = (detail) => {
  try {
    window.dispatchEvent(new CustomEvent(EVT, { detail }));
  } catch (_) {}
};

/* ------------------------------------------------------------------ */
/* Bus core                                                             */
/* ------------------------------------------------------------------ */

export const UnityBus = {
  EVT,
  read() {
    if (typeof localStorage === 'undefined') return {};
    return safeParse(localStorage.getItem(BUS_KEY), {});
  },
  write(partial = {}) {
    if (typeof localStorage === 'undefined') return partial;
    const next = { ...this.read(), ...partial, updated_at: new Date().toISOString() };
    localStorage.setItem(BUS_KEY, JSON.stringify(next));
    emit(next);
    return next;
  },
  subscribe(handler) {
    if (typeof window === 'undefined') return () => {};
    const fn = (e) => handler(e.detail || this.read());
    const onStorage = (e) => {
      if (e.key === BUS_KEY || e.key === EMP_KEY || e.key === MAT_LEDGER_KEY || e.key === COA_KEY) {
        handler(this.read());
      }
    };
    window.addEventListener(EVT, fn);
    window.addEventListener('storage', onStorage);
    return () => {
      window.removeEventListener(EVT, fn);
      window.removeEventListener('storage', onStorage);
    };
  },
};

/* ------------------------------------------------------------------ */
/* 1) HR employees → Finance / Material                                 */
/* ------------------------------------------------------------------ */

export function normalizeHrEmployee(e) {
  if (!e || typeof e !== 'object') return null;
  const name = e.name || e.employee_name || e.full_name || '';
  if (!name) return null;
  return {
    ...e,
    name,
    employee_name: name,
    id_number: e.id_number || e.idNo || e.employee_id || '',
    job_title: e.job_title || e.position || e.current_position || '',
    position: e.job_title || e.position || e.current_position || '',
    department: e.department || e.dept || '',
    work_location: e.work_location || e.location || '',
    gross_salary: e.gross_salary ?? e.current_salary ?? e.basic_salary ?? e.salary ?? 0,
    basic_salary: e.gross_salary ?? e.current_salary ?? e.basic_salary ?? e.salary ?? 0,
    pension_id: e.pension_id || e.pension_ac || '',
    tin_no: e.tin_no || e.tin || '',
    national_id: e.national_id || '',
    telephone_number: e.telephone_number || e.mobile || '',
    seniority_date_gc: e.seniority_date_gc || e.seniority_date || '',
    employment_type: e.employment_type || '',
    employment_category: e.employment_category || '',
    status: e.status || 'Active',
  };
}

export function publishHrEmployees(list) {
  const arr = (Array.isArray(list) ? list : []).map(normalizeHrEmployee).filter(Boolean);
  try {
    localStorage.setItem(EMP_KEY, JSON.stringify(arr));
    UnityBus.write({ hr_employee_count: arr.length, hr_published_at: new Date().toISOString() });
    emit({ type: 'hr_employees', count: arr.length });
  } catch (e) {
    console.warn('[UnityBridge] publishHrEmployees', e);
  }
  return arr;
}

export function loadHrEmployees() {
  return safeParse(typeof localStorage !== 'undefined' ? localStorage.getItem(EMP_KEY) : null, [])
    .map(normalizeHrEmployee)
    .filter(Boolean);
}

/** Finance payroll: merge HR shared list into process employee options + salary master rows */
export function mapHrToPayMaster(emp, index = 0) {
  const n = normalizeHrEmployee(emp);
  if (!n) return null;
  return {
    id: n.id || `pay-master-${n.id_number || index}`,
    employee_name: n.name,
    id_no: n.id_number,
    id_number: n.id_number,
    position: n.job_title,
    department: n.department,
    tin_no: n.tin_no,
    pension_id: n.pension_id,
    basic_salary: money(n.basic_salary || n.gross_salary),
    gross_salary: money(n.gross_salary || n.basic_salary),
    work_location: n.work_location,
    employment_type: n.employment_type,
    status: n.status || 'Active',
    source: 'HR_EMPLOYEE_CRUD',
  };
}

export function buildPayMasterFromHr(employees = null) {
  const list = employees || loadHrEmployees();
  return list.map(mapHrToPayMaster).filter(Boolean);
}

/* ------------------------------------------------------------------ */
/* 2) Material → Finance journals                                       */
/* ------------------------------------------------------------------ */

export const MAT_FINANCE_DEFAULT_GL = {
  inventory: '1474-01-10',
  goodsInTransit: '1490-01-10',
  grni: '2115-01-10',
  tradeCreditors: '2110-01-10',
  cashBank: '1141-01-10',
  materialExpense: '6310-01-10',
  maintenanceExpense: '6330-01-10',
};

export const MAT_SOURCE_COLLECTION_MAP = {
  GRV: 'MAT_GRV_CRUD',
  GIV: 'MAT_GIV_CRUD',
  LOCAL_PO: 'MAT_PURCHASE_LOCAL_CRUD',
  FOREIGN_PO: 'MAT_PURCHASE_FOREIGN_CRUD',
  ROBBING: 'MAT_ROBBING_CRUD',
};

const resolveMatGL = (raw, fallbackKey = 'inventory') => {
  const s = String(raw || '').trim();
  if (s) {
    const m = s.match(/(\d{4}-\d{2}-\d{2})/);
    if (m) return m[1];
    if (/^\d{4}/.test(s)) return s.split(/\s+/)[0];
  }
  return MAT_FINANCE_DEFAULT_GL[fallbackKey] || MAT_FINANCE_DEFAULT_GL.inventory;
};

/**
 * Build balanced dual-leg journal lines for a material event.
 * @param {'GRV'|'GIV'|'LOCAL_PO'|'FOREIGN_PO'|'ROBBING'} source
 */
export function buildMaterialJournalLegs(source, record) {
  if (!record) return [];
  const date = String(record.date || record.arrival_date || new Date().toISOString().slice(0, 10));
  const partner = record.supplier || record.department || record.preferred_supplier || 'Unity Stores';
  const descBase = record.item_name || record.name || record.part_taken || 'Material movement';
  const uid = String(record.id || Date.now());
  const legs = [];

  const pushPair = (ref, amount, drCode, drName, crCode, crName, narration) => {
    const amt = money(amount);
    if (amt <= 0) return;
    const common = {
      date,
      partner_name: partner,
      amount: amt,
      status: 'Posted',
      module: 'MATERIAL',
      source: `MAT_${source}`,
      material_ref: ref,
      material_record_id: uid,
      description: narration,
      narration,
    };
    legs.push({
      ...common,
      id: `mat-${source}-${uid}-L1`,
      account_id: drCode,
      account_code: drCode,
      account_name: drName,
      type: 'Debit',
      debit: amt,
      credit: 0,
      balancing_account: crCode,
    });
    legs.push({
      ...common,
      id: `mat-${source}-${uid}-L2`,
      account_id: crCode,
      account_code: crCode,
      account_name: crName,
      type: 'Credit',
      debit: 0,
      credit: amt,
      balancing_account: drCode,
    });
  };

  if (source === 'GRV') {
    const amt =
      money(record.line_value || record.lineValue || record.total_value) ||
      money(record.qty_received || record.quantity || record.qty) *
        money(record.unit_cost || record.unitCost);
    const inv = resolveMatGL(record.coa_account, 'inventory');
    const grni = resolveMatGL(null, 'grni');
    pushPair(
      record.grv_no || uid,
      amt,
      inv,
      `Inventory · ${inv}`,
      grni,
      'GRNI / Accrued receipts',
      `GRV ${record.grv_no || ''} · ${descBase}`
    );
  } else if (source === 'GIV') {
    const amt =
      money(record.line_value) || money(record.quantity) * money(record.unit_cost);
    const exp = resolveMatGL(record.coa_expense || record.coa_account, 'materialExpense');
    const inv = resolveMatGL(record.coa_account, 'inventory');
    pushPair(
      record.giv_no || uid,
      amt,
      exp,
      `Material expense · ${exp}`,
      inv,
      `Inventory · ${inv}`,
      `GIV ${record.giv_no || ''} · ${descBase} → ${record.department || 'Dept'}`
    );
  } else if (source === 'LOCAL_PO') {
    const amt = money(record.total_value);
    const inv = resolveMatGL(record.coa_account, 'inventory');
    const ap = resolveMatGL(null, 'tradeCreditors');
    pushPair(
      record.lpo_no || uid,
      amt,
      inv,
      `Inventory · ${inv}`,
      ap,
      'Trade Creditors',
      `Local PO ${record.lpo_no || ''} · ${descBase}`
    );
  } else if (source === 'FOREIGN_PO') {
    const amt =
      money(record.total_landed) ||
      money(record.fob_value) +
        money(record.freight_cost) +
        money(record.insurance_cost) +
        money(record.customs_duty) +
        money(record.shipping_cost);
    const git = resolveMatGL(record.coa_account, 'goodsInTransit');
    const ap = resolveMatGL(null, 'tradeCreditors');
    pushPair(
      record.import_ref || uid,
      amt,
      git,
      `Goods in transit · ${git}`,
      ap,
      'Trade Creditors',
      `Foreign import ${record.import_ref || ''} · ${descBase}`
    );
  } else if (source === 'ROBBING') {
    const amt = money(record.part_value || record.line_value || 0);
    if (amt > 0) {
      const exp = resolveMatGL(record.coa_note, 'maintenanceExpense');
      const inv = resolveMatGL(record.coa_note, 'inventory');
      pushPair(
        record.robbing_no || uid,
        amt,
        exp,
        `Maintenance · ${exp}`,
        inv,
        `Inventory · ${inv}`,
        `Robbing ${record.robbing_no || ''} · ${record.part_taken || 'part'}`
      );
    }
  }
  return legs;
}

/** Material posts legs → bus + local ledger; Finance merges into transactions */
export function postMaterialToFinance(source, record) {
  const legs = buildMaterialJournalLegs(source, record);
  if (!legs.length) return { ok: false, legs: [], message: 'No amount to post' };

  try {
    const prev = safeParse(localStorage.getItem(MAT_LEDGER_KEY), []);
    const next = [...legs, ...(Array.isArray(prev) ? prev : [])].slice(0, 5000);
    localStorage.setItem(MAT_LEDGER_KEY, JSON.stringify(next));
  } catch (_) {}

  UnityBus.write({
    pending_journals: [
      {
        id: `batch-${Date.now()}`,
        created_at: new Date().toISOString(),
        status: 'pending',
        module: 'material',
        source,
        ref: record.grv_no || record.giv_no || record.lpo_no || record.import_ref || record.robbing_no || record.id,
        recordId: record.id,
        legs,
      },
      ...(UnityBus.read().pending_journals || []),
    ].slice(0, 80),
  });

  emit({ type: 'material_post', source, legs: legs.length });
  return { ok: true, legs, message: `Posted ${legs.length} leg(s) to Finance bridge` };
}

export function loadMaterialLedger() {
  return safeParse(typeof localStorage !== 'undefined' ? localStorage.getItem(MAT_LEDGER_KEY) : null, []);
}

/** Finance DataProvider: merge pending material legs into transactions */
export function consumePendingMaterialJournals({ clear = true } = {}) {
  const bus = UnityBus.read();
  const batches = Array.isArray(bus.pending_journals) ? bus.pending_journals : [];
  const legs = batches.flatMap((b) => b.legs || []);
  const ledger = loadMaterialLedger();
  if (clear) {
    UnityBus.write({ pending_journals: [] });
  }
  // Deduplicate by id
  const byId = new Map();
  [...legs, ...ledger].forEach((L) => {
    if (L && L.id) byId.set(L.id, L);
  });
  return Array.from(byId.values());
}

/* ------------------------------------------------------------------ */
/* 3) Finance COA → Material                                            */
/* ------------------------------------------------------------------ */

export function publishFinanceCoa(coaList) {
  const list = (Array.isArray(coaList) ? coaList : []).map((a) => {
    const code = String(a.account_no || a.account_code || a.id || '').replace(/^acc-/i, '');
    return {
      account_no: code,
      name: a.name || a.desc || a.description || code,
      type: a.type || a.account_type || '',
      label: `${code} ${a.name || a.desc || ''}`.trim(),
    };
  }).filter((a) => /^\d{4}-\d{2}-\d{2}/.test(a.account_no));
  try {
    localStorage.setItem(COA_KEY, JSON.stringify(list));
    UnityBus.write({ coa_count: list.length });
  } catch (_) {}
  return list;
}

export function loadSharedCoa() {
  return safeParse(typeof localStorage !== 'undefined' ? localStorage.getItem(COA_KEY) : null, []);
}

/** Material UI options: prefer shared COA, else static fallbacks */
export function getMaterialCoaOptions(fallback = []) {
  const shared = loadSharedCoa();
  if (shared.length) return shared.map((a) => a.label || `${a.account_no} ${a.name}`);
  return fallback;
}

/* ------------------------------------------------------------------ */
/* 4) HR → Finance payroll journal helpers                              */
/* ------------------------------------------------------------------ */

export const PAY_GL = {
  salaries: '6110-01-10',
  overtime: '6140-01-10',
  allowances: '6220-01-10',
  companyPension: '6260-01-10',
  indemnity: '6291-01-10',
  netPayable: '2240-01-10',
  incomeTax: '2270-01-10',
  pensionPayable: '2273-01-10',
  cashBank: '1141-01-10',
};

/** Build payroll process → GL legs (gross, deductions, net) */
export function buildPayrollJournalLegs(processRow) {
  if (!processRow) return [];
  const date = String(processRow.payroll_process_date || processRow.date || new Date().toISOString().slice(0, 10));
  const name = processRow.employee_name || 'Employee';
  const uid = String(processRow.id || Date.now());
  const basic = money(processRow.basic_salary || processRow.total_basic);
  const ot = money(processRow.overtime || processRow.total_ot_pay);
  const allow = money(processRow.allowance || processRow.total_allowances);
  const gross = money(processRow.gross_amount) || basic + ot + allow;
  const tax = money(processRow.income_tax || processRow.total_tax);
  const pension = money(processRow.pension_contribution || processRow.total_pension);
  const net = money(processRow.total_net || processRow.net_pay) || Math.max(0, gross - tax - pension);
  const legs = [];
  const push = (code, nameAcc, debit, credit, narr) => {
    const d = money(debit);
    const c = money(credit);
    if (d <= 0 && c <= 0) return;
    legs.push({
      id: `pay-${uid}-${code}-${d > 0 ? 'D' : 'C'}`,
      date,
      account_code: code,
      account_id: code,
      account_name: nameAcc,
      type: d > 0 ? 'Debit' : 'Credit',
      debit: d,
      credit: c,
      amount: d || c,
      status: 'Posted',
      module: 'PAYROLL',
      source: 'PAY_PROCESS',
      partner_name: name,
      narration: narr,
      payroll_process_id: uid,
    });
  };
  push(PAY_GL.salaries, 'Salaries & Wages', gross, 0, `Payroll gross · ${name}`);
  push(PAY_GL.incomeTax, 'Income Tax Payable', 0, tax, `PAYE · ${name}`);
  push(PAY_GL.pensionPayable, 'Pension Payable', 0, pension, `Pension · ${name}`);
  push(PAY_GL.netPayable, 'Salary Payable', 0, net, `Net pay · ${name}`);
  return legs;
}

export function postPayrollBatchToFinance(processRows = []) {
  const legs = (processRows || []).flatMap(buildPayrollJournalLegs);
  if (!legs.length) return { ok: false, legs: [] };
  UnityBus.write({
    pending_journals: [
      {
        id: `pay-batch-${Date.now()}`,
        created_at: new Date().toISOString(),
        module: 'payroll',
        source: 'PAY_PROCESS',
        legs,
      },
      ...(UnityBus.read().pending_journals || []),
    ].slice(0, 80),
  });
  return { ok: true, legs };
}

/* ------------------------------------------------------------------ */
/* 5) Bootstrap merge helper for Finance DataProvider                   */
/* ------------------------------------------------------------------ */

/**
 * Call after API bootstrap in Finance:
 *   setAllData(prev => mergeFinanceIntegrations(prev))
 */
export function mergeFinanceIntegrations(allData = {}) {
  const next = { ...allData };
  // HR employees into payroll visibility
  const hr = loadHrEmployees();
  if (hr.length) {
    next.HR_EMPLOYEE_CRUD = hr;
    const masters = buildPayMasterFromHr(hr);
    if (!next.PAY_MASTER_CRUD || !next.PAY_MASTER_CRUD.length) {
      next.PAY_MASTER_CRUD = masters;
    } else {
      // fill missing names from HR
      const byName = new Map(masters.map((m) => [String(m.employee_name).toLowerCase(), m]));
      next.PAY_MASTER_CRUD = next.PAY_MASTER_CRUD.map((row) => {
        const hit = byName.get(String(row.employee_name || '').toLowerCase());
        return hit ? { ...hit, ...row, basic_salary: row.basic_salary || hit.basic_salary } : row;
      });
    }
  }
  // Material + payroll pending legs into transactions
  const matLegs = consumePendingMaterialJournals({ clear: false });
  if (matLegs.length) {
    const existing = Array.isArray(next.transactions) ? next.transactions : [];
    const ids = new Set(existing.map((t) => t.id));
    const add = matLegs.filter((L) => L.id && !ids.has(L.id));
    next.transactions = [...add, ...existing];
  }
  // Publish COA outward
  const coa = next.FIN_COA_CRUD || next.chartsOfAccounts || [];
  if (coa.length) publishFinanceCoa(coa);
  return next;
}

/**
 * Call after API bootstrap in HR:
 *   publishHrEmployees(allData.HR_EMPLOYEE_CRUD)
 */
export function afterHrBootstrap(allData = {}) {
  const emp = allData.HR_EMPLOYEE_CRUD || [];
  if (emp.length) publishHrEmployees(emp);
  return allData;
}

/**
 * Call after Material bootstrap — attach shared COA labels if useful
 */
export function afterMaterialBootstrap(allData = {}) {
  const coa = loadSharedCoa();
  if (coa.length) {
    allData._shared_coa_options = coa.map((a) => a.label || a.account_no);
  }
  return allData;
}

export default {
  UnityBus,
  publishHrEmployees,
  loadHrEmployees,
  buildPayMasterFromHr,
  buildMaterialJournalLegs,
  postMaterialToFinance,
  consumePendingMaterialJournals,
  publishFinanceCoa,
  loadSharedCoa,
  getMaterialCoaOptions,
  buildPayrollJournalLegs,
  postPayrollBatchToFinance,
  mergeFinanceIntegrations,
  afterHrBootstrap,
  afterMaterialBootstrap,
};
