/**
 * Unity SIS — SharedDataProvider
 * Single allData store for Finance + HR + Material.
 * Drop in: frontend/src/unity/SharedDataProvider.jsx
 *
 * Usage in App.jsx:
 *   import { SharedDataProvider, useData } from './unity/SharedDataProvider';
 *   <SharedDataProvider><AppShell /></SharedDataProvider>
 *
 * Module UIs should call useData() instead of their own DataProvider.
 */

import React, {
  createContext,
  useContext,
  useState,
  useCallback,
  useMemo,
  useEffect,
} from 'react';
import { UnityIntegrationBus } from './IntegrationBus';

/* ------------------------------------------------------------------ */
/* GL defaults (align with institutional COA)                         */
/* ------------------------------------------------------------------ */

export const HR_FINANCE_DEFAULT_GL = {
  salaryExpense: '6110-01-10',
  overtime: '6140-01-10',
  allowances: '6220-01-10',
  trainingExpense: '6250-01-10',
  awardExpense: '6280-01-10',
  penaltyIncome: '4280-01-10',
  insuranceExpense: '6270-01-10',
  salaryPayable: '2240-01-10',
  sundryPayable: '2241-01-10',
  cashBank: '1141-01-10',
  payrollFund: '1130-01-10',
  incomeTax: '2270-01-10',
  pensionPayable: '2273-01-10',
  companyPension: '6260-01-10',
};

export const MAT_FINANCE_DEFAULT_GL = {
  inventory: '1474-01-10',
  goodsInTransit: '1490-01-10',
  grni: '2115-01-10',
  tradeCreditors: '2110-01-10',
  cashBank: '1141-01-10',
  materialExpense: '6310-01-10',
  maintenanceExpense: '6330-01-10',
};

export const HR_SOURCE_COLLECTION_MAP = {
  PENALTY: 'HR_PENALTY_CRUD',
  AWARD: 'HR_AWARD_CRUD',
  AWARD_HR: 'HR_AWARD_CRUD_HR',
  TRAINING: 'HR_TRAINING_CRUD',
  INSURANCE: 'HR_INSURANCE_CRUD',
};

export const HR_COLLECTION_SOURCE_MAP = Object.fromEntries(
  Object.entries(HR_SOURCE_COLLECTION_MAP).map(([k, v]) => [v, k])
);

export const MAT_SOURCE_COLLECTION_MAP = {
  GRV: 'MAT_GRV_CRUD',
  GIV: 'MAT_GIV_CRUD',
  LOCAL_PO: 'MAT_PURCHASE_LOCAL_CRUD',
  FOREIGN_PO: 'MAT_PURCHASE_FOREIGN_CRUD',
  ROBBING: 'MAT_ROBBING_CRUD',
};

export const MAT_COLLECTION_SOURCE_MAP = Object.fromEntries(
  Object.entries(MAT_SOURCE_COLLECTION_MAP).map(([k, v]) => [v, k])
);

const money = (v) => {
  const n = Number(v);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
};

const resolveGL = (raw, fallback, map) => {
  const s = String(raw || '').trim();
  if (s) {
    const m = s.match(/(\d{4}-\d{2}-\d{2})/);
    if (m) return m[1];
    if (/^\d{4}/.test(s)) return s.split(/\s+/)[0];
  }
  return map[fallback] || fallback;
};

const pickAmount = (record, keys = []) => {
  for (const k of keys) {
    if (record?.[k] != null && record[k] !== '' && money(record[k]) > 0) return money(record[k]);
  }
  return 0;
};

/* ------------------------------------------------------------------ */
/* HR employee → Basic Salary Master                                  */
/* ------------------------------------------------------------------ */

export const mapHrEmployeeToSalaryMaster = (emp, index = 0) => {
  if (!emp) return null;
  const name = emp.name || emp.employee_name || emp.username || '';
  if (!name) return null;
  const basic = money(
    emp.basic_salary != null && emp.basic_salary !== ''
      ? emp.basic_salary
      : emp.gross_salary != null && emp.gross_salary !== ''
        ? emp.gross_salary
        : 0
  );
  const locMap = {
    Gerji: '12', Burayu: '13', Dessie: '15', Adama: '14',
    West: '10', East: '11', HQ: '10',
  };
  const wl = String(emp.work_location || emp.location || '');
  let location_code = '10';
  for (const [k, v] of Object.entries(locMap)) {
    if (wl.toLowerCase().includes(k.toLowerCase())) {
      location_code = v;
      break;
    }
  }
  return {
    id: emp.pay_master_id || `hr-sal-${emp.id || emp.id_number || index}`,
    s_no: String(index + 1),
    employee_name: name,
    id_number: emp.id_number || emp.employee_id || '',
    position: emp.job_title || emp.position || '',
    department: emp.department || emp.dept || '',
    cost_center: emp.cost_center || '01',
    location_code: emp.location_code || location_code,
    basic_salary: basic,
    doctoral_incentive: money(emp.doctoral_incentive),
    pos_allowance: money(emp.pos_allowance || emp.position_allowance),
    cash_indemnity: money(emp.cash_indemnity),
    tin_no: emp.tin_no || emp.tin || '',
    pension_id: emp.pension_id || '',
    saving_acc: emp.saving_acc || emp.bank_account || '',
    seniority_date: emp.seniority_date_gc || emp.seniority_date || '',
    effective_date:
      emp.salary_effective_date ||
      emp.effective_date ||
      new Date().toISOString().slice(0, 10),
    status: /inactive|terminat|resign/i.test(String(emp.status || ''))
      ? 'Inactive'
      : emp.pay_status || 'Active',
    source: 'HR_EMPLOYEE_CRUD',
    hr_employee_id: emp.id,
    integrated_from_hr: true,
  };
};

export const buildSalaryMasterFromHR = (all) => {
  const hr = Array.isArray(all.HR_EMPLOYEE_CRUD) ? all.HR_EMPLOYEE_CRUD : [];
  const local = Array.isArray(all.PAY_MASTER_CRUD) ? all.PAY_MASTER_CRUD : [];
  if (!hr.length) return local;

  const byId = new Map();
  const byName = new Map();
  local.forEach((m) => {
    if (m.id_number) byId.set(String(m.id_number).trim().toLowerCase(), m);
    if (m.employee_name) byName.set(String(m.employee_name).trim().toLowerCase(), m);
  });

  const merged = [];
  const seen = new Set();
  hr.forEach((emp, i) => {
    const base = mapHrEmployeeToSalaryMaster(emp, i);
    if (!base) return;
    const idKey = String(base.id_number || '').trim().toLowerCase();
    const nameKey = String(base.employee_name || '').trim().toLowerCase();
    const override = (idKey && byId.get(idKey)) || (nameKey && byName.get(nameKey));
    if (override) {
      merged.push({
        ...base,
        ...override,
        employee_name: base.employee_name || override.employee_name,
        id_number: base.id_number || override.id_number,
        basic_salary:
          money(base.basic_salary) > 0 ? base.basic_salary : money(override.basic_salary),
        department: base.department || override.department,
        position: base.position || override.position,
        source: 'HR+PAY_MASTER',
        integrated_from_hr: true,
        hr_employee_id: emp.id,
      });
    } else {
      merged.push(base);
    }
    if (idKey) seen.add(`id:${idKey}`);
    if (nameKey) seen.add(`name:${nameKey}`);
  });

  local.forEach((m) => {
    const idKey = String(m.id_number || '').trim().toLowerCase();
    const nameKey = String(m.employee_name || '').trim().toLowerCase();
    if ((idKey && seen.has(`id:${idKey}`)) || (nameKey && seen.has(`name:${nameKey}`))) return;
    merged.push({ ...m, source: m.source || 'PAY_MASTER_ONLY' });
  });
  return merged;
};

/* ------------------------------------------------------------------ */
/* Journal builders                                                   */
/* ------------------------------------------------------------------ */

const pushPair = (legs, { date, partner, ref, uid, module, source, amt, drCode, drName, crCode, crName, narration }) => {
  if (amt <= 0) return;
  const common = {
    date,
    partner_name: partner,
    amount: amt,
    status: 'Posted',
    module,
    source,
    description: narration,
    narration,
  };
  legs.push({
    ...common,
    id: `${module.toLowerCase()}-${source}-${uid}-L1`,
    account_id: drCode,
    account_code: drCode,
    account_name: drName,
    type: 'Debit',
    debit: amt,
    credit: 0,
    balancing_account: crCode,
    balancing_account_name: crName,
    balancing_account_type: 'Credit',
    balancing_amount: amt,
    material_ref: module === 'MATERIAL' ? ref : undefined,
    hr_ref: module === 'HR' ? ref : undefined,
    reference: ref,
  });
  legs.push({
    ...common,
    id: `${module.toLowerCase()}-${source}-${uid}-L2`,
    account_id: crCode,
    account_code: crCode,
    account_name: crName,
    type: 'Credit',
    debit: 0,
    credit: amt,
    balancing_account: drCode,
    balancing_account_name: drName,
    balancing_account_type: 'Debit',
    balancing_amount: amt,
    material_ref: module === 'MATERIAL' ? ref : undefined,
    hr_ref: module === 'HR' ? ref : undefined,
    reference: ref,
  });
};

export const buildHRJournalLegs = (source, record) => {
  if (!record) return [];
  const date = String(
    record.date || record.award_date || record.training_date || record.penalty_date ||
    new Date().toISOString().slice(0, 10)
  );
  const partner = record.employee_name || record.name || 'Unity Staff';
  const uid = String(record.id || Date.now());
  const legs = [];
  const GL = HR_FINANCE_DEFAULT_GL;
  const ref = record.penalty_no || record.award_no || record.training_no || record.policy_no || uid;

  if (source === 'PENALTY') {
    const amount = pickAmount(record, ['fine_amount', 'amount', 'penalty_amount']);
    pushPair(legs, {
      date, partner, ref, uid, module: 'HR', source: `HR_${source}`, amt: amount,
      drCode: GL.cashBank, drName: 'Cash / recovery', crCode: GL.penaltyIncome, crName: 'Penalty income',
      narration: `HR Penalty · ${partner}`,
    });
  } else if (source === 'AWARD' || source === 'AWARD_HR') {
    const amount = pickAmount(record, ['award_amount', 'amount', 'value', 'monetary_value']);
    pushPair(legs, {
      date, partner, ref, uid, module: 'HR', source: `HR_${source}`, amt: amount,
      drCode: GL.awardExpense, drName: 'Award expense', crCode: GL.salaryPayable, crName: 'Awards payable',
      narration: `HR Award · ${partner}`,
    });
  } else if (source === 'TRAINING') {
    const amount = pickAmount(record, ['cost', 'training_cost', 'amount', 'budget']);
    pushPair(legs, {
      date, partner, ref, uid, module: 'HR', source: `HR_${source}`, amt: amount,
      drCode: GL.trainingExpense, drName: 'Training expense', crCode: GL.sundryPayable, crName: 'Training payable',
      narration: `HR Training · ${partner}`,
    });
  } else if (source === 'INSURANCE') {
    const amount = pickAmount(record, ['premium', 'amount', 'contribution', 'premium_amount']);
    pushPair(legs, {
      date, partner, ref, uid, module: 'HR', source: `HR_${source}`, amt: amount,
      drCode: GL.insuranceExpense, drName: 'Insurance expense', crCode: GL.sundryPayable, crName: 'Insurance payable',
      narration: `HR Insurance · ${partner}`,
    });
  }
  return legs;
};

export const buildMaterialJournalLegs = (source, record) => {
  if (!record) return [];
  const date = String(record.date || record.arrival_date || new Date().toISOString().slice(0, 10));
  const partner = record.supplier || record.department || 'Unity Stores';
  const uid = String(record.id || Date.now());
  const legs = [];
  const GL = MAT_FINANCE_DEFAULT_GL;
  const desc = record.item_name || record.name || record.part_taken || 'Material';

  if (source === 'GRV') {
    const amt = money(record.line_value) || money(record.qty_received || record.quantity) * money(record.unit_cost);
    const inv = resolveGL(record.coa_account, 'inventory', GL);
    pushPair(legs, {
      date, partner, ref: record.grv_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: inv, drName: `Inventory · ${inv}`, crCode: GL.grni, crName: 'GRNI',
      narration: `GRV ${record.grv_no || ''} · ${desc}`,
    });
  } else if (source === 'GIV') {
    const amt = money(record.line_value) || money(record.quantity) * money(record.unit_cost);
    const exp = resolveGL(record.coa_expense || record.coa_account, 'materialExpense', GL);
    const inv = resolveGL(record.coa_account, 'inventory', GL);
    pushPair(legs, {
      date, partner, ref: record.giv_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: exp, drName: `Expense · ${exp}`, crCode: inv, crName: `Inventory · ${inv}`,
      narration: `GIV ${record.giv_no || ''} · ${desc} → ${record.department || ''}`,
    });
  } else if (source === 'LOCAL_PO') {
    const amt = money(record.total_value);
    const inv = resolveGL(record.coa_account, 'inventory', GL);
    pushPair(legs, {
      date, partner, ref: record.lpo_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: inv, drName: `Inventory · ${inv}`, crCode: GL.tradeCreditors, crName: 'Trade Creditors',
      narration: `Local PO ${record.lpo_no || ''} · ${desc}`,
    });
  } else if (source === 'FOREIGN_PO') {
    const amt =
      money(record.total_landed) ||
      money(record.fob_value) + money(record.freight_cost) + money(record.insurance_cost) +
      money(record.customs_duty) + money(record.shipping_cost);
    const git = resolveGL(record.coa_account, 'goodsInTransit', GL);
    pushPair(legs, {
      date, partner, ref: record.import_ref || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
      drCode: git, drName: `GIT · ${git}`, crCode: GL.tradeCreditors, crName: 'Trade Creditors',
      narration: `Foreign import ${record.import_ref || ''} · ${desc}`,
    });
  } else if (source === 'ROBBING') {
    const amt = money(record.part_value || record.line_value);
    if (amt > 0) {
      pushPair(legs, {
        date, partner, ref: record.robbing_no || uid, uid, module: 'MATERIAL', source: `MAT_${source}`, amt,
        drCode: GL.maintenanceExpense, drName: 'Maintenance', crCode: GL.inventory, crName: 'Inventory',
        narration: `Robbing ${record.robbing_no || ''} · ${record.part_taken || ''}`,
      });
    }
  }
  return legs;
};

/* ------------------------------------------------------------------ */
/* Initial shell collections                                          */
/* ------------------------------------------------------------------ */

const EMPTY_COLLECTIONS = {
  // Finance
  transactions: [],
  dayJournals: [],
  FIN_COA_CRUD: [],
  chartsOfAccounts: [],
  FIN_PARTNERS_CRUD: [],
  PAY_MASTER_CRUD: [],
  PAY_PROCESS_CRUD: [],
  PAY_OVERTIME_CRUD: [],
  PAY_AD_CRUD: [],
  PAY_RECON_CRUD: [],
  // HR
  HR_EMPLOYEE_CRUD: [],
  HR_LEAVE_CRUD: [],
  HR_PENALTY_CRUD: [],
  HR_AWARD_CRUD: [],
  HR_AWARD_CRUD_HR: [],
  HR_TRAINING_CRUD: [],
  HR_INSURANCE_CRUD: [],
  HR_TRANSFER_CRUD: [],
  HR_VACANCY_CRUD: [],
  HR_RECRUITMENT_CRUD: [],
  DEPARTMENTS_CRUD: [],
  DEPARTMENTS_HIRING_REQUESTS_CRUD: [],
  RETIREMENT_CRUD: [],
  RESIGNATION_CRUD: [],
  DEATH_CRUD: [],
  ROLE_CRUD: [],
  HR_MATERIAL_SIGNALS: [],
  // Material
  MAT_ITEM_CRUD: [],
  MAT_GRV_CRUD: [],
  MAT_GIV_CRUD: [],
  MAT_ROBBING_CRUD: [],
  MAT_PURCHASE_LOCAL_CRUD: [],
  MAT_PURCHASE_FOREIGN_CRUD: [],
};

const DataContext = createContext(null);

export const useData = () => {
  const ctx = useContext(DataContext);
  if (!ctx) {
    throw new Error('useData() must be used within SharedDataProvider');
  }
  return ctx;
};

/**
 * @param {{ children: React.ReactNode, initialData?: object, persistKey?: string }} props
 * initialData — merge your MOCK / seed objects from any module
 * persistKey — if set, persist allData to localStorage
 */
export function SharedDataProvider({ children, initialData = {}, persistKey = 'uu_sis_all_data_v1' }) {
  const [allData, setAllData] = useState(() => {
    let base = { ...EMPTY_COLLECTIONS, ...initialData };
    if (persistKey && typeof localStorage !== 'undefined') {
      try {
        const raw = localStorage.getItem(persistKey);
        if (raw) base = { ...base, ...JSON.parse(raw) };
      } catch (_) {}
    }
    return base;
  });
  const [message, setMessage] = useState(null);
  const [notificationTrigger, setNotificationTrigger] = useState(0);

  useEffect(() => {
    if (!persistKey || typeof localStorage === 'undefined') return;
    try {
      localStorage.setItem(persistKey, JSON.stringify(allData));
    } catch (_) {}
  }, [allData, persistKey]);

  // Import pending journals from IntegrationBus (other windows / late posts)
  useEffect(() => {
    const importPending = () => {
      if (importPending._busy) return;
      importPending._busy = true;
      try {
        const batches = UnityIntegrationBus.consumePendingJournals({ clear: true });
        if (!batches.length) return;
        const legs = batches.flatMap((b) => (Array.isArray(b.legs) ? b.legs : [])).filter(Boolean);
        if (!legs.length) return;
        setAllData((prev) => {
          const existing = Array.isArray(prev.transactions) ? prev.transactions : [];
          const ids = new Set(existing.map((x) => x && x.id).filter(Boolean));
          const add = legs.filter((L) => L && L.id && !ids.has(L.id));
          if (!add.length) return prev;
          return { ...prev, transactions: [...add, ...existing] };
        });
      } finally {
        importPending._busy = false;
      }
    };
    importPending();
    return UnityIntegrationBus.subscribe(() => importPending());
  }, []);

  const getCollection = useCallback(
    (collectionName) => {
      if (collectionName === 'PAY_MASTER_CRUD') {
        return buildSalaryMasterFromHR(allData);
      }
      if (collectionName === 'chartsOfAccounts' || collectionName === 'FIN_COA_CRUD') {
        return allData.FIN_COA_CRUD?.length
          ? allData.FIN_COA_CRUD
          : allData.chartsOfAccounts || [];
      }
      return allData[collectionName] || [];
    },
    [allData]
  );

  const addRecord = useCallback(async (collectionName, newData) => {
    const id = newData.id || `${collectionName}-${Date.now()}`;
    const record = { ...newData, id };
    setAllData((prev) => ({
      ...prev,
      [collectionName]: [record, ...(prev[collectionName] || [])],
    }));
    setNotificationTrigger((n) => n + 1);
    return true;
  }, []);

  const updateRecord = useCallback(async (collectionName, id, updatedFields) => {
    setAllData((prev) => ({
      ...prev,
      [collectionName]: (prev[collectionName] || []).map((r) =>
        String(r.id) === String(id) ? { ...r, ...updatedFields } : r
      ),
    }));
    setNotificationTrigger((n) => n + 1);
    return true;
  }, []);

  const deleteRecord = useCallback(async (collectionName, id) => {
    setAllData((prev) => ({
      ...prev,
      [collectionName]: (prev[collectionName] || []).filter((r) => String(r.id) !== String(id)),
    }));
    setNotificationTrigger((n) => n + 1);
    return true;
  }, []);

  const setCollection = useCallback((collectionName, rows) => {
    setAllData((prev) => ({ ...prev, [collectionName]: rows }));
  }, []);

  const appendJournalLegs = useCallback((legs, meta = {}) => {
    if (!legs?.length) return;
    setAllData((prev) => ({
      ...prev,
      transactions: [...legs, ...(prev.transactions || [])],
      dayJournals: [...legs, ...(prev.dayJournals || [])],
    }));
    UnityIntegrationBus.pushJournalBatch({
      module: meta.module || 'SHARED',
      source: meta.source || 'append',
      legs,
      ref: meta.ref,
    });
    setNotificationTrigger((n) => n + 1);
  }, []);

  /** Payroll process row → finance (simplified institutional outline) */
  const postPayrollToFinance = useCallback((record) => {
    if (!record) return false;
    if (String(record.status || '').trim() !== 'Committed') {
      setMessage({ type: 'error', text: 'Payroll row must be Committed before post.' });
      return false;
    }
    const GL = HR_FINANCE_DEFAULT_GL;
    const date = new Date().toISOString().slice(0, 10);
    const partner = record.employee_name || 'Staff';
    const uid = String(record.id || Date.now());
    const ref = `PAYROLL-${record.period || record.month || date}-${uid}`;
    const basic = money(record.total_basic ?? record.basic_salary);
    const ot = money(record.total_ot_pay ?? record.overtime_pay ?? record.overtime);
    const allow = money(record.extra_allowance ?? record.allowance ?? record.total_allowances);
    const net = money(record.net_pay);
    const tax = money(record.income_tax);
    const penEE = money(record.pension_7 ?? record.employee_pension);
    const penER = money(record.company_pension ?? record.employer_pension);
    const legs = [];
    const push = (amt, dr, drN, cr, crN, narr) =>
      pushPair(legs, {
        date, partner, ref, uid, module: 'PAYROLL', source: 'PAY_PROCESS', amt,
        drCode: dr, drName: drN, crCode: cr, crName: crN, narration: narr,
      });
    if (basic > 0) push(basic, GL.salaryExpense, 'Salaries', GL.salaryPayable, 'Salary payable', `${ref} basic`);
    if (ot > 0) push(ot, GL.overtime, 'Overtime', GL.salaryPayable, 'Salary payable', `${ref} OT`);
    if (allow > 0) push(allow, GL.allowances, 'Allowances', GL.salaryPayable, 'Salary payable', `${ref} allow`);
    if (tax > 0) push(tax, GL.salaryPayable, 'Salary payable', GL.incomeTax, 'Tax payable', `${ref} tax`);
    if (penEE + penER > 0) {
      push(penER, GL.companyPension, 'Company pension', GL.pensionPayable, 'Pension payable', `${ref} ER pen`);
      push(penEE, GL.salaryPayable, 'Salary payable', GL.pensionPayable, 'Pension payable', `${ref} EE pen`);
    }
    if (!legs.length) {
      setMessage({ type: 'error', text: 'No payroll amounts to post.' });
      return false;
    }
    appendJournalLegs(legs, { module: 'PAYROLL', source: 'PAY_PROCESS', ref });
    setAllData((prev) => ({
      ...prev,
      PAY_PROCESS_CRUD: (prev.PAY_PROCESS_CRUD || []).map((r) =>
        String(r.id) === uid
          ? { ...r, posted_to_finance: true, posted_at: date, posted_ref: ref, status: 'Posted' }
          : r
      ),
    }));
    setMessage({ type: 'success', text: `Payroll posted: ${ref}` });
    return true;
  }, [appendJournalLegs]);

  const postHRToFinance = useCallback(
    (source, recordId) => {
      const collection = HR_SOURCE_COLLECTION_MAP[source];
      if (!collection) {
        setMessage({ type: 'error', text: `Unknown HR source ${source}` });
        return false;
      }
      const record = (allData[collection] || []).find((r) => String(r.id) === String(recordId));
      if (!record) {
        setMessage({ type: 'error', text: 'HR record not found' });
        return false;
      }
      if (record.posted_to_finance) {
        setMessage({ type: 'error', text: 'Already posted to finance' });
        return false;
      }
      const legs = buildHRJournalLegs(source, record);
      if (!legs.length) {
        setMessage({ type: 'error', text: 'No amount on HR record to post' });
        return false;
      }
      const ref = legs[0].reference;
      const postedAt = new Date().toISOString().slice(0, 10);
      appendJournalLegs(legs, { module: 'HR', source, ref });
      setAllData((prev) => ({
        ...prev,
        [collection]: (prev[collection] || []).map((r) =>
          String(r.id) === String(recordId)
            ? { ...r, posted_to_finance: true, posted_at: postedAt, posted_ref: ref, finance_status: 'Posted' }
            : r
        ),
      }));
      setMessage({ type: 'success', text: `HR ${source} posted · ${ref}` });
      return true;
    },
    [allData, appendJournalLegs]
  );

  const postMaterialToFinance = useCallback(
    (source, recordId) => {
      const collection = MAT_SOURCE_COLLECTION_MAP[source];
      if (!collection) {
        setMessage({ type: 'error', text: `Unknown material source ${source}` });
        return false;
      }
      const record = (allData[collection] || []).find((r) => String(r.id) === String(recordId));
      if (!record) {
        setMessage({ type: 'error', text: 'Material record not found' });
        return false;
      }
      if (record.posted_to_finance) {
        setMessage({ type: 'error', text: 'Already posted to finance' });
        return false;
      }
      const legs = buildMaterialJournalLegs(source, record);
      if (!legs.length) {
        setMessage({ type: 'error', text: 'No amount to post (check qty × cost)' });
        return false;
      }
      const ref = legs[0].reference;
      const postedAt = new Date().toISOString().slice(0, 10);
      appendJournalLegs(legs, { module: 'MATERIAL', source, ref });
      setAllData((prev) => ({
        ...prev,
        [collection]: (prev[collection] || []).map((r) =>
          String(r.id) === String(recordId)
            ? { ...r, posted_to_finance: true, posted_at: postedAt, posted_ref: ref, finance_status: 'Posted' }
            : r
        ),
      }));
      setMessage({ type: 'success', text: `Material ${source} posted · ${ref}` });
      return true;
    },
    [allData, appendJournalLegs]
  );

  const postAllHRToFinance = useCallback(
    (collectionName) => {
      const source = HR_COLLECTION_SOURCE_MAP[collectionName];
      if (!source) return 0;
      let n = 0;
      (allData[collectionName] || [])
        .filter((r) => !r.posted_to_finance)
        .forEach((r) => {
          if (postHRToFinance(source, r.id)) n += 1;
        });
      return n;
    },
    [allData, postHRToFinance]
  );

  const postAllMaterialToFinance = useCallback(
    (collectionName) => {
      const source = MAT_COLLECTION_SOURCE_MAP[collectionName];
      if (!source) return 0;
      let n = 0;
      (allData[collectionName] || [])
        .filter((r) => !r.posted_to_finance)
        .forEach((r) => {
          if (postMaterialToFinance(source, r.id)) n += 1;
        });
      return n;
    },
    [allData, postMaterialToFinance]
  );

  const syncBasicSalaryMasterFromHR = useCallback(() => {
    setAllData((prev) => ({
      ...prev,
      PAY_MASTER_CRUD: buildSalaryMasterFromHR(prev),
      _pay_master_synced_from_hr_at: new Date().toISOString(),
    }));
    if (typeof window !== 'undefined') {
      window.__uuSyncSalaryFromHR = () =>
        setAllData((prev) => ({
          ...prev,
          PAY_MASTER_CRUD: buildSalaryMasterFromHR(prev),
        }));
    }
    setMessage({ type: 'success', text: 'Basic Salary Master refreshed from HR.' });
    return true;
  }, []);

  useEffect(() => {
    if (typeof window !== 'undefined') {
      window.__uuSyncSalaryFromHR = syncBasicSalaryMasterFromHR;
    }
  }, [syncBasicSalaryMasterFromHR]);

  const emitMaterialSignal = useCallback((collectionName, record) => {
    const signal = {
      module: 'HR',
      collection: collectionName,
      record_id: record?.id,
      employee_name: record?.employee_name || record?.name,
      signal:
        collectionName === 'HR_TRANSFER_CRUD'
          ? 'STAFF_LOCATION_CHANGE'
          : ['RESIGNATION_CRUD', 'RETIREMENT_CRUD', 'DEATH_CRUD'].includes(collectionName)
            ? 'STAFF_EXIT'
            : 'STAFF_MASTER_UPSERT',
      from_location: record?.fromPlace || record?.from_location,
      to_location: record?.toPlace || record?.to_location,
      work_location: record?.work_location,
      department: record?.department,
    };
    UnityIntegrationBus.pushMaterialSignal(signal);
    setAllData((prev) => ({
      ...prev,
      HR_MATERIAL_SIGNALS: [signal, ...(prev.HR_MATERIAL_SIGNALS || [])],
    }));
    setMessage({ type: 'success', text: `Material signal: ${signal.signal}` });
    return true;
  }, []);

  const value = useMemo(
    () => ({
      allData,
      setAllData,
      getCollection,
      addRecord,
      updateRecord,
      deleteRecord,
      setCollection,
      message,
      setMessage,
      notificationTrigger,
      setNotificationTrigger,
      postPayrollToFinance,
      postHRToFinance,
      postMaterialToFinance,
      postAllHRToFinance,
      postAllMaterialToFinance,
      syncBasicSalaryMasterFromHR,
      emitMaterialSignal,
      appendJournalLegs,
      buildSalaryMasterFromHR,
      buildHRJournalLegs,
      buildMaterialJournalLegs,
    }),
    [
      allData,
      getCollection,
      addRecord,
      updateRecord,
      deleteRecord,
      setCollection,
      message,
      notificationTrigger,
      postPayrollToFinance,
      postHRToFinance,
      postMaterialToFinance,
      postAllHRToFinance,
      postAllMaterialToFinance,
      syncBasicSalaryMasterFromHR,
      emitMaterialSignal,
      appendJournalLegs,
    ]
  );

  return <DataContext.Provider value={value}>{children}</DataContext.Provider>;
}

export default SharedDataProvider;
// Named: SharedDataProvider already via `export function SharedDataProvider`
// useData, build* already exported above — do not re-export SharedDataProvider
