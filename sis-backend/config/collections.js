/**
 * Maps frontend collection names → DB strategy.
 * kind: table | bag
 * bag uses hr_records / fin_records / mat_records with collection column.
 */
const COLLECTION_MAP = {
  // Auth
  ROLE_CRUD: { module: 'shared', table: 'roles', kind: 'table' },

  // Org
  DEPARTMENTS_CRUD: { module: 'shared', table: 'departments', kind: 'table' },

  // HR
  HR_EMPLOYEE_CRUD: { module: 'hr', table: 'hr_employees', kind: 'table' },
  HR_LEAVE_CRUD: { module: 'hr', table: 'hr_leave', kind: 'table' },
  HR_RECRUITMENT_CRUD: { module: 'hr', table: 'hr_recruitment', kind: 'table' },
  HR_VACANCY_CRUD: { module: 'hr', table: 'hr_vacancies', kind: 'table' },
  HR_TRAINING_CRUD: { module: 'hr', table: 'hr_training', kind: 'table' },
  HR_APPRAISAL_CRUD: { module: 'hr', table: 'hr_appraisal', kind: 'table' },
  HR_INSURANCE_CRUD: { module: 'hr', table: 'hr_insurance', kind: 'table' },
  HR_STAFF_ATTRITION_CRUD: { module: 'hr', table: 'hr_attrition', kind: 'table' },
  HR_AWARD_CRUD: { module: 'hr', table: 'hr_awards', kind: 'table' },
  HR_PENALTY_CRUD: { module: 'hr', table: 'hr_penalties', kind: 'table' },

  // HR bag collections (full UI modules)
  DEPARTMENTS_HIRING_REQUESTS_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'DEPARTMENTS_HIRING_REQUESTS_CRUD' },
  RETIREMENT_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'RETIREMENT_CRUD' },
  RESIGNATION_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'RESIGNATION_CRUD' },
  DEATH_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'DEATH_CRUD' },
  HR_COMPETENCE_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_COMPETENCE_CRUD' },
  HR_TRANSFER_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_TRANSFER_CRUD' },
  HR_REPORT_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_REPORT_CRUD' },
  HR_STAFF_ASSIGNMENT_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_STAFF_ASSIGNMENT_CRUD' },
  HR_STAFF_DEVELOPMENT_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_STAFF_DEVELOPMENT_CRUD' },
  HR_STAFF_OVERVIEW_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_STAFF_OVERVIEW_CRUD' },
  HR_AWARD_CRUD_HR: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_AWARD_CRUD_HR' },
  HR_ROLE_ASSIGNMENT_CRUD: { module: 'hr', table: 'hr_records', kind: 'bag', collection: 'HR_ROLE_ASSIGNMENT_CRUD' },



  // Finance
  FIN_COA_CRUD: { module: 'finance', table: 'fin_coa', kind: 'table' },
  chartsOfAccounts: { module: 'finance', table: 'fin_coa', kind: 'table' },
  FIN_PARTNERS_CRUD: { module: 'finance', table: 'fin_partners', kind: 'table' },
  transactions: { module: 'finance', table: 'fin_transactions', kind: 'table' },
  dayJournals: { module: 'finance', table: 'fin_day_journals', kind: 'table' },
  PAY_MASTER_CRUD: { module: 'finance', table: 'pay_master', kind: 'table' },
  PAY_PROCESS_CRUD: { module: 'finance', table: 'pay_process', kind: 'table' },
  PAY_OVERTIME_CRUD: { module: 'finance', table: 'pay_overtime', kind: 'table' },
  PAY_AD_CRUD: { module: 'finance', table: 'pay_ad', kind: 'table' },
  PAY_FACTOR_MASTER_CRUD: { module: 'finance', table: 'pay_factors', kind: 'table' },
  PAY_RECON_CRUD: { module: 'finance', table: 'fin_records', kind: 'bag', collection: 'PAY_RECON_CRUD' },

  // Material
  MAT_INVENTORY_CRUD: { module: 'material', table: 'mat_inventory', kind: 'table' },
  MAT_ITEM_CRUD: { module: 'material', table: 'mat_inventory', kind: 'table' },
  MAT_GRV_CRUD: { module: 'material', table: 'mat_grv', kind: 'table' },
  MAT_GIV_CRUD: { module: 'material', table: 'mat_giv', kind: 'table' },
  MAT_PURCHASE_LOCAL_CRUD: { module: 'material', table: 'mat_po_local', kind: 'table' },
  MAT_PURCHASE_FOREIGN_CRUD: { module: 'material', table: 'mat_po_foreign', kind: 'table' },
  MAT_ROBBING_CRUD: { module: 'material', table: 'mat_robbing', kind: 'table' },
};

const BAG_TABLE = {
  hr: 'hr_records',
  finance: 'fin_records',
  material: 'mat_records',
  shared: 'fin_records',
};

function resolveCollection(name) {
  if (COLLECTION_MAP[name]) return COLLECTION_MAP[name];
  // fallback bag by prefix
  if (name.startsWith('HR_') || name.startsWith('DEPART')) {
    return { module: 'hr', table: BAG_TABLE.hr, kind: 'bag', collection: name };
  }
  if (name.startsWith('MAT_')) {
    return { module: 'material', table: BAG_TABLE.material, kind: 'bag', collection: name };
  }
  if (name.startsWith('FIN_') || name.startsWith('PAY_')) {
    return { module: 'finance', table: BAG_TABLE.finance, kind: 'bag', collection: name };
  }
  return { module: 'finance', table: BAG_TABLE.finance, kind: 'bag', collection: name };
}

module.exports = { COLLECTION_MAP, BAG_TABLE, resolveCollection };
