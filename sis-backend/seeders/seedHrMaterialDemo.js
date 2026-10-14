/**
 * Unity SIS — HR + Material typical demo seed (English names, ETB, full cells).
 * Fills typed tables + hr_records bags so UI columns are not empty.
 *
 * Run:  node seeders/seedHrMaterialDemo.js
 *   or: npm run db:seed:hrmat
 */
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const { pool } = require('../config/db');
const { randomUUID } = require('crypto');

const uid = () => randomUUID();

const DEPTS = [
  'Finance', 'Engineering', 'Human Resource', 'President Office', 'Registrar',
  'IT Department', 'Material Management', 'Academic Affairs', 'Marketing', 'Operations',
];
const LOCS = ['Torhailoch', 'Gerji', 'Burayu', 'Dessie', 'Adama', 'Bole', 'Addis Ababa'];
const FIRST = [
  'Netsanet', 'Damte', 'Abebe', 'Sara', 'Kalab', 'Hanna', 'Tadesse', 'Lomi',
  'Bereket', 'Aster', 'Chala', 'Tigist', 'Fasil', 'Eleni', 'Jemal', 'Saba',
  'Henok', 'Mekdes', 'Yonas', 'Rahel',
];
const LAST = [
  'Abebe', 'Tesfaye', 'Kebede', 'Hailu', 'Mekonnen', 'Assefa', 'Girma', 'Lemma',
  'Tadesse', 'Wolde', 'Bekele', 'Getachew', 'Alemu', 'Desta', 'Negash',
];

function empName(i) {
  return FIRST[i % FIRST.length] + ' ' + LAST[i % LAST.length];
}

function isoDate(y, m, d) {
  return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

async function clearHrMat() {
  const tables = [
    'mat_robbing', 'mat_po_local', 'mat_giv', 'mat_grv', 'mat_inventory',
    'hr_leave', 'hr_vacancies', 'hr_recruitment', 'hr_training', 'hr_appraisal',
    'hr_insurance', 'hr_awards', 'hr_penalties', 'hr_attrition', 'hr_records',
    'hr_employees', 'departments',
  ];
  for (const t of tables) {
    try {
      await pool.query('DELETE FROM ' + t);
    } catch (e) {
      console.warn('skip clear', t, e.message);
    }
  }
}

async function seedDepartments() {
  for (let i = 0; i < DEPTS.length; i++) {
    const name = DEPTS[i];
    const code = 'D' + String(i + 1).padStart(2, '0');
    await pool.query(
      'INSERT INTO departments (id, code, name, head_name, status, payload) VALUES ($1,$2,$3,$4,\'Active\',$5::jsonb)',
      [uid(), code, name, empName(i), JSON.stringify({ campus: LOCS[i % LOCS.length], cost_center: 'CC-' + (10 + i) })]
    );
  }
}

async function seedEmployees() {
  const rows = [];
  for (let i = 1; i <= 20; i++) {
    const name = empName(i - 1);
    const academic = i % 2 === 0;
    const job = academic ? 'Lecturer ' + ((i % 3) + 1) : 'Officer ' + ((i % 3) + 1);
    const category = academic ? 'Academic' : 'Non-Academic';
    let empType = i % 4 === 0 ? 'Non-Permanent' : 'Permanent';
    let dobY = 1975 + (i % 15);
    let dobM = (i % 12) + 1;
    let dobD = 5 + (i % 20);
    if (i <= 2) { dobY = 1966; dobM = 10 + (i % 2); dobD = 15; }
    if (i === 3 || i === 4) empType = 'Non-Permanent';
    const seniorityY = (i === 3 || i === 4) ? 2026 : 2015 + (i % 10);
    const seniorityM = (i === 3 || i === 4) ? 3 + (i % 3) : (i % 12) + 1;
    const salary = 12000 + i * 850;
    const idNumber = 'UU-' + String(i).padStart(3, '0');
    const role = i === 8 ? 'hr_manager' : i === 9 ? 'hr_staff' : 'employee';
    const payload = {
      full_name: name,
      employee_name: name,
      current_position: job,
      current_salary: salary,
      basic_salary: salary,
      mobile: '+2519' + String(10000000 + i * 137).slice(0, 8),
      education_entries: [{ title: academic ? 'Masters Degree' : 'Bachelors Degree', institution: 'Unity University', year: String(2005 + (i % 15)) }],
      beneficiaries: [{ name: 'Family of ' + name, relationship: 'Spouse', percentage: 100 }],
    };

    await pool.query(
      `INSERT INTO hr_employees (
        id, id_number, name, gender, dob_gc, dob_ec, age, marital_status,
        seniority_date_gc, seniority_date_ec, work_location, department, job_title,
        reporting_to, employment_category, employment_type, educational_status,
        cumulative_gpa, auth_status, field_of_study, gross_salary, family_contact,
        pension_id, retirement_date_ec, national_id, tin_no, telephone_number,
        office_tel, ext_no, emergency_contact_name, emergency_contact_telephone,
        emergency_contact_address, remarks, system_role, status, payload
      ) VALUES (
        $1,$2,$3,$4,$5::date,$6,$7,$8,
        $9::date,$10,$11,$12,$13,
        $14,$15,$16,$17,
        $18,$19,$20,$21,$22,
        $23,$24,$25,$26,$27,
        $28,$29,$30,$31,
        $32,$33,$34,$35,$36::jsonb
      )`,
      [
        uid(), idNumber, name, i % 2 === 0 ? 'Male' : 'Female',
        isoDate(dobY, dobM, dobD),
        String(dobD).padStart(2, '0') + '/' + String(dobM).padStart(2, '0') + '/' + (dobY - 8),
        2026 - dobY,
        ['Single', 'Married', 'Divorced', 'Widowed'][i % 4],
        isoDate(seniorityY, seniorityM, 10),
        '10/' + String(seniorityM).padStart(2, '0') + '/' + (seniorityY - 8),
        LOCS[i % LOCS.length], DEPTS[i % DEPTS.length], job,
        'Department Head', category, empType,
        academic ? 'Masters Degree' : 'Bachelors Degree',
        (2.7 + (i % 12) * 0.1).toFixed(2), 'Approved',
        academic ? 'Advanced Academics' : 'Operational Management',
        salary, 'Family of ' + name, 'PEN-' + (1000 + i),
        String(dobD).padStart(2, '0') + '/' + String(dobM).padStart(2, '0') + '/' + (dobY + 52),
        'NID-' + (55000 + i), 'TIN-' + (30000 + i * 17),
        '+2519' + String(10000000 + i * 137).slice(0, 8),
        '0116' + String(2000 + i).slice(-4), String(100 + i),
        'Emergency Contact ' + i, '+251911' + String(100000 + i).slice(0, 6),
        'Addis Ababa, Ethiopia', 'Verified Unity University staff record',
        role, i % 11 === 0 ? 'On Leave' : 'Active', JSON.stringify(payload),
      ]
    );
    rows.push({ id_number: idNumber, name, department: DEPTS[i % DEPTS.length], job });
  }
  return rows;
}

async function seedLeave(emps) {
  const types = ['Annual Leave', 'Sick Leave', 'Emergency Leave', 'Maternity Leave', 'Study Leave'];
  for (let i = 0; i < 15; i++) {
    const e = emps[i % emps.length];
    let start = isoDate(2026, 6, 1 + (i % 20));
    let end = isoDate(2026, 6, 5 + (i % 20));
    let status = 'Approved';
    if (i < 4) {
      start = '2026-08-25';
      end = '2026-09-' + String(10 + i).padStart(2, '0');
      status = 'Approved';
    } else if (i < 8) {
      start = '2026-09-05';
      end = '2026-09-' + String(12 + (i % 5)).padStart(2, '0');
      status = 'Pending';
    }
    const days = Math.max(1, Math.round((new Date(end) - new Date(start)) / 86400000) + 1);
    await pool.query(
      `INSERT INTO hr_leave (id, employee_name, id_number, leave_type, start_date, end_date, days, status, department, payload)
       VALUES ($1,$2,$3,$4,$5::date,$6::date,$7,$8,$9,$10::jsonb)`,
      [uid(), e.name, e.id_number, types[i % types.length], start, end, days, status, e.department,
        JSON.stringify({ reason: ['Family matter', 'Medical', 'Personal', 'Conference', 'Exam'][i % 5], approved_by: status === 'Approved' ? 'hr_manager' : '', remarks: status === 'Pending' ? 'Awaiting supervisor approval' : 'Recorded', days_taken: String(days) })]
    );
  }
}

async function seedVacancies() {
  const rows = [
    ['VAC-2026-001', 'Senior Software Engineer', 'IT Department', 'Published', '2026-10-15'],
    ['VAC-2026-002', 'HR Specialist', 'Human Resource', 'Open', '2026-09-30'],
    ['VAC-2026-003', 'Lecturer – Engineering', 'Engineering', 'Published', '2026-11-01'],
    ['VAC-2026-004', 'Registrar Officer', 'Registrar', 'Open', '2026-09-20'],
    ['VAC-2026-005', 'Store Keeper', 'Material Management', 'Closed', '2026-08-01'],
    ['VAC-2026-006', 'Finance Officer', 'Finance', 'Open', '2026-10-05'],
  ];
  for (const [no, title, dept, status, close] of rows) {
    await pool.query(
      `INSERT INTO hr_vacancies (id, title, department, status, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), title, dept, status, JSON.stringify({ vacancy_no: no, position_title: title, no_of_positions: 1, employment_type: 'Full-Time', closing_date: close, description: title + ' — Unity University institutional vacancy.' })]
    );
  }
}

async function seedRecruitment(emps) {
  const statuses = ['In Progress', 'Shortlisted', 'Completed', 'In Progress', 'Shortlisted'];
  for (let i = 0; i < 10; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_recruitment (id, vacancy_title, department, candidate, status, payload)
       VALUES ($1,$2,$3,$4,$5,$6::jsonb)`,
      [uid(), ['HR Specialist', 'Lecturer – Engineering', 'Finance Officer', 'Registrar Officer'][i % 4], e.department, e.name, statuses[i % statuses.length],
        JSON.stringify({ form_no: 'HRS/036-0' + (i + 1), recruit_name: e.name, total_score: 70 + (i % 25), rank: (i % 5) + 1, recommendation: 'Recommended for appointment', report_date: isoDate(2026, 7, 1 + i) })]
    );
  }
}

async function seedTraining(emps) {
  for (let i = 0; i < 12; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_training (id, title, employee_name, status, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), ['IFRS Update', 'Pedagogy Workshop', 'Leadership Skills', 'OSH Awareness', 'SIS User Training'][i % 5], e.name, ['Completed', 'Planned', 'In Progress'][i % 3],
        JSON.stringify({
          employee_id: e.id_number, department: e.department,
          training_name: ['IFRS Update', 'Pedagogy Workshop', 'Leadership Skills', 'OSH Awareness', 'SIS User Training'][i % 5],
          date_completed: isoDate(2026, 3 + (i % 6), 12), duration_days: 5 + (i % 3), score: 75 + (i % 20),
          start_date: isoDate(2026, 3 + (i % 6), 5), end_date: isoDate(2026, 3 + (i % 6), 12),
          provider: 'Unity University CPD', cost_etb: 2500 + i * 300, location: LOCS[i % LOCS.length]
        })]
    );
  }
}

async function seedAppraisal(emps) {
  for (let i = 0; i < 12; i++) {
    const e = emps[i % emps.length];
    const score = 60 + (i % 35);
    await pool.query(
      `INSERT INTO hr_appraisal (id, employee_name, score, period, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), e.name, score, '2025/26', JSON.stringify({
          employee_id: e.id_number, department: e.department, job_title: e.job,
          appraisal_date: isoDate(2026, 6, 15), appraisal_score: score,
          feedback_manager: 'Aligned with institutional KPIs.', feedback_peer: 'Collaborative team member.',
          feedback_self: 'Committed to continuous improvement.', rating: score >= 85 ? 'Excellent' : score >= 70 ? 'Good' : 'Satisfactory',
          reviewer: 'Department Head', status: 'Completed'
        })]
    );
  }
}

async function seedInsurance(emps) {
  for (let i = 0; i < 12; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_insurance (id, employee_name, amount, department, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), e.name, 1500 + i * 50, e.department, JSON.stringify({
          employee_id: e.id_number, policy_no: 'POL-2026-' + (1000 + i), type: 'Medical', coverage: 'Group Medical',
          effective_date: isoDate(2026, 1, 1), premium_amount: 1500 + i * 50, status: 'Active',
          provider: 'Ethiopian Insurance Corporation', currency: 'ETB'
        })]
    );
  }
}

async function seedAwardsPenalties(emps) {
  for (let i = 0; i < 8; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_awards (id, employee_name, category, amount, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), e.name, ['Best Lecturer', 'Service Excellence', 'Innovation Award', 'Community Service'][i % 4], 2000 + i * 500,
        JSON.stringify({
          employee_id: e.id_number, department: e.department, date: isoDate(2026, 5, 1 + i),
          award_type: ['Best Lecturer', 'Service Excellence', 'Innovation Award', 'Community Service'][i % 4],
          description: 'Institutional recognition', monetary_value: 2000 + i * 500, status: 'Granted'
        })]
    );
    await pool.query(
      `INSERT INTO hr_penalties (id, employee_name, reason, amount, payload)
       VALUES ($1,$2,$3,$4,$5::jsonb)`,
      [uid(), emps[(i + 3) % emps.length].name, ['Late attendance', 'Policy breach', 'Incomplete report'][i % 3], i % 2 === 0 ? 0 : 500 + i * 50,
        JSON.stringify({ employee_id: emps[(i + 3) % emps.length].id_number, severity: ['Low', 'Medium', 'High'][i % 3], date: isoDate(2026, 4, 1 + i), action_taken: ['Written warning', 'Verbal warning', 'Fine'][i % 3] })]
    );
  }
}

async function seedBags(emps) {
  for (let i = 0; i < 8; i++) {
    const e = emps[i % emps.length];
    const last = i < 3 ? '2026-09-' + String(1 + i).padStart(2, '0') : i < 6 ? '2026-09-20' : '2026-10-10';
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'RESIGNATION_CRUD', JSON.stringify({ id: uid(), employee_name: e.name, employee_id: e.id_number, department: e.department, resignation_date: '2026-08-15', last_working_day: last, reason: ['Better Opportunity', 'Personal Reasons', 'Relocation', 'Further Study'][i % 4], acceptance_status: i < 6 ? 'Accepted' : 'Pending', handover_completed: i < 2 ? 'No' : 'Yes', remarks: i < 3 ? 'Final working day this month — complete clearance' : 'Standard exit' })]
    );
  }
  for (let i = 0; i < 6; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'RETIREMENT_CRUD', JSON.stringify({ id: uid(), employee_name: e.name, employee_id: e.id_number, department: e.department, retirement_date: i < 2 ? '2026-10-15' : i < 4 ? '2026-12-01' : '2027-03-01', reason: i % 2 === 0 ? 'Mandatory age — Standard' : 'Early retirement request', pension_details: 'Pension Plan Alpha — PEN-' + (1000 + i), status: i < 4 ? 'In Progress' : 'Completed', notes: 'Handover with HR in progress' })]
    );
  }
  for (let i = 0; i < 6; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'HR_TRANSFER_CRUD', JSON.stringify({
          id: uid(), employee_name: e.name, employee_id: e.id_number,
          current_department: e.department, new_department: DEPTS[(i + 3) % DEPTS.length],
          current_location: LOCS[i % LOCS.length], new_location: LOCS[(i + 2) % LOCS.length],
          transfer_reason: 'Institutional redistribution', effective_date: isoDate(2026, 7, 1 + i), status: 'Approved'
        })]
    );
  }
  for (let i = 0; i < 10; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'HR_COMPETENCE_CRUD', JSON.stringify({
          id: uid(), employee_name: e.name, employee_id: e.id_number,
          skill: ['Teaching', 'Research', 'Leadership', 'IT Literacy', 'Customer Service'][i % 5],
          competency: ['Teaching', 'Research', 'Leadership', 'IT Literacy', 'Customer Service'][i % 5],
          score: 65 + (i % 30), last_assessed: isoDate(2026, 5, 1 + i), assessed_on: isoDate(2026, 5, 1 + i),
          department: e.department, evaluator: 'Department Head', status: 'Completed'
        })]
    );
  }
  for (let i = 0; i < 5; i++) {
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'DEPARTMENTS_HIRING_REQUESTS_CRUD', JSON.stringify({ id: uid(), department: DEPTS[i % DEPTS.length], position_title: ['Assistant Lecturer', 'Accountant', 'Lab Technician'][i % 3], positions: 1 + (i % 2), urgency: ['Normal', 'High', 'Critical'][i % 3], requested_by: empName(i), request_date: isoDate(2026, 6, 1 + i), status: i < 3 ? 'Open' : 'Approved', justification: 'Workload and institutional expansion' })]
    );
  }
  for (let i = 0; i < 8; i++) {
    const e = emps[i % emps.length];
    await pool.query(
      `INSERT INTO hr_records (id, collection, payload) VALUES ($1,$2,$3::jsonb)`,
      [uid(), 'HR_STAFF_ASSIGNMENT_CRUD', JSON.stringify({ id: uid(), employee_name: e.name, employee_id: e.id_number, assignment: ['Committee Member', 'Course Coordinator', 'Campus Duty', 'Project Lead'][i % 4], department: e.department, start_date: isoDate(2026, 1, 15), end_date: isoDate(2026, 12, 31), status: 'Active' })]
    );
  }
}

async function seedMaterial() {
  const items = [
    ['SKU-A4-80', 'A4 Copy Paper (Ream)', 'Stationery', 'Ream', 120, 40, 280, 'Main Store – West', 'West Campus'],
    ['SKU-HP-26A', 'HP Laser Toner 26A', 'IT Consumable', 'Pcs', 18, 5, 4100, 'IT Store', 'West Campus'],
    ['SKU-CHAIR-STD', 'Office Chair Standard', 'Furniture', 'Pcs', 24, 6, 3400, 'Furniture Store', 'Gerji Campus'],
    ['SKU-PEN-BL', 'Ballpoint Pen Blue', 'Stationery', 'Box', 80, 20, 115, 'Main Store – West', 'West Campus'],
    ['SKU-DET-5L', 'Floor Detergent 5L', 'Cleaning', 'Bottle', 40, 10, 200, 'Main Store – West', 'West Campus'],
    ['SKU-USB-32', 'USB Flash 32GB', 'IT Consumable', 'Pcs', 50, 15, 340, 'IT Store', 'West Campus'],
    ['SKU-MARKER', 'Whiteboard Marker Set', 'Stationery', 'Set', 35, 10, 180, 'Main Store – West', 'West Campus'],
    ['SKU-DESK-OAK', 'Executive Desk Oak', 'Furniture', 'Pcs', 8, 2, 12500, 'Furniture Store', 'Gerji Campus'],
    ['SKU-NET-CAT6', 'Cat6 Network Cable 305m', 'IT Consumable', 'Box', 6, 2, 9800, 'IT Store', 'West Campus'],
    ['SKU-PAPER-TOWEL', 'Paper Towel Roll', 'Cleaning', 'Roll', 90, 25, 45, 'Main Store – West', 'West Campus'],
  ];
  for (const [sku, name, cat, uom, qty, reorder, cost, loc, campus] of items) {
    await pool.query(
      `INSERT INTO mat_inventory (id, sku, item_name, category, uom, qty_on_hand, unit_cost, location, coa_account, status, payload)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,'1520-01-10','Active',$9::jsonb)`,
      [uid(), sku, name, cat, uom, qty, cost, loc, JSON.stringify({
          name, item_name: name, sku, category: cat, item_type: 'Consumable', unit: uom, uom,
          quantity: qty, qty_on_hand: qty, reorder_level: reorder, unit_cost: cost, unit_price: cost,
          total_value: qty * cost, location: loc, campus, coa_account: '1520-01-10',
          asset_tag: sku, condition: 'New', status: 'Active', preferred_supplier: 'Office Supplies Plc',
          last_updated: '2026-08-01', remarks: 'Typical stocked item — Unity University'
        })]
    );
  }

  const grvs = [
    ['GRV-2026-001', '2026-03-05', 'Office Supplies Plc', 'A4 Copy Paper (Ream)', 100, 280],
    ['GRV-2026-002', '2026-03-18', 'Import Trade ET', 'HP Laser Toner 26A', 10, 4100],
    ['GRV-2026-003', '2026-04-02', 'Clean Pro ET', 'Floor Detergent 5L', 20, 200],
    ['GRV-2026-004', '2026-05-12', 'Office Supplies Plc', 'Ballpoint Pen Blue', 40, 115],
    ['GRV-2026-005', '2026-06-01', 'Tech Distributors', 'USB Flash 32GB', 30, 340],
    ['GRV-2026-006', '2026-07-10', 'Furniture House', 'Office Chair Standard', 8, 3400],
  ];
  for (const [no, date, supplier, item, qty, cost] of grvs) {
    await pool.query(
      `INSERT INTO mat_grv (id, grv_no, date, supplier, item_name, qty_received, unit_cost, line_value, coa_account, posted_to_finance, finance_status, payload)
       VALUES ($1,$2,$3::date,$4,$5,$6,$7,$8,'1520-01-10',FALSE,'Draft',$9::jsonb)`,
      [uid(), no, date, supplier, item, qty, cost, qty * cost, JSON.stringify({
          po_ref: 'LPO-2026-0' + no.slice(-2), po_number: 'LPO-2026-0' + no.slice(-2),
          purchase_type: 'Local Purchase', item_sku: 'SKU-' + no.slice(-3),
          quantity: qty, qty_received: qty, unit: 'Pcs', unit_cost: cost, line_value: qty * cost,
          receiving_store: 'Main Store – West', store_location: 'Main Store – West',
          campus: 'West Campus', coa_account: '1520-01-10', inspected_by: 'Store Keeper',
          status: 'Complete', inspection_notes: 'Inspected and accepted', remarks: 'Inspected and accepted'
        })]
    );
  }

  const givs = [
    ['GIV-2026-001', '2026-04-10', 'Registrar', 'A4 Copy Paper (Ream)', 15, 280],
    ['GIV-2026-002', '2026-05-02', 'IT Department', 'HP Laser Toner 26A', 2, 4100],
    ['GIV-2026-003', '2026-06-15', 'Human Resource', 'Ballpoint Pen Blue', 10, 115],
    ['GIV-2026-004', '2026-07-20', 'Engineering', 'USB Flash 32GB', 5, 340],
  ];
  for (const [no, date, dept, item, qty, cost] of givs) {
    await pool.query(
      `INSERT INTO mat_giv (id, giv_no, date, department, item_name, quantity, unit_cost, line_value, coa_account, coa_expense, posted_to_finance, finance_status, payload)
       VALUES ($1,$2,$3::date,$4,$5,$6,$7,$8,'1520-01-10','6210-01-10',FALSE,'Draft',$9::jsonb)`,
      [uid(), no, date, dept, item, qty, cost, qty * cost, JSON.stringify({
          request_no: 'REQ-' + no.slice(-3), requested_by: 'Department Head', item_sku: 'SKU-' + no.slice(-3),
          unit: 'Pcs', store_location: 'Main Store – West', campus: 'West Campus',
          purpose: 'Departmental operational use', issued_by: 'Store Keeper', received_by: dept,
          status: 'Issued', remarks: 'Issued to department', coa_expense: '6210-01-10'
        })]
    );
  }

  const pos = [
    ['LPO-2026-011', '2026-02-20', 'Office Supplies Plc', 'A4 Copy Paper (Ream)', 35000],
    ['LPO-2026-012', '2026-03-01', 'Furniture House', 'Office Chair Standard', 27200],
    ['LPO-2026-013', '2026-03-15', 'Clean Pro ET', 'Floor Detergent 5L', 8000],
    ['LPO-2026-014', '2026-04-01', 'Import Trade ET', 'HP Laser Toner 26A', 41000],
    ['LPO-2026-015', '2026-05-10', 'Tech Distributors', 'USB Flash 32GB', 12000],
  ];
  for (const [no, date, supplier, item, total] of pos) {
    await pool.query(
      `INSERT INTO mat_po_local (id, lpo_no, date, supplier, item_name, total_value, coa_account, posted_to_finance, finance_status, payload)
       VALUES ($1,$2,$3::date,$4,$5,$6,'1520-01-10',FALSE,'Draft',$7::jsonb)`,
      [uid(), no, date, supplier, item, total, JSON.stringify({
          status: 'Approved', currency: 'ETB', payment_terms: 'Net 30', quantity: 10,
          unit_cost: total / 10, campus: 'West Campus', remarks: 'Local purchase approved'
        })]
    );
  }

  const robs = [
    ['ROB-2026-001', '2026-05-20', 'CPU Fan from PC-Lab-03', 450],
    ['ROB-2026-002', '2026-06-11', 'Chair armrest — Room 204', 800],
    ['ROB-2026-003', '2026-07-05', 'Printer tray assembly', 1200],
  ];
  for (const [no, date, part, val] of robs) {
    await pool.query(
      `INSERT INTO mat_robbing (id, robbing_no, date, part_taken, part_value, line_value, posted_to_finance, finance_status, payload)
       VALUES ($1,$2,$3::date,$4,$5,$5,FALSE,'Draft',$6::jsonb)`,
      [uid(), no, date, part, val, JSON.stringify({
          reason: 'Cannibalization for urgent repair', source_asset: 'Institutional equipment',
          destination: 'Repair pool', authorized_by: 'Material Manager', status: 'Completed',
          remarks: 'Part transferred under robbing voucher'
        })]
    );
  }
}

async function main() {
  console.log('Seeding HR + Material typical demo (full cells, ETB)…');
  await clearHrMat();
  await seedDepartments();
  const emps = await seedEmployees();
  await seedLeave(emps);
  await seedVacancies();
  await seedRecruitment(emps);
  await seedTraining(emps);
  await seedAppraisal(emps);
  await seedInsurance(emps);
  await seedAwardsPenalties(emps);
  await seedBags(emps);
  await seedMaterial();

  const counts = await pool.query(`
    SELECT 'hr_employees' t, count(*)::int c FROM hr_employees
    UNION ALL SELECT 'departments', count(*) FROM departments
    UNION ALL SELECT 'hr_leave', count(*) FROM hr_leave
    UNION ALL SELECT 'hr_vacancies', count(*) FROM hr_vacancies
    UNION ALL SELECT 'hr_recruitment', count(*) FROM hr_recruitment
    UNION ALL SELECT 'hr_training', count(*) FROM hr_training
    UNION ALL SELECT 'hr_appraisal', count(*) FROM hr_appraisal
    UNION ALL SELECT 'hr_insurance', count(*) FROM hr_insurance
    UNION ALL SELECT 'hr_awards', count(*) FROM hr_awards
    UNION ALL SELECT 'hr_penalties', count(*) FROM hr_penalties
    UNION ALL SELECT 'hr_records', count(*) FROM hr_records
    UNION ALL SELECT 'mat_inventory', count(*) FROM mat_inventory
    UNION ALL SELECT 'mat_grv', count(*) FROM mat_grv
    UNION ALL SELECT 'mat_giv', count(*) FROM mat_giv
    UNION ALL SELECT 'mat_po_local', count(*) FROM mat_po_local
    UNION ALL SELECT 'mat_robbing', count(*) FROM mat_robbing
    ORDER BY 1
  `);
  console.log('Demo counts:');
  counts.rows.forEach((r) => console.log(' ', String(r.t).padEnd(18), r.c));
  await pool.end();
  console.log('Done. Restart API if needed, then hard-refresh frontend (bootstrap).');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
