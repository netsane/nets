import ModuleAccessGate from '../../unity/ModuleAccessGate.jsx';
import { attemptSharedLogin, upsertSharedUsers, upsertSharedUser, saveSharedRoles, saveSharedSession, loadSharedSession, clearSharedSession, canAccessModule, getSharedRoles, mergeSharedRolesInto } from '../../unity/SharedCredentials.js';
import { postMaterialToFinance as bridgePostMaterialToFinance, afterMaterialBootstrap, getMaterialCoaOptions, MAT_SOURCE_COLLECTION_MAP } from '../../unity/UnityModuleBridge.js';
import LoginModuleSwitcher from '../../unity/LoginModuleSwitcher.jsx';
import { MobileStyles, useIsMobile } from '../../unity/mobileCompat.js';
               
import React, { createContext, useContext, useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { createPortal } from 'react-dom';
import * as Icons from 'lucide-react';

import {

/* Gender helpers (hoisted early — used by mock generators and unityUniversity AI) */

    BarChart,
    Bar,
    XAxis,
    YAxis,
    CartesianGrid,
    Tooltip,
    ComposedChart ,
    ResponsiveContainer,
    Legend,
    LineChart,
    PieChart, // Add this
    Pie,      // Add this
    Cell ,     // Add this
    Line
} from 'recharts';
import html2canvas from 'html2canvas'; // Import html2canvas
import { CSVLink } from "react-csv";
import { ChromePicker } from 'react-color';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as docx from 'docx';
import { saveAs } from 'file-saver';

/* Gender helpers (early — used by mock generators and AI) */
const UU_FEMALE_FIRST_EARLY = new Set(['lomi','saba','tigist','eleni','hanna','aster','netsanet','sara','sarah','bethlehem','mahlet','selam','rahel','helen','meseret','hiwot','kidist','marta','martha','yeshi','almaz','tsion','beten','sosina','liya','ruth','meron','eden','abeba','tsehay','worknesh','senait','bezawit','hawi','meaza','frehiwot','genet','ayelech','mekdes','selamawit','tsehaynesh','bethel','hirut','firehiwot','yelena']);
const UU_MALE_FIRST_EARLY = new Set(['abebe','kebede','chala','jemal','fasil','bereket','yonas','henok','kalab','tadesse','dawit','lemma','daniel','samuel','solomon','yared','mulugeta','getachew','tesfaye','hailu','mekonnen','assefa','girma','wolde','negash','bekele','yonatan','nahom','abel','kaleb','biruk','natnael','robelt','kirubel','ammanuel','amanuel','eyob','yohannes','melaku','sisay','asfaw','desta','haile','tewodros','mengistu','berhanu','gashaw','tamiru','fikadu','teklu','damte']);
function uuGuessGenderFromName(fullName) {
    const first = String(fullName || '').trim().split(/\s+/)[0].toLowerCase();
    if (!first) return '';
    if (UU_FEMALE_FIRST_EARLY.has(first)) return 'Female';
    if (UU_MALE_FIRST_EARLY.has(first)) return 'Male';
    return '';
}
function uuCorrectEmployeeGender(e) {
    if (!e || typeof e !== 'object') return e;
    const name = e.full_name || e.employee_name || e.name || e.Name || '';
    const guessed = uuGuessGenderFromName(name);
    if (!guessed) return e;
    return Object.assign({}, e, { gender: guessed, Gender: guessed });
}
function uuCorrectEmployeeList(list) {
    return Array.isArray(list) ? list.map(uuCorrectEmployeeGender) : [];
}


/* ===== Material ↔ Finance integration (existing modules only) ===== */
const UU_MAT_LEDGER_KEY = 'uu_shared_material_ledger_v1';
const UU_FIN_COA_KEY = 'uu_shared_finance_coa_v1';
const uuLoadMatLedger = () => {
    try {
        const v = JSON.parse(localStorage.getItem(UU_MAT_LEDGER_KEY) || '[]');
        return Array.isArray(v) ? v : [];
    } catch { return []; }
};
const uuAppendMatLedger = (legs) => {
    try {
        const prev = uuLoadMatLedger();
        const next = [...(legs || []), ...prev].slice(0, 5000);
        localStorage.setItem(UU_MAT_LEDGER_KEY, JSON.stringify(next));
        return next;
    } catch (e) { return legs || []; }
};
const uuLoadSharedCoa = () => {
    try {
        const v = JSON.parse(localStorage.getItem(UU_FIN_COA_KEY) || '[]');
        return Array.isArray(v) ? v : [];
    } catch { return []; }
};


/* Global corner scrollbar — long track on main content */
const SCROLLBAR_STYLE_ID = 'uu-mat-scrollbar-styles';
if (typeof document !== 'undefined' && !document.getElementById(SCROLLBAR_STYLE_ID)) {
    const s = document.createElement('style');
    s.id = SCROLLBAR_STYLE_ID;
    s.textContent = `
      .custom-scrollbar { scrollbar-width: auto; scrollbar-color: #22d3ee99 #0f172a; }
      .custom-scrollbar::-webkit-scrollbar { width: 14px; height: 14px; }
      .custom-scrollbar::-webkit-scrollbar-track { background: #0f172a; border-radius: 8px; }
      .custom-scrollbar::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #22d3ee, #0891b2);
        border-radius: 8px;
        border: 3px solid #0f172a;
        min-height: 48px;
      }
      .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #67e8f9; }
      .custom-scrollbar::-webkit-scrollbar-corner { background: #0f172a; }
    `;
    document.head.appendChild(s);
}



// --- Configuration Constants ---

const MAT_LOCATIONS = ['Burayu (Main Store)', 'Gerji Campus (Satellite)', 'Adama Campus (Satellite)', 'Desie Campus (Satellite)', 'West Office (Satellite)', 'Main Office/HQ'];
const MAT_CAMPUSES = ['West Campus', 'East Campus', 'Gerji Campus', 'Burayu Campus', 'Adama Campus', 'Dessie Campus', 'Distance Ed', 'Main Office/HQ'];
const ITEM_UNITS = ['Pcs', 'Kg', 'Ltr', 'Roll', 'Box', 'Set', 'Rim', 'Pack', 'Carton'];
const GRV_STATUSES = ['Completed', 'Partial', 'Pending Inspection', 'Rejected', 'Returned to Vendor'];
const GIV_STATUSES = ['Issued', 'Partial Issue', 'Pending Approval', 'Returned to Store', 'Cancelled'];
const PURCHASE_STATUSES = ['Draft', 'Ordered', 'In-Transit', 'Partial Receipt', 'Received', 'Cancelled', 'Closed'];
const ROBBING_STATUSES = ['Planned', 'In Progress', 'Part Fitted', 'Donor Decommissioned', 'Recipient Restored', 'Scrapped Donor', 'Cancelled'];
const MAT_CATEGORIES = ['Printing & Stationery', 'Computer & Accessories', 'Electrical Materials', 'Cleaning & Sanitary', 'Building Materials', 'Oil & Lubricants', 'Household Materials', 'Medical Supplies', 'Books & Modules', 'Spare Parts', 'Raw Materials', 'Finished Goods', 'Uniform & Outfits', 'Cafeteria Items', 'IT Equipment (Asset)', 'Furniture (Asset)', 'Other Materials'];
const MAT_DEPARTMENTS = ['Engineering', 'Academic', 'Finance', 'Logistics', 'IT Department', 'Registrar', 'Human Resource', 'President Office', 'Material Management', 'Research and Development', 'Operations', 'Legal'];
const MAT_COA_OPTIONS = [
    '1471-01-10 Uniform & Outfits', '1472-01-10 Printing & stationary', '1473-01-10 Electrical Materials',
    '1474-01-10 Computer & Accessories', '1475-01-10 Cleaning & Sanitary', '1476-01-10 Building Materials',
    '1477-01-10 Cafeteria Items', '1478-01-10 Oil & Lubricants', '1479-01-10 Household Materials',
    '1480-01-10 Other Materials', '1410-01-13 Raw Materials', '1430-01-13 Finished Goods',
    '1440-01-13 Spare Part', '1450-01-12 Modules and Books', '1460-01-13 Medicines',
    '1490-01-10 Goods In Transit', '1520-01-10 Plant & M/C', '1510-01-12 Building'
];

/* ======================================================================
   MATERIAL ↔ FINANCE INTEGRATION OUTLETS
   Compatible with App_FIXED finance: transactions, FIN_COA_CRUD, dayJournals
   Journal legs match finance txn shape (debit/credit dual legs, Posted status)
   ====================================================================== */

/** Default GL map when record has no coa_account (Unity inventory / AP / expense) */
const MAT_FINANCE_DEFAULT_GL = {
    inventory: '1474-01-10',           // Computer & Accessories (fallback stock)
    goodsInTransit: '1490-01-10',
    grni: '2115-01-10',                // Goods Received Not Invoiced (AP control)
    tradeCreditors: '2110-01-10',      // Trade Creditors
    cashBank: '1141-01-10',            // Cash at Bank CBE
    materialExpense: '6310-01-10',     // Materials / consumables expense (GIV)
    maintenanceExpense: '6330-01-10',  // Maintenance / repairs (robbing residual)
};

/**
 * Resolve account code from material record or default key.
 * Accepts full labels like "1474-01-10 Computer & Accessories".
 */
const resolveMatGL = (raw, fallbackKey = 'inventory') => {
    const s = String(raw || '').trim();
    if (s) {
        const m = s.match(/(\d{4}-\d{2}-\d{2})/);
        if (m) return m[1];
        if (/^\d{4}/.test(s)) return s.split(/\s+/)[0];
    }
    return MAT_FINANCE_DEFAULT_GL[fallbackKey] || MAT_FINANCE_DEFAULT_GL.inventory;
};

const matMoney = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
};

/**
 * Build balanced dual-leg journal lines for a material event.
 * @param {'GRV'|'GIV'|'LOCAL_PO'|'FOREIGN_PO'|'ROBBING'} source
 * @param {object} record
 * @returns {Array} finance-compatible transaction legs
 */
const buildMaterialJournalLegs = (source, record) => {
    if (!record) return [];
    const date = String(record.date || record.arrival_date || new Date().toISOString().slice(0, 10));
    const partner = record.supplier || record.department || record.preferred_supplier || 'Unity Stores';
    const descBase = record.item_name || record.name || record.part_taken || 'Material movement';
    const uid = String(record.id || Date.now());
    const legs = [];

    const pushPair = (ref, amount, drCode, drName, crCode, crName, narration) => {
        const amt = matMoney(amount);
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
            balancing_account_name: crName,
            balancing_account_type: 'Credit',
            balancing_amount: amt,
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
            balancing_account_name: drName,
            balancing_account_type: 'Debit',
            balancing_amount: amt,
        });
    };

    if (source === 'GRV') {
        const amt = matMoney(record.line_value || record.lineValue || record.total_value)
            || matMoney(record.qty_received || record.quantity || record.qty) * matMoney(record.unit_cost || record.unitCost);
        const inv = resolveMatGL(record.coa_account, 'inventory');
        const grni = resolveMatGL(null, 'grni');
        pushPair(record.grv_no || uid, amt, inv, `Inventory · ${inv}`, grni, 'GRNI / Accrued receipts',
            `GRV ${record.grv_no || ''} · ${descBase} · qty ${record.qty_received || record.quantity || 0}`);
    } else if (source === 'GIV') {
        const amt = matMoney(record.line_value) || matMoney(record.quantity) * matMoney(record.unit_cost);
        const exp = resolveMatGL(record.coa_expense || record.coa_account, 'materialExpense');
        const inv = resolveMatGL(record.coa_account, 'inventory');
        pushPair(record.giv_no || uid, amt, exp, `Material expense · ${exp}`, inv, `Inventory · ${inv}`,
            `GIV ${record.giv_no || ''} · ${descBase} → ${record.department || 'Dept'}`);
    } else if (source === 'LOCAL_PO') {
        // Commitment / invoice recognition: Dr inventory (or GIT), Cr Trade Creditors
        const amt = matMoney(record.total_value);
        const inv = resolveMatGL(record.coa_account, 'inventory');
        const ap = resolveMatGL(null, 'tradeCreditors');
        pushPair(record.lpo_no || uid, amt, inv, `Inventory / commitment · ${inv}`, ap, 'Trade Creditors',
            `Local PO ${record.lpo_no || ''} · ${descBase}`);
    } else if (source === 'FOREIGN_PO') {
        const amt = matMoney(record.total_landed)
            || (matMoney(record.fob_value) + matMoney(record.freight_cost) + matMoney(record.insurance_cost)
                + matMoney(record.customs_duty) + matMoney(record.shipping_cost));
        const git = resolveMatGL(record.coa_account, 'goodsInTransit');
        const ap = resolveMatGL(null, 'tradeCreditors');
        pushPair(record.import_ref || uid, amt, git, `Goods in transit · ${git}`, ap, 'Trade Creditors',
            `Foreign import ${record.import_ref || ''} · ${descBase} · landed cost`);
    } else if (source === 'ROBBING') {
        // Memorandum / zero-value movement by default; optional residual value on record.part_value
        const amt = matMoney(record.part_value || record.line_value || 0);
        if (amt > 0) {
            const exp = resolveMatGL(record.coa_note, 'maintenanceExpense');
            const inv = resolveMatGL(record.coa_note, 'inventory');
            pushPair(record.robbing_no || uid, amt, exp, `Maintenance · ${exp}`, inv, `Inventory · ${inv}`,
                `Robbing ${record.robbing_no || ''} · ${record.part_taken || 'part'} · ${record.donor_item || ''} → ${record.recipient_item || ''}`);
        }
    }
    return legs;
};

const SOURCE_COLLECTION_MAP = {
    GRV: 'MAT_GRV_CRUD',
    GIV: 'MAT_GIV_CRUD',
    LOCAL_PO: 'MAT_PURCHASE_LOCAL_CRUD',
    FOREIGN_PO: 'MAT_PURCHASE_FOREIGN_CRUD',
    ROBBING: 'MAT_ROBBING_CRUD',
};

const COLLECTION_SOURCE_MAP = {
    'MAT_GRV_CRUD': 'GRV',
    'MAT_GIV_CRUD': 'GIV',
    'MAT_PURCHASE_LOCAL_CRUD': 'LOCAL_PO',
    'MAT_PURCHASE_FOREIGN_CRUD': 'FOREIGN_PO',
    'MAT_ROBBING_CRUD': 'ROBBING',
};


const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';
const DEFAULT_ROLE = 'employee';
const ASSIGNABLE_ROLES = ['employee', 'hr_staff', 'hr_manager', 'admin'];
const MOCK_FIREBASE_CONFIG = {
    apiKey: "AIzaSyB_Mock_API_Key_For_Testing",
    authDomain: "mock-project-12345.firebaseapp.com",
    projectId: "mock-project-12345",
    storageBucket: "mock-project-12345.appspot.com",
    messagingSenderId: "1234567890",
    appId: "1:1234567890:web:mockwebid",
    measurementId: "G-MOCKMEASUREMENT"
};
const firebaseConfig = MOCK_FIREBASE_CONFIG;
// NEW: MOCK USERS for Auth (Replaces localStorage user storage)
const MOCK_USERS = [
    { username: 'admin', password: 'password', role: 'admin' },
    { username: 'hr_manager', password: 'password', role: 'hr_manager' },
    { username: 'hr_staff', password: 'password', role: 'hr_staff' },
    { username: 'employee', password: 'password', role: 'employee' },
];
const MOCK_USER_ID = 'MOCK-USER-ID-12345';
const HR_PENALTIES = ['Warning', 'Suspension', 'Termination', 'Fine', 'Demotion', 'Counseling'];
const LEAVE_TYPES = ['Annual Leave', 'Sick Leave', 'Maternity Leave', 'Paternity Leave', 'Study Leave', 'Emergency Leave', 'Unpaid Leave'];
const COMPETENCE_LEVELS = [1, 2, 3, 4, 5];
const MOCK_SKILLS = ['React', 'NodeJS', 'Finance', 'Audit', 'Marketing', 'Sales', 'HR Policy', 'Project Mgmt', 'Negotiation'];
const AMHARIC_TOP_EMPLOYEES = ['Netsanet Abebe', 'Damte Tesfaye', 'Kenu Unity', 'Kalab Unity', 'Abel Unity', 'Teferi Unity'];
const MOCK_AMHARIC_NAMES = ['Abebe', 'Kebede', 'Chala', 'Lomi', 'Saba', 'Jemal', 'Tigist', 'Fasil', 'Eleni', 'Bereket', 'Hanna'];
const MOCK_HOLIDAYS = [{ name: 'Ethiopian New Year', date: 'Sep 11' }, { name: 'Meskel', date: 'Sep 27' }, { name: 'Christmas', date: 'Jan 07' }, { name: 'Easter', date: 'May 05' }];MOCK_HOLIDAYS.sort((a, b) => a.date - b.date);
const MOCK_STAFF_ON_LEAVE = [{ name: 'Dawit Bekele', period: '10/24 - 10/28', reason: 'Annual Leave' }, { name: 'Tigist Fanta', period: '11/01 - 11/05', reason: 'Sick Leave' }];
const MOCK_RETIREMENT_APPROACHING = [{ name: 'Lemma Desta', days_left: 30, department: 'Finance' }, { name: 'Aster Getachew', days_left: 90, department: 'HR' }];
// ---- Data from Excel
const ACADEMIC_YEARS = ['201617', '201718', '201819', '201920', '202021', '202122', '202223', '202324', '202425'];
const ACADEMIC_PERIODS = ['S1', 'S2', 'SS', 'FS'];
const STAFF_ATTRITION_CAUSES = ['1391', '89604', '1404', '1401', '1400', '1402', '1399', '1403', '1398', '1390', '1392', '1394', '1403', '1395', '1397', '1396', '9999', '1405', '1389', '73201', '73202', '73203', '9997'];
const AWARD_CATEGORIES = ['1', '2', '3', '4', '5', '6', '7', '8'];
const COUNTRIES = ['AF', 'AX', 'AL', 'DZ', 'AS', 'AD', 'AO', 'AI', 'AQ', 'AG', 'AC', 'AR', 'AM', 'AW', 'AU', 'AT', 'AZ', 'BS', 'BH', 'BD', 'BB', 'BY', 'BE', 'BZ', 'BJ', 'BM', 'BT', 'BO', 'BA', 'BW', 'BR', 'IO', 'BN', 'BG', 'BF', 'BI', 'KH', 'CM', 'CA', 'CV', 'CF', 'TD', 'CL', 'CN', 'TW', 'CX', 'CC', 'CO', 'KM', 'CG', 'CD', 'CK', 'CR', 'CI', 'HR', 'CU', 'CY', 'CZ', 'DK', 'DJ', 'DM', 'DO', 'EC', 'EG', 'SV', 'EU', 'GQ', 'ER', 'EE', 'ET', 'FK', 'FO', 'FJ', 'FI', 'FR', 'FS', 'GF', 'PF', 'GA', 'GM', 'GE', 'DE', 'GH', 'GI', 'GR', 'GL', 'GD', 'GP', 'GU', 'GT', 'GN', 'GW', 'GY', 'HT', 'HM', 'HN',
    'HK',
    'HU', 'IS', 'IN', 'ID', 'IR', 'IQ', 'IE', 'IT', 'JM', 'JP', 'JO', 'KZ', 'KE', 'KI', 'KW', 'KG', 'LA', 'LV', 'LB', 'LS', 'LR', 'LY', 'LI', 'LT', 'LU', 'MO', 'MK', 'MG', 'MW', 'MY', 'MV', 'ML', 'MT', 'MR', 'MU', 'YT', 'MX', 'FM', 'MD', 'MC', 'MN', 'MS', 'MA', 'MZ', 'MM', 'NA', 'NR', 'NP', 'NL', 'AN', 'NC', 'NZ', 'NI', 'NE', 'NG', 'NU', 'NF', 'KP', 'MP', 'NO', 'NS', 'OM', 'OC', 'PK', 'PW', 'PS', 'PA', 'PG', 'PY', 'PE', 'PH', 'PN', 'PL', 'PT', 'PR', 'QA', 'KS', 'ME', 'RS', 'RE', 'RO', 'RU', 'RW', 'HL', 'LC', 'PM', 'WS', 'SM', 'GS', 'ST', 'SA', 'SN', 'CS', 'SC',
    'SL', 'SG', 'SK', 'SI', 'SB',
    'SO', 'ZA',
    'KR', 'ES', 'LK', 'KN', 'VC', 'SD', 'SR', 'SJ', 'SZ', 'SE', 'CH', 'SY', 'TJ', 'TZ', 'TH', 'TL', 'TG', 'TK', 'TO', 'TT', 'TN', 'TR', 'TM', 'TC', 'TV', 'UG', 'UA', 'DU', 'AE', 'GB', 'UM', 'US', 'UY', 'UZ', 'VU', 'VA', 'VE', 'VN', 'VG', 'VI', 'WF', 'EH', 'WN', 'YE', 'ZM', 'ZW', 'IC'];
const GENDERS = ['M', 'F'];
const ETH_BANDS = ['Band I', 'Band II', 'Band III', 'Band IV', 'Band V', 'Band VI'];
const SUB_STUDY_FIELDS = [
    '0111', '0112', '0113', '0114', '0188', '0211', '0212', '0213', '0214', '0215', '0221', '0222', '0223', '0231', '0232',
    '0288', '0311', '0312', '0313', '0314', '0321', '0322', '0388', '0411', '0412', '0413', '0414', '0415', '0416', '0417',
    '0421', '0488', '0511', '0512', '0521', '0522', '0531', '0532', '0533', '0541', '0542', '0588', '0611', '0612', '0613',
    '0688', '0711', '0712', '0713', '0714', '0715', '0716', '0721', '0722', '0723', '0724', '0731', '0732', '0788', '0811',
    '0812', '0821', '0831', '0841', '0888', '0911', '0912', '0913', '0914', '0915', '0916', '0917', '0921', '0922', '0923',
    '0988', '1011', '1012',
    '1013',
    '1014', '1015', '1021', '1022', '1031', '1032', '1041', '1088'
];
const PROGRAM_STATUSES = ['C', 'G', 'T'];
const PROGRAM_TYPES = ['ELIP', 'CPDHDP', 'IEXT', 'OTH'];
/* ----------------------------------------------------------------------
   1. ROLE-BASED ACCESS CONTROL (RBAC) CONFIGURATION
   ---------------------------------------------------------------------- */


const PERMISSIONS_MAP = {
    'MAT_DASHBOARD_VIEW': ['admin', 'store_manager', 'inventory_clerk'],
    'MAT_ITEM_CRUD': ['admin', 'store_manager'],
    'MAT_GRV_CRUD': ['admin', 'store_manager', 'inventory_clerk'],
    'MAT_GIV_CRUD': ['admin', 'store_manager', 'inventory_clerk'],
    'MAT_ROBBING_CRUD': ['admin', 'maintenance_tech', 'store_manager'],
    'MAT_PURCHASE_LOCAL_CRUD': ['admin', 'procurement_officer'],
    'MAT_PURCHASE_FOREIGN_CRUD': ['admin', 'procurement_officer'],
    'UI_CUSTOMIZATION': ['admin', 'store_manager', 'hr_manager'],
    'MAT_GENERIC_REPORTS': ['admin', 'store_manager'],
    'ROLE_CRUD': ['admin'],
    'HR_ROLE_ASSIGNMENT_CRUD': ['admin', 'hr_manager'],
    'STAFF_SIGNUP': ['admin', 'hr_manager'],
    'STAFF_REGISTER': ['admin', 'hr_manager'],
    'STAFF_CHANGE_PASSWORD': ['admin', 'hr_manager', 'hr_staff', 'employee', 'store_manager', 'inventory_clerk', 'procurement_officer', 'maintenance_tech'],
    'STAFF_UPDATE_PROFILE': ['admin', 'hr_manager', 'hr_staff', 'employee', 'store_manager', 'inventory_clerk', 'procurement_officer', 'maintenance_tech'],
    'CAN_ADD_RECORD': ['admin', 'store_manager', 'inventory_clerk'],
    'CAN_EDIT_RECORD': ['admin', 'store_manager'],
    'CAN_DELETE_RECORD': ['admin'],
};
const MENU_ITEMS = [
    {
        id: 'MAT_CORE', label: 'Material Management', icon: 'Package', children: [
            { id: 'MAT_DASHBOARD_VIEW', label: 'Inventory Dashboard', icon: 'LayoutDashboard' },
            { id: 'MAT_ITEM_CRUD', label: 'Inventory Master', icon: 'HardHat' },
            { id: 'MAT_GRV_CRUD', label: 'Goods Receiving (GRV)', icon: 'Truck' },
            { id: 'MAT_GIV_CRUD', label: 'Goods Issuing (GIV)', icon: 'Store' },
            { id: 'MAT_ROBBING_CRUD', label: 'Parts Robbing (Cannibalization)', icon: 'Hammer' },
        ]
    },
    {
        id: 'PROCUREMENT', label: 'Procurement', icon: 'ShoppingCart', children: [
            { id: 'MAT_PURCHASE_LOCAL_CRUD', label: 'Local Purchase', icon: 'ShoppingBag' },
            { id: 'MAT_PURCHASE_FOREIGN_CRUD', label: 'Foreign Purchase', icon: 'Globe' },
        ]
    },
    {
        id: 'SYSTEM', label: 'System Settings', icon: 'Settings', children: [
            { id: 'MAT_GENERIC_REPORTS', label: 'Reports Generator', icon: 'FileBarChart' },
            { id: 'HR_ROLE_ASSIGNMENT_CRUD', label: 'Role Assignment Manager', icon: 'Fingerprint' },
            { id: 'UI_CUSTOMIZATION', label: 'UI Customization', icon: 'Palette' },
            { id: 'ROLE_CRUD', label: 'Manage Roles', icon: 'ShieldAlert' },
            { id: 'STAFF_SIGNUP', label: 'Signup / System Registration', icon: 'UserPlus' },
            { id: 'STAFF_CHANGE_PASSWORD', label: 'Change Password', icon: 'Key' },
            { id: 'STAFF_UPDATE_PROFILE', label: 'Update Profile', icon: 'Edit' },
        ]
    }
];

/* ----------------------------------------------------------------------
   2. UI CONTEXT FOR THEME CUSTOMIZATION (Standard)
   ---------------------------------------------------------------------- */

const UIContext = createContext();
export const useUI = () => useContext(UIContext);

const UIProvider = ({ children }) => {
    const defaultColor = '#06b6d4';
    const [backgroundImage, setBackgroundImage] = useState(localStorage.getItem('bgImage') || null);
    const [logoURL, setLogoURL] = useState(localStorage.getItem('logoURL') || null);
    const [backgroundColor, setBackgroundColor] = useState(localStorage.getItem('bgColor') || defaultColor);
    const [textColor, setTextColor] = useState(localStorage.getItem('textColor') || '#ffffff');

    // GAUGE STATES
    const [bgScale, setBgScale] = useState(1);
    const [bgBlur, setBgBlur] = useState(0);
    const [bgBrightness, setBgBrightness] = useState(100);
    const [aiDockEnabled, setAiDockEnabled] = useState(() => {
        try { return localStorage.getItem('uu_ai_dock') === '1'; } catch (e) { return false; }
    }); 

    // --- NEW: Global Transparency Injector ---
    // This forcibly removes solid backgrounds from your Main, Sidebar, and Nav components
    useEffect(() => {
        const id = 'global-ui-overrides';
        let style = document.getElementById(id);
        if (!style) {
            style = document.createElement('style');
            style.id = id;
            document.head.appendChild(style);
        }

        if (backgroundImage) {
            style.innerHTML = `
                /* Force all major containers to be transparent */
                main, aside, nav, .glassClass, [class*="bg-cyan"], [class*="bg-white"], [class*="bg-slate"] {
                    background-color: rgba(0, 0, 0, 0.05) !important;
                    backdrop-filter: blur(${bgBlur}px) saturate(120%) !important;
                    -webkit-backdrop-filter: blur(${bgBlur}px) saturate(120%) !important;
                    border-color: rgba(255, 255, 255, 0.1) !important;
                }
                
                /* Fix for the text contrast */
                p, h1, h2, h3, span { text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
            `;
        } else {
            style.innerHTML = '';
        }
    }, [backgroundImage, bgBlur]);

    const resetToDefault = () => {
        localStorage.clear();
        window.location.reload();
    };

    const updateUISettings = (settings) => {
        if (settings.backgroundImage !== undefined) {
            setBackgroundImage(settings.backgroundImage);
            localStorage.setItem('bgImage', settings.backgroundImage);
        }
        if (settings.logoURL !== undefined) {
            setLogoURL(settings.logoURL);
            localStorage.setItem('logoURL', settings.logoURL);
        }
        if (settings.primaryColor) {
            setBackgroundColor(settings.primaryColor);
            localStorage.setItem('bgColor', settings.primaryColor);
        }
        if (settings.textColor) {
            setTextColor(settings.textColor);
            localStorage.setItem('textColor', settings.textColor);
        }
        if (settings.bgScale !== undefined) setBgScale(settings.bgScale);
        if (settings.bgBlur !== undefined) setBgBlur(settings.bgBlur);
        if (settings.aiDockEnabled !== undefined) {
            setAiDockEnabled(!!settings.aiDockEnabled);
            try { localStorage.setItem('uu_ai_dock', settings.aiDockEnabled ? '1' : '0'); } catch (e) {}
        }
        if (settings.bgBrightness !== undefined) setBgBrightness(settings.bgBrightness);
    };

    // Modified rootStyle to prevent "Magnification"
    const rootStyle = useMemo(() => ({
        backgroundImage: backgroundImage ? `url(${backgroundImage})` : 'none',
        backgroundColor: backgroundImage ? 'transparent' : backgroundColor,
        backgroundSize: '100% 100%', // Forces the image to fit the screen exactly
        backgroundPosition: 'center',
        backgroundAttachment: 'fixed',
        backgroundRepeat: 'no-repeat',
        position: 'fixed',
        inset: 0,
        zIndex: -1,
        // Using transform scale allows the user to manually adjust the "closeness"
        filter: `brightness(${bgBrightness}%) blur(${bgBlur}px)`,
        transform: `scale(${bgScale})`,
        transition: 'all 0.5s cubic-bezier(0.4, 0, 0.2, 1)',
    }), [backgroundImage, backgroundColor, bgBlur, bgBrightness, bgScale]);

    return (
        <UIContext.Provider value={{ 
            rootStyle, updateUISettings, resetToDefault, aiDockEnabled,
            uiSettings: { backgroundImage, logoURL, primaryColor: backgroundColor, textColor, bgBlur, bgScale, bgBrightness, aiDockEnabled } 
        }}>
            {/* THE BACKGROUND IMAGE LAYER */}
            <div style={rootStyle} />

            {/* VIGNETTE OVERLAY: Prevents the image from clashing with edges */}
            {backgroundImage && (
                <div style={{
                    position: 'fixed',
                    inset: 0,
                    background: 'radial-gradient(circle, transparent 20%, rgba(0,0,0,0.4) 100%)',
                    zIndex: 0,
                    pointerEvents: 'none'
                }} />
            )}

            {/* UI CONTENT LAYER */}
            <div 
                className="h-full transition-all duration-700"
                style={{ 
                    backgroundColor: 'transparent',
                    color: textColor,
                    position: 'relative',
                    zIndex: 1
                }}
            >
                {children}
            </div>
        </UIContext.Provider>
    );
};

/* ----------------------------------------------------------------------
   3. AUTHENTICATION CONTEXT & HOOKS (Standard)
   ---------------------------------------------------------------------- */

const AuthContext = createContext();
export const useAuth = () => useContext(AuthContext);
export const AuthProvider = ({ children }) => {
    const _sharedSession = (typeof loadSharedSession === 'function' ? loadSharedSession() : null);
    const [isLoggedIn, setIsLoggedIn] = useState(!!_sharedSession);
    const [currentUser, setCurrentUser] = useState(_sharedSession);
    const [loading, setLoading] = useState(false);
    const [currentView, setCurrentView] = useState('MAT_DASHBOARD_VIEW');

    const logout = useCallback(() => {
        setIsLoggedIn(false);
        setCurrentUser(null);
        try { clearSharedSession(); } catch (e) {}
        setCurrentView('AUTH');
    }, []);

    const changePassword = async (oldPassword, newPassword, getCollection, updateRecord) => {
        if (!getCollection || !updateRecord) return { success: false, message: 'Database failure.' };
        const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
        const userRecord = employees.find(u => u.username === currentUser.username);

        if (!userRecord || userRecord.password !== oldPassword.trim()) {
            return { success: false, message: 'Current password incorrect.' };
        }

        const success = await updateRecord('HR_EMPLOYEE_CRUD', userRecord.id, { 
            ...userRecord, 
            password: newPassword.trim() 
        });
        return success ? { success: true, message: 'Updated! Please re-login.' } : { success: false };
    };

    /**
     * THE APRIL 21 CORE LOGIC
     * Dynamically checks if the current user's role ID exists in the database
     * and if that role contains the requested permission.
     */
    const checkPermission = useCallback((permissionId, allRoles = []) => {
        if (!isLoggedIn || !currentUser) return false;
        const role = String(currentUser.role || '');
        const uname = String(currentUser.username || '');
        if (role === 'admin' || uname === 'admin' || role.toLowerCase() === 'admin') return true;

        // Merge live + shared roles (custom role ids like 1787056523255 live here)
        let roles = Array.isArray(allRoles) ? [...allRoles] : [];
        try {
            if (typeof mergeSharedRolesInto === 'function') roles = mergeSharedRolesInto(roles);
            else if (typeof getSharedRoles === 'function') roles = [...roles, ...(getSharedRoles() || [])];
        } catch (e) {}

        const roleDef = roles.find(
            (r) =>
                r &&
                (String(r.id) === role ||
                    String(r.name) === role ||
                    String(r.roleName || '') === role ||
                    String(r.id || '').toLowerCase() === role.toLowerCase() ||
                    String(r.name || '').toLowerCase() === role.toLowerCase())
        );

        if (roleDef) {
            const perms = roleDef.permissions || roleDef.permission_ids || roleDef.allowed || [];
            // Dashboard always allowed if the role exists for this user
            const dashIds = [
                'FIN_DASHBOARD', 'FIN_DASHBOARD_VIEW', 'HR_DASHBOARD_VIEW', 'HR_DASHBOARD',
                'MAT_DASHBOARD_VIEW', 'MAT_DASHBOARD', 'DASHBOARD', 'dashboard',
            ];
            if (dashIds.includes(permissionId) || String(permissionId || '').toUpperCase().includes('DASHBOARD')) {
                return true;
            }
            if (Array.isArray(perms) && perms.length > 0) {
                return perms.includes(permissionId) || perms.includes(String(permissionId).toUpperCase());
            }
            // Role exists but permissions not listed → grant access (role assigned intentionally)
            return true;
        }

        try {
            if (typeof DEFAULT_PERMISSIONS_MAP !== 'undefined' && DEFAULT_PERMISSIONS_MAP[permissionId]) {
                return DEFAULT_PERMISSIONS_MAP[permissionId]
                    .map((x) => String(x).toLowerCase())
                    .includes(role.toLowerCase());
            }
        } catch (e) {}

        // Last resort: logged-in user can open dashboards only
        if (String(permissionId || '').toUpperCase().includes('DASHBOARD')) return true;
        return false;
    }, [isLoggedIn, currentUser]);

    const login = async (username, password, allEmployees = []) => {
        setLoading(true);
        await new Promise((resolve) => setTimeout(resolve, 300));
        try {
            const result = attemptSharedLogin(username, password, allEmployees || []);
            if (result.success && result.user) {
                setCurrentUser(result.user);
                setIsLoggedIn(true);
                try { saveSharedSession(result.user); } catch (e) {}
                setCurrentView('MAT_DASHBOARD_VIEW');
                setLoading(false);
                return { success: true };
            }
            setLoading(false);
            return { success: false, message: (result && result.message) || 'Invalid username or password.' };
        } catch (err) {
            console.error(err);
            setLoading(false);
            return { success: false, message: 'Login error.' };
        }
    };


    const authValue = useMemo(() => ({
        isLoggedIn, currentUser, loading, login, changePassword, 
        checkPermission, currentView, setCurrentView, logout
    }), [isLoggedIn, currentUser, loading, login, checkPermission, logout, currentView]);

    return <AuthContext.Provider value={authValue}>{children}</AuthContext.Provider>;
};
// --- Access Control Component ---
export const AccessControl = ({ permissionId, children, fallback = null }) => {
    const { checkPermission } = useAuth();
    const isAuthorized = checkPermission(permissionId);

    if (isAuthorized) {
        return <>{children}</>;
    }
    return <>{fallback}</>;
};

const generateMockData = (collectionName, count = 35) => {
    const data = [];
    let defaultDate = new Date().toISOString().split('T')[0];
    const demoEmployeeNames = [...AMHARIC_TOP_EMPLOYEES, 'admin', 'hr_manager', 'hr_staff', 'employee'];
    const employeeNames = [...demoEmployeeNames, ...Array.from({ length: count - demoEmployeeNames.length }, (_, i) => `${MOCK_AMHARIC_NAMES[(i + 1) % MOCK_AMHARIC_NAMES.length]} ${i + 7}`)];

    const getRandomDate = () => {
        const year = 2025 - Math.floor(Math.random() * 5);
        const month = String(Math.floor(Math.random() * 12) + 1).padStart(2, '0');
        const day = String(Math.floor(Math.random() * 28) + 1).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    // Helper for Ethiopian Calendar Mock Data (approx 8 years behind GC)
    const getRandomECDate = () => {
        const day = String(Math.floor(Math.random() * 30) + 1).padStart(2, '0');
        const month = String(Math.floor(Math.random() * 12) + 1).padStart(2, '0');
        const year = 2010 + Math.floor(Math.random() * 7);
        return `${day}/${month}/${year}`;
    };

    const departments = [
        'Finance', 'Engineering', 'Human Resource', 'President Office', 'CEO Office',
        'Material Management', 'Academic Department', 'Registrar', 'IT Department', 
        'Marketing', 'Operations', 'Legal', 'Research and Development'
    ];

    for (let i = 1; i <= count; i++) {
        const id = `mock-id-${collectionName}-${i}`;
        const itemDate = getRandomDate();
        const salary = Math.floor(Math.random() * 20000) + 5000;
        
        const education = [
            { title: 'BSc Computer Science', institution: 'AAU', year: '2015' },
            { title: 'Certification in X', institution: 'Online', year: '2022' }
        ];
        if (i % 3 === 0) education.push({ title: 'MSc Business', institution: 'Unity', year: '2020' });
        
        const beneficiaries = [
            { name: 'Spouse', relationship: 'Spouse', percentage: '100' },
            { name: 'Child A', relationship: 'Child', percentage: '50' }
        ];
        if (i % 4 === 0) beneficiaries.push({ name: 'Child B', relationship: 'Child', percentage: '50' });
        
        const employeeName = employeeNames[i - 1] || `Employee ${i}`;
        
        // Logical birthday for age/retirement calculation
        const birthYear = 1965 + Math.floor(Math.random() * 30);
        const birthMonth = Math.floor(Math.random() * 12) + 1;
        const birthDay = Math.floor(Math.random() * 28) + 1;

        // Formatted dates
        let dob_gc = `${birthYear}-${String(birthMonth).padStart(2, '0')}-${String(birthDay).padStart(2, '0')}`;
        let dob_ec = `${String(birthDay).padStart(2, '0')}/${String(birthMonth).padStart(2, '0')}/${birthYear - 8}`;
        let retirement_date_ec = `${String(birthDay).padStart(2, '0')}/${String(birthMonth).padStart(2, '0')}/${(birthYear + 60) - 8}`;

        switch (collectionName) {
            case 'employees':
            case 'HR_EMPLOYEE_CRUD':
                // ROLE LOGIC
                const roleForDemo = employeeName.includes('admin') ? 'admin' : 
                                   employeeName.includes('manager') ? 'hr_manager' : 
                                   employeeName.includes('staff') ? 'hr_staff' : 'employee';

                // LOGICAL ACADEMIC MATCHING
                // If i is even, make them a Lecturer (Academic), otherwise an Officer (Non-Academic)
                const jobTitle = i % 2 === 0 ? `Lecturer ${i % 3 + 1}` : `Officer ${i % 3 + 1}`;
                const category = jobTitle.includes('Lecturer') ? 'Academic' : 'Non-Academic';

                data.push({
                    id, 
                    // Page 2 Fields
                    id_number: `UU-${String(i).padStart(3, '0')}`,
                    name: employeeName, 
                    gender: uuGuessGenderFromName(typeof employeeName !== 'undefined' ? employeeName : (typeof fullName !== 'undefined' ? fullName : '')) || (i % 2 === 0 ? 'Male' : 'Female'),
                    dob_gc: dob_gc,
                    dob_ec: dob_ec,
                    age: (2026 - birthYear).toString(),
                    marital_status: ['Single', 'Married', 'Divorced', 'Widowed'][i % 4],
                    seniority_date_gc: getRandomDate(),
                    seniority_date_ec: getRandomECDate(),

                    // Workplace & Position
                    work_location: ['Gerji', 'Burayu', 'Dessie', 'Adama'][i % 4],
                    department: departments[i % departments.length],
                    job_title: jobTitle, 
                    reporting_to: 'Department Head',
                    employment_category: category, // MODIFIED: Logical link to title
                    employment_type: i % 3 === 0 ? 'Non-Permanent' : 'Permanent',

                    // Page 3 Fields
                    educational_status: jobTitle.includes('Lecturer') ? 'Masters Degree' : 'Bachelors Degree',
                    auth_status: 'Approved',
                    field_of_study: jobTitle.includes('Lecturer') ? 'Advanced Academics' : 'Operational Management',
                    
                    // Page 1 Fields (Financial/Legal/Contact)
                    gross_salary: (salary + 5000).toFixed(2),
                    family_contact: `Family Member ${i}`,
                    pension_id: `PEN-${1000 + i}`,
                    retirement_date_ec: retirement_date_ec,
                    national_id: `NID-${Math.floor(Math.random() * 90000)}`,
                    tin_no: `TIN-${Math.floor(Math.random() * 80000)}`,
                    
                    telephone_number: `+2519${Math.floor(Math.random() * 99999999)}`,
                    office_tel: `0116${String(i).padStart(4, '0')}`,
                    ext_no: `${100 + i}`,
                    emergency_contact_name: `Contact Person ${i}`, 
                    emergency_contact_telephone: `+25111${Math.floor(Math.random() * 9999999)}`,
                    emergency_contact_address: `Addis Ababa, Sub-city ${i % 10}`, 

                    // Array Fields (Multiple Entries)
                    education_entries: education, 
                    remarks: i % 5 === 0 ? "Employee data verified" : "",
                    beneficiaries: beneficiaries,
                    system_role: roleForDemo
                });
                break;

            case 'departments':
                // Only generate records for the length of the departement list to avoid repeats
                if (i <= departments.length) {
                    data.push({
                        id,
                        name: departments[i - 1],
                        head: employeeNames[i % employeeNames.length],
                        location: ['Gerji', 'Burayu', 'Dessie', 'Adama'][i % 4],
                        employees_count: Math.floor(Math.random() * 100) + 10 
                    });
                }
                break;

            case 'departmentsHiringRequests':
                data.push({
                    id,
                    department_id: `mock-id-departments-${(i % 5) + 1}`,
                    department_name: departments[i % departments.length],
                    position_title: `Specialist ${i % 3}`,
                    required_skills: MOCK_SKILLS.slice(0, i % 4 + 1),
                    no_of_vacancies: Math.floor(Math.random() * 5) + 1,
                    status: ['Open', 'Closed', 'On Hold'][Math.floor(Math.random() * 3)],
                    request_date: itemDate,
                    requested_by: employeeNames[i % employeeNames.length],
                    hiring_manager: employeeNames[(i + 1) % employeeNames.length]
                });
                break;
case 'HR_VACANCY_CRUD':
    const positions = ['Software Engineer', 'Senior Accountant', 'HR Specialist', 'Lecturer', 'Administrative Assistant'];
    const vacancyStatuses = ['Published', 'Draft', 'Closed'];
    
    data.push({
        id,
        vacancy_no: `VAC-2026-${String(i).padStart(3, '0')}`,
        position_title: positions[i % positions.length],
        department: departments[i % departments.length],
        no_of_positions: Math.floor(Math.random() * 5) + 1,
        employment_type: ['Full-Time', 'Part-Time', 'Contract'][i % 3],
        closing_date: getRandomDate(),
        status: vacancyStatuses[i % vacancyStatuses.length],
        description: `Responsible for maintaining excellence in the ${departments[i % departments.length]} department.`
    });
    break;
            case 'retirementRecords':
                data.push({
                    id,
                    employee_name: employeeNames[i % employeeNames.length],
                    retirement_date: getRandomDate(),
                    reason: `Reached mandatory retirement age - ${i % 2 === 0 ? 'Standard' : 'Early'}`,
                    pension_details: `Pension Plan Alpha - ID ${Math.floor(Math.random() * 1000)}`,
                    notes: 'Handover process initiated. All documentation pending.'
                });
                break;

            case 'resignationRecords':
                data.push({
                    id,
                    employee_name: employeeName,
                    resignation_date: getRandomDate(),
                    last_working_day: getRandomDate(),
                    reason: ['Better Opportunity', 'Personal Reasons', 'Relocation'][Math.floor(Math.random() * 3)],
                    acceptance_status: ['Accepted', 'Pending', 'Rejected'][Math.floor(Math.random() * 3)],
                    handover_completed: ['Yes', 'No'][Math.floor(Math.random() * 2)],
                });
                break;

            case 'deathRecords':
                data.push({
                    id,
                    employee_name: employeeName,
                    date_of_death: getRandomDate(),
                    cause_of_death: ['Natural Causes', 'Accident', 'Illness'][Math.floor(Math.random() * 3)],
                    reported_by: employeeNames[(i + 1) % employeeNames.length],
                    legal_documentation_status: ['Provided', 'Pending'][Math.floor(Math.random() * 2)],
                });
                break;

          case 'recruitmentRecords':
case 'HR_RECRUITMENT_CRUD':
    // 1. Determine Gender and correct Ethiopian Title
    const isMaleApplicant = i % 2 === 0;
    const ethiopianTitle = isMaleApplicant ? "Ato" : "Wz";
    
    // Pick realistic names based on gender parity
    const maleNames = ['Abebe', 'Kebede', 'Chala', 'Jemal', 'Fasil', 'Bereket'];
    const femaleNames = ['Lomi', 'Saba', 'Tigist', 'Eleni', 'Hanna', 'Aster'];
    
    const firstNameApp = isMaleApplicant 
        ? maleNames[i % maleNames.length] 
        : femaleNames[i % femaleNames.length];
    const lastNameApp = MOCK_AMHARIC_NAMES[(i + 5) % MOCK_AMHARIC_NAMES.length];
    
    const applicantFullName = `${firstNameApp} ${lastNameApp}`;
    const jobTitleRaw = `Senior Specialist - ${departments[i % departments.length]}`;

    // 2. Generate Logical Scores
    const sEdu = Math.floor(Math.random() * 10) + 15; 
    const sExp = Math.floor(Math.random() * 10) + 15; 
    const sKno = Math.floor(Math.random() * 10) + 10; 
    const sCom = Math.floor(Math.random() * 5) + 5;   
    const sPer = Math.floor(Math.random() * 5) + 5;   
    const sSpe = Math.floor(Math.random() * 5) + 5;   
    const total = sEdu + sExp + sKno + sCom + sPer + sSpe;

    data.push({
        id,
        form_no: `HRS/036-01${i}`,
        report_date: itemDate,
        department: departments[i % departments.length],
        position: jobTitleRaw,
        recruit_name: applicantFullName,
        score_education: sEdu,
        score_experience: sExp,
        score_knowledge: sKno,
        score_communication: sCom,
        score_personality: sPer,
        score_special: sSpe,
        total_score: total,
        rank: (i % 5) + 1,
        remark: 'Candidate meets all standard requirements.',
        recommendation: 'Highly recommended for immediate appointment.',
        
        // Matches the phrasing on the paper form label
        selected_candidate_name: `${ethiopianTitle} ${applicantFullName}`, 
        
        committee_comment: 'The candidate demonstrated strong subject matter expertise.',
        
        // Multi-entry data matching the Name, Position, Signature columns on paper
        committee_records: [
            { name: 'Dr. Tadesse Moges', position: 'Academic Dean', signature: 'TM' },
            { name: 'W/ro Selamawit K.', position: 'HR Director', signature: 'SK' },
            { name: 'Ato Henok Ayele', position: 'Department Head', signature: 'HA' }
        ],
        
        verified_by: employeeNames[0],
        verified_by_comment: 'Credentials and professional history verified.',
        approved_by: 'Deputy President Treasury, Manpower and Enterprise Management',
        approved_by_signature: 'SIGNED_E-STAMP_ID_882',
        approval_date: itemDate
    });
    break;   
    // Inside your generateMockData function switch statement:
case 'HR_COMPETENCE_CRUD':
case 'competenceMatrix': 
    // Flatten all skills from the categories into one array
    const allAvailableSkills = Object.values(NEW_MOCK_SKILLS).flat();
    
    data.push({ 
        id,
        // Use the consistent employee name
        employee_name: employeeName,
        employee_id: `UU-00${i}`,
        department: departments[i % departments.length],
        // FIX: This ensures a skill is picked from the full list for every mock record
        skill: allAvailableSkills[i % allAvailableSkills.length], 
        score: (Math.floor(Math.random() * 5) + 1).toString(),
        last_assessed: itemDate,
        evaluator: employeeNames[(i + 5) % employeeNames.length] || 'HR Manager'
    });
    break;
            case 'appraisalRecords':
                data.push({ id, employee_name: employeeName, appraisal_date: itemDate, appraisal_score: employeeName === 'employee' ? '4.5' : (Math.random() * 5).toFixed(1), feedback_manager: `Manager Feedback ${i}`, feedback_peer: `Peer Feedback ${i}`, feedback_self: `Self Review ${i}`, status: employeeName === 'employee' && i % 2 === 0 ? 'Pending' : 'Completed' });
                break;
            case 'leaveRecords':
                data.push({ id, employee_name: employeeName, leave_type: ['Annual Leave', 'Sick Leave', 'Emergency Leave'][i % 3], start_date: itemDate, end_date: itemDate, days_taken: (Math.floor(Math.random() * 10) + 1).toFixed(0), status: 'Approved' });
                break;
            case 'penaltyRecords': data.push({ id, employee_name: employeeName, reason: `Disciplinary Issue ${i}`, date: itemDate, action_taken: HR_PENALTIES[i % HR_PENALTIES.length], severity: ['Low', 'Medium', 'High'][i % 3] });
                break;
            case 'awardRecords': data.push({ id, employee_name: employeeName, award_type: ['Employee of the Month', 'Innovation', 'Service Award'][i % 3], date: itemDate, description: `Awarded for excellent service ${i}`, monetary_value: (Math.random() * 5000).toFixed(2) });
                break;
            case 'transferRecords':
                data.push({ id, employee_name: employeeName, from_place: 'Dessie', to_place: 'Adama', transfer_date: itemDate, reason: `Organizational need ${i}`, status: 'Completed' });
                break;
            case 'trainingRecords': data.push({ id, employee_name: employeeName, training_name: `Skill Training ${i % 5}`, date_completed: itemDate, duration_days: (Math.floor(Math.random() * 5) + 1).toFixed(0), status: ['Completed', 'Scheduled'][i % 2], score: (Math.random() * 100).toFixed(0) });
                break;
            case 'salaryRecords': data.push({ id, employee_name: employeeName, effective_date: itemDate, preview_salary: (salary - 500).toFixed(2), new_salary: salary.toFixed(2), change_reason: ['Annual Increment', 'Promotion', 'Market Adjustment'][i % 3] });
                break;
            case 'insuranceRecords': data.push({ id, employee_name: employeeName, policy_no: `POL-00${i}`, type: ['Health', 'Life', 'Disability'][i % 3], effective_date: itemDate, premium_amount: (Math.random() * 2000).toFixed(2), status: ['Active', 'Expired'][i % 2] });
                break;
            case 'reportHRRecords': data.push({ id, report_name: `HR Report ${i % 3}`, report_type: ['Monthly', 'Quarterly', 'Annual'][i % 3], date_generated: itemDate, department: departments[i % departments.length], data_points: (Math.random() * 1000).toFixed(0), status: ['Generated', 'Draft'][i % 2] });
                break;
            case 'staffAssignmentRecords':
                data.push({ id, institution_code: `I-00${i}`, staff_national_id: `SNID${i}`, academic_year: ACADEMIC_YEARS[i % ACADEMIC_YEARS.length], academic_period: ACADEMIC_PERIODS[i % ACADEMIC_PERIODS.length], campus: `Campus ${i}`, staff_category: ['ACD', 'ADM'][i % 2], job_title: `Job ${i % 5}`, academic_rank: `Rank ${i % 3}`, employment_date: itemDate, employment_type: ['FT', 'PT'][i % 2], monthly_salary: salary.toFixed(2), isced_education_level: `Level ${i % 4}`, is_under_qualified: i % 2 === 0, nbr_courses_taught: i % 3, nbr_credit_hours_taught: i % 4, nbr_modules_taught: i % 5, nbr_postgraduates_advised: i % 6, standard_teaching_load: i % 7, community_service_load: i % 8, research_load: i % 9, teaching_overload: i % 10, overload_reason: `Reason ${i % 3}`, science_training: i % 2 === 0, research_training: i % 2 === 0, teaching_license: i % 2 === 0, completed_cpd_hdp: i % 2 === 0, outgoing_exchange: i % 2 === 0, exchange_country: 'Ethiopia', industry_experience: i % 2 === 0, leadership_training: i % 2 === 0, merit_based: i % 2 === 0, digital_skills_training: i % 2 === 0, professional_training: i % 2 === 0, research_ethics_training: i % 2 === 0, digital_instruction_training: i % 2 === 0, elip_training: i % 2 === 0 });
                break;
            case 'staffAttritionRecords':
                data.push({ id, institution_code: `I-00${i}`, staff_national_id: `SNID${i}`, academic_year: ACADEMIC_YEARS[i % ACADEMIC_YEARS.length], academic_period: ACADEMIC_PERIODS[i % ACADEMIC_PERIODS.length], attrition_cause: STAFF_ATTRITION_CAUSES[i % STAFF_ATTRITION_CAUSES.length] });
                break;
            case 'staffAwardsRecords':
                data.push({ id, institution_code: `I-00${i}`, staff_national_id: `SNID${i}`, academic_year: ACADEMIC_YEARS[i % ACADEMIC_YEARS.length], academic_period: ACADEMIC_PERIODS[i % ACADEMIC_PERIODS.length], award_category: AWARD_CATEGORIES[i % AWARD_CATEGORIES.length], award_date: itemDate, award_amount: salary.toFixed(2) });
                break;
            case 'staffDevelopmentRecords':
                data.push({ id, institution_code: `I-00${i}`, staff_national_id: `SNID${i}`, program_name: `Dev Program ${i}`, program_country: COUNTRIES[i % COUNTRIES.length], target_education_level: `Level ${i % 5}`, special_program_type: PROGRAM_TYPES[i % PROGRAM_TYPES.length], program_status: PROGRAM_STATUSES[i % PROGRAM_STATUSES.length], start_date: itemDate, end_date: itemDate, nbr_participants: i % 10, program_cost: (Math.random() * 10000).toFixed(2) });
                break;
            case 'staffOverviewRecords':
                data.push({ id, institution_code: `I-00${i}`, staff_national_id: `SNID${i}`, name: employeeName, date_of_birth: dob_gc, gender: GENDERS[i % GENDERS.length], nationality: COUNTRIES[i % COUNTRIES.length], eth_band: ETH_BANDS[i % ETH_BANDS.length], sub_study_field_code: SUB_STUDY_FIELDS[i % SUB_STUDY_FIELDS.length], appointment_date: itemDate, employment_type: ['FT', 'PT'][i % 2], academic_rank: `Rank ${i % 3}`, isced_education_level: `Level ${i % 4}`, is_under_qualified: i % 2 === 0, nbr_courses_taught: i % 3, nbr_credit_hours_taught: i % 4 });
                break;
            case 'roleRecords':
                const initialRoles = [
                    { id: 'admin', name: 'admin', permissions: Object.keys(PERMISSIONS_MAP) },
                    { id: 'hr_manager', name: 'hr_manager', permissions: PERMISSIONS_MAP.HR_DASHBOARD_VIEW.concat(PERMISSIONS_MAP.HR_PENALTY_CRUD, PERMISSIONS_MAP.HR_ROLE_ASSIGNMENT_CRUD, PERMISSIONS_MAP.UI_CUSTOMIZATION, PERMISSIONS_MAP.HR_GENERIC_REPORTS, PERMISSIONS_MAP.CAN_DELETE_RECORD) },
                    { id: 'hr_staff', name: 'hr_staff', permissions: PERMISSIONS_MAP.HR_DASHBOARD_VIEW.concat(PERMISSIONS_MAP.HR_RECRUITMENT_CRUD, PERMISSIONS_MAP.HR_COMPETENCE_CRUD, PERMISSIONS_MAP.HR_APPRAISAL_CRUD, PERMISSIONS_MAP.HR_LEAVE_CRUD, PERMISSIONS_MAP.HR_AWARD_CRUD, PERMISSIONS_MAP.HR_TRANSFER_CRUD, PERMISSIONS_MAP.HR_TRAINING_CRUD, PERMISSIONS_MAP.HR_REPORT_CRUD, PERMISSIONS_MAP.HR_GENERIC_REPORTS, PERMISSIONS_MAP.HR_STAFF_ASSIGNMENT_CRUD, PERMISSIONS_MAP.HR_STAFF_DEVELOPMENT_CRUD, PERMISSIONS_MAP.HR_STAFF_OVERVIEW_CRUD, PERMISSIONS_MAP.CAN_ADD_RECORD, PERMISSIONS_MAP.CAN_EDIT_RECORD) },
                    { id: 'employee', name: 'employee', permissions: PERMISSIONS_MAP.HR_DASHBOARD_VIEW.concat(PERMISSIONS_MAP.HR_EMPLOYEE_CRUD, PERMISSIONS_MAP.HR_COMPETENCE_CRUD, PERMISSIONS_MAP.HR_APPRAISAL_CRUD, PERMISSIONS_MAP.HR_LEAVE_CRUD, PERMISSIONS_MAP.HR_TRAINING_CRUD, PERMISSIONS_MAP.HR_STAFF_DEVELOPMENT_CRUD, PERMISSIONS_MAP.CAN_ADD_RECORD, PERMISSIONS_MAP.CAN_EDIT_RECORD, PERMISSIONS_MAP.STAFF_CHANGE_PASSWORD, PERMISSIONS_MAP.STAFF_UPDATE_PROFILE, PERMISSIONS_MAP.STAFF_SIGNUP) }
                ];
                data.push(...initialRoles);
                if (!initialRoles.find(r => r.id === `mock-role-${i}`)) {
                    data.push({ id: `mock-role-${i}`, name: `Mock Role ${i}`, permissions: [Object.keys(PERMISSIONS_MAP)[i % Object.keys(PERMISSIONS_MAP).length]] })
                }
                break;
            default:
                data.push({ id, name: `Generic Item ${i}`, date: itemDate });
                break;
        }
    }
    return data.filter((v, i, a) => a.findIndex(t => t.id === v.id) === i); 
};




// Mock Skills to the Core/Foundational Competencies
// Mock data for new modules

const NEW_MOCK_SKILLS = {
   'Core / Foundational Competencies (Universal for all jobs)': [
      'Cognitive / Thinking Skills',
       'Critical thinking',
       'Analytical reasoning',
       'Problem-solving',
       'Decision-making',
       'Logical reasoning',
       'Creativity & innovation',
       'Systems thinking',
       'Numerical reasoning',
       'Abstract reasoning',
       'Learning agility'
   ],
   'Communication Skills': [
       'Verbal communication',
       'Written communication',
       'Presentation skills',
       'Active listening',
       'Negotiation',
       'Persuasion',
       'Conflict resolution',
       'Cross-cultural communication',
       'Public speaking',
       'Storytelling'
   ],
   'Interpersonal Skills': [
       'Teamwork',
       'Collaboration',
       'Empathy',
       'Emotional intelligence',
       'Relationship building',
       'Influence',
       'Customer orientation',
       'People awareness',
       'Mediation',
       'Diplomacy'
   ],
   'Self-Management / Personal Effectiveness': [
       'Time management',
       'Stress management',
       'Adaptability',
       'Resilience',
       'Accountability',
       'Ethics & integrity',
       'Self-discipline',
       'Self-awareness',
       'Initiative',
       'Work-life balance'
   ]
};

// ... (MOCK_EMPLOYEES definition)

const MOCK_EMPLOYEES = generateMockData('employees', 10);
const MOCK_DEPARTMENTS = generateMockData('departments');
const MOCK_DEPARTMENTS_HIRING_REQUESTS = generateMockData('departmentsHiringRequests');
const MOCK_RETIREMENT_RECORDS = generateMockData('retirementRecords');
const MOCK_RESIGNATION_RECORDS = generateMockData('resignationRecords');
const MOCK_DEATH_RECORDS = generateMockData('deathRecords');

const MOCK_RECRUITMENT_RECORDS = generateMockData('recruitmentRecords');
const MOCK_COMPETENCE_MATRIX = generateMockData('competenceMatrix');
const MOCK_APPRAISAL_RECORDS = generateMockData('appraisalRecords');
const MOCK_LEAVE_RECORDS = generateMockData('leaveRecords');
const MOCK_PENALTY_RECORDS = generateMockData('penaltyRecords');
const MOCK_AWARD_RECORDS = generateMockData('awardRecords');
const MOCK_TRANSFER_RECORDS = generateMockData('transferRecords');
const MOCK_TRAINING_RECORDS = generateMockData('trainingRecords');
const MOCK_INSURANCE_RECORDS = generateMockData('insuranceRecords');
const MOCK_REPORT_HR_RECORDS = generateMockData('reportHRRecords');
const MOCK_STAFF_ASSIGNMENT_RECORDS = generateMockData('staffAssignmentRecords');
const MOCK_STAFF_ATTRITION_RECORDS = generateMockData('staffAttritionRecords');
const MOCK_STAFF_AWARDS_RECORDS = generateMockData('staffAwardsRecords');
const MOCK_STAFF_DEVELOPMENT_RECORDS = generateMockData('staffDevelopmentRecords');
const MOCK_STAFF_OVERVIEW_RECORDS = generateMockData('staffOverviewRecords');
const MOCK_ROLE_RECORDS = [
    {
        id: 'admin',
        name: 'admin',
        permissions: [
            'MAT_DASHBOARD_VIEW',  'HR_RECRUITMENT_CRUD',  'HR_EMPLOYEE_CRUD',
            'HR_COMPETENCE_CRUD',  'HR_APPRAISAL_CRUD',   'HR_LEAVE_CRUD',
            'HR_PENALTY_CRUD',   'HR_AWARD_CRUD',     'HR_TRANSFER_CRUD',
            'HR_TRAINING_CRUD',  'HR_INSURANCE_CRUD',   'HR_REPORT_CRUD',
            'HR_STAFF_ASSIGNMENT_CRUD', 'HR_STAFF_ATTRITION_CRUD',
            'HR_STAFF_AWARDS_CRUD',   'HR_STAFF_DEVELOPMENT_CRUD',
            'HR_STAFF_OVERVIEW_CRUD',  'ROLE_CRUD',         'HR_ROLE_ASSIGNMENT_CRUD',
            'UI_CUSTOMIZATION',    'HR_GENERIC_REPORTS',  'CAN_ADD_RECORD',
            'CAN_EDIT_RECORD',     'CAN_DELETE_RECORD', 'STAFF_REGISTER',
            'STAFF_CHANGE_PASSWORD', 'STAFF_UPDATE_PROFILE','STAFF_SIGNUP'
          ]
      },
      {
        id: 'hr_manager',
        name: 'hr_manager',
        permissions: [
            'MAT_DASHBOARD_VIEW',
            'HR_EMPLOYEE_CRUD',
            'HR_COMPETENCE_CRUD',
            'HR_APPRAISAL_CRUD',
            'HR_LEAVE_CRUD',
            'HR_TRAINING_CRUD',
            'HR_REPORT_CRUD',
            'HR_STAFF_DEVELOPMENT_CRUD',
            'CAN_ADD_RECORD',
            'CAN_EDIT_RECORD',
            'STAFF_CHANGE_PASSWORD',
            'STAFF_UPDATE_PROFILE',
            'STAFF_SIGNUP'
          ]
      },
      {
        id: 'hr_staff',
        name: 'hr_staff',
        permissions: [
            'MAT_DASHBOARD_VIEW',
            'HR_EMPLOYEE_CRUD',
            'HR_COMPETENCE_CRUD',
            'HR_APPRAISAL_CRUD',
            'HR_LEAVE_CRUD',
            'HR_TRAINING_CRUD',
            'HR_REPORT_CRUD',
            'HR_STAFF_DEVELOPMENT_CRUD',
            'CAN_ADD_RECORD',
            'CAN_EDIT_RECORD',
            'STAFF_CHANGE_PASSWORD',
            'STAFF_UPDATE_PROFILE',
            'STAFF_SIGNUP'
          ]
      },
      {
        id: 'employee',
        name: 'employee',
        permissions: [
            'MAT_DASHBOARD_VIEW',
            'HR_EMPLOYEE_CRUD',
            'HR_COMPETENCE_CRUD',
            'HR_APPRAISAL_CRUD',
            'HR_LEAVE_CRUD',
            'HR_TRAINING_CRUD',
            'HR_STAFF_DEVELOPMENT_CRUD',
            'CAN_ADD_RECORD',
            'CAN_EDIT_RECORD',
            'STAFF_CHANGE_PASSWORD',
            'STAFF_UPDATE_PROFILE',
            'STAFF_SIGNUP'
          ]
      },
    ];
const MOCK_VACANCY_RECORDS = [
    {
        id: 'vac-1',
        vacancy_no: 'VAC-2026-001',
        position_title: 'Senior Software Engineer',
        department: 'IT Department',
        no_of_positions: 2,
        employment_type: 'Full-Time',
        closing_date: '2026-05-15',
        status: 'Published',
        description: 'Lead the development of our core HR systems.'
    },
    {
        id: 'vac-2',
        vacancy_no: 'VAC-2026-002',
        position_title: 'HR Specialist',
        department: 'Human Resource',
        no_of_positions: 1,
        employment_type: 'Contract',
        closing_date: '2026-06-01',
        status: 'Draft',
        description: 'Handle recruitment and employee relations.'
    }
];
 
const generateMatMockData = (_slug) => [];;


/** Typical Material demo rows — used when API/bootstrap has no MAT rows yet */
const TYPICAL_MAT_SEED = {
  MAT_ITEM_CRUD: [
    { id: 'sku-a4', sku: 'SKU-A4-80', name: 'A4 Copy Paper (Ream)', category: 'Stationery', item_type: 'Consumable', unit: 'Ream', quantity: 120, reorder_level: 40, unit_cost: 280, total_value: 33600, location: 'Main Store – West', campus: 'West Campus', coa_account: '1520-01-10', condition: 'New', status: 'Active', preferred_supplier: 'Office Supplies Plc', last_updated: '2026-08-01', remarks: '80gsm' },
    { id: 'sku-toner', sku: 'SKU-HP-26A', name: 'HP Laser Toner 26A', category: 'IT Consumable', item_type: 'Consumable', unit: 'Pcs', quantity: 18, reorder_level: 5, unit_cost: 4100, total_value: 73800, location: 'IT Store', campus: 'West Campus', coa_account: '1520-01-10', condition: 'New', status: 'Active', preferred_supplier: 'Import Trade ET', last_updated: '2026-07-15', remarks: '' },
    { id: 'sku-chair', sku: 'SKU-CHR-STD', name: 'Office Chair Standard', category: 'Furniture', item_type: 'Serialized Asset', unit: 'Pcs', quantity: 24, reorder_level: 4, unit_cost: 3400, total_value: 81600, location: 'Furniture Store', campus: 'Gerji Campus', coa_account: '1520-01-10', condition: 'Good', status: 'Active', preferred_supplier: 'Furniture House', last_updated: '2026-07-10', remarks: '' },
    { id: 'sku-usb', sku: 'SKU-USB-32', name: 'USB Flash 32GB', category: 'IT Consumable', item_type: 'Consumable', unit: 'Pcs', quantity: 45, reorder_level: 15, unit_cost: 340, total_value: 15300, location: 'IT Store', campus: 'West Campus', coa_account: '1520-01-10', condition: 'New', status: 'Active', preferred_supplier: 'Tech Distributors', last_updated: '2026-06-01', remarks: '' },
    { id: 'sku-det', sku: 'SKU-DET-5L', name: 'Floor Detergent 5L', category: 'Cleaning', item_type: 'Consumable', unit: 'Bottle', quantity: 30, reorder_level: 10, unit_cost: 200, total_value: 6000, location: 'Main Store – West', campus: 'West Campus', coa_account: '6210-01-10', condition: 'New', status: 'Active', preferred_supplier: 'Clean Pro ET', last_updated: '2026-04-02', remarks: '' },
  ],
  MAT_GRV_CRUD: [
    { id: 'grv-1', grv_no: 'GRV-2026-001', date: '2026-03-05', po_number: 'LPO-2026-011', purchase_type: 'Local Purchase', supplier: 'Office Supplies Plc', item_sku: 'SKU-A4-80', item_name: 'A4 Copy Paper (Ream)', quantity: 100, qty_received: 100, unit: 'Ream', unit_cost: 280, line_value: 28000, store: 'Main Store – West', campus: 'West Campus', coa_account: '1520-01-10', inspected_by: 'Store Officer', status: 'Complete', inspection_notes: 'Inspected and accepted', post_status: 'Posted' },
    { id: 'grv-2', grv_no: 'GRV-2026-002', date: '2026-03-18', po_number: 'LPO-2026-014', purchase_type: 'Local Purchase', supplier: 'Import Trade ET', item_sku: 'SKU-HP-26A', item_name: 'HP Laser Toner 26A', quantity: 10, qty_received: 10, unit: 'Pcs', unit_cost: 4100, line_value: 41000, store: 'IT Store', campus: 'West Campus', coa_account: '1520-01-10', inspected_by: 'IT Store', status: 'Complete', inspection_notes: 'Inspected and accepted', post_status: 'Posted' },
    { id: 'grv-3', grv_no: 'GRV-2026-003', date: '2026-04-02', po_number: 'LPO-2026-013', purchase_type: 'Local Purchase', supplier: 'Clean Pro ET', item_sku: 'SKU-DET-5L', item_name: 'Floor Detergent 5L', quantity: 20, qty_received: 20, unit: 'Bottle', unit_cost: 200, line_value: 4000, store: 'Main Store – West', campus: 'West Campus', coa_account: '1520-01-10', inspected_by: 'Store Officer', status: 'Complete', inspection_notes: 'Inspected and accepted', post_status: 'Posted' },
    { id: 'grv-4', grv_no: 'GRV-2026-005', date: '2026-06-01', po_number: 'LPO-2026-015', purchase_type: 'Local Purchase', supplier: 'Tech Distributors', item_sku: 'SKU-USB-32', item_name: 'USB Flash 32GB', quantity: 30, qty_received: 30, unit: 'Pcs', unit_cost: 340, line_value: 10200, store: 'IT Store', campus: 'West Campus', coa_account: '1520-01-10', inspected_by: 'IT Store', status: 'Complete', inspection_notes: 'Inspected and accepted', post_status: 'Posted' },
    { id: 'grv-5', grv_no: 'GRV-2026-006', date: '2026-07-10', po_number: 'LPO-2026-012', purchase_type: 'Local Purchase', supplier: 'Furniture House', item_sku: 'SKU-CHR-STD', item_name: 'Office Chair Standard', quantity: 8, qty_received: 8, unit: 'Pcs', unit_cost: 3400, line_value: 27200, store: 'Furniture Store', campus: 'Gerji Campus', coa_account: '1520-01-10', inspected_by: 'Store Officer', status: 'Complete', inspection_notes: 'Inspected and accepted', post_status: 'Posted' },
  ],
  MAT_GIV_CRUD: [
    { id: 'giv-1', giv_no: 'GIV-2026-001', date: '2026-03-12', department: 'Registrar', requested_by: 'Abebe Kebede', item_sku: 'SKU-A4-80', item_name: 'A4 Copy Paper (Ream)', qty_issued: 10, unit: 'Ream', unit_cost: 280, line_value: 2800, store: 'Main Store – West', campus: 'West Campus', status: 'Issued', remarks: 'Exam packs' },
    { id: 'giv-2', giv_no: 'GIV-2026-002', date: '2026-06-05', department: 'IT Department', requested_by: 'Sara Hailu', item_sku: 'SKU-USB-32', item_name: 'USB Flash 32GB', qty_issued: 5, unit: 'Pcs', unit_cost: 340, line_value: 1700, store: 'IT Store', campus: 'West Campus', status: 'Issued', remarks: '' },
  ],
  MAT_PURCHASE_LOCAL_CRUD: [
    { id: 'lpo-1', po_number: 'LPO-2026-011', date: '2026-02-20', supplier: 'Office Supplies Plc', item_name: 'A4 Copy Paper (Ream)', quantity: 100, unit_cost: 280, total: 28000, status: 'Received', campus: 'West Campus' },
    { id: 'lpo-2', po_number: 'LPO-2026-014', date: '2026-03-10', supplier: 'Import Trade ET', item_name: 'HP Laser Toner 26A', quantity: 10, unit_cost: 4100, total: 41000, status: 'Received', campus: 'West Campus' },
  ],
  MAT_PURCHASE_FOREIGN_CRUD: [
    { id: 'fpo-1', po_number: 'FPO-2026-001', date: '2026-01-15', supplier: 'Import Trade ET', item_name: 'Network Switch 24-Port', quantity: 2, unit_cost: 18500, total: 37000, status: 'In Transit', campus: 'West Campus', currency: 'USD', fx_rate: 120 },
  ],
  MAT_ROBBING_CRUD: [
    { id: 'rob-1', rob_no: 'ROB-2026-001', date: '2026-05-20', donor_sku: 'SKU-OLD-PC', donor_name: 'Desktop PC (failed PSU)', receiver_sku: 'SKU-LAB-PC', receiver_name: 'Lab PC Build', parts: 'RAM 8GB, HDD 500GB', qty: 1, status: 'Completed', campus: 'West Campus', remarks: 'Cannibalize non-functional unit' },
  ],
};

const INITIAL_ALL_DATA = {
    'MAT_ITEM_CRUD': TYPICAL_MAT_SEED.MAT_ITEM_CRUD,
    'MAT_GRV_CRUD': TYPICAL_MAT_SEED.MAT_GRV_CRUD,
    'MAT_GIV_CRUD': TYPICAL_MAT_SEED.MAT_GIV_CRUD,
    'MAT_ROBBING_CRUD': TYPICAL_MAT_SEED.MAT_ROBBING_CRUD,
    'MAT_PURCHASE_LOCAL_CRUD': TYPICAL_MAT_SEED.MAT_PURCHASE_LOCAL_CRUD,
    'MAT_PURCHASE_FOREIGN_CRUD': TYPICAL_MAT_SEED.MAT_PURCHASE_FOREIGN_CRUD,
    /* Finance outlets — merge with App_FIXED finance store when unified */
    'transactions': [],
    'dayJournals': [],
    'FIN_COA_CRUD': MAT_COA_OPTIONS.map((label, i) => {
        const code = String(label).split(/\s+/)[0];
        return {
            id: code,
            account_no: code,
            account_code: code,
            name: label,
            short_desc: label,
            base: code.split('-')[0],
            status: 'Active',
            module: 'MATERIAL',
        };
    }),
    'chartsOfAccounts': [], // filled from FIN_COA_CRUD on first finance merge
    'ROLE_CRUD': [{ id: 'admin', name: 'admin', permissions: Object.keys(PERMISSIONS_MAP) }],
    'HR_EMPLOYEE_CRUD': [{ id: 'e1', name: 'Super Admin', username: 'admin', password: 'password', role: 'admin' }]
};


/* ----------------------------------------------------------------------
   5. DATA CONTEXT & CRUD HOOKS (Standard)
   ---------------------------------------------------------------------- */

const DataContext = createContext();
export const useData = () => useContext(DataContext);
export const DataProvider = ({ children }) => {
    const [allData, setAllData] = useState(() => ({ ...INITIAL_ALL_DATA }));

    // Load all collections from PostgreSQL via API (same pattern as Finance)
    useEffect(() => {
        let cancelled = false;
        (async () => {
            try {
                const { bootstrap } = await import('../../unity/api');
                const res = await bootstrap();
                const base = (res && res.allData) ? res.allData : {};
                // Material: inventory also exposed as MAT_ITEM_CRUD
                if (Array.isArray(base.MAT_INVENTORY_CRUD) && (!base.MAT_ITEM_CRUD || !base.MAT_ITEM_CRUD.length)) {
                    base.MAT_ITEM_CRUD = base.MAT_INVENTORY_CRUD.map((r) => ({
                        ...r,
                        id: r.id || r.sku,
                        item_name: r.item_name || r.name,
                        name: r.item_name || r.name,
                        quantity: r.qty_on_hand,
                        qty: r.qty_on_hand,
                        unit_price: r.unit_cost,
                        unit_cost: r.unit_cost,
                    }));
                }
                // HR employee field aliases for UI columns
                if (Array.isArray(base.HR_EMPLOYEE_CRUD)) {
                    base.HR_EMPLOYEE_CRUD = base.HR_EMPLOYEE_CRUD.map((e) => ({
                        ...e,
                        full_name: e.full_name || e.name,
                        id_no: e.id_no || e.id_number,
                        current_position: e.current_position || e.job_title,
                        mobile: e.mobile || e.telephone_number,
                    }));
                }
                if (!cancelled && base && Object.keys(base).length) {
                    setAllData((prev) => {
                    const next = { ...prev };
                    Object.keys(base).forEach((k) => {
                      // Backend wins only when it has rows; empty arrays must NOT wipe INITIAL mocks
                      if (Array.isArray(base[k])) {
                        if (base[k].length > 0) next[k] = base[k];
                        else if (!(k in next) || next[k] == null) next[k] = base[k];
                      } else if (base[k] != null) next[k] = base[k];
                    });
                    return next;
                });
                    console.info('[bootstrap HR/MAT]', {
                        emp: (base.HR_EMPLOYEE_CRUD || []).length,
                        leave: (base.HR_LEAVE_CRUD || []).length,
                        inv: (base.MAT_INVENTORY_CRUD || base.MAT_ITEM_CRUD || []).length,
                        grv: (base.MAT_GRV_CRUD || []).length,
                        giv: (base.MAT_GIV_CRUD || []).length,
                    });
                }
            } catch (e) {
                console.warn('[bootstrap HR/MAT] failed — using local empty state', e?.message || e);
            }
        })();
        return () => { cancelled = true; };
    }, []);



    useEffect(() => {
        try {
            const shared = typeof getSharedRoles === 'function' ? getSharedRoles() : [];
            if (!shared.length) return;
            setAllData((prev) => {
                const local = prev.ROLE_CRUD || [];
                const merged = typeof mergeSharedRolesInto === 'function' ? mergeSharedRolesInto(local) : [...shared, ...local];
                if (JSON.stringify(merged) === JSON.stringify(local)) return prev;
                return { ...prev, ROLE_CRUD: merged };
            });
        } catch (e) {}
    }, []);
    useEffect(() => {
        try {
            if (Array.isArray(allData?.ROLE_CRUD) && allData.ROLE_CRUD.length && typeof saveSharedRoles === 'function') {
                saveSharedRoles(allData.ROLE_CRUD);
            }
        } catch (e) {}
    }, [allData?.ROLE_CRUD]);

    const [message, setMessage] = useState(null); // { text, type: 'success'|'error' }
    const [notificationTrigger, setNotificationTrigger] = useState(0);
    const [leaveUpdateTrigger, setLeaveUpdateTrigger] = useState(0);
    const [roleUpdateTrigger, setRoleUpdateTrigger] = useState(0);

    const getCollection = useCallback((collectionName) => {
        const rows = allData[collectionName];
        if (Array.isArray(rows) && rows.length > 0) return rows;
        const seed = typeof TYPICAL_MAT_SEED !== 'undefined' ? TYPICAL_MAT_SEED : null;
        if (seed && Array.isArray(seed[collectionName]) && seed[collectionName].length) {
            return seed[collectionName];
        }
        return Array.isArray(rows) ? rows : [];
    }, [allData]);

    /* --- SYNCHRONIZED DATA ENGINE --- */

    
    const MAT_POST_SOURCE = {
        MAT_GRV_CRUD: 'GRV',
        MAT_GIV_CRUD: 'GIV',
        MAT_PURCHASE_LOCAL_CRUD: 'LOCAL_PO',
        MAT_PURCHASE_FOREIGN_CRUD: 'FOREIGN_PO',
        MAT_ROBBING_CRUD: 'ROBBING',
    };
    const postMatToFinIfNeeded = (collectionName, record) => {
        try {
            const source = MAT_POST_SOURCE[collectionName];
            if (!source || !record) return;
            // Need amount: qty * unit_cost or line_value / total_value
            const r = { ...record };
            if (!r.line_value && (r.quantity || r.qty_received) && r.unit_cost) {
                r.line_value = Number(r.quantity || r.qty_received || 0) * Number(r.unit_cost || 0);
            }
            const result = bridgePostMaterialToFinance(source, r);
            if (result && result.ok) {
                console.info('[Material→Finance]', source, result.message, result.legs?.length);
            } else {
                console.warn('[Material→Finance] no legs', source, result?.message, {
                    qty: r.quantity || r.qty_received,
                    unit_cost: r.unit_cost,
                    line_value: r.line_value,
                });
            }
        } catch (e) {
            console.warn('[Material→Finance] failed', e);
        }
    };

    const addRecord = async (collectionName, newData) => {
        const id = Date.now().toString();
        const currentCollection = allData[collectionName] || [];
        
        let recordToSave = { ...newData, id };

        // 1. Auto-ID logic for Employees (UU-001 format)
        if (collectionName === 'HR_EMPLOYEE_CRUD') {
            const lastNum = currentCollection.reduce((max, emp) => {
                const num = parseInt(emp.id_number?.split('-')[1]) || 0;
                return Math.max(max, num);
            }, 0);
            recordToSave.id_number = `UU-${String(lastNum + 1).padStart(3, '0')}`;
        }

        // 2. Auto-ID logic for Vacancies (VAC-1 format)
        if (collectionName === 'HR_VACANCY_CRUD') {
            const lastNum = currentCollection.reduce((max, vac) => {
                // Splits "VAC-12" or "VAC-2026-12" and gets the numeric part at the end
                const parts = vac.vacancy_no?.split('-');
                const numStr = parts ? parts[parts.length - 1] : "0";
                const num = parseInt(numStr) || 0;
                return Math.max(max, num);
            }, 0);
            recordToSave.vacancy_no = `VAC-${lastNum + 1}`;
        }
        
        // Update the state using the correct state setter
        setAllData(prev => ({
            ...prev,
            [collectionName]: [recordToSave, ...currentCollection]
        }));

        // Material → Finance (GRV / GIV / PO / Robbing)
        postMatToFinIfNeeded(collectionName, recordToSave);

        try {
            if (['HR_EMPLOYEE_CRUD', 'EMPLOYEE_CRUD', 'SYS_USERS_CRUD', 'USERS_CRUD', 'HR_USER_CRUD'].includes(collectionName)) {
                upsertSharedUser({ ...recordToSave, modules: recordToSave.modules || ['material'] });
            }
            if (collectionName === 'ROLE_CRUD') {
                const prev = allData.ROLE_CRUD || [];
                saveSharedRoles([recordToSave, ...prev.filter(r => String(r.id) !== String(recordToSave.id))]);
            }
        } catch (e) {}

        
        return true; // Important for FormModal to know it finished
    };

    const updateRecord = useCallback(async (collectionName, id, updatedFields) => {
        try {
            setAllData(prev => {
                let newData = { ...prev };
                const original = prev[collectionName]?.find(r => r.id === id);
                
                // 1. Apply the update to the current list
                newData[collectionName] = prev[collectionName].map(r => 
                    r.id === id ? { ...r, ...updatedFields } : r
                );

                if (!original) return newData;

                // IDENTITY DETECTION: Capture the employee's identity from the record being edited
                const empName = original.employee_name || original.name;
                const empId = original.employee_id || original.staff_national_id || original.id_number;

                // ---------------------------------------------------------
                // SYNC RULE 1: COMPETENCE MATRIX SCORE -> APPRAISAL SCORE
                // ---------------------------------------------------------
                if (collectionName === 'HR_COMPETENCE_CRUD' && updatedFields.score) {
                    console.log(`Syncing Score: ${updatedFields.score} for ${empName}`);
                    
                    if (newData['HR_APPRAISAL_CRUD']) {
                        newData['HR_APPRAISAL_CRUD'] = newData['HR_APPRAISAL_CRUD'].map(app => {
                            const matches = (app.employee_name === empName) || 
                                            (app.name === empName) || 
                                            (app.employee_id === empId) || 
                                            (app.staff_national_id === empId);

                            if (matches) {
                                return { 
                                    ...app, 
                                    score: updatedFields.score,
                                    manager_feedback: `Updated via Competence Sync: ${new Date().toLocaleDateString()}`
                                };
                            }
                            return app;
                        });
        // Material → Finance on update
        try {
            const mergedRow = { ...(allData[collectionName] || []).find(r => r.id === id), ...updatedFields, id };
            postMatToFinIfNeeded(collectionName, mergedRow);
        } catch (e) {}

                    }
                }

                // ---------------------------------------------------------
                // SYNC RULE 2: GLOBAL EMPLOYEE MASTER SYNC
                // ---------------------------------------------------------
                if (collectionName === 'HR_EMPLOYEE_CRUD') {
                    const oldName = original.name;
                    const oldId = original.id_number;

                    Object.keys(newData).forEach(colKey => {
                        if (colKey === 'HR_EMPLOYEE_CRUD' || colKey === 'ROLE_CRUD') return;
                        
                        newData[colKey] = newData[colKey].map(item => {
                            const isMatch = item.employee_name === oldName || item.name === oldName || 
                                            item.employee_id === oldId || item.staff_national_id === oldId;
                            if (isMatch) {
                                return {
                                    ...item,
                                    ...(item.employee_name !== undefined && { employee_name: updatedFields.name || item.employee_name }),
                                    ...(item.name !== undefined && { name: updatedFields.name || item.name }),
                                    ...(item.employee_id !== undefined && { employee_id: updatedFields.id_number || item.employee_id }),
                                    ...(item.staff_national_id !== undefined && { staff_national_id: updatedFields.id_number || item.staff_national_id }),
                                    ...(item.department !== undefined && { department: updatedFields.department || item.department })
                                };
                            }
                            return item;
                        });
                    });
                }

                return newData;
            });

            setMessage({ text: "System Synchronized Successfully", type: 'success' });
            return true;
        } catch (e) {
            console.error("Update Sync Error: ", e);
            return false;
        }
    }, []);

    const deleteRecord = useCallback(async (collectionName, id) => {
        if (!collectionName || !id) {
            setMessage({ text: `Data missing to delete record from ${collectionName}.`, type: 'error' });
            return false;
        }

        try {
            setAllData(prev => ({
                ...prev,
                [collectionName]: prev[collectionName].filter(record => record.id !== id)
            }));

            if (collectionName === 'HR_LEAVE_CRUD') {
                setLeaveUpdateTrigger(prev => prev + 1);
            }

            setMessage({ text: `${collectionName} record deleted successfully!`, type: 'success' });
            return true;
        } catch (error) {
            setMessage({ text: `Failed to delete record from ${collectionName}: ${error.message}`, type: 'error' });
            return false;
        }
    }, []);

    const updateEmployeeRole = useCallback(async (employeeId, newRole) => {
        const employees = allData['HR_EMPLOYEE_CRUD'];
        const employeeIndex = employees.findIndex(emp => emp.id === employeeId);
        if (employeeIndex !== -1) {
            setAllData(prev => {
                const updatedEmployees = [...prev['HR_EMPLOYEE_CRUD']];
                updatedEmployees[employeeIndex] = { ...updatedEmployees[employeeIndex], system_role: newRole };
                return { ...prev, ['HR_EMPLOYEE_CRUD']: updatedEmployees };
            });
            return { success: true, message: `Role for ${employees[employeeIndex].name} updated to ${newRole}` };
        }
        return { success: false, message: `Employee ID ${employeeId} not found.` };
    }, [allData]);

    const addRole = useCallback(async (roleData) => {
        if (!roleData.name) {
            setMessage({ text: `Role Name is required.`, type: 'error' });
            return false;
        }
        if (allData['ROLE_CRUD'].find(role => role.name === roleData.name)) {
            setMessage({ text: `Role ${roleData.name} already exists.`, type: 'error' });
            return false;
        }

        try {
            const newRole = {
                id: roleData.name,
                name: roleData.name,
                username: roleData.username,
                password: roleData.password,
                permissions: Object.keys(PERMISSIONS_MAP)
            };
            setAllData(prev => {
                const updatedRoles = [...prev['ROLE_CRUD'], newRole];
                return { ...prev, 'ROLE_CRUD': updatedRoles };
            });

            setMessage({ text: `Role ${roleData.name} created successfully!`, type: 'success' });
            return true;
        } catch (error) {
            setMessage({ text: `Failed to add role: ${error.message}`, type: 'error' });
            return false;
        }
    }, [allData]);

    const updateRole = useCallback(async (roleId, updatedPermissions) => {
        try {
            setAllData(prev => {
                const updatedRoles = prev['ROLE_CRUD'].map(role =>
                    role.id === roleId ? { ...role, permissions: updatedPermissions } : role
                );
                return { ...prev, 'ROLE_CRUD': updatedRoles };
            });

            setRoleUpdateTrigger(prev => prev + 1);
            setMessage({ text: `Role ${roleId} updated successfully!`, type: 'success' });
            return true;
        } catch (error) {
            setMessage({ text: `Failed to update role: ${error.message}`, type: 'error' });
            return false;
        }
    }, []);

    const deleteRole = useCallback(async (roleId) => {
        try {
            setAllData(prev => ({
                ...prev,
                'ROLE_CRUD': prev['ROLE_CRUD'].filter(role => role.id !== roleId)
            }));
            setRoleUpdateTrigger(prev => prev + 1);
            setMessage({ text: `Role ${roleId} deleted successfully!`, type: 'success' });
            return true;
        } catch (error) {
            setMessage({ text: `Failed to delete role: ${error.message}`, type: 'error' });
            return false;
        }
    }, []);

    /**
     * Post a material document to Finance (transactions + dayJournals).
     * Mirrors postPayrollToFinance outlet in App_FIXED.
     * @param {'GRV'|'GIV'|'LOCAL_PO'|'FOREIGN_PO'|'ROBBING'} source
     * @param {string} recordId
     */
    const postMaterialToFinance = useCallback((source, recordId) => {
        const collection = SOURCE_COLLECTION_MAP[source];
        if (!collection) {
            setMessage({ text: `Unknown material source: ${source}`, type: 'error' });
            return false;
        }
        const rows = allData[collection] || [];
        const record = rows.find(r => String(r.id) === String(recordId));
        if (!record) {
            setMessage({ text: 'Material record not found for finance post.', type: 'error' });
            return false;
        }
        if (record.posted_to_finance) {
            setMessage({ text: `Already posted (${record.posted_ref || 'see ledger'}).`, type: 'error' });
            return false;
        }
        const legs = buildMaterialJournalLegs(source, record);
        if (!legs.length) {
            setMessage({
                text: source === 'ROBBING'
                    ? 'Robbing posted as memorandum only (set part_value to create value legs).'
                    : 'No amount to post — check qty × cost / line value.',
                type: 'error'
            });
            if (source !== 'ROBBING') return false;
        }
        const postedAt = new Date().toISOString().slice(0, 10);
        const ref = legs[0]?.material_ref || `${source}-${recordId}`;

        // Share journal legs with Finance module (ledger + UnityBus for FinanceApp)
        try {
            if (legs.length) uuAppendMatLedger(legs);
        } catch (e) {}
        try {
            // Bridge writes uu_shared_material_ledger_v1 + pending_journals (Finance reads these)
            bridgePostMaterialToFinance(source, record);
        } catch (e) {
            console.warn('[Material→Finance bridge]', e);
        }
        setAllData(prev => {
            const currentTx = Array.isArray(prev.transactions) ? prev.transactions : [];
            const currentDj = Array.isArray(prev.dayJournals) ? prev.dayJournals : [];
            const col = SOURCE_COLLECTION_MAP[source];
            return {
                ...prev,
                transactions: [...legs, ...currentTx],
                dayJournals: [...legs, ...currentDj],
                [col]: (prev[col] || []).map(r =>
                    String(r.id) === String(recordId)
                        ? {
                            ...r,
                            posted_to_finance: true,
                            posted_at: postedAt,
                            posted_ref: ref,
                            finance_status: 'Posted',
                            module: 'MATERIAL',
                        }
                        : r
                ),
            };
        });
        setMessage({
            text: legs.length
                ? `Posted to Finance: ${source} ${ref} · ${legs.length / 2} balanced journal(s)`
                : `Robbing ${ref} flagged (no value legs).`,
            type: 'success',
        });
        return true;
    }, [allData]);

    /** Batch-post all unposted rows for a material collection */
    const postAllMaterialToFinance = useCallback((collectionName) => {
        const source = COLLECTION_SOURCE_MAP[collectionName];
        if (!source) return 0;
        const rows = (allData[collectionName] || []).filter(r => !r.posted_to_finance);
        let n = 0;
        rows.forEach(r => {
            if (postMaterialToFinance(source, r.id)) n += 1;
        });
        return n;
    }, [allData, postMaterialToFinance]);


    const dataValue = useMemo(() => ({
        allData, getCollection, addRecord, updateRecord, deleteRecord, message, setMessage, notificationTrigger, setNotificationTrigger, leaveUpdateTrigger, setLeaveUpdateTrigger, updateEmployeeRole, addRole, updateRole, deleteRole, roleUpdateTrigger, postMaterialToFinance, postAllMaterialToFinance, buildMaterialJournalLegs
    }), [allData, getCollection, addRecord, updateRecord, deleteRecord, message, notificationTrigger, leaveUpdateTrigger, updateEmployeeRole, addRole, updateRole, deleteRole, roleUpdateTrigger, postMaterialToFinance, postAllMaterialToFinance]);

    return (
        <DataContext.Provider value={dataValue}>
            {children}
        </DataContext.Provider>
    );

    useEffect(() => {
        try {
            const coa = uuLoadSharedCoa();
            if (!coa.length) return;
            setAllData((prev) => ({
                ...prev,
                chartsOfAccounts: coa,
                FIN_COA_CRUD: coa,
            }));
        } catch (e) {}
    }, []);

};

/* ----------------------------------------------------------------------
   6. General Utility Components (Standard)
   ---------------------------------------------------------------------- */

// Helper to convert an icon name string to a lucide-react component
const getIcon = (name) => Icons[name] || Icons.File;

/* --- FIXED INPUT COMPONENT (SELECT + DUAL CALENDARS + REDIRECT) --- */

/** Dark in-app select (filters + forms) — never uses OS white picker on mobile */
const UUSelect = ({ value, onChange, options = [], allLabel = 'All', className = '', compact = false }) => {
    const [open, setOpen] = useState(false);
    const ref = useRef(null);
    const list = Array.isArray(options) ? options : [];

    useEffect(() => {
        if (!open) return undefined;
        const onDoc = (e) => {
            if (ref.current && !ref.current.contains(e.target)) setOpen(false);
        };
        document.addEventListener('mousedown', onDoc);
        document.addEventListener('touchstart', onDoc);
        return () => {
            document.removeEventListener('mousedown', onDoc);
            document.removeEventListener('touchstart', onDoc);
        };
    }, [open]);

    const labelFor = (v) => {
        if (v === 'ALL' || v === 'All') return allLabel;
        return String(v);
    };
    const currentLabel = labelFor(value);

    return (
        <div className={`relative ${className}`} ref={ref}>
            <button
                type="button"
                onClick={() => setOpen((o) => !o)}
                className={`w-full flex items-center justify-between gap-2 bg-black/50 border border-white/10 rounded-xl
                            text-white outline-none focus:border-cyan-500 focus:ring-1 focus:ring-cyan-500/40
                            ${compact ? 'px-3 py-2 text-xs min-h-[40px] max-w-[180px]' : 'px-3.5 py-3 text-base min-h-[48px]'}`}
                style={{ fontSize: compact ? undefined : 16 }}
                aria-haspopup="listbox"
                aria-expanded={open}
            >
                <span className="truncate">{currentLabel}</span>
                <span className="text-cyan-400/80 text-[10px] shrink-0">{open ? '▲' : '▼'}</span>
            </button>
            {open && (
                <div
                    className="absolute left-0 right-0 top-full mt-1 z-[10050] min-w-[160px] max-h-56 overflow-y-auto
                               rounded-xl border border-white/20 bg-slate-900 shadow-2xl shadow-black/60"
                    role="listbox"
                >
                    {list.map((s) => {
                        const val = s;
                        const selected = String(value) === String(val);
                        return (
                            <button
                                key={String(val)}
                                type="button"
                                role="option"
                                aria-selected={selected}
                                className={`w-full text-left px-4 py-3 min-h-[44px] text-sm
                                            border-b border-white/10 last:border-b-0
                                            flex items-center justify-between gap-2
                                            ${selected ? 'bg-cyan-600/35 text-cyan-100' : 'text-white hover:bg-white/5 active:bg-cyan-500/20'}`}
                                style={{ fontSize: 16 }}
                                onClick={() => {
                                    onChange(val);
                                    setOpen(false);
                                }}
                            >
                                <span className="truncate">{labelFor(val)}</span>
                                <span className={`w-4 h-4 rounded-full border-2 shrink-0 flex items-center justify-center
                                    ${selected ? 'border-cyan-400 bg-cyan-500' : 'border-white/30'}`}>
                                    {selected ? <span className="w-1.5 h-1.5 rounded-full bg-white" /> : null}
                                </span>
                            </button>
                        );
                    })}
                </div>
            )}
        </div>
    );
};


const Input = ({ label, type, value, onChange, placeholder, isRequired = false, options = [] }) => {
    const id = React.useId();
    const [showECPicker, setShowECPicker] = useState(false);
    const [selectOpen, setSelectOpen] = useState(false);
    const [selectMenuBox, setSelectMenuBox] = useState(null);
    const selectRootRef = useRef(null);
    const selectBtnRef = useRef(null);
    useEffect(() => {
        if (!selectOpen) {
            setSelectMenuBox(null);
            return undefined;
        }
        const place = () => {
            const el = selectBtnRef.current || selectRootRef.current;
            if (!el || !el.getBoundingClientRect) return;
            const r = el.getBoundingClientRect();
            const width = Math.max(r.width, 240);
            let left = r.left;
            if (left + width > window.innerWidth - 8) left = Math.max(8, window.innerWidth - width - 8);
            if (left < 8) left = 8;
            const gap = 6;
            const spaceBelow = window.innerHeight - r.bottom - gap;
            const openUp = spaceBelow < 220 && r.top > spaceBelow;
            const maxHeight = Math.min(320, openUp ? Math.max(180, r.top - gap - 8) : Math.max(180, spaceBelow));
            const top = openUp ? Math.max(8, r.top - gap - maxHeight) : r.bottom + gap;
            setSelectMenuBox({ top, left, width, maxHeight, zIndex: 99999 });
        };
        place();
        const onDoc = (e) => {
            if (selectRootRef.current && selectRootRef.current.contains(e.target)) return;
            if (e.target && e.target.closest && e.target.closest('[data-uu-select-menu]')) return;
            setSelectOpen(false);
        };
        document.addEventListener('mousedown', onDoc);
        document.addEventListener('touchstart', onDoc);
        window.addEventListener('resize', place);
        window.addEventListener('scroll', place, true);
        return () => {
            document.removeEventListener('mousedown', onDoc);
            document.removeEventListener('touchstart', onDoc);
            window.removeEventListener('resize', place);
            window.removeEventListener('scroll', place, true);
        };
    }, [selectOpen]);

    
    const ethMonths = [
        "Meskerem", "Tikimt", "Hidar", "Tahsas", "Tir", "Yekatit", 
        "Megabit", "Miazia", "Ginbot", "Sene", "Hamle", "Nehasse", "Pagume"
    ];

    const handleLabelClick = () => {
        const lowerLabel = label?.toLowerCase() || "";
        let targetId = "";
        // Redirection logic to specific 3-input MultiEntry sectors
        if (lowerLabel.includes('experience')) targetId = "new-entry-area-education_entries";
        else if (lowerLabel.includes('beneficiary')) targetId = "new-entry-area-beneficiaries";

        if (targetId) {
            const element = document.getElementById(targetId);
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                element.style.outline = "2px solid #06b6d4";
                setTimeout(() => element.style.outline = "none", 2000);
            }
        }
    };

    const isEC = label?.includes('(E.C.)');
    const isGC = label?.includes('(G.C.)');

    const selectECDate = (day, monthIdx, year) => {
        const formatted = `${String(day).padStart(2, '0')}/${String(monthIdx + 1).padStart(2, '0')}/${year}`;
        onChange(formatted);
        setShowECPicker(false);
    };

    // --- PRIORITY 1: CUSTOM SELECT (page forms + modal — no OS white picker) ---
    if (type === 'select') {
        const normalized = (Array.isArray(options) ? options : []).map((opt) => {
            if (opt && typeof opt === 'object') {
                return {
                    value: String(opt.value ?? opt.id ?? opt.name ?? ''),
                    label: String(opt.label ?? opt.name ?? opt.value ?? opt.id ?? ''),
                };
            }
            return { value: String(opt), label: String(opt) };
        }).filter((o) => o.value !== '' || o.label !== '');

        const current = normalized.find((o) => o.value === String(value ?? '')) || null;

        return (
            <div className="flex flex-col space-y-2 mb-4 group relative" ref={selectRootRef}>
                {label && (
                <label 
                    onClick={handleLabelClick}
                    className="text-[10px] font-black uppercase tracking-[0.2em] text-white/50 group-focus-within:text-cyan-400"
                >
                    {label} {isRequired && <span className="text-red-500">*</span>}
                </label>
                )}
                <button
                    type="button"
                    id={id}
                    ref={selectBtnRef}
                    onClick={() => setSelectOpen((o) => !o)}
                    className="w-full min-h-[48px] flex items-center justify-between gap-2
                               p-3.5 rounded-xl bg-black/60 border border-white/20 text-left
                               text-white text-base focus:ring-2 focus:ring-cyan-500/50 outline-none
                               hover:bg-black/80 transition-all cursor-pointer"
                    style={{ fontSize: 16 }}
                    aria-haspopup="listbox"
                    aria-expanded={selectOpen}
                >
                    <span className={current ? 'text-white truncate' : 'text-white/40'}>
                        {current ? current.label : (placeholder || '-- SELECT --')}
                    </span>
                    <span className="text-cyan-400/80 text-xs shrink-0">{selectOpen ? '▲' : '▼'}</span>
                </button>
                {selectOpen && selectMenuBox && createPortal(
                    <div
                        data-uu-select-menu="1"
                        role="listbox"
                        className="fixed overflow-y-auto rounded-xl border border-cyan-500/30
                                   bg-slate-950 text-white shadow-2xl shadow-black/80 ring-1 ring-white/10"
                        style={{
                            top: selectMenuBox.top,
                            left: selectMenuBox.left,
                            width: selectMenuBox.width,
                            maxHeight: selectMenuBox.maxHeight,
                            zIndex: 99999,
                        }}
                    >
                        <button
                            type="button"
                            className="w-full text-left px-4 py-3.5 min-h-[48px] text-base
                                       border-b border-white/10 text-white/45
                                       hover:bg-white/5 active:bg-cyan-500/20"
                            style={{ fontSize: 16 }}
                            onClick={() => { onChange(''); setSelectOpen(false); }}
                        >
                            {placeholder || '-- SELECT --'}
                        </button>
                        {normalized.map((o) => {
                            const selected = String(value ?? '') === o.value;
                            return (
                                <button
                                    key={o.value}
                                    type="button"
                                    role="option"
                                    aria-selected={selected}
                                    className={`w-full text-left px-4 py-3.5 min-h-[48px] text-base
                                                border-b border-white/10 last:border-b-0
                                                flex items-center justify-between gap-2
                                                ${selected
                                                    ? 'bg-cyan-600/35 text-cyan-100'
                                                    : 'text-white hover:bg-white/5 active:bg-cyan-500/20'}`}
                                    style={{ fontSize: 16 }}
                                    onClick={() => { onChange(o.value); setSelectOpen(false); }}
                                >
                                    <span className="truncate">{o.label}</span>
                                    <span className={`w-5 h-5 rounded-full border-2 shrink-0 flex items-center justify-center
                                        ${selected ? 'border-cyan-400 bg-cyan-500' : 'border-white/30'}`}>
                                        {selected ? <span className="w-2 h-2 rounded-full bg-white" /> : null}
                                    </span>
                                </button>
                            );
                        })}
                    </div>,
                    document.body
                )}
            </div>
        );
    }


    // --- PRIORITY 2: CALENDAR AND TEXT INPUTS ---
    return (
        <div className="flex flex-col space-y-2 mb-4 group relative">
            {label && (
                <label 
                    onClick={handleLabelClick}
                    className={`text-[10px] font-black uppercase tracking-[0.2em] transition-all
                        ${(label.toLowerCase().includes('experience') || label.toLowerCase().includes('beneficiary')) ? 'text-cyan-400 cursor-pointer hover:underline' : 'text-white/50'}`}
                >
                    {label} {isRequired && <span className="text-red-500">*</span>}
                </label>
            )}

            <div className="relative">
                <input
                    id={id}
                    type={isEC ? "text" : (isGC ? "date" : (type || "text"))} autoComplete="off" spellCheck="false"
                    value={value ?? ''}
                    readOnly={isEC} 
                    onClick={() => isEC && setShowECPicker(!showECPicker)}
                    onChange={(e) => !isEC && onChange(e.target.value)}
                    placeholder={isEC ? "DD/MM/YYYY (E.C.)" : placeholder}
                    className={`w-full p-3.5 rounded bg-black/60 border ${isEC ? 'border-cyan-500/50 cursor-pointer' : 'border-white/20'} text-white text-sm focus:ring-2 focus:ring-cyan-500/50 outline-none transition-all`}
                    required={isRequired}
                />
                
                {(isEC || isGC) && (
                    <Icons.Calendar className={`absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 ${isEC ? 'text-cyan-500' : 'text-gray-400'} pointer-events-none`} />
                )}

                {/* ETHIOPIAN CALENDAR PICKER OVERLAY */}
                {isEC && showECPicker && (
                    <div className="absolute top-full left-0 mt-2 p-4 bg-slate-900 border border-cyan-500/50 rounded-xl shadow-2xl z-[999] w-72 backdrop-blur-2xl">
                        <div className="flex justify-between items-center mb-4">
                            <span className="text-[11px] font-black text-white uppercase italic">Power Geez · Ethiopian Calendar</span>
                            <button onClick={() => setShowECPicker(false)}><Icons.X size={16} className="text-red-400"/></button>
                        </div>
                        
                        <div className="space-y-4">
                            <div className="flex gap-2">
                                <input id={`${id}-eth-year`} type="number" defaultValue="2017" className="w-1/2 bg-black border border-white/10 rounded p-2 text-white text-xs" placeholder="Year" />
                                <input id={`${id}-eth-day`} type="number" defaultValue="01" min="1" max="30" className="w-1/2 bg-black border border-white/10 rounded p-2 text-white text-xs" placeholder="Day" />
                            </div>
                            <div className="grid grid-cols-3 gap-1">
                                {ethMonths.map((month, idx) => (
                                    <button
                                        key={month}
                                        type="button"
                                        onClick={() => {
                                            const y = document.getElementById(`${id}-eth-year`).value;
                                            const d = document.getElementById(`${id}-eth-day`).value;
                                            selectECDate(d, idx, y);
                                        }}
                                        className="text-[9px] font-bold p-2 bg-white/5 hover:bg-cyan-600 text-white rounded transition-colors"
                                    >
                                        {month}
                                    </button>
                                ))}
                            </div>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};


// Component for Generic Reports (Only available to Admins/Managers/Staff)
const ReportsGenerator = ({ glassClass }) => {
    const { allData, getCollection } = useData();
    const { currentUser, checkPermission } = useAuth();

    const REPORT_TYPES = [
        { label: 'All Time', value: 'all' },
        { label: 'Daily', value: 'daily' },
        { label: 'Weekly', value: 'weekly' },
        { label: 'Monthly', value: 'monthly' },
        { label: 'Quarterly', value: 'quarterly' },
        { label: 'Annually', value: 'annually' },
    ];

        const REPORT_PRESETS = [
        { id: '', label: 'Custom selection' },
        { id: 'MAT_ITEM', label: 'Material · Inventory / SKU', collection: 'MAT_ITEM_CRUD' },
        { id: 'MAT_GRV', label: 'Material · GRV (Goods Received)', collection: 'MAT_GRV_CRUD' },
        { id: 'MAT_GIV', label: 'Material · GIV (Goods Issued)', collection: 'MAT_GIV_CRUD' },
        { id: 'MAT_PO', label: 'Material · Local Purchase', collection: 'MAT_PURCHASE_CRUD' },
        { id: 'MAT_FOREIGN', label: 'Material · Foreign Purchase', collection: 'MAT_FOREIGN_PURCHASE_CRUD' },
        { id: 'MAT_STORE', label: 'Material · Stores', collection: 'MAT_STORE_CRUD' },
        { id: 'MAT_ROB', label: 'Material · Parts Robbing', collection: 'MAT_PARTS_ROBBING_CRUD' },
        { id: 'MAT_VENDOR', label: 'Material · Vendors', collection: 'MAT_VENDOR_CRUD' },
    ];

    const getCollectionLabel = useCallback((name) => {
        if (!name || typeof name !== 'string') return '';
        return name.replace(/_/g, ' ').replace('CRUD', '(Data)').trim();
    }, []);

    const collectionNames = useMemo(() => Object.keys(allData || {}).sort(), [allData]);
    const initialSelected = collectionNames.length > 0 ? collectionNames[0] : '';

    const [selectedCollection, setSelectedCollection] = useState(initialSelected);
    const [sortKey, setSortKey] = useState('');
    const [groupByKey, setGroupByKey] = useState('');
    const [reportType, setReportType] = useState('all');
    const [dateFrom, setDateFrom] = useState('');
    const [dateTo, setDateTo] = useState('');
    const [preset, setPreset] = useState('');
    const [statusFilter, setStatusFilter] = useState('All');
    const [searchTerm, setSearchTerm] = useState('');

    const columns = useMemo(() => (typeof COLUMN_CONFIG !== 'undefined' ? (COLUMN_CONFIG[selectedCollection] || []) : []), [selectedCollection]);

    const filterableColumns = useMemo(
        () => (columns || []).filter((col) => !col.isHidden && col.type !== 'array'),
        [columns]
    );

    const [columnFilters, setColumnFilters] = useState({});

    useEffect(() => {
        if (selectedCollection) {
            setColumnFilters(
                filterableColumns.reduce((acc, col) => {
                    acc[col.key] = '';
                    return acc;
                }, {})
            );
            setSortKey('');
            setGroupByKey('');
            setReportType('all');
        }
    }, [selectedCollection, filterableColumns]);

    const currentData = useMemo(() => getCollection(selectedCollection) || [], [getCollection, selectedCollection]);

    const getColumnFilterOptions = useCallback((data, columnKey) => {
        const values = new Set(['All']);
        (data || []).forEach((item) => {
            const value = item[columnKey];
            if (value !== null && value !== undefined) values.add(String(value));
        });
        return Array.from(values).map((v) => ({ label: v, value: v }));
    }, []);

    const columnOptions = useMemo(() => {
        const options = {};
        filterableColumns.forEach((col) => {
            options[col.key] = getColumnFilterOptions(currentData, col.key);
        });
        return options;
    }, [currentData, filterableColumns, getColumnFilterOptions]);

    const filteredAndSortedData = useMemo(() => {
        let data = [...currentData];

        data = data.filter((item) =>
            (filterableColumns || []).every((col) => {
                const filterValue = columnFilters[col.key] || '';
                if (filterValue === 'All' || !filterValue) return true;
                return String(item[col.key] || '').toLowerCase() === filterValue.toLowerCase();
            })
        );

        if (statusFilter && statusFilter !== 'All') {
            const sf = statusFilter.toLowerCase();
            data = data.filter((item) => String(item.status || item.Status || '').toLowerCase() === sf);
        }

        const dateKeys = [
            'date', 'posting_date', 'createdAt', 'created_at', 'payroll_process_date',
            'process_date', 'leave_start', 'start_date', 'transaction_date', 'grv_date', 'period',
        ];

        if (dateFrom || dateTo) {
            const fromTs = dateFrom ? new Date(dateFrom).setHours(0, 0, 0, 0) : null;
            const toTs = dateTo ? new Date(dateTo).setHours(23, 59, 59, 999) : null;
            data = data.filter((item) => {
                let raw = null;
                for (const k of dateKeys) {
                    if (item[k] != null && String(item[k]).trim() !== '') {
                        raw = item[k];
                        break;
                    }
                }
                if (!raw) return true;
                const ts = new Date(raw).getTime();
                if (isNaN(ts)) return true;
                if (fromTs != null && ts < fromTs) return false;
                if (toTs != null && ts > toTs) return false;
                return true;
            });
        }

        if (reportType && reportType !== 'all') {
            const now = new Date();
            data = data.filter((item) => {
                let raw = null;
                for (const k of dateKeys) {
                    if (item[k] != null && String(item[k]).trim() !== '') {
                        raw = item[k];
                        break;
                    }
                }
                if (!raw) return true;
                const d = new Date(raw);
                if (isNaN(d.getTime())) return true;
                if (reportType === 'daily') return d.toDateString() === now.toDateString();
                if (reportType === 'weekly') {
                    const weekAgo = new Date(now);
                    weekAgo.setDate(now.getDate() - 7);
                    return d >= weekAgo && d <= now;
                }
                if (reportType === 'monthly') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
                if (reportType === 'quarterly') {
                    const q = Math.floor(now.getMonth() / 3);
                    return Math.floor(d.getMonth() / 3) === q && d.getFullYear() === now.getFullYear();
                }
                if (reportType === 'annually') return d.getFullYear() === now.getFullYear();
                return true;
            });
        }

        if (searchTerm && searchTerm.trim()) {
            const q = searchTerm.trim().toLowerCase();
            data = data.filter((item) =>
                Object.values(item || {}).some((v) => String(v ?? '').toLowerCase().includes(q))
            );
        }

        if (sortKey) {
            data.sort((a, b) => {
                const aVal = a[sortKey];
                const bVal = b[sortKey];
                const na = parseFloat(aVal);
                const nb = parseFloat(bVal);
                if (!isNaN(na) && !isNaN(nb) && isFinite(na) && isFinite(nb)) return na - nb;
                return String(aVal || '').toLowerCase().localeCompare(String(bVal || '').toLowerCase());
            });
        }

        return data;
    }, [currentData, columnFilters, sortKey, filterableColumns, searchTerm, dateFrom, dateTo, reportType, statusFilter]);

    const groupedData = useMemo(() => {
        if (!groupByKey) return { 'All Records': filteredAndSortedData };
        return filteredAndSortedData.reduce((acc, item) => {
            const key = String(item[groupByKey] || 'N/A');
            if (!acc[key]) acc[key] = [];
            acc[key].push(item);
            return acc;
        }, {});
    }, [filteredAndSortedData, groupByKey]);

    const displayColumns = useMemo(
        () => (columns || []).filter((col) => !col.isHidden && col.type !== 'array').slice(0, 8),
        [columns]
    );

    const handleFilterChange = (columnKey, value) => {
        setColumnFilters((prev) => ({ ...prev, [columnKey]: value }));
    };

    useEffect(() => {
        if (!selectedCollection && collectionNames.length > 0) {
            setSelectedCollection(collectionNames[0]);
        }
    }, [collectionNames, selectedCollection]);

    const exportCSV = () => {
        const rows = filteredAndSortedData;
        if (!rows || !rows.length) return;
        const cols = displayColumns.length
            ? displayColumns
            : Object.keys(rows[0] || {}).map((k) => ({ key: k, label: k }));
        const headers = cols.map((c) => c.label || c.key);
        const body = rows.map((r) =>
            cols
                .map((c) => {
                    const v = r[c.key];
                    const s = v == null ? '' : String(v).replace(/"/g, '""');
                    return `"${s}"`;
                })
                .join(',')
        );
        const csv = [headers.join(','), ...body].join('\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${selectedCollection || 'report'}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    };

    const partners = getCollection('FIN_PARTNERS_CRUD') || [];

    return (
        <ContainerCard Icon="FileText" glassClass={glassClass} className="h-full max-w-full min-w-0 overflow-x-hidden overflow-y-auto">
            <div className="space-y-3 mb-4 p-3 sm:p-4 border border-white/10 rounded-xl max-w-full min-w-0 overflow-visible bg-black/20">
                <h3 className="text-base font-semibold text-cyan-400 flex items-center shrink-0">
                    <Icons.FileBarChart className="w-5 h-5 mr-2 shrink-0" />
                    Institutional Report Studio
                </h3>
                <p className="text-[10px] text-gray-500 font-bold uppercase tracking-widest">
                    Presets · date range · column filters · group · export
                </p>

                <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2">
                    <div className="rounded-xl border border-cyan-500/20 bg-cyan-950/20 p-3">
                        <p className="text-[9px] text-cyan-400/80 uppercase font-black">Rows</p>
                        <p className="text-lg font-black text-white">{filteredAndSortedData.length}</p>
                    </div>
                    <div className="rounded-xl border border-white/10 bg-black/30 p-3">
                        <p className="text-[9px] text-gray-500 uppercase font-black">Source</p>
                        <p className="text-xs font-bold text-white truncate">{getCollectionLabel(selectedCollection) || '—'}</p>
                    </div>
                    <div className="rounded-xl border border-white/10 bg-black/30 p-3">
                        <p className="text-[9px] text-gray-500 uppercase font-black">Period</p>
                        <p className="text-xs font-bold text-white">{reportType}</p>
                    </div>
                    <div className="rounded-xl border border-white/10 bg-black/30 p-3">
                        <p className="text-[9px] text-gray-500 uppercase font-black">Groups</p>
                        <p className="text-lg font-black text-white">{Object.keys(groupedData || {}).length}</p>
                    </div>
                </div>

                <div className="w-full max-w-full min-w-0 overflow-visible">
                    <div className="relative z-[50] grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3 items-end min-w-0 w-full">
                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Report preset"
                                type="select"
                                options={REPORT_PRESETS.map((p) => ({ label: p.label, value: p.id }))}
                                value={preset}
                                onChange={(val) => {
                                    setPreset(val);
                                    const hit = REPORT_PRESETS.find((p) => p.id === val);
                                    if (hit && hit.collection && collectionNames.includes(hit.collection)) {
                                        setSelectedCollection(hit.collection);
                                    }
                                }}
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Data Collection"
                                type="select"
                                options={collectionNames.map((name) => ({ label: getCollectionLabel(name), value: name }))}
                                value={selectedCollection}
                                onChange={(val) => setSelectedCollection(val)}
                                isRequired
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <label className="text-[10px] font-black uppercase text-gray-500 tracking-widest">Date from</label>
                            <input
                                type="date"
                                value={dateFrom}
                                onChange={(e) => setDateFrom(e.target.value)}
                                className="w-full mt-1 bg-black/50 border border-white/15 rounded-xl px-3 py-2.5 text-white text-sm outline-none focus:ring-1 focus:ring-cyan-500/40"
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <label className="text-[10px] font-black uppercase text-gray-500 tracking-widest">Date to</label>
                            <input
                                type="date"
                                value={dateTo}
                                onChange={(e) => setDateTo(e.target.value)}
                                className="w-full mt-1 bg-black/50 border border-white/15 rounded-xl px-3 py-2.5 text-white text-sm outline-none focus:ring-1 focus:ring-cyan-500/40"
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Status"
                                type="select"
                                options={[
                                    { label: 'All', value: 'All' },
                                    { label: 'Draft', value: 'Draft' },
                                    { label: 'Posted', value: 'Posted' },
                                    { label: 'Approved', value: 'Approved' },
                                    { label: 'Pending', value: 'Pending' },
                                    { label: 'Active', value: 'Active' },
                                ]}
                                value={statusFilter}
                                onChange={(val) => setStatusFilter(val)}
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Filter by Partner/Entity"
                                type="select"
                                options={[
                                    { label: 'All Entities', value: 'All' },
                                    ...partners.map((p) => ({
                                        label: p.partner_name || p.name || p.id,
                                        value: p.partner_name || p.name || p.id,
                                    })),
                                ]}
                                value={columnFilters.partner_name || ''}
                                onChange={(val) => handleFilterChange('partner_name', val)}
                            />
                        </div>

                        {(filterableColumns || []).map((col) => (
                            <div key={col.key} className="min-w-0 max-w-full w-full overflow-visible">
                                <Input
                                    label={`Filter: ${col.label}`}
                                    type="select"
                                    options={columnOptions[col.key] || [{ label: 'All', value: 'All' }]}
                                    value={columnFilters[col.key] || ''}
                                    onChange={(val) => handleFilterChange(col.key, val)}
                                />
                            </div>
                        ))}

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Report Type"
                                type="select"
                                options={REPORT_TYPES}
                                value={reportType}
                                onChange={(val) => setReportType(val)}
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Sort By"
                                type="select"
                                options={[
                                    { label: 'None', value: '' },
                                    ...(filterableColumns || []).map((c) => ({ label: c.label, value: c.key })),
                                ]}
                                value={sortKey}
                                onChange={(val) => setSortKey(val)}
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-visible">
                            <Input
                                label="Group By"
                                type="select"
                                options={[
                                    { label: 'None', value: '' },
                                    ...(filterableColumns || []).map((c) => ({ label: c.label, value: c.key })),
                                ]}
                                value={groupByKey}
                                onChange={(val) => setGroupByKey(val)}
                            />
                        </div>
                    </div>
                </div>

                <div className="w-full pt-2 flex flex-wrap justify-end gap-2">
                    {typeof ExportButton === 'function' ? (
                        <ExportButton
                            data={filteredAndSortedData}
                            columns={columns}
                            filename={selectedCollection || 'report'}
                            isFinancial
                        />
                    ) : (
                        <button
                            type="button"
                            onClick={exportCSV}
                            className="bg-green-600 text-white px-4 py-2 rounded-lg text-sm hover:bg-green-700 transition font-semibold"
                        >
                            Export CSV
                        </button>
                    )}
                </div>
            </div>

            <div className="mb-4 flex flex-wrap gap-2 items-center">
                <input
                    type="text"
                    placeholder="Search records..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="border border-cyan-500/30 bg-black/50 text-white placeholder:text-gray-500 p-2 rounded-lg text-sm w-full md:w-64 outline-none focus:ring-2 focus:ring-cyan-500/40"
                />
                <button
                    type="button"
                    onClick={exportCSV}
                    className="bg-green-600 text-white px-4 py-2 rounded-lg text-sm hover:bg-green-700 transition font-semibold"
                >
                    Export to CSV
                </button>
            </div>

            {Object.entries(groupedData).map(([groupName, rows]) => (
                <div key={groupName} className="mb-6">
                    {groupByKey ? (
                        <h4 className="text-sm font-black text-cyan-400 uppercase tracking-widest mb-2 px-1">
                            {groupName} · {rows.length}
                        </h4>
                    ) : null}
                    <div className="uu-table-scroll w-full max-w-full min-w-0 overflow-x-auto overflow-y-auto max-h-[min(55vh,520px)] rounded-xl border border-white/10 bg-black/40">
                        <table className="w-max min-w-full border-collapse table-auto">
                            <thead className="bg-slate-950/95 sticky top-0 z-10">
                                <tr className="border-b border-white/10">
                                    {(displayColumns || []).map((col, i) => (
                                        <th
                                            key={i}
                                            className="px-3 py-2.5 text-left text-[10px] font-black uppercase tracking-wider text-cyan-300/90 whitespace-nowrap border-r border-white/5 last:border-r-0"
                                        >
                                            {col.label}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {(rows || []).map((row, ri) => (
                                    <tr key={row.id || ri} className="border-b border-white/5 hover:bg-white/5">
                                        {(displayColumns || []).map((col, j) => {
                                            const raw = row[col.key];
                                            let cell;
                                            if (Array.isArray(raw)) {
                                                cell = (raw || []).slice(0, 2).map((entry, idx) => (
                                                    <span
                                                        key={idx}
                                                        className="bg-gray-700 text-gray-200 rounded-full px-2 py-1 text-xs mr-1 inline-block"
                                                    >
                                                        {typeof entry === 'object' ? Object.values(entry).join(' / ') : String(entry)}
                                                    </span>
                                                ));
                                            } else if (raw === null || raw === undefined || raw === '') {
                                                cell = <span className="text-gray-500">—</span>;
                                            } else if (
                                                col.type === 'number' ||
                                                typeof raw === 'number' ||
                                                (typeof raw === 'string' &&
                                                    raw !== '' &&
                                                    !isNaN(Number(raw)) &&
                                                    /salary|pay|gross|net|amount|tax|pension|basic|ot|allowance/i.test(col.key || ''))
                                            ) {
                                                const n = Number(raw);
                                                cell = Number.isFinite(n)
                                                    ? n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                                                    : String(raw);
                                            } else {
                                                cell = String(raw);
                                            }
                                            return (
                                                <td key={j} className="px-4 py-3.5 whitespace-nowrap text-sm text-white font-medium">
                                                    {cell}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            ))}

            {filteredAndSortedData.length === 0 ? (
                <p className="text-center py-8 text-gray-500">No records match the current criteria.</p>
            ) : null}
        </ContainerCard>
    );
};




/* --- FULLY MODIFIED MULTIENTRYFIELD COMPONENT --- */
const MultiEntryField = ({ label, entries, fieldConfig, onChange, glassClass }) => {
    const safeEntries = Array.isArray(entries) ? entries : [];
    const [newEntry, setNewEntry] = React.useState(
        fieldConfig?.reduce((acc, field) => ({ ...acc, [field.key]: '' }), {}) || {}
    );

    const handleAdd = () => {
        // Validation: ensures all fields except readOnly ones are filled
        const isComplete = fieldConfig.every(field => field.isReadOnly || newEntry[field.key]?.toString().trim());
        
        if (isComplete) {
            onChange([...safeEntries, newEntry]);
            setNewEntry(fieldConfig.reduce((acc, field) => ({ ...acc, [field.key]: '' }), {})); 
        } else {
            alert(`Please complete the entry form below.`);
        }
    };

    return (
        <div className={`p-8 rounded-[2rem] space-y-6 border border-white/10 ${glassClass} mb-8 bg-black/40 shadow-2xl scroll-mt-20`}>
            <label className="block text-[14px] font-black text-cyan-400 border-b border-cyan-500/30 pb-4 uppercase italic tracking-widest">
                {label}
            </label>

            {/* List Table */}
            <div className="space-y-3 max-h-80 overflow-y-auto pr-2 custom-scrollbar">
                {safeEntries.length > 0 ? (
                    safeEntries.map((entry, index) => (
                        <div key={index} className="flex justify-between items-start p-5 rounded-2xl bg-white/5 border border-white/10 hover:bg-cyan-900/20 transition-all">
                            <div className="grid grid-cols-2 lg:grid-cols-4 gap-6 w-full text-[11px]">
                                {fieldConfig.map(f => (
                                    <div key={f.key} className="flex flex-col">
                                        <span className="text-cyan-500 font-black uppercase text-[9px] mb-1">{f.label}:</span>
                                        <span className="text-white font-medium break-words">
                                            {f.key === 'digital_stamp' ? (
                                                <code className="bg-cyan-500/10 px-1 rounded text-[8px] tracking-tighter">{entry[f.key]}</code>
                                            ) : entry[f.key]}
                                        </span>
                                    </div>
                                ))}
                            </div>
                            <button type="button" onClick={() => onChange(safeEntries.filter((_, i) => i !== index))} className="ml-4 text-red-500 hover:text-red-400 p-2 transition-transform active:scale-90">
                                <Icons.Trash2 size={20} />
                            </button>
                        </div>
                    ))
                ) : (
                    <p className="text-[10px] text-white/20 italic p-6 text-center border-2 border-dashed border-white/5 rounded-2xl uppercase tracking-widest font-black">Awaiting entry input...</p>
                )}
            </div>

            {/* SPACIOUS Input Form */}
            <div className="pt-10 border-t border-white/10 space-y-8">
                <p className="text-[10px] font-black text-cyan-400/50 uppercase tracking-[0.3em]">Data Entry Interface:</p>
                <div className="grid grid-cols-1 gap-y-8 bg-black/60 p-8 rounded-[2rem] border border-white/5 shadow-inner shadow-black">
                    {fieldConfig.map(field => {
                        const isSelect = field.type === 'select' || (field.options && field.options.length > 0);
                        
                        return (
                            <div key={field.key} className="flex flex-col space-y-3">
                                <label className="text-[11px] font-black text-gray-400 uppercase ml-2 tracking-widest">
                                    {field.label}
                                </label>
                                {isSelect ? (
                                    <select
                                        value={newEntry[field.key] || ''}
                                        onChange={(e) => setNewEntry(prev => ({ ...prev, [field.key]: e.target.value }))}
                                        className="w-full p-5 rounded-2xl bg-black border border-white/10 text-white text-sm focus:border-cyan-500 outline-none appearance-none"
                                    >
                                        <option value="">Select {field.label}...</option>
                                        {field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                    </select>
                                ) : (
                                    <input
                                        type={field.type || 'text'}
                                        autoComplete="off" 
                                        readOnly={field.isReadOnly}
                                        value={newEntry[field.key] || ''}
                                        onChange={(e) => {
                                            const val = e.target.value;
                                            setNewEntry(prev => {
                                                const updated = { ...prev, [field.key]: val };
                                                // AUTO-STAMP Generation Logic
                                                if (field.key === 'member_name' && val.length > 1) {
                                                    updated.digital_stamp = `AUTH-${val.toUpperCase().substring(0,3)}-${Date.now().toString().slice(-4)}`;
                                                }
                                                return updated;
                                            });
                                        }}
                                        className={`w-full p-5 rounded-2xl bg-black border border-white/10 text-white text-sm focus:border-cyan-500 transition-all ${field.isReadOnly ? 'opacity-40 cursor-not-allowed border-dashed' : ''}`}
                                        placeholder={field.isReadOnly ? 'Securely generated' : `Provide ${field.label}...`}
                                    />
                                )}
                            </div>
                        );
                    })}
                </div>
                <button ref={addBtnRef}
                    type="button"
                    onClick={handleAdd}
                    className="w-full bg-cyan-600 hover:bg-cyan-500 text-white font-black py-6 rounded-3xl text-[12px] uppercase tracking-widest transition-all shadow-2xl shadow-cyan-900/60 active:scale-95"
                >
                    + Append New Record To List
                </button>
            </div>
        </div>
    );
};

// New Component: ContainerCard
const ContainerCard = ({ title, Icon, children, glassClass, className = '' }) => {
    const CardIcon = Icon ? getIcon(Icon) : Icons.File;
    return (
        <div className={`p-5 rounded-xl ${glassClass} shadow-xl backdrop-blur-sm ${className}`}>
            <h2 className="flex items-center text-xl font-bold text-cyan-400 mb-4 border-b border-cyan-400/50 pb-2">
                <CardIcon className="w-6 h-6 mr-2" />
                {title}
            </h2>
            <div className="text-white">
                {children}
            </div>
        </div>
    );
};


const Modal = ({ isOpen, onClose, title, children, glassClass = '', footer, anchorRef }) => {
    const panelRef = React.useRef(null);
    const [box, setBox] = React.useState(null);

    React.useEffect(() => {
        if (!isOpen) return undefined;
        const place = () => {
            try {
                const btn = anchorRef && anchorRef.current ? anchorRef.current : null;
                const vw = window.innerWidth;
                const vh = window.innerHeight;

                // Prefer the table / main content panel — never the left submenu
                let panel = null;
                if (btn && btn.closest) {
                    panel =
                        btn.closest('[data-uu-table-panel]') ||
                        btn.closest('.table-panel') ||
                        btn.closest('main') ||
                        btn.closest('[class*="overflow"]') ||
                        btn.parentElement;
                    // Walk up until we get a reasonably wide content box (skip tight toolbars)
                    let walk = btn.parentElement;
                    for (let i = 0; i < 8 && walk; i++) {
                        const wr = walk.getBoundingClientRect();
                        if (wr.width >= Math.min(480, vw * 0.45) && wr.height >= 120) {
                            panel = walk;
                            break;
                        }
                        walk = walk.parentElement;
                    }
                }

                const pr = panel && panel.getBoundingClientRect
                    ? panel.getBoundingClientRect()
                    : { left: Math.min(280, vw * 0.22), top: 80, right: vw - 12, bottom: vh - 12, width: vw * 0.7, height: vh - 100 };

                // Content column only (exclude typical left nav ~0–260px if panel spans full)
                let contentLeft = pr.left;
                let contentRight = pr.right;
                // If panel is almost full width, inset past submenu
                if (pr.left < 40 && pr.width > vw * 0.85) {
                    contentLeft = Math.min(280, Math.max(200, vw * 0.18));
                    contentRight = vw - 12;
                }

                const contentW = Math.max(280, contentRight - contentLeft);
                const width = Math.min(Math.max(420, contentW * 0.58), contentW - 24, 640);

                let left;
                let top;
                if (btn && btn.getBoundingClientRect) {
                    const r = btn.getBoundingClientRect();
                    // Align toward the Add button within the table content band
                    left = r.right - width;
                    if (left < contentLeft + 8) left = contentLeft + 8;
                    if (left + width > contentRight - 8) left = contentRight - width - 8;

                    const estH = Math.min(vh * 0.88, 820);
                    const below = r.bottom + 8;
                    // Prefer a tall panel in the content column; pin near top of panel if needed
                    if (below + 320 < Math.min(pr.bottom, vh) - 8) {
                        top = Math.min(below, Math.max(pr.top + 12, vh * 0.06));
                    } else {
                        top = Math.max(12, Math.min(r.top, pr.top + 12, vh * 0.06));
                    }
                    // Ensure enough height: if remaining space short, move top up
                    if (Math.min(pr.bottom, vh - 8) - top < 360) {
                        top = Math.max(12, Math.min(pr.bottom, vh - 8) - Math.min(estH, 820));
                    }
                } else {
                    left = contentLeft + (contentW - width) / 2;
                    top = Math.max(pr.top + 24, vh * 0.12);
                }

                const maxBottom = Math.min(pr.bottom, vh - 8);
                const maxHeight = Math.max(360, Math.min(vh * 0.88, maxBottom - top - 8, 820));
                setBox({ top, left, width, maxHeight });
            } catch (e) {
                setBox(null);
            }
        };
        place();
        window.addEventListener('resize', place);
        window.addEventListener('scroll', place, true);
        return () => {
            window.removeEventListener('resize', place);
            window.removeEventListener('scroll', place, true);
        };
    }, [isOpen, anchorRef]);


    if (!isOpen) return null;

    const node = (
        <div
            className="fixed inset-0 z-[99999]"
            style={{ position: 'fixed', inset: 0, zIndex: 99999 }}
            role="presentation"
        >
            {/* Light transparent scrim — page stays clearly visible */}
            <div
                className="absolute inset-0 bg-slate-950/25 backdrop-blur-[2px]"
                style={{ background: 'rgba(2, 8, 20, 0.22)' }}
            />
            <div
                ref={panelRef}
                className={`absolute flex flex-col overflow-hidden rounded-2xl
                            border border-white/30 bg-slate-900/80 backdrop-blur-xl text-white
                            shadow-[0_20px_60px_rgba(0,0,0,0.45)] ring-1 ring-cyan-400/20
                            ${glassClass || ''}`}
                style={box ? {
                    top: box.top,
                    left: box.left,
                    width: box.width,
                    maxHeight: box.maxHeight, minHeight: Math.min(box.maxHeight || 480, 480),
                } : {
                    top: '12vh',
                    left: '50%',
                    transform: 'translateX(-50%)',
                    width: 'min(640px, calc(100vw - 48px))',
                    maxHeight: '78vh',
                }}
                role="dialog"
                aria-modal="true"
                onMouseDown={(e) => e.stopPropagation()}
                onClick={(e) => e.stopPropagation()}
            >
                <div className="h-1 w-full shrink-0 bg-gradient-to-r from-cyan-500/80 via-sky-400/80 to-violet-500/80" />
                <div className="shrink-0 flex items-center justify-between gap-3 px-4 py-2.5 border-b border-white/10 bg-white/5">
                    <h3 className="text-sm sm:text-base font-black uppercase tracking-wide text-cyan-300 truncate">
                        {title}
                    </h3>
                    <button
                        type="button"
                        onClick={onClose}
                        className="p-2 rounded-xl border border-white/15 bg-white/5 min-h-[40px] min-w-[40px]
                                   flex items-center justify-center text-gray-300 hover:text-red-300"
                        aria-label="Close"
                    >
                        <Icons.X className="w-5 h-5" />
                    </button>
                </div>
                <div className="flex-1 min-h-0 overflow-y-auto overflow-x-hidden px-4 py-3 overscroll-contain">
                    {children}
                </div>
                <div className="shrink-0 flex flex-wrap justify-end gap-2 px-4 py-2.5 border-t border-white/10 bg-slate-900/70
                                pb-[max(0.5rem,env(safe-area-inset-bottom))]">
                    {footer}
                    <button
                        type="button"
                        onClick={onClose}
                        className="px-4 py-2.5 rounded-xl text-xs font-bold uppercase border border-white/20
                                   bg-white/5 text-gray-200 min-h-[40px] flex-1 sm:flex-none"
                    >
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );

    try {
        if (typeof document !== 'undefined' && document.body) {
            const portalFn =
                (typeof createPortal === 'function' && createPortal) ||
                (typeof ReactDOM !== 'undefined' && ReactDOM.createPortal);
            if (typeof portalFn === 'function') return portalFn(node, document.body);
        }
    } catch (e) {}
    return node;
};

// New Component: Export Helpers
const exportTableWorkbook = async (data, columns, filename, options = {}) => {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Unity University SIS';
    wb.created = new Date();
    const sheetName = String(options.sheetName || filename || 'Data').replace(/[\\/*?:\[\]]/g, ' ').slice(0, 31) || 'Data';
    const ws = wb.addWorksheet(sheetName, {
        views: [{ state: 'frozen', ySplit: 4, showGridLines: false }],
        pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, paperSize: 9 }
    });
    const cols = (columns || []).filter(c => c && !c.isHidden && c.type !== 'array' && c.key);
    if (!cols.length && data && data[0]) {
        Object.keys(data[0]).forEach(k => { if (k !== 'id') cols.push({ key: k, label: k.replace(/_/g, ' ') }); });
    }
    const isMoneyKey = (key, type) =>
        type === 'number' || /amount|salary|pay|gross|net|tax|cost|price|total_|fee|value|qty|quantity|duty|freight|vat|landed/i.test(String(key || ''));
    const borderThin = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };
    const fillNavy = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
    const fillAlt = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    const fillTitle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    const lastCol = Math.max(cols.length, 1);
    ws.mergeCells(1, 1, 1, lastCol);
    const t = ws.getCell(1, 1);
    t.value = 'UNITY UNIVERSITY';
    t.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    t.fill = fillTitle;
    t.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 26;
    ws.mergeCells(2, 1, 2, lastCol);
    const s = ws.getCell(2, 1);
    s.value = String(filename || 'Export').replace(/_/g, ' ').toUpperCase() + ' (Official Record)';
    s.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF1E3A5F' } };
    s.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(2).height = 20;
    ws.mergeCells(3, 1, 3, lastCol);
    const meta = ws.getCell(3, 1);
    meta.value = `Material Management · Columns: ${cols.map(c => c.label || c.key).join(' · ')} · Generated ${new Date().toLocaleString()} · ${(data || []).length} record(s)`;
    meta.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF64748B' } };
    meta.alignment = { horizontal: 'center', wrapText: true };
    ws.getRow(3).height = 28;
    const headerRow = ws.getRow(4);
    cols.forEach((col, i) => {
        const cell = headerRow.getCell(i + 1);
        cell.value = col.label || col.key;
        cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = fillNavy;
        cell.border = borderThin;
        cell.alignment = { horizontal: isMoneyKey(col.key, col.type) ? 'right' : 'left', vertical: 'middle', wrapText: true };
        const w = Math.min(42, Math.max(12, String(col.label || col.key).length + 4));
        ws.getColumn(i + 1).width = isMoneyKey(col.key, col.type) ? Math.max(w, 14) : w;
    });
    headerRow.height = 22;
    const normVal = (col, raw) => {
        if (raw == null || raw === '') return '';
        if (Array.isArray(raw)) return raw.map(entry => (typeof entry === 'object' ? Object.values(entry).join(' / ') : String(entry))).join('; ');
        if (col.type === 'date' || (typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}/.test(String(raw)))) {
            try { const d = new Date(raw); if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10); } catch (e) {}
        }
        if (isMoneyKey(col.key, col.type) || typeof raw === 'number') {
            const num = Number(raw);
            return Number.isFinite(num) ? num : String(raw);
        }
        return String(raw);
    };
    (data || []).forEach((row, ri) => {
        const excelRow = ws.getRow(5 + ri);
        cols.forEach((col, ci) => {
            const cell = excelRow.getCell(ci + 1);
            const val = normVal(col, row[col.key]);
            cell.value = val;
            cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF1E293B' } };
            cell.border = borderThin;
            if (ri % 2 === 1) cell.fill = fillAlt;
            if (typeof val === 'number') {
                cell.numFmt = '#,##0.00';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
                cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
            }
        });
    });
    const footR = 5 + (data || []).length + 1;
    ws.mergeCells(footR, 1, footR, lastCol);
    ws.getCell(footR, 1).value = 'Unity University · Official system export · Columns match on-screen table · Confidential';
    ws.getCell(footR, 1).font = { name: 'Calibri', size: 8, italic: true, color: { argb: 'FF94A3B8' } };
    const buf = await wb.xlsx.writeBuffer();
    const safe = String(filename || 'Unity_Export').replace(/\s+/g, '_').replace(/[^\w\-]+/g, '');
    saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safe}_${new Date().toISOString().slice(0, 10)}.xlsx`);
};

const convertToExcel = (data, columns, filename) => {
    exportTableWorkbook(data, columns, filename, { sheetName: String(filename || 'Data').slice(0, 31) });
};

const convertToPDF = (data, columns, filename) => {
    const cols = (columns || []).filter(c => c && !c.isHidden && c.type !== 'array');
    const rows = Array.isArray(data) ? data : [];
    const doc = new jsPDF({ orientation: cols.length > 6 ? 'landscape' : 'portrait', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, pageW, 26, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.text('UNITY UNIVERSITY', 14, 12);
    doc.setFontSize(10);
    doc.text(String(filename || 'Export').toUpperCase() + ' (Official Record)', 14, 19);
    doc.setTextColor(60);
    doc.setFontSize(8);
    doc.text(`Generated ${new Date().toLocaleString()} · ${rows.length} records`, 14, 32);
    doc.autoTable({
        head: [cols.map(c => c.label || c.key)],
        body: rows.map(r => cols.map(c => {
            const v = r[c.key];
            if (v == null) return '';
            if (typeof v === 'number') return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            return String(v);
        })),
        startY: 36,
        theme: 'grid',
        styles: { fontSize: 7, cellPadding: 1.5 },
        headStyles: { fillColor: [30, 58, 95], textColor: 255, fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [241, 245, 249] },
        margin: { left: 10, right: 10 }
    });
    doc.save(`${String(filename || 'Export').replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.pdf`);
};

const convertToDOCX = async (data, columns, filename) => {
    const cols = (columns || []).filter(c => c && !c.isHidden && c.type !== 'array');
    const rows = Array.isArray(data) ? data : [];
    const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, TextRun, ShadingType } = docx;
    const border = {
        top: { style: docx.BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
        bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
        left: { style: docx.BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
        right: { style: docx.BorderStyle.SINGLE, size: 4, color: 'CBD5E1' },
    };
    const headerRow = new TableRow({
        children: cols.map(c => new TableCell({
            borders: border,
            children: [new Paragraph({ children: [new TextRun({ text: String(c.label || c.key).toUpperCase(), bold: true, color: 'FFFFFF', size: 16 })] })],
            shading: { fill: '1E3A5F', type: ShadingType.CLEAR }
        })),
        tableHeader: true
    });
    const dataRows = rows.map((r, i) => new TableRow({
        children: cols.map(c => {
            let v = r[c.key];
            if (v == null) v = '';
            else if (typeof v === 'number') v = v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            else v = String(v);
            return new TableCell({
                borders: border,
                children: [new Paragraph({ children: [new TextRun({ text: v, size: 16 })] })],
                shading: i % 2 === 0 ? { fill: 'F8FAFC', type: ShadingType.CLEAR } : undefined
            });
        })
    }));
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ children: [new TextRun({ text: 'UNITY UNIVERSITY', bold: true, size: 32, color: '1E3A5F' })] }),
                new Paragraph({ children: [new TextRun({ text: String(filename || 'Export').toUpperCase() + ' (Official Record)', bold: true, size: 22 })], spacing: { after: 120 } }),
                new Paragraph({ children: [new TextRun({ text: `${rows.length} records · ${new Date().toLocaleString()}`, size: 16, italics: true, color: '64748B' })], spacing: { after: 200 } }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] }),
                new Paragraph({ children: [new TextRun({ text: 'Unity University · Official system export · Confidential', size: 14, italics: true, color: '94A3B8' })], spacing: { before: 200 } })
            ]
        }]
    });
    saveAs(await Packer.toBlob(doc), `${String(filename || 'Export').replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.docx`);
};

const ExportButton = ({ data, columns, filename, chartRef, isFinancial = false }) => {
    const [isOpen, setIsOpen] = useState(false);
    const getFilename = () => String(filename || 'Unity_Export').replace(/\s+/g, '_');
    const visibleCols = () => (columns || []).filter(c => c && !c.isHidden && c.type !== 'array');

    const handleAction = async (type) => {
        setIsOpen(false);
        const fname = getFilename();
        const rows = Array.isArray(data) ? data : [];
        const cols = visibleCols();
        try {
            if (type === 'PNG') {
                if (!chartRef?.current) return alert('No capture area for PNG.');
                const canvas = await html2canvas(chartRef.current, { backgroundColor: '#020617', scale: 3 });
                saveAs(canvas.toDataURL('image/png'), `${fname}_${new Date().toISOString().slice(0, 10)}.png`);
                return;
            }
            if (type === 'XLSX') {
                await exportTableWorkbook(rows, cols.length ? cols : (columns || []), filename || fname, {
                    sheetName: String(filename || 'Data').slice(0, 31)
                });
                return;
            }
            if (type === 'PDF') {
                convertToPDF(rows, cols, filename || fname);
                return;
            }
            if (type === 'DOCX') {
                await convertToDOCX(rows, cols, filename || fname);
                return;
            }
            if (type === 'CSV') {
                const esc = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
                const lines = [
                    'UNITY UNIVERSITY',
                    String(filename || 'Export').toUpperCase() + ' (Official Record)',
                    `Generated ${new Date().toLocaleString()} · ${rows.length} records`,
                    '',
                    cols.map(c => esc(c.label || c.key)).join(','),
                    ...rows.map(r => cols.map(c => esc(r[c.key])).join(',')),
                    '',
                    'Unity University · Official system export · Confidential'
                ];
                saveAs(new Blob(['\ufeff' + lines.join(String.fromCharCode(10))], { type: 'text/csv;charset=utf-8' }), `${fname}_${new Date().toISOString().slice(0, 10)}.csv`);
            }
        } catch (err) {
            console.error('[EXPORT]', err);
            alert('Export failed: ' + (err && err.message ? err.message : String(err)));
        }
    };

    return (
        <div className="relative inline-block text-left z-30">
            <button type="button" onClick={() => setIsOpen(!isOpen)} className="bg-cyan-600 hover:bg-cyan-500 text-white font-bold py-2.5 px-6 rounded-xl flex items-center gap-2 shadow-lg transition-all active:scale-95 text-[11px] uppercase tracking-widest border border-white/20">
                <Icons.Download size={16}/> Export Options
            </button>
            {isOpen && (
                <div className="absolute right-0 top-full mt-2 w-64 rounded-2xl bg-slate-900 border border-white/10 shadow-2xl overflow-hidden backdrop-blur-xl z-[10001]">
                    <div className="p-2 grid grid-cols-1 gap-1">
                        <button type="button" onClick={() => handleAction('XLSX')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-emerald-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileSpreadsheet size={16}/> Excel Spreadsheet</button>
                        <button type="button" onClick={() => handleAction('DOCX')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-blue-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileBox size={16}/> Word Document</button>
                        <button type="button" onClick={() => handleAction('PDF')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-rose-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileText size={16}/> PDF Document</button>
                        <button type="button" onClick={() => handleAction('CSV')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-gray-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.Filter size={16}/> CSV (Official)</button>
                        <hr className="border-white/5 mx-2 my-1" />
                        <button type="button" onClick={() => handleAction('PNG')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-cyan-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.Camera size={16}/> Image Capture (PNG)</button>
                    </div>
                </div>
            )}
        </div>
    );
};

const DataTable = ({ collectionName, data, columns, glassClass, title, getCollection }) => {
    const { deleteRecord, updateRecord, postMaterialToFinance, postAllMaterialToFinance } = useData();
    const { checkPermission } = useAuth();

    const [isModalOpen, setIsModalOpen] = useState(false);
    const addBtnRef = useRef(null);
    const [editingRecord, setEditingRecord] = useState(null);
    const [searchTerm, setSearchTerm] = useState('');
    const [statusFilter, setStatusFilter] = useState('ALL');
    const [locationFilter, setLocationFilter] = useState('ALL');
    const [campusFilter, setCampusFilter] = useState('ALL');
    const tableWrapRef = useRef(null);

    const allRoles = useMemo(() => getCollection('ROLE_CRUD') || [], [getCollection]);
    const dataMemo = useMemo(() => getCollection(collectionName) || [], [getCollection, collectionName]);

    const getAddLabel = () => {
        const labels = {
            'MAT_ITEM_CRUD': 'Inventory Item',
            'MAT_GRV_CRUD': 'GRV',
            'MAT_GIV_CRUD': 'GIV',
            'MAT_ROBBING_CRUD': 'Robbing Voucher',
            'MAT_PURCHASE_LOCAL_CRUD': 'Local Purchase',
            'MAT_PURCHASE_FOREIGN_CRUD': 'Foreign Purchase',
            'DEPARTMENTS_HIRING_REQUESTS_CRUD': 'Hiring Request',
            'HR_EMPLOYEE_CRUD': 'Employee',
            'HR_APPRAISAL_CRUD': 'Appraisal',
            'HR_LEAVE_CRUD': 'Leave Request',
            'HR_RECRUITMENT_CRUD': 'Candidate',
            'HR_COMPETENCE_CRUD': 'Competence',
            'HR_PENALTY_CRUD': 'Penalty',
            'HR_AWARD_CRUD_HR': 'Award',
            'HR_TRANSFER_CRUD': 'Transfer',
            'DEPARTMENTS_CRUD': 'Department',
            'ROLE_CRUD': 'Role',
            'HR_INSURANCE_CRUD': 'Policy',
            'RETIREMENT_CRUD': 'Retirement',
            'RESIGNATION_CRUD': 'Resignation',
            'DEATH_CRUD': 'Incident',
        };
        return labels[collectionName] || 'Record';
    };

    const moduleTitle = title || {
        'MAT_ITEM_CRUD': 'Inventory Master',
        'MAT_GRV_CRUD': 'Goods Receiving (GRV)',
        'MAT_GIV_CRUD': 'Goods Issuing (GIV)',
        'MAT_ROBBING_CRUD': 'Parts Robbing (Cannibalization)',
        'MAT_PURCHASE_LOCAL_CRUD': 'Local Purchase',
        'MAT_PURCHASE_FOREIGN_CRUD': 'Foreign Purchase',
    }[collectionName] || `${getAddLabel()} Register`;

    const visibleColumns = useMemo(
        () => columns.filter(c => !c.isHidden && c.type !== 'array' && c.key !== 'actions'),
        [columns]
    );

    const statusOptions = useMemo(() => {
        const set = new Set();
        dataMemo.forEach(r => { if (r.status) set.add(String(r.status)); });
        return ['ALL', ...Array.from(set).sort()];
    }, [dataMemo]);

    const locationOptions = useMemo(() => {
        const set = new Set();
        dataMemo.forEach(r => {
            const loc = r.location || r.store_location || r.donor_location || r.recipient_location;
            if (loc) set.add(String(loc));
        });
        return ['ALL', ...Array.from(set).sort()];
    }, [dataMemo]);

    const campusOptions = useMemo(() => {
        const set = new Set();
        dataMemo.forEach(r => { if (r.campus) set.add(String(r.campus)); });
        return ['ALL', ...Array.from(set).sort()];
    }, [dataMemo]);

    const filteredData = useMemo(() => {
        const q = (searchTerm || '').trim().toLowerCase();
        return dataMemo
            .filter(item => {
                if (statusFilter !== 'ALL' && String(item.status || '') !== statusFilter) return false;
                if (locationFilter !== 'ALL') {
                    const loc = String(item.location || item.store_location || item.donor_location || item.recipient_location || '');
                    if (loc !== locationFilter) return false;
                }
                if (campusFilter !== 'ALL' && String(item.campus || '') !== campusFilter) return false;
                if (!q) return true;
                return visibleColumns.some(col => String(item[col.key] ?? '').toLowerCase().includes(q));
            })
            .sort((a, b) => String(b.id || '').localeCompare(String(a.id || '')));
    }, [dataMemo, searchTerm, visibleColumns, statusFilter, locationFilter, campusFilter]);

    /* Summary cards for current module */
    const cards = useMemo(() => {
        const rows = filteredData;
        const all = dataMemo;
        const n = (v) => Number(v) || 0;
        const sum = (arr, key) => arr.reduce((s, r) => s + n(r[key]), 0);

        if (collectionName === 'MAT_ITEM_CRUD') {
            const low = rows.filter(r => n(r.quantity) > 0 && n(r.quantity) <= (n(r.reorder_level) || 15)).length;
            const oos = rows.filter(r => n(r.quantity) <= 0).length;
            const val = sum(rows, 'total_value') || rows.reduce((s, r) => s + n(r.quantity) * n(r.unit_cost), 0);
            return [
                { label: 'SKUs (filtered)', value: rows.length, color: 'text-cyan-400', tip: 'Items matching filters' },
                { label: 'Low stock', value: low, color: 'text-amber-400', tip: 'At or below reorder level' },
                { label: 'Out of stock', value: oos, color: 'text-rose-400', tip: 'Zero quantity' },
                { label: 'Stock value (ETB)', value: val.toLocaleString(undefined, { maximumFractionDigits: 0 }), color: 'text-emerald-400', tip: 'Qty × unit cost' },
            ];
        }
        if (collectionName === 'MAT_GRV_CRUD') {
            return [
                { label: 'GRV lines', value: rows.length, color: 'text-cyan-400', tip: 'Receiving documents' },
                { label: 'Qty received', value: sum(rows, 'qty_received') || sum(rows, 'quantity'), color: 'text-violet-400', tip: 'Units inbound' },
                { label: 'Line value (ETB)', value: (sum(rows, 'line_value') || 0).toLocaleString(), color: 'text-emerald-400', tip: 'Received value' },
                { label: 'Pending / inspect', value: rows.filter(r => /pending|partial|inspect/i.test(String(r.status || ''))).length, color: 'text-amber-400', tip: 'Not fully completed' },
            ];
        }
        if (collectionName === 'MAT_GIV_CRUD') {
            return [
                { label: 'GIV lines', value: rows.length, color: 'text-cyan-400', tip: 'Issue documents' },
                { label: 'Qty issued', value: sum(rows, 'quantity'), color: 'text-violet-400', tip: 'Units to departments' },
                { label: 'Issue value (ETB)', value: (sum(rows, 'line_value') || 0).toLocaleString(), color: 'text-emerald-400', tip: 'Issued value' },
                { label: 'Departments', value: new Set(rows.map(r => r.department).filter(Boolean)).size, color: 'text-sky-400', tip: 'Distinct requesting depts' },
            ];
        }
        if (collectionName === 'MAT_ROBBING_CRUD') {
            const restored = rows.filter(r => /restored/i.test(String(r.recipient_status_after || r.status || ''))).length;
            const scrapped = rows.filter(r => /scrap/i.test(String(r.donor_status_after || r.status || ''))).length;
            return [
                { label: 'Robbing vouchers', value: rows.length, color: 'text-cyan-400', tip: 'Cannibalization records' },
                { label: 'Recipients restored', value: restored, color: 'text-emerald-400', tip: 'Recipient brought back to service' },
                { label: 'Donors scrapped', value: scrapped, color: 'text-rose-400', tip: 'Donor taken out of service' },
                { label: 'Parts moved', value: sum(rows, 'qty_parts') || rows.length, color: 'text-amber-400', tip: 'Components transferred' },
            ];
        }
        if (collectionName === 'MAT_PURCHASE_LOCAL_CRUD') {
            return [
                { label: 'LPOs', value: rows.length, color: 'text-cyan-400', tip: 'Local purchase orders' },
                { label: 'Order value (ETB)', value: sum(rows, 'total_value').toLocaleString(), color: 'text-emerald-400', tip: 'Sum of totals' },
                { label: 'VAT (ETB)', value: sum(rows, 'vat_amount').toLocaleString(), color: 'text-violet-400', tip: 'VAT amount' },
                { label: 'Open / transit', value: rows.filter(r => /order|transit|pending|partial/i.test(String(r.status || ''))).length, color: 'text-amber-400', tip: 'Not fully received' },
            ];
        }
        if (collectionName === 'MAT_PURCHASE_FOREIGN_CRUD') {
            return [
                { label: 'Import lines', value: rows.length, color: 'text-cyan-400', tip: 'Foreign PIs' },
                { label: 'Landed cost (ETB)', value: (sum(rows, 'total_landed') || sum(rows, 'customs_duty') + sum(rows, 'shipping_cost')).toLocaleString(), color: 'text-emerald-400', tip: 'FOB + freight + duty…' },
                { label: 'Customs duty', value: sum(rows, 'customs_duty').toLocaleString(), color: 'text-violet-400', tip: 'Duty portion' },
                { label: 'In transit', value: rows.filter(r => /transit|order|pending/i.test(String(r.status || ''))).length, color: 'text-amber-400', tip: 'Not yet received' },
            ];
        }
        return [
            { label: 'Records', value: rows.length, color: 'text-cyan-400', tip: 'Filtered rows' },
            { label: 'Total in module', value: all.length, color: 'text-gray-300', tip: 'Unfiltered count' },
            { label: 'Statuses', value: new Set(all.map(r => r.status).filter(Boolean)).size, color: 'text-violet-400', tip: 'Distinct status values' },
            { label: 'Showing', value: `${rows.length}/${all.length}`, color: 'text-sky-400', tip: 'Filter coverage' },
        ];
    }, [collectionName, filteredData, dataMemo]);

    const handleAdd = () => { setEditingRecord({}); setIsModalOpen(true); };
    const handleEdit = (record) => { setEditingRecord(record); setIsModalOpen(true); };

    const exportFilename = moduleTitle.replace(/[^\w\s\-]/g, '').replace(/\s+/g, '_');

    return (
        <div className={`rounded-[1.75rem] border border-white/10 ${glassClass} flex flex-col w-full min-h-0`}>
            {/* Title + cards + filters — fixed height, no horizontal scroll */}
            <div className="shrink-0 z-30 border-b border-white/10 bg-slate-950/95 backdrop-blur-xl px-4 sm:px-5 pt-4 pb-3">
                <div className="border-l-4 border-cyan-500 pl-4 mb-3">
                    <h1 className="text-xl sm:text-2xl md:text-3xl font-black text-white uppercase italic tracking-tighter leading-tight">
                        {moduleTitle}
                    </h1>
                    <p className="text-[10px] text-cyan-400 font-black uppercase tracking-[0.2em] mt-1 opacity-90">
                        Unity University · Material Management · Official register
                    </p>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-4 gap-2 sm:gap-3 mb-3">
                    {cards.map((c, i) => (
                        <div key={i} className="rounded-2xl border border-white/10 bg-black/40 px-3 py-2.5 min-w-0">
                            <p className="text-[9px] font-black uppercase tracking-widest text-gray-500 truncate">{c.label}</p>
                            <p className={`text-lg sm:text-xl font-black mt-0.5 truncate ${c.color}`}>{c.value}</p>
                            <p className="text-[9px] text-gray-600 mt-0.5 truncate">{c.tip}</p>
                        </div>
                    ))}
                </div>

                <div className="flex flex-col xl:flex-row flex-wrap gap-2 items-stretch xl:items-center justify-between">
                    <div className="flex flex-wrap gap-2 items-center flex-1 min-w-0">
                        <div className="relative min-w-[160px] flex-1 max-w-sm">
                            <Icons.Search className="w-4 h-4 text-gray-500 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                            <input
                                type="text"
                                placeholder={`Search ${getAddLabel()}s…`}
                                value={searchTerm}
                                onChange={e => setSearchTerm(e.target.value)}
                                className="w-full pl-9 pr-3 py-2 rounded-xl bg-black/50 border border-white/10 text-xs text-white outline-none focus:border-cyan-500"
                            />
                        </div>
                        {statusOptions.length > 1 && (
                            <UUSelect
                                value={statusFilter}
                                onChange={setStatusFilter}
                                options={statusOptions}
                                allLabel="All statuses"
                                compact
                                className="max-w-[160px]"
                            />
                        )}
                        {locationOptions.length > 1 && (
                            <UUSelect
                                value={locationFilter}
                                onChange={setLocationFilter}
                                options={locationOptions}
                                allLabel="All locations"
                                compact
                                className="max-w-[180px]"
                            />
                        )}
                        {campusOptions.length > 1 && (
                            <UUSelect
                                value={campusFilter}
                                onChange={setCampusFilter}
                                options={campusOptions}
                                allLabel="All campuses"
                                compact
                                className="max-w-[160px]"
                            />
                        )}
                        {(searchTerm || statusFilter !== 'ALL' || locationFilter !== 'ALL' || campusFilter !== 'ALL') && (
                            <button type="button" onClick={() => { setSearchTerm(''); setStatusFilter('ALL'); setLocationFilter('ALL'); setCampusFilter('ALL'); }}
                                className="text-[10px] font-black uppercase text-rose-400 hover:text-rose-300 px-2 whitespace-nowrap">
                                Clear filters
                            </button>
                        )}
                    </div>
                    <div className="flex flex-wrap gap-2 items-center shrink-0">
                        {COLLECTION_SOURCE_MAP[collectionName] && (
                            <button
                                type="button"
                                onClick={() => postAllMaterialToFinance?.(collectionName)}
                                className="bg-emerald-700 hover:bg-emerald-600 text-white font-bold py-2 px-3 rounded-xl flex items-center gap-2 text-[10px] uppercase tracking-widest border border-white/20"
                                title="Post unposted rows to Finance"
                            >
                                <Icons.Landmark size={14} /> Post All → Finance
                            </button>
                        )}
                        <ExportButton data={filteredData} columns={columns} filename={exportFilename} />
                        {checkPermission('CAN_ADD_RECORD', allRoles) && (
                            <button ref={addBtnRef} type="button" onClick={handleAdd}
                                className="bg-cyan-600 hover:bg-cyan-500 text-white font-bold py-2 px-4 rounded-xl flex items-center gap-2 text-[10px] uppercase tracking-widest border border-white/20">
                                <Icons.Plus size={14} /> Add {getAddLabel()}
                            </button>
                        )}
                    </div>
                </div>
            </div>

            {/* ONE scrollport: horizontal + vertical move header + body together */}
            <div
                ref={tableWrapRef}
                className="w-full overflow-x-auto overflow-y-visible custom-scrollbar relative border border-white/15 rounded-xl"
            >
                <table className="w-max min-w-full border-collapse text-left border border-white/20">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-slate-900 shadow-md shadow-black/40">
                            {visibleColumns.map(col => (
                                <th
                                    key={col.key}
                                    className="px-3 py-2.5 text-[9px] font-black uppercase tracking-widest text-cyan-300/95 border border-white/15 border-b-cyan-500/40 whitespace-nowrap bg-slate-900"
                                    style={{ minWidth: col.type === 'textarea' ? 180 : (col.type === 'number' ? 110 : 140), maxWidth: 280 }}
                                >
                                    {col.label}
                                </th>
                            ))}
                            <th className="px-3 py-2.5 text-[9px] font-black uppercase tracking-widest text-cyan-300/95 border-b border-cyan-500/30 whitespace-nowrap bg-slate-900 sticky right-0 z-30 shadow-[-8px_0_12px_rgba(0,0,0,0.35)]" style={{ minWidth: 108 }}>
                                Actions
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        {filteredData.length > 0 ? (
                            filteredData.map(item => (
                                <tr
                                    key={item.id}
                                    onClick={() => handleEdit(item)}
                                    className="border-b border-white/15 hover:bg-cyan-500/10 cursor-pointer transition-colors group"
                                    title="Click row to open record"
                                >
                                    {visibleColumns.map(col => (
                                        <td
                                            key={col.key}
                                            className="px-3 py-2 text-[11px] text-gray-200 border border-white/10 group-hover:text-white"
                                            style={{ minWidth: col.type === 'textarea' ? 180 : (col.type === 'number' ? 110 : 140), maxWidth: 280 }}
                                        >
                                            <div className="truncate" title={item[col.key] != null && item[col.key] !== '' ? String(item[col.key]) : ''}>
                                                {item[col.key] !== undefined && item[col.key] !== null && item[col.key] !== ''
                                                    ? String(item[col.key])
                                                    : <span className="text-white/15 italic">—</span>}
                                            </div>
                                        </td>
                                    ))}
                                    <td
                                        className="px-2 py-2 bg-slate-950/80 sticky right-0 z-10 shadow-[-8px_0_12px_rgba(0,0,0,0.25)] group-hover:bg-slate-900/95"
                                        onClick={e => e.stopPropagation()}
                                        style={{ minWidth: 108 }}
                                    >
                                        <div className="flex justify-center items-center gap-1.5">
                                            {COLLECTION_SOURCE_MAP[collectionName] && !item.posted_to_finance && (
                                                <button
                                                    type="button"
                                                    onClick={() => postMaterialToFinance?.(COLLECTION_SOURCE_MAP[collectionName], item.id)}
                                                    className="text-emerald-400 hover:text-emerald-200 p-1"
                                                    title="Post this document to Finance"
                                                >
                                                    <Icons.Landmark size={15} />
                                                </button>
                                            )}
                                            {item.posted_to_finance && (
                                                <span className="text-[9px] font-black uppercase text-emerald-500/80" title={item.posted_ref || 'Posted'}>Posted</span>
                                            )}
                                            {checkPermission('CAN_EDIT_RECORD', allRoles) && (
                                                <button type="button" onClick={() => handleEdit(item)} className="text-cyan-400 hover:text-cyan-200 p-1" title="Edit">
                                                    <Icons.Edit3 size={15} />
                                                </button>
                                            )}
                                            {checkPermission('CAN_DELETE_RECORD', allRoles) && (
                                                <button type="button" onClick={() => deleteRecord(collectionName, item.id)} className="text-red-500 hover:text-red-400 p-1" title="Delete">
                                                    <Icons.Trash2 size={15} />
                                                </button>
                                            )}
                                        </div>
                                    </td>
                                </tr>
                            ))
                        ) : (
                            <tr>
                                <td colSpan={visibleColumns.length + 1} className="p-16 text-center text-gray-600 text-[10px] font-black uppercase tracking-widest">
                                    No {getAddLabel()}s match the current filters
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>

            <Modal anchorRef={addBtnRef} isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={editingRecord && editingRecord.id ? `Edit ${getAddLabel()}` : `New ${getAddLabel()}`}>
                <FormModal
                    collectionName={collectionName}
                    initialData={editingRecord}
                    onSave={() => setIsModalOpen(false)}
                    onCancel={() => setIsModalOpen(false)}
                />
            </Modal>
        </div>
    );
};


const COLUMN_CONFIG = {
    'MAT_ITEM_CRUD': [
        { key: 'sku', label: 'SKU / Item Code', isRequired: true },
        { key: 'name', label: 'Item / Asset Name', isRequired: true },
        { key: 'category', label: 'Category', type: 'select', options: MAT_CATEGORIES, isRequired: true },
        { key: 'item_type', label: 'Stock Type', type: 'select', options: ['Consumable', 'Spare Part', 'Serialized Asset', 'Module/Book', 'Raw Material', 'Finished Good'] },
        { key: 'unit', label: 'Unit', type: 'select', options: ITEM_UNITS, isRequired: true },
        { key: 'quantity', label: 'In-Hand Qty', type: 'number', isRequired: true },
        { key: 'reorder_level', label: 'Reorder Level', type: 'number' },
        { key: 'unit_cost', label: 'Unit Cost (ETB)', type: 'number', isRequired: true },
        { key: 'total_value', label: 'Stock Value (ETB)', type: 'number', isReadOnly: true },
        { key: 'location', label: 'Store Location', type: 'select', options: MAT_LOCATIONS, isRequired: true },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'coa_account', label: 'GL / COA Account', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'asset_tag', label: 'Asset Tag / Serial' },
        { key: 'condition', label: 'Condition', type: 'select', options: ['New', 'Good', 'Fair', 'Poor', 'Non-Functional', 'Cannibalized (Donor)'] },
        { key: 'status', label: 'Status', type: 'select', options: ['Active', 'Low Stock', 'Out of Stock', 'Obsolete', 'Quarantine', 'In Transit', 'Reserved for Robbing'] },
        { key: 'preferred_supplier', label: 'Preferred Supplier' },
        { key: 'last_updated', label: 'Last Count Date', type: 'date' },
        { key: 'remarks', label: 'Remarks / Spec', type: 'textarea' },
    ],
    'MAT_GRV_CRUD': [
        { key: 'grv_no', label: 'GRV Number', isRequired: true },
        { key: 'date', label: 'GRV Date', type: 'date', isRequired: true },
        { key: 'po_number', label: 'PO / LPO Ref' },
        { key: 'purchase_type', label: 'Source', type: 'select', options: ['Local Purchase', 'Foreign Purchase', 'Donation', 'Transfer In', 'Return from Dept'] },
        { key: 'supplier', label: 'Vendor / Source', isRequired: true },
        { key: 'item_sku', label: 'Item SKU' },
        { key: 'item_name', label: 'Item Received', isRequired: true },
        { key: 'quantity', label: 'Qty Ordered', type: 'number' },
        { key: 'qty_received', label: 'Qty Received', type: 'number', isRequired: true },
        { key: 'unit', label: 'Unit', type: 'select', options: ITEM_UNITS },
        { key: 'unit_cost', label: 'Unit Cost (ETB)', type: 'number' },
        { key: 'line_value', label: 'Line Value (ETB)', type: 'number', isReadOnly: true },
        { key: 'store_location', label: 'Receiving Store', type: 'select', options: MAT_LOCATIONS },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'coa_account', label: 'GL / COA Account', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'inspected_by', label: 'Inspected By' },
        { key: 'status', label: 'Status', type: 'select', options: GRV_STATUSES },
        { key: 'remarks', label: 'Inspection Notes', type: 'textarea' },
    ],
    'MAT_GIV_CRUD': [
        { key: 'giv_no', label: 'GIV Number', isRequired: true },
        { key: 'date', label: 'Issue Date', type: 'date', isRequired: true },
        { key: 'request_no', label: 'Requisition No.' },
        { key: 'department', label: 'Requesting Dept', type: 'select', options: MAT_DEPARTMENTS, isRequired: true },
        { key: 'requested_by', label: 'Requested By' },
        { key: 'item_sku', label: 'Item SKU' },
        { key: 'item_name', label: 'Item Issued', isRequired: true },
        { key: 'quantity', label: 'Qty Issued', type: 'number', isRequired: true },
        { key: 'unit', label: 'Unit', type: 'select', options: ITEM_UNITS },
        { key: 'unit_cost', label: 'Unit Cost (ETB)', type: 'number' },
        { key: 'line_value', label: 'Issue Value (ETB)', type: 'number', isReadOnly: true },
        { key: 'store_location', label: 'Issuing Store', type: 'select', options: MAT_LOCATIONS },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'coa_expense', label: 'Expense / COA Charge', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'purpose', label: 'Usage Purpose', type: 'textarea', isRequired: true },
        { key: 'issued_by', label: 'Store Keeper' },
        { key: 'received_by', label: 'Received By (Dept)' },
        { key: 'status', label: 'Status', type: 'select', options: GIV_STATUSES },
        { key: 'remarks', label: 'Remarks', type: 'textarea' },
    ],
    'MAT_ROBBING_CRUD': [
        { key: 'robbing_no', label: 'Robbing Voucher No.', isRequired: true },
        { key: 'date', label: 'Robbing Date', type: 'date', isRequired: true },
        { key: 'reason', label: 'Reason for Cannibalization', type: 'select', options: ['Urgent Repair', 'Spare Unavailable', 'Donor End-of-Life', 'Upgrade Recipient', 'Classroom Continuity', 'Other'], isRequired: true },
        { key: 'donor_asset_tag', label: 'Donor Asset Tag / SKU', isRequired: true },
        { key: 'donor_item', label: 'Donor Asset (Fully / Partially Equipped)', isRequired: true },
        { key: 'donor_location', label: 'Donor Location', type: 'select', options: MAT_LOCATIONS },
        { key: 'donor_condition_before', label: 'Donor Condition Before', type: 'select', options: ['Fully Equipped', 'Partially Equipped', 'Already Incomplete'] },
        { key: 'part_taken', label: 'Part / Component Removed', isRequired: true },
        { key: 'part_serial', label: 'Part Serial / Batch' },
        { key: 'qty_parts', label: 'Qty of Parts', type: 'number' },
        { key: 'recipient_asset_tag', label: 'Recipient Asset Tag / SKU', isRequired: true },
        { key: 'recipient_item', label: 'Recipient Asset (To Be Maintained)', isRequired: true },
        { key: 'recipient_location', label: 'Recipient Location', type: 'select', options: MAT_LOCATIONS },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'work_order', label: 'Work Order / Ticket No.' },
        { key: 'performed_by', label: 'Technician / Staff', isRequired: true },
        { key: 'approved_by', label: 'Approved By (Stores/IT Head)' },
        { key: 'donor_status_after', label: 'Donor Status After', type: 'select', options: ['Still Serviceable', 'Degraded', 'Non-Functional', 'Scrapped', 'Held for Further Robbing'] },
        { key: 'recipient_status_after', label: 'Recipient Status After', type: 'select', options: ['Restored', 'Partially Restored', 'Awaiting Further Parts', 'Failed'] },
        { key: 'status', label: 'Voucher Status', type: 'select', options: ROBBING_STATUSES },
        { key: 'coa_note', label: 'COA / Asset Note', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'remarks', label: 'Technical Notes', type: 'textarea' },
    ],
    'MAT_PURCHASE_LOCAL_CRUD': [
        { key: 'lpo_no', label: 'LPO Number', isRequired: true },
        { key: 'date', label: 'Order Date', type: 'date', isRequired: true },
        { key: 'supplier', label: 'Local Vendor', isRequired: true },
        { key: 'tin', label: 'Vendor TIN' },
        { key: 'item_sku', label: 'Item SKU' },
        { key: 'item_name', label: 'Item', isRequired: true },
        { key: 'quantity', label: 'Qty Ordered', type: 'number', isRequired: true },
        { key: 'unit', label: 'Unit', type: 'select', options: ITEM_UNITS },
        { key: 'unit_cost', label: 'Unit Cost (ETB)', type: 'number' },
        { key: 'total_value', label: 'Total (ETB)', type: 'number', isRequired: true },
        { key: 'vat_amount', label: 'VAT (ETB)', type: 'number' },
        { key: 'store_location', label: 'Deliver To Store', type: 'select', options: MAT_LOCATIONS },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'coa_account', label: 'GL / COA Account', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'requested_by', label: 'Requested By' },
        { key: 'approved_by', label: 'Approved By' },
        { key: 'expected_date', label: 'Expected Delivery', type: 'date' },
        { key: 'status', label: 'Procurement Status', type: 'select', options: PURCHASE_STATUSES },
        { key: 'remarks', label: 'Remarks', type: 'textarea' },
    ],
    'MAT_PURCHASE_FOREIGN_CRUD': [
        { key: 'import_ref', label: 'Import / PI Ref', isRequired: true },
        { key: 'date', label: 'Order / PI Date', type: 'date', isRequired: true },
        { key: 'supplier', label: 'International Vendor', isRequired: true },
        { key: 'country', label: 'Country of Origin' },
        { key: 'item_sku', label: 'Item SKU' },
        { key: 'item_name', label: 'Item', isRequired: true },
        { key: 'quantity', label: 'Qty', type: 'number', isRequired: true },
        { key: 'unit', label: 'Unit', type: 'select', options: ITEM_UNITS },
        { key: 'fob_value', label: 'FOB / Invoice (ETB)', type: 'number' },
        { key: 'freight_cost', label: 'Freight (ETB)', type: 'number' },
        { key: 'insurance_cost', label: 'Insurance (ETB)', type: 'number' },
        { key: 'customs_duty', label: 'Customs Duty (ETB)', type: 'number' },
        { key: 'shipping_cost', label: 'Other Landing Cost (ETB)', type: 'number' },
        { key: 'total_landed', label: 'Total Landed (ETB)', type: 'number', isReadOnly: true },
        { key: 'store_location', label: 'Deliver To Store', type: 'select', options: MAT_LOCATIONS },
        { key: 'campus', label: 'Campus', type: 'select', options: MAT_CAMPUSES },
        { key: 'coa_account', label: 'GL / COA Account', type: 'select', options: MAT_COA_OPTIONS },
        { key: 'bill_of_lading', label: 'B/L or AWB No.' },
        { key: 'arrival_date', label: 'Arrival Date', type: 'date' },
        { key: 'status', label: 'Import Status', type: 'select', options: PURCHASE_STATUSES },
        { key: 'remarks', label: 'Customs / Logistics Notes', type: 'textarea' },
    ],
    'ROLE_CRUD': [{ key: 'name', label: 'Role Key', isRequired: true }],
};


/* ----------------------------------------------------------------------
   8. NOTIFICATION & UTILITY BAR (Standard)
   ---------------------------------------------------------------------- */


/** Relative time for notification rows (e.g. "1 min ago"). */
const formatRelativeTime = (isoOrDate) => {
    if (!isoOrDate) return '';
    const then = new Date(isoOrDate);
    if (isNaN(then.getTime())) return String(isoOrDate).slice(0, 16);
    const sec = Math.round((Date.now() - then.getTime()) / 1000);
    if (sec < 10) return 'just now';
    if (sec < 60) return sec + ' sec ago';
    const min = Math.floor(sec / 60);
    if (min < 60) return min + (min === 1 ? ' min ago' : ' mins ago');
    const hr = Math.floor(min / 60);
    if (hr < 24) return hr + (hr === 1 ? ' hour ago' : ' hours ago');
    const day = Math.floor(hr / 24);
    if (day < 30) return day + (day === 1 ? ' day ago' : ' days ago');
    return then.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

const NotificationBar = ({ glassClass, currentUser, onClose, onCountChange }) => {
    const { getCollection, notificationTrigger } = useData();
    const [activeTab, setActiveTab] = useState('notifications');

    const today = useMemo(() => {
        const d = new Date();
        d.setHours(0, 0, 0, 0);
        return d;
    }, []);

    /** Next 3 holidays from today, only within the current calendar year */
    const upcomingHolidays = useMemo(() => {
        const year = today.getFullYear();
        const yearEnd = new Date(year, 11, 31);
        yearEnd.setHours(0, 0, 0, 0);

        // Fixed + major Ethiopian / public holidays (month is 0-based in Date)
        const catalog = [
            { name: 'Ethiopian Christmas (Genna)', m: 1, d: 7 },
            { name: 'Timkat (Epiphany)', m: 1, d: 19 },
            { name: 'Adwa Victory Day', m: 3, d: 2 },
            { name: 'International Labour Day', m: 5, d: 1 },
            { name: 'Patriots Victory Day', m: 5, d: 5 },
            { name: 'Downfall of Derg', m: 5, d: 28 },
            { name: 'Ethiopian New Year (Enkutatash)', m: 9, d: 11 },
            { name: 'Meskel (Finding of the True Cross)', m: 9, d: 27 },
            { name: 'Christmas (Gregorian)', m: 12, d: 25 },
        ];

        const list = catalog
            .map(h => {
                const date = new Date(year, h.m - 1, h.d);
                date.setHours(0, 0, 0, 0);
                return { name: h.name, date, iso: date.toISOString().slice(0, 10) };
            })
            .filter(h => h.date >= today && h.date <= yearEnd)
            .sort((a, b) => a.date - b.date)
            .slice(0, 3);

        return list;
    }, [today]);

    const items = getCollection('MAT_ITEM_CRUD') || [];
    const grv = getCollection('MAT_GRV_CRUD') || [];
    const giv = getCollection('MAT_GIV_CRUD') || [];
    const robbing = getCollection('MAT_ROBBING_CRUD') || [];
    const localPO = getCollection('MAT_PURCHASE_LOCAL_CRUD') || [];
    const foreignPO = getCollection('MAT_PURCHASE_FOREIGN_CRUD') || [];

    const n = (v) => Number(v) || 0;
    const withinDays = (dateStr, days) => {
        if (!dateStr) return false;
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return false;
        d.setHours(0, 0, 0, 0);
        const diff = (d.getTime() - today.getTime()) / (1000 * 60 * 60 * 24);
        return diff >= 0 && diff <= days;
    };
    const isPast = (dateStr) => {
        if (!dateStr) return false;
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return false;
        d.setHours(0, 0, 0, 0);
        return d.getTime() < today.getTime();
    };

    const notifications = useMemo(() => {
        const list = [];

        // —— UPCOMING HOLIDAYS (next 3 in current year only) ——
        list.push({ type: 'Title', message: `UPCOMING HOLIDAYS · ${today.getFullYear()}` });
        if (upcomingHolidays.length === 0) {
            list.push({
                type: 'Info',
                message: `No remaining holidays in ${today.getFullYear()} after today.`,
                color: 'bg-gray-800/30',
            });
        } else {
            upcomingHolidays.forEach(h => {
                const label = h.date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
                const days = Math.round((h.date.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
                list.push({
                    type: 'Info',
                    message: `📅 ${h.name} — ${label}${days === 0 ? ' · Today' : days === 1 ? ' · Tomorrow' : ` · in ${days} days`}`,
                    color: 'bg-emerald-900/30',
                });
            });
        }

        // —— STOCK ALERTS ——
        list.push({ type: 'Title', message: 'STOCK ALERTS' });
        const oos = items.filter(i => n(i.quantity) <= 0);
        const low = items.filter(i => {
            const q = n(i.quantity);
            const rl = n(i.reorder_level) || 15;
            return q > 0 && q <= rl;
        });
        const reserved = items.filter(i => /reserved|robbing/i.test(String(i.status || '')));
        const obsolete = items.filter(i => /obsolete|quarantine/i.test(String(i.status || '')));

        if (oos.length) {
            oos.slice(0, 8).forEach(i => {
                list.push({
                    type: 'Warning',
                    message: `Out of stock: ${i.sku || i.name || 'Item'} · ${i.location || '—'}`,
                    color: 'bg-rose-900/40',
                });
            });
            if (oos.length > 8) list.push({ type: 'Info', message: `+${oos.length - 8} more out-of-stock SKUs`, color: 'bg-rose-900/20' });
        } else {
            list.push({ type: 'Info', message: 'No out-of-stock items.', color: 'bg-gray-800/30' });
        }
        if (low.length) {
            low.slice(0, 6).forEach(i => {
                list.push({
                    type: 'Warning',
                    message: `Low stock (≤ reorder): ${i.name || i.sku} · qty ${n(i.quantity)} · ${i.location || ''}`,
                    color: 'bg-amber-900/40',
                });
            });
        }
        if (reserved.length) {
            list.push({
                type: 'Info',
                message: `${reserved.length} item(s) reserved / held for robbing`,
                color: 'bg-orange-900/30',
            });
        }
        if (obsolete.length) {
            list.push({
                type: 'Info',
                message: `${obsolete.length} item(s) marked obsolete or quarantine`,
                color: 'bg-slate-800/40',
            });
        }

        // —— GOODS RECEIVING ——
        list.push({ type: 'Title', message: 'GOODS RECEIVING (GRV)' });
        const grvPending = grv.filter(g => /pending|partial|inspect/i.test(String(g.status || '')));
        const grvRejected = grv.filter(g => /reject|return/i.test(String(g.status || '')));
        if (grvPending.length) {
            grvPending.slice(0, 6).forEach(g => {
                list.push({
                    type: 'Warning',
                    message: `GRV ${g.grv_no || g.id}: ${g.item_name || 'item'} · ${g.status} · ${g.supplier || ''}`,
                    color: 'bg-violet-900/35',
                });
            });
        } else {
            list.push({ type: 'Info', message: 'No GRVs pending inspection.', color: 'bg-gray-800/30' });
        }
        if (grvRejected.length) {
            list.push({
                type: 'Warning',
                message: `${grvRejected.length} GRV line(s) rejected / returned to vendor`,
                color: 'bg-rose-900/30',
            });
        }

        // —— GOODS ISSUING ——
        list.push({ type: 'Title', message: 'GOODS ISSUING (GIV)' });
        const givPending = giv.filter(g => /pending|partial/i.test(String(g.status || '')));
        if (givPending.length) {
            givPending.slice(0, 6).forEach(g => {
                list.push({
                    type: 'Warning',
                    message: `GIV ${g.giv_no || g.id}: ${g.item_name || 'item'} → ${g.department || 'dept'} · ${g.status}`,
                    color: 'bg-sky-900/35',
                });
            });
        } else {
            list.push({ type: 'Info', message: 'No GIVs awaiting approval.', color: 'bg-gray-800/30' });
        }

        // —— PROCUREMENT ——
        list.push({ type: 'Title', message: 'PROCUREMENT' });
        const openLocal = localPO.filter(p => /order|transit|pending|partial|draft/i.test(String(p.status || '')));
        const openForeign = foreignPO.filter(p => /order|transit|pending|partial|draft/i.test(String(p.status || '')));
        const overdueLocal = localPO.filter(p => p.expected_date && isPast(p.expected_date) && !/received|closed|cancel/i.test(String(p.status || '')));
        const arrivingSoon = foreignPO.filter(p => withinDays(p.arrival_date, 14) && !/received|closed|cancel/i.test(String(p.status || '')));

        openLocal.slice(0, 5).forEach(p => {
            list.push({
                type: 'Info',
                message: `LPO ${p.lpo_no || p.id}: ${p.item_name || ''} · ${p.status} · ETB ${n(p.total_value).toLocaleString()}`,
                color: 'bg-emerald-900/30',
            });
        });
        if (!openLocal.length) list.push({ type: 'Info', message: 'No open local POs.', color: 'bg-gray-800/30' });
        if (overdueLocal.length) {
            list.push({
                type: 'Warning',
                message: `${overdueLocal.length} local PO(s) past expected delivery`,
                color: 'bg-rose-900/35',
            });
        }
        openForeign.slice(0, 5).forEach(p => {
            list.push({
                type: 'Info',
                message: `Import ${p.import_ref || p.id}: ${p.item_name || ''} · ${p.status} · ${p.country || ''}`,
                color: 'bg-indigo-900/30',
            });
        });
        arrivingSoon.forEach(p => {
            list.push({
                type: 'Warning',
                message: `Arrival due: ${p.import_ref || p.item_name} on ${p.arrival_date}`,
                color: 'bg-cyan-900/35',
            });
        });

        // —— PARTS ROBBING ——
        list.push({ type: 'Title', message: 'PARTS ROBBING (CANNIBALIZATION)' });
        const robOpen = robbing.filter(r => /plan|progress|fitted|awaiting/i.test(String(r.status || r.recipient_status_after || '')));
        const robDone = robbing.filter(r => /restored|decommission|scrap/i.test(String(r.status || '')));
        if (robOpen.length) {
            robOpen.slice(0, 6).forEach(r => {
                list.push({
                    type: 'Warning',
                    message: `Robbing ${r.robbing_no || r.id}: ${r.part_taken || 'part'} · ${r.donor_item || '?'} → ${r.recipient_item || '?'} · ${r.status || ''}`,
                    color: 'bg-amber-900/35',
                });
            });
        } else {
            list.push({ type: 'Info', message: 'No open robbing vouchers.', color: 'bg-gray-800/30' });
        }
        if (robDone.length) {
            list.push({
                type: 'Info',
                message: `${robDone.length} completed / closed robbing voucher(s) on file`,
                color: 'bg-gray-800/30',
            });
        }

        // —— FINANCE POSTING ——
        list.push({ type: 'Title', message: 'FINANCE POSTING' });
        const unposted = [
            ...grv.filter(r => !r.posted_to_finance),
            ...giv.filter(r => !r.posted_to_finance),
            ...localPO.filter(r => !r.posted_to_finance),
            ...foreignPO.filter(r => !r.posted_to_finance),
            ...robbing.filter(r => !r.posted_to_finance),
        ];
        if (unposted.length) {
            list.push({
                type: 'Warning',
                message: `${unposted.length} document(s) not yet posted to Finance (GRV/GIV/PO/Robbing)`,
                color: 'bg-emerald-900/40',
            });
        } else {
            list.push({ type: 'Info', message: 'All material docs flagged posted (or empty).', color: 'bg-gray-800/30' });
        }

        // —— VALUE SNAPSHOT ——
        list.push({ type: 'Title', message: 'INVENTORY SNAPSHOT' });
        const totalVal = items.reduce((s, i) => s + (n(i.total_value) || n(i.quantity) * n(i.unit_cost)), 0);
        list.push({
            type: 'Info',
            message: `SKUs ${items.length} · Stock value ~ ETB ${totalVal.toLocaleString(undefined, { maximumFractionDigits: 0 })}`,
            color: 'bg-cyan-900/25',
        });
        list.push({
            type: 'Info',
            message: `GRV ${grv.length} · GIV ${giv.length} · LPO ${localPO.length} · Import ${foreignPO.length} · Robbing ${robbing.length}`,
            color: 'bg-cyan-900/20',
        });

        return list;
    }, [items, grv, giv, robbing, localPO, foreignPO, today, upcomingHolidays, notificationTrigger]);
    const counts = useMemo(() => ({
        all: (notifications || []).length,
        alert: (notifications || []).filter(n => n.severity === 'alert').length,
        warning: (notifications || []).filter(n => n.severity === 'warning').length,
        info: (notifications || []).filter(n => n.severity === 'info').length,
    }), [notifications]);
    useEffect(() => {
        if (typeof onCountChange !== 'function') return;
        onCountChange(Number(counts.all) || 0, counts);
    }, [counts, onCountChange]);


    const activityLog = useMemo(() => {
        const stamp = (r, label) => ({
            id: r.id,
            label,
            title: r.grv_no || r.giv_no || r.lpo_no || r.import_ref || r.robbing_no || r.sku || r.name || r.id,
            sub: r.item_name || r.part_taken || r.supplier || r.department || r.status || '',
            date: r.date || r.last_updated || r.arrival_date || r.expected_date || '',
            posted: !!r.posted_to_finance,
        });
        const pool = [
            ...grv.map(r => stamp(r, 'GRV')),
            ...giv.map(r => stamp(r, 'GIV')),
            ...localPO.map(r => stamp(r, 'LPO')),
            ...foreignPO.map(r => stamp(r, 'IMP')),
            ...robbing.map(r => stamp(r, 'ROB')),
            ...items.slice(0, 10).map(r => stamp(r, 'SKU')),
        ];
        return pool
            .sort((a, b) => String(b.date).localeCompare(String(a.date)))
            .slice(0, 20);
    }, [grv, giv, localPO, foreignPO, robbing, items]);

    const badgeCount = useMemo(() => {
        return notifications.filter(x => x.type === 'Warning').length;
    }, [notifications]);

    return (
        <div className={`w-full max-w-md ${glassClass} border border-white/10 rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]`}>
            <div className="flex items-center justify-between px-4 py-3 border-b border-white/10 bg-slate-950/80">
                <div>
                    <h3 className="text-sm font-black uppercase tracking-widest text-white">Material Alerts</h3>
                    <p className="text-[10px] text-cyan-400/80 font-bold uppercase tracking-wider mt-0.5">
                        {badgeCount} attention · Unity Stores
                    </p>
                </div>
                <button type="button" onClick={onClose} className="text-gray-400 hover:text-white p-1">
                    <Icons.X className="w-5 h-5" />
                </button>
            </div>

            <div className="flex border-b border-white/10">
                <button
                    type="button"
                    onClick={() => setActiveTab('notifications')}
                    className={`flex-1 p-2.5 text-[10px] font-black uppercase tracking-widest transition-all ${
                        activeTab === 'notifications' ? 'bg-cyan-600 text-white border-b-2 border-white' : 'bg-gray-800/40 text-gray-400'
                    }`}
                >
                    Notifications {badgeCount > 0 && <span className="ml-1 text-amber-300">({badgeCount})</span>}
                </button>
                <button
                    type="button"
                    onClick={() => setActiveTab('activity')}
                    className={`flex-1 p-2.5 text-[10px] font-black uppercase tracking-widest transition-all ${
                        activeTab === 'activity' ? 'bg-cyan-600 text-white border-b-2 border-white' : 'bg-gray-800/40 text-gray-400'
                    }`}
                >
                    Activity
                </button>
            </div>

            <div className="flex-1 overflow-y-auto custom-scrollbar p-3 space-y-2 min-h-0">
                {activeTab === 'notifications' ? (
                    notifications.map((item, index) => (
                        item.type === 'Title' ? (
                            <div key={index} className="pt-2 pb-1 text-[10px] font-black uppercase tracking-[0.2em] text-cyan-400/90 border-b border-white/5">
                                {item.message}
                            </div>
                        ) : (
                            <div key={index} className={`rounded-xl px-3 py-2 text-[11px] text-gray-200 border border-white/5 ${item.color || 'bg-black/30'}`}>
                                {item.message}
                            </div>
                        )
                    ))
                ) : (
                    activityLog.length ? activityLog.map((row, i) => (
                        <div key={row.id || i} className="rounded-xl px-3 py-2 border border-white/5 bg-black/30 flex gap-2 items-start">
                            <span className="text-[9px] font-black uppercase text-cyan-400 shrink-0 w-8 pt-0.5">{row.label}</span>
                            <div className="min-w-0 flex-1">
                                <p className="text-[11px] font-semibold text-white truncate">{row.title}</p>
                                <p className="text-[10px] text-gray-500 truncate">{row.sub}</p>
                                <p className="text-[9px] text-gray-600 mt-0.5">{row.date || '—'}{row.posted ? ' · Posted to Finance' : ''}</p>
                            </div>
                        </div>
                    )) : (
                        <p className="text-center text-gray-500 text-xs py-8">No recent material activity.</p>
                    )
                )}
            </div>
        </div>
    );
};

/* ----------------------------------------------------------------------
   9. VIEW COMPONENTS (Standard)
   ---------------------------------------------------------------------- */


const AuthScreen = ({ glassClass }) => {
    const { login } = useAuth();
    const { getCollection } = useData();
    // ADDED: Get the logo from UI Context
    const { uiSettings } = useUI(); 
    
    const [username, setUsername] = useState('');
    const [password, setPassword] = useState('');
    const [error, setError] = useState('');
    const [isSubmitting, setIsSubmitting] = useState(false);

    const handleSignIn = async (e) => {
        e.preventDefault();
        setError('');
        setIsSubmitting(true);
        try {
            const allEmployees = getCollection('HR_EMPLOYEE_CRUD') || [];
            const result = await login(username, password, allEmployees);
            if (!result.success) {
                setError(result.message || 'Invalid username or password.');
            }
        } catch (err) {
            setError('An unexpected error occurred.');
            console.error("Login Error:", err);
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <div className="h-full w-full min-h-0 flex items-center justify-center p-4 bg-slate-950 overflow-y-auto">
            <div className={`max-w-md w-full p-8 rounded-3xl shadow-2xl border border-white/10 ${glassClass} transform transition-all`}>
                
                <div className="flex flex-col items-center mb-8">
                    {/* FIXED: Conditional Logo Logic */}
                    <div className="mb-4">
                        {uiSettings?.logoURL ? (
                            // Show your UPLOADED logo if it exists
                            <img 
                                src={uiSettings.logoURL} 
                                alt="University Logo" 
                                className="w-20 h-20 object-contain drop-shadow-lg"
                            />
                        ) : (
                            // Fallback to the icon if no logo is set in UICustomization
                            <div className="p-4 bg-cyan-500/20 rounded-2xl shadow-inner">
                                <Icons.Monitor className="w-12 h-12 text-cyan-400" />
                            </div>
                        )}
                    </div>
                    
                    <h1 className="text-3xl font-black text-white tracking-tighter italic uppercase text-center">
                        Material Management - Unity University
                    </h1>
                    <p className="text-gray-400 mt-2 text-sm font-medium">Material Management Terminal</p>
                </div>

                <form onSubmit={handleSignIn} className="space-y-5">
                    {error && (
                        <div className="p-4 rounded-xl bg-red-500/10 border border-red-500/50 text-red-400 text-xs flex items-center animate-shake">
                            <Icons.AlertCircle className="w-4 h-4 mr-2 flex-shrink-0" />
                            {error}
                        </div>
                    )}

                    <div className="space-y-1">
                        <Input 
                            label="Username"
                            value={username}
                            onChange={(val) => setUsername(val)}
                            placeholder="Enter your username"
                            isRequired
                        />
                    </div>

                    <div className="space-y-1">
                        <Input 
                            label="Password"
                            type="password"
                            placeholder="Password"
                            autoComplete="current-password"
                            value={password}
                            onChange={(val) => setPassword(val)}
                            isRequired
                        />
                    </div>

                    <button 
                        type="submit"
                        disabled={isSubmitting}
                        className={`w-full py-4 rounded-2xl font-bold text-white transition-all flex items-center justify-center text-lg
                            ${isSubmitting 
                                ? 'bg-gray-700 cursor-not-allowed' 
                                : 'bg-cyan-600 hover:bg-cyan-500 active:scale-[0.97] shadow-xl shadow-cyan-900/40'
                            }`}
                    >
                        {isSubmitting ? (
                            <Icons.Loader2 className="w-6 h-6 animate-spin" />
                        ) : (
                            <>
                                <Icons.LogIn className="w-5 h-5 mr-2" />
                                Sign In
                            </>
                        )}
                    </button>
                
                    

                
                    <LoginModuleSwitcher current="material" />
                </form>

                <div className="mt-8 pt-6 border-t border-white/5 text-center">
                    
                    

                    <p className="text-[10px] text-gray-500 uppercase tracking-[0.2em] font-bold opacity-60">
                        Secure Access Terminal
                    </p>
                </div>
            </div>
        </div>
    );
};


/* ----------------------------------------------------------------------
   MODIFIED HR DASHBOARD COMPONENT (With Custom Analytics & PNG Export)
   ---------------------------------------------------------------------- */
/* ----------------------------------------------------------------------
   MODIFIED HR DASHBOARD COMPONENT (With Explaining Text)
   ---------------------------------------------------------------------- */
/* ----------------------------------------------------------------------
   MODIFIED HR DASHBOARD COMPONENT (With Explaining Text)
   ---------------------------------------------------------------------- */


const HrDashboard = ({ glassClass }) => {
    const { getCollection } = useData();
    const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
    const appraisals = getCollection('HR_APPRAISAL_CRUD') || [];
    const leaves = getCollection('HR_LEAVE_CRUD') || [];
    const [dataSource, setDataSource] = useState('Insurance');
    const [chartType, setChartType] = useState('bar');
    const chartRef = useRef(null);

    const COLORS = ['#06b6d4', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];
    const DEPARTMENTS = ['Finance', 'Engineering', 'HR', 'Marketing', 'Operations'];

    const deptAnalytics = DEPARTMENTS.map(dept => {
        const deptEmps = employees.filter(e => e.department === dept);
        return {
            name: dept,
            insUnivShare: deptEmps.length * 1200,
            appQuality: (Math.random() * 5).toFixed(2),
            recVolume: Math.floor(Math.random() * 15),
            lveDays: deptEmps.reduce((a, e) => a + (leaves.filter(l => l.employee_name === e.name && l.status === 'Approved').length * 3), 0)
        };
    });

    const chartExplanations = {
        'Insurance': 'Institution-funded insurance contributions aggregated by department showing the total financial commitment.',
        'Recruitment': 'Number of successful hires and completed recruitment requests per department for workforce planning.',
        'Appraisal': 'Average performance scores based on competency assessments and appraisal ratings across departments.',
        'Leave': 'Total workdays consumed by approved leave requests (annual, sick, emergency) per department.'
    };

    const activeData = dataSource === 'Insurance' ? deptAnalytics.map(d => ({ name: d.name, value: d.insUnivShare })) :
                       dataSource === 'Recruitment' ? deptAnalytics.map(d => ({ name: d.name, value: d.recVolume })) :
                       dataSource === 'Appraisal' ? deptAnalytics.map(d => ({ name: d.name, value: parseFloat(d.appQuality) })) :
                       deptAnalytics.map(d => ({ name: d.name, value: d.lveDays }));

    return (
        <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <KPI  
                    title="University Pay"
                    value={`${deptAnalytics.reduce((a, b) => a + b.insUnivShare, 0).toLocaleString()} Br`} 
                    iconName="ShieldCheck" 
                    colorClass="text-cyan-400" 
                    trend="Univ. Pay" 
                    glassClass={glassClass} 
                    explanation="Total institution-funded insurance contribution for the current staff roster."
                />
                <KPI  
                    title="Avg Appraisal"
                    value={(deptAnalytics.reduce((a, b) => a + b.appQuality, 0) / DEPARTMENTS.length).toFixed(2)} 
                    iconName="Star" 
                    colorClass="text-yellow-400" 
                    trend="Appraisal" 
                    glassClass={glassClass} 
                    explanation="Performance index across departments based on average staff competency scores."
                />
                <KPI  
                    title="Recruitment"
                    value={deptAnalytics.reduce((a, b) => a + b.recVolume, 0)} 
                    iconName="UserPlus" 
                    colorClass="text-green-400" 
                    trend="Success Rate" 
                    glassClass={glassClass} 
                    explanation="Count of successful candidate placements and hiring request completions."
                />
                <KPI  
                    title="Leaves Taken"
                    value={`${deptAnalytics.reduce((a, b) => a + b.lveDays, 0)} Days`} 
                    iconName="Clock" 
                    colorClass="text-purple-400" 
                    trend="Total Volume" 
                    glassClass={glassClass} 
                    explanation="Aggregation of total workdays lost to annual, sick, and emergency leave."
                />
            </div>

            {/* CHART CONTROL SECTION WITH EXPLAINING TEXT */}
            <ContainerCard Icon={Icons.Activity} glassClass={glassClass} title={`${dataSource} Trend Analysis`}>
                <div className="mb-6 p-4 rounded-xl bg-cyan-500/5 border border-cyan-500/20">
                    <p className="text-xs text-cyan-100 italic leading-relaxed">
                        <Icons.Info className="inline w-3 h-3 mr-2 mb-0.5" />
                        {chartExplanations[dataSource]}
                    </p>
                </div>

                <div className="flex flex-wrap items-end justify-between gap-4 mb-8">
                    <div className="flex gap-4 items-end">
                        <div className="flex flex-col">
                            <label className="text-[10px] text-gray-400 mb-1 font-bold uppercase">Metric Analytics</label>
                            <select 
                                className="bg-black/40 border border-white/10 text-white text-xs rounded p-2 outline-none focus:border-cyan-500"
                                value={dataSource} 
                                onChange={(e) => setDataSource(e.target.value)}
                            >
                                <option value="Insurance">Insurance (Quantity/Money)</option>
                                <option value="Recruitment">Recruitment (Volume/Success)</option>
                                <option value="Appraisal">Performance (Quality Score)</option>
                                <option value="Leave">Leave Taken (Average Days)</option>
                            </select>
                        </div>
                        <div className="flex flex-col">
                            <label className="text-[10px] text-gray-400 mb-1 font-bold uppercase">Visual Method</label>
                            <select 
                                className="bg-black/40 border border-white/10 text-white text-xs rounded p-2 outline-none focus:border-cyan-500"
                                value={chartType} 
                                onChange={(e) => setChartType(e.target.value)}
                            >
                                <option value="bar">Bar Analysis</option>
                                <option value="line">Trend Analysis</option>
                                <option value="pie">Proportional Analysis</option>
                            </select>
                        </div>

                        <div className="pb-[1px]">
                            <ExportButton 
                                data={activeData} 
                                columns={[{ label: 'Department', key: 'name' }, { label: 'Value', key: 'value' }]} 
                                filename={`HR_LIVE_${dataSource}`}
                                chartRef={chartRef} 
                            />
                        </div>
                    </div>
                </div>

                {/* DYNAMIC REAL-TIME CHART */}
                <div ref={chartRef} className="h-80 w-full bg-black/20 rounded-2xl p-4 border border-white/5">
                    <ResponsiveContainer width="100%" height="100%">
                        {chartType === 'pie' ? (
                            <PieChart>
                                <Pie data={activeData} cx="50%" cy="50%" innerRadius={60} outerRadius={100} paddingAngle={5} dataKey="value">
                                    {activeData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                                </Pie>
                                <Legend />
                                <Tooltip contentStyle={{backgroundColor: '#000', border: '1px solid #333'}} />
                            </PieChart>
                        ) : chartType === 'line' ? (
                            <LineChart data={activeData}>
                                <CartesianGrid strokeDasharray="3 3" stroke="#ffffff05" vertical={false} />
                                <XAxis dataKey="name" stroke="#6b7280" fontSize={10} />
                                <YAxis stroke="#6b7280" fontSize={10} />
                                <Tooltip contentStyle={{backgroundColor: '#000', border: '1px solid #333'}} />
                                <Line type="monotone" dataKey="value" stroke="#06b6d4" strokeWidth={4} dot={{ r: 6, fill: '#06b6d4' }} animationDuration={300} />
                            </LineChart>
                        ) : (
                            <BarChart data={activeData}>
                                <CartesianGrid strokeDasharray="3 3" stroke="#ffffff05" vertical={false} />
                                <XAxis dataKey="name" stroke="#6b7280" fontSize={10} />
                                <YAxis stroke="#6b7280" fontSize={10} />
                                <Tooltip contentStyle={{backgroundColor: '#000', border: '1px solid #333'}} />
                                <Bar dataKey="value" fill="#06b6d4" radius={[6, 6, 0, 0]} barSize={40} animationDuration={300} />
                            </BarChart>
                        )}
                    </ResponsiveContainer>
                </div>
            </ContainerCard>
        </div>
    );
};
// Internal KPI Component Updated to show Explanations
const KPI = ({ title, value, iconName, glassClass, colorClass, explanation }) => {
    const IconComponent = Icons[iconName] || Icons.Activity;

    return (
        <div className={`p-4 rounded-2xl flex flex-col justify-between ${glassClass} border border-white/5 shadow-xl h-28 hover:border-cyan-500/50 transition-all`}>
            <div className="flex items-center justify-between w-full">
                <div className="flex-1">
                    <p className="text-[9px] font-black text-gray-400 uppercase tracking-[0.2em]">{title}</p>
                    <p className="text-sm font-black text-white mt-1 tabular-nums tracking-tighter uppercase">
                        {value}
                    </p>
                </div>
                <div className={`p-2 rounded-xl bg-white/5 ${colorClass}`}>
                    <IconComponent className="w-5 h-5" />
                </div>
            </div>
            {/* Added Explaining Text here */}
            <div className="mt-2 pt-2 border-t border-white/10">
                <p className="text-[8px] leading-tight text-gray-500 italic">
                    {explanation}
                </p>
            </div>
        </div>
    );
};
// // Internal KPI Component with text-sm
// const KPI = ({ title, value, iconName, glassClass, colorClass, trend }) => {
//     const IconComponent = Icons[iconName] || Icons.Activity;
//     const TrendIcon = trend === 'up' ? Icons.ArrowUp : trend === 'down' ? Icons.ArrowDown : Icons.Minus;
//     const trendColor = trend === 'up' ? 'text-green-500' : trend === 'down' ? 'text-red-500' : 'text-gray-400';

//     return (
//         <div className={`p-4 rounded-2xl flex items-center justify-between ${glassClass} border border-white/5 shadow-xl h-24`}>
//             <div className="flex-1">
//                 <p className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">{title}</p>
//                 <p className="text-sm font-black text-white mt-1 tabular-nums tracking-tighter uppercase">
//                     {value}
//                 </p>
//                 {/* {trend !== 'none' && (
//                     <div className={`flex items-center text-[8px] font-black mt-1 ${trendColor} uppercase`}>
//                         <TrendIcon className="w-2.5 h-2.5 mr-1" />
//                         Live
//                     </div>
//                 )} */}
//             </div>
//             <div className={`p-2 rounded-xl bg-white/5 ${colorClass}`}>
//                 <IconComponent className="w-5 h-5" />
//             </div>
//         </div>
//     );
// };
/* ----------------------------------------------------------------------
   FIXED STAT CARD COMPONENT
   ---------------------------------------------------------------------- */
const StatCard = ({ label, value, Icon, color, glassClass }) => {
    // If Icon is passed as a component from the Icons object (e.g. Icons.Clock)
    // We must ensure it is rendered as a JSX tag <Icon />
    return (
        <div className={`p-4 rounded-xl ${glassClass} border border-white/10 flex items-center space-x-4`}>
            <div className={`p-3 rounded-lg bg-gray-800/50 ${color}`}>
                {/* This is where the error happens if Icon is undefined */}
                {Icon && <Icon className="w-6 h-6" />}
            </div>
            <div>
                <p className="text-xs text-gray-400 font-medium uppercase tracking-wider">{label}</p>
                <p className="text-xl font-bold text-white">{value}</p>
            </div>
        </div>
    );
};
// This component uses the existing XLSX, jsPDF, and docx libraries from your file [cite: 380, 381]
const ChartExportButton = ({ data, chartTitle, glassClass }) => {
    
    // Handler for XLSX (Excel) [cite: 380]
    const handleExcelExport = () => {
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Chart Data");
        XLSX.writeFile(workbook, `${chartTitle}_Data.xlsx`);
    };

    // Handler for PDF using jsPDF [cite: 381]
    const handlePDFExport = () => {
        const doc = new jsPDF();
        doc.text(chartTitle, 14, 15);
        // Extracts keys for headers and values for rows [cite: 403-405]
        const headers = [Object.keys(data[0])];
        const rows = data.map(item => Object.values(item));
        
        doc.autoTable({
            head: headers,
            body: rows,
            startY: 20,
        });
        doc.save(`${chartTitle}.pdf`);
    };

    return (
        <div className="flex space-x-2 mb-4">
            {/* CSV Export using existing react-csv link [cite: 380] */}
            <CSVLink 
                data={data} 
                filename={`${chartTitle}.csv`}
                className={`flex items-center px-3 py-1 rounded-md bg-cyan-600 text-white text-xs font-bold hover:bg-cyan-500 transition`}
            >
                <Icons.Download className="w-4 h-4 mr-1" /> Export CSV
            </CSVLink>

            <button onClick={handleExcelExport} className="flex items-center px-3 py-1 rounded-md bg-green-600 text-white text-xs font-bold hover:bg-green-500">
                <Icons.FileSpreadsheet className="w-4 h-4 mr-1" /> Excel
            </button>

            <button onClick={handlePDFExport} className="flex items-center px-3 py-1 rounded-md bg-red-600 text-white text-xs font-bold hover:bg-red-500">
                <Icons.FileText className="w-4 h-4 mr-1" /> PDF
            </button>
        </div>
    );
};
const RoleManagement = ({ glassClass }) => {
    const { getCollection, addRecord, updateRecord, deleteRecord } = useData();
    const [roleName, setRoleName] = useState(''); // This state holds the display name
    const [permissions, setPermissions] = useState({});
    const [editingId, setEditingId] = useState(null);

    // Get all available permissions from MENU_ITEMS
    const allAvailablePermissions = useMemo(() => {
        const ids = [];
        const extract = (items) => {
            items.forEach(item => {
                // Filter out navigation items that are not meant to be permissions (e.g., 'AUTH')
                if (item.id && item.id !== 'AUTH' && !item.children) { // Ensure it's a leaf node that represents a permission
                    ids.push({ id: item.id, label: item.label });
                }
                if (item.children) {
                    extract(item.children);
                }
            });
        };
        extract(MENU_ITEMS);
        return ids;
    }, []);

    // FIX: Ensure all menu items have valid IDs before mapping
    const uniquePermissions = useMemo(() => {
        const unique = {};
        allAvailablePermissions.forEach(p => {
            if (p.id) unique[p.id] = p;
        });
        return Object.values(unique).sort((a, b) => a.label.localeCompare(b.label));
    }, [allAvailablePermissions]);

    // FIX: Define the required CRUD Action IDs used by DataTable
    const crudActions = [
        { id: 'CAN_ADD_RECORD', label: 'Allow Creating Records', color: 'text-green-400' },
        { id: 'CAN_EDIT_RECORD', label: 'Allow Editing Records', color: 'text-yellow-400' },
        { id: 'CAN_DELETE_RECORD', label: 'Allow Deleting Records', color: 'text-red-400' }
    ];

    const currentRoles = getCollection('ROLE_CRUD') || []; // Get current roles for checks

    const handleSave = async () => {
        const trimmedRoleName = roleName.trim();
        if (!trimmedRoleName) return alert("Enter role name");

        // Combine selected page IDs AND selected action IDs
        const activePermissions = Object.keys(permissions).filter(k => permissions[k] === true);

        if (activePermissions.length === 0) {
            return alert("Please select at least one permission or action.");
        }

        // FIX: Prepare roleData to include the 'name' field as expected by COLUMN_CONFIG
        const roleData = {
            name: trimmedRoleName, // Use 'name' field for the role's display name
            roleName: trimmedRoleName, // Also keep roleName for potential backward compatibility or other uses if any.
                                      // However, 'name' is the primary key expected by DataTable.
            permissions: activePermissions,
            updatedAt: new Date().toISOString()
        };
        
        try {
            // ADDED: Check for duplicate role names (using 'name' which is now canonical)
            if (!editingId && currentRoles.some(role => role.name?.toLowerCase() === trimmedRoleName.toLowerCase())) {
                return alert(`Role "${trimmedRoleName}" already exists.`);
            }

            if (editingId) {
                await updateRecord('ROLE_CRUD', editingId, roleData);
                alert("Role Updated Successfully!");
            } else {
                // For add, ensure 'id' is unique. 'name' can be used as ID if it's guaranteed unique.
                // We'll keep 'id' generation for safety, but can optimize if 'name' is always intended as unique ID.
                const newId = `role_${Date.now()}_${trimmedRoleName.toLowerCase().replace(/\s+/g, '-')}`;
                await addRecord('ROLE_CRUD', { ...roleData, id: newId });
                alert("New Role Created!");
            }
            // Clear state after save
            setRoleName('');
            setPermissions({});
            setEditingId(null);
        } catch (error) {
            console.error("Save error:", error);
            alert(`Failed to save role: ${error.message}`);
        }
    };

    const handleEdit = (role) => {
        // Ensure we are using role.name for consistency, as COLUMN_CONFIG expects 'name'
        setEditingId(role.id); // Keep the ID for update operation
        setRoleName(role.name || role.roleName || ''); // Use role.name, fallback to roleName
        const prefs = {};
        role.permissions?.forEach(p => prefs[p] = true);
        setPermissions(prefs);
    };

    return (
        <div className="space-y-6">
            <ContainerCard  glassClass={glassClass}>
                <div className="p-6 space-y-6">
                    {/* Input for Role Name, bound to roleName state */}
                    <Input 
                        label="Role Name" 
                        value={roleName} 
                        onChange={setRoleName} 
                        placeholder="e.g. HR Staff" 
                        isRequired
                    />
                    
                    {/* NEW SECTION: ACTION BUTTON PERMISSIONS */}
                    <div>
                        <h3 className="text-white font-bold mb-3 text-sm flex items-center">
                            <Icons.ShieldCheck className="w-4 h-4 mr-2 text-cyan-400" />
                            Global Action Permissions
                        </h3>
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                            {crudActions.map(action => (
                                <button 
                                    key={action.id}
                                    onClick={() => setPermissions(prev => ({...prev, [action.id]: !prev[action.id]}))}
                                    className={`p-3 rounded-xl border text-left transition-all ${
                                        permissions[action.id] ? 'bg-cyan-500/10 border-cyan-500 shadow-lg shadow-cyan-500/10' : 'bg-white/5 border-white/10'
                                    }`}
                                >
                                    <span className={`text-xs font-bold ${action.color}`}>{action.label}</span>
                                    <div className="text-[10px] text-gray-500 mt-1">ID: {action.id}</div>
                                </button>
                            ))}
                        </div>
                    </div>

                    {/* PAGE ACCESS SECTION */}
                    <div>
                        <h3 className="text-white font-bold mb-3 text-sm flex items-center">
                            <Icons.Layout className="w-4 h-4 mr-2 text-cyan-400" />
                            Page Access
                        </h3>
                        <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                            {uniquePermissions.map(p => (
                                <button 
                                    key={p.id}
                                    onClick={() => setPermissions(prev => ({...prev, [p.id]: !prev[p.id]}))}
                                    className={`p-2 text-[10px] rounded-lg border text-left truncate transition-all ${
                                        permissions[p.id] ? 'bg-cyan-500/20 border-cyan-500 text-white' : 'bg-white/5 border-white/10 text-gray-400'
                                    }`}
                                >
                                    {p.label}
                                </button>
                            ))}
                        </div>
                    </div>

                    <button onClick={handleSave} className="w-full bg-cyan-600 hover:bg-cyan-500 p-4 rounded-xl font-bold text-white shadow-xl transition-all">
                        {editingId ? "Update Configuration" : "Save New Role"}
                    </button>
                </div>
            </ContainerCard>

            <ContainerCard  glassClass={glassClass}>
                <div className="p-4 space-y-3">
                    {(currentRoles).map(role => ( // Use currentRoles variable
                        <div key={role.id} className="flex items-center justify-between p-4 border border-white/10 bg-white/5 rounded-xl">
                            {/* FIX: Display role.name as it's the canonical name now */}
                            <div className="font-bold text-white">{role.name}</div> 
                            <div className="flex gap-2">
                                <button onClick={() => handleEdit(role)} className="bg-cyan-500/10 px-3 py-1 rounded text-cyan-400 text-xs">Edit</button>
                                {/* Added confirmation dialog for delete */}
                                <button onClick={() => {
                                    if (window.confirm(`Are you sure you want to delete role "${role.name}"? This action cannot be undone.`)) {
                                        deleteRecord('ROLE_CRUD', role.id);
                                    }
                                }} className="bg-red-500/10 px-3 py-1 rounded text-red-400 text-xs">Delete</button>
                            </div>
                        </div>
                    ))}
                </div>
            </ContainerCard>
        </div>
    );
};


    const convertGCToEC = (gcDateString) => {
        const date = new Date(gcDateString);
        if (isNaN(date.getTime())) return '';
        
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const day = date.getDate();

        // Approximate Ethiopia Calendar logic:
        // New Year (Meskerem 1) is Sept 11 (or Sept 12 in leap years)
        let etYear = year - 8;
        if (month > 9 || (month === 9 && day >= 11)) {
            etYear = year - 7;
        }

        // Ethiopian months are 30 days each. 
        // We calculate the day of the year to find the exact month/day
        const start = new Date(year, 0, 0);
        const diff = (date - start) + ((start.getTimezoneOffset() - date.getTimezoneOffset()) * 60 * 1000);
        const dayOfYear = Math.floor(diff / (1000 * 60 * 60 * 24));
        
        // Meskerem 1 is Gregorian Day 254 (non-leap) or 255 (leap)
        const isLeap = (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0));
        const newYearDay = isLeap ? 255 : 254;
        
        let ethDayOfYear = dayOfYear - newYearDay + 1;
        if (ethDayOfYear <= 0) {
            ethDayOfYear += isLeap ? 366 : 365;
        }

        const etMonth = Math.ceil(ethDayOfYear / 30);
        const etDay = ethDayOfYear % 30 || 30;

        return `${etYear}-${String(etMonth).padStart(2, '0')}-${String(etDay).padStart(2, '0')}`;
    };

const FormModal = ({ collectionName, initialData, onSave, onCancel, glassClass }) => {
    const { addRecord, updateRecord, getCollection, allData } = useData();
    const [formData, setFormData] = React.useState({});
    const isNew = React.useMemo(() => !initialData || !initialData.id, [initialData]);
    const columns = COLUMN_CONFIG[collectionName] || [];

    const isEmployee = collectionName === 'HR_EMPLOYEE_CRUD';
    const isRecruitment = collectionName === 'HR_RECRUITMENT_CRUD';
    const isLeave = collectionName === 'HR_LEAVE_CRUD';
    const isTransfer = collectionName === 'HR_TRANSFER_CRUD';
    const isTraining = collectionName === 'HR_TRAINING_CRUD';
    const isRetirement = collectionName === 'RETIREMENT_CRUD';

    /* --- 100% EXACT CALENDAR ENGINE (JDN Algorithm) --- */
    const ethToGc = (ethStr) => {
        const parts = ethStr.split('/');
        if (parts.length !== 3) return null;
        const [d, m, y] = parts.map(n => parseInt(n));
        const era = 1724220; 
        const jdn = (y - 1) * 365 + Math.floor(y / 4) + (m - 1) * 30 + d + era;
        return new Date((jdn - 2440587.5) * 86400000).toISOString().split('T')[0];
    };

    const gcToEth = (gcStr) => {
        const date = new Date(gcStr);
        if (isNaN(date.getTime())) return null;
        const jdn = Math.floor(date.getTime() / 86400000) + 2440588;
        const era = 1723856;
        const r = (jdn - era) % 1461;
        const n = (r % 365) + 365 * Math.floor(r / 1460);
        const y = 4 * Math.floor((jdn - era) / 1461) + Math.floor(r / 365) - Math.floor(r / 1460);
        const m = Math.floor(n / 30) + 1;
        const d = (n % 30) + 1;
        return { 
            full: `${String(d).padStart(2, '0')}/${String(m > 13 ? 13 : m).padStart(2, '0')}/${y}`,
            year: y
        };
    };

    /* --- DATA LOOKUP LOGIC --- */
    const handleDataLookup = (lookupValue) => {
        if (!lookupValue || isLeave || isTransfer || isTraining || isRetirement) return;
        const allRecords = getCollection(collectionName) || [];
        
        let keyToMatch = 'id_number'; 
        if (isRecruitment) keyToMatch = 'form_no';
        if (collectionName === 'HR_VACANCY_CRUD') keyToMatch = 'vacancy_no';

        const existing = allRecords.find(r => r[keyToMatch]?.trim().toLowerCase() === lookupValue.trim().toLowerCase());
        if (existing) {
            setFormData(prev => ({ ...prev, ...existing, id: existing.id }));
        }
    };

    /* --- INITIALIZATION & AUTO-ID --- */
    React.useEffect(() => {
        const allRecords = getCollection(collectionName) || [];
        const sanitized = {};
        columns.forEach(field => {
            let val = initialData?.[field.key];
            if (isNew && !val) {
                if (field.key === 'id_number' && isEmployee) {
                    const max = allRecords.reduce((m, r) => Math.max(m, parseInt(r.id_number?.split('-')[1]) || 0), 0);
                    val = `UU-${String(max + 1).padStart(3, '0')}`;
                }
                if (field.key === 'form_no' && isRecruitment) val = `HRS/036-01${allRecords.length + 1}`;

                if (field.key === 'vacancy_no' && collectionName === 'HR_VACANCY_CRUD') {
                    const lastNum = allRecords.reduce((max, vac) => {
                        const parts = vac.vacancy_no?.split('-');
                        const numStr = parts ? parts[parts.length - 1] : "0";
                        const num = parseInt(numStr) || 0;
                        return Math.max(max, num);
                    }, 0);
                    val = `VAC-${lastNum + 1}`;
                }
            }
            sanitized[field.key] = field.type === 'array' ? (Array.isArray(val) ? val : []) : (val ?? '');
        });
        setFormData(sanitized);
    }, [initialData, columns, isNew, getCollection, collectionName]); 

    const renderField = (field) => {
        if (field.isHidden) return null;

        const handleJump = () => {
            const lbl = field.label.toLowerCase();
            let tid = "";
            if (lbl.includes('educational status')) tid = "anchor-education";
            if (lbl.includes('family contact')) tid = "anchor-beneficiaries";
            const el = document.getElementById(tid);
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        };

        if (field.type === 'array') {
            const anchorId = (field.key === 'education_entries' || field.key === 'experience_entries') 
                ? "anchor-education" 
                : (field.key === 'beneficiaries' ? "anchor-beneficiaries" : "");

            return (
                <div key={field.key} id={anchorId} className="md:col-span-2 pt-4">
                    <MultiEntryField
                        label={field.label}
                        entries={formData[field.key] || []}
                        fieldConfig={field.arrayConfig}
                        glassClass={glassClass}
                        onChange={(entries) => {
                            let updates = { [field.key]: entries };
                            if (isRecruitment && field.key === 'missing_documents' && entries.length > 0) {
                                const docList = entries.map(d => d.doc_name).join(', ');
                                updates.recommendation = (formData.recommendation || '') + `\n[PENDING DOCUMENTS: ${docList}]`;
                            }
                            setFormData(prev => ({ ...prev, ...updates }));
                        }}
                    />
                </div>
            );
        }

        /* --- DYNAMIC DROPDOWN SOURCING --- */
        let options = field.options || [];
        
        if (isLeave && field.key === 'leave_type') {
            options = ['Annual Leave', 'Sick Leave', 'Maternity Leave', 'Pre-maternity Leave', 'Paternity Leave', 'Wedding Leave', 'Funeral Leave', 'Educational Leave', 'Unpaid Leave', 'Special Leave'];
        }
        
        if ((field.key === 'department' || field.key === 'new_department' || field.key === 'current_department') && allData['DEPARTMENTS_CRUD']) {
            options = allData['DEPARTMENTS_CRUD'].map(d => d.name);
        }

        if ((isLeave || isTransfer || isTraining || isRetirement) && (field.key === 'employee_name' || field.key === 'name')) {
            options = (allData['HR_EMPLOYEE_CRUD'] || []).map(e => e.name);
        }

        if (isTraining && field.key === 'score') {
            options = ['Yes', 'No'];
        }

        // UPDATED: Added current_location to dropdown sourcing for Transfer CRUD
        if (isTransfer && (field.key === 'new_location' || field.key === 'current_location')) {
            options = ['Gerji', 'Adama', 'Burayu', 'Dessie'];
        }

        if (isEmployee && field.key === 'reporting_to' && allData['HR_EMPLOYEE_CRUD']) {
            options = allData['HR_EMPLOYEE_CRUD']
                .filter(e => e.system_role === 'hr_manager' || e.system_role === 'admin' || (e.job_title && e.job_title.toLowerCase().includes('head')))
                .map(e => e.name);
        }

        return (
            <div key={field.key} className={field.type === 'textarea' ? "md:col-span-2" : "md:col-span-1"}>
                <div onClick={handleJump} className={field.label.includes('Jump') ? "cursor-pointer hover:text-cyan-400 font-bold" : ""}>
                    <Input
                        label={isTraining && field.key === 'score' ? 'Obtained Certificate?' : field.label}
                        type={options.length > 0 ? 'select' : field.type}
                        options={options}
                        value={formData[field.key] || ''}
                        onChange={(v) => {
                            let updates = { [field.key]: v };
                            
                            const isEmpKey = field.key === 'id_number' && isEmployee;
                            const isRecKey = field.key === 'form_no' && isRecruitment;
                            const isVacKey = field.key === 'vacancy_no' && collectionName === 'HR_VACANCY_CRUD';

                            if (isEmpKey || isRecKey || isVacKey) {
                                handleDataLookup(v);
                            }

                            if ((isRetirement || isTransfer || isTraining) && (field.key === 'employee_name' || field.key === 'name')) {
                                const emp = (allData['HR_EMPLOYEE_CRUD'] || []).find(e => e.name === v);
                                if (emp) {
                                    updates.department = emp.department || '';
                                    updates.current_department = emp.department || '';
                                    updates.current_location = emp.work_location || '';
                                    updates.job_title = emp.job_title || '';
                                    updates.id_number = emp.id_number || '';
                                    if (isRetirement) {
                                        updates.dob_ec = emp.dob_ec || '';
                                        updates.retirement_date_ec = emp.retirement_date_ec || '';
                                    }
                                }
                            }

                            if (isLeave && (field.key === 'start_date' || field.key === 'end_date')) {
                                const start = field.key === 'start_date' ? v : formData.start_date;
                                const end = field.key === 'end_date' ? v : formData.end_date;
                                if (start && end) {
                                    const diff = Math.ceil(Math.abs(new Date(end) - new Date(start)) / (1000 * 60 * 60 * 24)) + 1;
                                    updates.days_taken = diff > 0 ? diff : 0;
                                }
                            }

                            if (isEmployee) {
                                if (field.key === 'dob_gc' && v) {
                                    const conv = gcToEth(v);
                                    if (conv) {
                                        updates.dob_ec = conv.full;
                                        updates.age = 2026 - new Date(v).getFullYear();
                                        updates.retirement_date_ec = `${conv.full.split('/')[0]}/${conv.full.split('/')[1]}/${conv.year + 60}`;
                                    }
                                }
                                if (field.key === 'dob_ec' && v.includes('/') && v.length >= 8) {
                                    const rev = ethToGc(v); if (rev) updates.dob_gc = rev;
                                }
                                if (field.key === 'seniority_date_gc' && v) {
                                    const conv = gcToEth(v); if (conv) updates.seniority_date_ec = conv.full;
                                }
                                if (field.key === 'seniority_date_ec' && v.includes('/') && v.length >= 8) {
                                    const rev = ethToGc(v); if (rev) updates.seniority_date_gc = rev;
                                }
                            }
                            setFormData(prev => ({ ...prev, ...updates }));
                        }}
                        isRequired={field.isRequired}
                    />
                </div>
            </div>
        );
    };

    return (
        <form onSubmit={async (e) => { e.preventDefault(); (formData.id ? await updateRecord(collectionName, formData.id, formData) : await addRecord(collectionName, formData)) && onSave(); }} 
              className="space-y-6 max-h-[85vh] overflow-y-auto p-6 bg-slate-900/60 rounded-3xl custom-scrollbar">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">{columns.map(renderField)}</div>
            <div className="flex justify-end gap-6 pt-10 sticky bottom-0 bg-gray-900 p-4 rounded-b-3xl border-t border-white/5 z-50">
                <button type="button" onClick={onCancel} className="text-gray-500 font-bold uppercase text-[10px]">Discard</button>
                <button type="submit" className="bg-cyan-600 px-14 py-4 rounded-2xl font-black text-white text-[11px] uppercase shadow-2xl hover:bg-cyan-500 transition-all">
                    Commit To Registry
                </button>
            </div>
        </form>
    );
};

// Helper for dynamic button text
const getAddLabel = (slug) => {
    const parts = slug.split('_');
    return parts[1] ? parts[1].charAt(0) + parts[1].slice(1).toLowerCase() : 'Record';
};
const Sidebar = ({ currentView, setCurrentView, glassClass }) => {
    const { checkPermission, logout } = useAuth();
    const { getCollection } = useData();
    
    // FETCH THE LIVE ROLES HERE
    const allRoles = getCollection('ROLE_CRUD') || [];

    const filteredMenuItems = useMemo(() => {
        const filterItems = (items) => {
            return items
                .filter(item => {
                    // CRITICAL: Pass the allRoles list into the checkPermission function
                    return checkPermission(item.id, allRoles);
                })
                .map(item => ({
                    ...item,
                    children: item.children ? filterItems(item.children) : null
                }));
        };
        return filterItems(MENU_ITEMS);
    }, [checkPermission, allRoles]); // Re-run whenever roles change

    // ... rest of Sidebar JSX
};
/* ===== unityUniversity AI (live + knowledge) ===== */
const UU_FEMALE_FIRST_DUP = new Set(['lomi','saba','tigist','eleni','hanna','aster','netsanet','sara','sarah','bethlehem','mahlet','selam','rahel','helen','meseret','hiwot','kidist','marta','martha','yeshi','almaz','tsion','beten','sosina','liya','ruth','meron','eden','abeba','tsehay','worknesh','senait','bezawit','hawi','meaza','frehiwot','genet','ayelech','mekdes','selamawit','tsehaynesh','bethel','hirut','firehiwot','yelena']);
const UU_MALE_FIRST_DUP = new Set(['abebe','kebede','chala','jemal','fasil','bereket','yonas','henok','kalab','tadesse','dawit','lemma','daniel','samuel','solomon','yared','mulugeta','getachew','tesfaye','hailu','mekonnen','assefa','girma','wolde','negash','bekele','yonatan','nahom','abel','kaleb','biruk','natnael','robelt','kirubel','ammanuel','amanuel','eyob','yohannes','melaku','sisay','asfaw','desta','haile','tewodros','mengistu','berhanu','gashaw','tamiru','fikadu','teklu','damte']);
const uuGuessGenderFromName_UNUSED = (fullName) => {
    const first = String(fullName || '').trim().split(/\s+/)[0].toLowerCase();
    if (!first) return '';
    if (UU_FEMALE_FIRST.has(first)) return 'Female';
    if (UU_MALE_FIRST.has(first)) return 'Male';
    return '';
};
const uuCorrectEmployeeGender_DUP = (e) => {
    if (!e || typeof e !== 'object') return e;
    const name = e.full_name || e.employee_name || e.name || e.Name || '';
    const guessed = uuGuessGenderFromName(name);
    if (!guessed) return e;
    return Object.assign({}, e, { gender: guessed, Gender: guessed });
};
const uuCorrectEmployeeList_DUP = (list) => Array.isArray(list) ? list.map(uuCorrectEmployeeGender) : [];



/* ---- unityUniversity AI: dictionary + literal/contextual analyzers ---- */
const UU_AI_DICTIONARY = [
  { term: 'SKU', aliases: ['stock keeping unit', 'item code'], def: 'Stock Keeping Unit — unique code for an inventory item in Material master.', tags: ['material', 'inventory'] },
  { term: 'GRV', aliases: ['goods received voucher', 'goods received'], def: 'Goods Received Voucher — stock received into store; may post to Finance.', tags: ['material', 'finance'] },
  { term: 'GIV', aliases: ['goods issued voucher', 'goods issued'], def: 'Goods Issued Voucher — stock issued from store to a department or job.', tags: ['material'] },
  { term: 'COA', aliases: ['chart of accounts', 'chart of account'], def: 'Chart of Accounts — list of ledger accounts in Finance.', tags: ['finance'] },
  { term: 'TB', aliases: ['trial balance'], def: 'Trial Balance — accounts with debit and credit totals that should balance.', tags: ['finance'] },
  { term: 'GL', aliases: ['general ledger'], def: 'General Ledger — account-level history of postings.', tags: ['finance'] },
  { term: 'PAYE', aliases: ['income tax', 'pay as you earn'], def: 'Pay As You Earn — employee income tax withheld in payroll.', tags: ['finance', 'hr'] },
  { term: 'Basic salary', aliases: ['basic pay', 'base salary'], def: 'Core contracted pay before allowances, overtime, and deductions (HR master → payroll).', tags: ['hr', 'finance'] },
  { term: 'Parts robbing', aliases: ['part robbing', 'cannibalization'], def: 'Taking parts from one item/asset to maintain another while tracking both sides in Material.', tags: ['material'] },
  { term: 'Power Geez', aliases: ['geez calendar', 'ethiopian calendar'], def: 'Ethiopian (Geez) ↔ Gregorian converter in UI Customization.', tags: ['ui'] },
  { term: 'RBAC', aliases: ['role based access', 'roles'], def: 'Role-Based Access Control — roles grant module and permission access.', tags: ['security'] },
  { term: 'Gerji campus', aliases: ['gerji', 'al-amoudi campus', 'alamoudi'], def: 'Main Unity University campus identity in Addis Ababa.', tags: ['university'] },
  { term: 'Bootstrap', aliases: ['api bootstrap'], def: 'API load of collections from PostgreSQL into the SPA session.', tags: ['system'] },
  { term: 'Appraisal', aliases: ['performance review', 'evaluation'], def: 'HR performance evaluation record (employee, period, score).', tags: ['hr'] },
  { term: 'Award', aliases: ['recognition', 'prize'], def: 'HR recognition record under Awards.', tags: ['hr'] },
  { term: 'Equity', aliases: ['share capital', 'capital', 'owners equity'], def: 'Owners residual interest; Share Capital and related equity accounts (often 3xxx).', tags: ['finance'] },
  { term: 'Question paper', aliases: ['exam paper', 'test paper', 'a4 paper', 'copy paper', 'paper ream'], def: 'Stationery inventory concept — match Material SKUs whose name/description mentions paper / A4 / ream / exam paper.', tags: ['material', 'inventory'] },
  { term: 'Chair', aliases: ['office chair', 'seat'], def: 'Furniture inventory — SKUs with chair/seat in name or description.', tags: ['material', 'inventory'] },
  { term: 'Table', aliases: ['desk', 'office table'], def: 'Furniture inventory — SKUs with table/desk in name or description.', tags: ['material', 'inventory'] },
  { term: 'Toner', aliases: ['cartridge', 'laser toner'], def: 'IT consumable inventory — toner/cartridge SKUs.', tags: ['material', 'inventory'] },
  { term: 'Unity University', aliases: ['uu', 'unity'], def: 'Private higher education institution in Ethiopia; modules Finance, HR, Material under shared SIS login.', tags: ['university'] },
  { term: 'Moral use', aliases: ['ethics'], def: 'Report only operational facts from accessible session data; do not invent quantities or confidential fields.', tags: ['system'] },
];

const UU_STOPWORDS = new Set(('a an the in on at to for of and or is are was were be been being how many much what is there store inventory stock sku item items quantity qty total sum count number please show list me our university unity have has had do does did with from by about than then also just only campus campuses president vice founder chairman leadership faculty college school'.split(/\s+/)));

const uuNormalizeText = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9\s\-./]/g, ' ').replace(/\s+/g, ' ').trim();

/** Extract subject phrase from count/amount questions */
const uuExtractSubject = (question) => {
    const q = uuNormalizeText(question);
    let m =
        q.match(/how many (.+?)(?:\s+are|\s+is|\s+in|\s+at|\s+do|\s+does|\?|$)/) ||
        q.match(/how much (.+?)(?:\s+is|\s+are|\s+in|\s+at|\s+do|\s+does|\?|$)/) ||
        q.match(/(?:quantity|qty|total|sum|count|number) (?:of )?(.+?)(?:\s+in|\s+at|\?|$)/) ||
        q.match(/(.+?) in (?:the )?store/) ||
        q.match(/store (?:has|have|hold) (.+)/);
    let sub = m ? m[1] : '';
    sub = sub.replace(/\b(are|is|there|available|currently|now|total|please)\b/g, ' ').replace(/\s+/g, ' ').trim();
    if (!sub || sub.length < 2) {
        // fallback: meaningful tokens
        sub = q.split(/\s+/).filter((w) => w.length > 2 && !UU_STOPWORDS.has(w)).join(' ');
    }
    return sub;
};

const uuTokenize = (s) => uuNormalizeText(s).split(/\s+/).filter((w) => w.length > 1 && !UU_STOPWORDS.has(w));

/** Literal + contextual score of query tokens against a text blob */

const uuStem = (w) => {
    let t = String(w || '').toLowerCase();
    if (t.length > 4 && t.endsWith('ies')) t = t.slice(0, -3) + 'y';
    else if (t.length > 3 && t.endsWith('es') && !t.endsWith('ss')) t = t.slice(0, -2);
    else if (t.length > 3 && t.endsWith('s') && !t.endsWith('ss')) t = t.slice(0, -1);
    return t;
};

const uuExpandTokens = (tokens) => {
    const out = [];
    const seen = {};
    const add = (t) => {
        t = String(t || '').toLowerCase().trim();
        if (!t || t.length < 2 || seen[t]) return;
        seen[t] = true;
        out.push(t);
    };
    (tokens || []).forEach((tok) => {
        add(tok);
        add(uuStem(tok));
        // domain synonyms
        if (/^papers?$/.test(tok) || uuStem(tok) === 'paper') {
            ['paper', 'papers', 'a4', 'ream', 'copy', 'exam', 'stationery', 'bond'].forEach(add);
        }
        if (/^chairs?$/.test(tok) || uuStem(tok) === 'chair') {
            ['chair', 'chairs', 'seat', 'stool', 'furniture'].forEach(add);
        }
        if (/^tables?$/.test(tok) || uuStem(tok) === 'table') {
            ['table', 'tables', 'desk', 'desks', 'furniture'].forEach(add);
        }
        if (/toner|cartridge/.test(tok)) ['toner', 'cartridge', 'laser', 'hp'].forEach(add);
        if (/^pens?$/.test(tok) || uuStem(tok) === 'pen') ['pen', 'pens', 'biro', 'ballpoint'].forEach(add);
    });
    return out;
};

/** Flatten all fields of a table row for search */
const uuRowBlob = (r) => {
    if (!r || typeof r !== 'object') return '';
    const parts = [];
    try {
        Object.keys(r).forEach((k) => {
            const v = r[k];
            if (v == null) return;
            if (typeof v === 'object') {
                try { parts.push(JSON.stringify(v)); } catch (e) {}
            } else {
                parts.push(String(v));
                parts.push(String(k)); // field names help contextual match
            }
        });
    } catch (e) {
        try { parts.push(JSON.stringify(r)); } catch (e2) {}
    }
    return parts.join(' ');
};

const uuMatchScore = (queryTokens, blob) => {
    const b = uuNormalizeText(blob);
    if (!b || !queryTokens || !queryTokens.length) return 0;
    const expanded = uuExpandTokens(queryTokens);
    let score = 0;
    const phrase = queryTokens.join(' ');
    if (phrase.length > 2 && b.indexOf(uuNormalizeText(phrase)) >= 0) score += 14;
    expanded.forEach((tok) => {
        if (!tok) return;
        if (b.indexOf(tok) >= 0) score += tok.length >= 4 ? 5 : 3;
        const st = uuStem(tok);
        if (st !== tok && b.indexOf(st) >= 0) score += 4;
        // soft: token inside longer words (copy-paper, papertowel)
        if (st.length >= 4 && b.indexOf(st) >= 0) score += 1;
    });
    return score;
};

/** Analyze Material inventory for any subject — uses ALL row fields */
const uuAnalyzeStoreSubject = (question, items) => {
    const subject = uuExtractSubject(question);
    const tokens = uuTokenize(subject);
    const focus = tokens.filter((t) => !/^(store|stores|stock|inventory|sku|item|items|qty|quantity)$/.test(t));
    if (!focus.length) return null;
    const rows = Array.isArray(items) ? items : [];
    const expanded = uuExpandTokens(focus);
    const scored = [];
    rows.forEach((r) => {
        const blob = uuRowBlob(r);
        const score = uuMatchScore(focus, blob);
        if (score >= 3) {
            const qty = Number(
                r.quantity != null ? r.quantity :
                r.qty != null ? r.qty :
                r.qty_on_hand != null ? r.qty_on_hand :
                r.on_hand != null ? r.on_hand :
                r.stock_qty != null ? r.stock_qty :
                r.balance_qty != null ? r.balance_qty : 0
            ) || 0;
            scored.push({ r: r, score: score, qty: qty });
        }
    });
    scored.sort((a, b) => b.score - a.score || b.qty - a.qty);
    const best = scored[0] ? scored[0].score : 0;
    const hit = scored.filter((s) => s.score >= Math.max(3, best * 0.4));
    const qtySum = hit.reduce((s, x) => s + x.qty, 0);
    const nl = String.fromCharCode(10);
    const lines = hit.slice(0, 25).map((x) => {
        const r = x.r;
        const sku = (r.sku || r.item_sku || r.item_code || r.code || r.id) || '—';
        const name = (r.name || r.item_name || r.description || r.desc) || '';
        const store = (r.location || r.store || r.warehouse || r.receiving_store || r.store_name) || '—';
        const unit = (r.unit || r.uom || r.unit_of_measure) || '';
        return '• ' + sku + (name ? ' — ' + name : '') + ' | qty ' + x.qty + (unit ? ' ' + unit : '') + ' | ' + store + ' | score ' + Math.round(x.score);
    });
    // Near misses when empty
    let near = '';
    if (!hit.length && rows.length) {
        const soft = rows.map((r) => {
            const blob = uuRowBlob(r);
            return { r: r, score: uuMatchScore(focus, blob) };
        }).sort((a, b) => b.score - a.score).slice(0, 5);
        near = nl + 'Closest rows (low score — may still help):' + nl + soft.map((x) => {
            const r = x.r;
            const sku = (r.sku || r.item_code || r.id) || '—';
            const name = (r.name || r.item_name || r.description) || '';
            return '• ' + sku + (name ? ' — ' + name : '') + ' | score ' + Math.round(x.score);
        }).join(nl);
    }
    const body = [
        'Store / inventory analysis (all table fields · literal + stem + synonyms):',
        '• Question subject: ' + (subject || focus.join(' ')),
        '• Focus tokens: ' + focus.join(', '),
        '• Expanded search tokens: ' + expanded.slice(0, 12).join(', '),
        '• Matching SKU types: ' + hit.length,
        '• Total quantity on matches: ' + qtySum,
        '• Inventory rows scanned: ' + rows.length,
        '',
        hit.length ? 'Matches:' : 'No strong SKU match for those tokens.',
        lines.join(nl) || '• —',
        near,
        '',
        'Logic: every column on the item row is searched; plurals stem (papers→paper); synonyms (paper↔A4/ream/copy). Quantities come only from qty fields.',
        'Source: MAT_ITEM_CRUD / inventory master.',
    ].filter((x) => x !== '').join(nl);
    return {
        title: 'Live · Store analysis · ' + (subject || focus.join(' ')) + ' (' + hit.length + ' SKU · qty ' + qtySum + ')',
        body: body,
        topic: 'store_item',
        source: 'session',
    };
};

const uuDictionarySearch = (question) => {
    const q = uuNormalizeText(question);
    const qTokens = uuTokenize(q);
    const hits = [];
    (UU_AI_DICTIONARY || []).forEach((d) => {
        const term = uuNormalizeText(d.term);
        const aliases = (d.aliases || []).map(uuNormalizeText);
        let score = 0;
        if (q.indexOf(term) >= 0) score += 10;
        aliases.forEach((a) => { if (a && q.indexOf(a) >= 0) score += 8; });
        score += uuMatchScore(qTokens, term + ' ' + aliases.join(' ') + ' ' + (d.def || '') + ' ' + (d.tags || []).join(' ')) * 0.35;
        if (score >= 3) hits.push({ d: d, score: score });
    });
    hits.sort((a, b) => b.score - a.score);
    return hits;
};

const uuDictionaryAnswer = (question) => {
    try {
        const q = String(question || '');
        const wantsDef = /\b(what is|what's|whats|define|definition|meaning of|explain|dictionary|glossary)\b/i.test(q);
        const hits = uuDictionarySearch(q);
        if (!hits.length) return null;
        if (!wantsDef && hits[0].score < 8) return null;
        const top = hits.slice(0, 6);
        const nl = String.fromCharCode(10);
        const body = [
            'Dictionary search (literal + contextual):',
            ...top.map((h, i) => (i + 1) + '. ' + h.d.term + ' — ' + h.d.def + (h.d.aliases && h.d.aliases.length ? ' (also: ' + h.d.aliases.join(', ') + ')' : '')),
            '',
            'Tip: for stock quantities ask “how many [item] in store?” — the analyzer searches SKU names/descriptions.',
        ].join(nl);
        return { title: 'Dictionary · ' + top[0].d.term, body: body, topic: 'dictionary', source: 'dictionary' };
    } catch (e) { return null; }
};


const UU_AI_KNOWLEDGE = [
    {
        id: 'uu_profile',
        title: 'Unity University profile',
        tags: ['unity university', 'university', 'profile', 'about', 'what is unity', 'private', 'ethiopia', 'quality house', 'motto', 'founder', 'president', 'chairman', 'midroc', 'accreditation', 'herqa'],
        body: 'Unity University (Amharic: ዩኒቲ ዩኒቨርስቲ) is Ethiopia’s pioneer private higher education institution. Motto: “Quality House”. Type: private, co-educational, English-medium instruction. Established 1991 as Unity Language School (English, Arabic, French, Italian). Became the first private college (1998), Unity University College (2002), and the first full-fledged private university (June 2008). Founder: Fisseha Eshetu. Chairman (historical/public listings): Mohammed Hussein Al Amoudi. President/CEO (public listings): Dr. Arega Yirdaw. Accredited by the Ethiopian Ministry of Education; quality agency HERQA is commonly cited. Official website: https://www.uu.edu.et · Student portal: sis.uu.edu.et · General: info@uu.edu.et · Registrar: registrar@uu.edu.et · P.O. Box 6722, Addis Ababa. Main campus identity: Gerji / Al-Amoudi Campus, Addis Ababa. Note: program and leadership titles can change—confirm on uu.edu.et for official decisions.'
    },
    {
        id: 'uu_history',
        title: 'Unity University history',
        tags: ['history', 'established', '1991', 'language school', 'college', '2008', 'timeline', 'phases', 'midroc', 'founded', 'founder', 'construct', 'constructed', 'construction', 'built', 'build', 'when', 'year', 'started', 'start', 'opened', 'opening', 'begin', 'began'],
        body: 'Founded / constructed (as language school): 1991. Full private university status: June 2008.\nTimeline (public institutional history):\n• 1991 — Unity Language School founded; small cohort, language courses.\n• 1993–1994 — Certificate/Diploma in Basic Law; Adama branch begins expanding reach.\n• 1997–1999 — Diploma programs (Accounting, Marketing, Business, Personnel Admin, Secretarial Science); upgraded to college level (~1998), first private college in Ethiopia.\n• 1999–2000 — Undergraduate degrees expand (Accounting, Management, Economics; MIS, Law, Marketing, Computer Technology).\n• Aug 2002 — Renamed Unity University College (first private university college).\n• 2003 — Architecture & Urban Planning, Civil Engineering, Mining, Construction Technology Management; distance education college expands.\n• Feb 2007 — Postgraduate launch: MBA and Master of Development Economics.\n• Jun 2008 — Full university status; later associated with MIDROC Technology Group (public accounts cite membership ~2008–2020).\n• 2011/12 — Health sciences (SalaM Nurses College absorbed); Dessie campus (~2012).\n• 2013/14 onward — Unity Academy (KG–secondary) linked to the university ecosystem.\n• 2016/17 — Burayu campus developments (printing, academy hub, distance).\nPublic sources also cite later expansion of postgraduate fields (e.g. marketing, economics, leadership, project management) and plans toward PhD pathways.'
    },
    {
        id: 'uu_campuses',
        title: 'Campuses and locations',
        tags: ['campus', 'campuses', 'gerji', 'al-amoudi', 'adama', 'dessie', 'burayu', 'keranio', 'cmc', 'location', 'address', 'where', 'bole', 'addis', 'phone', 'contact'],
        body: 'Primary campus identity: Gerji / Al-Amoudi Campus, Addis Ababa (main operational and academic hub). Other locations cited across official and public sources:\n• Adama (Oromia) — undergraduate and related programs; historical early branch.\n• Dessie (Amhara) — undergraduate programs such as Accounting & Finance, Business Administration, Computer Science; TVET offerings reported.\n• Burayu — distance hub / academy / support services.\n• Additional Addis names appearing in public lists: CMC, Keranio/Keranyo (academy/general education contexts).\nContacts (from university contact pages; verify current numbers):\n• Main / Al-Amoudi: +251 (0)11 629 8154 · Fax +251 (0)11 629 8156 · info@uu.edu.et · registrar@uu.edu.et · P.O. Box 6722.\n• Adama: +251 (0)22 111 3666 / 4650 · P.O. Box 1486.\n• Dessie: +251 (0)33 111 5050 / 8604.\n• Burayu: +251 (0)11 260 4983 / 4731.\nModes of study: Regular, Extension, Distance/Continuing; distance coordination centers reported in Addis, Adama, Burayu, Bahir Dar, Debre Berhan.'
    },
  { id: 'uu_leadership', title: 'Unity University leadership', tags: ['president', 'vice president', 'vice-president', 'vp', 'ceo', 'chairman', 'founder', 'leadership', 'fisseha', 'amoudi', 'arega', 'vice'], body: 'Publicly cited leadership (always verify on https://www.uu.edu.et for current office-holders):\n• Founder: Fisseha Eshetu\n• Chairman (historical/public listings): Mohammed Hussein Al Amoudi\n• President / CEO (public listings): Dr. Arega Yirdaw\nVice President: universities often have academic and administrative vice presidents. This SIS library does not invent a specific current vice-president name when not published here. Use the official website, board notices, or registrar for the authoritative current vice-president name.' },

    {
        id: 'uu_programs',
        title: 'Academic programs and faculties',
        tags: ['program', 'programs', 'faculty', 'course', 'courses', 'degree', 'undergraduate', 'postgraduate', 'mba', 'engineering', 'nursing', 'computer', 'accounting', 'law', 'tvet', 'masters', 'bachelor'],
        body: 'Unity organizes teaching mainly under Business/Economics/Social Sciences; Engineering/Technology/Computational Sciences; Health Sciences; and Distance & Continuing Education.\nUndergraduate examples (campus offerings vary): Accounting & Finance, Economics, Business Administration, Marketing Management, Management, Sociology & Social Anthropology, Computer Science, Architecture & Urban Planning, Civil Engineering, Nursing, Public Health, Medical Laboratory Sciences, Law (historical listings).\nPostgraduate examples: MBA, Development Economics, Project Management; expanded masters areas (marketing, economics, leadership) reported by the university.\nTVET examples: Accounting & Finance, Marketing & Sales, Nursing, Midwifery, Hardware Maintenance & Networking, Pharmacy (campus-dependent).\nGeneral education: Unity Academy — KG through primary/secondary at Gerji, Keranyo, Burayu, Adama (as reported).\nDelivery: Regular, Extension, Distance. Always confirm the live prospectus on uu.edu.et because program lists change by campus and year.'
    },
    {
        id: 'uu_vision_mission',
        title: 'Vision mission and values',
        tags: ['vision', 'mission', 'goal', 'values', 'strategy', '2030', 'quality'],
        body: 'Vision (university public statement): to be among leading African universities in teaching-learning and research with international reputation by 2030.\nMission (public statement): contribute to sustainable development of society by providing education from KG to PhD, conducting research at high standards, and participating in community development.\nMotto: Quality House. Positioning: pioneer private HEI in Ethiopia; English-medium; multi-campus; mix of academic, TVET, and academy pathways.'
    },
    {
        id: 'uu_students_admissions',
        title: 'Students admissions and SIS portal',
        tags: ['admission', 'admissions', 'student', 'students', 'tuition', 'fee', 'sis', 'registrar', 'enrollment', 'how to apply', 'portal'],
        body: 'Student information systems and registrar services are central to operations. Public guides point to sis.uu.edu.et for student portal services and registrar@uu.edu.et for academic records.\nAdmissions (typical private HEI pattern in Ethiopia — confirm current rules with Registrar):\n• Undergraduate: secondary completion (Grade 12 / equivalent), institutional requirements, English-medium study.\n• Postgraduate: accredited bachelor degree, institutional GPA/exam requirements (e.g. entrance processes for MBA), recommendations/statement as required.\nTuition is charged in ETB; amounts vary by program and year (media has reported fee adjustments—always use official fee schedules). Scholarships and staff/family support schemes have been described in university history materials (e.g. targeted grants).\nHistorical size notes: older public figures cited ~5,000 students; later media mentions substantially larger multi-campus enrollment—treat headcount as time-sensitive.'
    },
    {
        id: 'uu_notable',
        title: 'Leadership alumni and recognition',
        tags: ['alumni', 'notable', 'bethlehem', 'solereb', 'president', 'arega', 'fisseha', 'amoudi', 'ranking'],
        body: 'Public leadership associations: founder Fisseha Eshetu; chairman listings for Mohammed Hussein Al Amoudi; president/CEO Dr. Arega Yirdaw in contemporary public materials.\nNotable alumni often cited: Bethlehem Tilahun Alemu (entrepreneur, soleRebels), Maria Yusuf (activist).\nRecognition: first private college and first full-fledged private university in Ethiopia; MOE recognition; HERQA referenced for quality assurance. Rankings in global directories are secondary signals only—not substitutes for accreditation status.'
    },
    {
        id: 'sis_overview',
        title: 'Unity SIS software overview',
        tags: ['sis', 'system', 'dashboard', 'modules', 'rbac', 'login', 'integration', 'unity sis', 'software', 'app'],
        body: 'Unity SIS (this application) is an integrated institutional system for Unity University operations with three core modules under shared authentication and RBAC:\n1) Finance — chart of accounts, partners, journals/single transactions, trial balance, general ledger, payroll process & reconciliation, financial statements (position, P&L, cash flow), exports.\n2) HR — employee master, departments, leave, recruitment, training, transfer, appraisal, competence, awards/penalties, role assignment, notifications.\n3) Material — inventory/SKU master, GRV, GIV, purchases, parts robbing; posting hooks toward Finance.\nCross-cutting: module switcher after login, shared credentials/roles, bootstrap from API/PostgreSQL when backend is up, UI Customization (theme, Field Intelligence, Power Geez calendar, unityUniversity AI), mobile-friendly layouts, robust search/export on major tables.'
    },
    {
        id: 'finance_module',
        title: 'Finance module in depth',
        tags: ['finance', 'coa', 'chart of accounts', 'journal', 'transaction', 'trial balance', 'tb', 'ledger', 'gl', 'payroll', 'payslip', 'reconciliation', 'partner', 'supplier', 'customer', 'fp', 'p&l', 'cashflow', 'cash flow', 'statement', 'export'],
        body: 'Finance module purpose: institutional accounting and payroll finance for Unity University.\nKey masters/screens:\n• Chart of Accounts (COA) — account codes (e.g. 6110 salaries, 1141 bank, 1520 inventory-related), types Asset/Liability/Equity/Revenue/Expense.\n• Partners — customers, suppliers, banks, government (ERCA, pension) with default AR/AP/bank/expense/income COA maps.\n• Single transactions & day journals — source lines that feed TB/GL (not synthetic COA-only amounts).\n• Trial Balance & General Ledger — balances from posted transactions/journals (+ payroll postings when integrated).\n• Payroll process — select employees from HR master; earnings (basic, doctoral, allowances, OT, cash indemnity), deductions (PAYE, pension EE/ER, cost sharing, credit, sundry), net pay; payslip print.\n• Payroll reconciliation — I Payment totals vs II difference analysis so III net ≈ 0 when changes are explained.\n• Statements — Financial Position, Profit or Loss, Cash Flow (print-oriented layouts).\n• Exports — Excel/PDF/Word/CSV with institutional header framing.\nOutputs users expect: balanced TB, GL movement, payslips, recon sheets, partner balances, statement packs, audit-friendly exports.'
    },
    {
        id: 'hr_module',
        title: 'HR module in depth',
        tags: ['hr', 'human resource', 'employee', 'staff', 'leave', 'recruitment', 'training', 'transfer', 'appraisal', 'competence', 'award', 'department', 'role', 'notification', 'retirement', 'resignation'],
        body: 'HR module purpose: staff system of record and workforce processes.\nEmployee master fields (representative): ID NO, full name, gender, DOB GC/EC, age, marital status, seniority GC/EC, work location, department, position/academic rank, reporting to, employment category/type, education, GPA, field of study, salary, family contact, pension ID, retirement EC, national ID, TIN, phones, emergency contacts, remarks, system access role.\nOther areas: Departments; Leave; Recruitment / vacancies; Training; Transfer; Appraisal & competence matrix scoring; Awards/penalties; Insurance; attrition/resignation/retirement notifications; Role assignment on users (not only staff names); HR dashboard analytics; Report generator.\nIntegration: employee list and salary attributes feed Finance payroll process; roles feed module RBAC.\nOutputs: complete staff registers, leave balances, recruitment pipelines, appraisal scores, role maps, exportable HR reports, notification triggers (e.g. approaching retirement).'
    },
    {
        id: 'material_module',
        title: 'Material module in depth',
        tags: ['material', 'inventory', 'sku', 'grv', 'giv', 'store', 'warehouse', 'purchase', 'po', 'lpo', 'robbing', 'stock', 'item'],
        body: 'Material module purpose: inventory and store logistics for campuses/stores.\nCore ideas:\n• Item/SKU master — codes, descriptions, units, costs, store locations (e.g. Main Store West, IT Store, Furniture Store), campus.\n• GRV (Goods Received Voucher) — receiving against PO/LPO, vendor, qty, unit cost, line value, inspection, GL/COA, status; can post toward Finance.\n• GIV (Goods Issued Voucher) — issues to departments/jobs.\n• Local/foreign purchase documents; parts robbing (take parts from one asset/item to maintain another—Unity operational meaning).\n• Dashboard cards, search, filters, robust exports matching Finance/HR style.\nIntegration: GRV/GIV postings create finance legs (inventory/expense vs payable/bank) via bridge/API when configured.\nOutputs: stock positions by SKU/store, GRV/GIV registers, purchase follow-up, posting status to Finance, export packs.'
    },
    {
        id: 'integration_flow',
        title: 'Cross-module integration',
        tags: ['integration', 'posting', 'bridge', 'sync', 'payroll post', 'material post', 'bootstrap', 'api', 'postgres', 'backend'],
        body: 'Integration design:\n• Shared Auth/RBAC — one login model; module access by role; users created/assigned roles in settings.\n• SharedData / bootstrap — frontend can load collections from Node/Express API + PostgreSQL (`/api/bootstrap`, `/api/collections/...`).\n• HR → Finance: employee master drives payroll employee select; basic salary and identifiers (TIN, pension) flow into process; payroll commit posts journal legs (wages, tax, pension, net payable).\n• Material → Finance: GRV/GIV post creates balanced transaction legs mapped to COA (e.g. inventory/stock and AP/bank).\n• Finance statements/TB should reflect only real transactions/journals/postings—not random COA seed amounts.\n• When API is down, modules may keep local/session data; production intent is DB as source of truth after seed/migrate.'
    },
    {
        id: 'ui_ai_geez',
        title: 'UI customization AI and Power Geez',
        tags: ['ui', 'customization', 'theme', 'geez', 'calendar', 'ethiopian', 'power geez', 'ai', 'assistant', 'field intelligence'],
        body: 'UI Customization (all modules): atmospheric theme (scale, glass, logo/background), Field Intelligence tooltips, Power Geez Ethiopian↔Gregorian converter with month grid, and unityUniversity AI.\nPower Geez: convert Gregorian→Geez and Geez→Gregorian; calendar click fills converter; month names Meskerem–Pagumen.\nunityUniversity AI: answers from (1) live session data (employees, SKUs, GRV, payroll rows, COA counts, etc.) and (2) this knowledge library about the university and the three modules. Prefer live numbers when the question is about current records; prefer library for institutional facts (history, campuses, programs, contacts).'
    },
    {
        id: 'outputs_reports',
        title: 'Reports exports and outputs',
        tags: ['report', 'export', 'excel', 'pdf', 'word', 'csv', 'print', 'payslip', 'output', 'dashboard'],
        body: 'Common outputs across modules:\n• On-screen dashboards with KPI cards (clickable where implemented to retarget charts/tables).\n• Data tables with search, filters, status chips, add/edit modals.\n• Export dropdown: Excel (formatted workbook headers), Word, PDF, CSV, optional PNG capture—Finance-style framing reused in HR/Material.\n• Finance-specific: TB, GL, FP, P&L, Cash Flow print layouts; payslip full-page print; payroll recon sheet.\n• HR-specific: staff lists, leave/training/appraisal extracts.\n• Material-specific: inventory and GRV/GIV registers.\n• Report generator screens build presentable filtered extracts by collection/fields.'
    },
];


const uuSafeList = (data, key) => {
    try {
        if (!data) return [];
        if (Array.isArray(data[key])) return data[key];
    } catch (e) {}
    return [];
};
const uuPull = (allData, getCollection, key) => {
    try {
        if (typeof getCollection === 'function') {
            const v = getCollection(key);
            if (Array.isArray(v)) return v;
        }
    } catch (e) {}
    const aliases = {
        MAT_ITEM_CRUD: ['MAT_ITEM_CRUD', 'MAT_INVENTORY_CRUD', 'MAT_MASTER_CRUD', 'inventoryItems', 'items'],
        HR_EMPLOYEE_CRUD: ['HR_EMPLOYEE_CRUD', 'employees'],
        MAT_GRV_CRUD: ['MAT_GRV_CRUD', 'grv'],
        MAT_GIV_CRUD: ['MAT_GIV_CRUD', 'giv'],
        FIN_PARTNERS_CRUD: ['FIN_PARTNERS_CRUD', 'partners'],
        FIN_COA_CRUD: ['FIN_COA_CRUD', 'chartsOfAccounts', 'coa'],
        transactions: ['transactions', 'dayJournals', 'FIN_TRANSACTIONS'],
        PAY_PROCESS_CRUD: ['PAY_PROCESS_CRUD'],
        HR_AWARD_CRUD: ['HR_AWARD_CRUD', 'HR_AWARD_CRUD_HR', 'HR_STAFF_AWARDS_CRUD', 'awards'],
        HR_PENALTY_CRUD: ['HR_PENALTY_CRUD', 'penalties'],
        HR_TRAINING_CRUD: ['HR_TRAINING_CRUD', 'training'],
        HR_APPRAISAL_CRUD: ['HR_APPRAISAL_CRUD', 'appraisals'],
        HR_LEAVE_CRUD: ['HR_LEAVE_CRUD'],
    };
    try {
        for (const k of (aliases[key] || [key])) {
            const arr = uuSafeList(allData || {}, k);
            if (arr.length) return arr;
        }
    } catch (e) {}
    return uuSafeList(allData || {}, key);
};
const uuEmpName = (e) => { try { return String((e && (e.full_name || e.employee_name || e.name || e.Name)) || '—').trim(); } catch (err) { return '—'; } };
const uuEmpId = (e) => { try { return String((e && (e.id_number || e.employee_id || e.emp_id || e.id)) || '—').trim(); } catch (err) { return '—'; } };
const uuEmpDept = (e) => { try { return String((e && (e.department || e.dept || e.Department)) || '—').trim(); } catch (err) { return '—'; } };
const uuEmpPos = (e) => { try { return String((e && (e.current_position || e.job_title || e.position || e.title)) || '—').trim(); } catch (err) { return '—'; } };
const uuEmpGenderRaw = (e) => {
    try {
        if (!e || typeof e !== 'object') return '';
        if (e.gender != null) return e.gender;
        if (e.Gender != null) return e.Gender;
        if (e.sex != null) return e.sex;
        if (e.Sex != null) return e.Sex;
        return '';
    } catch (err) { return ''; }
};
const uuNormalizeGender = (g) => {
    try {
        const x = String(g == null ? '' : g).toLowerCase().trim();
        if (!x) return 'unknown';
        if (x === 'f' || x === 'female' || x === 'woman' || x === 'women' || x === 'girl') return 'female';
        if (x === 'm' || x === 'male' || x === 'man' || x === 'men' || x === 'boy') return 'male';
        if (/\bfemale\b|\bwomen\b|\bwoman\b/.test(x)) return 'female';
        if (/\bmale\b|\bmen\b|\bman\b/.test(x) && !/\bfemale\b|\bwomen\b|\bwoman\b/.test(x)) return 'male';
        if (x.startsWith('fem')) return 'female';
        if (x.startsWith('mal')) return 'male';
        return 'unknown';
    } catch (err) { return 'unknown'; }
};
const uuMatchGender = (recordGender, want) => !want || uuNormalizeGender(recordGender) === want;
const uuParseEmployeeFilters = (q) => {
    try {
        const s = String(q || '').toLowerCase();
        let gender = null;
        if (/\bfemale\b|\bwomen\b|\bwoman\b|\bfemales\b/.test(s)) gender = 'female';
        else if (/\bmale\b|\bmen\b|\bman\b|\bmales\b/.test(s) && !/\bfemale\b|\bwomen\b|\bwoman\b/.test(s)) gender = 'male';
        let department = null;
        const deptMatch = s.match(/\b(?:in|from|at)\s+([a-z][a-z\s]{1,40}?)(?:\s+department|\s+dept)?(?:\s|$)/);
        if (deptMatch) {
            const d = deptMatch[1].trim();
            if (!/^(unity|university|the|a|an|hr|finance|material)$/.test(d)) department = d;
        }
        let location = null;
        for (const loc of ['gerji', 'burayu', 'dessie', 'adama', 'torhailoch', 'bole', 'addis ababa', 'west campus']) {
            if (s.includes(loc)) { location = loc; break; }
        }
        return { gender, department, location };
    } catch (err) {
        return { gender: null, department: null, location: null };
    }
};
const uuFilterEmployees = (list, filters) => {
    try {
        const rows = Array.isArray(list) ? list : [];
        const f = filters || {};
        return rows.filter((e) => {
            if (!e) return false;
            if (f.gender && !uuMatchGender(uuEmpGenderRaw(e), f.gender)) return false;
            if (f.department && !uuEmpDept(e).toLowerCase().includes(String(f.department).toLowerCase())) return false;
            if (f.location) {
                const loc = String((e.work_location || e.location || '')).toLowerCase();
                if (!loc.includes(String(f.location).toLowerCase())) return false;
            }
            return true;
        });
    } catch (err) { return []; }
};
const uuGenderCensus = (list) => {
    try {
        const rows = Array.isArray(list) ? list : [];
        let female = 0, male = 0, unknown = 0;
        rows.forEach((e) => {
            const g = uuNormalizeGender(uuEmpGenderRaw(e));
            if (g === 'female') female += 1;
            else if (g === 'male') male += 1;
            else unknown += 1;
        });
        return { total: rows.length, female, male, unknown };
    } catch (err) { return { total: 0, female: 0, male: 0, unknown: 0 }; }
};
const uuResolveQuery = (question, priorTopic, priorQuestion) => {
    try {
        const raw = String(question || '').trim();
        const q = raw.toLowerCase();
        if (!q) return { query: '', topic: '', filters: {}, followUp: false };

        let topic = '';
        if (/\b(how many|how much)\b/.test(q) && /\b(store|stock|inventory)\b/.test(q)) topic = 'store_item';
        else if (/\b(question paper|copy paper|a4 paper|toner|chairs?|tables?)\b/.test(q)) topic = 'store_item';
        else if (/\b(chairs?|stools?)\b/.test(q)) topic = 'chairs';
        else if (/\b(basic salary|total basic|total salary)\b/.test(q) || (/\bsalary\b/.test(q) && /\b(total|sum|how much)\b/.test(q))) topic = 'salary_total';
        else if (/\b(award|awards|awarded|recognition)\b/.test(q)) topic = 'award';
        else if (/\b(appraisals?|performance review|competence)\b/.test(q)) topic = 'appraisal';
        else if (/\b(on leave|leave records?|how many .*leave|annual leave|sick leave)\b/.test(q) || (/\bleave\b/.test(q) && /\b(how many|count|list|who)\b/.test(q))) topic = 'leave';
        else if (/\b(share capital|paid[- ]?up capital|equity capital|how much .*capital|capital now)\b/.test(q) || (/\bcapital\b/.test(q) && /\b(how much|total|current|now|balance)\b/.test(q))) topic = 'capital';
        else if (/\b(sku|skus|item|items|inventory|stock|product|tables? in store|store tables)\b/.test(q)) topic = 'sku';
        else if (/\b(how many tables?|tables? in (the )?store)\b/.test(q)) topic = 'tables';
        else if (/\b(employee|employees|staff|worker|personnel|female|male|women|men|academic|non[- ]?academic)\b/.test(q)) topic = 'employee';
        else if (/\b(grv|goods received)\b/.test(q)) topic = 'grv';
        else if (/\b(giv|goods issued)\b/.test(q)) topic = 'giv';
        else if (/\bleave\b/.test(q)) topic = 'leave';
        else if (/\b(partner|supplier|customer|vendor)\b/.test(q)) topic = 'partner';
        else if (/\b(coa|chart of account)/.test(q)) topic = 'coa';
        else if (/\b(payroll|payslip|salary process)\b/.test(q)) topic = 'payroll';
        else if (/\b(profit|loss|p&l|pnl|net income)\b/.test(q)) topic = 'profit';
        else if (/\b(trial\s*balance|\btb\b)\b/.test(q)) topic = 'tb';
        else if (/\b(transaction|journal)\b/.test(q)) topic = 'tx';

        const softFollow = /^(be )?specific|more detail|tell me more|what about (them|those|it)|elaborate|details?$|narrow|focus on that|same as above|specific number|exact number|how many exactly/.test(q);
        const followUp = softFollow && !!priorTopic && !topic;
        if (followUp) topic = priorTopic;

        let query = raw;
        if (followUp && priorTopic) {
            query = raw + ' (about ' + priorTopic + (priorQuestion ? '; earlier: ' + priorQuestion : '') + ')';
        }
        // Follow-ups like "specific number of awards?" keep awards topic
        if (!topic && priorTopic && /\b(specific|number|count|how many|exact)\b/.test(q)) topic = priorTopic;

        const filterSrc = query + (followUp && priorQuestion ? ' ' + priorQuestion : '');
        return { query, topic, followUp: !!followUp, filters: typeof uuParseEmployeeFilters === 'function' ? uuParseEmployeeFilters(filterSrc) : {} };
    } catch (err) {
        return { query: String(question || ''), topic: '', filters: {}, followUp: false };
    }
};


const unityUniversityLiveAnswer = (question, allData, getCollection, topicHint, filtersIn) => {
    try {
        const q = String(question || '').toLowerCase().trim();
        if (!q && !topicHint) return null;
        const filters = filtersIn || (typeof uuParseEmployeeFilters === 'function' ? uuParseEmployeeFilters(q) : {}) || {};
                let topic = topicHint || '';
        if (!topic) {
            // Specific operational questions FIRST (before broad employee/module matches)
            if (/(campus|campuses|president|vice[ -]?president|vp|chairman|founder|leadership|faculty|college|university only)/.test(q)) topic = 'institutional';
            else if (/(how many|how much|quantity|qty)/.test(q) && /(store|stock|inventory|sku|warehouse)/.test(q) && !/(campus|president|university leadership)/.test(q)) topic = 'store_item';
            else if (/(question paper|copy paper|a4|toner|chairs?|tables?|desks?|pens?|paper towel)/.test(q) && /(store|stock|how many|how much|inventory)/.test(q)) topic = 'store_item';
            else if (/\b(chairs?|stools?)\b/.test(q)) topic = 'chairs';
            else if (/\b(basic salary|total basic|sum of (the )?salary|total salary|payroll basic)\b/.test(q) || (/\bsalary\b/.test(q) && /\b(total|sum|how much|aggregate)\b/.test(q))) topic = 'salary_total';
            else if (/\b(award|awards|awarded|recognition|medal|prize)\b/.test(q)) topic = 'award';
            else if (/\b(appraisals?|performance review|competence matrix|evaluation)\b/.test(q)) topic = 'appraisal';
            else if (/\b(on leave|leave records?|how many .*leave|leave balance|annual leave|sick leave)\b/.test(q) || (/\bleave\b/.test(q) && /\b(how many|count|list|who|employee|staff)\b/.test(q))) topic = 'leave';
            else if (/\b(share capital|paid[- ]?up capital|equity capital|how much .*capital|capital now|owners'? equity)\b/.test(q) || (/\bcapital\b/.test(q) && /\b(how much|total|current|now|balance)\b/.test(q))) topic = 'capital';
            else if (/\b(profit|loss|p&l|pnl|surplus|deficit|net income|revenue|income statement)\b/.test(q)) topic = 'profit';
            else if (/\b(trial\s*balance|\btb\b)\b/.test(q)) topic = 'tb';
            else if (/\b(transaction|journal|ledger entry|list of transaction)\b/.test(q)) topic = 'tx';
            else if (/\b(academic|acedamic|non[-\s]?academic|non[-\s]?acedamic|nonacademic|nonacedamic)\b/.test(q)) topic = 'emp_category';
            else if (/\b(how many tables?|tables? in (the )?store|office tables?|furniture tables?)\b/.test(q)) topic = 'tables';
            else if (/\btables?\b/.test(q) && /\bstore\b/.test(q)) topic = 'tables';
            else if (/\b(sku|skus|item|items|inventory|stock)\b/.test(q)) topic = 'sku';
            else if (/\bgrv\b/.test(q)) topic = 'grv';
            else if (/\bgiv\b/.test(q)) topic = 'giv';
            else if (/\bleave\b/.test(q)) topic = 'leave';
            else if (/\b(employee|staff|female|male|women|men|personnel)\b/.test(q) || filters.gender) topic = 'employee';
            else if (/\b(partner|supplier|customer)\b/.test(q)) topic = 'partner';
            else if (/\b(coa|chart of account)/.test(q)) topic = 'coa';
            else if (/\b(payroll|payslip)\b/.test(q)) topic = 'payroll';
            else if (/\b(dashboard|overview|summary|kpi)\b/.test(q)) topic = 'dashboard';
        }

        const employees = (typeof uuCorrectEmployeeList === 'function'
            ? uuCorrectEmployeeList(uuPull(allData, getCollection, 'HR_EMPLOYEE_CRUD'))
            : uuPull(allData, getCollection, 'HR_EMPLOYEE_CRUD'));
        const leave = uuPull(allData, getCollection, 'HR_LEAVE_CRUD');
        const items = uuPull(allData, getCollection, 'MAT_ITEM_CRUD');
        const grv = uuPull(allData, getCollection, 'MAT_GRV_CRUD');
        const giv = uuPull(allData, getCollection, 'MAT_GIV_CRUD');
        const partners = uuPull(allData, getCollection, 'FIN_PARTNERS_CRUD');
        const coa = uuPull(allData, getCollection, 'FIN_COA_CRUD');
        let tx = uuPull(allData, getCollection, 'transactions');
        if (!tx.length) tx = uuPull(allData, getCollection, 'dayJournals');
        const process = uuPull(allData, getCollection, 'PAY_PROCESS_CRUD');
        const awards = uuPull(allData, getCollection, 'HR_AWARD_CRUD');
        const penalties = uuPull(allData, getCollection, 'HR_PENALTY_CRUD');
        const training = uuPull(allData, getCollection, 'HR_TRAINING_CRUD');
        const appraisals = uuPull(allData, getCollection, 'HR_APPRAISAL_CRUD');
        const yearNow = new Date().getFullYear();

        const money = (n) => {
            const v = Number(n) || 0;
            try { return v.toLocaleString('en-ET', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
            catch (e) { return String(Math.round(v * 100) / 100); }
        };
        const txDebit = (r) => Number(r && (r.debit != null ? r.debit : (r.type === 'Debit' || r.txn_type === 'Debit' ? r.amount : 0))) || 0;
        const txCredit = (r) => Number(r && (r.credit != null ? r.credit : (r.type === 'Credit' || r.txn_type === 'Credit' ? r.amount : 0))) || 0;
        const txAcc = (r) => String((r && (r.account_code || r.account_no || r.account_id || r.accountCode || r.coa)) || '—').replace(/^acc-/i, '');
        const txName = (r) => String((r && (r.account_name || r.narration || r.description || r.name)) || '').trim();
        const txDate = (r) => String((r && (r.date || r.txn_date || r.posting_date || r.created_at)) || '').slice(0, 10);
        const txYear = (r) => {
            const d = txDate(r);
            const y = parseInt(d.slice(0, 4), 10);
            return Number.isFinite(y) ? y : yearNow;
        };

        // Build simple TB from transactions
        const buildTb = () => {
            const ledger = {};
            (tx || []).forEach((r) => {
                const acc = txAcc(r);
                if (!acc || acc === '—') return;
                if (!ledger[acc]) ledger[acc] = { account: acc, name: txName(r) || acc, debit: 0, credit: 0 };
                ledger[acc].debit += txDebit(r);
                ledger[acc].credit += txCredit(r);
                if (!ledger[acc].name || ledger[acc].name === acc) {
                    const nm = txName(r);
                    if (nm) ledger[acc].name = nm;
                }
            });
            // enrich names from COA
            (coa || []).forEach((c) => {
                const code = String((c && (c.account_no || c.account_code || c.id)) || '').replace(/^acc-/i, '');
                if (code && ledger[code]) {
                    ledger[code].name = c.name || c.desc || c.description || ledger[code].name;
                    ledger[code].type = c.account_type || c.type || c.category || '';
                }
            });
            return Object.values(ledger).sort((a, b) => String(a.account).localeCompare(String(b.account)));
        };

        const classifyAcc = (code, typeHint) => {
            const t = String(typeHint || '').toLowerCase();
            if (/revenue|income/.test(t)) return 'revenue';
            if (/expense|cost/.test(t)) return 'expense';
            if (/asset/.test(t)) return 'asset';
            if (/liab/.test(t)) return 'liability';
            if (/equity|capital/.test(t)) return 'equity';
            const c = String(code || '');
            const head = parseInt(c.slice(0, 1), 10) || parseInt(c.slice(0, 4), 10);
            // UU-style: 1xxx assets, 2xxx liabilities, 3xxx equity, 4xxx revenue, 5-6xxx expense (approx)
            if (/^4/.test(c)) return 'revenue';
            if (/^[56]/.test(c)) return 'expense';
            if (/^1/.test(c)) return 'asset';
            if (/^2/.test(c)) return 'liability';
            if (/^3/.test(c)) return 'equity';
            return 'other';
        };

        const empCategory = (e) => {
            const raw = String((e && (e.employment_category || e.category || e.emp_category || e.EmploymentCategory || e.staff_type)) || '').toLowerCase();
            if (/non[-\s]?academic|nonacademic|admin|support/.test(raw)) return 'non-academic';
            if (/academic|lecturer|faculty|teaching/.test(raw)) return 'academic';
            // infer from position
            const pos = String((e && (e.current_position || e.position || e.job_title || e.title)) || '').toLowerCase();
            if (/lecturer|professor|dean|academic/.test(pos)) return 'academic';
            if (/officer|registrar|finance|hr |human|store|material|it |admin/.test(pos)) return 'non-academic';
            return raw ? raw : 'unspecified';
        };

        // —— Dashboard overview ——
        if (topic === 'dashboard' || (!topic && /\b(overview|summary)\b/.test(q))) {
            const census = typeof uuGenderCensus === 'function' ? uuGenderCensus(employees) : { total: employees.length, female: 0, male: 0, unknown: 0 };
            const tb = buildTb();
            let td = 0, tc = 0;
            tb.forEach((r) => { td += r.debit; tc += r.credit; });
            return {
                title: 'Live · SIS dashboard snapshot',
                body: [
                    'Institutional live snapshot (session data):',
                    '• Employees: ' + census.total + ' (F ' + census.female + ' · M ' + census.male + ')',
                    '• Leave rows: ' + leave.length,
                    '• Inventory SKUs: ' + items.length,
                    '• GRV: ' + grv.length + ' · GIV: ' + giv.length,
                    '• Partners: ' + partners.length,
                    '• COA rows: ' + coa.length,
                    '• Transactions/journals: ' + tx.length,
                    '• Payroll process rows: ' + process.length,
                    '• Trial balance accounts with movement: ' + tb.length,
                    '• TB totals — Debit ' + money(td) + ' · Credit ' + money(tc),
                    '',
                    'Ask follow-ups: trial balance · list of transactions · profit this year · academic vs non-academic · female staff · SKUs · GRV',
                ].join(String.fromCharCode(10)),
            };
        }

        // —— Academic vs non-academic ——
        if (topic === 'emp_category' || /\b(academic|acedamic|non[-\s]?academic|non[-\s]?acedamic)\b/.test(q)) {
            const groups = { academic: [], 'non-academic': [], unspecified: [] };
            (employees || []).forEach((e) => {
                const c = empCategory(e);
                if (!groups[c]) groups[c] = [];
                groups[c].push(e);
            });
            const linePeople = (arr) => arr.slice(0, 12).map((e) => {
                const id = typeof uuEmpId === 'function' ? uuEmpId(e) : (e.id_number || e.id);
                const name = typeof uuEmpName === 'function' ? uuEmpName(e) : (e.full_name || e.name);
                const pos = typeof uuEmpPos === 'function' ? uuEmpPos(e) : (e.current_position || e.position || '—');
                return '  • ' + id + ' | ' + name + ' | ' + pos;
            }).join(String.fromCharCode(10));
            const body = [
                'Staff by employment category (live HR master):',
                '• Academic: ' + groups.academic.length,
                '• Non-academic: ' + groups['non-academic'].length,
                '• Unspecified: ' + groups.unspecified.length,
                '• Total employees: ' + employees.length,
                '',
                'Academic sample:',
                linePeople(groups.academic) || '  • —',
                '',
                'Non-academic sample:',
                linePeople(groups['non-academic']) || '  • —',
                '',
                'Category is read from employment_category when present, else inferred from position (Lecturer→academic, Officer→non-academic).',
                'Source: HR_EMPLOYEE_CRUD.',
            ].join(String.fromCharCode(10));
            return { title: 'Live · Academic vs non-academic (' + groups.academic.length + ' / ' + groups['non-academic'].length + ')', body };
        }

        // —— Employees (gender / general) ——
        if (topic === 'employee' || filters.gender || /\b(employee|staff|female|male|women|men|personnel)\b/.test(q)) {
            const census = typeof uuGenderCensus === 'function' ? uuGenderCensus(employees) : { total: employees.length, female: 0, male: 0, unknown: 0 };
            const filtered = typeof uuFilterEmployees === 'function' ? uuFilterEmployees(employees, filters) : employees;
            const filteredCensus = typeof uuGenderCensus === 'function' ? uuGenderCensus(filtered) : { total: filtered.length, female: 0, male: 0, unknown: 0 };
            const n = filtered.length;
            let logicNote = filters.gender
                ? ('Logic check: list length ' + n + ' vs ' + filters.gender + ' in list ' + (filters.gender === 'female' ? filteredCensus.female : filteredCensus.male) + '.')
                : '';
            const depts = {};
            filtered.forEach((e) => {
                const d = typeof uuEmpDept === 'function' ? uuEmpDept(e) : ((e && e.department) || '—');
                depts[d] = (depts[d] || 0) + 1;
            });
            const deptLines = Object.entries(depts).sort((a, b) => b[1] - a[1]).map(([d, c]) => '• ' + d + ': ' + c).join(String.fromCharCode(10));
            const lines = filtered.slice(0, 30).map((e) => {
                const id = typeof uuEmpId === 'function' ? uuEmpId(e) : (e.id_number || '—');
                const name = typeof uuEmpName === 'function' ? uuEmpName(e) : (e.full_name || '—');
                const g = typeof uuNormalizeGender === 'function' ? uuNormalizeGender(e.gender) : String(e.gender || '');
                const pos = typeof uuEmpPos === 'function' ? uuEmpPos(e) : (e.current_position || '—');
                const dept = typeof uuEmpDept === 'function' ? uuEmpDept(e) : (e.department || '—');
                return '• ' + id + ' | ' + name + ' | ' + g + ' | ' + pos + ' | ' + dept;
            });
            const filterBits = [];
            if (filters.gender) filterBits.push('gender=' + filters.gender);
            if (filters.department) filterBits.push('department~' + filters.department);
            if (filters.location) filterBits.push('location~' + filters.location);
            // mini category split on filtered set
            let ac = 0, nac = 0;
            filtered.forEach((e) => { const c = empCategory(e); if (c === 'academic') ac++; else if (c === 'non-academic') nac++; });
            let head = 'Employees (live HR data)\n• Matching: ' + n + '\n• Total in master: ' + census.total + '\n• Full split: female ' + census.female + ' · male ' + census.male + ' · unspecified ' + census.unknown + '\n• In this set — academic ' + ac + ' · non-academic ' + nac + '\n';
            if (n === 0) {
                return { title: filters.gender ? ('Live · ' + filters.gender + ' employees') : 'Live · HR employees', body: head + '• Filters: ' + (filterBits.join(', ') || 'none') + '\n\nNo rows match.\n' + logicNote + '\nSource: HR_EMPLOYEE_CRUD.' };
            }
            return {
                title: filters.gender ? ('Live · ' + filters.gender + ' employees (' + n + ')') : ('Live · HR employees (' + n + ')'),
                body: head + '• Active filters: ' + (filterBits.join(', ') || 'none') + '\n\nBy department:\n' + (deptLines || '• —') + '\n\nPeople (' + Math.min(n, 30) + (n > 30 ? ' of ' + n : '') + '):\n' + lines.join(String.fromCharCode(10)) + String.fromCharCode(10,10) + logicNote + '\nGender uses the record field, corrected from first name when the name is a known male/female Ethiopian name.\nSource: HR_EMPLOYEE_CRUD.',
            };
        }

        // —— Transactions list ——
        if (topic === 'tx' || /\b(transaction|journal|list of transaction)\b/.test(q)) {
            const n = tx.length;
            if (!n) {
                return { title: 'Live · Transactions (0)', body: 'No transaction/journal rows in session data.\nPost single transactions, day journals, payroll, or Material GRV so they appear here.\nSources tried: transactions, dayJournals.' };
            }
            let sumD = 0, sumC = 0;
            const byYear = {};
            tx.forEach((r) => {
                sumD += txDebit(r);
                sumC += txCredit(r);
                const y = txYear(r);
                byYear[y] = (byYear[y] || 0) + 1;
            });
            const yearLines = Object.entries(byYear).sort((a, b) => String(b[0]).localeCompare(String(a[0]))).map(([y, c]) => '• ' + y + ': ' + c + ' line(s)').join(String.fromCharCode(10));
            const lines = tx.slice(0, 25).map((r, i) => {
                const d = txDebit(r);
                const c = txCredit(r);
                return (i + 1) + '. ' + txDate(r) + ' | ' + txAcc(r) + ' | ' + (txName(r) || '—').slice(0, 40) + ' | Dr ' + money(d) + ' | Cr ' + money(c);
            });
            return {
                title: 'Live · Transactions (' + n + ')',
                body: [
                    'Finance transaction / journal lines (live):',
                    '• Lines: ' + n,
                    '• Total debit: ' + money(sumD) + ' ETB',
                    '• Total credit: ' + money(sumC) + ' ETB',
                    '• Imbalance (Dr − Cr): ' + money(sumD - sumC) + ' ETB',
                    '',
                    'By year:',
                    yearLines || '• —',
                    '',
                    'Sample lines (up to 25):',
                    lines.join(String.fromCharCode(10)),
                    '',
                    'Dashboard tip: ask “trial balance” or “what profit did we achieve this year?” for analysed views.',
                    'Source: transactions / dayJournals.',
                ].join(String.fromCharCode(10)),
            };
        }

        // —— Trial balance ——
        if (topic === 'tb' || /\b(trial\s*balance|\btb\b)\b/.test(q)) {
            const rows = buildTb();
            if (!rows.length) {
                return { title: 'Live · Trial balance (empty)', body: 'No movement to build TB. Need posted transactions/journals.\nCOA rows available: ' + coa.length + ' (structure only; balances come from transactions).' };
            }
            let td = 0, tc = 0;
            rows.forEach((r) => { td += r.debit; tc += r.credit; });
            const lines = rows.slice(0, 40).map((r) => {
                const bal = r.debit - r.credit;
                const side = bal >= 0 ? ('Dr ' + money(bal)) : ('Cr ' + money(-bal));
                return '• ' + r.account + ' | ' + String(r.name || '').slice(0, 36) + ' | Dr ' + money(r.debit) + ' | Cr ' + money(r.credit) + ' | Net ' + side;
            });
            return {
                title: 'Live · Trial balance (' + rows.length + ' accounts)',
                body: [
                    'Trial balance built from live transactions (not COA seed zeros):',
                    '• Accounts with movement: ' + rows.length,
                    '• Total debit: ' + money(td) + ' ETB',
                    '• Total credit: ' + money(tc) + ' ETB',
                    '• Difference: ' + money(td - tc) + ' ETB ' + (Math.abs(td - tc) < 0.05 ? '(balanced ✓)' : '(review postings)'),
                    '',
                    'Accounts:',
                    lines.join(String.fromCharCode(10)),
                    rows.length > 40 ? '\n… +' + (rows.length - 40) + ' more accounts' : '',
                    '',
                    'Source: sum of transaction debits/credits by account_code; names enriched from FIN_COA_CRUD when present.',
                ].join(String.fromCharCode(10)),
            };
        }

        // —— Profit / P&L current year ——

        if (topic === 'chairs' || (/\bchairs?\b/.test(q) && /\b(store|inventory|sku|how many|stock|furniture)\b/.test(q)) || topic === 'furniture') {
            const rows = Array.isArray(items) ? items : [];
            const matchWord = topic === 'chairs' || /\bchairs?\b/.test(q) ? /chair|stool|seat/i : /chair|desk|table|furniture|sofa|cabinet|stool/i;
            const hit = rows.filter((r) => matchWord.test(JSON.stringify(r || {})));
            let qtySum = 0;
            const lines = hit.slice(0, 20).map((r) => {
                const sku = (r && (r.sku || r.item_sku || r.item_code || r.code || r.id)) || '—';
                const name = (r && (r.name || r.item_name || r.description)) || '';
                const qty = Number(r && (r.quantity != null ? r.quantity : r.qty)) || 0;
                qtySum += qty;
                const store = (r && (r.location || r.store || r.warehouse || r.receiving_store)) || '—';
                return '• ' + sku + (name ? ' — ' + name : '') + ' | qty ' + qty + ' | ' + store;
            });
            // If filtering chairs only, recompute qtySum from chair rows only
            if (/\bchairs?\b/.test(q)) {
                qtySum = hit.reduce((s, r) => s + (Number(r && (r.quantity != null ? r.quantity : r.qty)) || 0), 0);
            }
            const nl = String.fromCharCode(10);
            const body = [
                'Inventory interpretation (Material master — not database tables):',
                '• Matching SKU types: ' + hit.length,
                '• Total quantity on those SKUs: ' + qtySum,
                '• All inventory SKU rows: ' + rows.length,
                '',
                'Matching items:',
                lines.join(nl) || '• No matching furniture/chair SKU in MAT_ITEM_CRUD.',
                '',
                'Logical note: “how many chairs in store” sums quantity fields on chair-related SKUs.',
                'Source: MAT_ITEM_CRUD / inventory master.',
            ].join(nl);
            return { title: 'Live · Store items (' + hit.length + ' SKU · qty ' + qtySum + ')', body: body, topic: 'chairs', source: 'session' };
        }


        if (topic === 'salary_total' || (/\b(basic salary|total basic|sum of salary|total salary|payroll basic)\b/.test(q)) || (/\bsalary\b/.test(q) && /\b(total|sum|how much|aggregate)\b/.test(q))) {
            const rows = Array.isArray(employees) ? employees : [];
            let total = 0;
            let n = 0;
            const lines = [];
            rows.forEach((r) => {
                const sal = Number(r && (r.current_salary != null ? r.current_salary : (r.basic_salary != null ? r.basic_salary : (r.salary != null ? r.salary : r.basic)))) || 0;
                if (sal > 0) n += 1;
                total += sal;
            });
            rows.slice(0, 15).forEach((r) => {
                const who = (r && (r.full_name || r.name || r.employee_name || r.id_no || r.id)) || '—';
                const sal = Number(r && (r.current_salary != null ? r.current_salary : (r.basic_salary != null ? r.basic_salary : (r.salary != null ? r.salary : r.basic)))) || 0;
                lines.push('• ' + who + ' · ' + (sal ? sal.toLocaleString('en-ET', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0.00') + ' ETB');
            });
            const nl = String.fromCharCode(10);
            const money = (v) => {
                try { return Number(v).toLocaleString('en-ET', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
                catch (e) { return String(v); }
            };
            const body = [
                'Total basic / current salary from HR employee master (live):',
                '• Employees in master: ' + rows.length,
                '• Employees with salary > 0: ' + n,
                '• Sum of current/basic salary: ' + money(total) + ' ETB',
                '',
                'Sample (up to 15):',
                lines.join(nl) || '• No employee rows.',
                '',
                'Logical note: this is the HR master total, not necessarily the latest payroll process batch net pay.',
                'Moral note: salary figures are sensitive — only visible according to your role access in the system.',
                'Source: HR_EMPLOYEE_CRUD (current_salary / basic_salary / salary).',
            ].join(nl);
            return { title: 'Live · Total basic salary (' + money(total) + ' ETB)', body: body, topic: 'salary_total', source: 'session' };
        }


        if (topic === 'profit' || /\b(profit|loss|p&l|pnl|surplus|net income)\b/.test(q)) {
            const yMatch = q.match(/\b(20\d{2})\b/);
            const year = yMatch ? parseInt(yMatch[1], 10) : yearNow;
            let revenue = 0, expense = 0;
            const revLines = [];
            const expLines = [];
            const tb = buildTb();
            // Prefer year-filtered txs
            const yearTx = (tx || []).filter((r) => txYear(r) === year);
            const useRows = yearTx.length ? yearTx : (tx || []);
            const usedYearFilter = yearTx.length > 0;
            useRows.forEach((r) => {
                const code = txAcc(r);
                const coaRow = (coa || []).find((c) => String(c.account_no || c.account_code || '').replace(/^acc-/i, '') === code);
                const kind = classifyAcc(code, coaRow && (coaRow.account_type || coaRow.type));
                const d = txDebit(r);
                const c = txCredit(r);
                if (kind === 'revenue') {
                    // revenue normally credit-balance
                    const amt = c - d;
                    revenue += amt;
                    if (Math.abs(amt) > 0.001) revLines.push({ code, name: txName(r) || (coaRow && coaRow.name) || code, amt });
                } else if (kind === 'expense') {
                    const amt = d - c;
                    expense += amt;
                    if (Math.abs(amt) > 0.001) expLines.push({ code, name: txName(r) || (coaRow && coaRow.name) || code, amt });
                }
            });
            // Aggregate duplicate accounts in samples
            const agg = (arr) => {
                const m = {};
                arr.forEach((x) => {
                    const k = x.code;
                    if (!m[k]) m[k] = { code: k, name: x.name, amt: 0 };
                    m[k].amt += x.amt;
                });
                return Object.values(m).sort((a, b) => Math.abs(b.amt) - Math.abs(a.amt));
            };
            const revA = agg(revLines).slice(0, 12);
            const expA = agg(expLines).slice(0, 12);
            const profit = revenue - expense;
            const body = [
                'Profit / loss analysis (live transactions' + (usedYearFilter ? ', year ' + year : ', all dated lines — few/no rows tagged ' + year) + '):',
                '• Revenue (Cr − Dr on income accounts): ' + money(revenue) + ' ETB',
                '• Expense (Dr − Cr on expense accounts): ' + money(expense) + ' ETB',
                '• Net ' + (profit >= 0 ? 'profit' : 'loss') + ': ' + money(Math.abs(profit)) + ' ETB',
                '',
                'Top revenue accounts:',
                revA.length ? revA.map((r) => '• ' + r.code + ' ' + String(r.name).slice(0, 40) + ' · ' + money(r.amt)).join(String.fromCharCode(10)) : '• — none classified',
                '',
                'Top expense accounts:',
                expA.length ? expA.map((r) => '• ' + r.code + ' ' + String(r.name).slice(0, 40) + ' · ' + money(r.amt)).join(String.fromCharCode(10)) : '• — none classified',
                '',
                'Classification uses COA type when available, else account code head (4xxx≈revenue, 5–6xxx≈expense).',
                'This is an operational snapshot from session postings — not a signed statutory statement.',
                'Transactions used: ' + useRows.length + ' · TB accounts overall: ' + tb.length,
            ].join(String.fromCharCode(10));
            return { title: 'Live · P&L snapshot ' + year + ' (' + (profit >= 0 ? 'profit' : 'loss') + ' ' + money(Math.abs(profit)) + ')', body };
        }

        // —— SKU / inventory ——
        
        // —— 

        
        // —— Generic store item analysis (question paper, chairs, toner, …) ——
        if (topic === 'institutional' || /\b(campus|campuses|president|vice[ -]?president|chairman|founder)\b/.test(q)) {
            // fall through — university facts from knowledge library
        } else if (
            topic === 'store_item' || topic === 'chairs' || topic === 'tables' || topic === 'furniture' ||
            (/\b(in store|in stock|inventory|warehouse)\b/.test(q) && /\b(how many|how much|quantity|qty)\b/.test(q) && !/\b(campus|employee|staff|salary)\b/.test(q))
        ) {
            const analyzed = typeof uuAnalyzeStoreSubject === 'function' ? uuAnalyzeStoreSubject(q, items) : null;
            if (analyzed && analyzed.body) return analyzed;
        }



        if (topic === 'tables' || (/\\btable/.test(q) && /\\b(store|sku|inventory|furniture|how many)\\b/.test(q))) {
            const furniture = (items || []).filter((r) => {
                const blob = JSON.stringify(r || {}).toLowerCase();
                return /table|desk|chair|furniture|sofa|cabinet/.test(blob);
            });
            const tablesOnly = (items || []).filter((r) => {
                const blob = JSON.stringify(r || {}).toLowerCase();
                return /\\btable\\b|desk/.test(blob);
            });
            const lines = tablesOnly.slice(0, 15).map((r) => {
                const sku = (r && (r.sku || r.item_sku || r.item_code || r.code || r.id)) || '—';
                const name = (r && (r.name || r.item_name || r.description)) || '';
                const qty = (r && (r.quantity != null ? r.quantity : r.qty));
                return '• ' + sku + (name ? ' — ' + name : '') + ' | qty ' + (qty ?? '—');
            });
            const storeSet = {};
            (items || []).forEach((r) => {
                const s = (r && (r.location || r.store || r.warehouse || r.receiving_store)) || 'Main store';
                storeSet[s] = true;
            });
            return {
                title: 'Live · Tables / furniture in store (' + tablesOnly.length + ' SKU types)',
                body: [
                    'Interpreted “tables in store” as furniture inventory (not database tables).',
                    '• Table/desk SKU types: ' + tablesOnly.length,
                    '• Broader furniture SKU types (chair/desk/table/cabinet…): ' + furniture.length,
                    '• Total inventory SKU rows: ' + items.length,
                    '• Distinct store locations: ' + Object.keys(storeSet).length + ' (' + Object.keys(storeSet).join(', ') + ')',
                    '',
                    'Matching table/desk items:',
                    lines.join(String.fromCharCode(10)) || '• No SKU name contains table/desk — check Material master descriptions.',
                    '',
                    'Also try: How many SKUs? · GRV count · inventory by store',
                    'Source: MAT_ITEM_CRUD.',
                ].join(String.fromCharCode(10)),
            };
        }

        // —— Awards ——
                if (topic === 'award' || /\b(award|awards|awarded|recognition)\b/.test(q)) {
            const rows = Array.isArray(awards) ? awards : [];
            const n = rows.length;
            const lines = rows.slice(0, 25).map((r, i) => {
                const who = (r && (r.employee_name || r.full_name || r.name || r.staff_name || r.employee_id || r.id_no)) || '—';
                const title = (r && (r.award_title || r.title || r.award_type || r.type || r.category)) || 'Award';
                const when = (r && (r.date || r.award_date || r.created_at || r.year)) || '';
                return '• ' + (i + 1) + '. ' + who + ' — ' + title + (when ? ' (' + when + ')' : '');
            });
            const body = [
                'Awards / recognition (live HR data):',
                '• Exact count: ' + n,
                '• Employees in master (for context): ' + employees.length,
                '',
                n ? 'Records:' : 'No award rows in session.',
                lines.join(String.fromCharCode(10)) || '• Seed HR_AWARD_CRUD or add awards under HR → Awards.',
                '',
                'Source: HR_AWARD_CRUD.',
            ].join(String.fromCharCode(10));
            return { title: 'Live · Awards (' + n + ')', body: body, topic: 'award', source: 'session' };
        }

        if (topic === 'penalty' || /\\bpenalt/.test(q)) {
            return { title: 'Live · Penalties (' + penalties.length + ')', body: 'Penalty records: ' + penalties.length + '.\\nSource: HR_PENALTY_CRUD.' };
        }
        if (topic === 'training' || /\\btraining\\b/.test(q)) {
            return { title: 'Live · Training (' + training.length + ')', body: 'Training records: ' + training.length + '.\\nSource: HR_TRAINING_CRUD.' };
        }
        if (topic === 'appraisal' || /\\bappraisal\\b/.test(q)) {
            return { title: 'Live · Appraisals (' + appraisals.length + ')', body: 'Appraisal records: ' + appraisals.length + '.\\nSource: HR_APPRAISAL_CRUD.' };
        }

if (topic === 'sku' || /\b(sku|skus|inventory|stock|tables?)\b/.test(q)) {
            const n = items.length;
            const lines = items.slice(0, 15).map((r) => {
                const sku = (r && (r.sku || r.item_sku || r.item_code || r.code || r.id)) || '—';
                const name = (r && (r.name || r.item_name || r.description)) || '';
                return '• ' + sku + (name ? ' — ' + name : '') + ' | qty ' + ((r && (r.quantity != null ? r.quantity : r.qty)) ?? '—');
            });
            const stores = {};
            let qtySum = 0;
            items.forEach((r) => {
                const s = (r && (r.location || r.store || r.warehouse || r.receiving_store)) || 'Main store';
                stores[s] = (stores[s] || 0) + 1;
                qtySum += Number(r && (r.quantity != null ? r.quantity : r.qty)) || 0;
            });
            const storeLines = Object.entries(stores).sort((a, b) => b[1] - a[1]).map(([s, c]) => '• ' + s + ': ' + c + ' item row(s)').join(String.fromCharCode(10));
            return {
                title: 'Live · Store / SKUs (' + n + ')',
                body: n
                    ? ('Inventory dashboard:\n• SKU rows: ' + n + '\n• Qty sum (numeric fields): ' + qtySum + '\n• GRV docs: ' + grv.length + ' · GIV docs: ' + giv.length + '\n\nBy store:\n' + (storeLines || '• —') + '\n\nSample SKUs:\n' + lines.join(String.fromCharCode(10)) + '\nSource: MAT_ITEM_CRUD.')
                    : 'No inventory rows in session.',
            };
        }

        if (topic === 'grv' || /\bgrv\b/.test(q)) {
            const n = grv.length;
            let value = 0;
            const lines = grv.slice(0, 12).map((r) => {
                const v = Number(r.line_value || r.total || r.amount || ((r.qty_received || r.qty || 0) * (r.unit_cost || 0))) || 0;
                value += v;
                return '• ' + (r.grv_number || r.grv_no || r.id || '—') + ' | ' + String(r.item_name || r.item || r.description || '').slice(0, 32) + ' | ' + money(v) + ' | ' + (r.status || '—');
            });
            // if value still 0 recount
            if (value === 0) grv.forEach((r) => { value += Number(r.line_value || r.total || r.amount || 0) || 0; });
            return {
                title: 'Live · GRV (' + n + ')',
                body: 'Goods received dashboard:\n• Documents: ' + n + '\n• Approx line value sum: ' + money(value) + ' ETB\n• Linked inventory SKUs: ' + items.length + '\n\nSample:\n' + (lines.join(String.fromCharCode(10)) || '• —') + '\nSource: MAT_GRV_CRUD.',
            };
        }
        if (topic === 'giv' || /\bgiv\b/.test(q)) {
            return { title: 'Live · GIV (' + giv.length + ')', body: 'Goods issued documents: ' + giv.length + '.\nInventory SKUs: ' + items.length + '.\nSource: MAT_GIV_CRUD.' };
        }
                if (topic === 'leave' || (/\bleave\b/.test(q) && !/\b(employee master|staff master)\b/.test(q))) {
            const rows = Array.isArray(leave) ? leave : [];
            const activeStatuses = /approved|active|ongoing|on\s*leave|taken|current/i;
            const active = rows.filter((r) => {
                const st = String((r && (r.status || r.leave_status || r.state)) || '');
                if (activeStatuses.test(st)) return true;
                // date window: start <= today <= end when dates exist
                try {
                    const a = r.start_date || r.from_date || r.leave_from;
                    const b = r.end_date || r.to_date || r.leave_to;
                    if (a && b) {
                        const t0 = new Date(a).getTime();
                        const t1 = new Date(b).getTime();
                        const now = Date.now();
                        if (!isNaN(t0) && !isNaN(t1) && t0 <= now && now <= t1) return true;
                    }
                } catch (e) {}
                return false;
            });
            const lines = (active.length ? active : rows).slice(0, 20).map((r) => {
                const who = (r && (r.employee_name || r.full_name || r.name || r.staff_name || r.employee_id || r.id_no)) || '—';
                const typ = (r && (r.leave_type || r.type || r.category)) || 'Leave';
                const st = (r && (r.status || r.leave_status)) || '—';
                const days = (r && (r.days != null ? r.days : r.no_of_days));
                return '• ' + who + ' · ' + typ + ' · ' + st + (days != null ? ' · ' + days + ' day(s)' : '');
            });
            const body = [
                'Leave records in HR (live session / DB bootstrap):',
                '• Total leave rows: ' + rows.length,
                '• Currently on leave / approved-active: ' + active.length,
                '• Employees in master: ' + employees.length,
                '',
                active.length ? 'Active / on-leave sample:' : 'All leave rows (none marked active by status/dates):',
                lines.join(String.fromCharCode(10)) || '• No leave rows — open HR → Leave or seed HR_LEAVE_CRUD.',
                '',
                'Source: HR_LEAVE_CRUD (and status/date fields when present).',
            ].join(String.fromCharCode(10));
            return { title: 'Live · Leave (' + rows.length + ' rows · ' + active.length + ' active)', body: body, topic: 'leave', source: 'session' };
        }

                if (topic === 'appraisal' || /\b(appraisals?|performance review|competence)\b/.test(q)) {
            const rows = Array.isArray(appraisals) ? appraisals : [];
            const n = rows.length;
            const lines = rows.slice(0, 20).map((r, i) => {
                const who = (r && (r.employee_name || r.full_name || r.name || r.employee_id || r.id_no)) || '—';
                const score = (r && (r.score != null ? r.score : (r.total_score != null ? r.total_score : r.rating)));
                const period = (r && (r.period || r.appraisal_period || r.year || r.cycle)) || '';
                const st = (r && (r.status || r.state)) || '';
                return '• ' + (i + 1) + '. ' + who + (score != null ? ' · score ' + score : '') + (period ? ' · ' + period : '') + (st ? ' · ' + st : '');
            });
            const body = [
                'Appraisal / performance records (live HR data):',
                '• Exact count: ' + n,
                '• Employees in master: ' + (Array.isArray(employees) ? employees.length : 0),
                '',
                n ? 'Records:' : 'No appraisal rows in session.',
                lines.join(String.fromCharCode(10)) || '• Open HR → Appraisal / Competence or seed HR_APPRAISAL_CRUD.',
                '',
                'Source: HR_APPRAISAL_CRUD.',
            ].join(String.fromCharCode(10));
            return { title: 'Live · Appraisals (' + n + ')', body: body, topic: 'appraisal', source: 'session' };
        }

        if (topic === 'capital' || (/\bcapital\b/.test(q) && /\b(how much|total|current|now|balance|share)\b/.test(q))) {
            const coa = (uuPull(allData, getCollection, 'FIN_COA_CRUD') || []).length
                ? uuPull(allData, getCollection, 'FIN_COA_CRUD')
                : (uuPull(allData, getCollection, 'chartsOfAccounts') || []);
            const txs = [].concat(
                uuPull(allData, getCollection, 'transactions') || [],
                uuPull(allData, getCollection, 'dayJournals') || [],
                uuPull(allData, getCollection, 'FIN_TRANSACTIONS_CRUD') || []
            );
            const ledger = {};
            const clean = (id) => (id == null ? '' : String(id).replace(/^acc-/i, '').trim());
            (Array.isArray(coa) ? coa : []).forEach((acc) => {
                const no = clean(acc.account_no || acc.account_code || acc.code || acc.id);
                if (!no) return;
                const typ = String(acc.account_type || acc.type || acc.category || '').toLowerCase();
                const isEq = /equity|capital|share/.test(typ) || /^3/.test(no);
                if (!isEq) return;
                ledger[no] = { no: no, name: acc.name || acc.desc || acc.description || no, credit: 0, debit: 0 };
            });
            (Array.isArray(txs) ? txs : []).forEach((tx) => {
                const no = clean(tx.account_code || tx.account_no || tx.account_id || tx.accountId);
                if (!no) return;
                if (!ledger[no] && /^3/.test(no)) {
                    ledger[no] = { no: no, name: tx.account_name || no, credit: 0, debit: 0 };
                }
                if (!ledger[no]) return;
                const d = Number(tx.debit != null ? tx.debit : (tx.type === 'Debit' ? tx.amount : 0)) || 0;
                const c = Number(tx.credit != null ? tx.credit : (tx.type === 'Credit' ? tx.amount : 0)) || 0;
                ledger[no].debit += d;
                ledger[no].credit += c;
            });
            const rows = Object.values(ledger).map((r) => Object.assign({}, r, { balance: (r.credit || 0) - (r.debit || 0) }));
            rows.sort((a, b) => String(a.no).localeCompare(String(b.no)));
            const total = rows.reduce((s, r) => s + (Number(r.balance) || 0), 0);
            const money = (n) => {
                const v = Number(n) || 0;
                try { return v.toLocaleString('en-ET', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
                catch (e) { return String(Math.round(v * 100) / 100); }
            };
            const lines = rows.filter((r) => Math.abs(r.balance) > 0.001).slice(0, 20).map((r) => '• ' + r.no + ' · ' + r.name + ' · ' + money(r.balance) + ' ETB');
            const body = [
                'Capital / equity snapshot (live COA + transactions):',
                '• Total equity / capital (credit − debit on equity accounts): ' + money(total) + ' ETB',
                '• Equity accounts touched: ' + rows.length,
                '• Transaction lines scanned: ' + txs.length,
                '',
                'By equity account:',
                lines.join(String.fromCharCode(10)) || '• No equity balances — post capital journals or seed opening balances.',
                '',
                'Source: FIN_COA_CRUD + transactions / dayJournals.',
            ].join(String.fromCharCode(10));
            return { title: 'Live · Capital / equity (' + money(total) + ' ETB)', body: body, topic: 'capital', source: 'session' };
        }

if (topic === 'partner' || /\b(partner|supplier|customer)\b/.test(q)) {
            const types = {};
            partners.forEach((p) => {
                const t = String((p && (p.type || p.partner_type || p.category)) || 'Other');
                types[t] = (types[t] || 0) + 1;
            });
            const typeLines = Object.entries(types).map(([k, v]) => '• ' + k + ': ' + v).join(String.fromCharCode(10));
            const sample = partners.slice(0, 10).map((p) => '• ' + (p.name || p.partner_name || p.company || p.id) + ' | ' + (p.type || '—')).join(String.fromCharCode(10));
            return { title: 'Live · Partners (' + partners.length + ')', body: 'Partner directory:\n• Count: ' + partners.length + '\nBy type:\n' + (typeLines || '• —') + '\n\nSample:\n' + (sample || '• —') + '\nSource: FIN_PARTNERS_CRUD.' };
        }
        if (topic === 'coa' || /\bcoa\b/.test(q) || /chart of account/.test(q)) {
            const types = {};
            coa.forEach((c) => {
                const t = String((c && (c.account_type || c.type || c.category)) || 'General');
                types[t] = (types[t] || 0) + 1;
            });
            const typeLines = Object.entries(types).sort((a, b) => b[1] - a[1]).map(([k, v]) => '• ' + k + ': ' + v).join(String.fromCharCode(10));
            return { title: 'Live · COA (' + coa.length + ')', body: 'Chart of accounts:\n• Rows: ' + coa.length + '\nBy type:\n' + (typeLines || '• —') + '\n\nTB accounts with movement: ' + buildTb().length + '.\nSource: FIN_COA_CRUD.' };
        }
        if (topic === 'payroll' || /\bpayroll\b/.test(q)) {
            let gross = 0, net = 0;
            process.forEach((r) => {
                gross += Number(r.gross || r.gross_amount || r.total_basic || r.basic_salary || 0) || 0;
                net += Number(r.net || r.net_pay || r.net_salary || 0) || 0;
            });
            return {
                title: 'Live · Payroll process (' + process.length + ')',
                body: 'Payroll process dashboard:\n• Rows: ' + process.length + '\n• Approx gross sum: ' + money(gross) + ' ETB\n• Approx net sum: ' + money(net) + ' ETB\n• HR employees available: ' + employees.length + '\nSource: PAY_PROCESS_CRUD.',
            };
        }

        return null;
    } catch (err) {
        return null;
    }
};

const unityUniversityAnswer = (question, allData, getCollection, priorTopic, priorQuestion) => {
    try {
        const raw = String(question || '').trim();
        const resolved = uuResolveQuery(raw, priorTopic, priorQuestion);
        // Institutional questions (campuses, president, VP) → knowledge, not store
        const instForce = /\b(campus|campuses|president|vice[ -]?president|\bvp\b|chairman|founder|who (is|was) the|leadership)\b/i.test(resolved.query || raw);
        if (instForce) {
            resolved.topic = resolved.topic || 'institutional';
        }

        // Dictionary / term definitions first when asked
        try {
            const dictHit = typeof uuDictionaryAnswer === 'function' ? uuDictionaryAnswer(resolved.query || raw) : null;
            if (dictHit && dictHit.body && /\b(what is|what\'s|define|meaning|explain)\b/i.test(resolved.query || raw)) {
                return dictHit;
            }
        } catch (eDict) {}
        // Quantitative / inventory questions: never answer with generic module essays
        const quantQ = (/\b(how many|how much|total|sum|count|number of|quantity)\b/i.test(resolved.query || raw)
            || /\b(chairs?|tables?|sku|grv|giv|award|appraisal|leave|salary|capital|profit|transaction|paper|toner)\b/i.test(resolved.query || raw))
            && !/\b(campus|campuses|president|vice[ -]?president|chairman|founder)\b/i.test(resolved.query || raw);

        if (!raw) {
            return { title: 'unityUniversity AI', body: 'Ask about Unity University or live SIS data:\n• What is Unity University? · campuses · programs · history\n• Finance / HR / Material modules\n• how many female employees · how many SKUs · GRV count', topic: '' };
        }
        const instQ = /\b(unity university|history of|who founded|founder|president|chairman|campus|campuses|gerji|adama|dessie|burayu|program|faculty|mba|admission|registrar|motto|vision|mission|accreditation|uu\.edu|sis\.uu|when was|when were|established|founded|constructed|construction|built|private university)\b/i.test(resolved.query);
        const live = (!instQ || resolved.topic) ? unityUniversityLiveAnswer(resolved.query, allData, getCollection, resolved.topic, resolved.filters) : null;
        if (live && live.body && !(typeof instForce !== 'undefined' && instForce) && String(live.topic || '') !== 'institutional' && String(live.title || '').indexOf('Store analysis') < 0) return Object.assign({}, live, { topic: live.topic || resolved.topic || priorTopic || '', source: live.source || 'session' });
        if (typeof quantQ !== 'undefined' && quantQ && !(live && live.body)) {
            return {
                title: 'Live data · no matching rows',
                body: 'I treated this as a live count/amount question, but no matching session rows were found for: ' + (resolved.query || raw) + String.fromCharCode(10) +
                    'Check that HR/Material/Finance collections are loaded (bootstrap/API) and field names contain the item (e.g. chair in SKU description).' + String.fromCharCode(10) +
                    'I will not invent numbers. Moral rule: only report data present in collections you can access.',
                topic: resolved.topic || '',
                source: 'session',
            };
        }

        const q = String(resolved.query || '').toLowerCase();
        const words = q.split(/[^a-z0-9]+/).filter((w) => w.length > 2);
        const knowledge = Array.isArray(UU_AI_KNOWLEDGE) ? UU_AI_KNOWLEDGE : [];
        const scored = [];
        for (let i = 0; i < knowledge.length; i++) {
            const item = knowledge[i];
            if (!item) continue;
            let score = 0;
            const titleL = String(item.title || '').toLowerCase();
            const bodyL = String(item.body || '').toLowerCase();
            if (q.indexOf(titleL) >= 0) score += 8;
            const tags = Array.isArray(item.tags) ? item.tags : [];
            for (let ti = 0; ti < tags.length; ti++) {
                const tag = String(tags[ti] || '').toLowerCase();
                if (!tag) continue;
                if (q.indexOf(tag) >= 0) score += 5;
                for (let w = 0; w < words.length; w++) {
                    const word = words[w];
                    if (tag === word) score += 4;
                    else if (tag.indexOf(word) >= 0 || word.indexOf(tag) >= 0) score += 1.5;
                }
            }
            for (let w = 0; w < words.length; w++) {
                const word = words[w];
                if (titleL.indexOf(word) >= 0) score += 2.5;
                if (bodyL.indexOf(word) >= 0) score += 0.35;
            }
            if (/unity|university|campus|gerji|adama|dessie|burayu|program|faculty|mba|admission|registrar|history|founder|president/.test(q)) {
                if (String(item.id || '').indexOf('uu_') === 0 || String(item.id || '').indexOf('sis_') === 0) score += 1.5;
            }
            if (score >= 1.2) scored.push({ item: item, score: score });
        }
        scored.sort(function(a, b) { return b.score - a.score; });
        if (scored.length && scored[0].score >= 1.5) {
            const top = [];
            for (let i = 0; i < scored.length && i < 3; i++) {
                if (i === 0 || scored[i].score >= Math.max(2, scored[0].score * 0.45)) top.push(scored[i]);
            }
            if (top.length === 1) {
                return { title: top[0].item.title, body: top[0].item.body, topic: resolved.topic || top[0].item.id };
            }
            const body = top.map(function(s, idx) {
                return (idx + 1) + ') ' + s.item.title + String.fromCharCode(10) + s.item.body;
            }).join(String.fromCharCode(10,10)+'——'+String.fromCharCode(10,10));
            return {
                title: 'unityUniversity AI · ' + top.map(function(s) { return s.item.title; }).join(' · '),
                body: body,
                topic: resolved.topic || top[0].item.id,
            };
        }
        return {
            title: 'unityUniversity AI',
            body: 'No strong match in the library or live data.\nTry:\n• About Unity University / history / campuses / programs\n• Finance module · HR module · Material GRV\n• how many female employees · how many SKUs\n• contacts / registrar / sis.uu.edu.et',
            topic: resolved.topic || priorTopic || '',
        };
    } catch (err) {
        return { title: 'unityUniversity AI', body: 'Answer error: ' + String(err && err.message ? err.message : err), topic: priorTopic || '' };
    }
};

const uuFetchLiveUniversityFacts = async (question) => {
    try {
        const q = String(question || '').toLowerCase();
        const wantsWeb = /unity|university|campus|founded|establish|construct|built|history|president|program|gerji|adama|dessie|when|who|where|motto|accreditation/.test(q);
        if (!wantsWeb) return null;
        const url = 'https://en.wikipedia.org/api/rest_v1/page/summary/Unity_University';
        const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
        const timer = ctrl ? setTimeout(() => { try { ctrl.abort(); } catch (e) {} }, 4500) : null;
        const res = await fetch(url, { method: 'GET', headers: { Accept: 'application/json' }, signal: ctrl ? ctrl.signal : undefined });
        if (timer) clearTimeout(timer);
        if (!res || !res.ok) return null;
        const data = await res.json();
        const extract = String((data && (data.extract || data.description)) || '').trim();
        if (!extract) return null;
        const title = String((data && data.title) || 'Unity University');
        const page = (data && data.content_urls && data.content_urls.desktop && data.content_urls.desktop.page) || 'https://en.wikipedia.org/wiki/Unity_University';
        return {
            title: 'Live web · ' + title,
            body: extract + '\n\nSource: Wikipedia summary (refresh each ask). Official site: https://www.uu.edu.et · ' + page + '\nNote: For admissions/fees/current programs prefer uu.edu.et.',
            topic: 'uu_profile',
            source: 'web',
        };
    } catch (e) {
        return null;
    }
};

const uuFormatAiBody = (s) => {
    let t = String(s == null ? '' : s);
    // Turn literal backslash-n / backslash-t into real control chars if present
    t = t.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n').replace(/\\t/g, '  ');
    t = t.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    t = t.replace(/\n{3,}/g, '\n\n');
    return t;
};

const UnityUniversityAI = ({ glassClass = '', dockMode = false }) => {
    let allData = {};
    let getCollection = null;
    try {
        const dataApi = useData();
        allData = (dataApi && dataApi.allData) || {};
        getCollection = dataApi && dataApi.getCollection;
    } catch (e) { allData = {}; }
    try {
        if ((!allData || !Object.keys(allData).length) && typeof window !== 'undefined' && window.__UU_ALL_DATA__) allData = window.__UU_ALL_DATA__;
    } catch (e) {}

    const [question, setQuestion] = useState('');
    const [messages, setMessages] = useState([]);
    const [priorTopic, setPriorTopic] = useState('');
    const [priorQuestion, setPriorQuestion] = useState('');
    const [busy, setBusy] = useState(false);
    const threadRef = useRef(null);
    const bottomRef = useRef(null);
    const inputRef = useRef(null);
    const askSeq = useRef(0);

    const scrollChat = () => {
        window.setTimeout(() => {
            try {
                if (bottomRef.current && bottomRef.current.scrollIntoView) {
                    bottomRef.current.scrollIntoView({ behavior: 'smooth', block: 'end' });
                } else if (threadRef.current) {
                    threadRef.current.scrollTop = threadRef.current.scrollHeight;
                }
            } catch (e) {}
        }, 40);
    };

    const ask = async (q) => {
        const query = String(q != null && String(q).length ? q : question).trim();
        if (!query || busy) return;
        const seq = ++askSeq.current;
        const at = new Date().toISOString();
        const userMsg = { id: 'u-' + seq + '-' + at, role: 'user', text: query, at: at };

        setBusy(true);
        setMessages((prev) => [...(Array.isArray(prev) ? prev : []), userMsg]);
        setQuestion('');
        scrollChat();

        let ad = allData;
        let gc = getCollection;
        try { if (typeof window !== 'undefined' && window.__UU_ALL_DATA__) ad = window.__UU_ALL_DATA__; } catch (e) {}

        try {
            let res = unityUniversityAnswer(query, ad, gc, priorTopic, priorQuestion);
            try {
                const liveWeb = await uuFetchLiveUniversityFacts(query);
                if (liveWeb && liveWeb.body) {
                    if (res && res.body && String(res.title || '').indexOf('Live') !== 0) {
                        res = {
                            title: (res.title || 'unityUniversity AI') + ' · Live web note',
                            body: uuFormatAiBody(String(res.body || '') + '\n\n——\nLive web (Wikipedia summary, may lag official site):\n' + liveWeb.body),
                            topic: res.topic || liveWeb.topic || '',
                            source: 'library+web',
                        };
                    } else if (!res || !res.body) {
                        res = Object.assign({}, liveWeb, { body: uuFormatAiBody(liveWeb.body) });
                    }
                }
            } catch (webErr) { /* ignore */ }

            if (seq !== askSeq.current) return;

            const safe = (res && res.body)
                ? Object.assign({}, res, { body: uuFormatAiBody(res.body) })
                : { title: 'unityUniversity AI', body: 'No answer for: ' + query + '\nTry: how many employees · trial balance · how many SKUs · campuses', topic: '', source: 'none' };

            const botMsg = {
                id: 'a-' + seq + '-' + Date.now(),
                role: 'assistant',
                text: uuFormatAiBody(safe.body),
                title: String(safe.title || 'unityUniversity AI'),
                source: safe.source || 'session',
                topic: safe.topic || '',
                at: new Date().toISOString(),
            };
            setMessages((prev) => [...(Array.isArray(prev) ? prev : []), botMsg]);
            if (safe.topic) setPriorTopic(safe.topic);
            else setPriorTopic('');
            setPriorQuestion(query);
            scrollChat();
        } catch (err) {
            if (seq !== askSeq.current) return;
            setMessages((prev) => [...(Array.isArray(prev) ? prev : []), {
                id: 'e-' + seq,
                role: 'assistant',
                title: 'unityUniversity AI',
                text: 'Could not answer: ' + query + ' — ' + String(err && err.message ? err.message : err),
                source: 'error',
                at: new Date().toISOString(),
            }]);
            scrollChat();
        } finally {
            if (seq === askSeq.current) setBusy(false);
            try { inputRef.current && inputRef.current.focus && inputRef.current.focus(); } catch (e) {}
        }
    };

    const quick = [
        'How many employees?',
        'How many female employees?',
        'Academic vs non-academic staff?',
        'List of transactions',
        'Trial balance',
        'What profit this year?',
        'How many SKUs?',
        'How many GRV?',
        'Who was awarded?',
        'Tables in store?',
        'What is Unity University?',
        'Finance module',
        'HR module',
        'Material module',
    ];

    const timeAgo = (iso) => {
        try {
            const ms = Date.now() - new Date(iso).getTime();
            if (ms < 60000) return 'just now';
            if (ms < 3600000) return Math.floor(ms / 60000) + ' min ago';
            if (ms < 86400000) return Math.floor(ms / 3600000) + ' h ago';
            return new Date(iso).toLocaleString();
        } catch (e) { return ''; }
    };

    const renderBody = (text) => {
        const lines = uuFormatAiBody(text).split('\n');
        return (
            <div className="text-[13px] text-gray-50 leading-relaxed space-y-1 break-words">
                {lines.map((line, idx) => (
                    <div key={idx} className={line.trim() ? '' : 'h-2'}>{line || '\u00a0'}</div>
                ))}
            </div>
        );
    };

    const shellClass = dockMode
        ? 'flex flex-col h-full min-h-0 max-h-full rounded-none border-0 bg-transparent'
        : 'mt-6 rounded-[1.75rem] border border-cyan-500/30 bg-gradient-to-br from-cyan-950/50 to-slate-950/80 ' + (glassClass || '') + ' flex flex-col relative z-[80] min-h-[28rem]';

    return (
        <div className={shellClass} style={{ isolation: 'isolate' }}>
            <div className="px-4 sm:px-5 py-3 border-b border-cyan-500/20 flex flex-wrap items-center justify-between gap-2 shrink-0 bg-slate-950/80 backdrop-blur-md">
                <div className="flex items-center gap-3 min-w-0">
                    <div className="w-9 h-9 rounded-2xl bg-cyan-500/20 border border-cyan-400/40 flex items-center justify-center text-cyan-300 text-sm font-black shrink-0">AI</div>
                    <div className="min-w-0">
                        <h3 className="text-sm font-black uppercase tracking-widest text-cyan-300 truncate">unityUniversity AI</h3>
                        <p className="text-[10px] text-gray-400 font-bold uppercase tracking-wider truncate">
                            Chat · live data · library{priorTopic ? ' · ' + priorTopic : ''}
                        </p>
                    </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                    {messages.length ? (
                        <button type="button" onClick={() => { setMessages([]); setPriorTopic(''); setPriorQuestion(''); }}
                            className="text-[9px] font-black uppercase px-2 py-1 rounded-lg bg-white/5 text-gray-400 border border-white/10 hover:text-white">
                            Clear
                        </button>
                    ) : null}
                    <span className="text-[9px] font-black uppercase px-2 py-1 rounded-lg bg-cyan-500/10 text-cyan-200/90 border border-cyan-500/20">
                        {busy ? 'Thinking…' : 'Ready'}
                    </span>
                </div>
            </div>

            <div
                ref={threadRef}
                className="p-3 sm:p-4 space-y-3 flex-1 min-h-0 overflow-y-auto overscroll-contain bg-black/25"
                style={{ maxHeight: dockMode ? undefined : 'min(42vh, 420px)' }}
            >
                {!messages.length ? (
                    <div className="rounded-2xl border border-white/10 bg-black/40 p-4 text-[13px] text-gray-300 leading-relaxed">
                        Ask anything about Unity University or this SIS. Each question gets its own answer in this chat.
                        <div className="mt-2 text-gray-400 text-[12px]">Examples: how many female employees · trial balance · how many SKUs · campuses</div>
                    </div>
                ) : null}
                {messages.map((m) => (
                    <div
                        key={m.id}
                        className={
                            m.role === 'user'
                                ? 'ml-4 sm:ml-10 rounded-2xl border border-amber-400/30 bg-amber-500/10 px-3 py-2.5 shadow-lg shadow-black/20'
                                : 'mr-2 sm:mr-8 rounded-2xl border border-cyan-400/35 bg-slate-950/90 px-3 py-3 shadow-xl shadow-cyan-950/30'
                        }
                    >
                        <div className="flex items-center justify-between gap-2 mb-1.5">
                            <span className={'text-[9px] font-black uppercase tracking-widest ' + (m.role === 'user' ? 'text-amber-300' : 'text-cyan-300')}>
                                {m.role === 'user' ? 'You' : (m.title || 'unityUniversity AI')}
                            </span>
                            <span className="text-[9px] text-gray-500 font-bold whitespace-nowrap">{timeAgo(m.at)}{m.source ? ' · ' + m.source : ''}</span>
                        </div>
                        {m.role === 'user' ? (
                            <p className="text-[13px] text-gray-50 leading-relaxed break-words">{m.text}</p>
                        ) : (
                            renderBody(m.text)
                        )}
                    </div>
                ))}
                {busy ? (
                    <div className="mr-8 rounded-2xl border border-cyan-400/20 bg-cyan-500/5 px-3 py-2 text-[12px] text-cyan-200/90 font-bold animate-pulse">
                        Thinking…
                    </div>
                ) : null}
                <div ref={bottomRef} />
            </div>

            <div
                className="p-3 sm:p-4 space-y-2 border-t border-cyan-500/30 shrink-0 bg-slate-950"
                style={{ zIndex: 200, position: 'relative', paddingBottom: 'max(0.75rem, env(safe-area-inset-bottom))' }}
            >
                <div className="flex flex-wrap gap-1.5 max-h-16 overflow-y-auto">
                    {quick.map((label) => (
                        <button
                            key={label}
                            type="button"
                            disabled={busy}
                            onClick={() => ask(label)}
                            className="px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wide bg-white/5 border border-white/10 text-cyan-100/90 hover:bg-cyan-500/20 hover:border-cyan-400/40 disabled:opacity-40"
                        >
                            {label}
                        </button>
                    ))}
                </div>
                <div className="flex gap-2 items-end" style={{ zIndex: 210, position: 'relative' }}>
                    <textarea
                        ref={inputRef}
                        rows={2}
                        value={question}
                        onChange={(e) => setQuestion(e.target.value)}
                        onKeyDown={(e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault();
                                ask(question);
                            }
                        }}
                        placeholder="Type a question… (Enter to send)"
                        className="flex-1 min-w-0 rounded-xl bg-black border border-cyan-400/40 px-3 py-2.5 text-sm text-white placeholder:text-gray-500 focus:outline-none focus:border-cyan-300 focus:ring-2 focus:ring-cyan-500/30 resize-none min-h-[3rem]"
                        style={{ zIndex: 220, position: 'relative' }}
                        disabled={busy}
                    />
                    <button
                        type="button"
                        onClick={() => ask(question)}
                        disabled={busy || !String(question || '').trim()}
                        className="h-12 px-5 rounded-xl text-xs font-black uppercase tracking-widest bg-cyan-600 hover:bg-cyan-500 text-white disabled:opacity-40 shrink-0"
                        style={{ zIndex: 220, position: 'relative' }}
                    >
                        {busy ? '…' : 'Ask'}
                    </button>
                </div>
            </div>
        </div>
    );
};


const UnityUniversityAIDock = ({ glassClass = '' }) => {
    const ui = React.useContext(UIContext) || {};
    const enabled = !!(ui.aiDockEnabled || (ui.uiSettings && ui.uiSettings.aiDockEnabled));
    const [open, setOpen] = React.useState(false);
    if (!enabled) return null;
    return (
        <>
            <button
                type="button"
                onClick={() => setOpen((v) => !v)}
                className="fixed bottom-5 right-5 w-14 h-14 rounded-2xl bg-cyan-600 hover:bg-cyan-500 text-white font-black shadow-2xl shadow-cyan-900/50 border border-cyan-300/50 flex items-center justify-center text-sm"
                title="unityUniversity AI"
                style={{ zIndex: 99950 }}
            >
                AI
            </button>
            {open ? (
                <div
                    className="fixed bottom-24 right-3 sm:right-5 w-[min(440px,calc(100vw-1.25rem))] h-[min(78vh,720px)] flex flex-col overflow-hidden rounded-2xl shadow-2xl border border-cyan-500/40 bg-slate-950"
                    style={{ zIndex: 99960 }}
                >
                    <div className="flex items-center justify-between px-3 py-2 border-b border-white/10 bg-slate-950 shrink-0" style={{ zIndex: 99970 }}>
                        <span className="text-[10px] font-black uppercase tracking-widest text-cyan-300">unityUniversity AI · dock</span>
                        <button type="button" onClick={() => setOpen(false)} className="text-gray-400 hover:text-white text-xs font-bold px-2 py-1">Close</button>
                    </div>
                    <div className="flex-1 min-h-0 overflow-hidden flex flex-col">
                        <UnityUniversityAI glassClass={glassClass || ''} dockMode />
                    </div>
                </div>
            ) : null}
        </>
    );
};


/* ---- Ethiopian (Geez) ↔ Gregorian calendar helpers (Power Geez) ---- */
const ETH_MONTHS = ['Meskerem','Tikimt','Hidar','Tahsas','Tir','Yekatit','Megabit','Miazia','Genbot','Sene','Hamle','Nehasse','Pagumen'];
const ethiopianMonthNames = ETH_MONTHS;
const isEthiopianLeap = (y) => (y % 4) === 3;
const ethMonthLength = (y, m) => {
    if (m < 1 || m > 13) return 0;
    if (m === 13) return isEthiopianLeap(y) ? 6 : 5;
    return 30;
};
const gregorianToJdn = (y, m, d) => {
    const a = Math.floor((14 - m) / 12);
    const yy = y + 4800 - a;
    const mm = m + 12 * a - 3;
    return d + Math.floor((153 * mm + 2) / 5) + 365 * yy + Math.floor(yy / 4) - Math.floor(yy / 100) + Math.floor(yy / 400) - 32045;
};
const jdnToGregorian = (jdn) => {
    const a = jdn + 32044;
    const b = Math.floor((4 * a + 3) / 146097);
    const c = a - Math.floor((146097 * b) / 4);
    const d = Math.floor((4 * c + 3) / 1461);
    const e = c - Math.floor((1461 * d) / 4);
    const m = Math.floor((5 * e + 2) / 153);
    const day = e - Math.floor((153 * m + 2) / 5) + 1;
    const month = m + 3 - 12 * Math.floor(m / 10);
    const year = 100 * b + d - 4800 + Math.floor(m / 10);
    return { year, month, day };
};
const ethiopianToJdn = (y, m, d) => {
    // Ethiopian epoch: Meskerem 1, 1 = Aug 29, 8 CE (JDN 1724220.5 ≈ 1724221 civil)
    return (Math.floor(y / 4) * 1461) + ((y % 4) * 365) + ((m - 1) * 30) + d + 1723855;
};
const jdnToEthiopian = (jdn) => {
    const r = jdn - 1723856;
    const n = Math.floor(r / 1461);
    const a = r % 1461;
    const y = n * 4 + Math.floor(a / 365) - (a === 1460 ? 1 : 0);
    const dayOfYear = a === 1460 ? 365 : a % 365;
    const m = Math.floor(dayOfYear / 30) + 1;
    const d = (dayOfYear % 30) + 1;
    return { year: y + 1, month: m, day: d };
};
const gregorianToEthiopian = (y, m, d) => {
    try {
        const jdn = gregorianToJdn(Number(y), Number(m), Number(d));
        return jdnToEthiopian(jdn);
    } catch (e) {
        return { year: 0, month: 1, day: 1 };
    }
};
const ethiopianToGregorian = (y, m, d) => {
    try {
        const jdn = ethiopianToJdn(Number(y), Number(m), Number(d));
        return jdnToGregorian(jdn);
    } catch (e) {
        return { year: 1970, month: 1, day: 1 };
    }
};

const EthiopianCalendarPanel = ({ glassClass }) => {
    const now = new Date();
    const [gY, setGY] = React.useState(now.getFullYear());
    const [gM, setGM] = React.useState(now.getMonth() + 1);
    const [gD, setGD] = React.useState(now.getDate());
    const ethToday = gregorianToEthiopian(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const [eY, setEY] = React.useState(ethToday.year);
    const [eM, setEM] = React.useState(ethToday.month);
    const [eD, setED] = React.useState(ethToday.day);
    const [mode, setMode] = React.useState('g2e'); // g2e | e2g
    const [calMonth, setCalMonth] = React.useState(now.getMonth() + 1);
    const [calYear, setCalYear] = React.useState(now.getFullYear());

    const g2e = gregorianToEthiopian(gY, gM, gD);
    const e2g = ethiopianToGregorian(eY, eM, eD);

    const daysInGregorianMonth = (y, m) => new Date(y, m, 0).getDate();
    const firstWeekday = (y, m) => new Date(y, m - 1, 1).getDay();
    const dim = daysInGregorianMonth(calYear, calMonth);
    const start = firstWeekday(calYear, calMonth);
    const cells = [];
    for (let i = 0; i < start; i++) cells.push(null);
    for (let d = 1; d <= dim; d++) cells.push(d);

    return (
        <div className={`p-6 rounded-[2.5rem] border border-amber-500/20 ${glassClass} shadow-2xl`}>
            <div className="flex items-center gap-4 mb-6">
                <div className="p-3 bg-amber-500/20 rounded-2xl">
                    <Icons.Calendar className="w-6 h-6 text-amber-400" />
                </div>
                <div>
                    <h2 className="text-xl font-black uppercase italic tracking-tighter text-white">Power Geez · Geez ↔ Gregorian</h2>
                    <p className="text-[10px] text-amber-400/80 font-black uppercase tracking-[0.25em]">Power Geez · Ethiopian calendar converter</p>
                </div>
            </div>

            <div className="flex gap-2 mb-4">
                <button type="button" onClick={() => setMode('g2e')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase ${mode === 'g2e' ? 'bg-amber-500 text-black' : 'bg-white/10 text-gray-300'}`}>
                    Gregorian → Geez
                </button>
                <button type="button" onClick={() => setMode('e2g')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase ${mode === 'e2g' ? 'bg-amber-500 text-black' : 'bg-white/10 text-gray-300'}`}>
                    Geez → Gregorian
                </button>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="space-y-3 p-4 rounded-2xl bg-black/30 border border-white/10">
                    {mode === 'g2e' ? (
                        <>
                            <p className="text-[10px] font-black text-cyan-400 uppercase tracking-widest">Gregorian input</p>
                            <div className="grid grid-cols-3 gap-2">
                                <label className="text-[9px] text-gray-400 uppercase">Year<input type="number" value={gY} onChange={e => setGY(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                                <label className="text-[9px] text-gray-400 uppercase">Month<input type="number" min={1} max={12} value={gM} onChange={e => setGM(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                                <label className="text-[9px] text-gray-400 uppercase">Day<input type="number" min={1} max={31} value={gD} onChange={e => setGD(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                            </div>
                            <div className="mt-3 p-3 rounded-xl bg-amber-500/10 border border-amber-500/30">
                                <p className="text-[9px] text-amber-400/80 uppercase font-black">Ethiopian (Geez)</p>
                                <p className="text-lg font-black text-white mt-1">
                                    {g2e.day} {ethiopianMonthNames[Math.max(0, Math.min(12, g2e.month - 1))] || g2e.month} {g2e.year}
                                </p>
                                <p className="text-[10px] text-gray-400 font-mono mt-1">{g2e.year}-{String(g2e.month).padStart(2,'0')}-{String(g2e.day).padStart(2,'0')}</p>
                            </div>
                        </>
                    ) : (
                        <>
                            <p className="text-[10px] font-black text-amber-400 uppercase tracking-widest">Ethiopian input</p>
                            <div className="grid grid-cols-3 gap-2">
                                <label className="text-[9px] text-gray-400 uppercase">Year<input type="number" value={eY} onChange={e => setEY(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                                <label className="text-[9px] text-gray-400 uppercase">Month (1–13)<input type="number" min={1} max={13} value={eM} onChange={e => setEM(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                                <label className="text-[9px] text-gray-400 uppercase">Day<input type="number" min={1} max={30} value={eD} onChange={e => setED(+e.target.value)} className="w-full mt-1 bg-black/50 border border-white/15 rounded-lg px-2 py-2 text-white text-sm" /></label>
                            </div>
                            <div className="mt-3 p-3 rounded-xl bg-cyan-500/10 border border-cyan-500/30">
                                <p className="text-[9px] text-cyan-400/80 uppercase font-black">Gregorian</p>
                                <p className="text-lg font-black text-white mt-1">
                                    {e2g.day}/{e2g.month}/{e2g.year}
                                </p>
                                <p className="text-[10px] text-gray-400 font-mono mt-1">{e2g.year}-{String(e2g.month).padStart(2,'0')}-{String(e2g.day).padStart(2,'0')}</p>
                            </div>
                        </>
                    )}
                </div>

                <div className="p-4 rounded-2xl bg-black/30 border border-white/10">
                    <div className="flex items-center justify-between mb-3">
                        <p className="text-[10px] font-black text-white uppercase tracking-widest">Gregorian month</p>
                        <div className="flex gap-1">
                            <button type="button" className="px-2 py-1 rounded bg-white/10 text-xs text-white" onClick={() => {
                                if (calMonth === 1) { setCalMonth(12); setCalYear(y => y - 1); }
                                else setCalMonth(m => m - 1);
                            }}>‹</button>
                            <span className="text-xs text-cyan-300 font-bold px-2">{calMonth}/{calYear}</span>
                            <button type="button" className="px-2 py-1 rounded bg-white/10 text-xs text-white" onClick={() => {
                                if (calMonth === 12) { setCalMonth(1); setCalYear(y => y + 1); }
                                else setCalMonth(m => m + 1);
                            }}>›</button>
                        </div>
                    </div>
                    <div className="grid grid-cols-7 gap-1 text-center text-[9px] text-gray-500 mb-1">
                        {['Su','Mo','Tu','We','Th','Fr','Sa'].map(d => <div key={d}>{d}</div>)}
                    </div>
                    <div className="grid grid-cols-7 gap-1">
                        {cells.map((d, i) => {
                            if (!d) return <div key={'e'+i} />;
                            const eth = gregorianToEthiopian(calYear, calMonth, d);
                            const isToday = d === now.getDate() && calMonth === now.getMonth()+1 && calYear === now.getFullYear();
                            return (
                                <button
                                    key={d}
                                    type="button"
                                    onClick={() => { setGY(calYear); setGM(calMonth); setGD(d); setMode('g2e'); }}
                                    className={`rounded-lg p-1 border text-left min-h-[42px] ${isToday ? 'border-cyan-400 bg-cyan-500/20' : 'border-white/10 bg-white/5 hover:bg-white/10'}`}
                                >
                                    <div className="text-[11px] font-bold text-white">{d}</div>
                                    <div className="text-[8px] text-amber-400/90 truncate">{eth.day} {ethiopianMonthNames[eth.month-1]?.slice(0,3)}</div>
                                </button>
                            );
                        })}
                    </div>
                    <p className="text-[9px] text-gray-500 mt-3">Click a day to convert · Ethiopian month names shown under each Gregorian date.</p>
                </div>
            </div>
        </div>
    );
};


const UICustomization = ({ glassClass }) => {
    const { uiSettings, updateUISettings, resetToDefault } = useContext(UIContext);
    const [qrText, setQrText] = useState('');

    // --- Asset Upload Handlers ---
    const handleBackgroundUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            updateUISettings({ backgroundImage: event.target.result });
        };
        reader.readAsDataURL(file);
    };

    const handleLogoUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            updateUISettings({ logoURL: event.target.result });
        };
        reader.readAsDataURL(file);
    };

    return (
        <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <UnityUniversityAI glassClass={glassClass} />
            <div className={`p-5 rounded-[1.75rem] border border-cyan-500/25 ${glassClass} shadow-xl`}>
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                        <h3 className="text-sm font-black uppercase tracking-widest text-cyan-300">AI access on all screens</h3>
                        <p className="text-[11px] text-gray-400 mt-1">When enabled, a floating AI button appears on every page of this module (not only UI Customization).</p>
                    </div>
                    <button
                        type="button"
                        onClick={() => updateUISettings({ aiDockEnabled: !uiSettings.aiDockEnabled })}
                        className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest border transition ${uiSettings.aiDockEnabled ? 'bg-cyan-600 border-cyan-400 text-white' : 'bg-white/5 border-white/15 text-gray-300'}`}
                    >
                        {uiSettings.aiDockEnabled ? 'AI dock ON' : 'AI dock OFF'}
                    </button>
                </div>
            </div>

            <EthiopianCalendarPanel glassClass={glassClass} />

            {/* 1. THE ATMOSPHERIC ENGINE (Scale & Transparency Fixes) */}
            <div className={`p-6 rounded-[2.5rem] border border-white/10 ${glassClass} shadow-2xl`}>
                <div className="flex items-center gap-4 mb-8">
                    <div className="p-3 bg-cyan-500/20 rounded-2xl">
                        <Icons.Settings2 className="w-6 h-6 text-cyan-400" />
                    </div>
                    <div>
                        <h2 className="text-2xl font-black uppercase italic tracking-tighter text-white">System Architect</h2>
                        <p className="text-[10px] text-cyan-400 font-black uppercase tracking-[0.3em] opacity-70">Atmospheric Tuning Protocol</p>
                    </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-10">
                    {/* SCALE: Fixes the "too close" magnification */}
                    <div className="space-y-4">
                        <div className="flex justify-between items-center">
                            <span className="text-[10px] font-black uppercase text-white/40 tracking-widest flex items-center gap-2">
                                <Icons.Maximize2 size={12}/> Viewport Scale
                            </span>
                            <span className="text-xs font-mono text-cyan-400">{Math.round(uiSettings.bgScale * 100)}%</span>
                        </div>
                        <input 
                            type="range" min="0.7" max="1.3" step="0.01" 
                            value={uiSettings.bgScale}
                            onChange={(e) => updateUISettings({ bgScale: parseFloat(e.target.value) })}
                            className="w-full h-1.5 bg-white/10 rounded-lg appearance-none cursor-pointer accent-cyan-500"
                        />
                    </div>

                    {/* BLUR: Fixes the transparency depth */}
                    <div className="space-y-4">
                        <div className="flex justify-between items-center">
                            <span className="text-[10px] font-black uppercase text-white/40 tracking-widest flex items-center gap-2">
                                <Icons.Layers size={12}/> Glass Blur
                            </span>
                            <span className="text-xs font-mono text-cyan-400">{uiSettings.bgBlur}px</span>
                        </div>
                        <input 
                            type="range" min="0" max="25" 
                            value={uiSettings.bgBlur}
                            onChange={(e) => updateUISettings({ bgBlur: parseInt(e.target.value) })}
                            className="w-full h-1.5 bg-white/10 rounded-lg appearance-none cursor-pointer accent-cyan-500"
                        />
                    </div>

                    {/* EXPOSURE: Brightness Control */}
                    <div className="space-y-4">
                        <div className="flex justify-between items-center">
                            <span className="text-[10px] font-black uppercase text-white/40 tracking-widest flex items-center gap-2">
                                <Icons.Sun size={12}/> Exposure
                            </span>
                            <span className="text-xs font-mono text-cyan-400">{uiSettings.bgBrightness}%</span>
                        </div>
                        <input 
                            type="range" min="50" max="150" 
                            value={uiSettings.bgBrightness}
                            onChange={(e) => updateUISettings({ bgBrightness: parseInt(e.target.value) })}
                            className="w-full h-1.5 bg-white/10 rounded-lg appearance-none cursor-pointer accent-cyan-500"
                        />
                    </div>
                </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* 2. ASSET SLOTS (Background & Logo) */}
                <div className={`p-8 rounded-[2.5rem] border border-white/10 ${glassClass} space-y-6`}>
                    <h4 className="text-[10px] font-black uppercase tracking-[0.3em] text-cyan-400">Identity Assets</h4>
                    
                    <div className="grid grid-cols-2 gap-6">
                        {/* Background Slot */}
                        <div className="relative group aspect-square rounded-3xl border-2 border-dashed border-white/10 bg-black/40 flex flex-col items-center justify-center transition-all hover:border-cyan-500/50 cursor-pointer overflow-hidden">
                            {uiSettings.backgroundImage ? (
                                <img src={uiSettings.backgroundImage} className="absolute inset-0 w-full h-full object-cover opacity-40 group-hover:scale-110 transition-transform duration-700" alt="BG" />
                            ) : null}
                            <div className="relative z-10 flex flex-col items-center text-center p-4">
                                <Icons.Image className="text-white/20 mb-2 group-hover:text-cyan-400 transition-colors" size={32} />
                                <span className="text-[9px] font-bold uppercase text-white/40 tracking-tighter">Environment Map</span>
                            </div>
                            <input type="file" accept="image/*" onChange={handleBackgroundUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                        </div>

                        {/* Logo Slot (Added per your request) */}
                        <div className="relative group aspect-square rounded-3xl border-2 border-dashed border-white/10 bg-black/40 flex flex-col items-center justify-center transition-all hover:border-cyan-500/50 cursor-pointer overflow-hidden">
                            {uiSettings.logoURL ? (
                                <div className="absolute inset-0 flex items-center justify-center bg-black/20 p-6">
                                    <img src={uiSettings.logoURL} className="max-w-full max-h-full object-contain drop-shadow-[0_0_15px_rgba(255,255,255,0.3)]" alt="Logo" />
                                </div>
                            ) : null}
                            <div className="relative z-10 flex flex-col items-center text-center p-4">
                                <Icons.ShieldCheck className="text-white/20 mb-2 group-hover:text-cyan-400 transition-colors" size={32} />
                                <span className="text-[9px] font-bold uppercase text-white/40 tracking-tighter">Corporate Seal</span>
                            </div>
                            <input type="file" accept="image/*" onChange={handleLogoUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                        </div>
                    </div>

                    <div className="flex gap-3">
                        <button 
                            onClick={() => updateUISettings({ backgroundImage: null, logoURL: null })}
                            className="flex-1 py-4 rounded-2xl border border-red-500/20 text-red-500 text-[10px] font-black uppercase tracking-[0.2em] hover:bg-red-500 hover:text-white transition-all shadow-lg"
                        >
                            Purge All Assets
                        </button>
                    </div>
                </div>

                {/* 3. UTILITY MODULES (QR & Reset) */}
                <div className={`p-8 rounded-[2.5rem] border border-white/10 ${glassClass} flex flex-col justify-between`}>
                    <div className="space-y-6">
                        <h4 className="text-[10px] font-black uppercase tracking-[0.3em] text-cyan-400">System Utilities</h4>
                        
                        <div className="group">
                            <label className="text-[9px] font-black text-white/30 uppercase ml-1 tracking-widest">QR Transmission Key</label>
                            <div className="flex gap-2 mt-2">
                                <input 
                                    className="flex-1 h-14 bg-black/40 border border-white/5 rounded-2xl px-5 text-sm text-white focus:border-cyan-500/50 outline-none transition-all placeholder:text-white/10"
                                    placeholder="Input secure data..." autoComplete="off" spellCheck="false"
                                    value={qrText}
                                    onChange={(e) => setQrText(e.target.value)}
                                />
                                <button 
                                    onClick={() => qrText && window.open(`https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${qrText}`, '_blank')}
                                    className="h-14 w-14 flex items-center justify-center bg-indigo-600 hover:bg-indigo-500 text-white rounded-2xl transition-all shadow-xl shadow-indigo-500/20 active:scale-95"
                                >
                                    <Icons.QrCode size={24} />
                                </button>
                            </div>
                        </div>
                    </div>

                    <div className="pt-8 mt-8 border-t border-white/5">
                        <button 
                            onClick={resetToDefault}
                            className="w-full py-5 bg-white/5 hover:bg-red-500/10 text-white/40 hover:text-red-500 text-[10px] font-black uppercase rounded-2xl transition-all border border-transparent hover:border-red-500/20 flex items-center justify-center gap-3"
                        >
                            <Icons.RefreshCw size={14} /> Full System Re-Initialization
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
};

const StaffChangePassword = ({ glassClass }) => {
    const { changePassword, logout } = useAuth();
    const { getCollection, updateRecord } = useData(); // <--- Get the missing functions here
    
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [status, setStatus] = useState({ type: '', msg: '' });

    const [formData, setFormData] = useState({
        oldPassword: '',
        newPassword: '',
        confirmNewPassword: ''
    });

    const handleChange = (field, value) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setStatus({ type: '', msg: '' });

        if (formData.newPassword !== formData.confirmNewPassword) {
            setStatus({ type: 'error', msg: "New password confirmation does not match." });
            return;
        }

        setIsSubmitting(true);
        
        // PASS THE FUNCTIONS HERE:
        const result = await changePassword(
            formData.oldPassword, 
            formData.newPassword, 
            getCollection, 
            updateRecord
        );
        
        setIsSubmitting(false);

        if (result.success) {
            setStatus({ type: 'success', msg: result.message });
            setFormData({ oldPassword: '', newPassword: '', confirmNewPassword: '' });
            
            // Re-login requirement: Logout after 2 seconds
            setTimeout(() => {
                logout();
            }, 2000);
        } else {
            setStatus({ type: 'error', msg: result.message });
        }
    };

    // ... (rest of your return JSX remains exactly the same)
    return (
        <ContainerCard Icon="Lock" glassClass={glassClass}>
            <form onSubmit={handleSubmit} className="space-y-5 max-w-md mx-auto p-6">
                {/* ... existing status and input fields ... */}
                {status.msg && (
                    <div className={`p-4 rounded-xl flex items-center text-sm ${
                        status.type === 'success' ? 'bg-green-500/10 text-green-400' : 'bg-red-500/10 text-red-400'
                    }`}>
                        {status.msg}
                    </div>
                )}
                
                <Input label="Current Password" type="password" placeholder="Enter password" value={formData.oldPassword} onChange={(v)=>handleChange('oldPassword',v)} isRequired />
                <Input label="New Password" type="password" placeholder="Enter password" value={formData.newPassword} onChange={(v)=>handleChange('newPassword',v)} isRequired />
                <Input label="Confirm New Password" type="password" placeholder="Enter password" value={formData.confirmNewPassword} onChange={(v)=>handleChange('confirmNewPassword',v)} isRequired />

                <button type="submit" disabled={isSubmitting} className="w-full bg-cyan-600 p-4 rounded-xl font-bold text-white">
                    {isSubmitting ? "Processing..." : "Confirm Password Change"}
                </button>
            </form>
        </ContainerCard>
    );
};
const StaffUpdateProfile = ({ glassClass }) => {
    const { currentUser } = useAuth();
    const { getCollection, updateRecord, addRecord } = useData();
    const [loading, setLoading] = useState(false);
    const [message, setMessage] = useState({ text: '', type: '' });
    const [formData, setFormData] = useState({ id: '', name: '', email: '', telephone_number: '' });

    // Try to find the user record in the database
    const loadUser = useCallback(() => {
        const allEmployees = getCollection('HR_EMPLOYEE_CRUD') || [];
        const myRecord = allEmployees.find(emp => 
            emp.username === currentUser.username || emp.id === currentUser.id
        );

        if (myRecord) {
            setFormData({
                id: myRecord.id,
                name: myRecord.name || '',
                email: myRecord.email || '',
                telephone_number: myRecord.telephone_number || ''
            });
            return true;
        }
        return false;
    }, [currentUser, getCollection]);

    useEffect(() => { loadUser(); }, [loadUser]);

    // FALLBACK: Create a record if one doesn't exist
    const handleInitializeRecord = async () => {
        setLoading(true);
        const success = await addRecord('HR_EMPLOYEE_CRUD', {
            username: currentUser.username,
            name: currentUser.username,
            role: currentUser.role,
            status: 'active'
        });
        if (success) {
            loadUser();
            setMessage({ text: 'Database record initialized!', type: 'success' });
        }
        setLoading(false);
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setLoading(true);
        const success = await updateRecord('HR_EMPLOYEE_CRUD', formData.id, formData);
        setLoading(false);
        setMessage(success ? { text: 'Profile saved!', type: 'success' } : { text: 'Save failed.', type: 'error' });
    };

    if (!formData.id) {
        return (
            <div className="p-10 text-center space-y-4">
                <Icons.Database size={48} className="mx-auto text-yellow-500 opacity-50" />
                <h2 className="text-xl font-bold text-white">No Database Link Found</h2>
                <p className="text-gray-400">Your current account ({currentUser.username}) is not yet linked to an employee record.</p>
                <button onClick={handleInitializeRecord} className="bg-cyan-600 px-6 py-3 rounded-xl font-bold text-white hover:bg-cyan-500">
                    Initialize Database Record
                </button>
            </div>
        );
    }

    return (
        <ContainerCard Icon='Edit' glassClass={glassClass}>
            <form onSubmit={handleSubmit} className="p-6 space-y-4 max-w-xl mx-auto">
                {message.text && <div className={`p-3 rounded text-sm ${message.type === 'success' ? 'bg-green-500/10 text-green-400' : 'bg-red-500/10 text-red-400'}`}>{message.text}</div>}
                <Input label="Full Name" value={formData.name} onChange={(v) => setFormData({...formData, name: v})} />
                <Input label="Email" value={formData.email} onChange={(v) => setFormData({...formData, email: v})} />
                <Input label="Phone" value={formData.telephone_number} onChange={(v) => setFormData({...formData, telephone_number: v})} />
                <button type="submit" className="w-full bg-cyan-600 p-3 rounded font-bold text-white uppercase tracking-widest">{loading ? 'Saving...' : 'Save Profile'}</button>
            </form>
        </ContainerCard>
    );
};

const StaffSignup = ({ glassClass }) => {
    const { getCollection, addRecord } = useData();
    const [formData, setFormData] = useState({
        username: '',
        fullName: '',
        password: '',
        role: '' // Stores the ID from ROLE_CRUD
    });

    // Populate dropdown with Dynamic Roles from ROLE_CRUD
    const dynamicRoleOptions = useMemo(() => {
        const savedRoles = getCollection('ROLE_CRUD') || [];
        const base = [
            { label: 'Admin (System)', value: 'admin' },
            { label: 'Standard Employee', value: 'employee' }
        ];
        
        const custom = savedRoles.map(role => ({
            label: role.name || role.roleName,
            value: role.id 
        }));

        return [...base, ...custom];
    }, [getCollection]);

    const handleSubmit = async (e) => {
        e.preventDefault();
        try {
            // Find display name for system_role
            const roleObj = dynamicRoleOptions.find(o => o.value === formData.role);
            
            await addRecord('HR_EMPLOYEE_CRUD', {
                ...formData,
                name: formData.fullName,
                status: 'active',
                system_role: roleObj?.label || 'Employee',
                createdAt: new Date().toISOString()
            });
            
            alert(`Staff Registered Successfully: ${formData.username}`);
            setFormData({ username: '', fullName: '', password: '', role: '' });
        } catch (error) {
            alert("Signup failed.");
        }
    };

    return (
        <ContainerCard Icon="UserPlus" glassClass={glassClass}>
            <form onSubmit={handleSubmit} className="p-6 space-y-4 max-w-2xl mx-auto">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <Input label="Full Name" value={formData.fullName} onChange={(v) => setFormData({...formData, fullName: v})} isRequired />
                    <Input label="Username" value={formData.username} onChange={(v) => setFormData({...formData, username: v})} isRequired />
                </div>
                <Input label="Initial Password" type="password" placeholder="Enter password" value={formData.password} onChange={(v) => setFormData({...formData, password: v})} isRequired />
                
                <div className="flex flex-col space-y-2">
                    <label className="text-[10px] font-black text-cyan-400 uppercase tracking-widest">Assign Permission Role</label>
                    <select 
                        value={formData.role} 
                        onChange={(e) => setFormData({...formData, role: e.target.value})} 
                        className="w-full bg-black/60 border border-white/20 p-4 rounded-xl text-white outline-none focus:border-cyan-500 transition-all font-bold"
                    >
                        <option value="">-- Choose Access Level --</option>
                        {dynamicRoleOptions.map(opt => <option key={opt.value} value={opt.value} className="bg-slate-900">{opt.label}</option>)}
                    </select>
                </div>

                <button type="submit" className="w-full bg-cyan-600 hover:bg-cyan-500 p-4 rounded-xl font-black text-white shadow-lg mt-6 uppercase tracking-tighter transition-transform active:scale-95">
                    Generate Staff Account
                </button>
            </form>
        </ContainerCard>
    );
};
// Access Denied Fallback
const AccessDenied = ({ glassClass }) => (
    <ContainerCard  Icon='Lock' glassClass={glassClass} className="h-full flex items-center justify-center">
        <div className="text-center text-red-400 p-8">
            <Icons.AlertTriangle className="w-16 h-16 mx-auto mb-4" />
            <p className="text-2xl font-bold">You do not have permission to view this resource.</p>
            <p className="text-lg text-gray-400 mt-2">Please contact your administrator for assistance.</p>
        </div>
    </ContainerCard>
);

/* ----------------------------------------------------------------------
   10. MAIN APP & NAVIGATION (Standard)
   ---------------------------------------------------------------------- */
const RoleAssignmentManager = ({ glassClass }) => {
    const { getCollection, updateRecord } = useData();
    
    // Use getCollection directly to ensure we get the latest database state
    const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
    const roles = getCollection('ROLE_CRUD') || []; 
    
    const [selectedEmployeeId, setSelectedEmployeeId] = useState('');
    const [selectedRoleId, setSelectedRoleId] = useState('');

    const selectStyle = {
        backgroundColor: '#000000',
        color: '#ffffff',         
        appearance: 'none', 
        WebkitAppearance: 'none'
    };

    const handleAssignment = async () => {
        if (!selectedEmployeeId || !selectedRoleId) {
            return alert("Please select both an employee and a role.");
        }

        const employee = employees.find(e => e.id === selectedEmployeeId);
        const role = roles.find(r => r.id === selectedRoleId);

        if (!employee) {
            alert("Selected employee not found. Please try again.");
            return;
        }
        if (!role) {
            alert("Selected role not found. Please try again.");
            return;
        }

        try {
            await updateRecord('HR_EMPLOYEE_CRUD', selectedEmployeeId, {
                ...employee, 
                role: selectedRoleId, 
                system_role: role.roleName || role.name 
            });
            alert(`Role "${role.roleName || role.name}" assigned to ${employee.name}`);
            setSelectedEmployeeId('');
            setSelectedRoleId('');
        } catch (error) {
            console.error("Role assignment failed:", error);
            alert("Role assignment failed. Please check the console for errors.");
        }
    };

    return (
        <div className="space-y-6 pb-40"> 
            <ContainerCard glassClass={glassClass}>
                <div className="p-6 space-y-6">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        
                        {/* EMPLOYEE SELECT */}
                        <div className="relative">
                            <label className="text-cyan-400 text-[10px] font-bold mb-2 block uppercase tracking-widest">
                                1. Select Staff Member ({employees.length} users found)
                            </label>
                            <select 
                                style={selectStyle} 
                                className="w-full bg-black border border-white/20 rounded-xl p-4 text-white focus:border-cyan-500 outline-none cursor-pointer shadow-2xl"
                                value={selectedEmployeeId}
                                onChange={(e) => setSelectedEmployeeId(e.target.value)}
                            >
                                <option value="">-- Click to view list --</option> 
                                {employees.map(emp => (
                                    <option key={emp.id} value={emp.id}>
                                        {/* FIXED THE CHARACTER ENCODING BELOW */}
                                        {emp.name} - {emp.email || 'No Email'}
                                    </option>
                                ))}
                            </select>
                        </div>

                        {/* ROLE SELECT */}
                        <div className="relative">
                            <label className="text-cyan-400 text-[10px] font-bold mb-2 block uppercase tracking-widest">
                                2. Choose New Permissions Role
                            </label>
                            <select 
                                style={selectStyle} 
                                className="w-full bg-black border border-white/20 rounded-xl p-4 text-white focus:border-cyan-500 outline-none cursor-pointer shadow-2xl"
                                value={selectedRoleId}
                                onChange={(e) => setSelectedRoleId(e.target.value)}
                            >
                                <option value="">-- Click to view roles --</option> 
                                {roles.map(role => (
                                    <option key={role.id} value={role.id}>
                                        {role.roleName || role.name}
                                    </option>
                                ))}
                            </select>
                        </div>
                    </div>

                    <button 
                        onClick={handleAssignment}
                        className="w-full bg-cyan-600 hover:bg-cyan-500 py-4 rounded-xl font-black text-white transition-all shadow-lg flex items-center justify-center gap-2 uppercase tracking-tighter"
                    >
                        <Icons.UserCheck className="w-5 h-5" /> Confirm Role Assignment
                    </button>
                </div>
            </ContainerCard>

            <ContainerCard glassClass={glassClass}>
                <div className="p-4">
                    <table className="w-full text-left">
                        <thead>
                            <tr className="text-gray-500 text-[10px] uppercase border-b border-white/10">
                                <th className="p-3">Staff</th>
                                <th className="p-3">Current Role</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-white/5">
                            {employees.length === 0 ? (
                                <tr><td colSpan="2" className="p-4 text-center text-gray-500">No users in database.</td></tr>
                            ) : (
                                employees.map(emp => (
                                    <tr key={emp.id} className="hover:bg-white/5 transition-colors">
                                        <td className="p-3 text-white font-medium">{emp.name}</td>
                                        <td className="p-3">
                                            <span className="bg-cyan-500/10 text-cyan-400 border border-cyan-500/20 px-2 py-1 rounded text-xs">
                                                {emp.system_role || emp.role || 'N/A'}
                                            </span>
                                        </td>
                                    </tr>
                                ))
                            )}
                        </tbody>
                    </table>
                </div>
            </ContainerCard>
        </div>
    );
};
// 1. STANDALONE HELPER FUNCTION (Fixes the ReferenceError)
const findMenuItemTitle = (items, currentId) => {
    if (!items || !Array.isArray(items)) return null;
    for (const item of items) {
        if (item.id === currentId) return item.label || item.title;
        if (item.children) {
            const childTitle = findMenuItemTitle(item.children, currentId);
            if (childTitle) return childTitle;
        }
    }
    return null;
};

const renderView = (currentView, glassClass, getCollection, checkPermission, allRoles = []) => {
    if (!checkPermission(currentView, allRoles)) return <AccessDenied glassClass={glassClass} />;
    switch (currentView) {
        case 'MAT_DASHBOARD_VIEW': return <MatDashboard glassClass={glassClass} />;
        case 'UI_CUSTOMIZATION': return <UICustomization glassClass={glassClass} />;
        case 'HR_ROLE_ASSIGNMENT_CRUD':
        case 'ROLE_ASSIGNMENT_VIEW':
            return <RoleAssignmentManager glassClass={glassClass} />;
        case 'ROLE_CRUD': return <RoleManagement glassClass={glassClass} />;
        case 'MAT_GENERIC_REPORTS': return <ReportsGenerator glassClass={glassClass} />;
        case 'STAFF_SIGNUP':
        case 'STAFF_REGISTER':
        case 'REGISTER':
            return <StaffSignup glassClass={glassClass} />;
        case 'STAFF_CHANGE_PASSWORD': return <StaffChangePassword glassClass={glassClass} />;
        case 'STAFF_UPDATE_PROFILE': return <StaffUpdateProfile glassClass={glassClass} />;
        default:
            if (COLUMN_CONFIG[currentView]) return (
                <div className="w-full">
                    <DataTable
                        collectionName={currentView}
                        columns={COLUMN_CONFIG[currentView] || []}
                        data={typeof getCollection === 'function' ? getCollection(currentView) : []}
                        glassClass={glassClass}
                        getCollection={getCollection}
                    />
                </div>
            );
            return <AccessDenied glassClass={glassClass} />;
    }
};

// Navigation Component
// Navigation Component
const NavItem = ({ item, currentView, setCurrentView, checkPermission }) => {
    const IconComponent = getIcon(item.icon);
    const isSelected = currentView === item.id;
    const isAuthorized = checkPermission(item.id);

    if (!isAuthorized) return null;

    return (
        <li className="mb-1">
            <button
                onClick={() => setCurrentView(item.id)}
                className={`flex items-center w-full p-2 rounded-lg transition duration-200 text-sm font-medium ${isSelected
                    ? 'bg-cyan-600 text-white shadow-lg'
                    : 'text-gray-300 hover:bg-cyan-900/50'
                    }`}
            >
                <IconComponent className="w-5 h-5 mr-3" />
                {item.label}
            </button>
        </li>
    );
};

const Navigation = ({ menuItems, currentView, setCurrentView, expandedFolders, setExpandedFolders, checkPermission, glassClass, logoURL }) => {
    const toggleFolder = (id) => {
        setExpandedFolders(prev =>
            prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
        );
    };

    return (
        <nav className="h-full flex flex-col p-4 overflow-y-auto custom-scrollbar">
            {/* LOGO SECTION - PERSISTENT AT TOP */}
            <div className="flex items-center gap-3 mb-8 px-3 py-4 border-b border-white/10 shrink-0">
                {logoURL ? (
                    <img 
                        src={logoURL} 
                        alt="Logo" 
                        className="h-12 w-auto object-contain" 
                    />
                ) : (
                    <Icons.Activity className="w-8 h-8 text-cyan-400" />
                )}
                <span className="text-xl font-bold text-white tracking-tight">Unity Material Management</span>
            </div>

            <ul className="space-y-2 flex-grow">
                {menuItems.map(menu => {
                    const isMenuAuthorized = menu.children.some(child => checkPermission(child.id));
                    if (!isMenuAuthorized) return null;

                    const isExpanded = expandedFolders.includes(menu.id);
                    const MenuIcon = getIcon(menu.icon);
                    const hasActiveChild = menu.children.some(child => currentView === child.id);

                    return (
                        <li key={menu.id}>
                            <button
                                onClick={() => toggleFolder(menu.id)}
                                className={`flex items-center justify-between w-full p-3 font-semibold rounded-lg transition duration-200 text-lg ${hasActiveChild || isExpanded
                                    ? 'bg-cyan-800/70 text-white shadow-lg'
                                    : 'text-cyan-400 hover:bg-cyan-900/50'
                                    }`}
                            >
                                <div className="flex items-center">
                                    <MenuIcon className="w-6 h-6 mr-3" />
                                    {menu.label}
                                </div>
                                {isExpanded ? <Icons.ChevronUp className="w-5 h-5" /> : <Icons.ChevronDown className="w-5 h-5" />}
                            </button>

                            {isExpanded && (
                                <ul className="ml-3 mt-1 space-y-1 transition-all">
                                    {menu.children.map(item => {
                                        if (!checkPermission(item.id)) return null;
                                        const ItemIconComponent = getIcon(item.icon);
                                        const isSelected = currentView === item.id;

                                        return (
                                            <li key={item.id} className="mb-1">
                                                <button
                                                    onClick={() => setCurrentView(item.id)}
                                                    className={`flex items-center w-full p-2 rounded-lg transition duration-200 text-sm font-medium 
                                                        ${isSelected 
                                                            ? 'bg-cyan-600 text-white shadow-md' 
                                                            : 'text-gray-300 bg-black'
                                                        }
                                                    `}
                                                >
                                                    <ItemIconComponent className="w-5 h-5 mr-3" />
                                                    {item.label}
                                                </button>
                                            </li>
                                        );
                                    })}
                                </ul>
                            )}
                        </li>
                    );
                })}
            </ul>
        </nav>
    );
};


 

const HrmAppContent = () => {
    const { isLoggedIn, currentView, setCurrentView, currentUser, checkPermission, logout } = useAuth();
    const { glassOpacityClass, rootStyle, uiSettings } = useUI();
    const { getCollection } = useData();

    // FETCH LIVE ROLES: This is the heartbeat of the April 21 logic
    const allRoles = getCollection('ROLE_CRUD') || [];
    const [showNotifications, setShowNotifications] = useState(false);
    const [notifCount, setNotifCount] = useState(0);
    const [mobileNavOpen, setMobileNavOpen] = useState(false);
    const isMobile = useIsMobile(768); 
    const [expandedFolders, setExpandedFolders] = useState(['HR', 'SETTINGS']); 

    const glassClass = `bg-gray-900/70 backdrop-blur-md`;

    
    // If current view is not allowed, fall back to a dashboard the role can open
    useEffect(() => {
        if (!isLoggedIn || !currentUser) return;
        try {
            const roles = typeof getCollection === 'function' ? (getCollection('ROLE_CRUD') || []) : [];
            const ok = checkPermission(currentView, roles);
            if (!ok) {
                const candidates = ['MAT_DASHBOARD_VIEW', 'HR_DASHBOARD_VIEW', 'FIN_DASHBOARD', 'FIN_DASHBOARD_VIEW', 'HR_DASHBOARD', 'MAT_DASHBOARD'];
                const next = candidates.find((id) => checkPermission(id, roles)) || candidates[0];
                if (next && next !== currentView) setCurrentView(next);
            }
        } catch (e) {}
    }, [isLoggedIn, currentUser, currentView, checkPermission, getCollection]);

    if (!isLoggedIn) {
        return <AuthScreen glassClass={glassClass} />;
    }

    // Role / modules[] must allow this app
    if (currentUser && typeof canAccessModule === 'function' && !canAccessModule(currentUser.role, 'material', currentUser)) {
        return (
            <ModuleAccessGate moduleId="material" user={currentUser} onLogout={logout} />
        );
    }
    if (currentUser && typeof canAccessModule !== 'function') {
        const _r = String(currentUser.role || '').toLowerCase().replace(/\s+/g, '_');
        const _admin = _r === 'admin' || _r === 'super_admin';
        const _allow = { finance: ['finance_officer','finance_manager','accountant','payroll_officer','hr_manager'], hr: ['hr_manager','hr_staff','hr_officer','recruiter'], material: ['store_manager','inventory_clerk','procurement_officer','warehouse'] }['material'] || [];
        if (!_admin && !_allow.includes(_r) && !(Array.isArray(currentUser.modules) && currentUser.modules.includes('material'))) {
            return (
                <ModuleAccessGate moduleId="material" user={currentUser} onLogout={logout} />
            );
        }
    }


    return (
        <div className="h-full w-full flex flex-col overflow-hidden relative uu-module-root" style={{ ...(rootStyle || {}), position: "relative", inset: "auto", zIndex: "auto", pointerEvents: "auto" }}>
            <MobileStyles />
            <header className={`uu-app-header h-14 shrink-0 flex items-center justify-between px-3 sm:px-6 border-b border-white/10 z-20 ${glassClass}`}>
                <div className="flex items-center gap-2 sm:gap-4 min-w-0">
                    <button type="button" className="md:hidden p-2 rounded-lg bg-white/10 text-cyan-300 shrink-0" aria-label="Open menu" onClick={() => setMobileNavOpen(true)}>
                        <Icons.Menu className="w-5 h-5" />
                    </button>
                    {uiSettings?.logoURL && <img src={uiSettings.logoURL} alt="Logo" className="h-7 sm:h-8 w-auto shrink-0" />}
                    <h1 className="text-sm sm:text-xl font-black text-cyan-400 tracking-tighter italic truncate">UNITY UNIVERSITY</h1>
                </div>
                <div className="flex items-center gap-2 sm:gap-6 shrink-0">
                    <div className="uu-user-block text-right border-r border-white/10 pr-2 sm:pr-6 hidden sm:block">
                        <p className="text-[10px] sm:text-xs font-bold text-white truncate max-w-[7rem]">{currentUser?.username}</p>
                        <p className="text-[8px] sm:text-[9px] text-cyan-500 uppercase font-black truncate">{currentUser?.role}</p>
                    </div>
                    <button type="button" onClick={logout} className="p-2 bg-red-500/10 text-red-500 rounded-lg hover:bg-red-500/20 text-[10px] sm:text-xs font-bold transition-all">LOGOUT</button>
                    <button type="button" onClick={() => setShowNotifications(!showNotifications)} className={`relative p-2 rounded-full transition-all ${showNotifications ? 'bg-cyan-500 text-white' : 'text-gray-400 hover:text-white hover:bg-white/10'}`}>
                        <span className="relative inline-flex">
                            <Icons.Bell className="w-5 h-5" />
                            {notifCount > 0 && (
                                <span className="absolute -top-1.5 -right-1.5 min-w-[18px] h-[18px] px-1 rounded-full bg-rose-500 text-white text-[9px] font-black flex items-center justify-center shadow-lg shadow-rose-500/50 ring-2 ring-slate-950 animate-pulse" title={`${notifCount} notifications`}>
                                    {notifCount > 99 ? '99+' : notifCount}
                                </span>
                            )}
                        </span>
                    </button>
                </div>
            </header>

            <div className="flex flex-grow overflow-hidden min-h-0 relative">
                {mobileNavOpen && <div className="uu-sidebar-backdrop md:hidden" onClick={() => setMobileNavOpen(false)} />}
                <aside className={`uu-sidebar-desktop hidden md:block w-[20%] min-w-[200px] max-w-[280px] border-r border-white/10 ${glassClass}`}>
                    <Navigation
                        menuItems={MENU_ITEMS}
                        currentView={currentView}
                        setCurrentView={(id) => checkPermission(id, allRoles) && setCurrentView(id)}
                        expandedFolders={expandedFolders}
                        setExpandedFolders={setExpandedFolders}
                        // PASS ROLES DOWN TO SIDEBAR
                        checkPermission={(id) => checkPermission(id, allRoles)}
                        glassClass={glassClass}
                    />
                </aside>
                <aside className={`uu-sidebar-drawer md:hidden border-r border-white/10 ${glassClass} ${mobileNavOpen ? 'uu-open' : ''}`}>
                    <div className="flex items-center justify-between p-3 border-b border-white/10">
                        <span className="text-xs font-black text-cyan-400 uppercase">Menu</span>
                        <button type="button" className="p-2 text-gray-400" onClick={() => setMobileNavOpen(false)} aria-label="Close menu">
                            <Icons.X className="w-5 h-5" />
                        </button>
                    </div>
                    <Navigation
                        menuItems={MENU_ITEMS}
                        currentView={currentView}
                        setCurrentView={(id) => {
                            if (checkPermission(id, allRoles)) {
                                setCurrentView(id);
                                setMobileNavOpen(false);
                            }
                        }}
                        expandedFolders={expandedFolders}
                        setExpandedFolders={setExpandedFolders}
                        checkPermission={(id) => checkPermission(id, allRoles)}
                        glassClass={glassClass}
                    />
                </aside>


                <main className={`uu-main flex-1 min-w-0 min-h-0 overflow-y-auto overflow-x-hidden ${glassClass} p-2 sm:p-4 md:p-5 custom-scrollbar`}>
                    {/* PASS ROLES DOWN TO RENDERVIEW */}
                    {renderView(currentView, glassClass, getCollection, checkPermission, allRoles)}
                </main>

                <aside
                    className={`border-l border-white/10 ${glassClass} ${
                        showNotifications
                            ? 'w-[25%] min-w-[280px] max-w-[360px] animate-in slide-in-from-right'
                            : 'hidden w-0 min-w-0 overflow-hidden'
                    }`}
                    aria-hidden={!showNotifications}
                >
                    <NotificationBar
                        glassClass={glassClass}
                        currentUser={currentUser}
                        onClose={() => setShowNotifications(false)}
                        onCountChange={setNotifCount}
                    />
                </aside>
                    <UnityUniversityAIDock glassClass={glassClass} />

            </div>
        </div>
    );
};

// Provider Wrapper
const App = () => {
    return (
        <AuthProvider>
            <UIProvider>
                <DataProvider>
                    <HrmAppContent />
                </DataProvider>
            </UIProvider>
        </AuthProvider>
    )
};

// Main entry point is the App component


const MatDashboard = ({ glassClass }) => {
    const { getCollection } = useData();
    const [chartType, setChartType] = useState('bar');
    const [selectedMetric, setSelectedMetric] = useState('stock');
    const [locationFilter, setLocationFilter] = useState('ALL');
    const [searchTerm, setSearchTerm] = useState('');
    const chartRef = useRef(null);

    const items = getCollection('MAT_ITEM_CRUD') || [];
    const grv = getCollection('MAT_GRV_CRUD') || [];
    const giv = getCollection('MAT_GIV_CRUD') || [];
    const robbing = getCollection('MAT_ROBBING_CRUD') || [];
    const localPO = getCollection('MAT_PURCHASE_LOCAL_CRUD') || [];
    const foreignPO = getCollection('MAT_PURCHASE_FOREIGN_CRUD') || [];

    const COLORS = ['#06b6d4', '#8b5cf6', '#f59e0b', '#10b981', '#ef4444', '#3b82f6', '#ec4899', '#14b8a6'];
    const fmt = (n) => Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 0 });
    const fmt2 = (n) => Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const q = (searchTerm || '').trim().toLowerCase();

    const analytics = useMemo(() => {
        const safeItems = Array.isArray(items) ? items : [];
        const safeGrv = Array.isArray(grv) ? grv : [];
        const safeGiv = Array.isArray(giv) ? giv : [];
        const safeRob = Array.isArray(robbing) ? robbing : [];
        const safeLocal = Array.isArray(localPO) ? localPO : [];
        const safeForeign = Array.isArray(foreignPO) ? foreignPO : [];

        const matchQ = (obj, keys) => {
            if (!q) return true;
            return keys.some(k => String(obj[k] ?? '').toLowerCase().includes(q));
        };

        let filteredItems = locationFilter === 'ALL'
            ? safeItems
            : safeItems.filter(i => String(i.location || '').toLowerCase().includes(locationFilter.toLowerCase()));
        filteredItems = filteredItems.filter(i => matchQ(i, ['name', 'location', 'unit', 'id']));

        const fGrv = safeGrv.filter(g => matchQ(g, ['item_name', 'supplier', 'po_number', 'status']));
        const fGiv = safeGiv.filter(g => matchQ(g, ['item_name', 'department', 'purpose']));
        const fRob = safeRob.filter(r => matchQ(r, ['donor_item', 'part_taken', 'recipient_item', 'performed_by', 'status']));
        const fLocal = safeLocal.filter(p => matchQ(p, ['supplier', 'item_name', 'status']));
        const fForeign = safeForeign.filter(p => matchQ(p, ['supplier', 'item_name']));

        const totalQty = filteredItems.reduce((s, i) => s + (Number(i.quantity) || 0), 0);
        const totalValue = filteredItems.reduce((s, i) => s + ((Number(i.quantity) || 0) * (Number(i.unit_cost) || 0)), 0);
        const lowStock = filteredItems.filter(i => (Number(i.quantity) || 0) > 0 && (Number(i.quantity) || 0) <= 15);
        const outOfStock = filteredItems.filter(i => (Number(i.quantity) || 0) <= 0);
        const locations = [...new Set(safeItems.map(i => i.location).filter(Boolean))];

        const grvPending = fGrv.filter(g => {
            const st = String(g.status || '').toLowerCase();
            return st.includes('pending') || st.includes('partial') || st.includes('inspection');
        });
        const givQty = fGiv.reduce((s, g) => s + (Number(g.quantity) || 0), 0);
        const grvQty = fGrv.reduce((s, g) => s + (Number(g.quantity) || 0), 0);
        const purchaseValue = fLocal.reduce((s, p) => s + (Number(p.total_value) || 0), 0)
            + fForeign.reduce((s, p) => s + (Number(p.customs_duty) || 0) + (Number(p.shipping_cost) || 0), 0);

        const byLocation = {};
        filteredItems.forEach(i => {
            const loc = i.location || 'Unknown';
            if (!byLocation[loc]) byLocation[loc] = { name: loc, qty: 0, value: 0, skus: 0 };
            byLocation[loc].qty += Number(i.quantity) || 0;
            byLocation[loc].value += (Number(i.quantity) || 0) * (Number(i.unit_cost) || 0);
            byLocation[loc].skus += 1;
        });

        const topStock = [...filteredItems]
            .sort((a, b) => (Number(b.quantity) || 0) - (Number(a.quantity) || 0))
            .slice(0, 12)
            .map(i => ({
                name: String(i.name || 'Item').slice(0, 24),
                value: Number(i.quantity) || 0,
                cost: Number(i.unit_cost) || 0,
                location: i.location || '—',
                unit: i.unit || 'Pcs',
                id: i.id
            }));

        const grvByItem = {};
        fGrv.forEach(g => {
            const k = g.item_name || 'Unknown';
            grvByItem[k] = (grvByItem[k] || 0) + (Number(g.quantity) || 0);
        });
        const topGrv = Object.entries(grvByItem).sort((a, b) => b[1] - a[1]).slice(0, 12)
            .map(([name, value]) => ({ name: name.slice(0, 24), value }));

        const givByDept = {};
        fGiv.forEach(g => {
            const k = g.department || 'Unknown';
            givByDept[k] = (givByDept[k] || 0) + (Number(g.quantity) || 0);
        });
        const topGiv = Object.entries(givByDept).sort((a, b) => b[1] - a[1]).slice(0, 12)
            .map(([name, value]) => ({ name: name.slice(0, 24), value }));

        const robByDonor = {};
        fRob.forEach(r => {
            const k = r.donor_item || r.part_taken || 'Unknown';
            robByDonor[k] = (robByDonor[k] || 0) + 1;
        });
        const topRob = Object.entries(robByDonor).sort((a, b) => b[1] - a[1]).slice(0, 12)
            .map(([name, value]) => ({ name: String(name).slice(0, 24), value }));

        const locationChart = Object.values(byLocation).sort((a, b) => b.value - a.value).slice(0, 10)
            .map(l => ({ name: String(l.name).slice(0, 20), value: Math.round(l.value), qty: l.qty, skus: l.skus }));

        const lowChart = [...lowStock, ...outOfStock]
            .sort((a, b) => (Number(a.quantity) || 0) - (Number(b.quantity) || 0))
            .slice(0, 12)
            .map(i => ({
                name: String(i.name || 'Item').slice(0, 24),
                value: Number(i.quantity) || 0,
                location: i.location || '—',
                cost: Number(i.unit_cost) || 0
            }));

        const oosChart = outOfStock.slice(0, 12).map(i => ({
            name: String(i.name || 'Item').slice(0, 24),
            value: 0,
            location: i.location || '—',
            cost: Number(i.unit_cost) || 0
        }));

        const poBySupplier = {};
        [...fLocal, ...fForeign].forEach(p => {
            const k = p.supplier || 'Unknown';
            poBySupplier[k] = (poBySupplier[k] || 0) + (Number(p.total_value) || Number(p.customs_duty) || 0) + (Number(p.shipping_cost) || 0);
        });
        const topPO = Object.entries(poBySupplier).sort((a, b) => b[1] - a[1]).slice(0, 12)
            .map(([name, value]) => ({ name: name.slice(0, 24), value: Math.round(value) }));

        const valueChart = [...filteredItems]
            .map(i => ({
                name: String(i.name || 'Item').slice(0, 24),
                value: Math.round((Number(i.quantity) || 0) * (Number(i.unit_cost) || 0)),
                qty: Number(i.quantity) || 0,
                location: i.location || '—'
            }))
            .sort((a, b) => b.value - a.value)
            .slice(0, 12);

        return {
            filteredItems, totalQty, totalValue, lowStock, outOfStock, locations,
            grvPending, givQty, grvQty, purchaseValue,
            topStock, topGrv, topGiv, topRob, locationChart, lowChart, oosChart, topPO, valueChart,
            fGrv, fGiv, fRob, fLocal, fForeign,
            skuCount: filteredItems.length,
            grvCount: fGrv.length,
            givCount: fGiv.length,
            robCount: fRob.length,
            poCount: fLocal.length + fForeign.length,
        };
    }, [items, grv, giv, robbing, localPO, foreignPO, locationFilter, q]);

    /* Metric registry — every card maps here */
    const METRICS = {
        stock: {
            label: 'Top Stock by Quantity',
            tip: 'Highest on-hand quantities in filtered stores',
            data: analytics.topStock,
            columns: [
                { key: 'name', label: 'Item' },
                { key: 'value', label: 'Qty' },
                { key: 'cost', label: 'Unit Cost' },
                { key: 'location', label: 'Location' },
            ],
        },
        grv: {
            label: 'Goods Received by Item',
            tip: 'Inbound quantities aggregated by item name',
            data: analytics.topGrv,
            columns: [
                { key: 'name', label: 'Item' },
                { key: 'value', label: 'Qty Received' },
            ],
        },
        giv: {
            label: 'Issues by Department',
            tip: 'Outbound quantities issued to consuming departments',
            data: analytics.topGiv,
            columns: [
                { key: 'name', label: 'Department' },
                { key: 'value', label: 'Qty Issued' },
            ],
        },
        maintenance: {
            label: 'Parts Robbing (Cannibalization)',
            tip: 'How often donor assets / parts were reallocated',
            data: analytics.topRob,
            columns: [
                { key: 'name', label: 'Donor / Part' },
                { key: 'value', label: 'Events' },
            ],
        },
        location: {
            label: 'Inventory Value by Location',
            tip: 'Stock carrying value (qty × unit cost) by store',
            data: analytics.locationChart,
            columns: [
                { key: 'name', label: 'Location' },
                { key: 'value', label: 'Value (ETB)' },
                { key: 'qty', label: 'Units' },
                { key: 'skus', label: 'SKUs' },
            ],
        },
        low: {
            label: 'Low & Critical Stock',
            tip: 'Items at or below 15 units (including zero)',
            data: analytics.lowChart,
            columns: [
                { key: 'name', label: 'Item' },
                { key: 'value', label: 'Qty' },
                { key: 'location', label: 'Location' },
                { key: 'cost', label: 'Unit Cost' },
            ],
        },
        oos: {
            label: 'Out of Stock',
            tip: 'SKUs with zero quantity on hand',
            data: analytics.oosChart.length ? analytics.oosChart : [{ name: 'None', value: 0 }],
            columns: [
                { key: 'name', label: 'Item' },
                { key: 'value', label: 'Qty' },
                { key: 'location', label: 'Location' },
            ],
        },
        purchase: {
            label: 'Purchase Value by Supplier',
            tip: 'Local + foreign purchase exposure by supplier',
            data: analytics.topPO,
            columns: [
                { key: 'name', label: 'Supplier' },
                { key: 'value', label: 'Value (ETB)' },
            ],
        },
        value: {
            label: 'Highest Inventory Value (SKU)',
            tip: 'Items ranked by qty × unit cost',
            data: analytics.valueChart,
            columns: [
                { key: 'name', label: 'Item' },
                { key: 'value', label: 'Value (ETB)' },
                { key: 'qty', label: 'Qty' },
                { key: 'location', label: 'Location' },
            ],
        },
    };

    const active = METRICS[selectedMetric] || METRICS.stock;
    const activeData = active.data || [];

    const detailRows = useMemo(() => {
        switch (selectedMetric) {
            case 'stock':
            case 'value':
            case 'low':
            case 'oos':
            case 'location':
                return (selectedMetric === 'location'
                    ? analytics.locationChart
                    : selectedMetric === 'low'
                        ? [...analytics.outOfStock, ...analytics.lowStock]
                        : selectedMetric === 'oos'
                            ? analytics.outOfStock
                            : analytics.filteredItems
                ).slice(0, 14);
            case 'grv':
                return analytics.fGrv.slice(0, 14);
            case 'giv':
                return analytics.fGiv.slice(0, 14);
            case 'maintenance':
                return analytics.fRob.slice(0, 14);
            case 'purchase':
                return [...analytics.fLocal, ...analytics.fForeign].slice(0, 14);
            default:
                return analytics.filteredItems.slice(0, 14);
        }
    }, [selectedMetric, analytics]);

    const cardClass = (id) =>
        `cursor-pointer transition-all ${selectedMetric === id ? 'ring-2 ring-cyan-400/70 scale-[1.02]' : 'opacity-80 hover:opacity-100 hover:ring-1 hover:ring-white/20'}`;

    const locationOptions = ['ALL', ...(analytics.locations || [])];

    return (
        <div className="space-y-6 animate-in fade-in duration-500 pb-16 min-h-0 w-full">
            {/* Command header — finance formality */}
            <div className={`p-6 rounded-[2rem] ${glassClass} border border-cyan-500/25 bg-gradient-to-br from-cyan-950/40 to-black/40`}>
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                    <div>
                        <h2 className="text-2xl font-black text-white italic uppercase tracking-tight flex items-center gap-3">
                            <span className="p-2 rounded-2xl bg-cyan-500/20 border border-cyan-400/30">
                                <Icons.LayoutDashboard className="w-6 h-6 text-cyan-400" />
                            </span>
                            Material Command Center
                        </h2>
                        <p className="text-[10px] text-cyan-300/80 font-black uppercase tracking-[0.2em] mt-1">
                            Unity University · Inventory · GRV · GIV · Robbing · Procurement
                        </p>
                    </div>
                    <div className="flex flex-wrap items-end gap-3">
                        <div>
                            <label className="text-[10px] font-black uppercase tracking-widest mb-1 text-cyan-400 block">Search Master Stores</label>
                            <div className="relative">
                                <Icons.Search className="w-4 h-4 text-gray-500 absolute left-3 top-1/2 -translate-y-1/2" />
                                <input
                                    type="text"
                                    placeholder="Item, location, supplier, dept…"
                                    value={searchTerm}
                                    onChange={e => setSearchTerm(e.target.value)}
                                    className="pl-9 pr-3 py-2.5 rounded-xl bg-black/50 border border-cyan-500/30 text-xs text-white w-56 outline-none focus:border-cyan-400"
                                />
                            </div>
                        </div>
                        <div>
                            <label className="text-[10px] font-black uppercase tracking-widest mb-1 text-cyan-400 block">Location</label>
                            <UUSelect
                                value={locationFilter}
                                onChange={setLocationFilter}
                                options={locationOptions && locationOptions.length ? locationOptions : ['ALL']}
                                allLabel="All locations"
                                compact
                            />
                        </div>
                    </div>
                </div>
            </div>

            {/* Primary KPI row — click switches metric + chart + detail */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div onClick={() => setSelectedMetric('stock')} className={cardClass('stock')}>
                    <KPI title="SKUs / Assets" value={fmt(analytics.skuCount)} iconName="Package" colorClass="text-cyan-400" glassClass={glassClass}
                        explanation={`${fmt(analytics.totalQty)} units · ETB ${fmt2(analytics.totalValue)}`} />
                </div>
                <div onClick={() => setSelectedMetric('grv')} className={cardClass('grv')}>
                    <KPI title="Goods Received" value={fmt(analytics.grvCount)} iconName="Truck" colorClass="text-violet-400" glassClass={glassClass}
                        explanation={`${fmt(analytics.grvQty)} units in · ${analytics.grvPending.length} pending`} />
                </div>
                <div onClick={() => setSelectedMetric('giv')} className={cardClass('giv')}>
                    <KPI title="Goods Issued" value={fmt(analytics.givCount)} iconName="Store" colorClass="text-emerald-400" glassClass={glassClass}
                        explanation={`${fmt(analytics.givQty)} units to departments`} />
                </div>
                <div onClick={() => setSelectedMetric('maintenance')} className={cardClass('maintenance')}>
                    <KPI title="Parts Robbing" value={fmt(analytics.robCount)} iconName="Hammer" colorClass="text-amber-400" glassClass={glassClass}
                        explanation="Cannibalization: parts from donor assets restore recipients" />
                </div>
            </div>

            {/* Secondary cards — also drive the same graph + info */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <div onClick={() => setSelectedMetric('low')} className={cardClass('low')}>
                    <StatCard label="Low stock (≤15)" value={fmt(analytics.lowStock.length)} Icon={Icons.AlertTriangle}
                        color={analytics.lowStock.length ? 'text-amber-400' : 'text-emerald-400'} glassClass={glassClass} />
                </div>
                <div onClick={() => setSelectedMetric('oos')} className={cardClass('oos')}>
                    <StatCard label="Out of stock" value={fmt(analytics.outOfStock.length)} Icon={Icons.XCircle}
                        color={analytics.outOfStock.length ? 'text-rose-400' : 'text-emerald-400'} glassClass={glassClass} />
                </div>
                <div onClick={() => setSelectedMetric('purchase')} className={cardClass('purchase')}>
                    <StatCard label="Purchase docs" value={fmt(analytics.poCount)} Icon={Icons.ShoppingCart}
                        color="text-sky-400" glassClass={glassClass} />
                </div>
                <div onClick={() => setSelectedMetric('value')} className={cardClass('value')}>
                    <StatCard label="Inventory value (ETB)" value={fmt2(analytics.totalValue)} Icon={Icons.Banknote}
                        color="text-cyan-300" glassClass={glassClass} />
                </div>
            </div>

            {/* Quick metric chips including location */}
            <div className="flex flex-wrap gap-2">
                {Object.keys(METRICS).map(m => (
                    <button
                        key={m}
                        type="button"
                        onClick={() => setSelectedMetric(m)}
                        className={`px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-wider ${
                            selectedMetric === m ? 'bg-cyan-600 text-white shadow-lg shadow-cyan-900/40' : 'bg-white/5 text-gray-400 hover:bg-white/10'
                        }`}
                    >
                        {m}
                    </button>
                ))}
            </div>

            {/* Chart panel */}
            <div className={`rounded-[1.75rem] border border-white/10 ${glassClass} p-5 space-y-4 bg-black/40`}>
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                        <h3 className="text-sm font-black uppercase tracking-widest text-white">{active.label}</h3>
                        <p className="text-[10px] text-gray-500 mt-0.5">{active.tip} · {activeData.length} series points</p>
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                        {['bar', 'line', 'pie'].map(t => (
                            <button
                                key={t}
                                type="button"
                                onClick={() => setChartType(t)}
                                className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase ${
                                    chartType === t ? 'bg-cyan-600 text-white' : 'bg-white/5 text-gray-400 hover:bg-white/10'
                                }`}
                            >
                                {t}
                            </button>
                        ))}
                        <ExportButton
                            data={(activeData || []).map(r => ({
                                name: r.name,
                                value: r.value,
                                cost: r.cost,
                                location: r.location,
                                qty: r.qty,
                                skus: r.skus,
                                unit: r.unit,
                            }))}
                            columns={active.columns}
                            filename={`Material_${selectedMetric}_${chartType}_dashboard`}
                            chartRef={chartRef}
                        />
                    </div>
                </div>

                <div ref={chartRef} className="h-[360px] w-full bg-black/30 rounded-2xl p-3 border border-white/5">
                    {!activeData.length ? (
                        <div className="h-full flex items-center justify-center text-gray-500 text-sm">
                            No data for this metric. Adjust search / location or add store transactions.
                        </div>
                    ) : (
                        <ResponsiveContainer width="100%" height="100%">
                            {chartType === 'pie' ? (
                                <PieChart>
                                    <Pie data={activeData} cx="50%" cy="50%" innerRadius={50} outerRadius={110} paddingAngle={2}
                                        dataKey="value" nameKey="name"
                                        label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>
                                        {activeData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                                    </Pie>
                                    <Legend />
                                    <Tooltip contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #334155', borderRadius: 12, fontSize: 11 }} formatter={(v) => fmt(v)} />
                                </PieChart>
                            ) : chartType === 'line' ? (
                                <LineChart data={activeData} margin={{ top: 10, right: 10, left: 0, bottom: 50 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="#ffffff08" vertical={false} />
                                    <XAxis dataKey="name" stroke="#6b7280" fontSize={10} angle={-25} textAnchor="end" interval={0} height={60} />
                                    <YAxis stroke="#6b7280" fontSize={10} tickFormatter={fmt} />
                                    <Tooltip contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #334155', borderRadius: 12, fontSize: 11 }} formatter={(v) => fmt(v)} />
                                    <Line type="monotone" dataKey="value" stroke="#06b6d4" strokeWidth={3} dot={{ r: 5, fill: '#06b6d4', strokeWidth: 0 }} activeDot={{ r: 7 }} />
                                </LineChart>
                            ) : (
                                <BarChart data={activeData} margin={{ top: 10, right: 10, left: 0, bottom: 50 }}>
                                    <CartesianGrid strokeDasharray="3 3" stroke="#ffffff08" vertical={false} />
                                    <XAxis dataKey="name" stroke="#6b7280" fontSize={10} angle={-25} textAnchor="end" interval={0} height={60} />
                                    <YAxis stroke="#6b7280" fontSize={10} tickFormatter={fmt} />
                                    <Tooltip contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #334155', borderRadius: 12, fontSize: 11 }} formatter={(v) => fmt(v)} />
                                    <Bar dataKey="value" radius={[6, 6, 0, 0]} barSize={36}>
                                        {activeData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                                    </Bar>
                                </BarChart>
                            )}
                        </ResponsiveContainer>
                    )}
                </div>
            </div>

            {/* Live detail grid — changes with selected card */}
            <div className={`rounded-[1.75rem] border border-white/10 ${glassClass} p-5 bg-black/40`}>
                <div className="flex items-center justify-between mb-4 gap-3 flex-wrap">
                    <div>
                        <h3 className="text-xs font-black uppercase tracking-widest text-white">Live register · {active.label}</h3>
                        <p className="text-[10px] text-gray-500 mt-0.5">Rows follow the active card · search & location applied</p>
                    </div>
                    <span className="text-[10px] font-black uppercase text-cyan-400">{detailRows.length} rows</span>
                </div>
                <div className="uu-table-scroll overflow-x-auto max-h-80 custom-scrollbar">
                    <table className="w-full text-left text-[11px]">
                        <thead className="sticky top-0 bg-slate-950/90 text-gray-400 uppercase tracking-widest text-[9px]">
                            <tr className="border-b border-white/10">
                                {selectedMetric === 'grv' && (<><th className="p-2">Date</th><th className="p-2">Item</th><th className="p-2">Supplier</th><th className="p-2 text-right">Qty</th><th className="p-2">Status</th></>)}
                                {selectedMetric === 'giv' && (<><th className="p-2">Date</th><th className="p-2">Item</th><th className="p-2">Department</th><th className="p-2 text-right">Qty</th><th className="p-2">Purpose</th></>)}
                                {selectedMetric === 'maintenance' && (<><th className="p-2">Date</th><th className="p-2">Donor</th><th className="p-2">Part</th><th className="p-2">Recipient</th><th className="p-2">By</th></>)}
                                {selectedMetric === 'purchase' && (<><th className="p-2">Date</th><th className="p-2">Supplier</th><th className="p-2">Item</th><th className="p-2 text-right">Value</th><th className="p-2">Status</th></>)}
                                {selectedMetric === 'location' && (<><th className="p-2">Location</th><th className="p-2 text-right">SKUs</th><th className="p-2 text-right">Units</th><th className="p-2 text-right">Value (ETB)</th></>)}
                                {['stock', 'value', 'low', 'oos'].includes(selectedMetric) && (
                                    <><th className="p-2">Item</th><th className="p-2">Location</th><th className="p-2 text-right">Qty</th><th className="p-2 text-right">Unit Cost</th><th className="p-2 text-right">Value</th></>
                                )}
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-white/5 text-gray-200">
                            {!detailRows.length ? (
                                <tr><td colSpan={6} className="p-6 text-center text-gray-500">No rows for this view.</td></tr>
                            ) : detailRows.map((r, idx) => {
                                if (selectedMetric === 'grv') {
                                    return (
                                        <tr key={r.id || idx} className="hover:bg-white/5">
                                            <td className="p-2 text-gray-400">{r.date || '—'}</td>
                                            <td className="p-2 font-semibold">{r.item_name || '—'}</td>
                                            <td className="p-2">{r.supplier || '—'}</td>
                                            <td className="p-2 text-right font-mono">{fmt(r.quantity)}</td>
                                            <td className="p-2"><span className="text-cyan-400 text-[10px] font-bold">{r.status || '—'}</span></td>
                                        </tr>
                                    );
                                }
                                if (selectedMetric === 'giv') {
                                    return (
                                        <tr key={r.id || idx} className="hover:bg-white/5">
                                            <td className="p-2 text-gray-400">{r.date || '—'}</td>
                                            <td className="p-2 font-semibold">{r.item_name || '—'}</td>
                                            <td className="p-2">{r.department || '—'}</td>
                                            <td className="p-2 text-right font-mono">{fmt(r.quantity)}</td>
                                            <td className="p-2 text-gray-400">{r.purpose || '—'}</td>
                                        </tr>
                                    );
                                }
                                if (selectedMetric === 'maintenance') {
                                    return (
                                        <tr key={r.id || idx} className="hover:bg-white/5">
                                            <td className="p-2 text-gray-400">{r.date || '—'}</td>
                                            <td className="p-2 font-semibold">{r.donor_item || '—'}</td>
                                            <td className="p-2">{r.part_taken || '—'}</td>
                                            <td className="p-2">{r.recipient_item || '—'}</td>
                                            <td className="p-2 text-gray-400">{r.performed_by || '—'}</td>
                                        </tr>
                                    );
                                }
                                if (selectedMetric === 'purchase') {
                                    return (
                                        <tr key={r.id || idx} className="hover:bg-white/5">
                                            <td className="p-2 text-gray-400">{r.date || '—'}</td>
                                            <td className="p-2 font-semibold">{r.supplier || '—'}</td>
                                            <td className="p-2">{r.item_name || '—'}</td>
                                            <td className="p-2 text-right font-mono text-cyan-300">{fmt2(r.total_value || r.customs_duty || 0)}</td>
                                            <td className="p-2"><span className="text-[10px] font-bold text-sky-400">{r.status || 'Foreign'}</span></td>
                                        </tr>
                                    );
                                }
                                if (selectedMetric === 'location') {
                                    return (
                                        <tr key={idx} className="hover:bg-white/5">
                                            <td className="p-2 font-semibold">{r.name}</td>
                                            <td className="p-2 text-right font-mono">{fmt(r.skus)}</td>
                                            <td className="p-2 text-right font-mono">{fmt(r.qty)}</td>
                                            <td className="p-2 text-right font-mono text-cyan-300">{fmt2(r.value)}</td>
                                        </tr>
                                    );
                                }
                                const qty = Number(r.quantity ?? r.value) || 0;
                                const cost = Number(r.unit_cost ?? r.cost) || 0;
                                return (
                                    <tr key={r.id || idx} className="hover:bg-white/5">
                                        <td className="p-2 font-semibold">{r.name || '—'}</td>
                                        <td className="p-2 text-gray-400">{r.location || '—'}</td>
                                        <td className={`p-2 text-right font-mono ${qty <= 0 ? 'text-rose-400' : qty <= 15 ? 'text-amber-400' : ''}`}>{fmt(qty)}</td>
                                        <td className="p-2 text-right font-mono">{fmt2(cost)}</td>
                                        <td className="p-2 text-right font-mono text-cyan-300">{fmt2(qty * cost)}</td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export function MaterialApp() {
    return (
        <AuthProvider>
            <UIProvider>
                <DataProvider>
                    <div className="h-full w-full min-h-0 relative overflow-hidden" style={{ pointerEvents: 'auto' }}>
                        <HrmAppContent />
                    </div>
                </DataProvider>
            </UIProvider>
        </AuthProvider>
    );
}
export default MaterialApp;