
import ModuleAccessGate from '../../unity/ModuleAccessGate.jsx';
import { attemptSharedLogin, upsertSharedUsers, getSharedUsers, resolveAssignedRole, changeSharedPassword, upsertSharedUser, saveSharedRoles, saveSharedSession, loadSharedSession, clearSharedSession, canAccessModule, getSharedRoles, mergeSharedRolesInto } from '../../unity/SharedCredentials.js';
import { postMaterialToFinance as bridgePostMaterialToFinance, afterMaterialBootstrap, getMaterialCoaOptions, MAT_SOURCE_COLLECTION_MAP } from '../../unity/UnityModuleBridge.js';
import LoginModuleSwitcher from '../../unity/LoginModuleSwitcher.jsx';
import { MobileStyles, useIsMobile } from '../../unity/mobileCompat.js';
import { IntelTip, IntelLabel, FieldIntelligenceSettings, FieldIntelligenceRuntime } from '../../unity/FieldIntelligence.jsx';
               
import React, { createContext, useContext, useState, useEffect, useCallback, useMemo, useRef } from 'react';
import ReactDOM, { createPortal } from 'react-dom';
import * as Icons from 'lucide-react';

import {
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
    'MAT_GENERIC_REPORTS': ['admin', 'store_manager', 'inventory_clerk', 'procurement_officer', 'hr_manager'],
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
            rootStyle, updateUISettings, resetToDefault,
            uiSettings: { backgroundImage, logoURL, primaryColor: backgroundColor, textColor, bgBlur, bgScale, bgBrightness } 
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
        try {
            if (!currentUser || !currentUser.username) {
                return { success: false, message: 'Not signed in.' };
            }
            const uName = String(currentUser.username).trim();
            const oldP = String(oldPassword ?? '');
            const newP = String(newPassword ?? '').trim();
            if (!oldP) return { success: false, message: 'Current password is required.' };
            if (!newP || newP.length < 4) {
                return { success: false, message: 'New password must be at least 4 characters.' };
            }
            if (oldP === newP) {
                return { success: false, message: 'New password must be different from the current password.' };
            }

            // Prefer shared helper; fall back to local implementation
            let applyShared = null;
            try {
                if (typeof changeSharedPassword === 'function') applyShared = changeSharedPassword;
            } catch (e) {}
            if (!applyShared) {
                applyShared = (username, o, n) => {
                    try {
                        const users = (typeof getSharedUsers === 'function' ? getSharedUsers() : []) || [];
                        const found = users.find(
                            (u) => String(u.username || '').trim().toLowerCase() === String(username).toLowerCase()
                        );
                        if (!found) return { success: false, message: 'User account not found.' };
                        const stored = String(found.password ?? '');
                        if (stored !== String(o ?? '') && stored.trim() !== String(o ?? '').trim()) {
                            return { success: false, message: 'Current password incorrect.' };
                        }
                        if (typeof upsertSharedUser === 'function') {
                            upsertSharedUser({
                                username: found.username,
                                password: String(n).trim(),
                                name: found.name,
                                role: found.role,
                                system_role: found.system_role || found.role,
                                role_id: found.role_id,
                                role_name: found.role_name,
                                modules: found.modules,
                                id: found.id,
                                _forceRole: true,
                            });
                        } else {
                            // raw localStorage write
                            const key = 'uu_shared_users_v1';
                            const raw = JSON.parse(localStorage.getItem(key) || '[]');
                            const list = Array.isArray(raw) ? raw : [];
                            const idx = list.findIndex(
                                (u) => String(u.username || '').toLowerCase() === String(username).toLowerCase()
                            );
                            const row = {
                                ...(idx >= 0 ? list[idx] : {}),
                                username: found.username || username,
                                password: String(n).trim(),
                                role: found.role,
                                name: found.name,
                            };
                            if (idx >= 0) list[idx] = { ...list[idx], ...row };
                            else list.push(row);
                            localStorage.setItem(key, JSON.stringify(list));
                        }
                        return { success: true, message: 'Password updated successfully. Please sign in again.' };
                    } catch (err) {
                        console.error(err);
                        return { success: false, message: 'Password change failed.' };
                    }
                };
            }

            let result = applyShared(uName, oldP, newP);

            // If not in shared store, try matching session-only / HR employee password
            if (!result.success && result.message === 'User account not found.') {
                // Seed shared user from current session + provided current password
                try {
                    if (typeof upsertSharedUser === 'function') {
                        // Only accept if they can prove old password against HR or we allow set
                        let hrOk = false;
                        if (getCollection) {
                            const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
                            const userRecord = employees.find(
                                (u) =>
                                    u &&
                                    String(u.username || u.user_name || '').toLowerCase() === uName.toLowerCase()
                            );
                            if (userRecord && String(userRecord.password ?? '') === oldP) hrOk = true;
                        }
                        // Also accept if they just logged in with this password (session active)
                        // User is logged in — trust oldP only when HR matches OR shared gets created with verify
                        if (hrOk) {
                            upsertSharedUser({
                                username: uName,
                                password: newP,
                                name: currentUser.name || uName,
                                role: currentUser.role,
                                role_id: currentUser.role_id,
                                role_name: currentUser.role_name,
                                _forceRole: true,
                            });
                            result = { success: true, message: 'Password updated successfully. Please sign in again.' };
                        }
                    }
                } catch (e) {}
            }

            // HR employee password sync
            try {
                if (result.success && getCollection && updateRecord) {
                    const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
                    const userRecord = employees.find(
                        (u) =>
                            u &&
                            String(u.username || u.user_name || '').toLowerCase() === uName.toLowerCase()
                    );
                    if (userRecord && userRecord.id != null) {
                        await updateRecord('HR_EMPLOYEE_CRUD', userRecord.id, {
                            ...userRecord,
                            password: newP,
                        });
                    }
                }
            } catch (e) {
                console.warn('[changePassword] HR sync', e);
            }

            // HR-only fallback when shared said incorrect / not found
            if (!result.success && getCollection && updateRecord) {
                const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
                const userRecord = employees.find(
                    (u) =>
                        u &&
                        String(u.username || u.user_name || '').toLowerCase() === uName.toLowerCase()
                );
                if (userRecord && String(userRecord.password ?? '') === oldP) {
                    const ok = await updateRecord('HR_EMPLOYEE_CRUD', userRecord.id, {
                        ...userRecord,
                        password: newP,
                    });
                    if (ok) {
                        try {
                            if (typeof upsertSharedUser === 'function') {
                                upsertSharedUser({
                                    username: uName,
                                    password: newP,
                                    name: userRecord.name || userRecord.employee_name || uName,
                                    role: currentUser.role,
                                    _forceRole: true,
                                });
                            }
                        } catch (e) {}
                        result = {
                            success: true,
                            message: 'Password updated successfully. Please sign in again.',
                        };
                    }
                }
            }

            return result.success
                ? result
                : { success: false, message: result.message || 'Current password incorrect.' };
        } catch (err) {
            console.error(err);
            return { success: false, message: 'Password change failed.' };
        }
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

        const norm = (x) => String(x || '').trim().toLowerCase().replace(/\s+/g, '_');
        const roleDef = roles.find(
            (r) =>
                r &&
                (String(r.id) === role ||
                    String(r.name) === role ||
                    String(r.roleName || '') === role ||
                    String(r.id || '').toLowerCase() === role.toLowerCase() ||
                    String(r.name || '').toLowerCase() === role.toLowerCase() ||
                    norm(r.roleName) === norm(role) ||
                    norm(r.name) === norm(role) ||
                    String(r.id) === String(currentUser.role_id || '') ||
                    norm(r.roleName) === norm(currentUser.role_name))
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
        await new Promise((resolve) => setTimeout(resolve, 200));
        const applyRole = (baseUser) => {
            const uName = (baseUser && baseUser.username) || username;
            let assigned = null;
            try {
                assigned = typeof resolveAssignedRole === 'function' ? resolveAssignedRole(uName) : null;
            } catch (e) {}
            const role = (assigned && assigned.role) || (baseUser && baseUser.role) || 'employee';
            return {
                ...(baseUser || {}),
                username: uName,
                name: (baseUser && baseUser.name) || (assigned && assigned.name) || uName,
                role,
                system_role: role,
                role_id: (assigned && assigned.role_id) || (baseUser && baseUser.role_id) || null,
                role_name: (assigned && assigned.role_name) || (baseUser && baseUser.role_name) || null,
                modules: (baseUser && baseUser.modules) || (assigned && assigned.modules) || null,
            };
        };
        try {
            try {
                const { login: apiLogin } = await import('../../unity/api.js');
                await apiLogin(username, password);
            } catch (apiErr) {
                console.warn('[MAT] API login failed (CRUD may stay local-only):', apiErr && apiErr.message);
            }
            const result = attemptSharedLogin(username, password, allEmployees || []);
            if (result.success && result.user) {
                const user = applyRole(result.user);
                setCurrentUser(user);
                setIsLoggedIn(true);
                try { saveSharedSession(user); } catch (e) {}
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
                    gender: i % 2 === 0 ? 'Male' : 'Female',
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
        // Do not inject hard-coded seed over live/DB state
        // (seed remains only as INITIAL_ALL_DATA until bootstrap)
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
        const id = (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now());
        const currentCollection = allData[collectionName] || [];
        let recordToSave = { ...newData, id: newData.id || id };

        // Material item aliases
        if (collectionName === 'MAT_ITEM_CRUD' || collectionName === 'MAT_INVENTORY_CRUD') {
            const nm = recordToSave.name || recordToSave.item_name || '';
            recordToSave.name = nm;
            recordToSave.item_name = nm;
            if (recordToSave.qty_on_hand == null && recordToSave.quantity != null) recordToSave.qty_on_hand = recordToSave.quantity;
            if (recordToSave.quantity == null && recordToSave.qty_on_hand != null) recordToSave.quantity = recordToSave.qty_on_hand;
            if (recordToSave.unit_cost == null && recordToSave.unit_price != null) recordToSave.unit_cost = recordToSave.unit_price;
        }

        setAllData(prev => ({
            ...prev,
            [collectionName]: [recordToSave, ...(prev[collectionName] || [])],
        }));

        try { postMatToFinIfNeeded(collectionName, recordToSave); } catch (e) {}
        try {
            if (['HR_EMPLOYEE_CRUD', 'EMPLOYEE_CRUD', 'SYS_USERS_CRUD', 'USERS_CRUD', 'HR_USER_CRUD'].includes(collectionName)) {
                upsertSharedUser({ ...recordToSave, modules: recordToSave.modules || ['material'] });
            }
            if (collectionName === 'ROLE_CRUD') {
                const prev = allData.ROLE_CRUD || [];
                saveSharedRoles([recordToSave, ...prev.filter(r => String(r.id) !== String(recordToSave.id))]);
            }
        } catch (e) {}

        // PostgreSQL — same pattern as Finance
        try {
            const mod = await import('../../unity/api.js').catch(() => import('../../unity/api'));
            if (mod && typeof mod.createRecord === 'function') {
                await mod.createRecord(collectionName, recordToSave);
                console.info('[MAT addRecord] DB ok', collectionName, recordToSave.id);
            }
        } catch (apiErr) {
            console.warn('[MAT addRecord] API persist skipped', collectionName, apiErr?.message || apiErr);
            setMessage({ text: 'Saved in Material; DB: ' + (apiErr?.message || apiErr), type: 'error' });
        }

        return true;
    };

    
const updateRecord = useCallback(async (collectionName, id, updatedFields) => {
        let fields = { ...updatedFields };
        if (collectionName === 'MAT_ITEM_CRUD' || collectionName === 'MAT_INVENTORY_CRUD') {
            const nm = fields.name || fields.item_name;
            if (nm) { fields.name = nm; fields.item_name = nm; }
            if (fields.qty_on_hand == null && fields.quantity != null) fields.qty_on_hand = fields.quantity;
            if (fields.quantity == null && fields.qty_on_hand != null) fields.quantity = fields.qty_on_hand;
            if (fields.unit_cost == null && fields.unit_price != null) fields.unit_cost = fields.unit_price;
        }

        setAllData(prev => {
            const list = prev[collectionName] || [];
            const updatedList = list.map(r =>
                String(r.id) === String(id) || String(r.sku || '') === String(id)
                    ? { ...r, ...fields }
                    : r
            );
            return { ...prev, [collectionName]: updatedList };
        });

        try {
            const mod = await import('../../unity/api.js').catch(() => import('../../unity/api'));
            if (mod && typeof mod.updateRecord === 'function') {
                await mod.updateRecord(collectionName, id, fields);
                console.info('[MAT updateRecord] DB ok', collectionName, id);
            }
        } catch (apiErr) {
            console.warn('[MAT updateRecord] API persist skipped', collectionName, apiErr?.message || apiErr);
            setMessage({ text: 'Updated in Material; DB: ' + (apiErr?.message || apiErr), type: 'error' });
        }

        setMessage({ text: 'System Synchronized Successfully', type: 'success' });
        return true;
    }, []);

    
const deleteRecord = useCallback(async (collectionName, id) => {
        if (!collectionName || !id) {
            setMessage({ text: `Data missing to delete record from ${collectionName}.`, type: 'error' });
            return false;
        }

        try {
            setAllData(prev => ({
                ...prev,
                [collectionName]: (prev[collectionName] || []).filter(record => String(record.id) !== String(id))
            }));

            if (collectionName === 'HR_LEAVE_CRUD') {
                setLeaveUpdateTrigger(prev => prev + 1);
            }

            try {
                const mod = await import('../../unity/api.js').catch(() => import('../../unity/api'));
                if (mod && typeof mod.deleteRecord === 'function') {
                    await mod.deleteRecord(collectionName, id);
                    console.info('[MAT deleteRecord] DB ok', collectionName, id);
                }
            } catch (apiErr) {
                console.warn('[MAT deleteRecord] API persist skipped', collectionName, apiErr?.message || apiErr);
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
/** Dark in-app select — opens above nearby UI (filters, tables, cards) */
/** Dark select — full readable width for long labels (campus, COA, employee, etc.) */
/** Dark select — menu portaled to document.body so it always stacks above tables/modals */
const UUSelect = ({ value, onChange, options = [], allLabel = 'All', className = '', compact = false, placeholder }) => {
    const [open, setOpen] = useState(false);
    const [menuBox, setMenuBox] = useState(null);
    const ref = useRef(null);
    const btnRef = useRef(null);
    const list = Array.isArray(options) ? options : [];

    const placeMenu = React.useCallback(() => {
        const el = btnRef.current;
        if (!el) return;
        const r = el.getBoundingClientRect();
        const gap = 4;
        const maxH = 320;
        const spaceBelow = window.innerHeight - r.bottom - gap;
        const spaceAbove = r.top - gap;
        const openUp = spaceBelow < 160 && spaceAbove > spaceBelow;
        const height = Math.min(maxH, openUp ? spaceAbove : spaceBelow, maxH);
        const width = Math.min(Math.max(r.width, 260), Math.min(window.innerWidth - 16, 520));
        let left = r.left;
        if (left + width > window.innerWidth - 8) left = Math.max(8, window.innerWidth - width - 8);
        if (left < 8) left = 8;
        const top = openUp ? Math.max(8, r.top - gap - height) : r.bottom + gap;
        setMenuBox({ top, left, width, maxHeight: height, openUp });
    }, []);

    useEffect(() => {
        if (!open) return undefined;
        placeMenu();
        const onDoc = (e) => {
            if (ref.current && !ref.current.contains(e.target)
                && !(e.target && e.target.closest && e.target.closest('[data-uu-select-menu]'))) {
                setOpen(false);
            }
        };
        const onReposition = () => placeMenu();
        document.addEventListener('mousedown', onDoc);
        document.addEventListener('touchstart', onDoc);
        window.addEventListener('resize', onReposition);
        window.addEventListener('scroll', onReposition, true);
        return () => {
            document.removeEventListener('mousedown', onDoc);
            document.removeEventListener('touchstart', onDoc);
            window.removeEventListener('resize', onReposition);
            window.removeEventListener('scroll', onReposition, true);
        };
    }, [open, placeMenu]);

    const labelFor = (v) => {
        if (v === 'ALL' || v === 'All' || v === 'all') return allLabel;
        return String(v ?? '');
    };
    const currentLabel = (value === '' || value == null || value === 'ALL' || value === 'All')
        ? (placeholder || allLabel || '-- SELECT --')
        : labelFor(value);

    const portalFn = (typeof createPortal === 'function')
        ? createPortal
        : (typeof ReactDOM !== 'undefined' && ReactDOM.createPortal ? ReactDOM.createPortal : null);

    const menu = (open && menuBox) ? (
        <div
            data-uu-select-menu="1"
            role="listbox"
            className="overflow-y-auto rounded-xl border border-white/25 bg-slate-900 shadow-2xl shadow-black/80 text-white"
            style={{
                position: 'fixed',
                top: menuBox.top,
                left: menuBox.left,
                width: menuBox.width,
                maxHeight: menuBox.maxHeight,
                zIndex: 2147483646,
            }}
        >
            {list.length === 0 ? (
                <div className="px-4 py-3 text-sm text-white/50">No options</div>
            ) : list.map((s) => {
                const val = (s && typeof s === 'object') ? (s.value ?? s.id ?? s.name) : s;
                const lab = (s && typeof s === 'object') ? (s.label ?? s.name ?? String(val)) : labelFor(s);
                const selected = String(value) === String(val);
                return (
                    <button
                        key={String(val)}
                        type="button"
                        role="option"
                        aria-selected={selected}
                        className={`w-full text-left px-4 py-3.5 min-h-[48px] text-sm border-b border-white/10 last:border-b-0
                                    flex items-center justify-between gap-3
                                    ${selected ? 'bg-cyan-600/40 text-white' : 'text-white hover:bg-white/10'}`}
                        style={{ fontSize: 15, lineHeight: 1.4 }}
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => { onChange(val); setOpen(false); }}
                    >
                        <span className="whitespace-normal break-words text-left pr-2">{lab}</span>
                        <span className={`w-4 h-4 rounded-full border-2 shrink-0 ${selected ? 'border-cyan-300 bg-cyan-400' : 'border-white/40'}`} />
                    </button>
                );
            })}
        </div>
    ) : null;

    const menuNode = menu && portalFn && typeof document !== 'undefined' && document.body
        ? portalFn(menu, document.body)
        : menu;

    return (
        <div
            className={`relative ${className || ''}`}
            ref={ref}
            style={{
                zIndex: open ? 50 : 2,
                isolation: 'isolate',
                minWidth: compact ? 200 : 260,
                maxWidth: '100%',
                width: compact ? undefined : '100%',
            }}
        >
            <button
                ref={btnRef}
                type="button"
                onClick={() => setOpen((o) => !o)}
                className={`w-full flex items-center justify-between gap-2 bg-slate-900/90 border border-white/20 rounded-xl
                            text-white outline-none focus:border-cyan-500 focus:ring-1 focus:ring-cyan-500/40
                            ${compact ? 'px-3 py-2.5 text-sm min-h-[44px]' : 'px-4 py-3 text-base min-h-[48px]'}`}
                style={{ fontSize: compact ? 14 : 15 }}
                aria-haspopup="listbox"
                aria-expanded={open}
            >
                <span
                    className={`text-left flex-1 min-w-0 ${(!value || value === 'ALL' || value === 'All') ? 'text-white/55' : 'text-white'}`}
                    style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                >
                    {currentLabel}
                </span>
                <span className="text-cyan-400 text-xs shrink-0 ml-2">{open ? '▲' : '▼'}</span>
            </button>
            {menuNode}
        </div>
    );
};







const Input = ({ label, type, value, onChange, placeholder, isRequired = false, options = [], hint, fieldKey, intelModule, intelCollection }) => {
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
            const el = selectBtnRef.current;
            if (!el) return;
            const r = el.getBoundingClientRect();
            const gap = 4;
            const maxH = 320;
            const spaceBelow = window.innerHeight - r.bottom - gap;
            const spaceAbove = r.top - gap;
            const openUp = spaceBelow < 160 && spaceAbove > spaceBelow;
            const height = Math.min(maxH, Math.max(120, openUp ? spaceAbove : spaceBelow));
            const width = Math.min(Math.max(r.width, 220), Math.min(window.innerWidth - 16, 480));
            let left = r.left;
            if (left + width > window.innerWidth - 8) left = Math.max(8, window.innerWidth - width - 8);
            if (left < 8) left = 8;
            const top = openUp ? Math.max(8, r.top - gap - height) : r.bottom + gap;
            setSelectMenuBox({ top, left, width, maxHeight: height });
        };
        place();
        const onDoc = (e) => {
            if (selectRootRef.current && selectRootRef.current.contains(e.target)) return;
            if (e.target && e.target.closest && e.target.closest('[data-uu-input-select-menu]')) return;
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
            <div className="flex flex-col space-y-1.5 mb-0 group relative min-w-0 max-w-full w-full overflow-visible" ref={selectRootRef}>
                {label && (
                <label 
                    onClick={handleLabelClick}
                    className="text-[9px] font-black uppercase tracking-[0.12em] text-white/50 group-focus-within:text-cyan-400 flex items-center gap-1 max-w-full cursor-help"
                    title={label}
                >
                    <span className="truncate">{label}{isRequired ? <span className="text-red-500"> *</span> : null}</span>
                    <IntelTip label={label} fieldKey={fieldKey || label} module={intelModule} collection={intelCollection} help={hint} />
                </label>
                )}
                <button
                    type="button"
                    id={id}
                    ref={selectBtnRef}
                    onClick={() => setSelectOpen((o) => !o)}
                    className="w-full max-w-full min-w-0 min-h-[42px] flex items-center justify-between gap-2
                               px-3 py-2.5 rounded-xl bg-black/60 border border-white/20 text-left
                               text-white text-sm focus:ring-2 focus:ring-cyan-500/50 outline-none
                               hover:bg-black/80 transition-all cursor-pointer box-border"
                    style={{ fontSize: 14 }}
                    aria-haspopup="listbox"
                    aria-expanded={selectOpen}
                >
                    <span className={current ? 'text-white' : 'text-white/50'} style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {current ? current.label : (placeholder || '-- SELECT --')}
                    </span>
                    <span className="text-cyan-400/80 text-xs shrink-0">{selectOpen ? '▲' : '▼'}</span>
                </button>
                {selectOpen && selectMenuBox && createPortal(
                    <div
                        data-uu-input-select-menu="1"
                        className="overflow-y-auto rounded-xl border border-white/25 bg-slate-900 shadow-2xl shadow-black/80"
                        style={{
                            position: 'fixed',
                            top: selectMenuBox.top,
                            left: selectMenuBox.left,
                            width: selectMenuBox.width,
                            maxHeight: selectMenuBox.maxHeight,
                            zIndex: 2147483646,
                        }}
                        role="listbox"
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
                    className={`text-[10px] font-black uppercase tracking-[0.2em] transition-all flex items-center gap-1 max-w-full cursor-help
                        ${(label.toLowerCase().includes('experience') || label.toLowerCase().includes('beneficiary')) ? 'text-cyan-400 cursor-pointer hover:underline' : 'text-white/50'}`}
                >
                    <span className="truncate">{label}{isRequired ? <span className="text-red-500"> *</span> : null}</span>
                    <IntelTip label={label} fieldKey={fieldKey || label} module={intelModule} collection={intelCollection} help={hint} />
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
                            <span className="text-[11px] font-black text-white uppercase italic">Ethiopian Calendar</span>
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
const ContainerCard = ({ children, className = '', title, subtitle, actions }) => (
    <div className={`rounded-[1.75rem] border border-white/10 bg-black/40 backdrop-blur-md shadow-xl ${className}`}>
        {(title || actions) && (
            <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-4 border-b border-white/10">
                <div>
                    {title ? <h3 className="text-sm font-black uppercase tracking-widest text-white">{title}</h3> : null}
                    {subtitle ? <p className="text-[10px] text-gray-500 mt-0.5">{subtitle}</p> : null}
                </div>
                {actions ? <div className="flex flex-wrap items-center gap-2">{actions}</div> : null}
            </div>
        )}
        <div className="p-5">{children}</div>
    </div>
);

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
        { id: 'MAT_ITEM', label: 'Material · Inventory Master', collection: 'MAT_ITEM_CRUD' },
        { id: 'MAT_GRV', label: 'Material · GRV', collection: 'MAT_GRV_CRUD' },
        { id: 'MAT_GIV', label: 'Material · GIV', collection: 'MAT_GIV_CRUD' },
        { id: 'MAT_ROB', label: 'Material · Robbing', collection: 'MAT_ROBBING_CRUD' },
        { id: 'MAT_LPO', label: 'Material · Local Purchase', collection: 'MAT_PURCHASE_LOCAL_CRUD' },
        { id: 'MAT_FPO', label: 'Material · Foreign Purchase', collection: 'MAT_PURCHASE_FOREIGN_CRUD' },
    ];

    const getCollectionLabel = useCallback((name) => {
        if (!name || typeof name !== 'string') return '';
        return name.replace(/_/g, ' ').replace('CRUD', '(Data)').trim();
    }, []);

    const collectionNames = useMemo(() => Object.keys(allData || {}).sort(), [allData]);
    const initialSelected = collectionNames.length > 0 ? collectionNames[0] : '';

    const [selectedCollection, setSelectedCollection] = useState(initialSelected);
    useEffect(() => {
        if (!selectedCollection && collectionNames.length) {
            setSelectedCollection(collectionNames[0]);
        }
    }, [collectionNames, selectedCollection]);
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

    const partners = (typeof getCollection === 'function' ? (getCollection('FIN_PARTNERS_CRUD') || getCollection('HR_PARTNERS_CRUD') || []) : []);

    return (
        <ContainerCard Icon="FileText" glassClass={glassClass} className="h-full max-w-full min-w-0 overflow-x-hidden overflow-y-auto">
            <div className="space-y-3 mb-4 p-3 sm:p-4 border border-white/10 rounded-xl max-w-full min-w-0 overflow-hidden bg-black/20">
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

                <div className="w-full max-w-full min-w-0 overflow-x-auto overflow-y-visible">
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3 items-end min-w-0 w-full">
                        <div className="min-w-0 max-w-full w-full overflow-hidden">
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

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
                            <Input
                                label="Data Collection"
                                type="select"
                                options={collectionNames.map((name) => ({ label: getCollectionLabel(name), value: name }))}
                                value={selectedCollection}
                                onChange={(val) => setSelectedCollection(val)}
                                isRequired
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
                            <label className="text-[10px] font-black uppercase text-gray-500 tracking-widest">Date from</label>
                            <input
                                type="date"
                                value={dateFrom}
                                onChange={(e) => setDateFrom(e.target.value)}
                                className="w-full mt-1 bg-black/50 border border-white/15 rounded-xl px-3 py-2.5 text-white text-sm outline-none focus:ring-1 focus:ring-cyan-500/40"
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
                            <label className="text-[10px] font-black uppercase text-gray-500 tracking-widest">Date to</label>
                            <input
                                type="date"
                                value={dateTo}
                                onChange={(e) => setDateTo(e.target.value)}
                                className="w-full mt-1 bg-black/50 border border-white/15 rounded-xl px-3 py-2.5 text-white text-sm outline-none focus:ring-1 focus:ring-cyan-500/40"
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
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

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
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
                            <div key={col.key} className="min-w-0 max-w-full w-full overflow-hidden">
                                <Input
                                    label={`Filter: ${col.label}`}
                                    type="select"
                                    options={columnOptions[col.key] || [{ label: 'All', value: 'All' }]}
                                    value={columnFilters[col.key] || ''}
                                    onChange={(val) => handleFilterChange(col.key, val)}
                                />
                            </div>
                        ))}

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
                            <Input
                                label="Report Type"
                                type="select"
                                options={REPORT_TYPES}
                                value={reportType}
                                onChange={(val) => setReportType(val)}
                            />
                        </div>

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
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

                        <div className="min-w-0 max-w-full w-full overflow-hidden">
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


/* --- HARDENED MULTIENTRYFIELD (add-modal sub-forms) --- */
const MultiEntryField = ({ label, entries, fieldConfig, onChange, glassClass }) => {
    const safeConfig = Array.isArray(fieldConfig) ? fieldConfig.filter(Boolean) : [];
    const safeEntries = Array.isArray(entries) ? entries : [];
    const emptyRow = () => safeConfig.reduce((acc, f) => {
        if (f && f.key) acc[f.key] = '';
        return acc;
    }, {});

    const [newEntry, setNewEntry] = React.useState(() => emptyRow());

    React.useEffect(() => {
        setNewEntry(emptyRow());
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [label]);

    const handleAdd = () => {
        if (!safeConfig.length) {
            alert('No fields configured for this sub-form.');
            return;
        }
        const isComplete = safeConfig.every((f) => f.isReadOnly || String(newEntry[f.key] ?? '').trim());
        if (isComplete) {
            onChange([...(safeEntries || []), { ...newEntry }]);
            setNewEntry(emptyRow());
        } else {
            alert('Please complete the entry form below.');
        }
    };

    return (
        <div className={`p-6 sm:p-8 rounded-[2rem] space-y-6 border border-white/10 ${glassClass || ''} mb-8 bg-black/40 shadow-2xl scroll-mt-20`}>
            <label className="block text-[14px] font-black text-cyan-400 border-b border-cyan-500/30 pb-4 uppercase italic tracking-widest">
                {label || 'Entries'}
            </label>
            <div className="space-y-3 max-h-80 overflow-y-auto pr-2 custom-scrollbar">
                {safeEntries.length > 0 ? (
                    safeEntries.map((entry, index) => (
                        <div key={index} className="flex justify-between items-start p-5 rounded-2xl bg-white/5 border border-white/10 hover:bg-cyan-900/20 transition-all">
                            <div className="grid grid-cols-2 lg:grid-cols-4 gap-6 w-full text-[11px]">
                                {safeConfig.map((f) => (
                                    <div key={f.key} className="flex flex-col">
                                        <span className="text-cyan-500 font-black uppercase text-[9px] mb-1">{f.label}:</span>
                                        <span className="text-white font-medium break-words">
                                            {entry?.[f.key] ?? '—'}
                                        </span>
                                    </div>
                                ))}
                            </div>
                            <button
                                type="button"
                                onClick={() => onChange(safeEntries.filter((_, i) => i !== index))}
                                className="ml-4 text-red-500 hover:text-red-400 p-2"
                            >
                                <Icons.Trash2 size={20} />
                            </button>
                        </div>
                    ))
                ) : (
                    <p className="text-[10px] text-white/20 italic p-6 text-center border-2 border-dashed border-white/5 rounded-2xl uppercase tracking-widest font-black">
                        Awaiting entry input...
                    </p>
                )}
            </div>
            <div className="pt-8 border-t border-white/10 space-y-6">
                <p className="text-[10px] font-black text-cyan-400/50 uppercase tracking-[0.3em]">Data Entry Interface:</p>
                <div className="grid grid-cols-1 gap-y-6 bg-black/60 p-6 sm:p-8 rounded-[2rem] border border-white/5">
                    {safeConfig.map((f) => {
                        if (!f || !f.key) return null;
                        const opts = Array.isArray(f.options) ? f.options : [];
                        const isSelect = f.type === 'select' || opts.length > 0;
                        return (
                            <div key={f.key} className="flex flex-col space-y-2">
                                <label className="text-[11px] font-black text-gray-400 uppercase ml-1 tracking-widest">
                                    {f.label || f.key}
                                </label>
                                {isSelect ? (
                                    <select
                                        value={newEntry[f.key] || ''}
                                        onChange={(e) => setNewEntry((prev) => ({ ...prev, [f.key]: e.target.value }))}
                                        className="w-full p-4 rounded-2xl bg-black border border-white/10 text-white text-sm focus:border-cyan-500 outline-none"
                                    >
                                        <option value="">Select {f.label || f.key}...</option>
                                        {opts.map((opt) => {
                                            const val = typeof opt === 'object' ? (opt.value ?? opt.label ?? '') : opt;
                                            const lab = typeof opt === 'object' ? (opt.label ?? opt.value ?? '') : opt;
                                            return <option key={String(val)} value={String(val)}>{String(lab)}</option>;
                                        })}
                                    </select>
                                ) : (
                                    <input
                                        type={f.type || 'text'}
                                        autoComplete="off"
                                        readOnly={!!f.isReadOnly}
                                        value={newEntry[f.key] || ''}
                                        onChange={(e) => setNewEntry((prev) => ({ ...prev, [f.key]: e.target.value }))}
                                        className={`w-full p-4 rounded-2xl bg-black border border-white/10 text-white text-sm focus:border-cyan-500 ${f.isReadOnly ? 'opacity-40 cursor-not-allowed' : ''}`}
                                        placeholder={f.isReadOnly ? 'Generated' : `Provide ${f.label || f.key}...`}
                                    />
                                )}
                            </div>
                        );
                    })}
                </div>
                <button
                    type="button"
                    onClick={handleAdd}
                    className="w-full py-4 rounded-2xl bg-cyan-600/90 hover:bg-cyan-500 text-white text-xs font-black uppercase tracking-widest"
                >
                    + Add Entry
                </button>
            </div>
        </div>
    );
};

/** Viewport-fixed glass modal — always centered on screen, not next to table rows */
/** Viewport-fixed glass modal — portaled to document.body so it is never stuck mid-table */
/** Content-sized glass modal — page stays visible behind; does not stretch full viewport */
/** Always viewport-centered via portal to document.body — never next to table/button */
const Modal = ({ isOpen, onClose, title, children, glassClass = '', footer }) => {
    if (!isOpen) return null;

    const panel = (
        <div
            className={`w-full max-w-lg sm:max-w-xl md:max-w-2xl max-h-[min(88dvh,720px)] flex flex-col
                        overflow-hidden rounded-2xl border border-white/25 bg-slate-900/95 backdrop-blur-xl
                        text-white shadow-[0_20px_60px_rgba(0,0,0,0.55)] ring-1 ring-cyan-400/20
                        ${glassClass || ''}`}
            role="dialog"
            aria-modal="true"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
        >
            <div className="h-1 w-full shrink-0 bg-gradient-to-r from-cyan-500 via-sky-400 to-violet-500" />
            <div className="shrink-0 flex items-center justify-between gap-3 px-4 py-3 border-b border-white/15 bg-white/5">
                <h3 className="text-sm sm:text-base font-black uppercase tracking-wide text-cyan-300 truncate">{title}</h3>
                <button type="button" onClick={onClose}
                    className="p-2 rounded-xl border border-white/20 bg-white/5 min-h-[40px] min-w-[40px]
                               flex items-center justify-center text-gray-300 hover:text-red-300"
                    aria-label="Close">
                    <Icons.X className="w-5 h-5" />
                </button>
            </div>
            <div className="flex-1 min-h-0 overflow-y-auto overflow-x-hidden px-4 py-4 overscroll-contain custom-scrollbar">
                {children}
            </div>
            <div className="shrink-0 flex flex-wrap justify-end gap-2 px-4 py-3 border-t border-white/15 bg-slate-900/95
                            pb-[max(0.5rem,env(safe-area-inset-bottom))]">
                {footer}
                <button type="button" onClick={onClose}
                    className="px-4 py-2.5 rounded-xl text-xs font-bold uppercase border border-white/20
                               bg-white/5 text-gray-200 min-h-[40px]">
                    Cancel
                </button>
            </div>
        </div>
    );

    const shell = (
        <div
            style={{
                position: 'fixed',
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                zIndex: 2147483000,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                padding: '12px',
                boxSizing: 'border-box',
            }}
            role="presentation"
        >
            <div
                style={{
                    position: 'absolute',
                    inset: 0,
                    background: 'rgba(2, 6, 23, 0.35)',
                    backdropFilter: 'blur(2px)',
                    WebkitBackdropFilter: 'blur(2px)',
                }}
                aria-hidden="true"
            />
            <div style={{ position: 'relative', zIndex: 1, width: '100%', maxWidth: '42rem', display: 'flex', justifyContent: 'center' }}>
                {panel}
            </div>
        </div>
    );

    const portalFn = (typeof createPortal === 'function')
        ? createPortal
        : (typeof ReactDOM !== 'undefined' && ReactDOM.createPortal ? ReactDOM.createPortal : null);

    if (portalFn && typeof document !== 'undefined' && document.body) {
        return portalFn(shell, document.body);
    }
    return shell;
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
    const btnRef = useRef(null);
    const [menuBox, setMenuBox] = useState(null);

    const getFilename = () => String(filename || 'Unity_Export').replace(/\s+/g, '_');
    const visibleCols = () => (columns || []).filter((c) => c && !c.isHidden && c.type !== 'array');

    const placeExportMenu = React.useCallback(() => {
        const el = btnRef.current;
        if (!el) return;
        const r = el.getBoundingClientRect();
        const width = 256;
        let left = r.right - width;
        if (left < 8) left = 8;
        if (left + width > window.innerWidth - 8) left = Math.max(8, window.innerWidth - width - 8);
        const gap = 8;
        const spaceBelow = window.innerHeight - r.bottom - gap;
        const openUp = spaceBelow < 220 && r.top > spaceBelow;
        const maxHeight = Math.min(360, openUp ? r.top - gap : Math.max(spaceBelow, 180));
        const top = openUp ? Math.max(8, r.top - gap - maxHeight) : r.bottom + gap;
        setMenuBox({ top, left, width, maxHeight });
    }, []);

    useEffect(() => {
        if (!isOpen) return undefined;
        placeExportMenu();
        const onDoc = (e) => {
            if (btnRef.current && btnRef.current.contains(e.target)) return;
            if (e.target && e.target.closest && e.target.closest('[data-uu-export-menu]')) return;
            setIsOpen(false);
        };
        const onRepo = () => placeExportMenu();
        document.addEventListener('mousedown', onDoc);
        document.addEventListener('touchstart', onDoc);
        window.addEventListener('resize', onRepo);
        window.addEventListener('scroll', onRepo, true);
        return () => {
            document.removeEventListener('mousedown', onDoc);
            document.removeEventListener('touchstart', onDoc);
            window.removeEventListener('resize', onRepo);
            window.removeEventListener('scroll', onRepo, true);
        };
    }, [isOpen, placeExportMenu]);

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
                if (typeof exportTableWorkbook === 'function') {
                    await exportTableWorkbook(rows, cols.length ? cols : (columns || []), filename || fname, {
                        sheetName: String(filename || 'Data').slice(0, 31),
                    });
                } else if (typeof generateRobustWorkbook === 'function' && isFinancial) {
                    await generateRobustWorkbook(rows, filename || fname);
                } else {
                    alert('Excel export is not available in this module.');
                }
                return;
            }
            if (type === 'PDF') {
                if (typeof convertToPDF === 'function') convertToPDF(rows, cols, filename || fname);
                else alert('PDF export is not available.');
                return;
            }
            if (type === 'DOCX') {
                if (typeof convertToDOCX === 'function') await convertToDOCX(rows, cols, filename || fname);
                else alert('Word export is not available.');
                return;
            }
            if (type === 'CSV') {
                const esc = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
                const lines = [
                    'UNITY UNIVERSITY',
                    String(filename || 'Export').toUpperCase() + ' (Official Record)',
                    `Generated ${new Date().toLocaleString()} · ${rows.length} records`,
                    '',
                    cols.map((c) => esc(c.label || c.key)).join(','),
                    ...rows.map((r) => cols.map((c) => esc(r[c.key])).join(',')),
                    '',
                    'Unity University · Official system export · Confidential',
                ];
                saveAs(new Blob(['\ufeff' + lines.join('\n')], { type: 'text/csv;charset=utf-8' }), `${fname}_${new Date().toISOString().slice(0, 10)}.csv`);
            }
        } catch (err) {
            console.error('[EXPORT]', err);
            alert('Export failed: ' + (err && err.message ? err.message : String(err)));
        }
    };

    const menuItems = (
        <div className="p-2 grid grid-cols-1 gap-1">
            <button type="button" onClick={() => handleAction('XLSX')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-emerald-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileSpreadsheet size={16}/> Excel Spreadsheet</button>
            <button type="button" onClick={() => handleAction('DOCX')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-blue-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileBox size={16}/> Word Document</button>
            <button type="button" onClick={() => handleAction('PDF')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-rose-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.FileText size={16}/> PDF Document</button>
            <button type="button" onClick={() => handleAction('CSV')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-gray-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.Filter size={16}/> CSV (Raw Data)</button>
            <hr className="border-white/5 mx-2 my-1" />
            <button type="button" onClick={() => handleAction('PNG')} className="w-full text-left px-4 py-3 text-[11px] font-bold text-cyan-400 hover:bg-white/5 transition flex items-center gap-3 rounded-xl"><Icons.Camera size={16}/> Image Capture (PNG)</button>
        </div>
    );

    return (
        <div className="relative inline-block text-left shrink-0">
            <button
                ref={btnRef}
                type="button"
                onClick={() => setIsOpen((v) => !v)}
                className="bg-cyan-600 hover:bg-cyan-500 text-white font-bold py-2.5 px-6 rounded-xl flex items-center gap-2 shadow-lg transition-all active:scale-95 text-[11px] uppercase tracking-widest border border-white/20"
            >
                <Icons.Download size={16}/> Export Options
            </button>
            {(() => {
                if (!isOpen || !menuBox) return null;
                const menu = (
                    <div
                        data-uu-export-menu="1"
                        className="rounded-2xl bg-slate-900 border border-white/15 shadow-2xl overflow-hidden backdrop-blur-xl"
                        style={{
                            position: 'fixed',
                            top: menuBox.top,
                            left: menuBox.left,
                            width: menuBox.width,
                            maxHeight: menuBox.maxHeight,
                            zIndex: 2147483646,
                            overflowY: 'auto',
                        }}
                        onMouseDown={(e) => e.stopPropagation()}
                    >
                        {menuItems}
                    </div>
                );
                const portalFn = (typeof createPortal === 'function')
                    ? createPortal
                    : (typeof ReactDOM !== 'undefined' && ReactDOM.createPortal ? ReactDOM.createPortal : null);
                if (portalFn && typeof document !== 'undefined' && document.body) {
                    return portalFn(menu, document.body);
                }
                return menu;
            })()}
        </div>
    );
};



const DataTable = ({ collectionName, data, columns, glassClass, title, getCollection }) => {
    const { deleteRecord, updateRecord, postMaterialToFinance, postAllMaterialToFinance } = useData();
    const { checkPermission } = useAuth();

    const [isModalOpen, setIsModalOpen] = useState(false);
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
                        <div className="relative min-w-[220px] flex-1 max-w-sm">
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
                            <button type="button" onClick={handleAdd}
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

            <Modal
                isOpen={isModalOpen}
                onClose={() => setIsModalOpen(false)}
                title={editingRecord && editingRecord.id ? `Edit ${getAddLabel()}` : `New ${getAddLabel()}`}
            >
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
    const { getCollection, notificationTrigger, leaveUpdateTrigger, roleUpdateTrigger, allData } = useData();
    const [activeTab, setActiveTab] = useState('all');
    const [filter, setFilter] = useState('all'); // all | alert | warning | info

    const today = useMemo(() => {
        const d = new Date();
        d.setHours(0, 0, 0, 0);
        return d;
    }, []);

    const daysUntil = (dateStr) => {
        if (!dateStr) return null;
        const t = new Date(String(dateStr).slice(0, 10));
        if (isNaN(t.getTime())) return null;
        t.setHours(0, 0, 0, 0);
        return Math.round((t - today) / 86400000);
    };

    const notifications = useMemo(() => {
        const list = [];
        const push = (sev, cat, message, meta = {}) => {
            list.push({
                id: `${cat}-${list.length}-${String(message).slice(0, 24)}`,
                severity: sev, // alert | warning | info | success
                category: cat,
                message,
                time: meta.time || new Date().toISOString(),
                ...meta
            });
        };

        // ——— Collections ———
        const txs = getCollection('transactions') || [];
        const journals = getCollection('dayJournals') || [];
        const coa = getCollection('chartsOfAccounts') || getCollection('FIN_COA_CRUD') || [];
        const partners = getCollection('FIN_PARTNERS_CRUD') || [];
        const process = getCollection('PAY_PROCESS_CRUD') || [];
        const master = getCollection('PAY_MASTER_CRUD') || [];
        const ot = getCollection('PAY_OVERTIME_CRUD') || [];
        const ad = getCollection('PAY_AD_CRUD') || [];
        const recon = getCollection('PAY_RECON_CRUD') || [];
        const employees = getCollection('HR_EMPLOYEE_CRUD') || [];
        const leave = getCollection('HR_LEAVE_CRUD') || [];
        const vacancies = getCollection('HR_VACANCY_CRUD') || [];
        const recruitment = getCollection('HR_RECRUITMENT_CRUD') || [];
        const attrition = getCollection('HR_STAFF_ATTRITION_CRUD') || [];
        const retirement = getCollection('RETIREMENT_CRUD') || [];
        const resignation = getCollection('RESIGNATION_CRUD') || [];
        const training = getCollection('HR_TRAINING_CRUD') || [];
        const appraisal = getCollection('HR_APPRAISAL_CRUD') || [];
        const insurance = getCollection('HR_INSURANCE_CRUD') || [];
        const penalties = getCollection('HR_PENALTY_CRUD') || [];
        const awards = getCollection('HR_AWARD_CRUD') || getCollection('HR_AWARD_CRUD_HR') || [];
        const depts = getCollection('DEPARTMENTS_CRUD') || [];
        const hiring = getCollection('DEPARTMENTS_HIRING_REQUESTS_CRUD') || [];
        const roles = getCollection('ROLE_CRUD') || [];

        // ——— FINANCE / BOOKS ———
        let tb = [];
        try {
            if (typeof calculateTrialBalance === 'function') {
                tb = calculateTrialBalance(txs, coa) || [];
            }
        } catch (e) { /* ignore */ }
        const totalD = tb.reduce((s, r) => s + Number(r.debit || 0), 0);
        const totalC = tb.reduce((s, r) => s + Number(r.credit || 0), 0);
        const gap = Math.abs(totalD - totalC);
        if (gap > 0.05) {
            push('alert', 'Finance', `Trial Balance unbalanced by ${gap.toLocaleString(undefined, { minimumFractionDigits: 2 })} ETB. Review journals and single transactions.`);
        } else if (totalD > 0) {
            push('success', 'Finance', `Books balanced · Σ Dr/Cr ${totalD.toLocaleString(undefined, { minimumFractionDigits: 2 })} ETB.`);
        } else {
            push('info', 'Finance', 'No posted ledger activity yet. Record single transactions or journals.');
        }

        const draftJournals = journals.filter(j => !['posted', 'verified'].includes(String(j.status || '').toLowerCase()));
        if (draftJournals.length) {
            push('warning', 'Finance', `${draftJournals.length} day journal(s) not Posted/Verified.`);
        }
        const unbalancedTxn = txs.filter(t => {
            const d = Number(t.debit || 0), c = Number(t.credit || 0);
            const a = Number(t.amount || 0);
            if (String(t.is_balanced || '').toLowerCase() === 'no') return true;
            return false;
        });
        if (unbalancedTxn.length) {
            push('alert', 'Finance', `${unbalancedTxn.length} transaction leg(s) marked not balanced.`);
        }
        const matLegs = txs.filter(t => String(t.module || '').toUpperCase() === 'MATERIAL' || String(t.source || '').startsWith('MAT_'));
        if (matLegs.length) {
            push('info', 'Finance', `${matLegs.length} material-linked ledger leg(s) in books (GRV/GIV/PO).`);
        }

        // Partners
        const inactiveP = partners.filter(p => String(p.status || '').toLowerCase() === 'inactive');
        const highAR = partners.filter(p => Number(p.outstanding_balance || 0) > 100000 && String(p.partner_type || '').toLowerCase().includes('customer'));
        const highAP = partners.filter(p => Number(p.outstanding_balance || 0) > 50000 && String(p.partner_type || '').toLowerCase().includes('supplier'));
        if (highAR.length) push('warning', 'Partners', `${highAR.length} customer(s) with AR outstanding over 100,000 ETB.`);
        if (highAP.length) push('warning', 'Partners', `${highAP.length} supplier(s) with AP outstanding over 50,000 ETB.`);
        if (inactiveP.length) push('info', 'Partners', `${inactiveP.length} inactive partner record(s).`);
        if (!partners.length) push('warning', 'Partners', 'Partner master is empty — single transactions cannot autofill entities.');

        // ——— PAYROLL ———
        const draftPay = process.filter(r => ['draft', 'pending'].includes(String(r.status || '').toLowerCase()));
        const committedPay = process.filter(r => String(r.status || '').toLowerCase() === 'committed');
        const postedPay = process.filter(r => String(r.status || '').toLowerCase().includes('post'));
        if (draftPay.length) push('warning', 'Payroll', `${draftPay.length} payroll process row(s) still Draft/Pending — commit to post to finance.`);
        if (committedPay.length && !postedPay.length) {
            push('info', 'Payroll', `${committedPay.length} committed process row(s). Confirm ledger legs on Trial Balance.`);
        }
        const missingMaster = process.filter(r => {
            const name = r.employee_name;
            return name && !(master || []).some(m => m.employee_name === name);
        });
        if (missingMaster.length) push('alert', 'Payroll', `${missingMaster.length} process row(s) without salary master match.`);
        if (ot.length) {
            const openOt = ot.filter(r => !['approved', 'posted', 'closed'].includes(String(r.status || '').toLowerCase()));
            if (openOt.length) push('info', 'Payroll', `${openOt.length} overtime record(s) awaiting approval/close.`);
        }
        if (ad.length) push('info', 'Payroll', `${ad.length} allowance/deduction master line(s) on file.`);
        const openRecon = recon.filter(r => String(r.status || '').toLowerCase() !== 'closed' && Math.abs(Number(r.net_balance || 0)) > 0.05);
        if (openRecon.length) push('warning', 'Payroll', `${openRecon.length} payroll reconciliation(s) with non-zero net balance.`);
        if (!process.length) push('info', 'Payroll', 'No payroll process rows this period.');

        // ——— HR: Leave ———
        const pendingLeave = leave.filter(r => ['pending', 'submitted', 'requested'].includes(String(r.status || '').toLowerCase()));
        if (pendingLeave.length) push('warning', 'HR Leave', `${pendingLeave.length} leave request(s) pending approval.`);
        const onLeaveNow = leave.filter(r => {
            const st = String(r.status || '').toLowerCase();
            if (!(st.includes('approv') || st === 'active' || st === 'ongoing')) return false;
            const a = daysUntil(r.start_date || r.from_date);
            const b = daysUntil(r.end_date || r.to_date);
            if (a == null || b == null) return false;
            return a <= 0 && b >= 0;
        });
        if (onLeaveNow.length) push('info', 'HR Leave', `${onLeaveNow.length} staff currently on approved leave.`);
        const leaveSoon = leave.filter(r => {
            const d = daysUntil(r.start_date || r.from_date);
            return d != null && d > 0 && d <= 7 && String(r.status || '').toLowerCase().includes('approv');
        });
        if (leaveSoon.length) push('info', 'HR Leave', `${leaveSoon.length} approved leave starting within 7 days.`);

        // ——— HR: Headcount / masters ———
        if (!employees.length) push('alert', 'HR', 'Employee master is empty.');
        else push('success', 'HR', `${employees.length} employee record(s) · ${depts.length} department(s).`);
        const noDept = employees.filter(e => !e.department && !e.dept);
        if (noDept.length) push('warning', 'HR', `${noDept.length} employee(s) missing department.`);
        const inactiveEmp = employees.filter(e => ['inactive', 'terminated', 'resigned'].includes(String(e.status || '').toLowerCase()));
        if (inactiveEmp.length) push('info', 'HR', `${inactiveEmp.length} inactive/terminated employee flag(s).`);

        // Vacancies / recruitment / hiring
        const openVac = vacancies.filter(v => ['open', 'published', 'active'].includes(String(v.status || '').toLowerCase()));
        if (openVac.length) push('info', 'Recruitment', `${openVac.length} open vacancy(ies).`);
        const pendingRec = recruitment.filter(r => !['hired', 'rejected', 'closed'].includes(String(r.status || '').toLowerCase()));
        if (pendingRec.length) push('warning', 'Recruitment', `${pendingRec.length} recruitment case(s) in progress.`);
        const openHiring = hiring.filter(h => ['pending', 'submitted', 'open'].includes(String(h.status || '').toLowerCase()));
        if (openHiring.length) push('warning', 'Recruitment', `${openHiring.length} department hiring request(s) pending.`);

        // Attrition / retirement / resignation
        if (attrition.length) push('info', 'HR Exit', `${attrition.length} attrition record(s) on file.`);
        const retireSoon = (retirement.length ? retirement : employees).filter(r => {
            const d = daysUntil(r.retirement_date || r.expected_date);
            return d != null && d >= 0 && d <= 90;
        });
        if (retireSoon.length) push('warning', 'HR Exit', `${retireSoon.length} retirement(s) within 90 days.`);
        const openRes = resignation.filter(r => !['processed', 'completed', 'closed'].includes(String(r.status || '').toLowerCase()));
        if (openRes.length) push('warning', 'HR Exit', `${openRes.length} resignation(s) not fully processed.`);

        // Training / appraisal / insurance / discipline / awards
        const openTrain = training.filter(t => !['completed', 'cancelled'].includes(String(t.status || '').toLowerCase()));
        if (openTrain.length) push('info', 'HR Dev', `${openTrain.length} training item(s) open.`);
        const dueAppraisal = appraisal.filter(a => {
            const d = daysUntil(a.due_date || a.appraisal_date);
            return d != null && d <= 14;
        });
        if (dueAppraisal.length) push('warning', 'HR Dev', `${dueAppraisal.length} appraisal(s) due within 14 days.`);
        const insExpire = insurance.filter(i => {
            const d = daysUntil(i.expiry_date || i.end_date);
            return d != null && d >= 0 && d <= 30;
        });
        if (insExpire.length) push('alert', 'HR', `${insExpire.length} insurance policy(ies) expiring within 30 days.`);
        const openPen = penalties.filter(p => !['closed', 'resolved'].includes(String(p.status || '').toLowerCase()));
        if (openPen.length) push('warning', 'HR', `${openPen.length} open disciplinary case(s).`);
        if (awards.length) push('success', 'HR', `${awards.length} award record(s) registered.`);

        // Roles
        if (!roles.length) push('warning', 'Security', 'ROLE_CRUD empty — relying on default permission map.');
        else push('info', 'Security', `${roles.length} role definition(s) loaded.`);

        // Holidays
        const holidays = [
            { name: 'Genna (Ethiopian Christmas)', date: '2026-01-07' },
            { name: 'Timkat (Epiphany)', date: '2026-01-19' },
            { name: 'Adwa Victory Day', date: '2026-03-02' },
            { name: 'Siklet (Good Friday)', date: '2026-04-10' },
            { name: 'Fasika (Easter)', date: '2026-04-12' },
            { name: 'Labour Day', date: '2026-05-01' },
            { name: 'Ethiopian New Year', date: '2026-09-11' },
            { name: 'Meskel', date: '2026-09-27' },
        ];
        holidays.forEach(h => {
            const d = daysUntil(h.date);
            if (d != null && d >= 0 && d <= 45) {
                push(d <= 7 ? 'warning' : 'info', 'Calendar', `${h.name} in ${d} day(s) (${h.date}).`);
            }
        });

        // Compliance reminders
        push('info', 'Compliance', 'VAT / withholding calendar: confirm monthly filing with Finance (Ato Eyob).');
        push('info', 'Compliance', 'Payroll statutory (tax & pension) must match process totals before bank advice.');

        // Sort: alerts first
        const order = { alert: 0, warning: 1, info: 2, success: 3 };
        list.sort((a, b) => (order[a.severity] ?? 9) - (order[b.severity] ?? 9));
        return list;
    }, [getCollection, notificationTrigger, leaveUpdateTrigger, roleUpdateTrigger, today]);

    const filtered = useMemo(() => {
        let rows = notifications;
        if (filter !== 'all') rows = rows.filter(n => n.severity === filter);
        if (activeTab !== 'all') rows = rows.filter(n => String(n.category || '').toLowerCase().includes(activeTab.toLowerCase()));
        return rows;
    }, [notifications, filter, activeTab]);

    const counts = useMemo(() => ({
        all: notifications.length,
        alert: notifications.filter(n => n.severity === 'alert').length,
        warning: notifications.filter(n => n.severity === 'warning').length,
        info: notifications.filter(n => n.severity === 'info').length,
        success: notifications.filter(n => n.severity === 'success').length,
    }), [notifications]);

    useEffect(() => {
        if (typeof onCountChange === 'function') {
            const badge = (counts.alert || 0) + (counts.warning || 0) || counts.all || 0;
            onCountChange(badge, counts);
        }
    }, [counts, onCountChange]);

    const sevStyle = {
        alert: 'border-rose-500/40 bg-rose-950/40 text-rose-100',
        warning: 'border-amber-500/40 bg-amber-950/30 text-amber-100',
        info: 'border-cyan-500/30 bg-cyan-950/30 text-cyan-100',
        success: 'border-emerald-500/30 bg-emerald-950/30 text-emerald-100',
    };
    const sevIcon = { alert: 'AlertTriangle', warning: 'AlertCircle', info: 'Info', success: 'CheckCircle2' };

    const tabs = [
        { id: 'all', label: 'All' },
        { id: 'Finance', label: 'Finance' },
        { id: 'Payroll', label: 'Payroll' },
        { id: 'HR', label: 'HR' },
        { id: 'Partners', label: 'Partners' },
        { id: 'Security', label: 'Security' },
        { id: 'Calendar', label: 'Calendar' },
    ];

    return (
        <div className={`h-full flex flex-col ${glassClass} border-l border-white/10`}>
            <div className="p-4 border-b border-white/10 flex items-center justify-between gap-2">
                <div>
                    <h3 className="text-sm font-black text-white uppercase tracking-widest flex items-center gap-2">
                        <Icons.Bell className="w-4 h-4 text-cyan-400" /> Notifications
                    </h3>
                    <p className="text-[9px] text-gray-500 mt-0.5">
                        {counts.alert} alert · {counts.warning} warning · {counts.info} info · signed in as {currentUser?.username || 'user'}
                    </p>
                </div>
                <button type="button" onClick={onClose} className="p-2 rounded-lg hover:bg-white/10 text-gray-400">
                    <Icons.X className="w-4 h-4" />
                </button>
            </div>

            <div className="px-3 py-2 flex flex-wrap gap-1.5 border-b border-white/5">
                {['all', 'alert', 'warning', 'info', 'success'].map(f => (
                    <button
                        key={f}
                        type="button"
                        onClick={() => setFilter(f)}
                        className={`px-2.5 py-1 rounded-lg text-[9px] font-black uppercase ${
                            filter === f ? 'bg-cyan-600 text-white' : 'bg-white/5 text-gray-400 hover:bg-white/10'
                        }`}
                    >
                        {f}{f !== 'all' ? ` (${counts[f] || 0})` : ` (${counts.all})`}
                    </button>
                ))}
            </div>

            <div className="px-3 py-2 flex flex-wrap gap-1 border-b border-white/5">
                {tabs.map(t => (
                    <button
                        key={t.id}
                        type="button"
                        onClick={() => setActiveTab(t.id)}
                        className={`px-2 py-1 rounded-md text-[9px] font-bold uppercase ${
                            activeTab === t.id ? 'bg-white/15 text-cyan-300' : 'text-gray-500 hover:text-gray-300'
                        }`}
                    >
                        {t.label}
                    </button>
                ))}
            </div>

            <div className="flex-1 overflow-y-auto custom-scrollbar p-3 space-y-2">
                {!filtered.length ? (
                    <p className="text-center text-gray-500 text-xs py-10">No notifications for this filter.</p>
                ) : filtered.map(n => {
                    const Icon = Icons[sevIcon[n.severity]] || Icons.Bell;
                    return (
                        <div key={n.id} className={`rounded-xl border p-3 ${sevStyle[n.severity] || sevStyle.info}`}>
                            <div className="flex items-start gap-2">
                                <Icon className="w-4 h-4 shrink-0 mt-0.5 opacity-90" />
                                <div className="min-w-0 flex-1">
                                    <div className="flex items-center gap-2 flex-wrap">
                                        <span className="text-[9px] font-black uppercase tracking-wider opacity-80">{n.category}</span>
                                        <span className="text-[8px] uppercase opacity-50">{n.severity}</span>
                                        <span className="text-[9px] text-white/50 ml-auto font-mono tabular-nums" title={n.time ? new Date(n.time).toLocaleString() : ''}>
                                            {formatRelativeTime(n.time)}
                                        </span>
                                    </div>
                                    <p className="text-[11px] font-medium mt-0.5 leading-snug">{n.message}</p>
                                </div>
                            </div>
                        </div>
                    );
                })}
            </div>

            <div className="p-3 border-t border-white/10 text-[9px] text-gray-500 text-center">
                Live from HR · Finance · Payroll masters · refresh on data change
            </div>
        </div>
    );
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
                        fieldConfig={Array.isArray(field.arrayConfig) ? field.arrayConfig : []}
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
                        fieldKey={field.key}
                        hint={field.help || field.intel}
                        intelCollection={collectionName}
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
/* —— Ethiopian (Geez) ↔ Gregorian calendar —— */
const ethiopianMonthNames = [
    'Meskerem', 'Tikimt', 'Hidar', 'Tahsas', 'Tir', 'Yekatit',
    'Megabit', 'Miazia', 'Genbot', 'Sene', 'Hamle', 'Nehase', 'Pagumen'
];

const isGregorianLeap = (y) => (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
const isEthiopianLeap = (y) => (y % 4) === 3;

const gregorianToJDN = (year, month, day) => {
    const a = Math.floor((14 - month) / 12);
    const y = year + 4800 - a;
    const m = month + 12 * a - 3;
    return day + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) - 32045;
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

const ethiopianToJDNCore = (year, month, day) => {
    // Standard: JDN = 1723856 + 365*(y-1) + floor(y/4) + 30*(m-1) + d
    return 1723856 + 365 * (year - 1) + Math.floor(year / 4) + 30 * (month - 1) + day;
};

const jdnToEthiopian = (jdn) => {
    const r = jdn - 1723855;
    const year = Math.floor((4 * r - 1) / 1461);
    const t = Math.floor((4 * r - 1) % 1461);
    const n = Math.floor(t / 4);
    const month = Math.floor(n / 30) + 1;
    const day = (n % 30) + 1;
    return { year, month, day };
};

const gregorianToEthiopian = (y, m, d) => jdnToEthiopian(gregorianToJDN(y, m, d));
const ethiopianToGregorian = (y, m, d) => jdnToGregorian(ethiopianToJDNCore(y, m, d));

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
                    <h2 className="text-xl font-black uppercase italic tracking-tighter text-white">Geez ↔ Gregorian</h2>
                    <p className="text-[10px] text-amber-400/80 font-black uppercase tracking-[0.25em]">Ethiopian calendar converter</p>
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

    const exportCSV = (slug) => {
        const data = allData[slug] || [];
        if (!data.length) return;
        const csv = [Object.keys(data[0]).join(','), ...data.map(r => Object.values(r).join(','))].join('\n');
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = `${slug}.csv`; a.click();
    };

    return (
        <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <FieldIntelligenceSettings glassClass={glassClass} />
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

const HrDashboard = ({ glassClass }) => {
    const { getCollection, allData } = useData();
    const chartRef = useRef(null);
    const [metric, setMetric] = useState('headcount');
    const [chartType, setChartType] = useState('bar');
    const [search, setSearch] = useState('');
    const [deptFilter, setDeptFilter] = useState('ALL');
    const [statusFilter, setStatusFilter] = useState('ALL');
    const [exportOpen, setExportOpen] = useState(false);

    const employees = useMemo(() => {
        const raw = (typeof getCollection === 'function' ? getCollection('HR_EMPLOYEE_CRUD') : null)
            || allData?.HR_EMPLOYEE_CRUD
            || [];
        try {
            if (typeof uuMergeHrEmployees === 'function') return uuMergeHrEmployees(raw);
        } catch (e) {}
        return Array.isArray(raw) ? raw : [];
    }, [getCollection, allData]);

    const statusOptions = useMemo(() => {
        const s = new Set(['ALL']);
        (employees || []).forEach((e) => {
            const v = e.employment_type || e.status || e.auth_status;
            if (v) s.add(String(v));
        });
        return Array.from(s);
    }, [employees]);

    const deptOptions = useMemo(() => {
        const s = new Set(['ALL']);
        (employees || []).forEach((e) => {
            if (e.department) s.add(String(e.department));
        });
        (typeof getCollection === 'function' ? getCollection('DEPARTMENTS_CRUD') : null)?.forEach?.((d) => {
            if (d?.name) s.add(String(d.name));
        });
        return Array.from(s);
    }, [employees, getCollection]);

    const leave = useMemo(() => getCollection?.('HR_LEAVE_CRUD') || allData?.HR_LEAVE_CRUD || [], [getCollection, allData]);
    const recruitment = useMemo(() => getCollection?.('HR_RECRUITMENT_CRUD') || allData?.HR_RECRUITMENT_CRUD || [], [getCollection, allData]);
    const vacancies = useMemo(() => getCollection?.('HR_VACANCY_CRUD') || allData?.HR_VACANCY_CRUD || [], [getCollection, allData]);
    const training = useMemo(() => getCollection?.('HR_TRAINING_CRUD') || allData?.HR_TRAINING_CRUD || [], [getCollection, allData]);
    const appraisal = useMemo(() => getCollection?.('HR_APPRAISAL_CRUD') || allData?.HR_APPRAISAL_CRUD || [], [getCollection, allData]);
    const insurance = useMemo(() => getCollection?.('HR_INSURANCE_CRUD') || allData?.HR_INSURANCE_CRUD || [], [getCollection, allData]);
    const attrition = useMemo(() => getCollection?.('HR_STAFF_ATTRITION_CRUD') || allData?.HR_STAFF_ATTRITION_CRUD || [], [getCollection, allData]);
    const awards = useMemo(() => getCollection?.('HR_AWARD_CRUD') || allData?.HR_AWARD_CRUD || [], [getCollection, allData]);
    const penalties = useMemo(() => getCollection?.('HR_PENALTY_CRUD') || allData?.HR_PENALTY_CRUD || [], [getCollection, allData]);
    const depts = useMemo(() => getCollection?.('DEPARTMENTS_CRUD') || allData?.DEPARTMENTS_CRUD || [], [getCollection, allData]);

    const deptNames = useMemo(() => {
        const fromEmp = employees.map((e) => e.department || e.dept).filter(Boolean);
        const fromMaster = depts.map((d) => d.name || d.department || d.title).filter(Boolean);
        return ['ALL', ...new Set([...fromMaster, ...fromEmp])].sort();
    }, [employees, depts]);

    const filteredEmployees = useMemo(() => {
        const q = search.trim().toLowerCase();
        return employees.filter((e) => {
            const name = String(e.name || e.employee_name || '');
            const idn = String(e.id_number || '');
            const dep = String(e.department || e.dept || '');
            const pos = String(e.job_title || e.position || '');
            const st = String(e.status || 'Active');
            if (deptFilter !== 'ALL' && dep !== deptFilter) return false;
            if (statusFilter !== 'ALL' && st.toLowerCase() !== statusFilter.toLowerCase()) return false;
            if (!q) return true;
            return [name, idn, dep, pos, e.tin_no, e.pension_id, e.work_location].join(' ').toLowerCase().includes(q);
        });
    }, [employees, search, deptFilter, statusFilter]);

    const stats = useMemo(() => {
        const active = employees.filter((e) => !/inactive|terminated|resigned/i.test(String(e.status || 'Active'))).length;
        const academic = employees.filter((e) => /academic/i.test(String(e.employment_category || ''))).length;
        const nonAc = employees.length - academic;
        const pendingLeave = leave.filter((r) => /pending|submitted|requested/i.test(String(r.status || ''))).length;
        const openVac = vacancies.filter((v) => /open|published|active/i.test(String(v.status || ''))).length;
        const openRec = recruitment.filter((r) => !/hired|rejected|closed/i.test(String(r.status || ''))).length;
        const openTrain = training.filter((t) => !/completed|cancelled/i.test(String(t.status || ''))).length;
        const payrollBase = employees.reduce((s, e) => s + (parseFloat(e.gross_salary || e.basic_salary || 0) || 0), 0);
        return {
            headcount: employees.length,
            active,
            academic,
            nonAc,
            pendingLeave,
            openVac,
            openRec,
            openTrain,
            attrition: attrition.length,
            awards: awards.length,
            penalties: penalties.length,
            insurance: insurance.length,
            payrollBase,
            depts: deptNames.length - 1,
        };
    }, [employees, leave, vacancies, recruitment, training, attrition, awards, penalties, insurance, deptNames]);

    const chartData = useMemo(() => {
        const map = {};
        employees.forEach((e) => {
            const d = e.department || e.dept || 'Unassigned';
            if (!map[d]) map[d] = { name: d, headcount: 0, salary: 0, academic: 0 };
            map[d].headcount += 1;
            map[d].salary += parseFloat(e.gross_salary || e.basic_salary || 0) || 0;
            if (/academic/i.test(String(e.employment_category || ''))) map[d].academic += 1;
        });
        leave.forEach((r) => {
            const d = r.department || 'Unassigned';
            if (!map[d]) map[d] = { name: d, headcount: 0, salary: 0, academic: 0, leave: 0 };
            map[d].leave = (map[d].leave || 0) + 1;
        });
        return Object.values(map).map((row) => ({
            ...row,
            value: metric === 'salary' ? row.salary : metric === 'leave' ? (row.leave || 0) : metric === 'academic' ? row.academic : row.headcount,
        }));
    }, [employees, leave, metric]);

    const metricCards = [
        { id: 'headcount', title: 'HEADCOUNT', value: stats.headcount, sub: `${stats.active} active · ${stats.depts} depts`, color: 'text-cyan-400' },
        { id: 'salary', title: 'SALARY MASS (ETB)', value: stats.payrollBase.toLocaleString(undefined, { maximumFractionDigits: 0 }), sub: 'Sum of current salary (HR master)', color: 'text-emerald-400' },
        { id: 'leave', title: 'LEAVE PIPELINE', value: stats.pendingLeave, sub: `${leave.length} leave records total`, color: 'text-amber-400' },
        { id: 'recruit', title: 'RECRUITMENT', value: stats.openRec, sub: `${stats.openVac} open vacancies`, color: 'text-violet-400' },
        { id: 'academic', title: 'ACADEMIC / NON-AC', value: `${stats.academic} / ${stats.nonAc}`, sub: 'Employment category split', color: 'text-sky-400' },
        { id: 'dev', title: 'TRAINING OPEN', value: stats.openTrain, sub: `${stats.awards} awards · ${stats.penalties} cases`, color: 'text-rose-400' },
    ];

    const exportWorkbook = async (format) => {
        const rows = filteredEmployees.map((e) => ({
            'ID NO': e.id_number || '',
            'Full Name': e.name || e.employee_name || '',
            Gender: e.gender || '',
            Department: e.department || e.dept || '',
            'Position / Rank': e.job_title || e.position || '',
            'Work Location': e.work_location || '',
            Category: e.employment_category || '',
            Type: e.employment_type || '',
            'Current Salary': e.gross_salary || e.basic_salary || '',
            'Pension ID': e.pension_id || '',
            'TIN NO': e.tin_no || '',
            Mobile: e.telephone_number || '',
            Status: e.status || 'Active',
            Role: e.system_role || e.role || '',
        }));
        const stamp = new Date().toISOString().slice(0, 10);
        if (format === 'CSV') {
            if (!rows.length) return;
            const keys = Object.keys(rows[0]);
            const csv = [keys.join(','), ...rows.map((r) => keys.map((k) => `"${String(r[k] ?? '').replace(/"/g, '""')}"`).join(','))].join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            saveAs(blob, `HR_Employees_${stamp}.csv`);
            return;
        }
        if (format === 'XLSX' || format === 'EXCEL') {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet('HR Employees');
            ws.mergeCells('A1:N1');
            ws.getCell('A1').value = 'UNITY UNIVERSITY — HR EMPLOYEE MASTER';
            ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF0E7490' } };
            ws.mergeCells('A2:N2');
            ws.getCell('A2').value = `Export ${stamp} · ${rows.length} record(s) · filter: ${deptFilter}/${statusFilter} · search: ${search || '—'}`;
            if (rows.length) {
                const keys = Object.keys(rows[0]);
                ws.addRow([]);
                const header = ws.addRow(keys);
                header.eachCell((c) => {
                    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF155E75' } };
                });
                rows.forEach((r) => ws.addRow(keys.map((k) => r[k])));
                keys.forEach((_, idx) => { ws.getColumn(idx + 1).width = 16; });
            }
            const buf = await wb.xlsx.writeBuffer();
            saveAs(new Blob([buf]), `HR_Employees_${stamp}.xlsx`);
            return;
        }
        if (format === 'PDF') {
            const doc = new jsPDF({ orientation: 'landscape' });
            doc.setFontSize(14);
            doc.text('Unity University — HR Employee Master', 14, 16);
            doc.setFontSize(9);
            doc.text(`Exported ${stamp} · ${rows.length} rows`, 14, 22);
            const body = rows.map((r) => [r['ID NO'], r['Full Name'], r.Department, r['Position / Rank'], r['Current Salary'], r.Status]);
            doc.autoTable({
                startY: 26,
                head: [['ID', 'Name', 'Department', 'Position', 'Salary', 'Status']],
                body,
                styles: { fontSize: 7 },
                headStyles: { fillColor: [21, 94, 117] },
            });
            doc.save(`HR_Employees_${stamp}.pdf`);
        }
    };

    const ChartBody = () => {
        if (!chartData.length) return <p className="text-center text-gray-500 text-sm py-16">No departmental data yet.</p>;
        const common = { data: chartData, margin: { top: 8, right: 12, left: 0, bottom: 0 } };
        if (chartType === 'pie') {
            return (
                <ResponsiveContainer width="100%" height={280}>
                    <PieChart>
                        <Pie data={chartData} dataKey="value" nameKey="name" outerRadius={100} label>
                            {chartData.map((_, i) => (
                                <Cell key={i} fill={['#06b6d4', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#6366f1', '#ec4899'][i % 7]} />
                            ))}
                        </Pie>
                        <Tooltip />
                        <Legend />
                    </PieChart>
                </ResponsiveContainer>
            );
        }
        const Comp = chartType === 'line' ? LineChart : chartType === 'area' ? AreaChart : BarChart;
        return (
            <ResponsiveContainer width="100%" height={280}>
                <Comp {...common}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                    <XAxis dataKey="name" tick={{ fill: '#94a3b8', fontSize: 10 }} />
                    <YAxis tick={{ fill: '#94a3b8', fontSize: 10 }} />
                    <Tooltip contentStyle={{ background: '#0f172a', border: '1px solid #334155' }} />
                    <Legend />
                    {chartType === 'line' ? (
                        <Line type="monotone" dataKey="value" stroke="#06b6d4" strokeWidth={2} name={metric} />
                    ) : chartType === 'area' ? (
                        <Area type="monotone" dataKey="value" fill="#06b6d4" stroke="#22d3ee" name={metric} />
                    ) : (
                        <Bar dataKey="value" fill="#06b6d4" name={metric} radius={[6, 6, 0, 0]} />
                    )}
                </Comp>
            </ResponsiveContainer>
        );
    };

    return (
        <div className="space-y-5 animate-in fade-in duration-500 pb-8">
            {/* Header */}
            <div className={`${glassClass} rounded-2xl border border-white/10 p-5 flex flex-col lg:flex-row lg:items-center gap-4 justify-between`}>
                <div>
                    <h2 className="text-xl font-black text-white tracking-tight flex items-center gap-2">
                        <span className="text-cyan-400">◆</span> HR COMMAND CENTER
                    </h2>
                    <p className="text-[10px] text-gray-400 mt-1 uppercase tracking-widest">
                        Unity University · Live employee master · leave · recruitment · development
                    </p>
                </div>
                <div className="flex flex-wrap gap-2 items-center">
                    <div className="relative">
                        <input
                            type="search"
                            value={search}
                            onChange={(e) => setSearch(e.target.value)}
                            placeholder="Search ID, name, dept, TIN, pension…"
                            className="bg-black/40 border border-white/10 rounded-xl px-3 py-2 text-xs text-white w-56 focus:outline-none focus:border-cyan-500"
                        />
                    </div>
                    <UUSelect
                                value={deptFilter}
                                onChange={setDeptFilter}
                                options={deptOptions && deptOptions.length ? deptOptions : ['ALL']}
                                allLabel="All departments"
                                compact
                            />
                    <UUSelect
                                value={statusFilter}
                                onChange={setStatusFilter}
                                options={statusOptions.length ? statusOptions : ['ALL']}
                                allLabel="All statuses"
                                compact
                            />
                    <div className="relative">
                        <button type="button" onClick={() => setExportOpen((v) => !v)} className="px-3 py-2 rounded-xl text-[10px] font-black uppercase bg-cyan-600 hover:bg-cyan-500 text-white">
                            Export ▾
                        </button>
                        {exportOpen && (
                            <div className="absolute right-0 mt-1 z-50 w-44 rounded-xl border border-white/10 bg-slate-900 shadow-xl overflow-hidden">
                                {['XLSX', 'CSV', 'PDF'].map((f) => (
                                    <button key={f} type="button" onClick={() => { exportWorkbook(f); setExportOpen(false); }} className="w-full text-left px-3 py-2 text-[11px] text-gray-200 hover:bg-white/10">
                                        {f === 'XLSX' ? 'Excel workbook' : f}
                                    </button>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            </div>

            {/* KPI cards */}
            <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-3">
                {metricCards.map((c) => (
                    <button
                        key={c.id}
                        type="button"
                        onClick={() => setMetric(c.id === 'recruit' || c.id === 'dev' ? 'headcount' : c.id === 'salary' ? 'salary' : c.id === 'leave' ? 'leave' : c.id === 'academic' ? 'academic' : 'headcount')}
                        className={`${glassClass} rounded-2xl border border-white/10 p-4 text-left hover:border-cyan-500/40 transition ${metric === c.id ? 'ring-1 ring-cyan-500/50' : ''}`}
                    >
                        <p className="text-[9px] font-black uppercase tracking-widest text-gray-500">{c.title}</p>
                        <p className={`text-xl font-black mt-1 ${c.color}`}>{c.value}</p>
                        <p className="text-[10px] text-gray-500 mt-1 leading-snug">{c.sub}</p>
                    </button>
                ))}
            </div>

            {/* Chart + summary */}
            <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
                <div className={`${glassClass} rounded-2xl border border-white/10 p-4 xl:col-span-2`} ref={chartRef}>
                    <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                        <p className="text-xs font-black uppercase tracking-widest text-white">Department profile · {metric}</p>
                        <div className="flex gap-1">
                            {['bar', 'line', 'area', 'pie'].map((t) => (
                                <button key={t} type="button" onClick={() => setChartType(t)} className={`px-2.5 py-1 rounded-lg text-[9px] font-black uppercase ${chartType === t ? 'bg-cyan-600 text-white' : 'bg-white/5 text-gray-400'}`}>{t}</button>
                            ))}
                        </div>
                    </div>
                    <ChartBody />
                </div>
                <div className={`${glassClass} rounded-2xl border border-white/10 p-4 space-y-3`}>
                    <p className="text-xs font-black uppercase tracking-widest text-white">Snapshot</p>
                    {[
                        ['Employees (filtered)', filteredEmployees.length],
                        ['Academic staff', stats.academic],
                        ['Non-academic', stats.nonAc],
                        ['Open vacancies', stats.openVac],
                        ['Recruitment in progress', stats.openRec],
                        ['Pending leave', stats.pendingLeave],
                        ['Attrition records', stats.attrition],
                        ['Insurance policies', stats.insurance],
                    ].map(([k, v]) => (
                        <div key={k} className="flex justify-between text-[11px] border-b border-white/5 pb-1.5">
                            <span className="text-gray-400">{k}</span>
                            <span className="font-bold text-cyan-300">{v}</span>
                        </div>
                    ))}
                </div>
            </div>

            {/* Employee table */}
            <div className={`${glassClass} rounded-2xl border border-white/10 overflow-hidden`}>
                <div className="px-4 py-3 border-b border-white/10 flex items-center justify-between">
                    <p className="text-xs font-black uppercase tracking-widest text-white">Employee master · full information</p>
                    <p className="text-[10px] text-gray-500">{filteredEmployees.length} / {employees.length} shown</p>
                </div>
                <div className="overflow-auto max-h-[420px] custom-scrollbar">
                    <table className="min-w-full text-left text-[11px]">
                        <thead className="sticky top-0 bg-slate-900/95 z-10">
                            <tr className="text-[9px] uppercase tracking-wider text-gray-400">
                                {['ID', 'Full name', 'Dept', 'Position', 'Location', 'Category', 'Salary', 'Pension', 'TIN', 'Mobile', 'Status'].map((h) => (
                                    <th key={h} className="px-3 py-2 font-black whitespace-nowrap">{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {!filteredEmployees.length ? (
                                <tr><td colSpan={11} className="px-3 py-10 text-center text-gray-500">No employees match search / filters. Open HR Employee master or clear filters.</td></tr>
                            ) : filteredEmployees.map((e) => (
                                <tr key={e.id || e.id_number || e.name} className="border-t border-white/5 hover:bg-white/5">
                                    <td className="px-3 py-2 font-mono text-cyan-300 whitespace-nowrap">{e.id_number || '—'}</td>
                                    <td className="px-3 py-2 font-semibold text-white whitespace-nowrap">{e.name || e.employee_name || '—'}</td>
                                    <td className="px-3 py-2 text-gray-300">{e.department || e.dept || '—'}</td>
                                    <td className="px-3 py-2 text-gray-300">{e.job_title || e.position || '—'}</td>
                                    <td className="px-3 py-2 text-gray-400">{e.work_location || '—'}</td>
                                    <td className="px-3 py-2 text-gray-400">{e.employment_category || '—'}</td>
                                    <td className="px-3 py-2 text-emerald-300 font-bold">{Number(e.gross_salary || e.basic_salary || 0).toLocaleString()}</td>
                                    <td className="px-3 py-2 text-gray-400">{e.pension_id || '—'}</td>
                                    <td className="px-3 py-2 text-gray-400">{e.tin_no || '—'}</td>
                                    <td className="px-3 py-2 text-gray-400 whitespace-nowrap">{e.telephone_number || '—'}</td>
                                    <td className="px-3 py-2"><span className="px-2 py-0.5 rounded-full bg-white/5 text-[9px] font-bold uppercase">{e.status || 'Active'}</span></td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

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
    const { getCollection, updateRecord, addRecord } = useData();
    const roles = getCollection('ROLE_CRUD') || [];
    const [userTick, setUserTick] = useState(0);

    // Re-read shared users when tick changes (after assign) or catalog event fires
    useEffect(() => {
        const onUpd = () => setUserTick((t) => t + 1);
        window.addEventListener('uu-shared-users-updated', onUpd);
        return () => window.removeEventListener('uu-shared-users-updated', onUpd);
    }, []);

    // System login users (created accounts) — NOT HR staff display names
    const users = useMemo(() => {
        const fromShared = (typeof getSharedUsers === 'function' ? getSharedUsers() : []) || [];
        const fromSys = getCollection('SYS_USERS_CRUD') || getCollection('USERS_CRUD') || [];
        const byUser = new Map();
        fromShared.forEach((u) => {
            if (!u || !u.username) return;
            byUser.set(String(u.username).toLowerCase(), {
                id: u.id || `user-${u.username}`,
                username: String(u.username).trim(),
                name: u.name || u.display_name || u.username,
                role: u.role || u.system_role || 'employee',
                modules: u.modules || null,
                password: u.password,
                source: 'shared',
            });
        });
        (Array.isArray(fromSys) ? fromSys : []).forEach((u) => {
            const username = String(u.username || u.user_name || '').trim();
            if (!username) return;
            const prev = byUser.get(username.toLowerCase()) || {};
            // Shared catalog role wins over stale collection role
            byUser.set(username.toLowerCase(), {
                ...prev,
                id: u.id || prev.id || `user-${username}`,
                username,
                name: u.name || u.full_name || prev.name || username,
                role: prev.role || u.role || u.system_role || 'employee',
                modules: u.modules || u.allowed_modules || prev.modules,
                password: u.password != null ? u.password : prev.password,
                source: 'collection',
                _raw: u,
            });
        });
        return Array.from(byUser.values()).sort((a, b) =>
            String(a.username).localeCompare(String(b.username))
        );
    }, [getCollection, userTick]);

    const [selectedUsername, setSelectedUsername] = useState('');
    const [selectedRoleId, setSelectedRoleId] = useState('');

    const selectedUser = users.find(
        (u) => String(u.username).toLowerCase() === String(selectedUsername).toLowerCase()
    );

    const handleAssignment = async () => {
        if (!selectedUsername || !selectedRoleId) {
            return alert('Please select both a system user and a role.');
        }
        const user = users.find(
            (u) => String(u.username).toLowerCase() === String(selectedUsername).toLowerCase()
        );
        const role = roles.find((r) => String(r.id) === String(selectedRoleId));
        if (!user) {
            alert('Selected user not found.');
            return;
        }
        if (!role) {
            alert('Selected role not found.');
            return;
        }
        const roleKey = role.roleName || role.name || role.id;
        // Prefer ROLE_CRUD id so checkPermission finds permissions; keep name for display
        const roleNorm = String(role.id != null ? role.id : roleKey).trim();
        const roleSlug = String(roleKey).trim().toLowerCase().replace(/\s+/g, '_');
        try {
            // Always write shared catalog first (login source of truth)
            if (typeof upsertSharedUser === 'function') {
                upsertSharedUser({
                    username: user.username,
                    password: user.password || '',
                    name: user.name,
                    role: roleNorm,
                    system_role: roleNorm,
                    role_id: role.id,
                    role_name: roleKey,
                    role_slug: roleSlug,
                    modules: user.modules,
                    id: user.id,
                    _forceRole: true,
                });
            }
            // Verify write
            const verify = (typeof getSharedUsers === 'function' ? getSharedUsers() : []).find(
                (x) => String(x.username).toLowerCase() === String(user.username).toLowerCase()
            );
            if (!verify || String(verify.role).toLowerCase() !== roleNorm) {
                // Force raw localStorage rewrite
                try {
                    const key = 'uu_shared_users_v1';
                    const raw = JSON.parse(localStorage.getItem(key) || '[]');
                    const list = Array.isArray(raw) ? raw : [];
                    const idx = list.findIndex((x) => String(x.username || '').toLowerCase() === String(user.username).toLowerCase());
                    const row = {
                        ...(idx >= 0 ? list[idx] : {}),
                        username: user.username,
                        password: user.password || (idx >= 0 ? list[idx].password : '') || '',
                        name: user.name || user.username,
                        role: roleNorm,
                        system_role: roleNorm,
                        role_id: role.id,
                        role_name: roleKey,
                        modules: user.modules,
                        id: user.id,
                    };
                    if (idx >= 0) list[idx] = row;
                    else list.push(row);
                    localStorage.setItem(key, JSON.stringify(list));
                    window.dispatchEvent(new CustomEvent('uu-shared-users-updated', { detail: { users: list } }));
                } catch (e) {
                    console.error(e);
                }
            }

            const collName =
                (getCollection('SYS_USERS_CRUD') || []).length
                    ? 'SYS_USERS_CRUD'
                    : (getCollection('USERS_CRUD') || []).length
                      ? 'USERS_CRUD'
                      : null;
            if (collName && user.id && !String(user.id).startsWith('user-')) {
                const raw = user._raw || user;
                await updateRecord(collName, user.id, {
                    ...raw,
                    username: user.username,
                    role: roleNorm,
                    system_role: roleNorm,
                    role_id: role.id,
                    role_name: roleKey,
                });
            }

            // Sync active sessions
            try {
                const sessKey = 'uu_shared_session_v1';
                const sess = JSON.parse(localStorage.getItem(sessKey) || 'null');
                if (sess && String(sess.username).toLowerCase() === String(user.username).toLowerCase()) {
                    localStorage.setItem(sessKey, JSON.stringify({ ...sess, role: roleNorm, system_role: roleNorm }));
                }
                const authKey = 'uu_auth_session_v1';
                const authSess = JSON.parse(localStorage.getItem(authKey) || 'null');
                if (authSess && String(authSess.username).toLowerCase() === String(user.username).toLowerCase()) {
                    localStorage.setItem(authKey, JSON.stringify({ ...authSess, role: roleNorm, system_role: roleNorm }));
                }
            } catch (e) {}

            setUserTick((t) => t + 1);
            alert(`Role "${roleKey}" assigned to user "${user.username}".
New role: ${roleNorm}
Log out and log in again if the menu does not update immediately.`);
            setSelectedRoleId('');
            // keep username selected so Current role line updates
            setSelectedUsername(user.username);
        } catch (error) {
            console.error('Role assignment failed:', error);
            alert('Role assignment failed. See console for details.');
        }
    };

    return (
        <div className="space-y-6 pb-40">
            <ContainerCard glassClass={glassClass}>
                <div className="p-6 space-y-6">
                    <div>
                        <h2 className="text-xl font-black text-white italic uppercase tracking-tight">Role Assignment</h2>
                        <p className="text-[11px] text-cyan-300/80 font-bold uppercase tracking-widest mt-1">
                            Assign roles to <span className="text-white">system users</span> (login accounts) — not HR staff names
                        </p>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="relative min-w-0">
                            <label className="text-cyan-400 text-[10px] font-bold mb-2 block uppercase tracking-widest">
                                1. Select system user ({users.length} accounts)
                            </label>
                            <select
                                className="w-full bg-black border border-white/20 rounded-xl p-4 text-white focus:border-cyan-500 outline-none cursor-pointer shadow-2xl [color-scheme:dark]"
                                value={selectedUsername}
                                onChange={(e) => setSelectedUsername(e.target.value)}
                            >
                                <option value="">-- Select username --</option>
                                {users.map((u) => (
                                    <option key={u.username} value={u.username}>
                                        {u.username}
                                        {u.name && u.name !== u.username ? ` — ${u.name}` : ''}
                                        {u.role ? ` [role: ${u.role}]` : ''}
                                    </option>
                                ))}
                            </select>
                            {selectedUser && (
                                <p className="mt-2 text-[10px] text-gray-400 font-mono">
                                    Current role: <span className="text-cyan-300">{selectedUser.role || '—'}</span>
                                </p>
                            )}
                        </div>
                        <div className="relative min-w-0">
                            <label className="text-cyan-400 text-[10px] font-bold mb-2 block uppercase tracking-widest">
                                2. Choose permissions role
                            </label>
                            <select
                                className="w-full bg-black border border-white/20 rounded-xl p-4 text-white focus:border-cyan-500 outline-none cursor-pointer shadow-2xl [color-scheme:dark]"
                                value={selectedRoleId}
                                onChange={(e) => setSelectedRoleId(e.target.value)}
                            >
                                <option value="">-- Select role --</option>
                                {roles.map((r) => (
                                    <option key={r.id} value={r.id}>
                                        {r.roleName || r.name || r.id}
                                    </option>
                                ))}
                            </select>
                        </div>
                    </div>
                    <div className="flex flex-wrap gap-3">
                        <button
                            type="button"
                            onClick={handleAssignment}
                            className="px-6 py-3 rounded-xl text-xs font-black uppercase bg-cyan-600 text-white hover:bg-cyan-500 shadow-lg shadow-cyan-900/40"
                        >
                            Assign role to user
                        </button>
                    </div>
                    <div className="overflow-x-auto rounded-xl border border-white/10">
                        <table className="w-full text-left text-sm min-w-[480px]">
                            <thead className="bg-slate-950 text-[10px] uppercase font-black tracking-widest text-cyan-400">
                                <tr>
                                    <th className="p-3">Username</th>
                                    <th className="p-3">Display name</th>
                                    <th className="p-3">Current role</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-white/5">
                                {users.map((u) => (
                                    <tr key={u.username} className="hover:bg-white/[0.03]">
                                        <td className="p-3 font-mono text-cyan-300">{u.username}</td>
                                        <td className="p-3 text-gray-200">{u.name || '—'}</td>
                                        <td className="p-3 text-gray-300">{u.role || '—'}</td>
                                    </tr>
                                ))}
                                {!users.length && (
                                    <tr>
                                        <td colSpan={3} className="p-6 text-center text-gray-500 text-xs uppercase">
                                            No system users found. Create users first, then assign roles here.
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </div>
            </ContainerCard>
        </div>
    );
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
}

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
            <FieldIntelligenceRuntime />
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
                    <button type="button" onClick={() => setShowNotifications(!showNotifications)} className={`relative p-2 rounded-full transition-all ${showNotifications ? 'bg-cyan-500 text-white' : 'text-gray-400 hover:text-cyan-300'}`} aria-label="Notifications">
                        <Icons.Bell className="w-5 h-5" />
                        {notifCount > 0 && (
                            <span className="absolute -top-0.5 -right-0.5 min-w-[1.1rem] h-[1.1rem] px-1 rounded-full bg-red-500 text-white text-[9px] font-black flex items-center justify-center ring-2 ring-slate-950">
                                {notifCount > 99 ? '99+' : notifCount}
                            </span>
                        )}
                    </button>
                </div>
            </header>

            <div className="flex flex-grow overflow-hidden min-h-0 relative">
                {mobileNavOpen && <div className="uu-sidebar-backdrop md:hidden" onClick={() => setMobileNavOpen(false)} />}
                <aside className={`uu-sidebar-desktop hidden md:block w-[20%] min-w-[260px] max-w-[280px] border-r border-white/10 ${glassClass}`}>
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

                {showNotifications && (
                    <aside className={`w-[25%] border-l border-white/10 animate-in slide-in-from-right ${glassClass}`}>
                        <NotificationBar 
                            glassClass={glassClass} 
                            currentUser={currentUser}
                            onCountChange={setNotifCount} 
                            onClose={() => setShowNotifications(false)} 
                        />
                    </aside>
                )}
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
}

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
}

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

