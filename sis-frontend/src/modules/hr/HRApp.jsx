import ModuleAccessGate from '../../unity/ModuleAccessGate.jsx';
import { attemptSharedLogin, upsertSharedUsers, upsertSharedUser, saveSharedRoles, saveSharedSession, loadSharedSession, clearSharedSession, canAccessModule, getSharedRoles, mergeSharedRolesInto } from '../../unity/SharedCredentials.js';
import { afterHrBootstrap, publishHrEmployees } from '../../unity/UnityModuleBridge.js';
import LoginModuleSwitcher from '../../unity/LoginModuleSwitcher.jsx';
import { MobileStyles, useIsMobile } from '../../unity/mobileCompat.js';
               
import React, { createContext, useContext, useState, useEffect, useCallback, useMemo, useRef } from 'react';
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
    AreaChart,
    Area,
    PieChart,
    Pie,
    Cell,
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


/* ===== HR ↔ Finance integration (no extra module) ===== */
const UU_HR_EMP_KEY = 'uu_shared_hr_employees_v1';
const uuLoadHrEmployees = () => {
    try {
        const raw = localStorage.getItem(UU_HR_EMP_KEY);
        if (!raw) return [];
        const v = JSON.parse(raw);
        return Array.isArray(v) ? v : [];
    } catch { return []; }
};
const uuSaveHrEmployees = (list) => {
    try {
        localStorage.setItem(UU_HR_EMP_KEY, JSON.stringify(Array.isArray(list) ? list : []));
    } catch (e) {}
};
const uuNormHrEmp = (e) => {
    if (!e || typeof e !== 'object') return null;
    const name = e.name || e.employee_name || e.full_name || '';
    if (!name) return null;
    return {
        ...e,
        name,
        employee_name: name,
        id_number: e.id_number || e.idNo || '',
        job_title: e.job_title || e.position || '',
        position: e.job_title || e.position || '',
        department: e.department || e.dept || '',
        work_location: e.work_location || e.location || '',
        gross_salary: e.gross_salary ?? e.current_salary ?? e.basic_salary ?? 0,
        basic_salary: e.gross_salary ?? e.current_salary ?? e.basic_salary ?? 0,
        pension_id: e.pension_id || e.pension_ac || '',
        tin_no: e.tin_no || e.tin || '',
        national_id: e.national_id || '',
        telephone_number: e.telephone_number || e.mobile || '',
        seniority_date_gc: e.seniority_date_gc || e.seniority_date || '',
        employment_type: e.employment_type || '',
        employment_category: e.employment_category || '',
        status: e.status || 'Active',
    };
};
const uuMergeHrEmployees = (local = []) => {
    const by = new Map();
    [...(local || []), ...uuLoadHrEmployees()].forEach((e) => {
        const n = uuNormHrEmp(e);
        if (!n) return;
        const k = String(n.id_number || n.name).toLowerCase();
        by.set(k, { ...(by.get(k) || {}), ...n });
    });
    return Array.from(by.values());
};
const uuPublishHrEmployees = (list) => {
    const merged = uuMergeHrEmployees(list || []);
    uuSaveHrEmployees(merged);
    return merged;
};


const SCROLLBAR_STYLE_ID = 'uu-hr-scrollbar-styles';
if (typeof document !== 'undefined' && !document.getElementById(SCROLLBAR_STYLE_ID)) {
    const s = document.createElement('style');
    s.id = SCROLLBAR_STYLE_ID;
    s.textContent = `
      .custom-scrollbar { scrollbar-width: auto; scrollbar-color: #22d3ee99 #0f172a; }
      .custom-scrollbar::-webkit-scrollbar { width: 14px; height: 14px; }
      .custom-scrollbar::-webkit-scrollbar-track { background: #0f172a; border-radius: 8px; }
      .custom-scrollbar::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #22d3ee, #0891b2);
        border-radius: 8px; border: 3px solid #0f172a; min-height: 48px;
      }
      .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #67e8f9; }
      .custom-scrollbar::-webkit-scrollbar-corner { background: #0f172a; }
    `;
    document.head.appendChild(s);
}



// --- Configuration Constants ---
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
const AMHARIC_TOP_EMPLOYEES = [
    'Netsanet Abebe',
    'Damte Tesfaye',
    'Kenu Hailu',
    'Kalab Mekonnen',
    'Abel Bekele',
    'Teferi Assefa',
    'Sara Hailu',
    'Abebe Kebede',
    'Tigist Fanta',
    'Dawit Lemma',
    'Hanna Girma',
    'Jemal Yusuf',
];
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
    // HR Dashboard is visible to all, but content is role-filtered
    'HR_DASHBOARD_VIEW': ['admin', 'hr_manager', 'hr_staff', 'employee'],
  'DEPARTMENTS_CRUD': ['admin', 'hr_manager'],
    'DEPARTMENTS_HIRING_REQUESTS_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'HR_VACANCY_CRUD': ['admin', 'hr_manager', 'hr_staff'], 
    // CRUD Permissions - Stricter 'Work Forwarding' Enforcement:
    'HR_RECRUITMENT_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'HR_EMPLOYEE_CRUD': ['admin', 'hr_manager', 'employee'],
    'HR_COMPETENCE_CRUD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'HR_APPRAISAL_CRUD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'HR_LEAVE_CRUD': ['admin', 'hr_manager', 'hr_staff', 'employee'],

    'HR_PENALTY_CRUD': ['admin', 'hr_manager'],

    'HR_AWARD_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'HR_TRANSFER_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'HR_TRAINING_CRUD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'HR_INSURANCE_CRUD': ['admin', 'hr_manager'],
    'HR_REPORT_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    // New Role Assignment Manager / Settings
    'HR_ROLE_ASSIGNMENT_CRUD': ['admin', 'hr_manager'],
    'UI_CUSTOMIZATION': ['admin', 'hr_manager'], // Only managers/admins can change UI theme
    'HR_GENERIC_REPORTS': ['admin', 'hr_manager', 'hr_staff'], // NEW: Report Generator Access
    'ROLE_CRUD': ['admin'], // Only admin can manage roles

    // Granular CRUD Permissions:
    'CAN_ADD_RECORD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'CAN_EDIT_RECORD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'CAN_DELETE_RECORD': ['admin', 'hr_manager'],


    // Additional Staff Modules
    'HR_STAFF_ASSIGNMENT_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'HR_STAFF_ATTRITION_CRUD': ['admin', 'hr_manager'],
    'HR_STAFF_DEVELOPMENT_CRUD': ['admin', 'hr_manager', 'hr_staff', 'employee'],
    'HR_STAFF_OVERVIEW_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'STAFF_REGISTER': ['admin', 'hr_manager'],
    'STAFF_CHANGE_PASSWORD': ['admin', 'hr_manager', 'employee'],
    'STAFF_UPDATE_PROFILE': ['admin', 'hr_manager', 'employee'],
    'STAFF_SIGNUP': ['admin', 'hr_manager', 'employee'],
    'RETIREMENT_CRUD': ['admin', 'hr_manager'],
    'RESIGNATION_CRUD': ['admin', 'hr_manager', 'hr_staff'],
    'DEATH_CRUD': ['admin', 'hr_manager'],
    // This key MUST match the ID used in your Sidebar MENU_ITEMS
    'ROLE_ASSIGNMENT_VIEW': ['admin', 'hr_manager'], 
    
    
 
};
// --- CORE MENU ITEMS ---
// --- CORE MENU ITEMS ---
const MENU_ITEMS = [
    //  {
        // id: 'ORG_STRUCTURE', label: 'Organization Structure', icon: 'Building', children: [
           // ]
    // },
    {
        id: 'HR', label: 'Human Resources', icon: 'Users', children: [
              { id: 'HR_DASHBOARD_VIEW', label: 'HR Dashboard', icon: 'LayoutDashboard' }, // NEW: Added HR Dashboard here
             { id: 'DEPARTMENTS_CRUD', label: 'Departments (CRUD)', icon: 'Building' },
            { id: 'DEPARTMENTS_HIRING_REQUESTS_CRUD', label: 'Departments Hiring Requests (CRUD)', icon: 'UserPlus' },
        { id: 'HR_VACANCY_CRUD', label: 'Vacancies (CRUD)', icon: 'Briefcase' }, 
            { id: 'HR_RECRUITMENT_CRUD', label: 'Recruitment (CRUD)', icon: 'UserPlus' },
            { id: 'HR_EMPLOYEE_CRUD', label: 'Employees Record (CRUD)', icon: 'UserCheck' },

            { id: 'HR_COMPETENCE_CRUD', label: 'Competence Matrix (CRUD)', icon: 'ListChecks' },
            { id: 'HR_APPRAISAL_CRUD', label: 'Appraisal (CRUD)', icon: 'HeartHandshake' },
            { id: 'HR_LEAVE_CRUD', label: 'Leave Management (CRUD)', icon: 'Clock' },
            { id: 'HR_PENALTY_CRUD', label: 'Disciplinary/Penalties Log (CRUD)', icon: 'Tag' },
            { id: 'HR_AWARD_CRUD_HR', label: 'Award     (CRUD)', icon: 'Trophy' },
            { id: 'HR_TRANSFER_CRUD', label: 'Transfer (CRUD)', icon: 'Send' },
            { id: 'HR_TRAINING_CRUD', label: 'Training (CRUD)', icon: 'GraduationCap' },
            { id: 'HR_INSURANCE_CRUD', label: 'Insurance (CRUD)', icon: 'ShieldCheck' },
            // { id: 'HR_REPORT_CRUD', label: 'Report HR (CRUD)', icon: 'FileText' },
            { id: 'HR_STAFF_ASSIGNMENT_CRUD', label: 'Staff Assignment (CRUD)',
                icon: 'Briefcase'
            },
            { id: 'HR_STAFF_ATTRITION_CRUD', label: 'Staff Attrition (CRUD)', icon: 'TrendingDown' },
            { id: 'HR_AWARD_CRUD', label: 'Staff Awards (CRUD)', icon: 'Award' },
            { id: 'HR_STAFF_DEVELOPMENT_CRUD', label: 'Staff Development (CRUD)', icon: 'Activity' },
            { id: 'HR_STAFF_OVERVIEW_CRUD', label: 'Staff Overview (CRUD)', icon: 'Eye' },
            { id: 'RETIREMENT_CRUD', label: 'Retirements (CRUD)', icon: 'Sunrise' },
            { id: 'RESIGNATION_CRUD', label: 'Resignations (CRUD)', icon: 'LogOut' },
            { id: 'DEATH_CRUD', label: 'Deaths (CRUD)', icon: 'Cross' },
        ]
    },

    {
        id: 'SETTINGS', label: 'Settings', icon: 'Settings', children: [
            { id: 'HR_ROLE_ASSIGNMENT_CRUD', label: 'Role Assignment Manager', icon: 'Fingerprint' },
            { id: 'UI_CUSTOMIZATION', label: 'UI Customization', icon: 'Palette' },
            { id: 'HR_GENERIC_REPORTS', label: 'Reports Generator', icon: 'FileBarChart' }, // NEW: Reports Module
            { id: 'ROLE_CRUD', label: 'Manage Roles',
                icon: 'ShieldAlert'
            }, // New: Role CRUD
            { id: 'STAFF_REGISTER', label: 'Register', icon: 'UserPlus' },
            { id: 'STAFF_CHANGE_PASSWORD', label: 'Change Password', icon: 'Key' },
            { id: 'STAFF_UPDATE_PROFILE', label: 'Update Profile', icon: 'Edit' },
            { id: 'STAFF_SIGNUP', label: 'Signup', icon: 'Edit' },

        ]
    },
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
    const [currentView, setCurrentView] = useState('HR_DASHBOARD_VIEW');

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
                setCurrentView('HR_DASHBOARD_VIEW');
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
    const extraSurnames = ['Bekele', 'Tesfaye', 'Hailu', 'Mekonnen', 'Assefa', 'Girma', 'Lemma', 'Tadesse', 'Wolde', 'Kebede'];
    const employeeNames = [...demoEmployeeNames, ...Array.from({ length: Math.max(0, count - demoEmployeeNames.length) }, (_, i) => {
        const first = MOCK_AMHARIC_NAMES[(i + 1) % MOCK_AMHARIC_NAMES.length];
        const last = extraSurnames[i % extraSurnames.length];
        return `${first} ${last}`;
    })];

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
            'HR_DASHBOARD_VIEW',  'HR_RECRUITMENT_CRUD',  'HR_EMPLOYEE_CRUD',
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
            'HR_DASHBOARD_VIEW',
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
            'HR_DASHBOARD_VIEW',
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
            'HR_DASHBOARD_VIEW',
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

/* ======================================================================
   CROSS-MODULE INTEGRATION OUTLETS
   HR  ↔  MATERIAL  ↔  FINANCE  (Unity University SIS)
   Same transaction leg shape as Material postMaterialToFinance /
   Finance postPayrollToFinance so a unified DataProvider can merge.
   ====================================================================== */

/** Finance GL defaults for HR-originated journals */
const HR_FINANCE_DEFAULT_GL = {
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
};

/** Material store touch-points from HR (staff location / issue eligibility) */
const HR_MATERIAL_OUTLETS = {
    /** Campus / work_location on employee → material store filter */
    employeeLocationField: 'work_location',
    /** Departments that may request GIV */
    issuableDepartmentsCollection: 'DEPARTMENTS_CRUD',
    /** Staff who can perform robbing / receive assets */
    technicianRoleHints: ['maintenance_tech', 'store_manager', 'inventory_clerk'],
};

/**
 * Registry: every HR collection that can post or notify other modules.
 * outlet: 'finance' | 'material' | 'both' | 'hr_only'
 */
const HR_INTEGRATION_REGISTRY = {
    'HR_EMPLOYEE_CRUD': {
        outlet: 'both',
        finance: 'Master for payroll cost centers / partner_name',
        material: 'work_location + department → GIV requester / asset custodian',
        postable: false,
    },
    'HR_LEAVE_CRUD': {
        outlet: 'finance',
        finance: 'Unpaid leave may reduce salary accrual (memo until payroll run)',
        material: null,
        postable: false,
    },
    'HR_PENALTY_CRUD': {
        outlet: 'finance',
        finance: 'Fine amount → Dr Sundry receivable / Cr penalty income (or payroll deduction)',
        material: null,
        postable: true,
        source: 'PENALTY',
        amountKeys: ['fine_amount', 'amount', 'penalty_amount'],
    },
    'HR_AWARD_CRUD': {
        outlet: 'finance',
        finance: 'Cash award → Dr award expense / Cr salary payable or cash',
        material: null,
        postable: true,
        source: 'AWARD',
        amountKeys: ['award_amount', 'amount', 'value'],
    },
    'HR_AWARD_CRUD_HR': {
        outlet: 'finance',
        finance: 'Same as awards',
        material: null,
        postable: true,
        source: 'AWARD',
        amountKeys: ['award_amount', 'amount', 'value'],
    },
    'HR_TRAINING_CRUD': {
        outlet: 'finance',
        finance: 'Training cost → Dr training expense / Cr sundry payable or cash',
        material: null,
        postable: true,
        source: 'TRAINING',
        amountKeys: ['cost', 'training_cost', 'amount', 'budget'],
    },
    'HR_INSURANCE_CRUD': {
        outlet: 'finance',
        finance: 'Premium → Dr insurance expense / Cr sundry payable',
        material: null,
        postable: true,
        source: 'INSURANCE',
        amountKeys: ['premium', 'amount', 'contribution'],
    },
    'HR_TRANSFER_CRUD': {
        outlet: 'material',
        finance: null,
        material: 'toPlace / fromPlace should realign asset custodian & store location',
        postable: false,
    },
    'HR_STAFF_ASSIGNMENT_CRUD': {
        outlet: 'both',
        finance: 'Cost center for payroll',
        material: 'Assigned campus for inventory responsibility',
        postable: false,
    },
    'RETIREMENT_CRUD': {
        outlet: 'both',
        finance: 'Final pay / benefits accrual',
        material: 'Clear asset custody before exit',
        postable: false,
    },
    'RESIGNATION_CRUD': {
        outlet: 'both',
        finance: 'Final settlement',
        material: 'Return issued items / clear GIV open lines',
        postable: false,
    },
    'DEATH_CRUD': {
        outlet: 'both',
        finance: 'Death benefits payable',
        material: 'Asset recovery from estate / department',
        postable: false,
    },
};

const HR_SOURCE_COLLECTION_MAP = {
    PENALTY: 'HR_PENALTY_CRUD',
    AWARD: 'HR_AWARD_CRUD',
    AWARD_HR: 'HR_AWARD_CRUD_HR',
    TRAINING: 'HR_TRAINING_CRUD',
    INSURANCE: 'HR_INSURANCE_CRUD',
};

const HR_COLLECTION_SOURCE_MAP = {
    'HR_PENALTY_CRUD': 'PENALTY',
    'HR_AWARD_CRUD': 'AWARD',
    'HR_AWARD_CRUD_HR': 'AWARD_HR',
    'HR_TRAINING_CRUD': 'TRAINING',
    'HR_INSURANCE_CRUD': 'INSURANCE',
};

const hrMoney = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
};

const pickAmount = (record, keys = []) => {
    for (const k of keys) {
        if (record[k] != null && record[k] !== '' && hrMoney(record[k]) > 0) return hrMoney(record[k]);
    }
    return 0;
};

/**
 * Build balanced finance legs for an HR monetary event.
 * Legs match Finance `transactions` documents (module: 'HR').
 */
const buildHRJournalLegs = (source, record) => {
    if (!record) return [];
    const date = String(record.date || record.award_date || record.training_date || record.penalty_date || new Date().toISOString().slice(0, 10));
    const partner = record.employee_name || record.name || record.staff_name || 'Unity Staff';
    const uid = String(record.id || Date.now());
    const legs = [];
    const meta = Object.values(HR_INTEGRATION_REGISTRY).find(r => r.source === source)
        || HR_INTEGRATION_REGISTRY[HR_SOURCE_COLLECTION_MAP[source]];
    const amountKeys = (meta && meta.amountKeys) || ['amount'];
    const amount = pickAmount(record, amountKeys);

    const pushPair = (ref, amt, drCode, drName, crCode, crName, narration) => {
        if (amt <= 0) return;
        const common = {
            date,
            partner_name: partner,
            amount: amt,
            status: 'Posted',
            module: 'HR',
            source: `HR_${source}`,
            hr_ref: ref,
            hr_record_id: uid,
            description: narration,
            narration,
        };
        legs.push({
            ...common,
            id: `hr-${source}-${uid}-L1`,
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
            id: `hr-${source}-${uid}-L2`,
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

    const GL = HR_FINANCE_DEFAULT_GL;
    const ref = record.penalty_no || record.award_no || record.training_no || record.policy_no || uid;

    if (source === 'PENALTY') {
        pushPair(ref, amount, GL.cashBank, 'Cash / Payroll recovery', GL.penaltyIncome, 'Penalty / fine income',
            `HR Penalty · ${partner} · ${record.penalty_type || record.reason || 'fine'}`);
    } else if (source === 'AWARD' || source === 'AWARD_HR') {
        pushPair(ref, amount, GL.awardExpense, 'Staff award expense', GL.salaryPayable, 'Awards payable',
            `HR Award · ${partner} · ${record.award_category || record.title || 'award'}`);
    } else if (source === 'TRAINING') {
        pushPair(ref, amount, GL.trainingExpense, 'Training expense', GL.sundryPayable, 'Training payable',
            `HR Training · ${partner} · ${record.course_name || record.title || 'course'}`);
    } else if (source === 'INSURANCE') {
        pushPair(ref, amount, GL.insuranceExpense, 'Staff insurance expense', GL.sundryPayable, 'Insurance payable',
            `HR Insurance · ${partner} · ${record.policy_type || record.provider || 'premium'}`);
    }
    return legs;
};

/**
 * Material-side hooks derived from HR (no stock movement — signals only).
 * Call when transfer / resignation / death is saved so Material can clear custody.
 */
const buildMaterialStaffSignals = (collectionName, record) => {
    if (!record) return null;
    const base = {
        module: 'HR',
        collection: collectionName,
        record_id: record.id,
        employee_name: record.employee_name || record.name,
        employee_id: record.employee_id || record.id_number || record.id,
        at: new Date().toISOString(),
    };
    if (collectionName === 'HR_TRANSFER_CRUD') {
        return {
            ...base,
            signal: 'STAFF_LOCATION_CHANGE',
            from_location: record.fromPlace || record.from_location,
            to_location: record.toPlace || record.to_location,
            material_action: 'Reassign open GIV custody / asset tags to new campus store',
        };
    }
    if (collectionName === 'RESIGNATION_CRUD' || collectionName === 'RETIREMENT_CRUD' || collectionName === 'DEATH_CRUD') {
        return {
            ...base,
            signal: 'STAFF_EXIT',
            material_action: 'Block new GIV; require return of issued items; clear asset custodian',
        };
    }
    if (collectionName === 'HR_EMPLOYEE_CRUD') {
        return {
            ...base,
            signal: 'STAFF_MASTER_UPSERT',
            work_location: record.work_location,
            department: record.department,
            material_action: 'Sync requester directory for GIV / robbing performed_by lists',
        };
    }
    return null;
};

/** Shared empty shells so unified app can merge Material + Finance stores */
const CROSS_MODULE_SHELLS = {
    transactions: [],
    dayJournals: [],
    FIN_COA_CRUD: [],
    chartsOfAccounts: [],
    MAT_ITEM_CRUD: [],
    MAT_GRV_CRUD: [],
    MAT_GIV_CRUD: [],
    MAT_ROBBING_CRUD: [],
    MAT_PURCHASE_LOCAL_CRUD: [],
    MAT_PURCHASE_FOREIGN_CRUD: [],
    HR_MATERIAL_SIGNALS: [],
};


const INITIAL_ALL_DATA = {
    'HR_EMPLOYEE_CRUD': typeof MOCK_EMPLOYEES !== 'undefined' ? MOCK_EMPLOYEES : [],
    'DEPARTMENTS_CRUD': typeof MOCK_DEPARTMENTS !== 'undefined' ? MOCK_DEPARTMENTS : [],
    'DEPARTMENTS_HIRING_REQUESTS_CRUD': typeof MOCK_DEPARTMENTS_HIRING_REQUESTS !== 'undefined' ? MOCK_DEPARTMENTS_HIRING_REQUESTS : [],
    'HR_VACANCY_CRUD': typeof MOCK_VACANCY_RECORDS !== 'undefined' ? MOCK_VACANCY_RECORDS : [],
    'RETIREMENT_CRUD': typeof MOCK_RETIREMENT_RECORDS !== 'undefined' ? MOCK_RETIREMENT_RECORDS : [],
    'RESIGNATION_CRUD': typeof MOCK_RESIGNATION_RECORDS !== 'undefined' ? MOCK_RESIGNATION_RECORDS : [],
    'DEATH_CRUD': typeof MOCK_DEATH_RECORDS !== 'undefined' ? MOCK_DEATH_RECORDS : [],
    'HR_RECRUITMENT_CRUD': typeof MOCK_RECRUITMENT_RECORDS !== 'undefined' ? MOCK_RECRUITMENT_RECORDS : [],
    'HR_COMPETENCE_CRUD': typeof MOCK_COMPETENCE_MATRIX !== 'undefined' ? MOCK_COMPETENCE_MATRIX : [],
    'HR_APPRAISAL_CRUD': typeof MOCK_APPRAISAL_RECORDS !== 'undefined' ? MOCK_APPRAISAL_RECORDS : [],
    'HR_LEAVE_CRUD': typeof MOCK_LEAVE_RECORDS !== 'undefined' ? MOCK_LEAVE_RECORDS : [],
    'HR_PENALTY_CRUD': typeof MOCK_PENALTY_RECORDS !== 'undefined' ? MOCK_PENALTY_RECORDS : [],
    'HR_AWARD_CRUD_HR': typeof MOCK_AWARD_RECORDS !== 'undefined' ? MOCK_AWARD_RECORDS : [],
    'HR_TRANSFER_CRUD': typeof MOCK_TRANSFER_RECORDS !== 'undefined' ? MOCK_TRANSFER_RECORDS : [],
    'HR_TRAINING_CRUD': typeof MOCK_TRAINING_RECORDS !== 'undefined' ? MOCK_TRAINING_RECORDS : [],
    'HR_INSURANCE_CRUD': typeof MOCK_INSURANCE_RECORDS !== 'undefined' ? MOCK_INSURANCE_RECORDS : [],
    'HR_STAFF_ASSIGNMENT_CRUD': typeof MOCK_STAFF_ASSIGNMENT_RECORDS !== 'undefined' ? MOCK_STAFF_ASSIGNMENT_RECORDS : [],
    'HR_STAFF_ATTRITION_CRUD': typeof MOCK_STAFF_ATTRITION_RECORDS !== 'undefined' ? MOCK_STAFF_ATTRITION_RECORDS : [],
    'HR_AWARD_CRUD': typeof MOCK_STAFF_AWARDS_RECORDS !== 'undefined' ? MOCK_STAFF_AWARDS_RECORDS : (typeof MOCK_AWARD_RECORDS !== 'undefined' ? MOCK_AWARD_RECORDS : []),
    'HR_STAFF_OVERVIEW_CRUD': typeof MOCK_STAFF_OVERVIEW_RECORDS !== 'undefined' ? MOCK_STAFF_OVERVIEW_RECORDS : [],
    'HR_STAFF_DEVELOPMENT_CRUD': typeof MOCK_STAFF_DEVELOPMENT_RECORDS !== 'undefined' ? MOCK_STAFF_DEVELOPMENT_RECORDS : [],
    'ROLE_CRUD': typeof MOCK_ROLE_RECORDS !== 'undefined' ? MOCK_ROLE_RECORDS : [],
    // Cross-module shells (Material + Finance) — merge when unifying apps
    ...CROSS_MODULE_SHELLS,
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
        // Fallback so tables never go blank if bootstrap wiped or state lagged
        const fallbacks = {
            HR_EMPLOYEE_CRUD: typeof MOCK_EMPLOYEES !== 'undefined' ? MOCK_EMPLOYEES : null,
            DEPARTMENTS_CRUD: typeof MOCK_DEPARTMENTS !== 'undefined' ? MOCK_DEPARTMENTS : null,
            DEPARTMENTS_HIRING_REQUESTS_CRUD: typeof MOCK_DEPARTMENTS_HIRING_REQUESTS !== 'undefined' ? MOCK_DEPARTMENTS_HIRING_REQUESTS : null,
            HR_VACANCY_CRUD: typeof MOCK_VACANCY_RECORDS !== 'undefined' ? MOCK_VACANCY_RECORDS : null,
            RETIREMENT_CRUD: typeof MOCK_RETIREMENT_RECORDS !== 'undefined' ? MOCK_RETIREMENT_RECORDS : null,
            RESIGNATION_CRUD: typeof MOCK_RESIGNATION_RECORDS !== 'undefined' ? MOCK_RESIGNATION_RECORDS : null,
            DEATH_CRUD: typeof MOCK_DEATH_RECORDS !== 'undefined' ? MOCK_DEATH_RECORDS : null,
            HR_RECRUITMENT_CRUD: typeof MOCK_RECRUITMENT_RECORDS !== 'undefined' ? MOCK_RECRUITMENT_RECORDS : null,
            HR_COMPETENCE_CRUD: typeof MOCK_COMPETENCE_MATRIX !== 'undefined' ? MOCK_COMPETENCE_MATRIX : null,
            HR_APPRAISAL_CRUD: typeof MOCK_APPRAISAL_RECORDS !== 'undefined' ? MOCK_APPRAISAL_RECORDS : null,
            HR_LEAVE_CRUD: typeof MOCK_LEAVE_RECORDS !== 'undefined' ? MOCK_LEAVE_RECORDS : null,
            HR_PENALTY_CRUD: typeof MOCK_PENALTY_RECORDS !== 'undefined' ? MOCK_PENALTY_RECORDS : null,
            HR_TRANSFER_CRUD: typeof MOCK_TRANSFER_RECORDS !== 'undefined' ? MOCK_TRANSFER_RECORDS : null,
            HR_TRAINING_CRUD: typeof MOCK_TRAINING_RECORDS !== 'undefined' ? MOCK_TRAINING_RECORDS : null,
            HR_INSURANCE_CRUD: typeof MOCK_INSURANCE_RECORDS !== 'undefined' ? MOCK_INSURANCE_RECORDS : null,
            HR_STAFF_ASSIGNMENT_CRUD: typeof MOCK_STAFF_ASSIGNMENT_RECORDS !== 'undefined' ? MOCK_STAFF_ASSIGNMENT_RECORDS : null,
            HR_STAFF_ATTRITION_CRUD: typeof MOCK_STAFF_ATTRITION_RECORDS !== 'undefined' ? MOCK_STAFF_ATTRITION_RECORDS : null,
            HR_AWARD_CRUD: typeof MOCK_AWARD_RECORDS !== 'undefined' ? MOCK_AWARD_RECORDS : null,
            HR_STAFF_OVERVIEW_CRUD: typeof MOCK_STAFF_OVERVIEW_RECORDS !== 'undefined' ? MOCK_STAFF_OVERVIEW_RECORDS : null,
            HR_STAFF_DEVELOPMENT_CRUD: typeof MOCK_STAFF_DEVELOPMENT_RECORDS !== 'undefined' ? MOCK_STAFF_DEVELOPMENT_RECORDS : null,
            ROLE_CRUD: typeof MOCK_ROLE_RECORDS !== 'undefined' ? MOCK_ROLE_RECORDS : null,
        };
        if (fallbacks[collectionName] && Array.isArray(fallbacks[collectionName]) && fallbacks[collectionName].length) {
            return fallbacks[collectionName];
        }
        return Array.isArray(rows) ? rows : [];
    }, [allData]);

 /* --- SYNCHRONIZED DATA ENGINE --- */



const addRecord = async (collectionName, newData) => {
    const id = Date.now().toString();
    // Use allData (not collections) to match your state variable name
    const currentCollection = allData[collectionName] || [];
    
    let recordToSave = { ...newData, id };

    // Auto-ID logic for Employees
    if (collectionName === 'HR_EMPLOYEE_CRUD') {
        const lastNum = currentCollection.reduce((max, emp) => {
            const num = parseInt(emp.id_number?.split('-')[1]) || 0;
            return Math.max(max, num);
        }, 0);
        recordToSave.id_number = `UU-${String(lastNum + 1).padStart(3, '0')}`;
    }
    
    // Update the state using the correct state setter
    setAllData(prev => ({
        ...prev,
        [collectionName]: [recordToSave, ...currentCollection]
    }));

        try {
            if (['HR_EMPLOYEE_CRUD', 'EMPLOYEE_CRUD', 'SYS_USERS_CRUD', 'USERS_CRUD', 'HR_USER_CRUD'].includes(collectionName)) {
                upsertSharedUser({ ...recordToSave, modules: recordToSave.modules || ['hr'] });
                try {
                    if (collectionName === 'HR_EMPLOYEE_CRUD') {
                        const next = [recordToSave, ...(allData.HR_EMPLOYEE_CRUD || []).filter(x => String(x.id) !== String(recordToSave.id))];
                        uuPublishHrEmployees(next);
                    }
                } catch (e) {}
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
                        // Check every possible name/id match field
                        const matches = (app.employee_name === empName) || 
                                        (app.name === empName) || 
                                        (app.employee_id === empId) || 
                                        (app.staff_national_id === empId);

                        if (matches) {
                            return { 
                                ...app, 
                                score: updatedFields.score,
                                // Also update visual feedback for confirmation
                                manager_feedback: `Updated via Competence Sync: ${new Date().toLocaleDateString()}`
                            };
                        }
                        return app;
                    });
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

    // MODIFIED: Removed localStorage usage for user registration
    const addRole = useCallback(async (roleData) => {
        console.log("Add Role triggered with", roleData);

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
                // Assign all available permissions by default:
                permissions: Object.keys(PERMISSIONS_MAP)
            };
            setAllData(prev => {
                const updatedRoles = [...prev['ROLE_CRUD'], newRole];
                // Removed local storage user registration logic (username/password)
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
          console.log("update Role triggered with", roleId, updatedPermissions);
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
     * Post HR monetary document → Finance transactions (same leg shape as Material/Payroll).
     * source: PENALTY | AWARD | AWARD_HR | TRAINING | INSURANCE
     */
    const postHRToFinance = useCallback((source, recordId) => {
        const collection = HR_SOURCE_COLLECTION_MAP[source];
        if (!collection) {
            setMessage({ text: `Unknown HR finance source: ${source}`, type: 'error' });
            return false;
        }
        const rows = allData[collection] || [];
        const record = rows.find(r => String(r.id) === String(recordId));
        if (!record) {
            setMessage({ text: 'HR record not found for finance post.', type: 'error' });
            return false;
        }
        if (record.posted_to_finance) {
            setMessage({ text: `Already posted (${record.posted_ref || 'ledger'}).`, type: 'error' });
            return false;
        }
        const legs = buildHRJournalLegs(source, record);
        if (!legs.length) {
            setMessage({ text: 'No amount to post — set fine/award/cost/premium on the record.', type: 'error' });
            return false;
        }
        const postedAt = new Date().toISOString().slice(0, 10);
        const ref = legs[0]?.hr_ref || `${source}-${recordId}`;
        setAllData(prev => {
            const currentTx = Array.isArray(prev.transactions) ? prev.transactions : [];
            const currentDj = Array.isArray(prev.dayJournals) ? prev.dayJournals : [];
            return {
                ...prev,
                transactions: [...legs, ...currentTx],
                dayJournals: [...legs, ...currentDj],
                [collection]: (prev[collection] || []).map(r =>
                    String(r.id) === String(recordId)
                        ? { ...r, posted_to_finance: true, posted_at: postedAt, posted_ref: ref, finance_status: 'Posted' }
                        : r
                ),
            };
        });
        setMessage({ text: `Posted to Finance: HR ${source} ${ref} · ${legs.length / 2} journal(s)`, type: 'success' });
        return true;
    }, [allData]);

    const postAllHRToFinance = useCallback((collectionName) => {
        const source = HR_COLLECTION_SOURCE_MAP[collectionName];
        if (!source) return 0;
        const rows = (allData[collectionName] || []).filter(r => !r.posted_to_finance);
        let n = 0;
        rows.forEach(r => { if (postHRToFinance(source, r.id)) n += 1; });
        return n;
    }, [allData, postHRToFinance]);

    /** Emit material signal (custody / location) for HR lifecycle events */
    const emitMaterialSignal = useCallback((collectionName, record) => {
        const signal = buildMaterialStaffSignals(collectionName, record);
        if (!signal) return false;
        setAllData(prev => ({
            ...prev,
            HR_MATERIAL_SIGNALS: [signal, ...(Array.isArray(prev.HR_MATERIAL_SIGNALS) ? prev.HR_MATERIAL_SIGNALS : [])],
        }));
        setMessage({ text: `Material signal: ${signal.signal}`, type: 'success' });
        return true;
    }, []);

    const getIntegrationMeta = useCallback((collectionName) => HR_INTEGRATION_REGISTRY[collectionName] || null, []);

    const dataValue = useMemo(() => ({
        allData, getCollection, addRecord, updateRecord, deleteRecord, message, setMessage, notificationTrigger, setNotificationTrigger, leaveUpdateTrigger, setLeaveUpdateTrigger, updateEmployeeRole, addRole, updateRole, deleteRole, roleUpdateTrigger, postHRToFinance, postAllHRToFinance, emitMaterialSignal, getIntegrationMeta, buildHRJournalLegs, HR_INTEGRATION_REGISTRY, HR_MATERIAL_OUTLETS
    }), [allData, getCollection, addRecord, updateRecord, deleteRecord, message, notificationTrigger, leaveUpdateTrigger, updateEmployeeRole, addRole, updateRole, deleteRole, roleUpdateTrigger, postHRToFinance, postAllHRToFinance, emitMaterialSignal, getIntegrationMeta]);
    return (
        <DataContext.Provider value={dataValue}>
            {children}
        </DataContext.Provider>
    );

    // Publish HR Employee Master for Finance payroll (inline integration)
    useEffect(() => {
        try {
            const list = allData?.HR_EMPLOYEE_CRUD;
            if (Array.isArray(list) && list.length) uuPublishHrEmployees(list);
        } catch (e) {}
    }, [allData?.HR_EMPLOYEE_CRUD]);

    useEffect(() => {
        try {
            const leave = allData?.HR_LEAVE_CRUD || [];
            localStorage.setItem('uu_shared_hr_leave_v1', JSON.stringify(leave));
        } catch (e) {}
    }, [allData?.HR_LEAVE_CRUD]);


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
    const selectRootRef = useRef(null);
    useEffect(() => {
        if (!selectOpen) return undefined;
        const onDoc = (e) => {
            if (selectRootRef.current && !selectRootRef.current.contains(e.target)) setSelectOpen(false);
        };
        document.addEventListener('mousedown', onDoc);
        document.addEventListener('touchstart', onDoc);
        return () => {
            document.removeEventListener('mousedown', onDoc);
            document.removeEventListener('touchstart', onDoc);
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
                {selectOpen && (
                    <div
                        className="absolute left-0 right-0 top-full mt-1 z-[10050]
                                   max-h-56 overflow-y-auto rounded-xl
                                   border border-white/20 bg-slate-900 shadow-2xl shadow-black/60"
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
                    </div>
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

    // New constants for Report Type
    const REPORT_TYPES = [
        { label: 'All Time', value: 'all' },
        { label: 'Daily', value: 'daily' },
        { label: 'Weekly', value: 'weekly' },
        { label: 'Monthly', value: 'monthly' },
        { label: 'Quarterly', value: 'quarterly' },
        { label: 'Annually', value: 'annually' }
    ];

    const getCollectionLabel = useCallback((name) => {
        if (!name) return '';
        if (typeof name !== 'string') return ''; 
        return name.replace(/_/g, ' ').replace('CRUD', '(Data)').trim();
    }, []);

    const collectionNames = useMemo(() => Object.keys(allData).sort(), [allData]);
    const initialSelectedCollectionName = collectionNames.length > 0 ? collectionNames[0] : '';
    
    const [selectedCollection, setSelectedCollection] = useState(initialSelectedCollectionName);
    const [sortKey, setSortKey] = useState('');
    const [groupByKey, setGroupByKey] = useState('');
    const [reportType, setReportType] = useState('all'); // New State for Type

    const columns = useMemo(() => COLUMN_CONFIG[selectedCollection] || [], [selectedCollection]);

    const filterableColumns = useMemo(() =>
        columns.filter(col => !col.isHidden && col.type !== 'array'),
        [columns]
    );

    const [columnFilters, setColumnFilters] = useState(() => {
        return filterableColumns.reduce((acc, col) => {
            acc[col.key] = '';  
            return acc;
        }, {});
    });

    const currentData = useMemo(() => getCollection(selectedCollection) || [], [getCollection, selectedCollection]);

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
            setReportType('all'); // Reset type on collection change
        }
    }, [selectedCollection, filterableColumns]);

    const getColumnFilterOptions = useCallback((data, columnKey) => {
        const values = new Set(['All']);
        data.forEach(item => {
            const value = item[columnKey];
            if (value !== null && value !== undefined) {
                values.add(String(value));
            }
        });
        return Array.from(values).map(v => ({ label: v, value: v }));
    }, []);

    const columnOptions = React.useMemo(() => {
        const options = {};
        filterableColumns.forEach(col => {
            options[col.key] = getColumnFilterOptions(currentData, col.key);
        });
        return options;
    }, [currentData, filterableColumns, getColumnFilterOptions]);

    const filteredAndSortedData = useMemo(() => {
        let data = [...currentData];

        data = data.filter(item => {
            return filterableColumns.every(col => {
                const filterValue = columnFilters[col.key] || '';
                if (filterValue === 'All' || !filterValue) {
                    return true;
                }
                const itemValue = String(item[col.key] || '').toLowerCase();
                const filterValueLower = filterValue.toLowerCase();
                return itemValue === filterValueLower;
            });
        });

        if (sortKey) {
            data.sort((a, b) => {
                const aVal = a[sortKey];
                const bVal = b[sortKey];
                if ((typeof aVal === 'number' || (!isNaN(parseFloat(aVal)) && isFinite(aVal))) && (typeof bVal === 'number' || (!isNaN(parseFloat(bVal)) && isFinite(bVal)))) {
                    return parseFloat(aVal) - parseFloat(bVal);
                }
                const stringA = String(aVal || '').toLowerCase();
                const stringB = String(bVal || '').toLowerCase();
                if (stringA < stringB) return -1;
                if (stringA > stringB) return 1;
                return 0;
            });
        }

        return data;
    }, [currentData, columnFilters, sortKey, filterableColumns]);

    const groupedData = useMemo(() => {
        if (!groupByKey) {
            return { 'All Records': filteredAndSortedData };
        }
        return filteredAndSortedData.reduce((acc, item) => {
            const key = String(item[groupByKey] || 'N/A');
            acc[key] = acc[key] || [];
            acc[key].push(item);
            return acc;
        }, {});
    }, [filteredAndSortedData, groupByKey]);

    const displayColumns = useMemo(() => columns.filter(col => !col.isHidden && col.type !== 'array').slice(0, 8), [columns]);

    const handleFilterChange = (columnKey, value) => {
        setColumnFilters(prevFilters => ({
            ...prevFilters,
            [columnKey]: value,
        }));
    };

    useEffect(() => {
        if (!selectedCollection && collectionNames.length > 0) {
            setSelectedCollection(collectionNames[0]);
        }
    }, [collectionNames, selectedCollection]);

    return (
        <ContainerCard  Icon='FileText' glassClass={glassClass} className="h-full">
            <div className="space-y-4 mb-6 p-4 border border-gray-700/50 rounded-xl">
                <h3 className="text-lg font-semibold text-cyan-400 flex items-center">
                    <Icons.Building className="w-5 h-5 mr-2" />
                    Report Configuration
                </h3>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                    <div className="col-span-full">
                        <div className="flex flex-wrap gap-4 justify-start items-end">
                            
                            {/* Data Collection Select */}
                            <div className="flex-grow min-w-[200px]">
                                <Input
                                    label="Data Collection"
                                    type="select"
                                    options={collectionNames.map(name => ({ label: getCollectionLabel(name), value: name }))}
                                    value={selectedCollection}
                                    onChange={(val) => setSelectedCollection(val)}
                                    isRequired
                                />
                            </div>

                            {/* Dynamic Filter Selects */}
                            {selectedCollection && filterableColumns.map(col => (
                                <div key={col.key} className="flex-grow min-w-[200px]">
                                    <Input
                                        label={`Filter: ${col.label}`}
                                        type="select"
                                        options={columnOptions[col.key] || [{ label: 'All', value: 'All' }]}
                                        value={columnFilters[col.key] || ''}
                                        onChange={(val) => handleFilterChange(col.key, val)}
                                    />
                                </div>
                            ))}

                            {/* Report Type Select (DAILY, WEEKLY, etc) */}
                            <div className="flex-grow min-w-[200px]">
                                <Input
                                    label="Report Type"
                                    type="select"
                                    options={REPORT_TYPES}
                                    value={reportType}
                                    onChange={(val) => setReportType(val)}
                                />
                            </div>

                            {/* Sort By Select */}
                            <div className="flex-grow min-w-[200px]">
                                <Input
                                    label="Sort By"
                                    type="select"
                                    options={[{ label: 'None', value: '' }, ...filterableColumns.map(c => ({ label: c.label, value: c.key }))]}
                                    value={sortKey}
                                    onChange={(val) => setSortKey(val)}
                                />
                            </div>

                            {/* Group By Select */}
                            <div className="flex-grow min-w-[200px]">
                                <Input
                                    label="Group By"
                                    type="select"
                                    options={[{ label: 'None', value: '' }, ...filterableColumns.map(c => ({ label: c.label, value: c.key }))]}
                                    value={groupByKey}
                                    onChange={(val) => setGroupByKey(val)}
                                />
                            </div>
                        </div>
                    </div>

                    {/* Export Button Integration */}
                    <div className="col-span-full pt-4 flex justify-end">
                        <ExportButton
                            data={filteredAndSortedData}
                            columns={columns}
                            filename={`${getCollectionLabel(selectedCollection)}-${reportType}-Report`}
                        />
                    </div>
                </div>
            </div>

            {/* Results Table */}
            {selectedCollection && (
                <div className="flex-grow overflow-y-auto p-2 mt-4">
                    <h3 className="text-xl font-bold text-white mb-4">
                        Report Results: {getCollectionLabel(selectedCollection)} ({filteredAndSortedData.length} Records)
                    </h3>

                    {Object.entries(groupedData).map(([groupKey, data]) => (
                        <div key={groupKey} className="mb-6 p-4 rounded-lg bg-gray-800/50">
                            <h4 className="text-lg font-semibold text-gray-300 sticky top-0 bg-gray-800/80 p-2 -m-2 mb-2 z-10">
                                {groupByKey ? `${filterableColumns.find(col => col.key === groupByKey)?.label || groupByKey}: ${groupKey} (${data.length})` : `All Records (${data.length})`}
                            </h4>
                            <div className="uu-table-scroll overflow-x-auto">
                                <table className="min-w-full divide-y divide-gray-700">
                                    <thead className="bg-gray-900/90 sticky top-12 z-[9]">
                                        <tr className="text-gray-400">
                                            {displayColumns.map((col, i) => (
                                                <th key={i} className="px-4 py-3 text-left text-xs font-medium uppercase tracking-wider">{col.label}</th>
                                            ))}
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-800 text-gray-300">
                                        {data.map((record) => (
                                            <tr key={record.id} className="transition duration-150 hover:bg-gray-700/50">
                                                {displayColumns.map((col, j) => (
                                                    <td key={j} className="px-4 py-3 whitespace-nowrap text-sm">
                                                        {Array.isArray(record[col.key])
                                                            ? record[col.key].slice(0, 2).map((entry, idx) => (
                                                                <span key={idx} className="bg-gray-700/50 text-gray-300 rounded-full px-2 py-1 text-xs mr-1">
                                                                    {typeof entry === 'object' ? Object.values(entry).join(' / ') : String(entry)}
                                                                </span>
                                                            ))
                                                            : String(record[col.key] || 'N/A')}
                                                    </td>
                                                ))}
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    ))}
                    {filteredAndSortedData.length === 0 && (
                        <p className="text-center py-8 text-gray-500">No records match the current criteria.</p>
                    )}
                </div>
            )}
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
                <button
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
const Modal = ({ isOpen, onClose, title, children, glassClass = '', footer }) => {
    if (!isOpen) return null;

    return (
        <div
            className="fixed inset-0 z-[9999] flex items-end sm:items-center justify-center
                       p-0 sm:p-5 bg-slate-950/50 backdrop-blur-md"
            role="presentation"
        >
            <div
                className={`relative z-[10000] w-full sm:max-w-2xl max-h-[92dvh] flex flex-col
                            overflow-hidden rounded-t-2xl sm:rounded-2xl
                            border-2 border-white/25 bg-slate-900/95 backdrop-blur-xl text-white
                            shadow-[0_25px_80px_rgba(0,0,0,0.65)] ring-1 ring-cyan-400/25
                            ${glassClass || ''}`}
                role="dialog"
                aria-modal="true"
                onMouseDown={(e) => e.stopPropagation()}
                onClick={(e) => e.stopPropagation()}
            >
                <div className="h-1 w-full shrink-0 bg-gradient-to-r from-cyan-500 via-sky-400 to-violet-500" />

                <div className="shrink-0 flex items-center justify-between gap-3 px-4 py-3 border-b border-white/15 bg-white/5">
                    <h3 className="text-sm sm:text-lg font-black uppercase tracking-wide text-cyan-300 truncate">
                        {title}
                    </h3>
                    <button
                        type="button"
                        onClick={onClose}
                        className="p-2.5 rounded-xl border border-white/20 bg-white/5 min-h-[44px] min-w-[44px]
                                   flex items-center justify-center text-gray-300 hover:text-red-300"
                        aria-label="Close"
                    >
                        <Icons.X className="w-5 h-5" />
                    </button>
                </div>

                <div className="flex-1 min-h-0 overflow-y-auto overflow-x-hidden px-4 py-4 overscroll-contain">
                    {children}
                </div>

                <div className="shrink-0 flex flex-wrap justify-end gap-2 px-4 py-3 border-t border-white/15 bg-slate-900/95
                                pb-[max(0.75rem,env(safe-area-inset-bottom))]">
                    {footer}
                    <button
                        type="button"
                        onClick={onClose}
                        className="px-4 py-3 rounded-xl text-xs font-bold uppercase border border-white/20
                                   bg-white/5 text-gray-200 min-h-[44px] flex-1 sm:flex-none"
                    >
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );
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
                saveAs(new Blob(['\ufeff' + lines.join('\n')], { type: 'text/csv;charset=utf-8' }), `${fname}_${new Date().toISOString().slice(0, 10)}.csv`);
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
    const { deleteRecord, updateRecord, postHRToFinance, postAllHRToFinance, emitMaterialSignal, getIntegrationMeta } = useData();
    const { checkPermission } = useAuth();
    const [isModalOpen, setIsModalOpen] = useState(false);
    const [editingRecord, setEditingRecord] = useState(null);
    const [searchTerm, setSearchTerm] = useState('');
    const [statusFilter, setStatusFilter] = useState('ALL');
    const [deptFilter, setDeptFilter] = useState('ALL');
    const tableWrapRef = useRef(null);

    const allRoles = useMemo(() => (getCollection ? getCollection('ROLE_CRUD') : []) || [], [getCollection]);
    const liveData = useMemo(() => {
        if (getCollection && collectionName) return getCollection(collectionName) || [];
        return Array.isArray(data) ? data : [];
    }, [getCollection, collectionName, data]);

    const getAddLabel = () => {
        const labels = {
            'HR_EMPLOYEE_CRUD': 'Employee',
            'HR_LEAVE_CRUD': 'Leave Request',
            'HR_VACANCY_CRUD': 'Vacancy',
            'HR_RECRUITMENT_CRUD': 'Candidate',
            'HR_APPRAISAL_CRUD': 'Appraisal',
            'HR_COMPETENCE_CRUD': 'Competence',
            'HR_PENALTY_CRUD': 'Penalty',
            'HR_AWARD_CRUD': 'Award',
            'HR_AWARD_CRUD_HR': 'Award',
            'HR_TRANSFER_CRUD': 'Transfer',
            'HR_TRAINING_CRUD': 'Training',
            'HR_INSURANCE_CRUD': 'Insurance Policy',
            'DEPARTMENTS_CRUD': 'Department',
            'DEPARTMENTS_HIRING_REQUESTS_CRUD': 'Hiring Request',
            'HR_STAFF_ASSIGNMENT_CRUD': 'Staff Assignment',
            'HR_STAFF_ATTRITION_CRUD': 'Attrition Record',
            'HR_STAFF_DEVELOPMENT_CRUD': 'Development Plan',
            'HR_STAFF_OVERVIEW_CRUD': 'Staff Overview',
            'RETIREMENT_CRUD': 'Retirement',
            'RESIGNATION_CRUD': 'Resignation',
            'DEATH_CRUD': 'Death Record',
            'ROLE_CRUD': 'Role',
        };
        return labels[collectionName] || (title || 'Record');
    };

    const moduleTitle = title || getAddLabel() + ' Register';

    const visibleColumns = useMemo(
        () => (columns || []).filter(c => !c.isHidden && c.type !== 'array' && c.key !== 'actions'),
        [columns]
    );

    const statusKey = useMemo(() => {
        const keys = ['status', 'auth_status', 'employment_type', 'leave_status'];
        return visibleColumns.find(c => keys.includes(c.key))?.key || null;
    }, [visibleColumns]);

    const deptKey = useMemo(() => {
        const keys = ['department', 'department_name', 'work_location'];
        return visibleColumns.find(c => keys.includes(c.key))?.key || null;
    }, [visibleColumns]);

    const statusOptions = useMemo(() => {
        if (!statusKey) return ['ALL'];
        const s = new Set(liveData.map(r => String(r[statusKey] || '').trim()).filter(Boolean));
        return ['ALL', ...Array.from(s).sort()];
    }, [liveData, statusKey]);

    const deptOptions = useMemo(() => {
        if (!deptKey) return ['ALL'];
        const s = new Set(liveData.map(r => String(r[deptKey] || '').trim()).filter(Boolean));
        return ['ALL', ...Array.from(s).sort()];
    }, [liveData, deptKey]);

    const filteredData = useMemo(() => {
        const q = searchTerm.trim().toLowerCase();
        return liveData.filter(row => {
            if (statusKey && statusFilter !== 'ALL' && String(row[statusKey] || '') !== statusFilter) return false;
            if (deptKey && deptFilter !== 'ALL' && String(row[deptKey] || '') !== deptFilter) return false;
            if (!q) return true;
            return visibleColumns.some(col => {
                const v = row[col.key];
                if (v == null) return false;
                if (Array.isArray(v)) return JSON.stringify(v).toLowerCase().includes(q);
                return String(v).toLowerCase().includes(q);
            });
        });
    }, [liveData, searchTerm, statusFilter, deptFilter, statusKey, deptKey, visibleColumns]);

    const cards = useMemo(() => {
        const rows = filteredData;
        const total = rows.length;
        const male = rows.filter(r => /^(m|male)$/i.test(String(r.gender || ''))).length;
        const female = rows.filter(r => /^(f|female)$/i.test(String(r.gender || ''))).length;
        const pending = rows.filter(r => /pending|draft|open/i.test(String(r.status || r.auth_status || ''))).length;
        const approved = rows.filter(r => /approv|active|published|completed/i.test(String(r.status || r.auth_status || ''))).length;
        const depts = new Set(rows.map(r => r.department || r.department_name).filter(Boolean)).size;

        if (collectionName === 'HR_EMPLOYEE_CRUD') {
            return [
                { label: 'Employees', value: total, color: 'text-cyan-400', tip: 'Records in view' },
                { label: 'Male', value: male, color: 'text-sky-400', tip: 'Gender M' },
                { label: 'Female', value: female, color: 'text-pink-400', tip: 'Gender F' },
                { label: 'Departments', value: depts, color: 'text-emerald-400', tip: 'Distinct units' },
            ];
        }
        if (collectionName === 'HR_LEAVE_CRUD') {
            return [
                { label: 'Leave requests', value: total, color: 'text-cyan-400', tip: 'In filtered set' },
                { label: 'Pending', value: pending, color: 'text-amber-400', tip: 'Awaiting decision' },
                { label: 'Approved', value: approved, color: 'text-emerald-400', tip: 'Approved leaves' },
                { label: 'Types', value: new Set(rows.map(r => r.leave_type || r.type).filter(Boolean)).size, color: 'text-violet-400', tip: 'Leave categories' },
            ];
        }
        if (collectionName === 'HR_VACANCY_CRUD' || collectionName === 'DEPARTMENTS_HIRING_REQUESTS_CRUD') {
            return [
                { label: 'Openings', value: total, color: 'text-cyan-400', tip: 'Vacancy / request lines' },
                { label: 'Open / Published', value: rows.filter(r => /open|publish/i.test(String(r.status || ''))).length, color: 'text-emerald-400', tip: 'Active hiring' },
                { label: 'Closed / Hold', value: rows.filter(r => /close|hold|draft/i.test(String(r.status || ''))).length, color: 'text-amber-400', tip: 'Not active' },
                { label: 'Departments', value: depts, color: 'text-sky-400', tip: 'Hiring units' },
            ];
        }
        return [
            { label: 'Records', value: total, color: 'text-cyan-400', tip: 'Matching filters' },
            { label: 'Pending', value: pending, color: 'text-amber-400', tip: 'Pending / draft' },
            { label: 'Approved / Active', value: approved, color: 'text-emerald-400', tip: 'Cleared items' },
            { label: 'Departments', value: depts || '—', color: 'text-violet-400', tip: 'Distinct departments' },
        ];
    }, [filteredData, collectionName]);

    const handleAdd = () => { setEditingRecord({}); setIsModalOpen(true); };
    const handleEdit = (item) => { setEditingRecord(item); setIsModalOpen(true); };

    const exportFilename = String(moduleTitle).replace(/[^\w\s\-]/g, '').replace(/\s+/g, '_');

    return (
        <div className={`rounded-[1.75rem] border border-white/10 ${glassClass} flex flex-col w-full min-h-0`}>
            <div className="shrink-0 z-30 border-b border-white/10 bg-slate-950/95 backdrop-blur-xl px-4 sm:px-5 pt-4 pb-3">
                <div className="border-l-4 border-cyan-500 pl-4 mb-3">
                    <h1 className="text-xl sm:text-2xl md:text-3xl font-black text-white uppercase italic tracking-tighter leading-tight">
                        {moduleTitle}
                    </h1>
                    <p className="text-[10px] text-cyan-400 font-black uppercase tracking-[0.2em] mt-1 opacity-90">
                        Unity University · Human Resources · Official register
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
                            />
                        )}
                        {deptOptions.length > 1 && (
                            <UUSelect
                                value={deptFilter}
                                onChange={setDeptFilter}
                                options={deptOptions}
                                allLabel="All departments"
                                compact
                            />
                        )}
                        {(searchTerm || statusFilter !== 'ALL' || deptFilter !== 'ALL') && (
                            <button type="button" onClick={() => { setSearchTerm(''); setStatusFilter('ALL'); setDeptFilter('ALL'); }}
                                className="text-[10px] font-black uppercase text-rose-400 hover:text-rose-300 px-2 whitespace-nowrap">
                                Clear filters
                            </button>
                        )}
                    </div>
                    <div className="flex flex-wrap gap-2 items-center shrink-0">
                        {HR_COLLECTION_SOURCE_MAP[collectionName] && (
                            <button
                                type="button"
                                onClick={() => postAllHRToFinance?.(collectionName)}
                                className="bg-emerald-700 hover:bg-emerald-600 text-white font-bold py-2 px-3 rounded-xl flex items-center gap-2 text-[10px] uppercase tracking-widest border border-white/20"
                                title="Post unposted HR docs to Finance"
                            >
                                <Icons.Landmark size={14} /> Post All → Finance
                            </button>
                        )}
                        {(collectionName === 'HR_TRANSFER_CRUD' || collectionName === 'RESIGNATION_CRUD' || collectionName === 'RETIREMENT_CRUD' || collectionName === 'DEATH_CRUD') && (
                            <button
                                type="button"
                                onClick={() => {
                                    filteredData.forEach(r => emitMaterialSignal?.(collectionName, r));
                                }}
                                className="bg-amber-700 hover:bg-amber-600 text-white font-bold py-2 px-3 rounded-xl flex items-center gap-2 text-[10px] uppercase tracking-widest border border-white/20"
                                title="Notify Material module (custody / location)"
                            >
                                <Icons.Package size={14} /> Signal Material
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

            <div ref={tableWrapRef} className="w-full overflow-x-auto overflow-y-visible custom-scrollbar relative border border-white/15 rounded-xl">
                <table className="w-max min-w-full border-collapse text-left border border-white/15">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-slate-900 shadow-md shadow-black/40">
                            {visibleColumns.map(col => (
                                <th key={col.key}
                                    className="px-3 py-2.5 text-[9px] font-black uppercase tracking-widest text-cyan-300/95 border border-white/15 border-b-cyan-500/40 whitespace-nowrap bg-slate-900"
                                    style={{ minWidth: col.type === 'textarea' ? 160 : (col.type === 'number' ? 100 : 130), maxWidth: 260 }}>
                                    {col.label}
                                </th>
                            ))}
                            <th className="px-3 py-2.5 text-[9px] font-black uppercase tracking-widest text-cyan-300/95 border-b border-cyan-500/30 whitespace-nowrap bg-slate-900 sticky right-0 z-30 shadow-[-8px_0_12px_rgba(0,0,0,0.35)]" style={{ minWidth: 100 }}>
                                Actions
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        {filteredData.length > 0 ? filteredData.map(item => (
                            <tr key={item.id} onClick={() => handleEdit(item)}
                                className="border-b border-white/15 hover:bg-cyan-500/10 cursor-pointer transition-colors group"
                                title="Click row to open record">
                                {visibleColumns.map(col => (
                                    <td key={col.key} className="px-3 py-2 text-[11px] text-gray-200 border border-white/10 group-hover:text-white"
                                        style={{ minWidth: col.type === 'textarea' ? 160 : (col.type === 'number' ? 100 : 130), maxWidth: 260 }}>
                                        <div className="truncate" title={item[col.key] != null && item[col.key] !== '' ? String(item[col.key]) : ''}>
                                            {item[col.key] !== undefined && item[col.key] !== null && item[col.key] !== ''
                                                ? (Array.isArray(item[col.key]) ? `${item[col.key].length} entries` : String(item[col.key]))
                                                : <span className="text-white/15 italic">—</span>}
                                        </div>
                                    </td>
                                ))}
                                <td className="px-2 py-2 bg-slate-950/80 sticky right-0 z-10 shadow-[-8px_0_12px_rgba(0,0,0,0.25)] group-hover:bg-slate-900/95"
                                    onClick={e => e.stopPropagation()} style={{ minWidth: 100 }}>
                                    <div className="flex justify-center items-center gap-1.5">
                                        {HR_COLLECTION_SOURCE_MAP[collectionName] && !item.posted_to_finance && (
                                            <button type="button"
                                                onClick={() => postHRToFinance?.(HR_COLLECTION_SOURCE_MAP[collectionName], item.id)}
                                                className="text-emerald-400 hover:text-emerald-200 p-1" title="Post to Finance">
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
                        )) : (
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
   
   'HR_EMPLOYEE_CRUD': [
    { key: 'id_number', label: 'ID NO ', isRequired: true, isReadOnly: true },
    { key: 'name', label: 'Full Name', isRequired: true },
    { key: 'gender', label: 'Gender', type: 'select', options: ['Male', 'Female'] },
    { key: 'dob_gc', label: 'Date of Birth (G.C.)', type: 'date' },
    { key: 'dob_ec', label: 'Date of Birth (E.C.)' }, 
    { key: 'age', label: 'Age', type: 'number' },
    { key: 'marital_status', label: 'Marital Status', type: 'select', options: ['Single', 'Married', 'Divorced', 'Widowed'] },
    { key: 'seniority_date_gc', label: 'Seniority Date (G.C.)', type: 'date' },
    { key: 'seniority_date_ec', label: 'Seniority Date (E.C.)' },
    { key: 'work_location', label: 'Work Location' },
    { key: 'department', label: 'Department', type: 'select' },
    { key: 'job_title', label: 'Current Position / Academic Rank', isRequired: true },
    { key: 'reporting_to', label: 'Reporting To', type: 'select' },
    { key: 'employment_category', label: 'Employment Category', type: 'select', options: ['Academic', 'Non-Academic', 'Semi-Academic'] },
    { key: 'employment_type', label: 'Employment Type', type: 'select', options: ['Permanent', 'Non-Permanent'] },
    { key: 'educational_status', label: 'Educational Status / Experience Year (Click label to Jump)' },
    { key: 'cumulative_gpa', label: 'Cumulative GPA' },
    { key: 'auth_status', label: 'Authentication of Educational Level', type: 'select', options: ['Approved', 'Not Approved'] },
    { key: 'field_of_study', label: 'Field of Study' },
    { key: 'gross_salary', label: 'Current Salary', type: 'number' },
    { key: 'family_contact', label: 'Family Contact' },
    { key: 'pension_id', label: 'Pension ID Number' },
    { key: 'retirement_date_ec', label: 'Retirement Date (E.C.)' },
    { key: 'national_id', label: 'National ID NO.' },
    { key: 'tin_no', label: 'TIN NO.' },
    { key: 'telephone_number', label: 'Mobile Tel. No' },
    { key: 'office_tel', label: 'Office Tel. No' },
    { key: 'ext_no', label: 'Ext. No.' },
    { key: 'emergency_contact_name', label: 'Emergency Contact Name' },
    { key: 'emergency_contact_telephone', label: 'Emergency Contact Phone' },
    { key: 'emergency_contact_address', label: 'Emergency Contact Address', type: 'textarea' },
    { 
        key: 'education_entries', 
        label: 'Education Entries (Experience History)', 
        type: 'array', 
        arrayConfig: [
            { key: 'title', label: 'Degree/Title' }, 
            { key: 'institution', label: 'Institution' }, 
            { key: 'year', label: 'Year' }
        ] 
    },
    { key: 'remarks', label: 'Remarks', type: 'textarea' },
    { 
        key: 'beneficiaries', 
        label: 'Beneficiaries/Family Contacts', 
        type: 'array', 
        arrayConfig: [
            { key: 'name', label: 'Name' }, 
            { key: 'relationship', label: 'Relationship' }, 
            { key: 'percentage', label: 'Percentage' }
        ] 
    },
    { key: 'system_role', label: 'System Access Role', isReadOnly: true }
],
     // Organization Structure Modules
    'DEPARTMENTS_CRUD': [
        { key: 'id', label: 'Department ID', isHidden: true, isReadOnly: true },
        { key: 'name', label: 'Department Name', isRequired: true },
        { key: 'head', label: 'Head of Department', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name), isRequired: true },
        { key: 'location', label: 'Location', type: 'select', options: ['Gerji', 'Burayu', 'Dessie', 'Adama'] },
        { key: 'employees_count', label: 'Number of Employees', type: 'number', isReadOnly: true },
    ],
    'DEPARTMENTS_HIRING_REQUESTS_CRUD': [
        { key: 'id', label: 'Request ID', isHidden: true, isReadOnly: true },
        { key: 'department_id', label: 'Department ID', isHidden: true }, // Primarily for internal linking
        { key: 'department_name', label: 'Department', type: 'select', options: MOCK_DEPARTMENTS.map(d => d.name), isRequired: true },
        { key: 'position_title', label: 'Position Title', isRequired: true },
        { key: 'required_skills', label: 'Required Skils', type: 'select', options: MOCK_SKILLS, isRequired: true, isMulti: true },
        { key: 'no_of_vacancies', label: 'Number', type: 'number', isRequired: true },
        { key: 'status', label: 'Status', type: 'select', options: ['Open', 'Closed', 'On Hold'], isRequired: true },
        { key: 'request_date', label: 'Request Date', type: 'date', isRequired: true },
        { key: 'requested_by', label: 'Requested By', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name), isReadOnly: true }, // Should be auto-populated or selected from employees
        { key: 'hiring_manager', label: 'Hiring Manager', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name), isRequired: true },
    ],
'HR_VACANCY_CRUD': [
        { key: 'id', label: 'Vacancy ID', isHidden: true, isReadOnly: true },
        { key: 'vacancy_no', label: 'Vacancy No', isRequired: true },
        { key: 'position_title', label: 'Job Position', isRequired: true },
        { key: 'department', label: 'Department', type: 'select' },
        { key: 'no_of_positions', label: 'No. of Openings', type: 'number' },
        { key: 'employment_type', label: 'Type', type: 'select', options: ['Full-Time', 'Part-Time', 'Contract'] },
        { key: 'closing_date', label: 'Closing Date', type: 'date' },
        { key: 'status', label: 'Status', type: 'select', options: ['Draft', 'Published', 'Closed'] },
        { key: 'description', label: 'Job Description', type: 'textarea' }
    ],
    'RETIREMENT_CRUD': [
    { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
    { 
        key: 'employee_name', 
        label: 'Employee Name', 
        type: 'select', 
        isRequired: true, 
        // Ensure onChange calls handleEmployeeSelect in your Modal component
        options: [] // Dynamically populated
    },
    { 
        key: 'retirement_date', 
        label: 'Retirement Date', 
        type: 'date', 
        isRequired: true 
    },
    // { key: 'reason', label: 'Reason' },
    { key: 'pension_details', label: 'Pension Details' },
    { key: 'notes', label: 'Notes', type: 'textarea' },
],
    'RESIGNATION_CRUD': [
        { key: 'id', label: 'Resignation ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name), isRequired: true },
        { key: 'resignation_date', label: 'Date of Resignation', type: 'date', isRequired: true },
        { key: 'last_working_day', label: 'Last Working Day', type: 'date' },
        { key: 'reason', label: 'Reason for Resignation', type: 'textarea', isRequired: true },
        { key: 'acceptance_status', label: 'Acceptance Status', type: 'select', options: ['Accepted', 'Pending', 'Rejected'], isRequired: true },
        { key: 'handover_completed', label: 'Handover Completed?', type: 'select', options: ['Yes', 'No'], isRequired: true },
    ],
    'DEATH_CRUD': [
        { key: 'id', label: 'Death ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name), isRequired: true },
        { key: 'date_of_death', label: 'Date of Death', type: 'date', isRequired: true },
        { key: 'cause_of_death', label: 'Cause of Death' },
        { key: 'reported_by', label: 'Reported By', type: 'select', options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'legal_documentation_status', label: 'Legal Documentation Status', type: 'select', options: ['Provided', 'Pending', 'Not Applicable'] },
    ],

 'HR_RECRUITMENT_CRUD': [
    { key: 'form_no', label: 'Protocol ID', placeholder: 'Type existing ID to auto-populate' },
    { key: 'report_date', label: 'Reporting Date', type: 'date', isRequired: true },
    
    // Explicit list for Department dropdown
    { 
        key: 'department', 
        label: 'Target Department', 
        type: 'select', 
        options: ['Human Resource', 'Finance', 'Engineering', 'Marketing', 'Legal', 'IT Support', 'Procurement', 'Academic Affairs'] 
    },
    
    // Explicit list for Position dropdown
    { 
        key: 'position', 
        label: 'Vacant Job Title', 
        type: 'select',
        options: ['Senior Manager', 'Junior Officer', 'Technical Specialist', 'Lecturer', 'Administrative Assistant', 'Department Head']
    },
    
    { key: 'recruit_name', label: 'Applicant Full Name', type: 'text', isRequired: true },
    
    // COMMITTEE EVALUATION
    { 
        key: 'committee_evaluations', 
        label: 'Consolidated Member Scoring & Authentication', 
        type: 'array', 
        arrayConfig: [
            { key: 'member_name', label: 'Full Name of the Committee Member' },
            { key: 'member_position', label: 'Position / Institutional Rank' },
            { key: 'edu', label: 'Education (25%)', type: 'number' },
            { key: 'exp', label: 'Experience (25%)', type: 'number' },
            { key: 'kno', label: 'Subject Matter Knowledge (20%)', type: 'number' },
            { key: 'com', label: 'Professional Communication (10%)', type: 'number' },
            { key: 'per', label: 'Personality Assessment (10%)', type: 'number' },
            { key: 'spe', label: 'Special Technical Skill (10%)', type: 'number' },
            { key: 'digital_stamp', label: 'Member Digital Validation Key', isReadOnly: true }
        ] 
    },

    {
        key: 'missing_documents',
        label: '⚠️ Missing Document Reminder (Follow-up List)',
        type: 'array',
        arrayConfig: [
            { key: 'doc_name', label: 'Missing Document Name' },
            { key: 'deadline', label: 'Submission Deadline', type: 'date' },
            { key: 'status', label: 'Current Status', type: 'select', options: ['Pending', 'Submitted', 'Partially Received'] }
        ]
    },

    { key: 'rank', label: 'Final Computed Rank', type: 'number' },
    { key: 'recommendation', label: 'Final Recommendation to President Office', type: 'textarea' },
    
    { key: 'approved_by', label: 'Authorizing Officer', placeholder: 'Name of Approver' },
    { key: 'approval_date', label: 'Approval Execution Date', type: 'date' },
    { key: 'approved_by_signature', label: 'Authority Master Stamp', isReadOnly: true, placeholder: 'MASTER_APPROVAL_PROTOCOL_SECURE' }
],   // Update this in your COLUMN_CONFIG object
'HR_COMPETENCE_CRUD': [
    { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
    { 
        key: 'skill', 
        label: 'Skill / Competency', 
        type: 'select', 
        isRequired: true, 
        // This ensures the dropdown in the "Add" form also has all skills
        options: Object.values(NEW_MOCK_SKILLS).flat() 
    },
    { key: 'score', label: 'Score (1-5)', type: 'select', options: COMPETENCE_LEVELS.map(String), isRequired: true },
    { key: 'last_assessed', label: 'Assessment Date', type: 'date', isRequired: true },
    { key: 'department', label: 'Department', isReadOnly: true },
    { key: 'evaluator', label: 'Evaluator', isReadOnly: true }
],
    'HR_APPRAISAL_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'appraisal_date', label: 'Appraisal Date', type: 'date', isRequired: true },
        { key: 'appraisal_score', label: 'Score (1.0 - 5.0)', type: 'number', isRequired: true },
        { key: 'feedback_manager', label: 'Manager Feedback', type: 'textarea' },
        { key: 'feedback_peer', label: 'Peer Feedback', type: 'textarea' },
        { key: 'feedback_self', label: 'Self Review', type: 'textarea' },
        { key: 'status', label: 'Status', type: 'select', options: ['Pending', 'Completed', 'Canceled'] },
    ],
    'HR_LEAVE_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'leave_type', label: 'Leave Type', type: 'select', isRequired: true, options: LEAVE_TYPES }, // Options dynamically filtered in form
        { key: 'start_date', label: 'Start Date', type: 'date', isRequired: true },
        { key: 'end_date', label: 'End Date', type: 'date', isRequired: true },
        { key: 'days_taken', label: 'Days Taken', type: 'number', isRequired: true, isReadOnly: true }, // Should be calculated
        { key: 'status', label: 'Status', type: 'select', options: ['Pending', 'Approved', 'Rejected'] },
    ],
    'HR_PENALTY_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'reason', label: 'Reason for Penalty', type: 'textarea', isRequired: true },
        { key: 'date', label: 'Date', type: 'date', isRequired: true },
        { key: 'action_taken', label: 'Action Taken', type: 'select', options: HR_PENALTIES, isRequired: true },
        { key: 'severity', label: 'Severity', type: 'select', options: ['Low', 'Medium', 'High'] },
    ],
    'HR_AWARD_CRUD_HR': [ // Same as HR_AWARD_CRUD but for the main menu
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'award_type', label: 'Award Type', isRequired: true },
        { key: 'date', label: 'Date Awarded', type: 'date', isRequired: true },
        { key: 'description', label: 'Description', type: 'textarea' },
        { key: 'monetary_value', label: 'Monetary Value (ETB)', type: 'number' },
    ],
  'HR_TRANSFER_CRUD': [
    { key: 'employee_name', label: 'Employee Name', isRequired: true },
    { key: 'current_department', label: 'Current Department', isReadOnly: true },
    { key: 'new_department', label: 'Destination Department', type: 'select', isRequired: true },
    { key: 'current_location', label: 'Current Location', isReadOnly: true },
    { key: 'new_location', label: 'Destination Location', type: 'select', isRequired: true },
    { key: 'transfer_reason', label: 'Transfer Reason', type: 'textarea' },
    { key: 'effective_date', label: 'Effective Date', type: 'date' }
],
    'HR_TRAINING_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'training_name', label: 'Training Name', isRequired: true },
        { key: 'date_completed', label: 'Date Completed', type: 'date', isRequired: true },
        { key: 'duration_days', label: 'Duration (Days)', type: 'number' },
        { key: 'status', label: 'Status', type: 'select', options: ['Completed', 'Scheduled', 'Canceled'] },
        { key: 'score', label: 'Score (%)', type: 'number' },
    ],
    'HR_INSURANCE_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'employee_name', label: 'Employee Name', type: 'select', isRequired: true, options: MOCK_EMPLOYEES.map(e => e.name) },
        { key: 'policy_no', label: 'Policy No', isRequired: true },
        { key: 'type', label: 'Type', type: 'select', options: ['Health', 'Life', 'Disability'] },
        { key: 'effective_date', label: 'Effective Date', type: 'date', isRequired: true },
        { key: 'premium_amount', label: 'Premium Amount (ETB)', type: 'number' },
        { key: 'status', label: 'Status', type: 'select', options: ['Active', 'Expired', 'Canceled'] },
    ],
    // 'HR_REPORT_CRUD': [
    //     { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
    //     { key: 'report_name', label: 'Report Name', isRequired: true },
    //     { key: 'report_type', label: 'Type', type: 'select', options: ['Monthly', 'Quarterly', 'Annual'] },
    //     { key: 'date_generated', label: 'Date Generated', type: 'date', isRequired: true },
    //     { key: 'department', label: 'Department'}, // options dynamically loaded },
    //     { key: 'data_points', label: 'Data Points', type: 'number' },
    //     { key: 'status', label: 'Status', type: 'select', options: ['Generated', 'Draft'] },
    // ],
    // Staff Modules (Data-intensive CRUDs)
    'HR_STAFF_ASSIGNMENT_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'institution_code', label: 'Institution Code', isRequired: true },
        { key: 'staff_national_id', label: 'Staff National ID', isRequired: true },
        { key: 'academic_year', label: 'Academic Year', type: 'select', options: ACADEMIC_YEARS, isRequired: true },
        { key: 'academic_period', label: 'Academic Period', type: 'select', options: ACADEMIC_PERIODS, isRequired: true },
        { key: 'campus', label: 'Campus' },
        { key: 'staff_category', label: 'Staff Category', type: 'select', options: ['ACD', 'ADM'] },
        { key: 'job_title', label: 'Job Title' },
        { key: 'academic_rank', label: 'Academic Rank' },
        { key: 'employment_date', label: 'Employment Date', type: 'date' },
        { key: 'employment_type', label: 'Employment Type', type: 'select', options: ['FT', 'PT'] },
        { key: 'monthly_salary', label: 'Monthly Salary (ETB)', type: 'number' },
        { key: 'isced_education_level', label: 'ISCED Education Level' },
        { key: 'is_under_qualified', label: 'Is Under-qualified?', type: 'select', options: ['true', 'false'] },
        { key: 'nbr_courses_taught', label: 'Courses Taught', type: 'number' },
        { key: 'nbr_credit_hours_taught', label: 'Credit Hours Taught', type: 'number' },
        { key: 'nbr_modules_taught', label: 'Modules Taught', type: 'number' },
        { key: 'nbr_postgraduates_advised', label: 'Postgraduates Advised', type: 'number' },
        { key: 'standard_teaching_load', label: 'Standard Teaching Load', type: 'number' },
        { key: 'community_service_load', label: 'Community Service Load', type: 'number' },
        { key:'research_load', label: 'Research Load', type: 'number' },
        { key: 'teaching_overload', label: 'Teaching Overload', type: 'number' },
        { key: 'overload_reason', label: 'Overload Reason' },
        { key: 'science_training', label: 'Science Training?', type: 'select', options: ['true', 'false'] },
        { key: 'research_training', label: 'Research Training?', type: 'select', options: ['true', 'false'] },
        { key: 'teaching_license', label: 'Teaching License?', type: 'select', options: ['true', 'false'] },
        { key: 'completed_cpd_hdp', label: 'Completed CPD/HDP?', type: 'select', options: ['true', 'false'] },
        { key: 'outgoing_exchange', label: 'Outgoing Exchange?', type: 'select', options: ['true', 'false'] },
        { key: 'exchange_country', label: 'Exchange Country', type: 'select', options: COUNTRIES },
        { key: 'industry_experience', label: 'Industry Experience?', type: 'select', options: ['true', 'false'] },
        { key: 'leadership_training', label: 'Leadership Training?', type: 'select', options: ['true', 'false'] },
        { key: 'merit_based', label: 'Merit Based?', type: 'select', options: ['true', 'false'] },
        { key: 'digital_skills_training', label: 'Digital Skills Training?', type: 'select', options: ['true', 'false'] },
        { key: 'professional_training', label: 'Professional Training?', type: 'select', options: ['true', 'false'] },
        { key: 'research_ethics_training', label: 'Research Ethics Training?', type: 'select', options: ['true', 'false'] },
        { key: 'digital_instruction_training', label: 'Digital Instruction Training?', type: 'select', options: ['true', 'false'] },
        { key: 'elip_training', label: 'ELIP Training?', type: 'select', options: ['true', 'false'] },
    ],
    'HR_STAFF_ATTRITION_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'institution_code', label: 'Institution Code', isRequired: true },
        { key: 'staff_national_id', label: 'Staff National ID', isRequired: true },
        { key: 'academic_year', label: 'Academic Year', type: 'select', options: ACADEMIC_YEARS, isRequired: true },
        { key: 'academic_period', label: 'Academic Period', type: 'select', options: ACADEMIC_PERIODS, isRequired: true },
        { key: 'attrition_cause', label: 'Attrition Cause Code', type: 'select', options: STAFF_ATTRITION_CAUSES, isRequired: true },
    ],
    'HR_AWARD_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'institution_code', label: 'Institution Code', isRequired: true },
        { key: 'staff_national_id', label: 'Staff National ID', isRequired: true },
        { key: 'academic_year', label: 'Academic Year', type: 'select', options: ACADEMIC_YEARS, isRequired: true },
        { key: 'academic_period', label: 'Academic Period', type: 'select', options: ACADEMIC_PERIODS, isRequired: true },
        { key: 'award_category', label: 'Award Category', type: 'select', options: AWARD_CATEGORIES, isRequired: true },
        { key: 'award_date', label: 'Award Date', type: 'date', isRequired: true },
        { key: 'award_amount', label: 'Award Amount (ETB)', type: 'number' },
    ],
    'HR_STAFF_DEVELOPMENT_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'institution_code', label: 'Institution Code', isRequired: true },
        { key: 'staff_national_id', label: 'Staff National ID', isRequired: true },
        { key: 'program_name', label: 'Program Name' },
        { key: 'program_country', label: 'Program Country', type: 'select', options: COUNTRIES, isRequired: true },
        { key: 'target_education_level', label: 'Target Education Level' },
        { key: 'special_program_type', label: 'Special Program Type', type: 'select', options: PROGRAM_TYPES },
        { key: 'program_status', label: 'Program Status', type: 'select', options: PROGRAM_STATUSES },
        { key: 'start_date', label: 'Start Date', type: 'date' },
        { key: 'end_date', label: 'End Date', type: 'date' },
        { key: 'nbr_participants', label: 'Number of Participants', type: 'number' },
        { key: 'program_cost', label: 'Program Cost (ETB)', type: 'number' },
    ],
    'HR_STAFF_OVERVIEW_CRUD': [
        { key: 'id', label: 'ID', isHidden: true, isReadOnly: true },
        { key: 'institution_code', label: 'Institution Code' },
        { key: 'staff_national_id', label: 'Staff National ID' },
        { key: 'name', label: 'Name' },
        { key: 'date_of_birth', label: 'Date of Birth', type: 'date' },
        { key: 'gender', label: 'Gender', type: 'select', options: GENDERS },
        { key: 'nationality', label: 'Nationality', type: 'select', options: COUNTRIES },
        { key: 'eth_band', label: 'ETH Band', type: 'select', options: ETH_BANDS },
        { key: 'sub_study_field_code', label: 'Sub Study Field Code', type: 'select', options: SUB_STUDY_FIELDS },
        { key: 'appointment_date', label: 'Appointment Date', type: 'date' },
        { key: 'employment_type', label: 'Employment Type', type: 'select', options: ['FT', 'PT'] },
        { key: 'academic_rank', label: 'Academic Rank' },
        { key: 'isced_education_level', label: 'ISCED Education Level' },
        { key: 'is_under_qualified', label: 'Is Under-qualified?', type: 'select', options: ['true', 'false'] },
        { key: 'nbr_courses_taught', label: 'Courses Taught', type: 'number' },
        { key: 'nbr_credit_hours_taught', label: 'Credit Hours Taught', type: 'number' },
    ],
    // Settings/Admin Modules
        'ROLE_CRUD': [
        { key: 'id', label: 'Role ID', isReadOnly: true, isHidden: true },
        { key: 'name', label: 'Role Name', isRequired: true },
        { key: 'username', label: 'Username', isHidden: false }, // Preserved for CRUD compliance
        { key: 'password', label: 'Password', type: 'password', isHidden: false }, // Preserved for CRUD compliance
        // Removed permissions from here
    ],
    // Staff Change Password
    'STAFF_CHANGE_PASSWORD': [
        { key: 'oldPassword', label: 'Old Password', type: 'password', isRequired: true },
        { key: 'newPassword', label: 'New Password', type: 'password', isRequired: true },
        { key: 'confirmPassword', label: 'Confirm New Password', type: 'password', isRequired: true },
    ],
    // Staff Update Profile
    'STAFF_UPDATE_PROFILE': [
        { key: 'firstName', label: 'First Name', type: 'text' },
        { key: 'lastName', label: 'Last Name', type: 'text' },
        { key: 'email', label: 'Email', type: 'email' },
        { key: 'phone', label: 'Phone Number', type: 'tel' },
    ],
    //Staff signup
    'STAFF_SIGNUP': [
        { key: 'username', label: 'Username', type: 'text', isRequired: true },
        { key: 'password', label: 'Password', type: 'password', isRequired: true },
        { key: 'confirmPassword', label: 'Confirm Password', type: 'password', isRequired: true },
        { key: 'role', label: 'Role', type: 'select', options: ASSIGNABLE_ROLES, isRequired: true },
    ]
};

/* ----------------------------------------------------------------------
   8. NOTIFICATION & UTILITY BAR (Standard)
   ---------------------------------------------------------------------- */

const NotificationBar = ({ glassClass, currentUser, onClose }) => {
    const { getCollection, notificationTrigger, leaveUpdateTrigger } = useData();
    const [activeTab, setActiveTab] = useState('notifications');

    // Get today's reference point and NORMALIZE to midnight
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // --- 1. ETHIOPIAN HOLIDAYS LOGIC ---
    const upcomingHolidays = useMemo(() => {
        const holidays = [
            { name: "Genna (Et. Christmas)", date: "2026-01-07" },
            { name: "Timkat (Et. Epiphany)", date: "2026-01-19" },
            { name: "Adwa Victory Day", date: "2026-03-02" },
            { name: "Siklet (Good Friday)", date: "2026-04-10" },
            { name: "Fasika (Easter)", date: "2026-04-12" },
            { name: "Labour Day", date: "2026-05-01" }
        ];

        return holidays
            .filter(h => {
                const hDate = new Date(h.date);
                hDate.setHours(0,0,0,0);
                return hDate >= today;
            })
            .sort((a, b) => new Date(a.date) - new Date(b.date))
            .slice(0, 3);
    }, [today]);

    // --- 2. GENERATE DYNAMIC NOTIFICATIONS ---
    const notifications = useMemo(() => {
        const list = [];
        const userRole = currentUser?.role;
        const isHR = userRole === 'hr_manager' || userRole === 'admin' || userRole === 'hr_staff';

        // SECTION: UPCOMING HOLIDAYS
        list.push({ type: 'Title', message: 'UPCOMING HOLIDAYS' });
        upcomingHolidays.forEach(h => {
            const displayDate = new Date(h.date).toLocaleDateString('en-US', { month: 'short', day: '2-digit' });
            list.push({ type: 'Holiday', message: `📅 ${h.name} - ${displayDate}`, color: "bg-emerald-900/30" });
        });

        // SECTION: PROVISION HR WATCH
        if (isHR) {
            list.push({ type: 'Title', message: 'PROVISION HR WATCH' });
            const employees = getCollection('HR_EMPLOYEE_CRUD');
            const provisionStaff = employees.filter(emp => {
                if (!emp.seniority_date_gc || emp.employment_type === 'Permanent') return false;
                const startDate = new Date(emp.seniority_date_gc);
                const provisionEnd = new Date(startDate.setMonth(startDate.getMonth() + 6));
                provisionEnd.setHours(0,0,0,0);
                return provisionEnd >= today;
            });

            if (provisionStaff.length > 0) {
                provisionStaff.forEach(emp => {
                    const startDate = new Date(emp.seniority_date_gc);
                    const provisionEnd = new Date(startDate.setMonth(startDate.getMonth() + 6));
                    list.push({ 
                        type: 'Info', 
                        message: `⏳ ${emp.name}: Provision ends ${provisionEnd.toLocaleDateString()}. Evaluate for permanency.`, 
                        color: 'bg-purple-900/30' 
                    });
                });
            } else {
                list.push({ type: 'Info', message: 'No staff currently in provision.', color: 'bg-gray-800/30' });
            }
        }

        // SECTION: STAFF CURRENTLY ON LEAVE
        list.push({ type: 'Title', message: 'STAFF CURRENTLY ON LEAVE' });
        const approvedLeaves = getCollection('HR_LEAVE_CRUD').filter(rec => {
            if (rec.status !== 'Approved') return false;
            const endDate = new Date(rec.end_date);
            endDate.setHours(0, 0, 0, 0);
            return endDate >= today;
        });

        if (approvedLeaves.length > 0) {
            approvedLeaves.forEach(onLeave => {
                list.push({ 
                    type: 'Info', 
                    message: `✈️ ${onLeave.employee_name} (${onLeave.leave_type}) until ${onLeave.end_date}`, 
                    color: 'bg-blue-900/30' 
                });
            });
        } else {
            list.push({ type: 'Info', message: 'No staff currently on leave.', color: 'bg-gray-800/30' });
        }

        // SECTION: RETIREMENT WATCH
        if (userRole === 'hr_manager' || userRole === 'admin') {
            list.push({ type: 'Title', message: `RETIREMENT WATCH` });
            const employees = getCollection('HR_EMPLOYEE_CRUD');
            let retirementAlertsCount = 0;

            employees.forEach(emp => {
                if (!emp.dob_gc) return;
                const dob = new Date(emp.dob_gc);
                const retirementDate = new Date(dob.getFullYear() + 60, dob.getMonth(), dob.getDate());
                retirementDate.setHours(0,0,0,0);
                
                const timeDiff = retirementDate.getTime() - today.getTime();
                const monthsRemaining = timeDiff / (1000 * 60 * 60 * 24 * 30.44);

                if (monthsRemaining >= 0 && monthsRemaining <= 1) {
                    retirementAlertsCount++;
                    list.push({ 
                        type: 'Warning', 
                        message: `🚨 FINAL CALL: ${emp.name} retires soon (${retirementDate.toLocaleDateString()}).`, 
                        color: 'bg-red-900/50' 
                    });
                } 
                else if (monthsRemaining > 1 && monthsRemaining <= 6) {
                    retirementAlertsCount++;
                    list.push({ 
                        type: 'Warning', 
                        message: `⚠️ PREPARATION: ${emp.name} retires in ~${Math.ceil(monthsRemaining)} months.`, 
                        color: 'bg-orange-900/30' 
                    });
                }
            });
            if (retirementAlertsCount === 0) list.push({ type: 'Info', message: 'No retirements due.', color: 'bg-gray-800/30' });
        }

        // SECTION: RESIGNATION REMINDERS (FIXED FOR TODAY)
        if (isHR) {
            list.push({ type: 'Title', message: 'RESIGNATION REMINDERS' });
            const resignations = getCollection('RESIGNATION_CRUD').filter(r => {
                const lastDay = new Date(r.last_working_day);
                lastDay.setHours(0, 0, 0, 0); // Normalize to midnight
                // Inclusion check: if lastDay is today or in the future
                return lastDay.getTime() >= today.getTime() && r.acceptance_status === 'Accepted';
            });

            if (resignations.length > 0) {
                resignations.forEach(r => {
                    const lDay = new Date(r.last_working_day);
                    lDay.setHours(0,0,0,0);
                    const isToday = lDay.getTime() === today.getTime();

                    list.push({ 
                        type: 'Alert', 
                        message: isToday 
                            ? `🔥 TODAY: ${r.employee_name}'s final working day! Complete handover.` 
                            : `🏃 ${r.employee_name}: Resignation final on ${r.last_working_day}.`, 
                        color: isToday ? 'bg-red-600/40' : 'bg-rose-900/30' 
                    });
                });
            } else {
                list.push({ type: 'Info', message: 'No upcoming resignations.', color: 'bg-gray-800/30' });
            }
        }

        // SECTION: PENDING ACTIONS
        list.push({ type: 'Title', message: 'PENDING ACTIONS' });
        if (isHR) {
            const pendingLeaves = getCollection('HR_LEAVE_CRUD').filter(rec => rec.status === 'Pending').length;
            if (pendingLeaves > 0) list.push({ type: 'Alert', message: `⚠️ ${pendingLeaves} Leave requests need review.`, color: 'bg-yellow-900/30' });

            const pendingAppraisals = getCollection('HR_APPRAISAL_CRUD').filter(rec => rec.status === 'Pending').length;
            if (pendingAppraisals > 0) list.push({ type: 'Alert', message: `📝 ${pendingAppraisals} Appraisals pending completion.`, color: 'bg-yellow-800/30' });
        }

        return list;
    }, [notificationTrigger, leaveUpdateTrigger, currentUser, getCollection, upcomingHolidays, today]);

    // SECTION: ACTIVITY LOG
    const activityLog = useMemo(() => {
        const logs = [];
        const logsData = getCollection('HR_LEAVE_CRUD').slice(-3);
        logsData.forEach(l => logs.push({ time: 'Recent', text: `Record updated for ${l.employee_name}.`, icon: 'Clock' }));
        return logs.length > 0 ? logs : [{ time: 'System', text: 'Database synchronized.', icon: 'RefreshCw' }];
    }, [getCollection]);

    return (
        <div className="h-full flex flex-col p-4 text-white relative">
            <div className="flex justify-between items-center mb-4 border-b border-cyan-500/50 pb-2">
                <div className="flex items-center text-cyan-400">
                    <button onClick={onClose} className="mr-2 hover:bg-white/10 p-1 rounded transition-colors"><Icons.ChevronRight className="w-5 h-5" /></button>
                    <h3 className="text-lg font-bold italic tracking-tight">Utility Hub</h3>
                </div>
            </div>
            
            <div className="flex space-x-2 mb-4">
                <button onClick={() => setActiveTab('notifications')} className={`flex-1 p-2 rounded-t-lg text-xs font-bold transition-all ${activeTab === 'notifications' ? 'bg-cyan-600 border-b-2 border-white' : 'bg-gray-700/50 text-gray-400'}`}>Notifications</button>
                <button onClick={() => setActiveTab('activity')} className={`flex-1 p-2 rounded-t-lg text-xs font-bold transition-all ${activeTab === 'activity' ? 'bg-cyan-600 border-b-2 border-white' : 'bg-gray-700/50 text-gray-400'}`}>Activity Log</button>
            </div>

            <div className="flex-grow overflow-y-auto space-y-2 custom-scrollbar pr-1">
                {activeTab === 'notifications' ? (
                    notifications.map((item, index) => (
                        item.type === 'Title' 
                        ? <div key={index} className="pt-4 pb-1"><span className="text-[10px] font-black text-cyan-400 uppercase tracking-[0.2em]">{item.message}</span></div>
                        : <div key={index} className={`p-3 rounded-lg text-[11px] border-l-2 border-white/20 hover:bg-white/5 transition-colors ${item.color}`}>{item.message}</div>
                    ))
                ) : (
                    activityLog.map((log, index) => (
                        <div key={index} className="p-3 rounded-lg bg-white/5 border border-white/10 flex items-start gap-3">
                            <Icons.Clock className="w-4 h-4 text-cyan-500 mt-0.5" />
                            <div>
                                <p className="text-[11px] text-gray-300">{log.text}</p>
                                <span className="text-[9px] text-gray-500 font-bold uppercase">{log.time}</span>
                            </div>
                        </div>
                    ))
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
                        HRM UNITY UNIVERSITY
                    </h1>
                    <p className="text-gray-400 mt-2 text-sm font-medium">Human Resources Terminal</p>
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
                
                    

                
                    <LoginModuleSwitcher current="hr" />
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
        const keyToMatch = isRecruitment ? 'form_no' : 'id_number';
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
            }
            sanitized[field.key] = field.type === 'array' ? (Array.isArray(val) ? val : []) : (val ?? '');
        });
        setFormData(sanitized);
    }, [initialData, columns, isNew, getCollection]);

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
        
        // Department Source for Transfer Destinations and Employee Records
        if ((field.key === 'department' || field.key === 'new_department' || field.key === 'current_department') && allData['DEPARTMENTS_CRUD']) {
            options = allData['DEPARTMENTS_CRUD'].map(d => d.name);
        }

        if ((isLeave || isTransfer || isTraining || isRetirement) && (field.key === 'employee_name' || field.key === 'name')) {
            options = (allData['HR_EMPLOYEE_CRUD'] || []).map(e => e.name);
        }

        if (isTraining && field.key === 'score') {
            options = ['Yes', 'No'];
        }

        if (isTransfer && field.key === 'new_location') {
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
                            
                            if ((field.key === 'form_no' && isRecruitment) || (field.key === 'id_number' && isEmployee)) {
                                handleDataLookup(v);
                            }

                            // --- AUTO-POPULATE (Transfer, Retirement, Training) ---
                            if ((isRetirement || isTransfer || isTraining) && (field.key === 'employee_name' || field.key === 'name')) {
                                const emp = (allData['HR_EMPLOYEE_CRUD'] || []).find(e => e.name === v);
                                if (emp) {
                                    updates.department = emp.department || '';
                                    updates.current_department = emp.department || ''; // TRANSFER: Current Dept
                                    updates.current_location = emp.work_location || ''; // TRANSFER: Current Loc
                                    updates.job_title = emp.job_title || '';
                                    updates.id_number = emp.id_number || '';
                                    if (isRetirement) {
                                        updates.dob_ec = emp.dob_ec || '';
                                        updates.retirement_date_ec = emp.retirement_date_ec || '';
                                    }
                                }
                            }

                            // --- LEAVE & DATE LOGIC ---
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

const UU_AI_KNOWLEDGE = [
    { id: 'sis', title: 'Unity SIS overview', tags: ['sis', 'unity', 'overview', 'modules', 'system'], body: 'Unity SIS integrates Finance, HR, and Material under shared login/RBAC. Finance: COA, journals, TB/GL, payroll, partners, FP/P&L/cash flow. HR: employee master, leave, recruitment, training, transfer, appraisal. Material: inventory, GRV, GIV, purchases; posts stock events to Finance.' },
    { id: 'finance', title: 'Finance module', tags: ['finance', 'coa', 'journal', 'payroll', 'trial', 'ledger'], body: 'Finance covers chart of accounts, single/journal transactions, trial balance, general ledger, partners, payroll process & reconciliation, and statements (FP, P&L, cash flow).' },
    { id: 'hr', title: 'HR module', tags: ['hr', 'employee', 'staff', 'leave', 'recruitment'], body: 'HR is the staff system of record: employee master (ID, names, department, position, salary, TIN, pension), leave, recruitment, training, transfer, appraisal, competence, awards.' },
    { id: 'material', title: 'Material module', tags: ['material', 'inventory', 'grv', 'giv', 'sku', 'store'], body: 'Material manages inventory SKUs, GRV (goods received), GIV (goods issued), local/foreign purchase, and parts robbing. GRV/GIV can post toward Finance stock/expense accounts.' },
    { id: 'students', title: 'SIS students', tags: ['student', 'students', 'tuition', 'registrar', 'qms'], body: 'Students connect via registrar/QMS/attendance concepts, tuition partners in Finance, and campus stores in Material.' },
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
        if (!q) return { query: '', topic: priorTopic || '', filters: {}, followUp: false };
        const followUp = /^(be )?specific|more detail|tell me more|what about|who are they|those|them|same|again|elaborate|details?|narrow|focus|only the/.test(q)
            || (q.split(/\s+/).length <= 8 && priorTopic && !/\b(grv|giv|payroll|inventory|sku|partner|coa|student)\b/.test(q));
        let topic = priorTopic || '';
        if (/\b(sku|skus|item|items|inventory|stock|product|tables?)\b/.test(q)) topic = 'sku';
        else if (/\b(employee|employees|staff|worker|personnel|female|male|women|men)\b/.test(q)) topic = 'employee';
        else if (/\b(grv|goods received)\b/.test(q)) topic = 'grv';
        else if (/\b(giv|goods issued)\b/.test(q)) topic = 'giv';
        else if (/\bleave\b/.test(q)) topic = 'leave';
        else if (/\b(partner|supplier|customer|vendor)\b/.test(q)) topic = 'partner';
        else if (/\b(coa|chart of account)/.test(q)) topic = 'coa';
        else if (/\b(payroll|payslip|salary process)\b/.test(q)) topic = 'payroll';
        else if (/\b(transaction|journal|trial balance|\btb\b)\b/.test(q)) topic = 'tx';
        else if (followUp && priorTopic) topic = priorTopic;
        let query = raw;
        if (followUp && priorTopic) query = raw + ' (about ' + priorTopic + (priorQuestion ? '; earlier: ' + priorQuestion : '') + ')';
        const filterSrc = query + (followUp && priorQuestion ? ' ' + priorQuestion : '');
        return { query, topic, followUp: !!followUp, filters: uuParseEmployeeFilters(filterSrc) };
    } catch (err) {
        return { query: String(question || ''), topic: priorTopic || '', filters: {}, followUp: false };
    }
};
const unityUniversityLiveAnswer = (question, allData, getCollection, topicHint, filtersIn) => {
    try {
        const q = String(question || '').toLowerCase().trim();
        if (!q && !topicHint) return null;
        const filters = filtersIn || uuParseEmployeeFilters(q) || {};
        let topic = topicHint || '';
        if (!topic) {
            if (/\b(sku|skus|item|items|inventory|stock|tables?)\b/.test(q)) topic = 'sku';
            else if (/\b(employee|staff|female|male|women|men|personnel)\b/.test(q) || filters.gender) topic = 'employee';
            else if (/\bgrv\b/.test(q)) topic = 'grv';
            else if (/\bgiv\b/.test(q)) topic = 'giv';
            else if (/\bleave\b/.test(q)) topic = 'leave';
            else if (/\b(partner|supplier|customer)\b/.test(q)) topic = 'partner';
            else if (/\b(coa|chart of account)/.test(q)) topic = 'coa';
            else if (/\b(payroll|payslip)\b/.test(q)) topic = 'payroll';
            else if (/\b(transaction|journal|\btb\b)\b/.test(q)) topic = 'tx';
        }
        const employees = uuCorrectEmployeeList(uuPull(allData, getCollection, 'HR_EMPLOYEE_CRUD'));
        const leave = uuPull(allData, getCollection, 'HR_LEAVE_CRUD');
        const items = uuPull(allData, getCollection, 'MAT_ITEM_CRUD');
        const grv = uuPull(allData, getCollection, 'MAT_GRV_CRUD');
        const giv = uuPull(allData, getCollection, 'MAT_GIV_CRUD');
        const partners = uuPull(allData, getCollection, 'FIN_PARTNERS_CRUD');
        const coa = uuPull(allData, getCollection, 'FIN_COA_CRUD');
        const tx = uuPull(allData, getCollection, 'transactions');
        const process = uuPull(allData, getCollection, 'PAY_PROCESS_CRUD');

        if (topic === 'employee' || filters.gender || /\b(employee|staff|female|male|women|men|personnel)\b/.test(q)) {
            const census = uuGenderCensus(employees);
            const filtered = uuFilterEmployees(employees, filters);
            const filteredCensus = uuGenderCensus(filtered);
            const n = filtered.length;
            let logicNote = filters.gender
                ? ('Logic check: list length ' + n + ' vs ' + filters.gender + ' in list ' + (filters.gender === 'female' ? filteredCensus.female : filteredCensus.male) + '.')
                : '';
            const depts = {};
            filtered.forEach((e) => { const d = uuEmpDept(e); depts[d] = (depts[d] || 0) + 1; });
            const deptLines = Object.entries(depts).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([d, c]) => '• ' + d + ': ' + c).join('\n');
            const lines = filtered.slice(0, 30).map((e) => '• ' + uuEmpId(e) + ' | ' + uuEmpName(e) + ' | ' + uuNormalizeGender(uuEmpGenderRaw(e)) + ' | ' + uuEmpPos(e) + ' | ' + uuEmpDept(e));
            const filterBits = [];
            if (filters.gender) filterBits.push('gender=' + filters.gender);
            if (filters.department) filterBits.push('department~' + filters.department);
            if (filters.location) filterBits.push('location~' + filters.location);
            let head = 'Employees (live HR data)\n• Matching: ' + n + '\n• Total in master: ' + census.total + '\n• Full split: female ' + census.female + ' · male ' + census.male + ' · unspecified ' + census.unknown + '\n';
            if (n === 0) {
                return { title: filters.gender ? ('Live · ' + filters.gender + ' employees') : 'Live · HR employees', body: head + '• Filters: ' + (filterBits.join(', ') || 'none') + '\n\nNo rows match.\n' + logicNote + '\nSource: HR_EMPLOYEE_CRUD.' };
            }
            return {
                title: filters.gender ? ('Live · ' + filters.gender + ' employees (' + n + ')') : ('Live · HR employees (' + n + ')'),
                body: head + '• Active filters: ' + (filterBits.join(', ') || 'none') + '\n\nBy department:\n' + (deptLines || '• —') + '\n\nPeople (' + Math.min(n, 30) + (n > 30 ? ' of ' + n : '') + '):\n' + lines.join('\n') + '\n\n' + logicNote + '\nGender uses the record field, corrected from first name when the name is a known male/female Ethiopian name.\nSource: HR_EMPLOYEE_CRUD.',
            };
        }
        if (topic === 'sku' || /\b(sku|skus|inventory|stock|tables?)\b/.test(q)) {
            const n = items.length;
            const lines = items.slice(0, 15).map((r) => {
                const sku = (r && (r.sku || r.item_sku || r.item_code || r.code || r.id)) || '—';
                const name = (r && (r.name || r.item_name || r.description)) || '';
                return '• ' + sku + (name ? ' — ' + name : '') + ' | qty ' + ((r && (r.quantity != null ? r.quantity : r.qty)) ?? '—');
            });
            const stores = {};
            items.forEach((r) => {
                const s = (r && (r.location || r.store || r.warehouse || r.receiving_store)) || 'Main store';
                stores[s] = (stores[s] || 0) + 1;
            });
            const storeLines = Object.entries(stores).sort((a, b) => b[1] - a[1]).map(([s, c]) => '• ' + s + ': ' + c + ' item row(s)').join('\n');
            const tableHint = /table/.test(q) ? '\n\n"Tables in store" = inventory item rows by store (not SQL tables).' : '';
            return { title: 'Live · Store / SKUs (' + n + ')', body: n ? ('Inventory SKU rows: ' + n + '\nBy store:\n' + (storeLines || '• —') + '\n\nSample:\n' + lines.join('\n') + tableHint) : ('No inventory rows.' + tableHint) };
        }
        if (topic === 'grv' || /\bgrv\b/.test(q)) return { title: 'Live · GRV', body: grv.length + ' GRV record(s). SKUs: ' + items.length + '.' };
        if (topic === 'giv' || /\bgiv\b/.test(q)) return { title: 'Live · GIV', body: giv.length + ' GIV record(s).' };
        if (topic === 'leave' || /\bleave\b/.test(q)) return { title: 'Live · Leave', body: leave.length + ' leave record(s). Employees: ' + employees.length + '.' };
        if (topic === 'partner' || /\b(partner|supplier|customer)\b/.test(q)) return { title: 'Live · Partners', body: partners.length + ' partner(s).' };
        if (topic === 'coa' || /\bcoa\b/.test(q)) return { title: 'Live · COA', body: coa.length + ' chart-of-account row(s).' };
        if (topic === 'payroll' || /\bpayroll\b/.test(q)) return { title: 'Live · Payroll process', body: process.length + ' payroll process row(s).' };
        if (topic === 'tx' || /\b(transaction|journal|\btb\b)\b/.test(q)) return { title: 'Live · Transactions', body: tx.length + ' transaction/journal row(s).' };
        return null;
    } catch (err) {
        return { title: 'unityUniversity AI', body: 'Live data step failed: ' + String(err && err.message ? err.message : err) };
    }
};
const unityUniversityAnswer = (question, allData, getCollection, priorTopic, priorQuestion) => {
    try {
        const raw = String(question || '').trim();
        const resolved = uuResolveQuery(raw, priorTopic, priorQuestion);
        if (!raw) {
            return { title: 'unityUniversity AI', body: 'Ask live questions:\n• how many female employees\n• how many male staff\n• how many skus\n• how many GRV', topic: '' };
        }
        const live = unityUniversityLiveAnswer(resolved.query, allData, getCollection, resolved.topic, resolved.filters);
        if (live && live.body) return Object.assign({}, live, { topic: resolved.topic || priorTopic || '' });
        const q = String(resolved.query || '').toLowerCase();
        const words = q.split(/[^a-z0-9]+/).filter((w) => w.length > 2);
        let best = null, bestScore = 0;
        const knowledge = Array.isArray(UU_AI_KNOWLEDGE) ? UU_AI_KNOWLEDGE : [];
        for (let i = 0; i < knowledge.length; i++) {
            const item = knowledge[i];
            if (!item) continue;
            let score = 0;
            const titleL = String(item.title || '').toLowerCase();
            if (q.includes(titleL)) score += 6;
            const tags = Array.isArray(item.tags) ? item.tags : [];
            for (let t = 0; t < tags.length; t++) {
                const tag = String(tags[t] || '');
                if (q.includes(tag)) score += 4;
                for (let w = 0; w < words.length; w++) {
                    if (tag === words[w]) score += 3;
                    else if (tag.includes(words[w]) || words[w].includes(tag)) score += 1;
                }
            }
            for (let w = 0; w < words.length; w++) {
                if (titleL.includes(words[w])) score += 2;
                if (String(item.body || '').toLowerCase().includes(words[w])) score += 0.2;
            }
            if (score > bestScore) { bestScore = score; best = item; }
        }
        if (best && bestScore >= 1.5) return { title: best.title, body: best.body, topic: resolved.topic || best.id };
        return { title: 'unityUniversity AI', body: 'No strong match. Try: how many female employees · how many SKUs · what is Unity SIS', topic: resolved.topic || priorTopic || '' };
    } catch (err) {
        return { title: 'unityUniversity AI', body: 'Answer error: ' + String(err && err.message ? err.message : err), topic: priorTopic || '' };
    }
};
const UnityUniversityAI = ({ glassClass = '' }) => {
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
    const [answer, setAnswer] = useState({ title: 'unityUniversity AI', body: 'Ask live questions — answers appear here at the top.\nExamples: how many female employees · how many SKUs · what is Unity SIS', topic: '' });
    const [history, setHistory] = useState([]);
    const [priorTopic, setPriorTopic] = useState('');
    const [priorQuestion, setPriorQuestion] = useState('');
    const [busy, setBusy] = useState(false);
    const answerRef = useRef(null);
    const ask = (q) => {
        const query = String(q != null && String(q).length ? q : question).trim();
        setBusy(true);
        let ad = allData;
        let gc = getCollection;
        try { if (typeof window !== 'undefined' && window.__UU_ALL_DATA__) ad = window.__UU_ALL_DATA__; } catch (e) {}
        window.setTimeout(() => {
            try {
                const res = unityUniversityAnswer(query, ad, gc, priorTopic, priorQuestion);
                const safe = (res && res.body) ? res : { title: 'unityUniversity AI', body: 'No answer.', topic: priorTopic };
                setAnswer(safe);
                if (safe.topic) setPriorTopic(safe.topic);
                if (query) {
                    setPriorQuestion(query);
                    setHistory((h) => [{ q: query, title: safe.title }, ...(Array.isArray(h) ? h : [])].slice(0, 12));
                }
                setQuestion('');
                window.setTimeout(() => { try { answerRef.current && answerRef.current.scrollIntoView && answerRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch (e2) {} }, 20);
            } catch (err) {
                setAnswer({ title: 'unityUniversity AI', body: 'Could not answer: ' + String(err && err.message ? err.message : err), topic: priorTopic });
            } finally { setBusy(false); }
        }, 40);
    };
    const quick = ['How many female employees?', 'How many male employees?', 'How many employees?', 'Who are the female staff?', 'How many SKUs?', 'How many GRV?', 'Payroll process count', 'What is Unity SIS?'];
    return (
        <div className={'mt-6 rounded-[1.75rem] border border-cyan-500/30 bg-gradient-to-br from-cyan-950/50 to-slate-950/80 ' + (glassClass || '') + ' overflow-hidden'}>
            <div className="px-5 py-4 border-b border-cyan-500/20 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-cyan-500/20 border border-cyan-400/40 flex items-center justify-center text-cyan-300 text-lg font-black">AI</div>
                    <div>
                        <h3 className="text-sm font-black uppercase tracking-widest text-cyan-300">unityUniversity AI</h3>
                        <p className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">Consistent live logic{priorTopic ? ' · ' + priorTopic : ''}</p>
                    </div>
                </div>
                <span className="text-[9px] font-black uppercase px-2 py-1 rounded-lg bg-cyan-500/10 text-cyan-200/90 border border-cyan-500/20">{busy ? 'Thinking…' : 'Ask anything'}</span>
            </div>
            <div className="p-5 space-y-4">
                <div ref={answerRef} className="rounded-2xl border border-cyan-400/30 bg-black/60 p-4 space-y-2 min-h-[8rem]">
                    <h4 className="text-xs font-black uppercase tracking-widest text-cyan-300">{(answer && answer.title) || 'unityUniversity AI'}</h4>
                    <p className="text-[13px] text-gray-100 leading-relaxed whitespace-pre-wrap">{(answer && answer.body) || 'Answer appears here at the top.'}</p>
                </div>
                <div className="flex flex-wrap gap-2">
                    {quick.map((label) => (
                        <button key={label} type="button" onClick={() => ask(label)} disabled={busy} className="px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-wide bg-white/5 border border-white/10 text-gray-300 hover:border-cyan-400/40 hover:text-cyan-200 transition disabled:opacity-50">{label}</button>
                    ))}
                </div>
                <div className="flex flex-col sm:flex-row gap-2">
                    <input type="text" value={question} onChange={(e) => setQuestion(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); ask(); } }} placeholder="e.g. how many female employees" className="flex-1 min-w-0 px-4 py-3 rounded-xl bg-black/50 border border-cyan-500/25 text-sm text-white outline-none focus:border-cyan-400" disabled={busy} />
                    <button type="button" onClick={() => ask()} disabled={busy} className="px-5 py-3 rounded-xl bg-cyan-600 hover:bg-cyan-500 disabled:bg-cyan-900 text-white text-xs font-black uppercase tracking-wider shrink-0">{busy ? '…' : 'Ask'}</button>
                </div>
                {history.length > 0 && (
                    <div className="space-y-1">
                        <p className="text-[9px] font-black uppercase tracking-widest text-gray-500">Recent</p>
                        <div className="flex flex-wrap gap-2">
                            {history.map((h, idx) => (
                                <button key={String(h.q) + '-' + idx} type="button" onClick={() => ask(h.q)} className="text-[10px] px-2 py-1 rounded-lg bg-white/5 text-gray-400 hover:text-cyan-200">{h.q}</button>
                            ))}
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};



/* —— Ethiopian (Geez) ↔ Gregorian · Power Geez calendar —— */
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
            <EthiopianCalendarPanel glassClass={glassClass} />

            <UnityUniversityAI glassClass={glassClass} />

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

// 2. MODIFIED RENDERVIEW - FIXED ACCESS DENIED & MODULE ROUTING
const renderView = (currentView, glassClass, getCollection, checkPermission, allRoles = []) => {
    const collectionName = currentView;
    
    // --- FIX: Access Denied Bypass for System Modules ---
    const isSystemModule = [
        'ROLE_ASSIGNMENT_VIEW', 
        'HR_ROLE_ASSIGNMENT_CRUD', 
        'UI_CUSTOMIZATION', 
        'STAFF_REGISTER', 
        'REGISTER'
    ].includes(currentView);

    // Only run security gate if it's NOT a core system module
    if (!isSystemModule && !checkPermission(currentView, allRoles)) {
        return <AccessDenied glassClass={glassClass} />;
    }

    const title = findMenuItemTitle(MENU_ITEMS, currentView) || currentView.replace(/_/g, ' ');
    const thStyle = "p-4 text-xs font-bold text-gray-400 uppercase tracking-wider border-b border-white/10";

    switch (currentView) {
        case 'HR_DASHBOARD_VIEW':
            return <HrDashboard glassClass={glassClass} />;
            
        case 'ROLE_CRUD':
            return <RoleManagement glassClass={glassClass} />;

        // --- FIXED: ROLE ASSIGNMENT MODULE ---
        case 'ROLE_ASSIGNMENT_VIEW':
        case 'HR_ROLE_ASSIGNMENT_CRUD':
            return <RoleAssignmentManager glassClass={glassClass} />;

        // --- FIXED: UI CUSTOMIZATION MODULE ---
        case 'UI_CUSTOMIZATION':
            return <UICustomization glassClass={glassClass} />;

        // --- FIXED: REGISTER / STAFF SIGNUP MODULE ---
        case 'STAFF_REGISTER':
        case 'REGISTER':
        case 'STAFF_SIGNUP':
            return <StaffSignup glassClass={glassClass} />;

        case 'STAFF_TRANSFER':
            const transfers = getCollection('STAFF_TRANSFERS') || [];
            return (
                <div className="space-y-6 animate-in fade-in duration-500">
                    <div className="flex justify-between items-center">
                        <h2 className="text-xl font-black text-white flex items-center gap-2">
                            <Icons.MapPin className="text-cyan-500" /> LOCATION TRANSFERS
                        </h2>
                    </div>
                    <div className={`rounded-xl border border-white/10 overflow-hidden ${glassClass}`}>
                        <div className="uu-table-scroll overflow-x-auto">
                            <table className="w-full text-left border-collapse table-fixed min-w-[900px]">
                                <thead>
                                    <tr className="bg-white/5">
                                        <th className={`${thStyle} w-[25%]`}>Staff Member</th>
                                        <th className={`${thStyle} w-[20%]`}>From Place</th>
                                        <th className={`${thStyle} w-[20%]`}>To Place</th>
                                        <th className={`${thStyle} w-[20%]`}>Date</th>
                                        <th className={`${thStyle} w-[15%] text-right`}>Actions</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-white/5">
                                    {transfers.map((item) => (
                                        <tr key={item.id} className="hover:bg-white/5 transition-colors">
                                            <td className="p-4 align-middle font-bold text-white text-sm">{item.employeeName}</td>
                                            <td className="p-4 align-middle text-red-400 text-xs font-bold">{item.fromPlace}</td>
                                            <td className="p-4 align-middle text-green-400 text-xs font-bold">{item.toPlace}</td>
                                            <td className="p-4 align-middle text-gray-400 text-xs">{item.transferDate}</td>
                                            <td className="p-4 text-right align-middle"><Icons.MoreVertical size={16} className="text-gray-500" /></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            );

        case 'HR_RESIGNATION_CRUD':
            const resignations = getCollection('HR_RESIGNATION_CRUD') || [];
            return (
                <div className="space-y-6 animate-in fade-in duration-500">
                    <h2 className="text-xl font-black text-white flex items-center gap-2">
                        <Icons.LogOut className="text-orange-500" /> RESIGNATION RECORDS
                    </h2>
                    <div className={`rounded-xl border border-white/10 overflow-hidden ${glassClass}`}>
                        <div className="uu-table-scroll overflow-x-auto">
                            <table className="w-full text-left border-collapse table-fixed min-w-[900px]">
                                <thead>
                                    <tr className="bg-white/5">
                                        <th className={`${thStyle} w-[25%]`}>Employee</th>
                                        <th className={`${thStyle} w-[20%]`}>Notice Date</th>
                                        <th className={`${thStyle} w-[20%]`}>Last Working Day</th>
                                        <th className={`${thStyle} w-[20%]`}>Status</th>
                                        <th className={`${thStyle} w-[15%] text-right`}>Actions</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-white/5">
                                    {resignations.map((res) => (
                                        <tr key={res.id} className="hover:bg-white/5 transition-colors">
                                            <td className="p-4 align-middle">
                                                <div className="font-bold text-white text-sm">{res.employee_name}</div>
                                                <div className="text-[10px] text-gray-500 uppercase">{res.employee_id}</div>
                                            </td>
                                            <td className="p-4 align-middle text-xs text-gray-300">{res.notice_date}</td>
                                            <td className="p-4 align-middle text-xs text-white font-medium">{res.last_working_day}</td>
                                            <td className="p-4 align-middle">
                                                <span className="px-2 py-1 rounded bg-orange-500/10 text-orange-400 text-[10px] font-bold border border-orange-500/20">
                                                    {res.status || 'Processed'}
                                                </span>
                                            </td>
                                            <td className="p-4 text-right align-middle">
                                                <button className="p-1.5 hover:bg-white/10 text-gray-400 rounded-lg border border-white/10"><Icons.Eye size={14} /></button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            );

        case 'HR_DEATH_CRUD':
            const deathRecords = getCollection('HR_DEATH_CRUD') || [];
            return (
                <div className="space-y-6 animate-in fade-in duration-500">
                    <h2 className="text-xl font-black text-white flex items-center gap-2">
                        <Icons.HeartOff className="text-red-500" /> DEATH RECORDS
                    </h2>
                    <div className={`rounded-xl border border-white/10 overflow-hidden ${glassClass}`}>
                        <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse table-fixed min-w-[900px]">
                                <thead>
                                    <tr className="bg-white/5">
                                        <th className={`${thStyle} w-[25%]`}>Employee</th>
                                        <th className={`${thStyle} w-[20%]`}>Date of Death</th>
                                        <th className={`${thStyle} w-[25%]`}>Cause / Notes</th>
                                        <th className={`${thStyle} w-[15%]`}>Benefits</th>
                                        <th className={`${thStyle} w-[15%] text-right`}>Actions</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-white/5">
                                    {deathRecords.map((record) => (
                                        <tr key={record.id} className="hover:bg-white/5 transition-colors">
                                            <td className="p-4 align-middle">
                                                <div className="font-bold text-white text-sm">{record.employee_name}</div>
                                                <div className="text-[10px] text-red-400 font-bold uppercase">{record.employee_id}</div>
                                            </td>
                                            <td className="p-4 align-middle text-xs text-gray-300">{record.date_of_death}</td>
                                            <td className="p-4 align-middle text-xs text-gray-400 truncate max-w-[200px]">{record.reason || record.notes || 'N/A'}</td>
                                            <td className="p-4 align-middle"><span className="text-[10px] font-bold text-green-400 uppercase">Paid</span></td>
                                            <td className="p-4 text-right align-middle">
                                                <button className="p-1.5 hover:bg-white/10 text-gray-400 rounded-lg border border-white/10"><Icons.FileText size={14} /></button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            );

        case 'HR_GENERIC_REPORTS':
            return <ReportsGenerator glassClass={glassClass} />;

        case 'STAFF_CHANGE_PASSWORD':
            return <StaffChangePassword glassClass={glassClass} />;

        case 'STAFF_UPDATE_PROFILE':
            return <StaffUpdateProfile glassClass={glassClass} />;

        default:
            if (COLUMN_CONFIG && COLUMN_CONFIG[collectionName]) {
                return (
                    <div className="w-full">
                    <DataTable 
                        collectionName={collectionName} 
                        columns={COLUMN_CONFIG[collectionName] || []} 
                        data={getCollection(collectionName)}
                        glassClass={glassClass}  
                        getCollection={getCollection} 
                        canAdd={checkPermission('CAN_ADD_RECORD', allRoles)}
                        canEdit={checkPermission('CAN_EDIT_RECORD', allRoles)}
                        canDelete={checkPermission('CAN_DELETE_RECORD', allRoles)}
                    />
                    </div>
                );
            }
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
                <span className="text-xl font-bold text-white tracking-tight">Unity HRMS</span>
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
    if (currentUser && typeof canAccessModule === 'function' && !canAccessModule(currentUser.role, 'hr', currentUser)) {
        return (
            <ModuleAccessGate moduleId="hr" user={currentUser} onLogout={logout} />
        );
    }
    if (currentUser && typeof canAccessModule !== 'function') {
        const _r = String(currentUser.role || '').toLowerCase().replace(/\s+/g, '_');
        const _admin = _r === 'admin' || _r === 'super_admin';
        const _allow = { finance: ['finance_officer','finance_manager','accountant','payroll_officer','hr_manager'], hr: ['hr_manager','hr_staff','hr_officer','recruiter'], material: ['store_manager','inventory_clerk','procurement_officer','warehouse'] }['hr'] || [];
        if (!_admin && !_allow.includes(_r) && !(Array.isArray(currentUser.modules) && currentUser.modules.includes('hr'))) {
            return (
                <ModuleAccessGate moduleId="hr" user={currentUser} onLogout={logout} />
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
                    <button type="button" onClick={() => setShowNotifications(!showNotifications)} className={`p-2 rounded-full transition-all ${showNotifications ? 'bg-cyan-500 text-white' : 'text-gray-400'}`}>
                        <Icons.Bell className="w-5 h-5" />
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

                {showNotifications && (
                    <aside className={`w-[25%] border-l border-white/10 animate-in slide-in-from-right ${glassClass}`}>
                        <NotificationBar 
                            glassClass={glassClass} 
                            currentUser={currentUser} 
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

export function HRApp() {
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
export default HRApp;

