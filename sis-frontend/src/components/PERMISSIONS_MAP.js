// PERMISSIONS_MAP.js (Updated to support specialized roles)

/**
 * Defines which roles are authorized for which features or actions (Permissions).
 * The list now includes highly specialized roles for fine-grained access.
 */
export const PERMISSIONS_MAP = {
    // ----------------------------------------------------
    // Existing Broad Roles:
    // 'admin': Full System Access (Superuser)
    // 'hr_manager': High-level HR Access
    // 'hr_staff': Mid-level HR Access
    // 'employee': Self-Service only
    // ----------------------------------------------------
    
    // HR Module View Access (Sidebar IDs)
    'HR_RECRUITMENT_CRUD': ['admin', 'hr_manager', 'hr_staff', 'recruiter'],
    'HR_EMPLOYEE_CRUD': ['admin', 'hr_manager', 'hr_data_entry'],
    'HR_COMPETENCE_CRUD': ['admin', 'hr_manager', 'hr_staff', 'comp_assessor'],
    'HR_APPRAISAL_CRUD': ['admin', 'hr_manager', 'hr_staff', 'appraisal_admin'],
    'HR_LEAVE_CRUD': ['admin', 'hr_manager', 'hr_staff', 'leave_processor', 'employee'], // Employee can access own leave records
    'HR_PENALTY_CRUD': ['admin', 'hr_manager', 'disciplinary_officer'],
    'HR_AWARD_CRUD': ['admin', 'hr_manager', 'awards_admin'],
    'HR_TRANSFER_CRUD': ['admin', 'hr_manager', 'transfer_officer'],
    'HR_TRAINING_CRUD': ['admin', 'hr_manager', 'hr_staff', 'training_officer'],
    'HR_INSURANCE_CRUD': ['admin', 'hr_manager', 'insurance_admin'],
    'HR_REPORT_CRUD': ['admin', 'hr_manager', 'report_viewer'],

    // Granular CRUD Permissions:
    
    // Only broad roles and some specialized roles are allowed to add/edit
    'CAN_ADD_RECORD': ['admin', 'hr_manager', 'hr_staff', 
                       'recruiter', 'hr_data_entry', 'comp_assessor', 
                       'appraisal_admin', 'leave_processor', 'disciplinary_officer',
                       'awards_admin', 'transfer_officer', 'training_officer', 'insurance_admin'],
                       
    'CAN_EDIT_RECORD': ['admin', 'hr_manager', 'hr_staff', 'employee', // Employee can edit their own profile
                        'recruiter', 'hr_data_entry', 'comp_assessor', 
                        'appraisal_admin', 'leave_processor', 'disciplinary_officer',
                        'awards_admin', 'transfer_officer', 'training_officer', 'insurance_admin'],
                        
    // Only Admins and HR Managers can delete records
    'CAN_DELETE_RECORD': ['admin', 'hr_manager'],
};

export const DEFAULT_ROLE = 'employee';