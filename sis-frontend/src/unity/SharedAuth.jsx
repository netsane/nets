import React, { createContext, useContext, useState, useCallback, useMemo } from 'react';
import { getSharedUsers, resolveAssignedRole, changeSharedPassword, attemptSharedLogin, saveSharedSession, loadSharedSession } from './SharedCredentials.js';

const AuthContext = createContext(null);

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used within AuthProvider');
  return ctx;
}

/** Demo users — roles drive RBAC. Password for all demos: password */
export const MOCK_USERS = [
  { username: 'admin', password: 'password', role: 'admin', name: 'System Administrator' },
  { username: 'hr_manager', password: 'password', role: 'hr_manager', name: 'HR Manager' },
  { username: 'hr_staff', password: 'password', role: 'hr_staff', name: 'HR Officer' },
  { username: 'store_manager', password: 'password', role: 'store_manager', name: 'Store Manager' },
  { username: 'inventory_clerk', password: 'password', role: 'inventory_clerk', name: 'Inventory Clerk' },
  { username: 'procurement_officer', password: 'password', role: 'procurement_officer', name: 'Procurement Officer' },
  { username: 'finance_officer', password: 'password', role: 'finance_officer', name: 'Finance Officer' },
  { username: 'employee', password: 'password', role: 'employee', name: 'Staff Member' },
];

export const DEFAULT_PERMISSIONS_MAP = {
  FIN_DASHBOARD: ['admin', 'finance_officer', 'hr_manager'],
  FIN_COA_CRUD: ['admin', 'finance_officer'],
  FIN_TXN_CRUD: ['admin', 'finance_officer'],
  FIN_JOURNAL_CRUD: ['admin', 'finance_officer'],
  FIN_TB_DISPLAY: ['admin', 'finance_officer', 'hr_manager'],
  PAY_MASTER_CRUD: ['admin', 'hr_manager', 'finance_officer'],
  PAY_PROCESS_CRUD: ['admin', 'hr_manager', 'finance_officer'],
  PAY_OVERTIME_CRUD: ['admin', 'hr_manager'],
  PAY_AD_CRUD: ['admin', 'hr_manager'],
  PAY_RECON_CRUD: ['admin', 'finance_officer'],
  HR_DASHBOARD_VIEW: ['admin', 'hr_manager', 'hr_staff', 'employee'],
  HR_EMPLOYEE_CRUD: ['admin', 'hr_manager', 'hr_staff'],
  HR_LEAVE_CRUD: ['admin', 'hr_manager', 'hr_staff', 'employee'],
  HR_PENALTY_CRUD: ['admin', 'hr_manager'],
  HR_AWARD_CRUD: ['admin', 'hr_manager', 'hr_staff'],
  HR_TRAINING_CRUD: ['admin', 'hr_manager', 'hr_staff', 'employee'],
  ROLE_CRUD: ['admin'],
  MAT_DASHBOARD_VIEW: ['admin', 'store_manager', 'inventory_clerk', 'procurement_officer'],
  MAT_ITEM_CRUD: ['admin', 'store_manager'],
  MAT_GRV_CRUD: ['admin', 'store_manager', 'inventory_clerk'],
  MAT_GIV_CRUD: ['admin', 'store_manager', 'inventory_clerk'],
  MAT_ROBBING_CRUD: ['admin', 'store_manager'],
  MAT_PURCHASE_LOCAL_CRUD: ['admin', 'procurement_officer'],
  MAT_PURCHASE_FOREIGN_CRUD: ['admin', 'procurement_officer'],
  UI_CUSTOMIZATION: ['admin'],
  CAN_ADD_RECORD: ['admin', 'hr_manager', 'hr_staff', 'store_manager', 'inventory_clerk', 'finance_officer', 'procurement_officer'],
  CAN_EDIT_RECORD: ['admin', 'hr_manager', 'hr_staff', 'store_manager', 'finance_officer'],
  CAN_DELETE_RECORD: ['admin', 'hr_manager'],
};

const STORAGE_KEY = 'uu_auth_session_v1';

function loadSession() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export function AuthProvider({ children, defaultView = 'FIN_DASHBOARD' }) {
  const saved = typeof localStorage !== 'undefined' ? loadSession() : null;
  const [isLoggedIn, setIsLoggedIn] = useState(!!saved);
  const [currentUser, setCurrentUser] = useState(saved || null);
  const [loading, setLoading] = useState(false);
  const [currentView, setCurrentView] = useState(defaultView);

  const persist = (user) => {
    try {
      if (user) localStorage.setItem(STORAGE_KEY, JSON.stringify(user));
      else localStorage.removeItem(STORAGE_KEY);
    } catch (_) {}
  };

  const logout = useCallback(() => {
    setIsLoggedIn(false);
    setCurrentUser(null);
    persist(null);
    setCurrentView('AUTH');
  }, []);

  /**
   * Role-based login.
   * 1) Optional HR employee list (username + password + role / system_role)
   * 2) MOCK_USERS fallback
   */
  const login = useCallback(
    async (username, password, allEmployees = []) => {
      setLoading(true);
      await new Promise((r) => setTimeout(r, 200));
      try {
        const result = attemptSharedLogin(username, password, allEmployees || []);
        if (result.success && result.user) {
          const assigned =
            typeof resolveAssignedRole === 'function'
              ? resolveAssignedRole(username)
              : null;
          const role =
            (assigned && assigned.role) || result.user.role || 'employee';
          const user = {
            ...result.user,
            role,
            system_role: role,
            role_id: (assigned && assigned.role_id) || result.user.role_id,
            role_name: (assigned && assigned.role_name) || result.user.role_name,
          };
          setCurrentUser(user);
          setIsLoggedIn(true);
          persist(user);
          try {
            saveSharedSession(user);
          } catch (_) {}
          setCurrentView(defaultView);
          setLoading(false);
          return { success: true, user };
        }
        setLoading(false);
        return {
          success: false,
          message: (result && result.message) || 'Invalid username or password.',
        };
      } catch (e) {
        console.error(e);
        setLoading(false);
        return { success: false, message: 'Login error.' };
      }
    },
    [defaultView]
  );

  const checkPermission = useCallback(
    (permissionId, allRoles = []) => {
      if (!isLoggedIn || !currentUser?.role) return false;
      const role = String(currentUser.role);
      if (role === 'admin' || currentUser.username === 'admin') return true;

      if (Array.isArray(allRoles) && allRoles.length) {
        const roleDef = allRoles.find(
          (r) =>
            r.id === role ||
            r.name === role ||
            (r.roleName && String(r.roleName).toLowerCase() === role.toLowerCase())
        );
        if (roleDef && Array.isArray(roleDef.permissions)) {
          return roleDef.permissions.includes(permissionId);
        }
      }

      const allowed = DEFAULT_PERMISSIONS_MAP[permissionId];
      if (allowed) {
        return allowed.map((x) => String(x).toLowerCase()).includes(role.toLowerCase());
      }
      return false;
    },
    [isLoggedIn, currentUser]
  );

  const value = useMemo(
    () => ({
      isLoggedIn,
      currentUser,
      loading,
      login,
      logout,
      checkPermission,
      currentView,
      setCurrentView,
      changePassword: async (oldPassword, newPassword) => {
        try {
          if (!currentUser?.username) return { success: false, message: 'Not signed in.' };
          return changeSharedPassword(currentUser.username, oldPassword, newPassword);
        } catch (e) {
          return { success: false, message: 'Password change failed.' };
        }
      },
      MOCK_USERS,
    }),
    [isLoggedIn, currentUser, loading, login, logout, checkPermission, currentView]
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export default AuthProvider;


/** Call after any module login so unity/api.js has JWT for PostgreSQL CRUD */
export async function ensureApiSession(username, password) {
  try {
    const { login } = await import('./api.js');
    await login(username, password);
    return true;
  } catch (e) {
    console.warn('[ensureApiSession]', e && e.message);
    return false;
  }
}
