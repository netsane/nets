/**
 * Shared users, roles, session across Finance / HR / Material.
 * Switching modules must NOT drop role or "expire" security context.
 */

const USERS_KEY = 'uu_shared_users_v1';
const ROLES_KEY = 'uu_shared_roles_v1';
const SESSION_KEY = 'uu_shared_session_v1';

const DEFAULT_USERS = [
  { username: 'admin', password: 'password', role: 'admin', name: 'System Administrator', modules: ['finance', 'hr', 'material'] },
  { username: 'hr_manager', password: 'password', role: 'hr_manager', name: 'HR Manager', modules: ['hr', 'finance'] },
  { username: 'hr_staff', password: 'password', role: 'hr_staff', name: 'HR Officer', modules: ['hr'] },
  { username: 'store_manager', password: 'password', role: 'store_manager', name: 'Store Manager', modules: ['material'] },
  { username: 'finance_officer', password: 'password', role: 'finance_officer', name: 'Finance Officer', modules: ['finance'] },
  { username: 'employee', password: 'password', role: 'employee', name: 'Staff Member', modules: ['hr'] },
];

export const MODULE_ACCESS = {
  finance: ['admin', 'finance_officer', 'finance_manager', 'accountant', 'payroll_officer', 'hr_manager'],
  hr: ['admin', 'hr_manager', 'hr_staff', 'hr_officer', 'recruiter'],
  material: ['admin', 'store_manager', 'inventory_clerk', 'procurement_officer', 'warehouse'],
};

function readJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    const v = JSON.parse(raw);
    return v == null ? fallback : v;
  } catch {
    return fallback;
  }
}

function writeJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.warn('[SharedCredentials] write failed', key, e);
  }
}

export function normalizeRole(role) {
  return String(role || 'employee').trim().toLowerCase().replace(/\s+/g, '_');
}


/** Raw assigned users from localStorage only (no DEFAULT merge). */
export function readAssignedUsers() {
  const saved = readJson(USERS_KEY, []);
  return Array.isArray(saved) ? saved : [];
}

/** Latest role for a username from assignment storage — source of truth for login. */
export function resolveAssignedRole(username) {
  const uName = String(username || '').trim().toLowerCase();
  if (!uName) return null;
  const saved = readAssignedUsers();
  const row = saved.find((u) => String(u.username || '').trim().toLowerCase() === uName);
  if (!row) return null;
  const role = row.role || row.system_role || null;
  if (role == null || String(role).trim() === '') return null;
  return {
    role: String(role).trim(),
    system_role: String(row.system_role || role).trim(),
    role_id: row.role_id != null ? row.role_id : null,
    role_name: row.role_name || null,
    name: row.name || null,
    modules: row.modules || null,
    password: row.password,
    id: row.id,
  };
}

export function getSharedUsers() {
  const saved = readAssignedUsers();
  const byName = new Map();
  DEFAULT_USERS.forEach((u) => byName.set(String(u.username).toLowerCase(), { ...u }));
  saved.forEach((u) => {
    if (!u || !u.username) return;
    const key = String(u.username).toLowerCase();
    const prev = byName.get(key) || {};
    const savedRole =
      u.role != null && String(u.role).trim() !== ''
        ? String(u.role).trim()
        : u.system_role != null && String(u.system_role).trim() !== ''
          ? String(u.system_role).trim()
          : null;
    byName.set(key, {
      ...prev,
      ...u,
      username: String(u.username).trim(),
      password:
        u.password != null && String(u.password).length
          ? String(u.password)
          : prev.password || '',
      // Assignment always wins over DEFAULT_USERS
      role: savedRole != null ? savedRole : prev.role || 'employee',
      system_role: savedRole != null ? savedRole : prev.system_role || prev.role || 'employee',
      role_id: u.role_id != null ? u.role_id : prev.role_id,
      role_name: u.role_name != null ? u.role_name : prev.role_name,
      name: u.name || u.employee_name || prev.name || u.username,
      modules: u.modules || u.allowed_modules || prev.modules || null,
    });
  });
  return Array.from(byName.values());
}

export function upsertSharedUsers(employees = [], { preserveRole = true } = {}) {
  const current = readAssignedUsers();
  const byName = new Map(
    (Array.isArray(current) ? current : []).map((u) => [
      String(u.username || '').toLowerCase(),
      u,
    ])
  );
  (employees || []).forEach((e) => {
    const username = String(e.username || e.user_name || '').trim();
    if (!username) return;
    const key = username.toLowerCase();
    const prev = byName.get(key) || {};

    const incomingRole =
      e.role != null && String(e.role).trim() !== ''
        ? String(e.role).trim()
        : e.system_role != null && String(e.system_role).trim() !== ''
          ? String(e.system_role).trim()
          : null;

    // Do NOT let HR employee sync wipe a role that was set in Role Assignment
    let nextRole = incomingRole != null ? incomingRole : prev.role || 'employee';
    if (
      preserveRole &&
      prev.role != null &&
      String(prev.role).trim() !== '' &&
      (prev.role_id != null || prev.role_name != null || prev.updated_at)
    ) {
      // Keep previous assigned role unless caller explicitly sends role_id (assignment)
      if (e.role_id == null && e.role_name == null && !e._forceRole) {
        nextRole = prev.role;
      }
    }
    if (e._forceRole && incomingRole != null) {
      nextRole = incomingRole;
    }

    const incomingMods = e.modules || e.allowed_modules || e.module_access;
    const modules = Array.isArray(incomingMods)
      ? incomingMods.map(String)
      : prev.modules || null;

    byName.set(key, {
      ...prev,
      username,
      password:
        e.password != null && String(e.password).length
          ? String(e.password)
          : prev.password || '',
      role: nextRole,
      system_role: nextRole,
      role_id:
        e.role_id != null
          ? e.role_id
          : e._forceRole
            ? e.role_id
            : prev.role_id,
      role_name:
        e.role_name != null
          ? e.role_name
          : e._forceRole
            ? e.role_name
            : prev.role_name,
      name: e.name || e.employee_name || prev.name || username,
      id: e.id || prev.id,
      modules,
      updated_at:
        e._forceRole || e.role_id != null
          ? new Date().toISOString()
          : prev.updated_at || null,
    });
  });
  const next = Array.from(byName.values());
  writeJson(USERS_KEY, next);
  try {
    if (typeof window !== 'undefined') {
      window.dispatchEvent(
        new CustomEvent('uu-shared-users-updated', { detail: { users: next } })
      );
    }
  } catch (_) {}
  return next;
}

export function upsertSharedUser(user) {
  if (!user || !(user.username || user.user_name)) return getSharedUsers();
  // Role Assignment and explicit role writes must force role
  const force =
    user._forceRole === true ||
    user.role_id != null ||
    user.role_name != null ||
    (user.role != null && String(user.role).trim() !== '');
  return upsertSharedUsers([{ ...user, _forceRole: force }], { preserveRole: !force });
}

export function getSharedRoles() {
  const roles = readJson(ROLES_KEY, []);
  return Array.isArray(roles) ? roles : [];
}

export function saveSharedRoles(roles = []) {
  const list = Array.isArray(roles) ? roles : [];
  writeJson(ROLES_KEY, list);
  return list;
}

export function mergeSharedRolesInto(existing = []) {
  const shared = getSharedRoles();
  if (!shared.length) return existing || [];
  const byId = new Map();
  (existing || []).forEach((r) => {
    const k = String(r.id || r.name || r.roleName || '');
    if (k) byId.set(k, r);
  });
  shared.forEach((r) => {
    const k = String(r.id || r.name || r.roleName || '');
    if (k) byId.set(k, { ...byId.get(k), ...r });
  });
  return Array.from(byId.values());
}

/** Full session snapshot — role must never be stripped on module switch */
export function saveSharedSession(user) {
  if (!user || !user.username) {
    clearSharedSession();
    return null;
  }
  const assigned = resolveAssignedRole(user.username);
  const catalog = getSharedUsers().find(
    (u) => String(u.username).toLowerCase() === String(user.username).toLowerCase()
  );
  const role =
    (assigned && assigned.role) ||
    user.role ||
    catalog?.role ||
    'employee';
  const snapshot = {
    username: user.username,
    name: user.name || assigned?.name || catalog?.name || user.username,
    role,
    system_role: role,
    role_id: (assigned && assigned.role_id) || user.role_id || catalog?.role_id || null,
    role_name: (assigned && assigned.role_name) || user.role_name || catalog?.role_name || null,
    id: user.id || catalog?.id || null,
    modules: user.modules || user.allowed_modules || assigned?.modules || catalog?.modules || null,
    at: Date.now(),
  };
  writeJson(SESSION_KEY, snapshot);
  return snapshot;
}

export function loadSharedSession() {
  const s = readJson(SESSION_KEY, null);
  if (!s || !s.username) return null;
  const assigned = resolveAssignedRole(s.username);
  const catalog = getSharedUsers().find(
    (u) => String(u.username).toLowerCase() === String(s.username).toLowerCase()
  );
  const role =
    (assigned && assigned.role) || s.role || catalog?.role || 'employee';
  return {
    ...s,
    role,
    system_role: role,
    role_id: (assigned && assigned.role_id) || s.role_id || catalog?.role_id || null,
    role_name: (assigned && assigned.role_name) || s.role_name || catalog?.role_name || null,
    name: s.name || (assigned && assigned.name) || catalog?.name || s.username,
    modules: s.modules || (assigned && assigned.modules) || catalog?.modules || null,
  };
}

export function clearSharedSession() {
  try {
    localStorage.removeItem(SESSION_KEY);
  } catch (_) {}
}

export function canAccessModule(userOrRole, moduleId, userObj) {
  const roleRaw =
    typeof userOrRole === 'string'
      ? userOrRole
      : userOrRole?.role || userOrRole?.system_role || userObj?.role || userObj?.system_role;
  const role = normalizeRole(roleRaw);
  const list = MODULE_ACCESS[moduleId] || [];
  if (role === 'admin' || String(userObj?.username || userOrRole?.username || '').toLowerCase() === 'admin') {
    return true;
  }
  // Custom ROLE_CRUD ids (numeric / uuid) — allow if user has any session role
  if (roleRaw && !list.map(normalizeRole).includes(role)) {
    // Still allow known module keys by role_name slug if present on user
    const slug = normalizeRole(userObj?.role_name || userOrRole?.role_name || '');
    if (slug && list.map(normalizeRole).includes(slug)) return true;
    // Assigned custom roles: allow module if modules[] lists it
    const mods = userObj?.modules || userOrRole?.modules;
    if (Array.isArray(mods) && mods.map(String).map((m) => m.toLowerCase()).includes(String(moduleId).toLowerCase())) {
      return true;
    }
  }
  return list.map(normalizeRole).includes(role);
}

export function moduleAccessMessage(moduleId) {
  const labels = {
    finance: 'Finance',
    hr: 'Human Resources (HR)',
    material: 'Material / Inventory',
  };
  const name = labels[moduleId] || String(moduleId || 'this module');
  const allowed = (MODULE_ACCESS[moduleId] || []).join(', ');
  return allowed
    ? `${name} is limited to roles: ${allowed}. Your account role is not on that list — ask an administrator to assign a suitable role.`
    : `You do not have permission to open ${name}. Contact your administrator.`;
}

/**
 * Login: password may match DEFAULT or assigned user.
 * Role ALWAYS comes from resolveAssignedRole when present.
 */
export function attemptSharedLogin(username, password, extraEmployees = []) {
  const uName = String(username || '').trim();
  const uPass = String(password || '');
  if (!uName || !uPass) {
    return { success: false, message: 'Username and password required.' };
  }

  // Sync passwords from employees only — never clobber assigned roles
  if (extraEmployees && extraEmployees.length) {
    const withCreds = extraEmployees.filter(
      (e) =>
        e &&
        (e.username || e.user_name) &&
        e.password != null &&
        String(e.password).length
    );
    if (withCreds.length) {
      upsertSharedUsers(
        withCreds.map((e) => ({
          username: e.username || e.user_name,
          password: e.password,
          name: e.name || e.employee_name,
          id: e.id,
          // omit role so preserveRole keeps assignment
        })),
        { preserveRole: true }
      );
    }
  }

  const pool = getSharedUsers();
  const found = pool.find(
    (u) =>
      u &&
      String(u.username || '').trim().toLowerCase() === uName.toLowerCase() &&
      String(u.password ?? '') === uPass
  );

  if (!found) {
    return { success: false, message: 'Invalid username or password.' };
  }

  const assigned = resolveAssignedRole(uName);
  const roleFinal = assigned
    ? String(assigned.role).trim()
    : String(found.role || found.system_role || 'employee').trim();

  const user = {
    username: found.username || uName,
    name: (assigned && assigned.name) || found.name || uName,
    role: roleFinal,
    system_role: roleFinal,
    role_id: (assigned && assigned.role_id) || found.role_id || null,
    role_name: (assigned && assigned.role_name) || found.role_name || null,
    id: found.id,
    modules: (assigned && assigned.modules) || found.modules || null,
  };

  // Persist session with assigned role
  saveSharedSession(user);
  return { success: true, user };
}



/**
 * Change password for a shared login user (Finance / HR / Material).
 * Verifies against getSharedUsers (includes DEFAULT + assignments).
 */
export function changeSharedPassword(username, oldPassword, newPassword) {
  const uName = String(username || '').trim();
  const oldP = String(oldPassword ?? '');
  const newP = String(newPassword ?? '').trim();
  if (!uName) return { success: false, message: 'Not signed in.' };
  if (!oldP) return { success: false, message: 'Current password is required.' };
  if (!newP || newP.length < 4) {
    return { success: false, message: 'New password must be at least 4 characters.' };
  }
  if (oldP === newP) {
    return { success: false, message: 'New password must be different from the current password.' };
  }

  // Prefer assigned catalog (raw), then merged getSharedUsers
  const assignedList = typeof readAssignedUsers === 'function' ? readAssignedUsers() : [];
  const merged = getSharedUsers();
  let found =
    (assignedList || []).find(
      (u) => String(u.username || '').trim().toLowerCase() === uName.toLowerCase()
    ) ||
    (merged || []).find(
      (u) => String(u.username || '').trim().toLowerCase() === uName.toLowerCase()
    );

  if (!found) {
    return { success: false, message: 'User account not found.' };
  }

  const stored = String(found.password ?? '');
  // Empty stored password: allow change if user is changing from whatever they type
  // only when stored is empty AND they have an active session (caller already logged in)
  if (stored !== '' && stored !== oldP && stored.trim() !== oldP.trim()) {
    return { success: false, message: 'Current password incorrect.' };
  }
  if (stored === '' && !oldP) {
    return { success: false, message: 'Current password is required.' };
  }

  upsertSharedUser({
    username: found.username,
    password: newP,
    name: found.name,
    role: found.role,
    system_role: found.system_role || found.role,
    role_id: found.role_id,
    role_name: found.role_name,
    modules: found.modules,
    id: found.id,
    _forceRole: true,
  });

  // Verify write
  const check = getSharedUsers().find(
    (u) => String(u.username || '').trim().toLowerCase() === uName.toLowerCase()
  );
  if (!check || String(check.password ?? '') !== newP) {
    return { success: false, message: 'Failed to save new password. Try again.' };
  }

  return { success: true, message: 'Password updated successfully. Please sign in again.' };
}


export { DEFAULT_USERS, USERS_KEY, ROLES_KEY, SESSION_KEY };
