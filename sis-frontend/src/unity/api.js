/**
 * Drop into sis-frontend/src/unity/api.js (or src/api.js)
 * Usage:
 *   import { api, login, bootstrap } from './api';
 */
const BASE = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';

function token() {
  return localStorage.getItem('uu_token') || '';
}

export async function api(path, options = {}) {
  const headers = {
    'Content-Type': 'application/json',
    ...(options.headers || {}),
  };
  const t = token();
  if (t) headers.Authorization = `Bearer ${t}`;

  const res = await fetch(`${BASE}${path}`, {
    ...options,
    headers,
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = data.message || data.error || res.statusText || 'Request failed';
    console.error('[SIS API]', res.status, path, msg, data);
    const err = new Error(`${res.status} ${msg}`);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

export async function login(username, password) {
  const data = await api('/auth/login', {
    method: 'POST',
    body: { username, password },
  });
  if (data.token) localStorage.setItem('uu_token', data.token);
  if (data.user) localStorage.setItem('uu_user', JSON.stringify(data.user));
  return data;
}

export function logout() {
  localStorage.removeItem('uu_token');
  localStorage.removeItem('uu_user');
}

export async function bootstrap() {
  return api('/bootstrap');
}

export async function listCollection(name) {
  const r = await api(`/collections/${encodeURIComponent(name)}`);
  return r.data || [];
}

export async function createRecord(name, body) {
  const r = await api(`/collections/${encodeURIComponent(name)}`, { method: 'POST', body });
  return r.data;
}

export async function updateRecord(name, id, body) {
  const r = await api(`/collections/${encodeURIComponent(name)}/${encodeURIComponent(id)}`, {
    method: 'PUT',
    body,
  });
  return r.data;
}

export async function deleteRecord(name, id) {
  return api(`/collections/${encodeURIComponent(name)}/${encodeURIComponent(id)}`, {
    method: 'DELETE',
  });
}

export async function postMaterialToFinance(source, recordId) {
  return api('/integration/material/post', {
    method: 'POST',
    body: { source, recordId },
  });
}

export async function syncHrPayMaster() {
  return api('/integration/hr/sync-pay-master', { method: 'POST', body: {} });
}

export async function postPayroll(ids = []) {
  return api('/integration/payroll/post', { method: 'POST', body: { ids } });
}
