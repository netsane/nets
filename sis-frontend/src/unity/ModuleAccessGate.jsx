import React from 'react';
import { canAccessModule, moduleAccessMessage, clearSharedSession } from './SharedCredentials.js';
import LoginModuleSwitcher from './LoginModuleSwitcher.jsx';

/**
 * Blocks dashboard when role cannot access this module.
 * Does NOT clear session on view — only Logout button does.
 */
export default function ModuleAccessGate({ moduleId, user, children, onLogout }) {
  if (!user) return children || null;

  if (canAccessModule(user.role, moduleId, user)) {
    return children || null;
  }

  return (
    <div className="h-full w-full min-h-0 flex items-center justify-center bg-slate-950 p-6">
      <div className="max-w-md w-full rounded-3xl border border-rose-500/30 bg-slate-900/90 p-8 text-center shadow-2xl">
        <p className="text-[10px] font-black uppercase tracking-widest text-rose-400 mb-2">
          Access restricted
        </p>
        <h2 className="text-xl font-black text-white mb-2">No permission for this module</h2>
        <p className="text-sm text-gray-400 mb-1">
          Signed in as <span className="text-cyan-400 font-bold">{user.username}</span>
          {' · '}
          <span className="text-amber-400 uppercase text-xs font-black">{user.role}</span>
        </p>
        <p className="text-sm text-gray-500 mb-6">{moduleAccessMessage(moduleId)}</p>
        <p className="text-[10px] text-gray-600 mb-4">
          Your session stays active. Open a module your role allows, or log out.
        </p>
        <div className="flex flex-col gap-3">
          <LoginModuleSwitcher current={moduleId} />
          <button
            type="button"
            onClick={() => {
              clearSharedSession();
              if (typeof onLogout === 'function') onLogout();
              else window.location.reload();
            }}
            className="w-full py-3 rounded-xl text-xs font-black uppercase bg-rose-600 hover:bg-rose-500 text-white"
          >
            Log out
          </button>
        </div>
      </div>
    </div>
  );
}
