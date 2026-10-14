import React, { useState, useEffect } from 'react';
import { MobileStyles } from './unity/mobileCompat.js';
import AuthProviderDefault, { AuthProvider as AuthProviderNamed } from './unity/SharedAuth.jsx';
import UIProviderDefault, { UIProvider as UIProviderNamed } from './unity/SharedUI.jsx';
import { SharedDataProvider } from './unity/SharedDataProvider.jsx';
import { FinanceApp } from './modules/finance/FinanceApp.jsx';
import { HRApp } from './modules/hr/HRApp.jsx';
import { MaterialApp } from './modules/material/MaterialApp.jsx';

const AuthProvider = AuthProviderNamed || AuthProviderDefault;
const UIProvider = UIProviderNamed || UIProviderDefault;

function readModule() {
  const id = (localStorage.getItem('uu_module') || 'finance').toLowerCase();
  if (id === 'finance' || id === 'hr' || id === 'material') return id;
  return 'finance';
}

function detailToModule(detail) {
  if (detail == null) return null;
  if (typeof detail === 'string') return detail.toLowerCase();
  if (typeof detail === 'object' && detail.module) return String(detail.module).toLowerCase();
  return null;
}

export default function App() {
  const [module, setModule] = useState(readModule);

  useEffect(() => {
    const onChange = (e) => {
      const fromEvent = detailToModule(e?.detail);
      const id = fromEvent || readModule();
      if (id === 'finance' || id === 'hr' || id === 'material') {
        setModule(id);
      }
    };
    window.addEventListener('uu-module-change', onChange);
    window.addEventListener('uu_module_change', onChange);
    // storage event (other tabs)
    const onStorage = (e) => {
      if (e.key === 'uu_module') onChange({ detail: e.newValue });
    };
    window.addEventListener('storage', onStorage);
    return () => {
      window.removeEventListener('uu-module-change', onChange);
      window.removeEventListener('uu_module_change', onChange);
      window.removeEventListener('storage', onStorage);
    };
  }, []);

  return (
    <AuthProvider>
      <MobileStyles />
      <UIProvider>
        <SharedDataProvider>
          {/* No top Finance/HR/Material bar — switch only on each module login form */}
          <div className="h-screen flex flex-col overflow-hidden bg-slate-950 text-white">
            <div className="flex-1 min-h-0 overflow-hidden">
              {module === 'finance' && <FinanceApp />}
              {module === 'hr' && <HRApp />}
              {module === 'material' && <MaterialApp />}
            </div>
          </div>
        </SharedDataProvider>
      </UIProvider>
    </AuthProvider>
  );
}
