/**
 * Example App.jsx — put in frontend/src/App.jsx
 * Adjust import paths to your modules once split.
 */
import React, { useState } from 'react';
import { SharedDataProvider } from './unity/SharedDataProvider';
import { UnityIntegrationBus } from './unity/IntegrationBus';

// After you split pastes:
// import { FinanceApp } from './modules/finance/FinanceApp';
// import { HRApp } from './modules/hr/HRApp';
// import { MaterialApp } from './modules/material/MaterialApp';

export default function App() {
  const [module, setModule] = useState(() => UnityIntegrationBus.getActiveModule());

  const select = (id) => {
    UnityIntegrationBus.setActiveModule(id);
    setModule(id);
  };

  return (
    <SharedDataProvider>
      <div className="h-screen flex flex-col bg-slate-950 text-white">
        <header className="h-12 flex items-center gap-2 px-4 border-b border-white/10">
          <span className="font-black italic text-cyan-400 mr-3">UNITY SIS</span>
          {[
            { id: 'finance', label: 'Finance' },
            { id: 'hr', label: 'HR' },
            { id: 'material', label: 'Material' },
          ].map((m) => (
            <button
              key={m.id}
              type="button"
              onClick={() => select(m.id)}
              className={`px-3 py-1 rounded-lg text-xs font-bold uppercase ${
                module === m.id ? 'bg-cyan-600' : 'text-gray-400 hover:bg-white/10'
              }`}
            >
              {m.label}
            </button>
          ))}
        </header>
        <main className="flex-1 min-h-0 overflow-auto p-4">
          <p className="text-sm text-gray-400 mb-4">
            Active: <strong className="text-cyan-300">{module}</strong> — wire FinanceApp / HRApp / MaterialApp here.
          </p>
          {/* {module === 'finance' && <FinanceApp />} */}
          {/* {module === 'hr' && <HRApp />} */}
          {/* {module === 'material' && <MaterialApp />} */}
        </main>
      </div>
    </SharedDataProvider>
  );
}
