import React from 'react';

/**
 * Module tabs on login forms only.
 * Switches uu_module so App.jsx mounts the correct module (login stays until user signs in).
 */
export default function LoginModuleSwitcher({ current = '' }) {
  const go = (id) => {
    if (!id || !['finance', 'hr', 'material'].includes(id)) return;
    localStorage.setItem('uu_module', id);
    // Support both event names / detail shapes used across versions
    try {
      window.dispatchEvent(new CustomEvent('uu-module-change', { detail: id }));
      window.dispatchEvent(new CustomEvent('uu-module-change', { detail: { module: id } }));
      window.dispatchEvent(new CustomEvent('uu_module_change', { detail: id }));
      window.dispatchEvent(new CustomEvent('uu_module_change', { detail: { module: id } }));
    } catch (e) { /* ignore */ }
  };

  const active = String(current || localStorage.getItem('uu_module') || '').toLowerCase();

  const btn = (id, label, onCls, offCls) => (
    <button
      key={id}
      type="button"
      onClick={(e) => {
        e.preventDefault();
        e.stopPropagation();
        go(id);
      }}
      className={
        'py-3 rounded-xl text-[11px] font-black uppercase tracking-wide transition ' +
        (active === id ? onCls : offCls)
      }
    >
      {label}
    </button>
  );

  return (
    <div className="mt-5 w-full relative z-[100]" style={{ pointerEvents: 'auto' }}>
      <p className="text-[10px] font-black uppercase tracking-widest text-cyan-400 text-center mb-2">
        Switch module (then sign in)
      </p>
      <div className="grid grid-cols-3 gap-2 w-full">
        {btn('finance', 'Finance', 'bg-cyan-500 text-white', 'bg-cyan-900/80 text-cyan-100 hover:bg-cyan-600')}
        {btn('hr', 'HR', 'bg-violet-500 text-white', 'bg-violet-900/80 text-violet-100 hover:bg-violet-600')}
        {btn('material', 'Material', 'bg-emerald-500 text-white', 'bg-emerald-900/80 text-emerald-100 hover:bg-emerald-600')}
      </div>
    </div>
  );
}
