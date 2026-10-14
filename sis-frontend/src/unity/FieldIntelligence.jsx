/**
 * Unity SIS — Field Intelligence (robust)
 * Hover / click the cyan (i) next to labels & headers for guidance.
 */
import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import { createPortal } from 'react-dom';
import * as Icons from 'lucide-react';

const STORAGE_KEY = 'uu_field_intelligence_v1';

export function loadIntelPrefs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { enabled: true, delayMs: 200, dense: false };
    const parsed = JSON.parse(raw);
    return {
      enabled: parsed.enabled !== false, // default ON
      delayMs: Number(parsed.delayMs) >= 0 ? Number(parsed.delayMs) : 200,
      dense: !!parsed.dense,
    };
  } catch {
    return { enabled: true, delayMs: 200, dense: false };
  }
}

export function saveIntelPrefs(partial = {}) {
  const next = { ...loadIntelPrefs(), ...partial };
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch (_) {}
  try {
    window.dispatchEvent(new CustomEvent('uu-intel-prefs', { detail: next }));
  } catch (_) {}
  return next;
}

export function useIntelPrefs() {
  const [prefs, setPrefs] = useState(() => loadIntelPrefs());
  useEffect(() => {
    const sync = (e) => setPrefs(e?.detail || loadIntelPrefs());
    window.addEventListener('uu-intel-prefs', sync);
    window.addEventListener('storage', sync);
    return () => {
      window.removeEventListener('uu-intel-prefs', sync);
      window.removeEventListener('storage', sync);
    };
  }, []);
  const update = useCallback((partial) => {
    const next = saveIntelPrefs(partial);
    setPrefs(next);
    return next;
  }, []);
  return [prefs, update];
}

export function humanizeKey(key) {
  return String(key || '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

export const INTEL_CATALOG = {
  ui_customization: {
    title: 'UI Customization',
    body: 'Theme, logo, background, and Field Intelligence. Settings are stored in this browser.',
    tips: ['Enable Field Intelligence to see (i) tips on forms'],
  },
  field_intelligence: {
    title: 'Field Intelligence',
    body: 'Cyan (i) icons explain fields, filters, and how HR, Finance, and Material share data.',
    tips: ['Hover or tap the (i)', 'Turn off anytime in UI Customization'],
  },
  export_options: {
    title: 'Export Options',
    body: 'Excel, Word, PDF, CSV, or PNG of the current view. Official Unity headers apply on formal exports.',
  },
  search: {
    title: 'Search',
    body: 'Filters the table by text across columns. Clear to show all rows again.',
  },
  status: {
    title: 'Status',
    body: 'Workflow state (Draft, Posted, Approved, Pending…). Affects posting and reports.',
  },
  employee_name: {
    title: 'Employee (HR Master)',
    body: 'Staff from the HR employee register — source of truth for payroll, competence, and leave.',
    tips: ['Update the person in HR if the name or ID is wrong'],
  },
  employee_id: {
    title: 'Employee ID',
    body: 'Institutional ID (e.g. UU-001) linking payroll, appraisal, and HR records.',
  },
  id_number: {
    title: 'ID Number',
    body: 'Primary staff number. Must be unique across the university.',
  },
  department: {
    title: 'Department',
    body: 'Organizational unit used in filters, dashboards, and cost context.',
  },
  basic_salary: {
    title: 'Basic Salary',
    body: 'Core monthly pay. Feeds payroll process, tax, pension, and Finance expense when posted.',
  },
  payroll_period: {
    title: 'Payroll Period',
    body: 'Processing month (YYYY-MM). Reconciliation compares prior vs current period.',
  },
  account_id: {
    title: 'Account (COA)',
    body: 'Chart of Accounts code. Postings here feed Trial Balance and General Ledger.',
  },
  account_no: {
    title: 'Account Number',
    body: 'Official COA key in ####-##-## form.',
  },
  partner_name: {
    title: 'Partner',
    body: 'Supplier, customer, bank, or agency from the Partners register.',
  },
  debit: {
    title: 'Debit',
    body: 'Debit amount. Journals must balance with equal credits.',
  },
  credit: {
    title: 'Credit',
    body: 'Credit amount. Balance every entry.',
  },
  narration: {
    title: 'Narration',
    body: 'Business description for audit and GL readability.',
  },
  skill: {
    title: 'Skill / Competency',
    body: 'One assessed competency (score 1–5) on the competence matrix.',
  },
  score: {
    title: 'Score (1–5)',
    body: '1 Novice · 2 Developing · 3 Competent · 4 Proficient · 5 Expert.',
  },
  appraisal_score: {
    title: 'Appraisal Score',
    body: 'Overall score, often from the competence matrix average for that employee.',
  },
  sku: {
    title: 'SKU',
    body: 'Stock keeping unit code for inventory, GRV, and issues.',
  },
  qty_received: {
    title: 'Qty Received',
    body: 'Quantity accepted on GRV. Can post into Finance when integrated.',
  },
  unit_cost: {
    title: 'Unit Cost (ETB)',
    body: 'Cost per unit in Birr. Line value = qty × unit cost.',
  },
  grv_number: {
    title: 'GRV Number',
    body: 'Goods received voucher reference for receiving and Finance postings.',
  },
};

export function resolveIntel({ module, collection, fieldKey, label, explicit } = {}) {
  if (explicit) {
    if (typeof explicit === 'string') {
      return { title: label || humanizeKey(fieldKey), body: explicit, tips: [] };
    }
    return {
      title: explicit.title || label || humanizeKey(fieldKey),
      body: explicit.body || explicit.text || '',
      tips: explicit.tips || [],
    };
  }
  const key = String(fieldKey || '').trim();
  const labelKey = String(label || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '');
  const candidates = [
    key,
    labelKey,
    key.toLowerCase(),
    // strip parenthetical labels e.g. "Basic Salary (HR)"
    key.replace(/\s*\([^)]*\)/g, '').trim(),
  ].filter(Boolean);

  // substring catalog match on label text
  const labelLower = String(label || key || '').toLowerCase();
  const fuzzy = [
    ['employee', 'employee_name'],
    ['basic salary', 'basic_salary'],
    ['payroll period', 'payroll_period'],
    ['account', 'account_id'],
    ['partner', 'partner_name'],
    ['debit', 'debit'],
    ['credit', 'credit'],
    ['sku', 'sku'],
    ['search', 'search'],
    ['status', 'status'],
  ];
  for (const [needle, cat] of fuzzy) {
    if (labelLower.includes(needle) && INTEL_CATALOG[cat]) {
      const hit = INTEL_CATALOG[cat];
      return { title: hit.title, body: hit.body, tips: hit.tips || [] };
    }
  }

  for (const c of candidates) {
    if (INTEL_CATALOG[c]) {
      const hit = INTEL_CATALOG[c];
      return { title: hit.title || label || humanizeKey(key), body: hit.body, tips: hit.tips || [] };
    }
  }

  const title = label || humanizeKey(key) || 'Field help';
  const mod = module ? String(module).toUpperCase() : 'Unity SIS';
  return {
    title,
    body: `${title} belongs to the ${mod}${collection ? ` · ${humanizeKey(collection)}` : ''} workspace. Enter accurate data — HR, Finance, and Material may share this value.`,
    tips: ['Hover other cyan (i) icons for more help'],
  };
}

export function IntelTip({
  module,
  collection,
  fieldKey,
  label,
  help,
  className = '',
  children,
  forceShow = false,
}) {
  const [prefs] = useIntelPrefs();
  const enabled = forceShow || prefs.enabled !== false;
  const [open, setOpen] = useState(false);
  const [box, setBox] = useState(null);
  const anchorRef = useRef(null);
  const timerRef = useRef(null);

  const intel = useMemo(
    () => resolveIntel({ module, collection, fieldKey, label, explicit: help }),
    [module, collection, fieldKey, label, help]
  );

  const place = useCallback(() => {
    const el = anchorRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const width = Math.min(340, Math.max(248, window.innerWidth - 24));
    const estHeight = prefs.dense ? 100 : 150;
    const gap = 8;
    let left = r.left + r.width / 2 - width / 2;
    if (left < 8) left = 8;
    if (left + width > window.innerWidth - 8) left = window.innerWidth - width - 8;
    const spaceBelow = window.innerHeight - r.bottom - gap;
    const openUp = spaceBelow < estHeight && r.top > spaceBelow;
    const top = openUp
      ? Math.max(8, r.top - gap - estHeight)
      : Math.min(window.innerHeight - 24, r.bottom + gap);
    setBox({
      top,
      left,
      width,
      maxHeight: Math.min(280, openUp ? Math.max(96, r.top - gap) : Math.max(96, spaceBelow)),
    });
  }, [prefs.dense]);

  const show = useCallback(() => {
    if (!enabled) return;
    clearTimeout(timerRef.current);
    timerRef.current = setTimeout(() => {
      place();
      setOpen(true);
    }, Math.min(Math.max(prefs.delayMs || 200, 0), 800));
  }, [enabled, prefs.delayMs, place]);

  const hide = useCallback(() => {
    clearTimeout(timerRef.current);
    setOpen(false);
  }, []);

  useEffect(() => () => clearTimeout(timerRef.current), []);

  useEffect(() => {
    if (!open) return undefined;
    place();
    const onRe = () => place();
    window.addEventListener('scroll', onRe, true);
    window.addEventListener('resize', onRe);
    return () => {
      window.removeEventListener('scroll', onRe, true);
      window.removeEventListener('resize', onRe);
    };
  }, [open, place]);

  if (!enabled) {
    return children ? <span className={className}>{children}</span> : null;
  }

  const InfoIcon = (Icons && Icons.Info) ? Icons.Info : null;

  const bubble =
    open && box && typeof document !== 'undefined'
      ? createPortal(
          <div
            role="tooltip"
            data-uu-intel-tip="1"
            className="rounded-2xl border border-white/10 bg-gray-900/90 text-white shadow-2xl shadow-black/50 backdrop-blur-md"
            style={{
              position: 'fixed',
              top: box.top,
              left: box.left,
              width: box.width,
              maxHeight: box.maxHeight,
              zIndex: 2147483647,
              padding: prefs.dense ? '10px 12px' : '12px 14px',
              pointerEvents: 'none',
              boxShadow: '0 12px 40px rgba(0,0,0,0.55), 0 0 0 1px rgba(34,211,238,0.12)',
            }}
          >
            <div className="flex items-center gap-2 mb-1.5">
              <span
                className="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded-lg
                           bg-cyan-500/15 border border-cyan-400/25 text-cyan-300"
              >
                {InfoIcon ? <InfoIcon className="h-3 w-3" strokeWidth={2.5} /> : (
                  <span className="text-[9px] font-black">i</span>
                )}
              </span>
              <div className="text-[10px] font-black uppercase tracking-[0.14em] text-cyan-300/95 leading-none">
                {intel.title}
              </div>
            </div>
            <p className="text-[11px] leading-relaxed text-white/75 pl-0.5">{intel.body}</p>
            {intel.tips?.length > 0 && (
              <ul className="mt-2 space-y-1 border-t border-white/5 pt-2">
                {intel.tips.map((tip, idx) => (
                  <li key={idx} className="text-[10px] text-cyan-100/70 flex gap-1.5 leading-snug">
                    <span className="text-cyan-500/80 shrink-0">·</span>
                    <span>{tip}</span>
                  </li>
                ))}
              </ul>
            )}
          </div>,
          document.body
        )
      : null;

  return (
    <span
      ref={anchorRef}
      className={`inline-flex items-center gap-1 align-middle ${className}`}
      onMouseEnter={show}
      onMouseLeave={hide}
      onFocus={show}
      onBlur={hide}
      data-uu-intel-anchor="1"
    >
      {children}
      <button
        type="button"
        tabIndex={0}
        aria-label={`Help: ${intel.title}`}
        title={intel.title}
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (open) hide();
          else {
            place();
            setOpen(true);
          }
        }}
        className={`group/intel relative inline-flex h-[15px] w-[15px] shrink-0 items-center justify-center
                    rounded-md border border-white/10 bg-white/[0.04]
                    text-white/35 hover:text-cyan-300 hover:border-cyan-400/40 hover:bg-cyan-500/10
                    focus:outline-none focus-visible:ring-1 focus-visible:ring-cyan-400/50
                    transition-all duration-150
                    ${open ? 'text-cyan-300 border-cyan-400/45 bg-cyan-500/15' : ''}`}
      >
        {InfoIcon ? (
          <InfoIcon className="h-[9px] w-[9px]" strokeWidth={2.5} />
        ) : (
          <span className="text-[8px] font-black leading-none">i</span>
        )}
      </button>
      {bubble}
    </span>
  );
}

export function IntelLabel({ label, isRequired, module, collection, fieldKey, help, className = '' }) {
  return (
    <span className={`inline-flex items-center gap-1.5 min-w-0 max-w-full ${className}`}>
      <IntelTip module={module} collection={collection} fieldKey={fieldKey || label} label={label} help={help}>
        <span className="truncate text-inherit">
          {label}
          {isRequired ? <span className="text-rose-400 ml-0.5">*</span> : null}
        </span>
      </IntelTip>
    </span>
  );
}

export function FieldIntelligenceSettings({ glassClass = '' }) {
  const [prefs, update] = useIntelPrefs();
  return (
    <div className={`rounded-2xl border border-cyan-500/40 p-4 sm:p-5 ${glassClass} bg-slate-950/80`}>
      <div className="flex items-start gap-3">
        <span className="p-2.5 rounded-2xl bg-cyan-500/10 border border-cyan-400/25 text-cyan-300 flex items-center justify-center">
          {(Icons && Icons.Info) ? <Icons.Info className="h-4 w-4" strokeWidth={2.25} /> : <span className="text-xs font-black">i</span>}
        </span>
        <div className="min-w-0 flex-1">
          <h4 className="text-sm font-black uppercase tracking-wide text-cyan-300">Field Intelligence</h4>
          <p className="mt-1 text-[12px] text-gray-300 leading-relaxed">
            Look for the cyan <strong className="text-cyan-300">(i)</strong> beside form labels and table headers.
            Hover or tap it for a short explanation. Works in Finance (including payroll), HR, and Material.
          </p>
          <div className="mt-4 flex flex-col sm:flex-row sm:items-center gap-3">
            <label className="inline-flex items-center gap-2 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={prefs.enabled !== false}
                onChange={(e) => update({ enabled: e.target.checked })}
                className="h-4 w-4 rounded border-cyan-500/50 bg-black accent-cyan-500"
              />
              <span className="text-xs font-bold text-white uppercase tracking-wide">Show intelligence tips</span>
            </label>
            <label className="inline-flex items-center gap-2 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={!!prefs.dense}
                onChange={(e) => update({ dense: e.target.checked })}
                className="h-4 w-4 rounded border-cyan-500/50 bg-black accent-cyan-500"
              />
              <span className="text-xs font-bold text-gray-200 uppercase tracking-wide">Compact tips</span>
            </label>
          </div>
          <div className="mt-3 flex items-center gap-2">
            <span className="text-[10px] font-black uppercase text-gray-500">Hover delay</span>
            <input
              type="range"
              min={0}
              max={600}
              step={50}
              value={prefs.delayMs ?? 200}
              onChange={(e) => update({ delayMs: Number(e.target.value) })}
              className="flex-1 accent-cyan-500"
            />
            <span className="text-[11px] font-mono text-cyan-300 w-12 text-right">{prefs.delayMs ?? 200}ms</span>
          </div>
          <div className="mt-4 flex flex-wrap items-center gap-3">
            <IntelTip fieldKey="field_intelligence" label="Field Intelligence" forceShow>
              <span className="text-[11px] text-cyan-300 underline">Preview tip (hover the i)</span>
            </IntelTip>
            <button
              type="button"
              className="text-[10px] font-bold uppercase tracking-wide text-gray-400 hover:text-white"
              onClick={() => {
                update({ enabled: true, delayMs: 150, dense: false });
                try {
                  localStorage.setItem(STORAGE_KEY, JSON.stringify({ enabled: true, delayMs: 150, dense: false }));
                } catch (_) {}
              }}
            >
              Reset to defaults
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}



/**
 * Global layer — works inside modules after login (tables, filters, native labels),
 * not only on shared <Input> used by the login form.
 */
export function FieldIntelligenceRuntime() {
  const [prefs] = useIntelPrefs();
  const [open, setOpen] = useState(false);
  const [box, setBox] = useState(null);
  const [intel, setIntel] = useState(null);
  const hideTimer = useRef(null);
  const lastEl = useRef(null);

  useEffect(() => {
    if (prefs.enabled === false) {
      setOpen(false);
      return undefined;
    }

    const interesting = (el) => {
      if (!el || el.nodeType !== 1) return null;
      const hit = el.closest(
        'label, th, [data-uu-intel], [data-field-key], .uu-intel-target, legend'
      );
      if (!hit) return null;
      if (hit.closest('[data-uu-intel-tip="1"]')) return null;
      // skip pure layout chrome without text
      const text = (hit.getAttribute('data-uu-intel') || hit.getAttribute('data-field-key') || hit.textContent || '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!text || text.length < 2 || text.length > 120) return null;
      return hit;
    };

    const showFor = (hit) => {
      if (!hit) return;
      lastEl.current = hit;
      const key =
        hit.getAttribute('data-field-key') ||
        hit.getAttribute('data-uu-intel') ||
        hit.getAttribute('data-key') ||
        '';
      const label = (hit.textContent || key || 'Help').replace(/\s+/g, ' ').trim().slice(0, 80);
      const resolved = resolveIntel({ fieldKey: key || label, label });
      setIntel(resolved);
      const r = hit.getBoundingClientRect();
      const width = Math.min(360, Math.max(260, window.innerWidth - 24));
      const estHeight = prefs.dense ? 110 : 160;
      const gap = 10;
      let left = r.left + r.width / 2 - width / 2;
      if (left < 8) left = 8;
      if (left + width > window.innerWidth - 8) left = window.innerWidth - width - 8;
      const spaceBelow = window.innerHeight - r.bottom - gap;
      const openUp = spaceBelow < estHeight && r.top > spaceBelow;
      const top = openUp
        ? Math.max(8, r.top - gap - estHeight)
        : Math.min(window.innerHeight - 24, r.bottom + gap);
      setBox({
        top,
        left,
        width,
        maxHeight: Math.min(300, openUp ? Math.max(100, r.top - gap) : Math.max(100, spaceBelow)),
      });
      setOpen(true);
    };

    const onMove = (e) => {
      const hit = interesting(e.target);
      if (!hit) return;
      clearTimeout(hideTimer.current);
      // slight delay so it does not flash while moving across the UI
      clearTimeout(hit._uuIntelShowTimer);
      hit._uuIntelShowTimer = setTimeout(() => showFor(hit), Math.min(prefs.delayMs ?? 200, 400));
    };

    const onOut = (e) => {
      const to = e.relatedTarget;
      if (to && lastEl.current && lastEl.current.contains(to)) return;
      if (to && to.closest && to.closest('[data-uu-intel-tip="1"]')) return;
      clearTimeout(hideTimer.current);
      hideTimer.current = setTimeout(() => setOpen(false), 120);
    };

    const onClick = (e) => {
      const hit = interesting(e.target);
      if (hit && (e.target.closest('[data-uu-intel-anchor]') || e.target.textContent === 'i')) {
        showFor(hit);
      }
    };

    document.addEventListener('mouseover', onMove, true);
    document.addEventListener('mouseout', onOut, true);
    document.addEventListener('focusin', onMove, true);
    document.addEventListener('click', onClick, true);
    return () => {
      document.removeEventListener('mouseover', onMove, true);
      document.removeEventListener('mouseout', onOut, true);
      document.removeEventListener('focusin', onMove, true);
      document.removeEventListener('click', onClick, true);
      clearTimeout(hideTimer.current);
    };
  }, [prefs.enabled, prefs.delayMs, prefs.dense]);

  if (prefs.enabled === false || !open || !box || !intel) return null;
  if (typeof document === 'undefined') return null;

  const InfoIcon = (Icons && Icons.Info) ? Icons.Info : null;

  return createPortal(
    <div
      role="tooltip"
      data-uu-intel-tip="1"
      className="rounded-2xl border border-white/10 bg-gray-900/90 text-white shadow-2xl backdrop-blur-md"
      style={{
        position: 'fixed',
        top: box.top,
        left: box.left,
        width: box.width,
        maxHeight: box.maxHeight,
        zIndex: 2147483647,
        padding: prefs.dense ? '10px 12px' : '12px 14px',
        pointerEvents: 'none',
        boxShadow: '0 12px 40px rgba(0,0,0,0.55), 0 0 0 1px rgba(34,211,238,0.12)',
      }}
    >
      <div className="flex items-center gap-2 mb-1.5">
        <span className="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded-lg bg-cyan-500/15 border border-cyan-400/25 text-cyan-300">
          {InfoIcon ? <InfoIcon className="h-3 w-3" strokeWidth={2.5} /> : <span className="text-[9px] font-black">i</span>}
        </span>
        <div className="text-[10px] font-black uppercase tracking-[0.14em] text-cyan-300/95 leading-none">{intel.title}</div>
      </div>
      <p className="text-[11px] leading-relaxed text-white/75">{intel.body}</p>
      {intel.tips?.length > 0 && (
        <ul className="mt-2 space-y-1 border-t border-white/5 pt-2">
          {intel.tips.map((tip, idx) => (
            <li key={idx} className="text-[10px] text-cyan-100/70 flex gap-1.5 leading-snug">
              <span className="text-cyan-500/80 shrink-0">·</span>
              <span>{tip}</span>
            </li>
          ))}
        </ul>
      )}
    </div>,
    document.body
  );
}

export default {
  IntelTip,
  IntelLabel,
  FieldIntelligenceSettings,
  FieldIntelligenceRuntime,
  resolveIntel,
  useIntelPrefs,
  loadIntelPrefs,
  saveIntelPrefs,
  INTEL_CATALOG,
};

