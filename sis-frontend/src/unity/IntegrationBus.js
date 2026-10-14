/**
 * Unity SIS — IntegrationBus
 * Lightweight cross-module bridge (no recursive event loops).
 */

const KEY = 'uu_sis_integration_v1';
const EVT = 'uu-sis-sync';

function safeParse(raw, fallback) {
  try {
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

export const UnityIntegrationBus = {
  KEY,
  EVT,

  read() {
    if (typeof localStorage === 'undefined') return {};
    return safeParse(localStorage.getItem(KEY), {});
  },

  /**
   * @param {object} partial
   * @param {{ silent?: boolean }} [opts] silent skips CustomEvent (prevents subscribe loops)
   */
  write(partial = {}, opts = {}) {
    if (typeof localStorage === 'undefined') return partial;
    const next = {
      ...this.read(),
      ...partial,
      updated_at: new Date().toISOString(),
    };
    localStorage.setItem(KEY, JSON.stringify(next));
    if (!opts.silent) {
      try {
        window.dispatchEvent(new CustomEvent(EVT, { detail: next }));
      } catch (_) {}
    }
    return next;
  },

  subscribe(handler) {
    if (typeof window === 'undefined') return () => {};
    const fn = (e) => {
      try {
        handler(e.detail || this.read());
      } catch (err) {
        console.warn('[IntegrationBus] subscriber error', err);
      }
    };
    window.addEventListener(EVT, fn);
    const onStorage = (e) => {
      if (e.key === KEY) {
        try {
          handler(this.read());
        } catch (err) {
          console.warn('[IntegrationBus] storage handler error', err);
        }
      }
    };
    window.addEventListener('storage', onStorage);
    return () => {
      window.removeEventListener(EVT, fn);
      window.removeEventListener('storage', onStorage);
    };
  },

  pushJournalBatch(batch) {
    const cur = this.read();
    const journals = Array.isArray(cur.pending_journals) ? cur.pending_journals : [];
    const entry = {
      id: `batch-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      created_at: new Date().toISOString(),
      status: 'pending',
      ...batch,
    };
    return this.write({ pending_journals: [entry, ...journals].slice(0, 80) });
  },

  consumePendingJournals({ clear = true } = {}) {
    const cur = this.read();
    const list = Array.isArray(cur.pending_journals) ? cur.pending_journals : [];
    // Silent clear — do NOT re-fire EVT or subscribers call consume again → stack overflow
    if (clear && list.length) {
      this.write({ pending_journals: [] }, { silent: true });
    }
    return list;
  },

  pushMaterialSignal(signal) {
    const cur = this.read();
    const signals = Array.isArray(cur.material_signals) ? cur.material_signals : [];
    return this.write({
      material_signals: [{ id: `sig-${Date.now()}`, at: new Date().toISOString(), ...signal }, ...signals].slice(0, 100),
    });
  },

  readMaterialSignals() {
    return this.read().material_signals || [];
  },

  setActiveModule(moduleId) {
    return this.write({ active_module: moduleId }, { silent: true });
  },

  getActiveModule() {
    return this.read().active_module || 'finance';
  },
};

export default UnityIntegrationBus;
