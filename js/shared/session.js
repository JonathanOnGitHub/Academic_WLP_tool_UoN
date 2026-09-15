// ═══════════════════════════════════════════════════════════
// SESSION PERSISTENCE — IndexedDB-backed session save/restore
// ═══════════════════════════════════════════════════════════
// Saves raw input file data so the user can restore their
// session on page reload.  No calculated results stored.
// Clearable via browser settings or the on-screen Clear button.
// ═══════════════════════════════════════════════════════════

const WLP_SESSION = (() => {
  const DB = 'wlp-session';
  const STORE = 'data';

  // ── IndexedDB helpers ────────────────────────────────────
  function openDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbPut(key, value) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function idbGet(key) {
    try {
      const db = await openDB();
      return new Promise(resolve => {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    } catch (_) { return null; }
  }

  async function idbKeys() {
    try {
      const db = await openDB();
      return new Promise(resolve => {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).getAllKeys();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => resolve([]);
      });
    } catch (_) { return []; }
  }

  async function idbClear() {
    try {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).clear();
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    } catch (_) {}
  }

  // ── Poll helper ──────────────────────────────────────────
  function waitFor(pred, timeoutMs = 15000) {
    const start = Date.now();
    return new Promise(resolve => {
      const check = () => {
        try {
          if (pred()) return resolve(true);
        } catch (_) {}
        if (Date.now() - start > timeoutMs) return resolve(false);
        setTimeout(check, 120);
      };
      check();
    });
  }

  // ── Tab-restore config ───────────────────────────────────
  // Each entry describes how to restore a saved tab.
  //   restore: async (saved) => … — re-trigger the tab's load/analyse
  //   waitFor / analyseBtn / autoAnalyse control the restore pipeline.

  const RESTORE = {
    assessment: {
      label: 'Non-timetabled Assessment',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseAssessmentXlsx(file);
      }
    },
    project: {
      label: 'Research Projects UG',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        projLoadFile(file);
        await waitFor(() => projRawProjects && projRawProjects.length > 0);
        document.getElementById('projAnalyseBtn').click();
      }
    },
    pgt: {
      label: 'Research Projects PGT',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        pgtLoadFile(file);
        await waitFor(() => pgtRawData && pgtRawData.length > 0);
        document.getElementById('pgtAnalyseBtn').click();
      }
    },
    tutorial: {
      label: 'Tutorials',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        tutProcessFile(file);
      }
    },
    mmi: {
      label: 'MMIs',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        mmiLoadFile(file);
        await waitFor(() => mmiRawWb !== null);
        document.getElementById('mmiAnalyseBtn').click();
      }
    },
    citizenship_file: {
      label: 'School Citizenship (file)',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseCitizenshipXlsx(file);
      }
    },
    citizenship_paste: {
      label: 'School Citizenship (paste)',
      async restore(saved) {
        if (saved.citTeaching) document.getElementById('cit-teaching').value = saved.citTeaching;
        if (saved.citResearch) document.getElementById('cit-research').value = saved.citResearch;
        if (saved.citSchool) document.getElementById('cit-school').value = saved.citSchool;
        document.getElementById('citAnalyseBtn').click();
      }
    },
    research: {
      label: 'Research Hours',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseResearchXlsx(file);
      }
    },
    pgr: {
      label: 'PGR Supervision',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || 'text/csv' });
        parsePgrCsv(file);
      }
    },
    aob: {
      label: 'Any Other Business',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseAobXlsx(file);
      }
    },
    welcomeweek: {
      label: 'Welcome Week',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseWwXlsx(file);
      }
    },
    simulation: {
      label: 'Simulations',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseSimXlsx(file);
      }
    },
    opendays: {
      label: 'Open Days',
      async restore(saved) {
        const file = new File([saved.data], saved.fileName, { type: saved.mime || '' });
        parseOpenDaysXlsx(file);
      }
    },
    pgr_training: {
      label: 'PGR Training',
      async restore(saved) {
        try {
          const data = JSON.parse(new TextDecoder().decode(saved.data));
          if (data.t1) document.getElementById('pgrtr-t1').value = data.t1;
          if (data.t2) document.getElementById('pgrtr-t2').value = data.t2;
          if (data.t3) document.getElementById('pgrtr-t3').value = data.t3;
          document.getElementById('pgrtrAnalyseBtn').click();
        } catch (_) {}
      }
    },
    teaching: {
      label: 'Timetabled Teaching',
      async restore(saved) {
        const fileArr = saved.files || [];
        // Push each restored file onto the upload list, then click Analyse
        for (const f of fileArr) {
          const file = new File(
            [new Blob([f.text], { type: f.mime || 'text/html' })],
            f.fileName,
            { type: f.mime || 'text/html' }
          );
          tlUploadedFiles.push(file);
        }
        tlUpdateFileList();
        tlAnalyseBtn.click();
      }
    }
  };

  // ── Public API ───────────────────────────────────────────
  return {

    /** Save a single file-based tab's data (called from FileReader onload) */
    async saveFile(tabKey, fileName, arrayBuffer, mime) {
      try {
        await idbPut(tabKey, {
          type: 'file',
          fileName,
          mime: mime || '',
          data: arrayBuffer,
          savedAt: Date.now()
        });
      } catch (_) { /* IndexedDB unavailable — silent */ }
    },

    /** Save teaching's accumulated file texts (called after Analyse) */
    async saveTeaching(files) {
      // files = [{ fileName, mime, text }, …]
      try {
        await idbPut('teaching', {
          type: 'files',
          files,
          savedAt: Date.now()
        });
      } catch (_) {}
    },

    /** Save citizenship paste data (called from its calculate handler) */
    async saveCitizenshipPaste(citTeaching, citResearch, citSchool) {
      try {
        await idbPut('citizenship_paste', {
          type: 'paste',
          citTeaching: citTeaching || '',
          citResearch: citResearch || '',
          citSchool: citSchool || '',
          savedAt: Date.now()
        });
      } catch (_) {}
    },

    /** Remove a single tab's saved data */
    async remove(tabKey) {
      try {
        const db = await openDB();
        return new Promise((resolve, reject) => {
          const tx = db.transaction(STORE, 'readwrite');
          tx.objectStore(STORE).delete(tabKey);
          tx.oncomplete = () => resolve();
          tx.onerror = () => reject(tx.error);
        });
      } catch (_) {}
    },

    /** Remove ALL saved session data */
    async clear() {
      await idbClear();
    },

    /** Get a summary of what's saved (for the banner) */
    async info() {
      const keys = await idbKeys();
      if (!keys.length) return null;
      const labels = [];
      for (const k of keys) {
        const entry = RESTORE[k];
        labels.push(entry ? entry.label : k);
      }
      return { count: keys.length, labels };
    },

    /** Restore every tab that has saved data */
    async restoreAll() {
      const keys = await idbKeys();
      if (!keys.length) return;
      for (const key of keys) {
        const saved = await idbGet(key);
        if (!saved) continue;
        const entry = RESTORE[key];
        if (!entry) continue;
        try {
          await entry.restore(saved);
        } catch (e) {
          console.warn(`WLP_SESSION: restore failed for "${key}":`, e);
        }
      }
    },

    /** Called once on DOMContentLoaded — shows banner if session exists */
    async initBanner() {
      const info = await this.info();
      if (!info) return;
      const banner = document.getElementById('sessionBanner');
      if (!banner) return;
      banner.style.display = 'flex';

      const desc = info.count === 1
        ? `📂 <strong>Session found</strong> — ${info.labels[0]}`
        : `📂 <strong>Session found</strong> — ${info.count} tabs (${info.labels.join(', ')})`;
      document.getElementById('sessionBannerText').innerHTML = desc;
    }
  };
})();

// ── Banner wiring (runs after DOM is ready) ────────────────
document.addEventListener('DOMContentLoaded', () => {
  const restoreBtn = document.getElementById('sessionRestore');
  const clearBtn  = document.getElementById('sessionClear');
  const banner    = document.getElementById('sessionBanner');

  if (restoreBtn) {
    restoreBtn.addEventListener('click', async e => {
      e.preventDefault();
      banner.style.display = 'none';
      restoreBtn.textContent = '⏳ Restoring…';
      restoreBtn.disabled = true;
      await WLP_SESSION.restoreAll();
      // Give a moment for rendering to settle
      setTimeout(() => { restoreBtn.textContent = '✓ Restored'; }, 500);
    });
  }

  if (clearBtn) {
    clearBtn.addEventListener('click', async e => {
      e.preventDefault();
      await WLP_SESSION.clear();
      banner.style.display = 'none';
    });
  }

  WLP_SESSION.initBanner();
});
