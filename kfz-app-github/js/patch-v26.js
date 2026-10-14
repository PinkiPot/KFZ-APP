// =========================================================================
// KFZ-APP PATCH-V26: DEMO-INITIALISIERUNG, VERWALTUNG AUF ZAHNRAD & SYNC
// =========================================================================

(function() {
  // --- 1. VERLAUFSSPEICHER (STRG + Z / STRG + Y) ---
  const undoStack = [];
  const redoStack = [];
  let isUndoRedoActive = false;

  window.recordAppSnapshot = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snap = JSON.stringify(window.state);
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snap) {
        undoStack.push(snap);
        if (undoStack.length > 40) undoStack.shift();
        redoStack.length = 0;
      }
    } catch(e) {}
  };

  window.executeUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte im Verlauf.');
      return;
    }
    try {
      isUndoRedoActive = true;
      redoStack.push(JSON.stringify(window.state));
      const prev = JSON.parse(undoStack.pop());
      Object.assign(window.state, prev);
      refreshView();
      if (typeof showToast === 'function') showToast('↩ Schritt rückgängig gemacht');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.executeRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));
      const next = JSON.parse(redoStack.pop());
      Object.assign(window.state, next);
      refreshView();
      if (typeof showToast === 'function') showToast('↪ Schritt wiederhergestellt');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  function refreshView() {
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') renderApp();
    else if (typeof renderAll === 'function') renderAll();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);
  }

  window.addEventListener('keydown', function(e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        if (e.shiftKey) {
          e.preventDefault();
          window.executeRedo();
        } else {
          const tag = (document.activeElement && document.activeElement.tagName) || '';
          if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            e.preventDefault();
            window.executeUndo();
          }
        }
      } else if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        window.executeRedo();
      }
    }
  });

  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.tagName === 'INPUT')) {
      window.recordAppSnapshot();
    }
  });

  // --- 2. ZAHNRAD MIT VERWALTUNGS-INHALT VERKNÜPFEN ---
  function wireSettingsGearToVerwaltung() {
    const settingsGear = document.getElementById('settingsGear');
    if (!settingsGear) return;

    settingsGear.onclick = function(e) {
      e.preventDefault();
      e.stopPropagation();

      // Klick auf den nativen Verwaltung-Button simulieren
      const adminBtn = document.getElementById('adminBtn') || document.querySelector('[data-action="admin"]');
      if (adminBtn) {
        adminBtn.click();
        return;
      }

      // Direkte Funktionsaufrufe von app.js als Fallback
      if (typeof openAdminModal === 'function') openAdminModal();
      else if (typeof openManagementModal === 'function') openManagementModal();
      else if (typeof showAdminModal === 'function') showAdminModal();
      else if (typeof renderAdminModal === 'function') renderAdminModal();
    };
  }

  // --- 3. INITIALISIERUNG & SUPABASE-SYNC ---
  async function launchDemo() {
    document.body.classList.add('dark');

    // Nativer App.js Login als Admin ausführen
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    // Cloudflare-Identität ermitteln
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
      window.state.darkMode = true;
    }

    // Aus Supabase laden, bestehende Admin-Daten schonend übernehmen
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          if (savedData.sheets && savedData.sheets.length > 0) {
            window.state.sheets = savedData.sheets;
          }
          if (savedData.settings) {
            window.state.settings = Object.assign({}, window.state.settings, savedData.settings);
          }
          document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // Erstes Sheet aktivieren
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    // Vollständig rendern
    refreshView();

    // Ersten Reiter aktivieren
    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, #tabs .tab');
      if (activeTab) activeTab.click();
      wireSettingsGearToVerwaltung();
    }, 150);

    // Speichern-Button verdrahten
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const origText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            saveInfo.textContent = 'Auto · ' + new Date().toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = origText;
        }, 1500);
      };
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(launchDemo, 100));
  } else {
    setTimeout(launchDemo, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: NATIVE DEMO-INITIALISIERUNG, SYNC & STRG + Z
// =========================================================================

(function() {
  // --- 1. VERLAUFSSPEICHER (STRG + Z / STRG + Y) ---
  const undoStack = [];
  const redoStack = [];
  let isUndoRedoActive = false;

  window.recordAppSnapshot = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snap = JSON.stringify(window.state);
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snap) {
        undoStack.push(snap);
        if (undoStack.length > 40) undoStack.shift();
        redoStack.length = 0;
      }
    } catch(e) {}
  };

  window.executeUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte im Verlauf.');
      return;
    }
    try {
      isUndoRedoActive = true;
      redoStack.push(JSON.stringify(window.state));
      const prev = JSON.parse(undoStack.pop());
      Object.assign(window.state, prev);
      refreshView();
      if (typeof showToast === 'function') showToast('↩ Schritt rückgängig gemacht');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.executeRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));
      const next = JSON.parse(redoStack.pop());
      Object.assign(window.state, next);
      refreshView();
      if (typeof showToast === 'function') showToast('↪ Schritt wiederhergestellt');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  function refreshView() {
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') renderApp();
    else if (typeof renderAll === 'function') renderAll();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);
  }

  window.addEventListener('keydown', function(e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        if (e.shiftKey) {
          e.preventDefault();
          window.executeRedo();
        } else {
          const tag = (document.activeElement && document.activeElement.tagName) || '';
          if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            e.preventDefault();
            window.executeUndo();
          }
        }
      } else if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        window.executeRedo();
      }
    }
  });

  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.tagName === 'INPUT')) {
      window.recordAppSnapshot();
    }
  });

  // --- 2. NATIVES BOOTEN IN DEN EXAKTEN DEMO-ZUSTAND ---
  async function launchExactDemo() {
    // Dunkelmodus sicherstellen
    document.body.classList.add('dark');

    // Nativer App.js Login als Admin (erzeugt die exakte Demo-Ansicht mit allen Reitern und Symbolen)
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    // Cloudflare E-Mail abrufen (ohne die Demo-Rolle zu stören)
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
      window.state.darkMode = true;
    }

    // Supabase Sync
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync:", err);
      }
    }

    // Erstes Sheet (A3 8P) aktivieren
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    // Rendern
    refreshView();

    // Den ersten Tab anklicken
    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, #tabs .tab');
      if (activeTab) activeTab.click();
    }, 150);

    // Speichern-Button mit Supabase verbinden
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const origText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            saveInfo.textContent = 'Auto · ' + new Date().toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = origText;
        }, 1500);
      };
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(launchExactDemo, 100));
  } else {
    setTimeout(launchExactDemo, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: ORIGINAL INITIALISIERUNG, SETTINGS, SYNC & STRG+Z
// =========================================================================

(function() {
  // --- 1. VERLAUFSSPEICHER (STRG + Z / STRG + Y) ---
  const undoStack = [];
  const redoStack = [];
  let isUndoRedoActive = false;

  window.recordAppSnapshot = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snap = JSON.stringify(window.state);
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snap) {
        undoStack.push(snap);
        if (undoStack.length > 40) undoStack.shift();
        redoStack.length = 0;
      }
    } catch(e) {}
  };

  window.executeUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte im Verlauf.');
      return;
    }
    try {
      isUndoRedoActive = true;
      redoStack.push(JSON.stringify(window.state));
      const prev = JSON.parse(undoStack.pop());
      Object.assign(window.state, prev);
      refreshView();
      if (typeof showToast === 'function') showToast('↩ Schritt rückgängig gemacht');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.executeRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));
      const next = JSON.parse(redoStack.pop());
      Object.assign(window.state, next);
      refreshView();
      if (typeof showToast === 'function') showToast('↪ Schritt wiederhergestellt');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  function refreshView() {
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') renderApp();
    else if (typeof renderAll === 'function') renderAll();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);
  }

  window.addEventListener('keydown', function(e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        if (e.shiftKey) {
          e.preventDefault();
          window.executeRedo();
        } else {
          const tag = (document.activeElement && document.activeElement.tagName) || '';
          if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            e.preventDefault();
            window.executeUndo();
          }
        }
      } else if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        window.executeRedo();
      }
    }
  });

  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.tagName === 'INPUT')) {
      window.recordAppSnapshot();
    }
  });

  // --- 2. ZENTRALES APPEINSTELLUNGEN-MODAL ---
  window.openSettingsDialog = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot) return;

    const isDark = document.body.classList.contains('dark');
    const densityY = getComputedStyle(document.documentElement).getPropertyValue('--density-cell-y').trim() || '5px';
    const fontScale = getComputedStyle(document.documentElement).getPropertyValue('--fontScale').trim() || '1';

    modalRoot.innerHTML = `
      <div class="modal" id="settingsOverlay" onclick="if(event.target===this) window.closeModalDialog()">
        <div class="modal-box" style="max-width: 660px;">
          <div class="modal-head">
            <h3>⚙ Appeinstellungen & Systemverwaltung</h3>
            <button class="ghost small" onclick="window.closeModalDialog()">✕</button>
          </div>
          <div class="modal-body" style="gap: 16px;">
            
            <div>
              <h4 style="margin: 0 0 10px; color: var(--primary);">Erscheinungsbild & Layout</h4>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <div>
                  <strong>Farbschema</strong>
                  <div class="hint">Dunkles oder helles Oberflächendesign</div>
                </div>
                <button id="setDarkBtn" class="btn secondary small">
                  ${isDark ? '☀️ Helles Design' : '🌙 Dunkles Design'}
                </button>
              </div>

              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <div>
                  <strong>Tabellendichte</strong>
                  <div class="hint">Zeilenhöhe der Tabellen anpassen</div>
                </div>
                <select id="setDensity" style="padding: 4px 8px;">
                  <option value="3px" ${densityY === '3px' ? 'selected' : ''}>Kompakt (3px)</option>
                  <option value="5px" ${densityY === '5px' ? 'selected' : ''}>Standard (5px)</option>
                  <option value="8px" ${densityY === '8px' ? 'selected' : ''}>Komfortabel (8px)</option>
                </select>
              </div>

              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <div>
                  <strong>Schriftgröße</strong>
                  <div class="hint">Skalierung aller Textinhalte</div>
                </div>
                <select id="setFontScale" style="padding: 4px 8px;">
                  <option value="0.9" ${fontScale === '0.9' ? 'selected' : ''}>90% (Klein)</option>
                  <option value="1" ${fontScale === '1' ? 'selected' : ''}>100% (Standard)</option>
                  <option value="1.1" ${fontScale === '1.1' ? 'selected' : ''}>110% (Groß)</option>
                </select>
              </div>
            </div>

            <hr style="border:0; border-top:1px solid var(--border);">

            <div>
              <h4 style="margin: 0 0 8px; color: var(--primary);">Versionsinformation & Release Notes</h4>
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span>Installierte Version: <strong>V26.07.26</strong></span>
                <button id="setReleaseBtn" class="btn secondary small">📝 Release Notes anzeigen</button>
              </div>
            </div>

            <hr style="border:0; border-top:1px solid var(--border);">

            <div>
              <h4 style="margin: 0 0 8px; color: var(--primary);">Benutzerkonto & Cloud-Sync</h4>
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span>Angemeldet via Cloudflare</span>
                <span class="pill ok">${(window.state && window.state.currentUser) || 'kirchbergpa@gmail.com'}</span>
              </div>
            </div>

          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.closeModalDialog()">Schließen</button>
          </div>
        </div>
      </div>
    `;

    document.getElementById('setDarkBtn').onclick = function() {
      document.body.classList.toggle('dark');
      const darkNow = document.body.classList.contains('dark');
      this.textContent = darkNow ? '☀️ Helles Design' : '🌙 Dunkles Design';
      if (window.state) window.state.darkMode = darkNow;
    };

    document.getElementById('setDensity').onchange = function() {
      document.documentElement.style.setProperty('--density-cell-y', this.value);
    };

    document.getElementById('setFontScale').onchange = function() {
      document.documentElement.style.setProperty('--fontScale', this.value);
    };

    document.getElementById('setReleaseBtn').onclick = function() {
      window.openReleaseDialog();
    };
  };

  // --- 3. RELEASE NOTES MODAL ---
  window.openReleaseDialog = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot) return;

    modalRoot.innerHTML = `
      <div class="modal" id="releaseOverlay" onclick="if(event.target===this) window.openSettingsDialog()">
        <div class="modal-box" style="max-width: 700px;">
          <div class="modal-head">
            <h3>📝 Release Notes · Versionshistorie</h3>
            <button class="ghost small" onclick="window.openSettingsDialog()">✕</button>
          </div>
          <div class="modal-body" style="max-height: 65vh; overflow-y: auto; line-height: 1.6;">
            <div style="margin-bottom: 14px;">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong>V26.07.26 (Aktuell)</strong>
                <span class="pill ok">Produktiv</span>
              </div>
              <ul style="margin: 6px 0 0 18px; padding: 0;">
                <li>Cloudflare Access Integration: Sicherer PIN-Zugang ohne manuelle Login-Maske.</li>
                <li>Supabase Synchronisation: Zuverlässiges Speichern aller KFZ-Reiter und Tabellendaten.</li>
                <li>Verlaufsspeicher: Änderungen jederzeit mit <code>Strg + Z</code> / <code>Strg + Y</code> zurücknehmen.</li>
                <li>Aufgeräumtes Design: Alle Systemfunktionen zentral im Appeinstellungen-Menü.</li>
              </ul>
            </div>
            <div style="margin-bottom: 14px;">
              <strong>V26.07.19</strong>
              <ul style="margin: 6px 0 0 18px; padding: 0;">
                <li>Dynamische Fälligkeitsberechnung anhand des aktuellen Fahrzeug-KM-Stands.</li>
                <li>Katalog- und Bildvorschau für Ersatzteile und Werkzeuge.</li>
                <li>Themen-Navigation mit Sprungmarken für schnelle Übersicht.</li>
              </ul>
            </div>
          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.openSettingsDialog()">Zurück zu den Einstellungen</button>
          </div>
        </div>
      </div>
    `;
  };

  window.closeModalDialog = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (modalRoot) modalRoot.innerHTML = '';
  };

  // --- 4. NATIVES BOOTING & INITIALISIERUNG ---
  async function launchApplication() {
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // Nativer Aufruf von app.js Quick-Login
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // Supabase laden
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          if (window.state.darkMode) document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // Erstes Blatt aktivieren
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    // Vollständiges Rendern
    refreshView();

    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, #tabs .tab');
      if (activeTab) activeTab.click();
    }, 150);

    // Zahnrad unten rechts verdrahten
    const settingsGear = document.getElementById('settingsGear');
    if (settingsGear) {
      settingsGear.onclick = function() {
        window.openSettingsDialog();
      };
    }

    // Speichern-Button mit Supabase verbinden
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const originalText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            saveInfo.textContent = 'Gespeichert · ' + new Date().toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = originalText;
        }, 1500);
      };
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(launchApplication, 100));
  } else {
    setTimeout(launchApplication, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: ORIGINALGETREUE WIEDERHERSTELLUNG & SYNC
// =========================================================================

(function() {
  // --- 1. UNDO / REDO SYSTEM (STRG + Z / STRG + Y) ---
  const undoStack = [];
  const redoStack = [];
  let isUndoRedoActive = false;

  window.captureHistorySnapshot = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snap = JSON.stringify(window.state);
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snap) {
        undoStack.push(snap);
        if (undoStack.length > 40) undoStack.shift();
        redoStack.length = 0;
      }
    } catch(e) {}
  };

  window.executeUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte im Verlauf.');
      return;
    }
    try {
      isUndoRedoActive = true;
      redoStack.push(JSON.stringify(window.state));
      const prev = JSON.parse(undoStack.pop());
      Object.assign(window.state, prev);
      refreshAppViews();
      if (typeof showToast === 'function') showToast('↩ Schritt rückgängig gemacht');
    } catch(e) {
      console.error("Undo Fehler:", e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.executeRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));
      const next = JSON.parse(redoStack.pop());
      Object.assign(window.state, next);
      refreshAppViews();
      if (typeof showToast === 'function') showToast('↪ Schritt wiederhergestellt');
    } catch(e) {
      console.error("Redo Fehler:", e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  function refreshAppViews() {
    if (typeof renderAll === 'function') renderAll();
    else if (typeof renderApp === 'function') renderApp();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);
    if (typeof renderTabs === 'function') renderTabs();
  }

  window.addEventListener('keydown', function(e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        if (e.shiftKey) {
          e.preventDefault();
          window.executeRedo();
        } else {
          const tag = (document.activeElement && document.activeElement.tagName) || '';
          if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            e.preventDefault();
            window.executeUndo();
          }
        }
      } else if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        window.executeRedo();
      }
    }
  });

  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.tagName === 'INPUT')) {
      window.captureHistorySnapshot();
    }
  });

  // --- 2. DIE ECHTEN ORIGINAL-APPEINSTELLUNGEN RENDERN ---
  window.openOriginalSettingsModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot || !window.state) return;

    const densityY = getComputedStyle(document.documentElement).getPropertyValue('--density-cell-y').trim() || '5px';
    const fontScale = getComputedStyle(document.documentElement).getPropertyValue('--fontScale').trim() || '1';
    const isDark = document.body.classList.contains('dark');

    modalRoot.innerHTML = `
      <div class="modal" id="appSettingsOverlay" onclick="if(event.target===this) window.closeOriginalModal()">
        <div class="modal-box" style="max-width: 680px;">
          <div class="modal-head">
            <h3>⚙ Appeinstellungen & Systemverwaltung</h3>
            <button class="ghost small" onclick="window.closeOriginalModal()">✕</button>
          </div>
          <div class="modal-body" style="gap: 16px;">

            <!-- Darstellung & Design -->
            <div>
              <h4 style="margin: 0 0 8px 0; color: var(--primary);">Erscheinungsbild</h4>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <span>Farbschema (Dark Mode)</span>
                <button id="cfgDarkModeBtn" class="btn secondary small">
                  ${isDark ? '☀️ Helles Design aktivieren' : '🌙 Dunkles Design aktivieren'}
                </button>
              </div>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <span>Tabellendichte</span>
                <select id="cfgDensitySelect" style="padding: 4px 8px;">
                  <option value="3px" ${densityY === '3px' ? 'selected' : ''}>Kompakt (3px)</option>
                  <option value="5px" ${densityY === '5px' ? 'selected' : ''}>Standard (5px)</option>
                  <option value="8px" ${densityY === '8px' ? 'selected' : ''}>Komfortabel (8px)</option>
                </select>
              </div>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 6px 0;">
                <span>Schriftgrößen-Skalierung</span>
                <select id="cfgFontScaleSelect" style="padding: 4px 8px;">
                  <option value="0.9" ${fontScale === '0.9' ? 'selected' : ''}>90% (Klein)</option>
                  <option value="1" ${fontScale === '1' ? 'selected' : ''}>100% (Standard)</option>
                  <option value="1.1" ${fontScale === '1.1' ? 'selected' : ''}>110% (Groß)</option>
                </select>
              </div>
            </div>

            <hr style="border:0; border-top:1px solid var(--border);">

            <!-- Release Notes & Version -->
            <div>
              <h4 style="margin: 0 0 8px 0; color: var(--primary);">Versionsverwaltung & Changelog</h4>
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span>Installierte Version: <strong>V26.07.26</strong></span>
                <button id="cfgOpenReleaseBtn" class="btn secondary small">📝 Release Notes anzeigen</button>
              </div>
            </div>

            <hr style="border:0; border-top:1px solid var(--border);">

            <!-- Authentifizierung & Cloud-Sync -->
            <div>
              <h4 style="margin: 0 0 8px 0; color: var(--primary);">Benutzer & Speicherstatus</h4>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 4px 0;">
                <span>Konto</span>
                <span class="pill ok">${window.state.currentUser || 'kirchbergpa@gmail.com'}</span>
              </div>
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 4px 0;">
                <span>Synchronisation</span>
                <span class="hint">Supabase Cloud aktiv</span>
              </div>
            </div>

          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.closeOriginalModal()">Schließen</button>
          </div>
        </div>
      </div>
    `;

    // Handler Dark Mode
    document.getElementById('cfgDarkModeBtn').onclick = function() {
      document.body.classList.toggle('dark');
      const darkNow = document.body.classList.contains('dark');
      this.textContent = darkNow ? '☀️ Helles Design aktivieren' : '🌙 Dunkles Design aktivieren';
      if (window.state) window.state.darkMode = darkNow;
    };

    // Handler Dichte
    document.getElementById('cfgDensitySelect').onchange = function() {
      document.documentElement.style.setProperty('--density-cell-y', this.value);
    };

    // Handler Skalierung
    document.getElementById('cfgFontScaleSelect').onchange = function() {
      document.documentElement.style.setProperty('--fontScale', this.value);
    };

    // Handler Release Notes
    document.getElementById('cfgOpenReleaseBtn').onclick = function() {
      window.openOriginalReleaseModal();
    };
  };

  // --- 3. DIE ECHTEN ORIGINAL-RELEASE-NOTES RENDERN ---
  window.openOriginalReleaseModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot) return;

    modalRoot.innerHTML = `
      <div class="modal" id="appReleaseOverlay" onclick="if(event.target===this) window.openOriginalSettingsModal()">
        <div class="modal-box" style="max-width: 720px;">
          <div class="modal-head">
            <h3>📝 Release Notes · Versionshistorie</h3>
            <button class="ghost small" onclick="window.openOriginalSettingsModal()">✕</button>
          </div>
          <div class="modal-body" style="max-height: 65vh; overflow-y: auto; line-height: 1.55;">
            
            <div style="margin-bottom: 14px;">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong>V26.07.26 (Aktuell)</strong>
                <span class="pill ok">Produktiv</span>
              </div>
              <ul style="margin: 6px 0 0 18px; padding: 0;">
                <li>Vollständige Cloudflare Access E-Mail-Authentifizierung (Zero Trust PIN).</li>
                <li>Permanente Supabase-Synchronisation aller Arbeitsblätter und Fahrzeuge.</li>
                <li>Globaler Rückgängig-Speicher via <code>Strg + Z</code> / <code>Strg + Y</code>.</li>
                <li>Zentrale Systemkonfiguration gebündelt im Appeinstellungen-Menü.</li>
              </ul>
            </div>

            <div style="margin-bottom: 14px;">
              <strong>V26.07.19</strong>
              <ul style="margin: 6px 0 0 18px; padding: 0;">
                <li>Dynamische Fälligkeitsberechnung anhand des aktuellen KM-Stands.</li>
                <li>Erweiterter Ersatzteilkatalog mit Bildvorschau und Verbauliste.</li>
                <li>Themen-Navigation mit Anker-Sprungmarken zu Tabellenblöcken.</li>
              </ul>
            </div>

            <div>
              <strong>V26.07.01</strong>
              <ul style="margin: 6px 0 0 18px; padding: 0;">
                <li>Initiales MRO-Datenmodell für Wartungsintervalle und Werkzeuge.</li>
                <li>Excel-Import und Drag-and-Drop Spaltenverwaltung.</li>
              </ul>
            </div>

          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.openOriginalSettingsModal()">Zurück zu den Einstellungen</button>
          </div>
        </div>
      </div>
    `;
  };

  window.closeOriginalModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (modalRoot) modalRoot.innerHTML = '';
  };

  // --- 4. SYSTEM-BOOTSTRAP & INITIALISIERUNG ---
  async function launchSystem() {
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // Nativer Aufruf von app.js Quick-Login für sauberen State
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    // Identität anwenden
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // Supabase laden
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          if (window.state.darkMode) document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // Erstes Blatt aktivieren
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    // Rendern
    refreshAppViews();

    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, .tabs .tab');
      if (activeTab) activeTab.click();
    }, 150);

    // Einstellungs-Zahnrad verbinden
    const settingsGear = document.getElementById('settingsGear');
    if (settingsGear) {
      settingsGear.onclick = function() {
        window.openOriginalSettingsModal();
      };
    }

    // Speichern-Button mit Supabase verbinden
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const originalText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            saveInfo.textContent = 'Gespeichert · ' + new Date().toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = originalText;
        }, 1500);
      };
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(launchSystem, 100));
  } else {
    setTimeout(launchSystem, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: NATIVE INITIALISIERUNG ALLER ORIGINAL-FUNKTIONEN
// =========================================================================

(function() {
  // --- GLOBALER UNDO / REDO STACK (STRG + Z / STRG + Y) ---
  const undoStack = [];
  const redoStack = [];
  let isUndoRedoActive = false;

  window.recordAppSnapshot = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snap = JSON.stringify(window.state);
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snap) {
        undoStack.push(snap);
        if (undoStack.length > 35) undoStack.shift();
        redoStack.length = 0;
      }
    } catch(e) {}
  };

  window.triggerUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte zum Rückgängigmachen.');
      return;
    }
    try {
      isUndoRedoActive = true;
      redoStack.push(JSON.stringify(window.state));
      const prev = JSON.parse(undoStack.pop());
      Object.assign(window.state, prev);
      if (typeof renderAll === 'function') renderAll();
      else if (typeof renderApp === 'function') renderApp();
      else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
      if (typeof showToast === 'function') showToast('↩ Schritt rückgängig gemacht');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.triggerRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));
      const nxt = JSON.parse(redoStack.pop());
      Object.assign(window.state, nxt);
      if (typeof renderAll === 'function') renderAll();
      else if (typeof renderApp === 'function') renderApp();
      else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
      if (typeof showToast === 'function') showToast('↪ Schritt wiederhergestellt');
    } catch(e) {
      console.error(e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  window.addEventListener('keydown', function(e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        if (e.shiftKey) { e.preventDefault(); window.triggerRedo(); }
        else {
          const tag = (document.activeElement && document.activeElement.tagName) || '';
          if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            e.preventDefault();
            window.triggerUndo();
          }
        }
      } else if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        window.triggerRedo();
      }
    }
  });

  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.tagName === 'INPUT')) {
      window.recordAppSnapshot();
    }
  });

  // --- NATIVE SYSTEM-BOOTSTRAP ---
  async function startOriginalApp() {
    // 1. Cloudflare Access Benutzer ermitteln
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // 2. Den echten Login von app.js ausführen (baut den kompletten State & Listener auf)
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    // 3. Benutzerdaten setzen
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // 4. Supabase-Daten synchronisieren
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          if (window.state.darkMode) document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // 5. Erstes Blatt aktiv schalten & Komplett rendern
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    // Ruft die echten Render-Routinen von app.js auf
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderAll === 'function') renderAll();
    else if (typeof renderApp === 'function') renderApp();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);

    // Ersten Reiter aktivieren
    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, .tabs .tab');
      if (activeTab) activeTab.click();
    }, 150);

    // 6. Die originalen Appeinstellungen von app.js öffnen
    const settingsGear = document.getElementById('settingsGear');
    if (settingsGear) {
      settingsGear.onclick = function() {
        if (typeof openSettingsModal === 'function') openSettingsModal();
        else if (typeof renderSettingsModal === 'function') renderSettingsModal();
        else if (typeof openSettings === 'function') openSettings();
        else if (typeof showSettings === 'function') showSettings();
      };
    }

    // 7. Speichern mit Supabase
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const origText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            saveInfo.textContent = 'Gespeichert · ' + new Date().toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = origText;
        }, 1500);
      };
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(startOriginalApp, 100));
  } else {
    setTimeout(startOriginalApp, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: ZENTRALE EINSTELLUNGEN & STRG + Z (UNDO/REDO)
// =========================================================================

(function() {
  // --- VERLAUFSSPEICHER (UNDO / REDO) ---
  const undoStack = [];
  const redoStack = [];
  const MAX_HISTORY = 40;
  let isUndoRedoActive = false;

  // Speichert den aktuellen Status vor einer Bearbeitung
  window.pushHistoryState = function() {
    if (isUndoRedoActive || !window.state) return;
    try {
      const snapshot = JSON.stringify(window.state);
      // Nur speichern, wenn sich zum vorherigen Schnappschuss tatsächlich etwas geändert hat
      if (undoStack.length === 0 || undoStack[undoStack.length - 1] !== snapshot) {
        undoStack.push(snapshot);
        if (undoStack.length > MAX_HISTORY) undoStack.shift();
        redoStack.length = 0; // Nach einer neuen manuellen Änderung Redo leeren
      }
    } catch(e) {}
  };

  // Rückgängig machen (Strg + Z)
  window.performUndo = function() {
    if (undoStack.length === 0 || !window.state) {
      if (typeof showToast === 'function') showToast('Keine weiteren Schritte zum Rückgängigmachen.');
      return;
    }
    try {
      isUndoRedoActive = true;
      // Aktuellen Stand im Redo-Stack sichern
      redoStack.push(JSON.stringify(window.state));

      // Letzten Stand wiederherstellen
      const previousState = JSON.parse(undoStack.pop());
      Object.assign(window.state, previousState);

      // Ansicht neu zeichnen
      if (typeof renderTabs === 'function') renderTabs();
      if (typeof renderApp === 'function') renderApp();
      else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
      else if (typeof renderSheet === 'function') renderSheet(window.state.activeSheetId);

      if (typeof showToast === 'function') {
        showToast('↩ Letzter Schritt rückgängig gemacht');
      }
    } catch(e) {
      console.error("Undo-Fehler:", e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  // Wiederholen (Strg + Y / Strg + Shift + Z)
  window.performRedo = function() {
    if (redoStack.length === 0 || !window.state) return;
    try {
      isUndoRedoActive = true;
      undoStack.push(JSON.stringify(window.state));

      const nextState = JSON.parse(redoStack.pop());
      Object.assign(window.state, nextState);

      if (typeof renderTabs === 'function') renderTabs();
      if (typeof renderApp === 'function') renderApp();
      else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
      else if (typeof renderSheet === 'function') renderSheet(window.state.activeSheetId);

      if (typeof showToast === 'function') {
        showToast('↪ Schritt wiederhergestellt');
      }
    } catch(e) {
      console.error("Redo-Fehler:", e);
    } finally {
      isUndoRedoActive = false;
    }
  };

  // Tastatur-Events global abfangen
  window.addEventListener('keydown', function(e) {
    const isCtrlOrCmd = e.ctrlKey || e.metaKey;
    if (!isCtrlOrCmd) return;

    // Strg + Z (Rückgängig)
    if (e.key === 'z' || e.key === 'Z') {
      if (e.shiftKey) {
        // Strg + Shift + Z = Wiederholen
        e.preventDefault();
        window.performRedo();
      } else {
        // Wenn man sich in einem normalen Textfeld befindet, greift Undo gezielt auf die App-Tabelle zu
        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
        if (activeTag !== 'input' && activeTag !== 'textarea') {
          e.preventDefault();
          window.performUndo();
        }
      }
    }
    // Strg + Y (Wiederholen)
    else if (e.key === 'y' || e.key === 'Y') {
      e.preventDefault();
      window.performRedo();
    }
  });

  // Tabellen-Eingaben automatisch für den Verlauf überwachen
  document.addEventListener('focusin', function(e) {
    if (e.target && (e.target.isContentEditable || e.target.classList.contains('cell-text') || e.target.tagName === 'INPUT')) {
      window.pushHistoryState();
    }
  });

  // --- EINSTELLUNGEN & MODAL-DIALOGE ---
  window.openCustomSettingsModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot) return;

    const isDark = document.body.classList.contains('dark');
    const version = "V26.07.26";

    modalRoot.innerHTML = `
      <div class="modal" id="settingsModalOverlay" onclick="if(event.target===this) window.closeCustomModal()">
        <div class="modal-box" style="max-width: 600px;">
          <div class="modal-head">
            <h3>⚙ Appeinstellungen & Verwaltung</h3>
            <button class="ghost small" onclick="window.closeCustomModal()">✕</button>
          </div>
          <div class="modal-body">
            
            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid var(--border);">
              <div>
                <strong>Design / Erscheinungsbild</strong>
                <div class="hint">Zwischen Hell- und Dunkelmodus wechseln</div>
              </div>
              <button id="toggleDarkBtn" class="btn secondary">
                ${isDark ? '☀️ Helles Design' : '🌙 Dunkles Design'}
              </button>
            </div>

            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid var(--border);">
              <div>
                <strong>Versionsinformation & Release Notes</strong>
                <div class="hint">Installierte Version: ${version}</div>
              </div>
              <button id="btnOpenReleaseNotes" class="btn secondary">
                📝 Release Notes
              </button>
            </div>

            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0;">
              <div>
                <strong>Angemeldeter Benutzer</strong>
                <div class="hint">${(window.state && window.state.currentUser) || 'kirchbergpa@gmail.com'}</div>
              </div>
              <span class="pill ok">Cloudflare Access</span>
            </div>

          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.closeCustomModal()">Schließen</button>
          </div>
        </div>
      </div>
    `;

    const toggleDarkBtn = document.getElementById('toggleDarkBtn');
    if (toggleDarkBtn) {
      toggleDarkBtn.onclick = function() {
        document.body.classList.toggle('dark');
        const nowDark = document.body.classList.contains('dark');
        toggleDarkBtn.textContent = nowDark ? '☀️ Helles Design' : '🌙 Dunkles Design';
        if (window.state) window.state.darkMode = nowDark;
      };
    }

    const btnOpenReleaseNotes = document.getElementById('btnOpenReleaseNotes');
    if (btnOpenReleaseNotes) {
      btnOpenReleaseNotes.onclick = function() {
        window.openCustomReleaseModal();
      };
    }
  };

  window.openCustomReleaseModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (!modalRoot) return;

    modalRoot.innerHTML = `
      <div class="modal" id="releaseModalOverlay" onclick="if(event.target===this) window.openCustomSettingsModal()">
        <div class="modal-box" style="max-width: 650px;">
          <div class="modal-head">
            <h3>📝 Release Notes · V26.07.26</h3>
            <button class="ghost small" onclick="window.openCustomSettingsModal()">✕</button>
          </div>
          <div class="modal-body" style="line-height:1.6;">
            <h4>Neuerungen & Optimierungen:</h4>
            <ul>
              <li><strong>Rückgängig-Funktion (Strg + Z):</strong> Letzte Bearbeitungsschritte können jederzeit mit <code>Strg + Z</code> rückgängig gemacht und mit <code>Strg + Y</code> wiederholt werden.</li>
              <li><strong>Cloudflare Access Integration:</strong> Nahtlose PIN-Authentifizierung ohne Demo-Overlay.</li>
              <li><strong>Supabase Cloud-Sync:</strong> Zuverlässige Speicherung aller Arbeitsblätter und Fahrzeugdaten.</li>
              <li><strong>Zentrale Appeinstellungen:</strong> Dark Mode und Systemverwaltung gebündelt im Einstellungsmenü.</li>
            </ul>
          </div>
          <div class="modal-foot">
            <button class="btn" onclick="window.openCustomSettingsModal()">Zurück zu den Einstellungen</button>
          </div>
        </div>
      </div>
    `;
  };

  window.closeCustomModal = function() {
    const modalRoot = document.getElementById('modalRoot');
    if (modalRoot) modalRoot.innerHTML = '';
  };

  // --- BOOTSTRAP ANWENDUNG ---
  async function bootAppNaturally() {
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // Nativer App.js Login
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');
    if (quickAdmin) quickAdmin.click();
    else if (loginBtn) loginBtn.click();
    else if (typeof window.login === 'function') window.login('admin', 'admin');

    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // Supabase laden
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          if (window.state.darkMode) document.body.classList.add('dark');
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // Erstes Sheet aktivieren & Rendern
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') renderApp();
    else if (typeof renderAll === 'function') renderAll();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);

    // Klick auf den ersten Reiter
    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, .tabs .tab.active');
      const firstTab = document.querySelector('#tabs .tab, .tabs .tab');
      if (activeTab) activeTab.click();
      else if (firstTab) {
        firstTab.classList.add('active');
        firstTab.click();
      }
    }, 150);

    // Zahnrad unten rechts verdrahten
    const settingsGear = document.getElementById('settingsGear');
    if (settingsGear) {
      settingsGear.onclick = function() {
        window.openCustomSettingsModal();
      };
    }

    // Speichern-Button verdrahten
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.addEventListener('click', async function() {
        const originalText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            const now = new Date();
            saveInfo.textContent = 'Gespeichert · ' + now.toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = originalText;
        }, 1500);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(bootAppNaturally, 100));
  } else {
    setTimeout(bootAppNaturally, 100);
  }
})();// =========================================================================
// KFZ-APP PATCH-V26: NATIVE DEMO-INITIALISIERUNG + CLOUDFLARE SYNC
// =========================================================================

(function() {
  async function launchFullApplication() {
    // 1. Echte E-Mail aus Cloudflare Access laden
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // 2. Den nativen Demo-Login von app.js ausführen
    // Dadurch initialisiert app.js den kompletten internen State, baut alle Tabellen und Modals auf
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    const loginBtn = document.getElementById('loginBtn');

    if (quickAdmin) {
      quickAdmin.click();
    } else if (loginBtn) {
      loginBtn.click();
    } else if (typeof window.login === 'function') {
      window.login('admin', 'admin');
    }

    // 3. Status auf die Cloudflare-Identität übertragen
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // 4. Supabase-Daten synchronisieren (falls bereits Spielstand vorhanden)
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
        }
      } catch (err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // 5. Erstes Blatt aktiv schalten & Rendern sicherstellen
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0 && !window.state.activeSheetId) {
        window.state.activeSheetId = sheets[0].id || sheets[0].name;
      }
    }

    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') renderApp();
    else if (typeof renderAll === 'function') renderAll();
    else if (typeof renderCurrentSheet === 'function') renderCurrentSheet();
    else if (typeof renderSheet === 'function' && window.state) renderSheet(window.state.activeSheetId);

    // 6. Klick auf den ersten Reiter ausführen
    setTimeout(() => {
      const activeTab = document.querySelector('#tabs .tab.active, .tabs .tab.active');
      const firstTab = document.querySelector('#tabs .tab, .tabs .tab');
      if (activeTab) {
        activeTab.click();
      } else if (firstTab) {
        firstTab.classList.add('active');
        firstTab.click();
      }
    }, 150);

    // 7. Appeinstellungen & Release Notes nativ verbinden
    const settingsGear = document.getElementById('settingsGear');
    if (settingsGear) {
      settingsGear.onclick = function() {
        if (typeof openSettingsModal === 'function') openSettingsModal();
        else if (typeof openSettings === 'function') openSettings();
        else if (typeof showSettings === 'function') showSettings();
      };
    }

    const releaseBtn = document.getElementById('releaseBtn');
    const versionChip = document.getElementById('versionChip');
    const triggerRelease = function() {
      if (typeof openReleaseModal === 'function') openReleaseModal();
      else if (typeof openReleaseNotes === 'function') openReleaseNotes();
      else if (typeof showReleaseNotes === 'function') showReleaseNotes();
    };
    if (releaseBtn) releaseBtn.onclick = triggerRelease;
    if (versionChip) versionChip.onclick = triggerRelease;

    // 8. Speichern mit Supabase verbinden
    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        const originalText = saveBtn.textContent;
        saveBtn.disabled = true;
        saveBtn.textContent = '⏳ Speichern...';
        try {
          if (typeof saveToSupabase === 'function' && window.state) {
            await saveToSupabase(window.state);
          }
          saveBtn.textContent = '✅ Gespeichert!';
          const saveInfo = document.getElementById('saveInfo');
          if (saveInfo) {
            const now = new Date();
            saveInfo.textContent = 'Gespeichert · ' + now.toLocaleTimeString();
          }
        } catch(e) {
          saveBtn.textContent = '❌ Fehler!';
        }
        setTimeout(() => {
          saveBtn.disabled = false;
          saveBtn.textContent = originalText;
        }, 1500);
      };
    }
  }

  // Sicherer Start
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(launchFullApplication, 100));
  } else {
    setTimeout(launchFullApplication, 100);
  }
})();
