// =========================================================================
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
