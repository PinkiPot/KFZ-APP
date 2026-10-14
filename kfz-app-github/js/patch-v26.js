// =========================================================================
// KFZ-APP PATCH-V26: VOLLSTÄNDIGE EVENT-INITIALISIERUNG & SUPABASE-SYNC
// =========================================================================

(function() {
  async function bootFullApp() {
    // 1. Echte Identität von Cloudflare ermitteln
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {}

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // 2. Demo-Trigger nativ auslösen, falls app.js noch im Login-Zustand verweilt
    const loginBtn = document.getElementById('loginBtn');
    const quickAdmin = document.querySelector('.quick-login[data-user="admin"]');
    
    if (quickAdmin) {
      quickAdmin.click();
    } else if (loginBtn) {
      loginBtn.click();
    }

    // 3. Sichtbarkeit erzwingen
    const appView = document.getElementById('appView');
    if (appView) {
      appView.classList.remove('hidden');
      appView.style.display = 'block';
    }
    const loginView = document.getElementById('loginView');
    if (loginView) {
      loginView.style.setProperty('display', 'none', 'important');
    }

    // 4. Status auf den echten Cloudflare-Admin fixieren
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // 5. Gespeicherte Daten aus Supabase laden (falls vorhanden)
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
        }
      } catch (err) {
        console.error("Supabase Sync:", err);
      }
    }

    // 6. Erstes Blatt aktiv setzen
    if (window.state) {
      const sheets = window.state.sheets || (window.WORKBOOK_DATA && window.WORKBOOK_DATA.sheets);
      if (sheets && sheets.length > 0) {
        if (!window.state.activeSheetId) {
          window.state.activeSheetId = sheets[0].id || sheets[0].name;
        }
      }
    }

    // 7. Render-Zyklus von app.js vollständig ausführen
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') {
      renderApp();
    } else if (typeof renderAll === 'function') {
      renderAll();
    } else if (typeof renderCurrentSheet === 'function') {
      renderCurrentSheet();
    } else if (typeof renderSheet === 'function' && window.state) {
      renderSheet(window.state.activeSheetId);
    }

    // 8. Klick auf den ersten Reiter simulieren, um alle Tabellencontainer zu rendern
    const activeTab = document.querySelector('#tabs .tab.active, .tabs .tab.active');
    const firstTab = document.querySelector('#tabs .tab, .tabs .tab');
    if (activeTab) {
      activeTab.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    } else if (firstTab) {
      firstTab.classList.add('active');
      firstTab.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    }

    // 9. Speichern-Funktion mit Supabase verdrahten
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

  // Sicherer Start sobald Skripte und DOM bereitstehen
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(bootFullApp, 100));
  } else {
    setTimeout(bootFullApp, 100);
  }
})();// ==========================================
// KFZ-APP PATCH-V26: AUTO-SELECT SHEET & RENDER
// ==========================================

(function() {
  // 1. Demo-Leiste und Demo-Elemente dauerhaft ausblenden
  const hideStyle = document.createElement('style');
  hideStyle.innerHTML = `
    .login-panel, .demo-bar, .demo-users { display: none !important; }
    #demoBar { display: none !important; }
  `;
  document.head.appendChild(hideStyle);

  async function initAndDisplayData() {
    // 2. Sicherstellen, dass appView sichtbar ist
    const appView = document.getElementById('appView');
    if (appView) {
      appView.classList.remove('hidden');
    }

    // 3. Echten Benutzer setzen
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
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    // Chip im Header aktualisieren
    const userChip = document.getElementById('userChip');
    if (userChip) userChip.textContent = email;

    // 4. Daten aus Supabase abrufen (falls vorhanden)
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
        }
      } catch (err) {
        console.error("Supabase Sync:", err);
      }
    }

    // 5. Erstes Sheet aktivieren und Inhalte rendern
    if (window.state) {
      // Falls noch kein aktiver Tab gesetzt ist, nimm das erste Sheet
      if (!window.state.activeSheetId && window.state.sheets && window.state.sheets.length > 0) {
        window.state.activeSheetId = window.state.sheets[0].id || window.state.sheets[0].name;
      }
    }

    // Rufe die Render-Funktionen der App auf
    if (typeof renderTabs === 'function') renderTabs();
    if (typeof renderApp === 'function') {
      renderApp();
    } else if (typeof renderAll === 'function') {
      renderAll();
    } else if (typeof renderCurrentSheet === 'function') {
      renderCurrentSheet();
    } else if (typeof renderSheet === 'function') {
      renderSheet();
    }

    // Falls ein Tab-Button existiert, diesen anklicken
    setTimeout(() => {
      const firstTab = document.querySelector('#tabs .tab, .tabs .tab');
      if (firstTab && !document.querySelector('#tabs .tab.active, .tabs .tab.active')) {
        firstTab.click();
      }
    }, 150);
  }

  // Beim Laden ausführen
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(initAndDisplayData, 200));
  } else {
    setTimeout(initAndDisplayData, 200);
  }
})();
