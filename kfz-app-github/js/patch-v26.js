// =========================================================================
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
