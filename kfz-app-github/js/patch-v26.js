// ==========================================
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
