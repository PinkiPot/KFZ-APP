// ==========================================
// KFZ-APP PATCH-V26: CLOUDFLARE AUTH & UNHIDE APP
// ==========================================

(function() {
  // 1. Demo-Leiste und Dummy-Logins unsichtbar schalten
  const hideStyle = document.createElement('style');
  hideStyle.innerHTML = `
    .login-panel, .demo-bar, .demo-users { display: none !important; }
    #demoBar { display: none !important; }
  `;
  document.head.appendChild(hideStyle);

  // 2. Hauptanwendung sichtbar schalten und initialisieren
  async function unlockAndRenderApp() {
    // Entfernt das 'hidden' von appView, damit die Seite nicht mehr weiß ist
    const appView = document.getElementById('appView');
    if (appView) {
      appView.classList.remove('hidden');
    }

    // Demo-Login-Container ausblenden, falls noch im DOM
    const loginPanels = document.querySelectorAll('.login-panel');
    loginPanels.forEach(el => el.style.setProperty('display', 'none', 'important'));

    // Echte Cloudflare-Identität ermitteln
    let email = "kirchbergpa@gmail.com";
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {
      console.log("Cloudflare Access lokal/Fallback:", e);
    }

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // App-Status auf echten Admin setzen
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    // Benutzer-Chip im Header mit deiner echten E-Mail beschriften
    const userChip = document.getElementById('userChip') || document.querySelector('.user-chip');
    if (userChip) {
      userChip.textContent = email;
    }

    // Echte Daten aus Supabase laden und Anzeige aktualisieren
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
        }
      } catch (err) {
        console.error("Fehler beim Laden aus Supabase:", err);
      }
    }

    // Render-Funktion der Original-App aufrufen
    if (typeof renderApp === 'function') {
      renderApp();
    } else if (typeof renderAll === 'function') {
      renderAll();
    } else if (typeof renderTabs === 'function') {
      renderTabs();
    }
  }

  // Direkt beim Starten ausführen
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', unlockAndRenderApp);
  } else {
    unlockAndRenderApp();
  }
})();// ==========================================
// KFZ-APP PATCH-V26: AUTO-LOGIN & CLOUDFLARE SYNC
// ==========================================

(function() {
  // 1. Demo-Leiste und Demo-Elemente sofort ausblenden
  const hideStyle = document.createElement('style');
  hideStyle.innerHTML = `
    .demo-bar, .demo-users { display: none !important; }
    #login-panel { display: none !important; }
  `;
  document.head.appendChild(hideStyle);

  // 2. Automatischer Login-Durchstich beim Start
  async function performAutoLoginAndInit() {
    let email = "kirchbergpa@gmail.com";

    // Versuche echte E-Mail von Cloudflare Access abzufragen
    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const idData = await res.json();
        if (idData && idData.email) email = idData.email;
      }
    } catch(e) {
      console.log("Cloudflare Identity Header Fallback:", e);
    }

    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');

    // Falls die Original-App eine Login-Funktion bereitstellt (z.B. doLogin oder login)
    if (typeof window.doLogin === 'function') {
      window.doLogin('admin', 'admin');
    } else if (typeof window.login === 'function') {
      window.login('admin', 'admin');
    } else {
      // Formularfelder automatisch befüllen und Klick simulieren
      const userInp = document.getElementById('login-user') || document.querySelector('input[type="text"]');
      const passInp = document.getElementById('login-pass') || document.querySelector('input[type="password"]');
      const loginBtn = document.getElementById('btn-login') || document.querySelector('.login-panel button') || document.querySelector('button[type="submit"]');

      if (userInp && passInp && loginBtn) {
        userInp.value = 'admin';
        passInp.value = 'admin';
        loginBtn.click();
      }
    }

    // Status auf den echten Cloudflare-Nutzer überschreiben
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    // Echte Daten aus Supabase nachladen
    if (typeof loadFromSupabase === 'function') {
      try {
        const cloudData = await loadFromSupabase();
        if (cloudData && window.state) {
          Object.assign(window.state, cloudData);
          if (typeof renderApp === 'function') renderApp();
          else if (typeof renderAll === 'function') renderAll();
          else if (typeof renderSheet === 'function') renderSheet();
        }
      } catch(err) {
        console.error("Supabase Sync Fehler:", err);
      }
    }

    // E-Mail-Chip in der Menüleiste aktualisieren
    const userChip = document.querySelector('.user-chip');
    if (userChip) {
      userChip.textContent = email;
    }
  }

  // Sobald das DOM bereit ist, automatisch ausführen
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(performAutoLoginAndInit, 100));
  } else {
    setTimeout(performAutoLoginAndInit, 100);
  }
})();
