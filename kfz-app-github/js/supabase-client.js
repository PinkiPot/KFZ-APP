// Supabase Konfiguration & Initialisierung
const SUPABASE_URL = "NEXT_PUBLIC_SUPABASE_URL=https://uxosshpiidqlyhxyjfol.supabase.co";
const SUPABASE_ANON_KEY = "b_publishable__bctjDJczPGfCllEvfLXfg_T_RouA_4";

// Erstellt den globalen Supabase-Client
window.supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Ermittelt die E-Mail des angemeldeten Nutzers (Cloudflare Access oder Fallback)
async function getActiveUserIdentifier() {
  try {
    // 1. Prüfen, ob Cloudflare Access den Identitäts-Header bereitstellt
    const res = await fetch('/cdn-cgi/access/get-identity');
    if (res.ok) {
      const data = await res.json();
      if (data && data.email) {
        return 'cf_' + data.email.toLowerCase().replace(/[^a-z0-9]/g, '_');
      }
    }
  } catch (e) {
    // Lokale Umgebung oder kein Cloudflare Access aktiv
  }
  
  // 2. Fallback: Aktueller lokaler Demo-Nutzer
  return (window.state && window.state.currentUserId) ? window.state.currentUserId : 'standard_user';
}

// Synchronisations-Helfer: Speichert den Zustand in der Supabase-Datenbank
async function saveToSupabase(stateData) {
  if (!window.supabaseClient) return;
  const recordId = await getActiveUserIdentifier();

  const { data, error } = await window.supabaseClient
    .from('app_state')
    .upsert({
      id: recordId,
      state: stateData,
      updated_at: new Date().toISOString()
    });

  if (error) {
    console.error('Fehler beim Speichern in Supabase:', error.message);
  } else {
    console.log('Erfolgreich in Supabase gesichert für:', recordId);
  }
}

// Synchronisations-Helfer: Lädt den Zustand aus Supabase
async function loadFromSupabase() {
  if (!window.supabaseClient) return null;
  const recordId = await getActiveUserIdentifier();

  const { data, error } = await window.supabaseClient
    .from('app_state')
    .select('state')
    .eq('id', recordId)
    .single();

  if (error || !data) {
    return null;
  }
  return data.state;
}
// =========================================================================
// DATENSCHUTZ- & MIGRATIONS-ROUTINE GEGEN DATENVERLUST BEI UPDATES
// =========================================================================

let isDataLoadedSuccessfully = false; // Sync-Lock-Flag

/**
 * Tiefes Zusammenführen: Behält alle existierenden Admin-Einstellungen
 * und ergänzt nur neue Felder aus Code-Updates.
 */
function preserveAdminData(currentSaved, newDefaults) {
  if (!currentSaved) return newDefaults;
  
  const merged = { ...newDefaults, ...currentSaved };

  // Spezieller Schutz für das Einstellungs-Objekt
  if (newDefaults.settings || currentSaved.settings) {
    merged.settings = {
      ...(newDefaults.settings || {}),
      ...(currentSaved.settings || {}) // Existierende Admin-Werte gewinnen IMMER
    };
  }

  // Spezieller Schutz für Tabellen und Reiter
  if (currentSaved.sheets && currentSaved.sheets.length > 0) {
    merged.sheets = currentSaved.sheets; // Niemals durch Code-Defaults ersetzen
  }

  return merged;
}

/**
 * Abgesichertes Laden beim Start
 */
async function loadUserDataSafely() {
  isDataLoadedSuccessfully = false; // Speichern sperren
  
  try {
    const { data, error } = await supabaseClient
      .from('kfz_app_state')
      .select('payload')
      .eq('user_id', window.state.currentUserId)
      .maybeSingle();

    if (error) {
      console.error("Fehler beim Laden aus Supabase:", error);
      // Kritischer Abbruch: Nicht speichern, um Datenverlust zu verhindern
      return;
    }

    if (data && data.payload) {
      // Vorhandene Daten gefunden: Admin-Daten schonend mit App-State verschmelzen
      window.state = preserveAdminData(data.payload, window.state);
      console.log("Admin-Daten & Einstellungen erfolgreich beibehalten.");
    } else {
      console.log("Keine Vorabdaten gefunden. Erstinstallation aktiv.");
    }

    // Erst NACH erfolgreicher Synchronisation wird die Speicherfunktion freigegeben
    isDataLoadedSuccessfully = true;

  } catch (err) {
    console.error("Netzwerkfehler beim Laden:", err);
    // isDataLoadedSuccessfully bleibt false -> Kein Überschreiben möglich!
  }
}

/**
 * Abgesichertes Speichern
 */
async function saveUserDataSafely() {
  if (!isDataLoadedSuccessfully) {
    alert("Sicherheitswarnung: Speichern abgebrochen, da die aktuellen Daten noch nicht sauber geladen wurden. Bitte Seite neu laden.");
    return false;
  }

  try {
    const { error } = await supabaseClient
      .from('kfz_app_state')
      .upsert({
        user_id: window.state.currentUserId,
        payload: window.state,
        updated_at: new Date().toISOString()
      }, { onConflict: 'user_id' });

    if (error) throw error;
    return true;
  } catch (err) {
    console.error("Fehler beim Speichern:", err);
    return false;
  }
}
