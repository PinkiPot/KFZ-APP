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
