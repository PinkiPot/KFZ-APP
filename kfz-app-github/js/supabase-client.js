// Supabase Konfiguration & Initialisierung
const SUPABASE_URL = "DEINE_PROJECT_URL_HIER_EINFUEGEN";
const SUPABASE_ANON_KEY = "DEIN_PUBLISHABLE_KEY_HIER_EINFUEGEN";

// Erstellt den globalen Supabase-Client
window.supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Synchronisations-Helfer: Speichert den Zustand in der Supabase-Datenbank
async function saveToSupabase(stateData) {
  if (!window.supabaseClient) return;
  const user = window.supabaseClient.auth.getUser ? (await window.supabaseClient.auth.getUser()).data.user : null;
  const userId = user ? user.id : 'demo-local-user';

  const { data, error } = await window.supabaseClient
    .from('app_state')
    .upsert({
      id: userId,
      state: stateData,
      updated_at: new Date().toISOString()
    });

  if (error) {
    console.error('Fehler beim Speichern in Supabase:', error.message);
  } else {
    console.log('Erfolgreich in Supabase gesichert.');
  }
}

// Synchronisations-Helfer: Lädt den Zustand aus Supabase
async function loadFromSupabase() {
  if (!window.supabaseClient) return null;
  const user = window.supabaseClient.auth.getUser ? (await window.supabaseClient.auth.getUser()).data.user : null;
  const userId = user ? user.id : 'demo-local-user';

  const { data, error } = await window.supabaseClient
    .from('app_state')
    .select('state')
    .eq('id', userId)
    .single();

  if (error || !data) {
    return null;
  }
  return data.state;
}
