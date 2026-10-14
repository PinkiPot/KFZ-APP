// ==========================================
// KFZ-APP PATCH-MODUL & CLOUDFLARE AUTH (WEG A)
// ==========================================

(function() {
  // 1. Cloudflare E-Mail abrufen und echte Daten aus Supabase laden
  async function initCloudflareAuth() {
    let email = "kirchbergpa@gmail.com";

    try {
      const res = await fetch('/cdn-cgi/access/get-identity');
      if (res.ok) {
        const identity = await res.json();
        if (identity && identity.email) {
          email = identity.email;
        }
      }
    } catch (e) {
      console.log("Cloudflare Access lokal oder nicht aktiv, nutze Standard-User:", email);
    }

    // Identität im Anwendungsstatus setzen
    const cleanId = 'cf_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');
    if (window.state) {
      window.state.isAuthenticated = true;
      window.state.currentUser = email;
      window.state.currentUserId = cleanId;
      window.state.currentUserRole = 'admin';
    }

    // Anzeige im Header aktualisieren
    const userDisplay = document.getElementById('user-display') || document.querySelector('.user-chip');
    if (userDisplay) {
      userDisplay.textContent = email;
    }

    // Echte Daten aus Supabase synchronisieren
    if (typeof loadFromSupabase === 'function') {
      try {
        const savedData = await loadFromSupabase();
        if (savedData && window.state) {
          Object.assign(window.state, savedData);
          console.log("Supabase-Daten erfolgreich geladen für:", email);
          if (typeof renderApp === 'function') renderApp();
          else if (typeof renderAll === 'function') renderAll();
        }
      } catch (err) {
        console.error("Fehler beim Abrufen aus Supabase:", err);
      }
    }
  }

  // 2. Speicher-Button mit Supabase verbinden
  function setupSaveHandler() {
    const saveBtn = document.getElementById('save-btn');
    if (saveBtn) {
      saveBtn.onclick = async function() {
        if (typeof saveToSupabase === 'function' && window.state) {
          saveBtn.disabled = true;
          saveBtn.textContent = '⏳ Speichern...';
          try {
            await saveToSupabase(window.state);
            saveBtn.textContent = '✅ Gespeichert!';
          } catch (e) {
            saveBtn.textContent = '❌ Fehler!';
          }
          setTimeout(() => {
            saveBtn.disabled = false;
            saveBtn.textContent = '💾 Speichern';
          }, 2000);
        }
      };
    }
  }

  // Initialisieren beim Laden der Seite
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      initCloudflareAuth();
      setupSaveHandler();
    });
  } else {
    initCloudflareAuth();
    setupSaveHandler();
  }
})();/* ---------------- V26.07.26: Hotspot-Mindestgröße, Beschreibung-Dropdown, Rechen/Farbfix, Werkzeug-Höhe, manuelle Speicherung mit Historie ---------------- */
(function(){
  'use strict';
  const VERSION_26='V26.07.26';
  const STORE_KEY_26='kfz_app_'+VERSION_26+'_state';
  const LIVE_KEY_26='kfz_app_live_state_preserved';
  const NEVER_LOSE_KEY_26='kfz_app_user_data_never_delete';
  const DRAFT_KEY_26='kfz_app_working_draft_preserved';
  const INDEX_KEY_26='kfz_app_save_history_index';
  const TODAY_26=new Date(2026,6,15);
  const MIN_HOTSPOT_PCT_26=3.0;

  function addCss26(){
    if(document.getElementById('v26_07_26_css')) return;
    const style=document.createElement('style');
    style.id='v26_07_26_css';
    style.textContent=`
      .kfz-table td,.kfz-table th{padding:3px 5px!important;line-height:1.12!important;vertical-align:middle!important;}
      .kfz-table tbody tr{height:auto!important;}
      .cell-inner{min-height:20px!important;padding:0!important;gap:3px!important;}
      .cell-text,.cell-link{line-height:1.12!important;}
      .cell-img{max-height:28px!important;max-width:58px!important;object-fit:contain!important;}
      .insert-row td{padding:1px 0!important;height:12px!important;background:transparent!important;border:0!important;}
      .insert-row button{min-width:22px!important;width:22px!important;height:18px!important;padding:0!important;border-radius:999px!important;font-size:12px!important;line-height:1!important;}
      body.sheet-werkzeug .table-wrap{overflow-x:visible!important;max-width:100%!important;}
      body.sheet-werkzeug .kfz-table{width:100%!important;table-layout:fixed!important;}
      body.sheet-werkzeug .kfz-table td,body.sheet-werkzeug .kfz-table th{height:24px!important;max-height:24px!important;padding:3px 4px!important;white-space:normal!important;word-break:break-word!important;overflow:hidden!important;text-overflow:ellipsis!important;}
      body.sheet-werkzeug .cell-inner{height:18px!important;min-height:18px!important;overflow:hidden!important;}
      .formula-gradient26{background-clip:padding-box!important;}
      .formula-gradient26 .cell-inner{position:relative;z-index:1;}
      .catalog21-hotspot{min-width:18px!important;min-height:18px!important;}
      .catalog21-hotspot[style*="width:0"],.catalog21-hotspot[style*="height:0"]{min-width:18px!important;min-height:18px!important;}
      .save-mode-badge26{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--border);border-radius:999px;padding:4px 8px;background:var(--surface2);font-size:12px;font-weight:800;}
      .save-history-list26{display:grid;gap:8px;margin-top:10px;}
      .save-history-item26{display:flex;justify-content:space-between;gap:10px;align-items:center;border:1px solid var(--border);background:var(--surface2);border-radius:12px;padding:8px;}
      .save-history-item26 strong{display:block;}
      .save-warning26{border:1px solid color-mix(in srgb,var(--warning,#f2c94c),#8a6d00 40%);background:color-mix(in srgb,var(--warning,#f2c94c) 18%, transparent);border-radius:12px;padding:9px;margin:8px 0;}
      body.dark .save-warning26{background:rgba(242,201,76,.13);}
      .description-dropdown26{font-weight:900;}
    `;
    document.head.appendChild(style);
  }

  function safeClone26(o){try{return JSON.parse(JSON.stringify(o));}catch(e){return o;}}
  function parseMaybe26(txt){try{const p=JSON.parse(txt||'null'); return p&&p.users?p:null;}catch(e){return null;}}
  function allStorageCandidates26(){
    const keys=[STORE_KEY_26,LIVE_KEY_26,NEVER_LOSE_KEY_26,DRAFT_KEY_26,'kfz_app_V26.07.26_state','kfz_app_V26.07.23_state','kfz_app_V26.07.22_state','kfz_app_V26.07.20_state'];
    try{for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i); if(k && /^kfz_app_.*state|kfz_app_live_state_preserved|kfz_app_user_data_never_delete|kfz_app_working_draft_preserved$/.test(k) && !keys.includes(k)) keys.push(k);}}catch(e){}
    const out=[];
    for(const k of keys){const p=parseMaybe26(localStorage.getItem(k)); if(p) out.push({key:k,state:p,ts:Date.parse(p.savedAt||p.updatedAt||p.createdAt||0)||0,version:String(p.version||'')});}
    out.sort((a,b)=>b.ts-a.ts);
    return out;
  }

  function collectDescriptions26(srcState){
    const vals=new Set();
    const pools=[];
    if(srcState){pools.push(...(srcState.adminSheets||[])); Object.values(srcState.userSheets||{}).forEach(arr=>pools.push(...(arr||[])));}
    pools.forEach(sheet=>{
      (sheet.themes||[]).forEach(t=>{
        (t.columns||[]).forEach(col=>{
          const label=(col.parts||[]).join(' / ').toLowerCase();
          const isDesc=label.includes('beschreibung') && !label.includes('kategorie');
          if(!isDesc) return;
          const rows=(t.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g).filter(Boolean);
          rows.forEach(r=>{const v=sheet.cells?.[`${r}:${col.col}`]?.text ?? sheet.cells?.[`${r}:${col.col}`]?.raw; if(v && String(v).trim() && String(v).trim()!=='-') vals.add(String(v).trim());});
        });
      });
    });
    return Array.from(vals).sort((a,b)=>a.localeCompare(b,'de'));
  }

  function migrate26(s){
    if(!s) return s;
    s.version=VERSION_26;
    s.savePrefs=s.savePrefs||{mode:'auto',history:true};
    if(!['auto','manual'].includes(s.savePrefs.mode)) s.savePrefs.mode='auto';
    if(s.savePrefs.history==null) s.savePrefs.history=true;
    s.saveHistoryIndex=Array.isArray(s.saveHistoryIndex)?s.saveHistoryIndex:[];
    s.dropdownLists=s.dropdownLists||{};
    const desc=collectDescriptions26(s);
    if(desc.length){
      const existing=Array.isArray(s.dropdownLists['Beschreibungen'])?s.dropdownLists['Beschreibungen']:[];
      s.dropdownLists['Beschreibungen']=Array.from(new Set([...existing,...desc])).sort((a,b)=>a.localeCompare(b,'de'));
    } else if(!s.dropdownLists['Beschreibungen']) s.dropdownLists['Beschreibungen']=[];
    s.dropdownAssignments=Array.isArray(s.dropdownAssignments)?s.dropdownAssignments:[];
    if(!s.dropdownAssignments.some(a=>String(a.match||'').toLowerCase()==='beschreibung')) s.dropdownAssignments.push({match:'Beschreibung',list:'Beschreibungen'});
    s.catalogHotspotStyle=s.catalogHotspotStyle||{};
    if(s.catalogHotspotStyle.minPct==null) s.catalogHotspotStyle.minPct=MIN_HOTSPOT_PCT_26;
    if(s.catalogHotspotStyle.borderWidth==null) s.catalogHotspotStyle.borderWidth=1;
    clampAllHotspots26(s);
    return s;
  }

  function clampAllHotspots26(srcState=state){
    const min=Number(srcState?.catalogHotspotStyle?.minPct)||MIN_HOTSPOT_PCT_26;
    const walk=(nodes)=>{(nodes||[]).forEach(n=>{if(n.previewHotspots){n.previewHotspots.forEach(h=>clampHotspot26(h,min));} if(n.children)walk(n.children);});};
    const sheets=[...(srcState?.adminSheets||[]),...Object.values(srcState?.userSheets||{}).flat()];
    sheets.forEach(s=>{
      if(s.catalog){ if(s.tree)walk(s.tree); Object.values(s.vehicleCatalogs||{}).forEach(b=>walk(b.tree)); }
    });
  }
  function clampHotspot26(h,min=MIN_HOTSPOT_PCT_26){
    h.x=Math.max(0,Math.min(99,Number(h.x)||0)); h.y=Math.max(0,Math.min(99,Number(h.y)||0));
    h.w=Math.max(min,Math.min(100-h.x,Number(h.w)||min)); h.h=Math.max(min,Math.min(100-h.y,Number(h.h)||min));
  }

  const baseInit26=initState;
  initState=function(){
    let best=allStorageCandidates26()[0]?.state||null;
    if(!best) best=baseInit26();
    return migrate26(best);
  };

  function writeLocal26(key,obj){try{localStorage.setItem(key,JSON.stringify(obj)); return true;}catch(e){console.warn('Speichern fehlgeschlagen',key,e); return false;}}
  function idb26(){return new Promise((resolve,reject)=>{try{const req=indexedDB.open('kfz_app_save_history_db',1); req.onupgradeneeded=()=>{const db=req.result; if(!db.objectStoreNames.contains('snapshots')) db.createObjectStore('snapshots',{keyPath:'id'});}; req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);}catch(e){reject(e);}});}
  async function idbPut26(snapshot){const db=await idb26(); return new Promise((resolve,reject)=>{const tx=db.transaction('snapshots','readwrite'); tx.objectStore('snapshots').put(snapshot); tx.oncomplete=()=>resolve(true); tx.onerror=()=>reject(tx.error);});}
  async function idbGet26(id){const db=await idb26(); return new Promise((resolve,reject)=>{const tx=db.transaction('snapshots','readonly'); const r=tx.objectStore('snapshots').get(id); r.onsuccess=()=>resolve(r.result); r.onerror=()=>reject(r.error);});}
  async function idbDelete26(id){const db=await idb26(); return new Promise((resolve,reject)=>{const tx=db.transaction('snapshots','readwrite'); tx.objectStore('snapshots').delete(id); tx.oncomplete=()=>resolve(true); tx.onerror=()=>reject(tx.error);});}

  function saveDraft26(){
    if(!state) return;
    state.version=VERSION_26; state.updatedAt=new Date().toISOString();
    writeLocal26(DRAFT_KEY_26,state);
    writeLocal26(NEVER_LOSE_KEY_26,state);
  }
  async function manualCommit26(reason='Manuelle Speicherung'){
    if(!state) return;
    state.version=VERSION_26; state.savedAt=new Date().toISOString(); state.updatedAt=state.savedAt;
    writeLocal26(STORE_KEY_26,state); writeLocal26(LIVE_KEY_26,state); writeLocal26(NEVER_LOSE_KEY_26,state); writeLocal26(DRAFT_KEY_26,state);
    const prefs=state.savePrefs||{};
    if(prefs.history!==false){
      const id='snap_'+Date.now()+'_'+Math.random().toString(36).slice(2,8);
      const meta={id,version:VERSION_26,savedAt:state.savedAt,userId:state.currentUserId||'',reason};
      state.saveHistoryIndex=[meta,...(state.saveHistoryIndex||[])].slice(0,20);
      try{await idbPut26({id,meta,state:safeClone26(state)});}catch(e){console.warn('Historie konnte nicht in IndexedDB gespeichert werden',e);}
      writeLocal26(INDEX_KEY_26,state.saveHistoryIndex);
      writeLocal26(STORE_KEY_26,state); writeLocal26(LIVE_KEY_26,state); writeLocal26(NEVER_LOSE_KEY_26,state);
    }
    // Nach: writeLocal26(NEVER_LOSE_KEY_26,state);
try {
  if (typeof saveToSupabase === 'function') {
    await saveToSupabase(state);
  }
} catch (err) {
  console.warn('Supabase Sync Fehler:', err);
}
    updateSaveInfo26(); toast('Gespeichert · Historie aktualisiert');
  }

  const baseSave26=save;
  save=function(){
    if(!state) return;
    state.version=VERSION_26; state.updatedAt=new Date().toISOString();
    const mode=state.savePrefs?.mode||'auto';
    if(mode==='manual'){
      state.unsavedChanges=true;
      saveDraft26();
      updateSaveInfo26();
      return;
    }
    state.unsavedChanges=false; state.savedAt=new Date().toISOString();
    writeLocal26(STORE_KEY_26,state); writeLocal26(LIVE_KEY_26,state); writeLocal26(NEVER_LOSE_KEY_26,state); writeLocal26(DRAFT_KEY_26,state);
    
    // --- HIER NEU EINGEFÜGT FÜR SUPABASE: ---
    if (typeof saveToSupabase === 'function' && state.savePrefs?.mode !== 'manual') {
      saveToSupabase(state).catch(e => console.warn(e));
    }
    // ----------------------------------------

    updateSaveInfo26();
  };
   
    state.unsavedChanges=false; state.savedAt=new Date().toISOString();
    writeLocal26(STORE_KEY_26,state); writeLocal26(LIVE_KEY_26,state); writeLocal26(NEVER_LOSE_KEY_26,state); writeLocal26(DRAFT_KEY_26,state);
    updateSaveInfo26();
  };
  function updateSaveInfo26(){
    const e=document.querySelector('#saveInfo'); if(!e||!state) return;
    const mode=state.savePrefs?.mode==='manual'?'Manuell':'Auto';
    const saved=state.savedAt?new Date(state.savedAt).toLocaleTimeString('de-DE'):'noch nicht gespeichert';
    e.textContent=(state.unsavedChanges?'Ungespeichert · ':'')+mode+' · '+saved;
  }
  updateSaveInfo=updateSaveInfo26;

  function parseVal26(v){if(v==null)return NaN; let s=String(v).trim(); if(!s||s==='-'||s==='—')return NaN; s=s.replace(/\s/g,'').replace(/\./g,'').replace(',', '.').replace('%',''); const n=Number(s); return Number.isFinite(n)?n:NaN;}
  function parseDate26(v){
    if(v instanceof Date && !isNaN(v)) return v;
    if(v==null) return null; let s=String(v).trim(); if(!s||s==='-'||s==='—') return null;
    let m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/); if(m){let y=+m[3]; if(y<100)y+=2000; return new Date(y,+m[2]-1,+m[1]);}
    m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); if(m) return new Date(+m[1],+m[2]-1,+m[3]);
    m=s.match(/^(\d{1,2})\/(\d{4})$/); if(m) return new Date(+m[2],+m[1]-1,1);
    const d=new Date(s); return isNaN(d)?null:d;
  }
  function days26(a,b){return Math.round((b-a)/(86400000));}
  function path26(col){return (col?.parts||[]).join(' / ').toLowerCase();}
  function short26(col){return (typeof shortHeader==='function'?shortHeader(col.parts):(col.parts||[]).join(' / ')).toLowerCase();}
  function fmtPct26(n){return (Math.max(0,Math.min(100,n))).toFixed(2).replace('.',',')+'%';}
  function findCol26(cols,pred){return (cols||[]).find(c=>pred(path26(c),short26(c),c));}
  recalcCellIfNeeded=function(sheet,t,r,col,cell){
    const p=path26(col), h=short26(col), cols=t.columns||[];
    const get=cc=>cc?sheet.cells?.[`${r}:${cc.col}`]?.text ?? sheet.cells?.[`${r}:${cc.col}`]?.raw:'';
    const installDate=findCol26(cols,(p,h)=>h.includes('einbaudatum')||p.includes('einbaudatum'));
    const installKm=findCol26(cols,(p,h)=>h.includes('kilometerstand')||p.includes('kilometerstand'));
    const nextDate=findCol26(cols,(p,h)=>p.includes('nächster tausch')&&h.includes('datum'));
    const nextKm=findCol26(cols,(p,h)=>p.includes('nächster tausch')&&h.includes('laufleistung')&&!h.includes('tage')&&!h.includes('anteil'));
    if(h.includes('tage bis')){
      const d=parseDate26(get(nextDate));
      if(d){ const diff=days26(TODAY_26,d); cell.text=diff<0?'':String(diff); cell.raw=cell.text; return; }
      const base=parseVal26(get(installKm)), target=parseVal26(get(nextKm));
      if(Number.isFinite(base)&&Number.isFinite(target)&&target>0&&sheet.mileage){ const diffKm=(base+target)-Number(sheet.mileage||0); cell.text=diffKm<0?'':String(Math.round(diffKm/50)); cell.raw=cell.text; return; }
      cell.text='-'; cell.raw='-'; return;
    }
    if(h.includes('anteil bis')){
      const d0=parseDate26(get(installDate)), d1=parseDate26(get(nextDate));
      if(d0&&d1&&d1>d0){ const pct=days26(d0,TODAY_26)/days26(d0,d1)*100; cell.text=fmtPct26(pct); cell.raw=cell.text; return; }
      const base=parseVal26(get(installKm)), target=parseVal26(get(nextKm));
      if(Number.isFinite(base)&&Number.isFinite(target)&&target>0&&sheet.mileage){ const pct=(Number(sheet.mileage||0)-base)/target*100; cell.text=fmtPct26(pct); cell.raw=cell.text; return; }
      cell.text='-'; cell.raw='-'; return;
    }
    // Simple Excel style formula support for Laufleistung cells like =$M$2/230000.
    if(p.includes('nächster tausch')&&h.includes('laufleistung')&&cell.formula){
      const m=String(cell.formula).match(/\$?M\$?2\s*\/\s*([0-9.]+)/i);
      if(m){ const denom=parseVal26(m[1]); if(denom>0&&sheet.mileage){ const pct=Number(sheet.mileage)/denom*100; cell.text=fmtPct26(pct); cell.raw=cell.text; } }
    }
  };

  const baseDropdown26=dropdownForColumn;
  dropdownForColumn=function(col){
    const p=(col?.parts||[]).map(x=>String(x||'').toLowerCase());
    const joined=p.join(' / ');
    const last=p[p.length-1]||'';
    if((last==='beschreibung' || joined.endsWith('beschreibung / beschreibung')) && state?.dropdownLists?.['Beschreibungen']){
      return {list:'Beschreibungen',values:state.dropdownLists['Beschreibungen']||[]};
    }
    return baseDropdown26(col);
  };

  function colorForPct26(pct){
    pct=Math.max(0,Math.min(100,Number(pct)||0));
    if(pct<55) return 'rgba(46, 160, 67, .22)';
    if(pct<80) return 'rgba(242, 201, 76, .26)';
    return 'rgba(220, 53, 69, .24)';
  }
  function applyDynamicColors26(){
    const sheet=typeof ensureActive==='function'?ensureActive():null; if(!sheet||!sheet.themes) return;
    const themeByKey=new Map((sheet.themes||[]).map(t=>[String(t.key||topicKey(t)),t]));
    document.body.classList.toggle('sheet-werkzeug',/werkzeug/i.test(sheet.name||''));
    document.querySelectorAll('.topic-card').forEach(card=>{
      const t=themeByKey.get(card.dataset.topic); if(!t) return;
      card.querySelectorAll('td[data-r][data-c]').forEach(td=>{
        const c=+td.dataset.c, r=+td.dataset.r; const col=(t.columns||[]).find(x=>Number(x.col)===c); if(!col) return;
        const cell=sheet.cells?.[`${r}:${c}`]; const h=short26(col); const val=cell?.text ?? cell?.raw ?? td.textContent;
        td.classList.remove('formula-gradient26'); td.style.removeProperty('background');
        if(h.includes('anteil bis')){
          const pct=parseVal26(val); if(Number.isFinite(pct)){td.classList.add('formula-gradient26');td.style.background=`linear-gradient(90deg, ${colorForPct26(pct)} 0%, ${colorForPct26(pct)} ${Math.max(0,Math.min(100,pct))}%, var(--cellBg,var(--surface)) ${Math.max(0,Math.min(100,pct))}%)`;}
        } else if(h.includes('tage bis')){
          const n=parseVal26(val); if(val===''||n<0) td.style.background='rgba(220,53,69,.18)'; else if(Number.isFinite(n)&&n<45) td.style.background='rgba(242,201,76,.22)'; else if(Number.isFinite(n)) td.style.background='rgba(46,160,67,.16)';
        } else if(/verfügbar|verfuegbar|zulauf|gekauft/.test(h)){
          const n=parseVal26(val); if(Number.isFinite(n)) td.style.background=n<=0?'rgba(220,53,69,.15)':'rgba(46,160,67,.13)';
        }
      });
    });
  }

  function applyHotspotVars26(){
    if(!state) return; const hs=state.catalogHotspotStyle||{}; const root=document.documentElement;
    const bw=(Number(hs.borderWidth)>=0?Number(hs.borderWidth):1);
    root.style.setProperty('--hotspot-border-width',bw+'px');
    if(bw===0){root.style.setProperty('--hotspot-border','transparent'); root.style.setProperty('--hotspot-active-border','transparent');}
    document.querySelectorAll('.catalog21-hotspot').forEach(el=>{
      const id=el.dataset.hotspotId; const sheet=ensureActive?.(); if(!id||!sheet) return;
      const all=[]; const walk=n=>{(n||[]).forEach(x=>{if(x.previewHotspots)all.push(...x.previewHotspots); if(x.children)walk(x.children);});};
      if(sheet.tree)walk(sheet.tree); Object.values(sheet.vehicleCatalogs||{}).forEach(b=>walk(b.tree));
      const h=all.find(x=>x.id===id); if(h){clampHotspot26(h,Number(hs.minPct)||MIN_HOTSPOT_PCT_26); el.style.width=h.w+'%'; el.style.height=h.h+'%';}
    });
  }

  function openHotspotStyle26(){
    const hs=state.catalogHotspotStyle=state.catalogHotspotStyle||{};
    const bw=Number(hs.borderWidth)>=0?Number(hs.borderWidth):1;
    const min=Number(hs.minPct)||MIN_HOTSPOT_PCT_26;
    const opacity=Number(hs.opacity)||0.14;
    showModal('Hotspot-Optik',`<div class="grid2">
      <div class="field"><label>Füllfarbe</label><input id="hsFill26" type="color" value="${esc(hs.fillColor||'#f2c94c')}"></div>
      <div class="field"><label>Rahmenfarbe</label><input id="hsBorder26" type="color" value="${esc(hs.borderColor||'#f2c94c')}"></div>
      <div class="field"><label>Transparenz</label><input id="hsOpacity26" type="range" min="0.03" max="0.45" step="0.01" value="${opacity}"><span id="hsOpacityOut26">${Math.round(opacity*100)}%</span></div>
      <div class="field"><label>Rahmenstärke</label><input id="hsBorderWidth26" type="range" min="0" max="4" step="1" value="${bw}"><span id="hsBorderOut26">${bw===0?'kein Rahmen':bw+'px'}</span></div>
      <div class="field"><label>Mindestgröße</label><input id="hsMin26" type="range" min="3" max="12" step="0.5" value="${min}"><span id="hsMinOut26">${min}%</span></div>
    </div><p class="hint">Rahmenstärke 0 entfernt den Rahmen komplett. Die Mindestgröße verhindert zu kleine, kaum anklickbare Bereiche.</p>`,[
      {text:'Speichern',action:()=>{const s=state.catalogHotspotStyle=state.catalogHotspotStyle||{}; s.fillColor=$('#hsFill26').value; s.borderColor=$('#hsBorder26').value; s.opacity=Number($('#hsOpacity26').value); s.activeOpacity=Math.min(.7,s.opacity+.16); s.borderWidth=Number($('#hsBorderWidth26').value); s.minPct=Number($('#hsMin26').value)||MIN_HOTSPOT_PCT_26; clampAllHotspots26(state); save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    const upd=()=>{const op=$('#hsOpacity26'),bw=$('#hsBorderWidth26'),mn=$('#hsMin26'); if(op)$('#hsOpacityOut26').textContent=Math.round(Number(op.value)*100)+'%'; if(bw)$('#hsBorderOut26').textContent=Number(bw.value)===0?'kein Rahmen':Number(bw.value)+'px'; if(mn)$('#hsMinOut26').textContent=Number(mn.value)+'%';};
    ['#hsOpacity26','#hsBorderWidth26','#hsMin26'].forEach(sel=>$(sel)?.addEventListener('input',upd));
  }

  function wireHotspotStyle26(){
    const old=document.getElementById('hotspotStyle25'); if(old){old.onclick=(e)=>{e.preventDefault(); e.stopPropagation(); openHotspotStyle26();}; old.textContent='Hotspot-Optik';}
  }

  function renderSaveSettings26(){
    const prefs=state.savePrefs=state.savePrefs||{mode:'auto',history:true};
    const idx=state.saveHistoryIndex||[];
    return `<h3>Speicherung & Versionshistorie</h3>
      <p class="hint">Wie in der LH-App kann zwischen automatischer und manueller Speicherung gewechselt werden. Nutzerdaten werden zusätzlich versionneutral gesichert, damit neue Demo-Versionen sie nicht verwerfen.</p>
      <div class="save-warning26"><strong>Datenerhalt:</strong> Diese Demo schreibt aktuelle Daten zusätzlich in einen versionsunabhängigen Speicher. Neue Versionen laden diesen Bestand bevorzugt und migrieren ihn weiter.</div>
      <div class="grid2">
        <div class="field"><label>Speichermodus</label><select id="saveMode26"><option value="auto" ${prefs.mode==='auto'?'selected':''}>Automatisch speichern</option><option value="manual" ${prefs.mode==='manual'?'selected':''}>Manuell speichern</option></select></div>
        <div class="field"><label>Historie</label><select id="saveHistory26"><option value="1" ${prefs.history!==false?'selected':''}>aktiv</option><option value="0" ${prefs.history===false?'selected':''}>deaktiviert</option></select></div>
      </div>
      <div style="margin:10px 0;display:flex;gap:8px;flex-wrap:wrap"><button id="manualSave26">Jetzt speichern + Historie</button><button class="secondary" id="snapshot26">Historiepunkt erstellen</button><button class="secondary" id="exportState26">Sicherung exportieren</button><label class="buttonlike secondary" style="display:inline-flex;align-items:center;gap:6px;cursor:pointer">Sicherung importieren<input id="importState26" type="file" accept="application/json,.json" hidden></label></div>
      <h4>Wiederherstellen</h4>
      <div class="save-history-list26">${idx.length?idx.map(m=>`<div class="save-history-item26"><div><strong>${esc(m.reason||'Speicherpunkt')}</strong><span class="hint">${esc(m.version||'')} · ${esc(m.savedAt?new Date(m.savedAt).toLocaleString('de-DE'):'')}</span></div><div><button class="small secondary" data-restore26="${esc(m.id)}">Wiederherstellen</button><button class="small danger" data-del-snap26="${esc(m.id)}">×</button></div></div>`).join(''):'<p class="hint">Noch keine Historie gespeichert.</p>'}</div>`;
  }

  const baseRenderSettingsPanel26=renderSettingsPanel;
  renderSettingsPanel=function(panel){
    if(panel==='daten'){
      const el=document.querySelector('#settingsPanel'); if(el) el.innerHTML=renderSaveSettings26();
      bindSaveSettings26(); return;
    }
    baseRenderSettingsPanel26(panel);
  };
  const baseOpenSettings26=openSettings;
  openSettings=function(panel='profil'){
    baseOpenSettings26(panel);
    const nav=document.querySelector('.settings-nav');
    if(nav && !nav.querySelector('[data-set-panel="daten"]')){
      const b=document.createElement('button'); b.dataset.setPanel='daten'; b.textContent='Speicherung'; b.className=panel==='daten'?'active':''; b.onclick=()=>openSettings('daten'); nav.appendChild(b);
    }
    if(panel==='daten') renderSettingsPanel('daten');
  };
  function bindSaveSettings26(){
    $('#saveMode26')?.addEventListener('change',e=>{state.savePrefs=state.savePrefs||{}; state.savePrefs.mode=e.target.value; save(); toast(e.target.value==='manual'?'Manuelle Speicherung aktiv':'Automatische Speicherung aktiv'); render();});
    $('#saveHistory26')?.addEventListener('change',e=>{state.savePrefs=state.savePrefs||{}; state.savePrefs.history=e.target.value==='1'; save(); openSettings('daten');});
    $('#manualSave26')?.addEventListener('click',()=>manualCommit26('Manuelle Speicherung'));
    $('#snapshot26')?.addEventListener('click',()=>manualCommit26('Manueller Historiepunkt'));
    $('#exportState26')?.addEventListener('click',()=>download('kfz_app_sicherung_'+VERSION_26+'.json',JSON.stringify(state,null,2),'application/json'));
    $('#importState26')?.addEventListener('change',e=>{const f=e.target.files[0]; if(!f)return; const rd=new FileReader(); rd.onload=()=>{try{const p=JSON.parse(rd.result); if(!p.users) throw new Error('Keine App-Sicherung'); state=migrate26(p); manualCommit26('Importierte Sicherung'); closeModal(); render(); toast('Sicherung importiert.');}catch(err){toast('Import fehlgeschlagen: '+err.message);}}; rd.readAsText(f);});
    $$('[data-restore26]').forEach(b=>b.onclick=async()=>{const snap=await idbGet26(b.dataset.restore26).catch(()=>null); if(!snap?.state){toast('Speicherpunkt nicht gefunden.');return;} if(confirm('Diesen Stand wiederherstellen? Aktueller Stand wird vorher als Historiepunkt gespeichert.')){await manualCommit26('Vor Wiederherstellung gesichert'); state=migrate26(snap.state); await manualCommit26('Wiederhergestellt: '+(snap.meta?.savedAt||'')); closeModal(); render();}});
    $$('[data-del-snap26]').forEach(b=>b.onclick=async()=>{if(confirm('Historiepunkt löschen?')){await idbDelete26(b.dataset.delSnap26).catch(()=>{}); state.saveHistoryIndex=(state.saveHistoryIndex||[]).filter(x=>x.id!==b.dataset.delSnap26); save(); openSettings('daten');}});
  }

  function wireSaveButtons26(){
    document.querySelectorAll('#saveBtn').forEach(btn=>{btn.textContent=state?.savePrefs?.mode==='manual'?'Speichern':'Speichern'; btn.onclick=(e)=>{e.preventDefault(); manualCommit26('Button Speichern');};});
  }

  const baseRender26=render;
  render=function(){
    baseRender26();
    addCss26(); migrate26(state); clampAllHotspots26(state); applyHotspotVars26(); wireHotspotStyle26(); wireSaveButtons26(); applyDynamicColors26();
    document.title='KFZ-APP '+VERSION_26;
    const chip=document.querySelector('.version-chip'); if(chip) chip.textContent=VERSION_26;
    updateSaveInfo26();
  };
  const baseApplyTheme26=applyTheme;
  applyTheme=function(){baseApplyTheme26(); addCss26(); applyHotspotVars26();};

  const baseBindSheetEvents26=bindSheetEvents;
  bindSheetEvents=function(sheet,access){
    baseBindSheetEvents26(sheet,access);
    const mi=$('#mileageInput');
    if(mi){
      mi.addEventListener('blur',()=>{autosaveRecalc?.(sheet); setTimeout(applyDynamicColors26,0);},true);
    }
  };

  const baseReleaseNotes26=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.26 · Speicherung, Berechnung und Tabellenkompaktheit',body:'Hotspots besitzen jetzt eine Mindestgröße, der Rahmen kann komplett entfernt werden, Beschreibungsspalten haben wieder Dropdowns, Laufleistungs-/Verarbeitungswerte und Farbverläufe werden nach Änderungen neu angewendet, Werkzeug-Zeilen sind kompakt wie die anderen Tabellen, und es gibt automatische/manuelle Speicherung mit wiederherstellbarer Historie sowie versionsunabhängigem Datenerhalt.',jump:'settings:daten'}, ...baseReleaseNotes26()];};

  // Extra Sicherheit: auch wenn die Datei nach einem alten Zustand startet, nie Nutzerdaten wegwerfen.
  window.addEventListener('beforeunload',()=>{try{saveDraft26();}catch(e){}});
  document.addEventListener('DOMContentLoaded',()=>{addCss26(); setTimeout(()=>{try{migrate26(state); wireSaveButtons26(); applyDynamicColors26(); applyHotspotVars26(); updateSaveInfo26();}catch(e){console.warn(e)}},150);});
})();
