'use strict';
const VERSION = 'V26.07.16';
const STORE_KEY = 'kfz_app_' + VERSION + '_state';
const TODAY = new Date(2026, 6, 13);
const EMPTY = '—';
const DEFAULT_LOGO = 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#1f4e79"/><stop offset="1" stop-color="#d9eaf7"/></linearGradient></defs><rect width="128" height="128" rx="26" fill="url(#g)"/><path d="M24 78h8l10-22c3-7 9-11 17-11h28c8 0 14 4 17 11l10 22h8v20h-12a14 14 0 0 1-28 0H47a14 14 0 0 1-28 0H12V84c0-3 2-6 5-6h7Zm28-10h48l-5-12c-1-3-4-5-8-5H59c-4 0-7 2-8 5l-5 12Zm-19 34a6 6 0 1 0 0-12 6 6 0 0 0 0 12Zm63 0a6 6 0 1 0 0-12 6 6 0 0 0 0 12Z" fill="white"/></svg>`);
const USERS = [
  {id:'admin',name:'Admin',username:'admin',email:'admin@kfz.local',password:'admin',role:'Admin',abo:'Adminzugang'},
  {id:'max',name:'Max Mustermann',username:'max',email:'max@kfz.local',password:'demo',role:'Nutzer',abo:'Testphase'},
  {id:'lea',name:'Lea Schrauber',username:'lea',email:'lea@kfz.local',password:'demo',role:'Nutzer',abo:'Monatlich'},
  {id:'gast',name:'Gast Ansicht',username:'gast',email:'gast@kfz.local',password:'demo',role:'Gast',abo:'Gast/Test'}
];
let state, tooltipEl, editingCell=null, dragged=null;
const $=sel=>document.querySelector(sel); const $$=sel=>Array.from(document.querySelectorAll(sel));
function clone(o){return JSON.parse(JSON.stringify(o));} function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function uid(p='id'){return p+'_'+Math.random().toString(36).slice(2,8)+Date.now().toString(36).slice(-4)}
function slug(s){return String(s||'').toLowerCase().replace(/[^a-z0-9äöüß]+/gi,'_').replace(/^_+|_+$/g,'')||'x'}
function rank(v){return {none:0,'':0,view:1,edit:2,manage:3}[v||'']||0} function canAccess(level,need){return rank(level)>=rank(need)}
function colLetter(n){let s='';while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-m)/26)}return s}
function currentUser(){return state.users.find(u=>u.id===state.currentUserId)||null}
function sheetId(name,owner='system'){return 'sh_'+slug(name)+'_'+owner}
function topicKey(t){return slug(t.displayName||t.name)+'_'+t.startCol+'_'+t.endCol}
function isVehicle(sheet){return !!sheet.vehicle}
function formatKm(value){let digits=String(value??'').replace(/[^0-9]/g,''); if(!digits) return ''; return Number(digits).toLocaleString('de-DE')}
function rawKm(value){return Number(String(value??'').replace(/[^0-9]/g,''))||0}
function parseNum(v){if(v==null)return NaN; let s=String(v).trim().replace(/\./g,'').replace(',', '.').replace('%',''); if(!s||s==='-')return NaN; return Number(s)}
function parseDate(v){if(!v||v==='-')return null; let s=String(v).trim(); let m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/); if(m){let y=+m[3]; if(y<100)y+=2000; return new Date(y,+m[2]-1,+m[1])} m=s.match(/^(\d{1,2})\/(\d{4})$/); if(m){return new Date(+m[2],+m[1]-1,1)} m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); if(m){return new Date(+m[1],+m[2]-1,+m[3])} return null}
function daysBetween(a,b){return Math.round((b-a)/(1000*60*60*24))}
function toast(msg){const s=$('#toastStack'); if(!s)return; const d=document.createElement('div'); d.className='toast'; d.textContent=msg; s.appendChild(d); setTimeout(()=>d.remove(),3400)}
function save(){state.savedAt=new Date().toISOString(); try{localStorage.setItem(STORE_KEY,JSON.stringify(state))}catch(e){console.warn(e)} updateSaveInfo()}
function updateSaveInfo(){const e=$('#saveInfo'); if(e)e.textContent=state.savedAt?'Gespeichert · '+new Date(state.savedAt).toLocaleTimeString('de-DE'):'Vorschau nicht gespeichert'}
function initState(){
  const saved=localStorage.getItem(STORE_KEY); if(saved){try{const p=JSON.parse(saved); if(p.version===VERSION)return p}catch(e){}}
  const wb=clone(window.WORKBOOK_DATA); const originals={}; wb.sheets.forEach(s=>{s.id=sheetId(s.name,'system'); originals[s.name]=s; (s.themes||[]).forEach(t=>{t.key=topicKey(t); prepTheme(t,s)});});
  const adminSheets=[]; const usersheets={};
  for(const s of wb.sheets){ if(s.name==='Datenpool') continue; const copy=clone(s); copy.ownerId='system'; copy.system=true; copy.permissions={admin:'manage'}; copy.vehicle = s.name==='A3 8P'; copy.mileage=copy.vehicle?143430:0; adminSheets.push(copy); }
  const catalog=createCatalogSheet(); adminSheets.push(catalog);
  for(const u of USERS.filter(u=>u.id!=='admin')){
    usersheets[u.id]=[];
    const auto=emptyFrom(originals['A3 8P'],'Auto',u.id,true); usersheets[u.id].push(auto);
    usersheets[u.id].push(emptyFrom(originals['Teilebestand'],'Teilebestand',u.id,false));
    usersheets[u.id].push(emptyFrom(originals['Werkzeug'],'Werkzeug',u.id,false));
    usersheets[u.id].push(catalog);
  }
  const settings={dark:false,logo:DEFAULT_LOGO,logoSize:42,colors:{primary:'#1f4e79',primaryDark:'#153a5a',secondary:'#d9eaf7',accent:'#fff2cc',bg:'#f4f7fb',surface:'#ffffff',text:'#19202a',tableText:'#111827',headText:'#12355b',border:'#d7dee8'},adminDefaultColors:null,density:'compact'};
  return {version:VERSION,currentUserId:null,users:clone(USERS),adminSheets,userSheets:usersheets,activeSheetId:null,settings,userColorOverrides:{},dropdownLists:clone(window.DROPDOWN_LISTS),dropdownAssignments:defaultDropdownAssignments(),messages:[],notices:[{title:'Willkommen in der KFZ-APP Vorschau',text:'Diese Demo zeigt die spätere Supabase-/Netlify-App ohne echte Cloud-Anbindung.',active:true}],selectedInspectionUserId:null,releaseTab:'notes',catalogOpen:{},fileOpen:{},catalogSelected:'cat_file_1',savedAt:null};
}
function prepTheme(t,sheet){
  t.displayName=t.displayName||String(t.name||'Thema').replace(/ \+ Laufleistung$/,'');
  for(const col of (t.columns||[])){ col.parts=headerParts(col.label,t.displayName); }
}
function headerParts(label, themeName){
  let s=String(label||''); s=s.replace(/^[A-Z]+\s*[–-]\s*/,''); let parts=s.split('/').map(x=>x.trim()).filter(Boolean);
  // Remove repeated table/sheet names and Excel-only Laufleistung current-km group, but preserve next Tausch/Laufleistung as actual maintenance column.
  const removeFirst = new Set([themeName,'Verbauliste + Laufleistung','gesammt Teilebestand','Wartungsarbeiten/Verbrauchsmatierial']);
  while(parts.length && removeFirst.has(parts[0])) parts.shift();
  if(parts[0]==='Laufleistung' && parts.length>=2 && /^\d+[\d\.]*$/.test(parts[1])) parts=parts.slice(2);
  if(parts.length===0) parts=['Eintrag'];
  return parts;
}
function emptyFrom(original,name,ownerId,vehicle){
  const s=clone(original); s.name=name; s.id=sheetId(name,ownerId); s.ownerId=ownerId; s.system=false; s.vehicle=vehicle; s.mileage=vehicle?0:0; s.permissions={[ownerId]:'manage'}; s.files=defaultFileTree();
  const newCells={};
  for(const t of s.themes||[]){ prepTheme(t,s); const cols=t.columns||[]; const blankRows=3; const oldStart=t.dataStart||5; t.rowGroups=[]; t.endRow=oldStart+blankRows-1; t.dataEnd=t.endRow; for(let i=0;i<blankRows;i++){const r=oldStart+i; t.rowGroups.push([r]); for(const col of cols){const c=col.col; const tmpl=s.cells?.[`${oldStart}:${c}`]||{}; newCells[`${r}:${c}`]={r,c,a1:colLetter(c)+r,text:'',raw:'',sort:{kind:'text',value:''},style:tmpl.style||{}};}}
  }
  s.cells=newCells; return s;
}
function defaultFileTree(){return [{id:uid('folder'),type:'folder',name:'Bilder',open:false,children:[{id:uid('file'),type:'file',name:'Beispielbild.jpg',note:'Demo-Datei',thumb:''}]},{id:uid('folder'),type:'folder',name:'Rechnungen',open:false,children:[]},{id:uid('folder'),type:'folder',name:'Dokumente',open:false,children:[]}];}
function createCatalogSheet(){return {id:'global_parts_catalog',name:'Teilekatalog',ownerId:'system',system:true,global:true,catalog:true,permissions:{admin:'manage','role:Nutzer':'view','role:Gast':'view'},tree:[{id:'cat_engine',type:'folder',name:'Motor',mode:'icon',icon:'🛠️',open:true,children:[{id:'cat_filter',type:'folder',name:'Filter',mode:'icon',icon:'🧰',open:true,children:[{id:'cat_file_1',type:'file',name:'Innenraumfilter MAHLE LAO 181',part:'LAO 181',thumb:'',rowId:'catrow1'}]}]},{id:'cat_body',type:'folder',name:'Karosserie',mode:'icon',icon:'🚗',open:false,children:[{id:'cat_file_2',type:'file',name:'Spiegelblinker ABAKUS',part:'003-18-862',thumb:'',rowId:'catrow2'}]}],rows:[{id:'catrow1',image:'',name:'Innenraumfilter MAHLE LAO 181',part:'LAO 181',category:'Filter / Luftfilter',note:'Klick auf Datei springt zu dieser Zeile.'},{id:'catrow2',image:'',name:'Spiegelblinker ABAKUS',part:'003-18-862',category:'Beleuchtung / Blinker',note:'Beispiel aus Originaldaten.'}]};}
function defaultDropdownAssignments(){return [{match:'kategorie',list:'Kategorie / Unterkategorie'},{match:'unter kategorie',list:'Kategorie / Unterkategorie'},{match:'einheit',list:'Einheiten'},{match:'fahrzeug',list:'Auto'},{match:'art',list:'Artikelart'},{match:'hersteller',list:'Hersteller'}];}
function allSheetsForUser(){const u=currentUser(); if(!u) return []; const arr=u.id==='admin'?state.adminSheets:(state.userSheets[u.id]||[]); return arr.filter(s=>canAccess(accessFor(s,u),'view'));}
function accessFor(sheet,user=currentUser()){ if(!sheet||!user)return ''; if(sheet.ownerId===user.id)return 'manage'; if(sheet.permissions?.[user.id])return sheet.permissions[user.id]; if(sheet.permissions?.['role:'+user.role])return sheet.permissions['role:'+user.role]; return '';}
function sheetById(id){return allSheetsForUser().find(s=>s.id===id)}
function ensureActive(){const sheets=allSheetsForUser(); if(!sheets.length){state.activeSheetId=null;return null} if(!sheets.some(s=>s.id===state.activeSheetId)) state.activeSheetId=sheets[0].id; return sheetById(state.activeSheetId)}
function applyTheme(){const st=state.settings; document.body.classList.toggle('dark',!!st.dark); const c=st.colors||{}; const root=document.documentElement; const map={primary:'--primary',primaryDark:'--primaryDark',secondary:'--secondary',accent:'--accent',bg:'--bg',surface:'--surface',text:'--text',tableText:'--tableText',headText:'--headText',border:'--border'}; Object.entries(map).forEach(([k,v])=>{if(c[k])root.style.setProperty(v,c[k])}); root.style.setProperty('--logoSize',(st.logoSize||42)+'px'); root.style.setProperty('--density-cell-y',st.density==='comfort'?'9px':st.density==='normal'?'7px':'5px'); root.style.setProperty('--density-cell-x',st.density==='comfort'?'11px':st.density==='normal'?'9px':'7px'); $('#appLogo').src=st.logo||DEFAULT_LOGO; $('#favicon').href=st.logo||DEFAULT_LOGO;}
function render(){applyTheme(); const u=currentUser(); $('#loginView').classList.toggle('hidden',!!u); $('#appView').classList.toggle('hidden',!u); if(!u)return; $('#userChip').innerHTML=`<span>👤</span><span>${esc(u.name)}</span><span class="hint" style="color:#fff9">${esc(u.role)}</span>`; updateSaveInfo(); renderTabs(); renderMain(); renderDemoBar();}
function renderTabs(){const tabs=$('#tabs'); const sheets=allSheetsForUser(); tabs.innerHTML=sheets.map(s=>`<button class="tab ${s.id===state.activeSheetId?'active':''}" draggable="true" data-tab="${esc(s.id)}" style="--tabW:${s.tabWidth||130}px"><span>${s.vehicle?'🚗':s.catalog?'🧩':'📄'}</span><span>${esc(s.name)}</span><span class="type">${s.vehicle?'Auto':s.catalog?'global':'Reiter'}</span><span class="tab-resize" data-resize-tab="${esc(s.id)}"></span></button>`).join('')+`<button class="tab secondary" id="addTabBtn">+ Reiter</button>`; tabs.onclick=e=>{const b=e.target.closest('[data-tab]'); if(b){state.activeSheetId=b.dataset.tab; save(); render();} if(e.target.closest('#addTabBtn')) addSheetDialog();}; tabs.ondragstart=e=>{const b=e.target.closest('[data-tab]'); if(b){dragged={type:'tab',id:b.dataset.tab};}}; tabs.ondragover=e=>{if(dragged?.type==='tab'){e.preventDefault(); const b=e.target.closest('[data-tab]'); if(b)b.classList.add('drag-over')}}; tabs.ondragleave=e=>e.target.closest('[data-tab]')?.classList.remove('drag-over'); tabs.ondrop=e=>{const b=e.target.closest('[data-tab]'); $$('.tab').forEach(x=>x.classList.remove('drag-over')); if(b&&dragged?.type==='tab')moveTab(dragged.id,b.dataset.tab); dragged=null;}; bindTabResize();}
function renderMain(){const sheet=ensureActive(); const main=$('#main'); if(!sheet){main.innerHTML='<div class="login-panel"><h2>Keine Reiter verfügbar</h2></div>'; return;} if(sheet.catalog){renderCatalog(sheet); return;} const access=accessFor(sheet); main.innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>${esc(sheet.name)}</h2><div class="sheet-sub"><span class="pill ${sheet.vehicle?'ok':''}">${sheet.vehicle?'Autoreiter':'Neutraler Reiter'}</span><span class="pill">Zugriff: ${accessLabel(access)}</span>${sheet.ownerId!=='system'?`<span class="pill warn">Besitzer: ${esc(userName(sheet.ownerId))}</span>`:''}</div></div><div class="sheet-tools"><button class="secondary" id="shareTabBtn">Teilen/Rechte</button><button class="secondary" id="archiveTabBtn">Archivieren</button><button class="danger" id="deleteTabBtn">Löschen</button></div></div>${sheet.vehicle?vehicleHtml(sheet):''}<div class="global-search"><strong>🔎 Teilenummer / Partnummer</strong><input id="partSearch" placeholder="PN, OEM, SN, Teilenummer suchen …"><button id="partSearchBtn">Suchen</button></div><div id="partResults" class="search-results"></div>${sheet.vehicle?renderFileSection(sheet,access):''}<div class="topic-nav">${(sheet.themes||[]).map(t=>`<a href="#${esc(topicDomId(sheet,t))}">${esc(t.displayName||t.name)}</a>`).join('')}</div><div class="topic-stack">${(sheet.themes||[]).map(t=>renderTopic(sheet,t,access)).join('')}</div>`; bindSheetEvents(sheet,access);}
function accessLabel(a){return {view:'Ansehen',edit:'Bearbeiten',manage:'Verwalten'}[a]||'Kein Zugriff'} function userName(id){return state.users.find(u=>u.id===id)?.name||id}
function vehicleHtml(sheet){return `<div class="vehicle-panel"><label for="mileageInput">Laufleistung</label><input id="mileageInput" value="${esc(formatKm(sheet.mileage||0))}" inputmode="numeric"><span class="pill">Km</span></div>`}
function topicDomId(sheet,t){return 'topic_'+slug(sheet.id+'_'+(t.key||topicKey(t)))}
function renderTopic(sheet,t,access){const canEdit=canAccess(access,'edit'); return `<section class="topic-card" id="${topicDomId(sheet,t)}" data-topic="${esc(t.key||topicKey(t))}" style="width:${t.uiWidth||'auto'};height:${t.uiHeight||'auto'}"><div class="topic-head" draggable="${canEdit}" data-drag-topic="${esc(t.key||topicKey(t))}"><h3>${esc(t.displayName||t.name)}</h3><div class="topic-actions"><select class="sort-select" data-sort-select>${(t.columns||[]).map((c,i)=>`<option value="${i}">${esc(shortHeader(c.parts))}</option>`).join('')}</select><button class="small secondary" data-sort="asc">↑</button><button class="small secondary" data-sort="desc">↓</button><button class="small secondary" data-add-col>+ Spalte</button><button class="small secondary" data-collapse>Ein-/ausklappen</button></div></div><div class="topic-body"><div class="table-wrap"><table class="kfz-table">${renderHeader(sheet,t,canEdit)}<tbody>${renderRows(sheet,t,canEdit)}</tbody></table></div></div><div class="topic-foot"><button class="secondary" data-add-row-end>+ Zeile unten</button><span class="hint">Doppelklick bearbeitet · Links bleiben anklickbar · Zeilen/Spalten per Griff verschieben</span></div></section>`}
function shortHeader(parts){return (parts||[]).filter(Boolean).slice(-2).join(' / ')||'Spalte'}
function renderHeader(sheet,t,canEdit){const cols=t.columns||[]; const maxDepth=Math.max(1,...cols.map(c=>(c.parts||[]).length)); let html='<thead>'; for(let tier=0;tier<maxDepth;tier++){html+='<tr>'; if(tier===0)html+=`<th class="row-tools" rowspan="${maxDepth}">↕</th>`; let i=0; while(i<cols.length){const parts=cols[i].parts||['Spalte']; if(tier>=parts.length){i++; continue;} const label=parts[tier]||''; const prefix=parts.slice(0,tier).join('|'); let span=1; while(i+span<cols.length){const p2=cols[i+span].parts||[]; if(tier>=p2.length)break; if((p2[tier]||'')!==label)break; if(p2.slice(0,tier).join('|')!==prefix)break; span++;} const rowspan=(tier===parts.length-1)?(maxDepth-tier):1; html+=`<th draggable="${canEdit}" data-head-tier="${tier}" data-head-start="${i}" data-head-span="${span}" data-col-index="${i}" colspan="${span}" rowspan="${rowspan}" title="Doppelklick zum Bearbeiten">${esc(label)}</th>`; i+=span;} html+='</tr>';} return html+'</thead>';}
function renderRows(sheet,t,canEdit){let rows=(t.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g).filter(Boolean); if(!rows.length){for(let r=t.dataStart||5;r<=t.endRow;r++)rows.push(r)} let html=''; rows.forEach((r,idx)=>{html+=`<tr class="insert-row"><td colspan="${(t.columns||[]).length+1}"><button data-insert-before="${r}">+ Zeile hier einfügen</button></td></tr>`; html+=`<tr draggable="${canEdit}" data-row="${r}"><td class="row-tools"><button class="row-handle" ${canEdit?'':'disabled'} draggable="${canEdit}" data-row-handle="${r}">↕</button></td>`; for(const col of (t.columns||[])){html+=renderCell(sheet,t,r,col,canEdit)} html+='</tr>';}); html+=`<tr class="insert-row"><td colspan="${(t.columns||[]).length+1}"><button data-insert-before="end">+ Zeile unten einfügen</button></td></tr>`; return html;}
function renderCell(sheet,t,r,col,canEdit){const c=col.col; const cell=sheet.cells?.[`${r}:${c}`] || {r,c,text:'',raw:'',style:{}}; recalcCellIfNeeded(sheet,t,r,col,cell); const style=cellStyleVars(cell); const cls=statusClass(cell,col); const txt=cell.text??cell.raw??''; const hasNote=!!cell.comment; const hasFormula=!!cell.formula; const dd=dropdownForColumn(col); const bar=barWidth(cell,col); const title=[hasNote?'Notiz: '+cell.comment:'', hasFormula?'Formel: '+cell.formula:'', dd?'Dropdown: '+dd.list:''].filter(Boolean).join('\n'); let body=''; if(cell.image)body+=`<img class="cell-img" src="${cell.image}" alt="Bild">`; if(cell.hyperlink) body+=`<a class="cell-link" href="${esc(cell.hyperlink)}" target="_blank" rel="noopener" data-link>${esc(txt||'Link')}</a>`; else body+=`<span class="cell-text" data-cell-text>${esc(displayText(txt))}</span>`; return `<td class="${cls}" data-r="${r}" data-c="${c}" data-col-index="${(t.columns||[]).indexOf(col)}" style="${style}" title="${esc(title)}"><div class="cell-inner">${bar?`<span class="bar" style="width:${bar}%"></span>`:''}${body}${hasNote?'<span class="note-dot"></span>':''}${hasFormula?'<span class="formula-dot">ƒ</span>':''}${dd?'<span class="formula-dot">▾</span>':''}<span class="cell-icons">${canEdit?'<button class="cell-icon" data-edit-cell>✎</button>':''}</span></div></td>`;}
function displayText(v){if(v===''||v==null)return ''; return String(v)}
function cellStyleVars(cell){let s=''; const bg=cell.style?.backgroundColor; const color=cell.style?.color; if(bg && bg!=='#FFFFFF' && bg!=='#ffffff')s+=`--cellBg:${bg};`; if(color)s+=`color:${color};`; return s}
function statusClass(cell,col){const v=String(cell.text??cell.raw??'').trim(); const low=(shortHeader(col.parts)||'').toLowerCase(); const n=parseNum(v); if(low.includes('tage bis') && !isNaN(n)){if(n<0)return 'status-danger'; if(n<45)return 'status-warn'; return 'status-ok'} if(low.includes('anteil') && !isNaN(n)){if(n>=100)return 'status-danger'; if(n>=80)return 'status-warn'; return 'status-ok'} if(['verfügbar','zulauf','gekauft'].some(k=>low.includes(k)) && !isNaN(n)){return n<=0?'status-danger':'status-ok'} return ''}
function barWidth(cell,col){const low=shortHeader(col.parts).toLowerCase(); if(!low.includes('anteil'))return 0; let n=parseNum(cell.text??cell.raw); if(isNaN(n))return 0; if(String(cell.text).includes('%')){} return Math.max(0,Math.min(100,n));}
function recalcCellIfNeeded(sheet,t,r,col,cell){const h=shortHeader(col.parts).toLowerCase(); if(!(h.includes('tage bis')||h.includes('anteil bis')))return; const cols=t.columns||[]; const find=(s)=>cols.find(cc=>shortHeader(cc.parts).toLowerCase().includes(s)); const installDate=find('einbaudatum'); const installKm=find('kilometerstand'); const nextKm=cols.find(cc=>{let l=(cc.parts||[]).join(' / ').toLowerCase(); return l.includes('nächster tausch')&&l.includes('laufleistung')}); const nextDate=cols.find(cc=>{let l=(cc.parts||[]).join(' / ').toLowerCase(); return l.includes('nächster tausch')&&l.includes('datum')}); const get=cc=>cc?sheet.cells?.[`${r}:${cc.col}`]?.text:''; if(h.includes('tage bis')){let days='-'; const d=parseDate(get(nextDate)); if(d) days=String(daysBetween(TODAY,d)); else {const base=parseNum(get(installKm)), interval=parseNum(get(nextKm)); if(!isNaN(base)&&!isNaN(interval)&&sheet.mileage){days=String(Math.round((base+interval-sheet.mileage)/50));}} cell.text=days; cell.raw=days;} if(h.includes('anteil bis')){let pct='-'; const d0=parseDate(get(installDate)), d1=parseDate(get(nextDate)); if(d0&&d1&&d1>d0){pct=(Math.max(0,Math.min(100,daysBetween(d0,TODAY)/daysBetween(d0,d1)*100))).toFixed(1).replace('.',',')+'%';} else {const base=parseNum(get(installKm)), interval=parseNum(get(nextKm)); if(!isNaN(base)&&!isNaN(interval)&&interval>0&&sheet.mileage){pct=(Math.max(0,Math.min(100,(sheet.mileage-base)/interval*100))).toFixed(1).replace('.',',')+'%';}} cell.text=pct; cell.raw=pct;}}
function dropdownForColumn(col){const label=(col.parts||[]).join(' ').toLowerCase(); const rule=state.dropdownAssignments.find(a=>label.includes(String(a.match).toLowerCase())); return rule?{list:rule.list,values:state.dropdownLists[rule.list]||[]}:null}
function bindSheetEvents(sheet,access){const canEdit=canAccess(access,'edit'); $('#shareTabBtn')?.addEventListener('click',()=>openSettings('rechte')); $('#archiveTabBtn')?.addEventListener('click',()=>archiveSheet(sheet)); $('#deleteTabBtn')?.addEventListener('click',()=>deleteSheet(sheet)); $('#mileageInput')?.addEventListener('blur',e=>{sheet.mileage=rawKm(e.target.value); e.target.value=formatKm(sheet.mileage); autosaveRecalc(sheet)}); $('#mileageInput')?.addEventListener('keydown',e=>{if(e.key==='Enter')e.currentTarget.blur()}); $('#partSearchBtn')?.addEventListener('click',()=>partSearch(sheet)); $('#partSearch')?.addEventListener('keydown',e=>{if(e.key==='Enter')partSearch(sheet)}); if(sheet.vehicle&&$('#fileTree')){bindTree($('#fileTree'),sheet.files,'file',access); $('[data-file-add-root]')?.addEventListener('click',()=>{sheet.files.push({id:uid('folder'),type:'folder',name:'Neuer Ordner',open:true,children:[],icon:'📁'});save();render();}); $('[data-file-upload-root]')?.addEventListener('click',()=>uploadVehicleFile(sheet));} $$('.topic-card').forEach(card=>bindTopic(card,sheet,access));}
function bindTopic(card,sheet,access){const canEdit=canAccess(access,'edit'); const key=card.dataset.topic; const t=(sheet.themes||[]).find(x=>(x.key||topicKey(x))===key); if(!t)return; card.addEventListener('click',e=>{if(e.target.matches('[data-collapse]'))card.classList.toggle('collapsed'); if(e.target.matches('[data-sort]'))sortTheme(sheet,t,e.target.dataset.sort); if(e.target.matches('[data-add-row-end]'))addRow(sheet,t,'end'); if(e.target.matches('[data-insert-before]'))addRow(sheet,t,e.target.dataset.insertBefore); if(e.target.matches('[data-add-col]'))addColumn(sheet,t); if(e.target.matches('[data-edit-cell]'))openCellEditor(sheet,t,e.target.closest('td'));}); card.addEventListener('dblclick',e=>{const th=e.target.closest('th[data-head-tier]'); if(th&&canEdit)return editHeader(t,th); const td=e.target.closest('td[data-r]'); if(td&&canEdit&&!e.target.closest('a'))return editCellInline(sheet,t,td);}); card.addEventListener('dragstart',e=>{const row=e.target.closest('[data-row]'); const th=e.target.closest('th[data-col-index]'); const head=e.target.closest('[data-drag-topic]'); if(row&&canEdit){dragged={type:'row',r:+row.dataset.row,topic:t.key,sheet:sheet.id};} else if(th&&canEdit){dragged={type:'col',idx:+th.dataset.colIndex,topic:t.key,sheet:sheet.id};} else if(head&&canEdit){dragged={type:'topic',key:t.key,sheet:sheet.id};}}); card.addEventListener('dragover',e=>{if(dragged){e.preventDefault();}}); card.addEventListener('drop',e=>{if(!dragged)return; if(dragged.type==='row'){const tr=e.target.closest('tr[data-row]'); if(tr)moveRow(sheet,t,dragged.r,+tr.dataset.row);} if(dragged.type==='col'){const th=e.target.closest('th[data-col-index]'); if(th)moveCol(t,dragged.idx,+th.dataset.colIndex);} dragged=null; render(); save();});}
function autosaveRecalc(sheet){for(const t of sheet.themes||[])for(const r of (t.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g))for(const col of t.columns||[]){const cell=sheet.cells?.[`${r}:${col.col}`]; if(cell)recalcCellIfNeeded(sheet,t,r,col,cell)} save(); render();}
function editHeader(t,th){const tier=+th.dataset.headTier, start=+th.dataset.headStart, span=+th.dataset.headSpan; const old=th.textContent; th.contentEditable='true'; th.focus(); document.execCommand?.('selectAll',false,null); const finish=()=>{const val=th.textContent.trim()||old; for(let i=start;i<start+span;i++){t.columns[i].parts[tier]=val; t.columns[i].label=t.columns[i].parts.join(' / ');} th.contentEditable='false'; save(); render();}; th.onblur=finish; th.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();th.blur()} if(e.key==='Escape'){th.textContent=old; th.blur()}};}
function editCellInline(sheet,t,td){const r=+td.dataset.r,c=+td.dataset.c; const col=t.columns.find(cc=>cc.col===c); const cell=sheet.cells[`${r}:${c}`]||(sheet.cells[`${r}:${c}`]={r,c,a1:colLetter(c)+r,text:'',raw:'',style:{}}); const dd=dropdownForColumn(col); const inner=td.querySelector('.cell-inner'); if(dd&&dd.values.length){const sel=document.createElement('select'); sel.innerHTML=[''].concat(dd.values).map(v=>`<option ${v===cell.text?'selected':''}>${esc(v)}</option>`).join(''); inner.innerHTML=''; inner.appendChild(sel); sel.focus(); sel.onchange=sel.onblur=()=>{cell.text=sel.value; cell.raw=sel.value; save(); render();}; return;} const span=td.querySelector('[data-cell-text]')||td.querySelector('a'); const old=cell.text??''; if(!span)return; span.classList.add('editing'); span.contentEditable='true'; span.focus(); document.execCommand?.('selectAll',false,null); const finish=()=>{cell.text=span.textContent.trim(); cell.raw=cell.text; span.contentEditable='false'; span.classList.remove('editing'); autosaveRecalc(sheet)}; span.onblur=finish; span.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault(); span.blur()} if(e.key==='Escape'){cell.text=old; span.textContent=old; span.blur()}};}
function openCellEditor(sheet,t,td){const r=+td.dataset.r,c=+td.dataset.c; const cell=sheet.cells[`${r}:${c}`]||(sheet.cells[`${r}:${c}`]={r,c,a1:colLetter(c)+r,text:'',raw:'',style:{}}); showModal('Zelle bearbeiten',`<div class="grid2"><div class="field"><label>Wert</label><input id="ceValue" value="${esc(cell.text||'')}"></div><div class="field"><label>Link</label><input id="ceLink" value="${esc(cell.hyperlink||'')}"></div><div class="field full"><label>Notiz / Hovertext</label><textarea id="ceNote">${esc(cell.comment||'')}</textarea></div><div class="field"><label>Hintergrundfarbe</label><input id="ceBg" type="color" value="${esc(cell.style?.backgroundColor||'#ffffff')}"></div><div class="field"><label>Textfarbe</label><input id="ceColor" type="color" value="${esc(cell.style?.color||'#111827')}"></div><div class="field full"><label>Bild hochladen</label><input id="ceImg" type="file" accept="image/*"></div></div>`,[{text:'Speichern',action:()=>{cell.text=$('#ceValue').value; cell.raw=cell.text; cell.hyperlink=$('#ceLink').value.trim(); cell.comment=$('#ceNote').value.trim(); cell.style=cell.style||{}; cell.style.backgroundColor=$('#ceBg').value; cell.style.color=$('#ceColor').value; const f=$('#ceImg').files[0]; if(f){const rd=new FileReader(); rd.onload=()=>{cell.image=rd.result; save(); closeModal(); render();}; rd.readAsDataURL(f);} else {save(); closeModal(); render();}}},{text:'Abbrechen',secondary:true,action:closeModal}]);}
function addRow(sheet,t,before){const cols=t.columns||[]; const existing=(t.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g); const newR=Math.max(sheet.maxRow||0,...existing)+1; sheet.maxRow=Math.max(sheet.maxRow||0,newR); t.rowGroups=t.rowGroups||[]; const pos=before==='end'||!before?t.rowGroups.length:t.rowGroups.findIndex(g=>(Array.isArray(g)?g[0]:g)===+before); const insertPos=pos<0?t.rowGroups.length:pos; t.rowGroups.splice(insertPos,0,[newR]); t.endRow=Math.max(t.endRow||0,newR); t.dataEnd=t.endRow; cols.forEach(col=>{const tmpl=sheet.cells?.[`${existing[0]}:${col.col}`]||{}; sheet.cells[`${newR}:${col.col}`]={r:newR,c:col.col,a1:colLetter(col.col)+newR,text:'',raw:'',sort:{kind:'text',value:''},style:clone(tmpl.style||{})};}); save(); render();}
function addColumn(sheet,t){const next=Math.max(...t.columns.map(c=>c.col))+1; const col={col:next,letter:colLetter(next),label:'Neue Spalte',parts:['Neue Spalte']}; t.columns.push(col); sheet.maxCol=Math.max(sheet.maxCol||0,next); for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; sheet.cells[`${r}:${next}`]={r,c:next,a1:colLetter(next)+r,text:'',raw:'',style:{}};} save(); render();}
function moveRow(sheet,t,from,to){const arr=t.rowGroups||[]; const a=arr.findIndex(g=>(Array.isArray(g)?g[0]:g)===from), b=arr.findIndex(g=>(Array.isArray(g)?g[0]:g)===to); if(a<0||b<0)return; const [x]=arr.splice(a,1); arr.splice(b,0,x);}
function moveCol(t,from,to){if(from===to)return; const [x]=t.columns.splice(from,1); t.columns.splice(to,0,x);}
function sortTheme(sheet,t,dir){const idx=+document.querySelector(`#${CSS.escape(topicDomId(sheet,t))} [data-sort-select]`)?.value||0; const col=t.columns[idx]; const arr=t.rowGroups||[]; arr.sort((ga,gb)=>{const a=sheet.cells?.[`${Array.isArray(ga)?ga[0]:ga}:${col.col}`]?.text||''; const b=sheet.cells?.[`${Array.isArray(gb)?gb[0]:gb}:${col.col}`]?.text||''; const na=parseNum(a), nb=parseNum(b); let cmp=!isNaN(na)&&!isNaN(nb)?na-nb:String(a).localeCompare(String(b),'de',{numeric:true}); return dir==='desc'?-cmp:cmp;}); save(); render();}
function partSearch(sheet){const q=$('#partSearch').value.trim().toLowerCase(); const out=$('#partResults'); if(!q){out.innerHTML='';return} const hits=[]; for(const t of sheet.themes||[])for(const col of t.columns||[]){const h=(col.parts||[]).join(' ').toLowerCase(); if(/teil|part|pn|oem|sn|nummer/.test(h)){for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; const v=String(sheet.cells?.[`${r}:${col.col}`]?.text||''); if(v.toLowerCase().includes(q))hits.push({theme:t,row:r,col:col,value:v});}}} out.innerHTML=hits.length?hits.slice(0,50).map(h=>`<div class="search-hit"><div><strong>${esc(h.value)}</strong><div class="hint">${esc(h.theme.displayName||h.theme.name)} · ${esc(shortHeader(h.col.parts))}</div></div><button class="small secondary" onclick="document.getElementById('${topicDomId(sheet,h.theme)}').scrollIntoView({behavior:'smooth'});">Springen</button></div>`).join(''):'<div class="search-hit">Keine Treffer.</div>';}
function renderFileSection(sheet,access){if(!sheet.files)sheet.files=defaultFileTree(); return `<section class="file-section" id="vehicleFiles"><div class="section-head"><h3>Dateien & Bilder</h3><div>${canAccess(access,'edit')?'<button class="small secondary" data-file-add-root>+ Ordner</button> <button class="small secondary" data-file-upload-root>+ Datei/Bild</button>':''}</div></div><div class="path-line" id="filePath">Pfad: /</div><div class="tree" id="fileTree">${renderTree(sheet.files,'file',access)}</div></section>`}
function renderTree(nodes,kind,access,depth=0,path=[]){return (nodes||[]).map(n=>{const canEdit=kind==='catalog'?currentUser()?.role==='Admin':canAccess(access,'edit'); const icon=n.thumb?`<img class="thumb" src="${n.thumb}">`:`<span class="thumb">${n.type==='folder'?(n.icon||'📁'):'📄'}</span>`; return `<div class="tree-node" data-node="${n.id}" data-kind="${kind}"><div class="tree-row" style="margin-left:${depth*4}px" data-tree-click><span>${n.type==='folder'?(n.open?'▾':'▸'):'•'}</span>${icon}<strong>${esc(n.name)}</strong>${n.part?`<span class="pill">${esc(n.part)}</span>`:''}<span style="flex:1"></span>${canEdit?`<button class="small secondary" data-tree-edit>✎</button>${n.type==='folder'?'<button class="small secondary" data-tree-add-folder>+ Ordner</button><button class="small secondary" data-tree-add-file>+ Datei</button>':''}`:''}</div>${n.type==='folder'&&n.open?`<div class="tree-children">${renderTree(n.children||[],kind,access,depth+1,path.concat(n.name))}</div>`:''}</div>`}).join('')}
function bindTree(container,root,kind,access){container.onclick=e=>{const nodeEl=e.target.closest('[data-node]'); if(!nodeEl)return; const node=findNode(root,nodeEl.dataset.node); if(e.target.closest('[data-tree-edit]'))return editTreeNode(node,kind); if(e.target.closest('[data-tree-add-folder]'))return addTreeChild(node,'folder',kind); if(e.target.closest('[data-tree-add-file]'))return addTreeChild(node,'file',kind); if(e.target.closest('[data-tree-click]')){if(node.type==='folder'){node.open=!node.open; save(); render();} else {if(kind==='catalog'){state.catalogSelected=node.id; save(); render(); setTimeout(()=>document.querySelector(`[data-catalog-row="${node.rowId}"]`)?.scrollIntoView({behavior:'smooth',block:'center'}),50)}}}}}
function findNode(nodes,id){for(const n of nodes||[]){if(n.id===id)return n; const c=findNode(n.children,id); if(c)return c}return null}
function addTreeChild(parent,type,kind){parent.children=parent.children||[]; parent.open=true; const node={id:uid(type),type,name:type==='folder'?'Neuer Ordner':'Neue Datei',open:false,children:type==='folder'?[]:undefined,thumb:'',part:'',rowId:type==='file'&&kind==='catalog'?uid('catrow'):undefined}; if(kind==='catalog'&&type==='file'){const catalog=state.adminSheets.find(s=>s.catalog); catalog.rows.push({id:node.rowId,image:'',name:node.name,part:'',category:'',note:''});} parent.children.push(node); save(); render();}
function uploadVehicleFile(sheet){const input=document.createElement('input'); input.type='file'; input.accept='image/*,.pdf,.txt,.doc,.docx,.xls,.xlsx'; input.onchange=()=>{const f=input.files[0]; if(!f)return; const file={id:uid('file'),type:'file',name:f.name,note:'hochgeladen in der Demo',thumb:''}; if(f.type.startsWith('image/')){const rd=new FileReader(); rd.onload=()=>{file.thumb=rd.result; sheet.files.push(file); save(); render();}; rd.readAsDataURL(f);} else {sheet.files.push(file); save(); render();}}; input.click();}

function editTreeNode(node,kind){showModal((node.type==='folder'?'Ordner':'Datei')+' bearbeiten',`<div class="grid2"><div class="field"><label>Name</label><input id="tnName" value="${esc(node.name)}"></div><div class="field"><label>Teilenummer / Bildtext</label><input id="tnPart" value="${esc(node.part||'')}"></div><div class="field"><label>Icon/Text</label><input id="tnIcon" value="${esc(node.icon||'')}"></div><div class="field"><label>Vorschau/Bild</label><input id="tnFile" type="file" accept="image/*"></div></div>`,[{text:'Speichern',action:()=>{node.name=$('#tnName').value; node.part=$('#tnPart').value; node.icon=$('#tnIcon').value; const f=$('#tnFile').files[0]; if(f){const rd=new FileReader(); rd.onload=()=>{node.thumb=rd.result; save(); closeModal(); render();}; rd.readAsDataURL(f);} else {save(); closeModal(); render();}}},{text:'Abbrechen',secondary:true,action:closeModal}]);}
function renderCatalog(sheet){const isAdmin=currentUser()?.role==='Admin'; const selectedNode=findNode(sheet.tree,state.catalogSelected); $('#main').innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>Teilekatalog</h2><div class="sheet-sub"><span class="pill ok">Global für alle Nutzer</span><span class="pill">${isAdmin?'Admin bearbeitet':'Nur ansehen'}</span></div></div><div class="sheet-tools">${isAdmin?'<button id="catAddRoot" class="secondary">+ Hauptordner</button>':''}</div></div><section class="catalog-section"><div class="section-head"><h3>Ordnerstruktur</h3><span class="hint">Ordner öffnen direkt untereinander. Datei anklicken springt zur Tabelle.</span></div><div class="catalog-grid"><div><div class="tree" id="catalogTree">${renderTree(sheet.tree,'catalog','manage')}</div></div><div><h3>Dateitabelle</h3><table class="catalog-table"><thead><tr><th>Bild</th><th>Bildtext / Datei</th><th>Teilenummer</th><th>Kategorie</th><th>Notiz</th></tr></thead><tbody>${sheet.rows.map(r=>`<tr data-catalog-row="${r.id}" class="${selectedNode?.rowId===r.id?'highlight-row':''}"><td>${r.image?`<img class="catalog-preview" src="${r.image}">`:'<span class="thumb">🧩</span>'}</td><td><span class="jump-link" data-jump-row="${r.id}">${esc(r.name)}</span></td><td>${esc(r.part)}</td><td>${esc(r.category)}</td><td>${esc(r.note)}</td></tr>`).join('')}</tbody></table></div></div></section>`; $('#catAddRoot')?.addEventListener('click',()=>{sheet.tree.push({id:uid('folder'),type:'folder',name:'Neuer Hauptordner',open:true,children:[],icon:'📁'});save();render()}); $$('#main [data-jump-row]').forEach(el=>el.onclick=()=>document.querySelector(`[data-catalog-row="${el.dataset.jumpRow}"]`)?.scrollIntoView({behavior:'smooth',block:'center'})); bindTree($('#catalogTree'),sheet.tree,'catalog','manage');}
function bindFileTrees(){const sheet=sheetById(state.activeSheetId); if(sheet?.vehicle&&$('#fileTree'))bindTree($('#fileTree'),sheet.files,'file',accessFor(sheet));}
function showModal(title,body,actions=[]){document.body.classList.add('modal-open'); $('#modalRoot').innerHTML=`<div class="modal"><div class="modal-card"><div class="modal-head"><h2>${esc(title)}</h2><button class="modal-close secondary" onclick="closeModal()">✕</button></div><div class="modal-body">${body}</div><div class="modal-actions">${actions.map((a,i)=>`<button class="${a.secondary?'secondary':''}" data-modal-action="${i}">${esc(a.text)}</button>`).join('')}</div></div></div>`; actions.forEach((a,i)=>$(`[data-modal-action="${i}"]`)?.addEventListener('click',a.action));}
function closeModal(){document.body.classList.remove('modal-open'); $('#modalRoot').innerHTML='';}
function openSettings(panel='profil'){const u=currentUser(); const nav=[['profil','Profil'],['design','Design & Logo'],['farben','Farben'],['dropdowns','Dropdowns'],['rechte','Reiter & Rechte'],['nutzer','Nutzerverwaltung'],['abo','Abo & Zahlung'],['hinweise','Hinweise'],['daten','Sicherung/Import']]; showModal('Appeinstellungen',`<div class="settings-layout"><nav class="settings-nav">${nav.map(([id,txt])=>`<button class="${id===panel?'active':''}" data-set-panel="${id}">${txt}</button>`).join('')}</nav><section class="settings-panel" id="settingsPanel"></section></div>`,[{text:'Schließen',secondary:true,action:closeModal}]); $('#modalRoot').addEventListener('click',e=>{const b=e.target.closest('[data-set-panel]'); if(b){openSettings(b.dataset.setPanel)}}); renderSettingsPanel(panel);}
function renderSettingsPanel(panel){const el=$('#settingsPanel'); if(!el)return; const isAdmin=currentUser()?.role==='Admin'; if(panel==='profil')el.innerHTML=`<h3>Profil</h3><div class="grid2"><div class="field"><label>Name</label><input value="${esc(currentUser().name)}"></div><div class="field"><label>Rolle</label><input value="${esc(currentUser().role)}" disabled></div><div class="field"><label>Abo</label><input value="${esc(currentUser().abo)}" disabled></div><div class="field"><label>Cloud</label><input value="Supabase-Vorschau / lokal simuliert" disabled></div></div>`;
 if(panel==='design')el.innerHTML=`<h3>Design & Logo</h3><div class="grid2"><div><div class="drop-zone" id="logoDrop"><img src="${state.settings.logo}" style="width:${state.settings.logoSize}px;height:${state.settings.logoSize}px;object-fit:contain"><p>Logo hier hineinziehen oder Datei auswählen</p><input id="logoFile" type="file" accept="image/*"></div><button class="secondary" id="logoReset">Standardlogo</button></div><div><div class="field"><label>Logo-Größe</label><input id="logoSize" type="range" min="28" max="96" value="${state.settings.logoSize}"></div><div class="field"><label>Darkmode</label><select id="darkSelect"><option value="0">Hell</option><option value="1" ${state.settings.dark?'selected':''}>Dunkel</option></select></div><div class="field"><label>Tabellendichte</label><select id="densitySelect"><option value="compact">Kompakt</option><option value="normal" ${state.settings.density==='normal'?'selected':''}>Normal</option><option value="comfort" ${state.settings.density==='comfort'?'selected':''}>Komfort</option></select></div></div></div>`;
 if(panel==='farben')el.innerHTML=`<h3>Farben</h3><p class="hint">Admin kann Standardfarben für neue Nutzer festlegen. Nutzeränderungen gelten nur für diesen Nutzer.</p><div class="color-grid">${Object.entries({primary:'Primärfarbe',primaryDark:'Dunkle Primärfarbe',secondary:'Sekundär / Tabellenkopf',accent:'Akzent / Bearbeitung',bg:'Hintergrund',surface:'Kartenfläche',text:'App-Textfarbe',tableText:'Tabellen-Textfarbe',headText:'Kopfzellen-Textfarbe',border:'Rahmen'}).map(([k,l])=>`<div class="color-item"><label>${l}</label><input type="color" data-color="${k}" value="${state.settings.colors[k]||'#ffffff'}"></div>`).join('')}</div><div style="margin-top:12px">${isAdmin?'<button id="saveDefaultColors">Als Admin-Standard für neue Nutzer speichern</button>':''}<button class="secondary" id="resetColors">Farben zurücksetzen</button></div>`;
 if(panel==='dropdowns')el.innerHTML=`<h3>Dropdowns</h3><div class="grid2"><div><h4>Listen</h4>${Object.entries(state.dropdownLists).map(([name,vals])=>`<div class="field"><label>${esc(name)}</label><textarea data-dd-list="${esc(name)}">${esc(vals.join('\n'))}</textarea></div>`).join('')}<button id="addDdList" class="secondary">+ Liste</button></div><div><h4>Spalten-Zuordnung</h4><table class="permission-table"><tr><th>Wenn Spaltenname enthält</th><th>Liste</th></tr>${state.dropdownAssignments.map((a,i)=>`<tr><td><input data-dd-match="${i}" value="${esc(a.match)}"></td><td><select data-dd-assignment="${i}">${Object.keys(state.dropdownLists).map(n=>`<option ${a.list===n?'selected':''}>${esc(n)}</option>`).join('')}</select></td></tr>`).join('')}</table><button id="addDdAssign" class="secondary">+ Zuordnung</button></div></div>`;
 if(panel==='rechte')el.innerHTML=renderRightsPanel();
 if(panel==='nutzer')el.innerHTML=`<h3>Nutzerverwaltung</h3><table class="permission-table"><tr><th>Name</th><th>Rolle</th><th>Abo</th><th>Status</th></tr>${state.users.map(u=>`<tr><td>${esc(u.name)}</td><td>${esc(u.role)}</td><td>${esc(u.abo)}</td><td>aktiv</td></tr>`).join('')}</table>`;
 if(panel==='abo')el.innerHTML=`<h3>Abo & Zahlung</h3><div class="plan-grid"><div class="plan"><h4>Testphase</h4><div class="price">0 €</div><p>zeitlich begrenzter Zugriff</p></div><div class="plan"><h4>Monatlich</h4><div class="price">4,99 €</div><p>PayPal, Karte/Stripe oder manuelle Freischaltung</p></div><div class="plan"><h4>Admin</h4><div class="price">Verwaltung</div><p>Nutzer, Rechte, Standardfarben, Katalog</p></div></div><h4>Zahlarten</h4><button>PayPal</button> <button>Stripe / Karte</button> <button class="secondary">Manuelle Zahlung</button>`;
 if(panel==='hinweise')el.innerHTML=`<h3>Hinweise & Nachrichten</h3><div class="field"><label>Nachricht an Admin</label><textarea id="msgText"></textarea><button id="sendMsg">Senden</button></div><h4>Admin-Hinweise</h4>${state.notices.map(n=>`<div class="release-card"><strong>${esc(n.title)}</strong><p>${esc(n.text)}</p></div>`).join('')}`;
 if(panel==='daten')el.innerHTML=`<h3>Sicherung / Import</h3><button id="exportJson">JSON exportieren</button> <button class="secondary" id="importJson">JSON importieren</button><p class="hint">In der finalen Supabase-Version werden Daten, Rechte, Dateien und Logos in Supabase gespeichert.</p>`;
 bindSettings(panel);}
function bindSettings(panel){ if(panel==='design'){const drop=$('#logoDrop'), file=$('#logoFile'); const handle=f=>{if(!f)return; const rd=new FileReader(); rd.onload=()=>{state.settings.logo=rd.result; save(); openSettings('design'); render();}; rd.readAsDataURL(f)}; file.onchange=e=>handle(e.target.files[0]); drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag')}; drop.ondragleave=()=>drop.classList.remove('drag'); drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');handle(e.dataTransfer.files[0])}; $('#logoReset').onclick=()=>{state.settings.logo=DEFAULT_LOGO;save();openSettings('design');render()}; $('#logoSize').oninput=e=>{state.settings.logoSize=+e.target.value;save();render();}; $('#darkSelect').onchange=e=>{state.settings.dark=e.target.value==='1';save();render()}; $('#densitySelect').onchange=e=>{state.settings.density=e.target.value;save();render()}; }
 if(panel==='farben'){$$('[data-color]').forEach(i=>i.oninput=e=>{state.settings.colors[e.target.dataset.color]=e.target.value;save();render()}); $('#saveDefaultColors')?.addEventListener('click',()=>{state.settings.adminDefaultColors=clone(state.settings.colors);save();toast('Admin-Standardfarben gespeichert')}); $('#resetColors')?.addEventListener('click',()=>{localStorage.removeItem(STORE_KEY);location.reload()});}
 if(panel==='dropdowns'){$$('[data-dd-list]').forEach(t=>t.onblur=e=>{state.dropdownLists[e.target.dataset.ddList]=e.target.value.split(/\n+/).map(x=>x.trim()).filter(Boolean);save();openSettings('dropdowns')}); $$('[data-dd-match]').forEach(i=>i.onblur=e=>{state.dropdownAssignments[+e.target.dataset.ddMatch].match=e.target.value;save()}); $$('[data-dd-assignment]').forEach(s=>s.onchange=e=>{state.dropdownAssignments[+e.target.dataset.ddAssignment].list=e.target.value;save()}); $('#addDdList').onclick=()=>{const n=prompt('Name der neuen Dropdownliste?'); if(n){state.dropdownLists[n]=[];save();openSettings('dropdowns')}}; $('#addDdAssign').onclick=()=>{state.dropdownAssignments.push({match:'',list:Object.keys(state.dropdownLists)[0]});save();openSettings('dropdowns')}}
 if(panel==='rechte'){$$('[data-perm]').forEach(s=>s.onchange=e=>{const [sheetId,uid]=e.target.dataset.perm.split('|'); const sh=[...state.adminSheets,...Object.values(state.userSheets).flat()].find(x=>x.id===sheetId); if(!sh.permissions)sh.permissions={}; if(e.target.value)sh.permissions[uid]=e.target.value; else delete sh.permissions[uid]; save(); openSettings('rechte'); render();});}
 $('#sendMsg')?.addEventListener('click',()=>{state.messages.push({from:currentUser().id,text:$('#msgText').value,date:new Date().toISOString()});save();toast('Nachricht gespeichert')}); $('#exportJson')?.addEventListener('click',()=>download('kfz_app_demo_state_'+VERSION+'.json',JSON.stringify(state,null,2),'application/json'))}
function renderRightsPanel(){const sheets=allSheetsForUser().filter(s=>!s.catalog); return `<h3>Reiter & Rechte</h3><p class="hint">Private Nutzerreiter sind für Admins unsichtbar, bis der Nutzer sie explizit freigibt.</p><table class="permission-table"><tr><th>Reiter</th>${state.users.map(u=>`<th>${esc(u.name)}</th>`).join('')}</tr>${sheets.map(s=>`<tr><td><strong>${esc(s.name)}</strong><br><span class="hint">${s.vehicle?'Autoreiter':'Neutral'}</span></td>${state.users.map(u=>`<td><select data-perm="${s.id}|${u.id}"><option value="">kein Zugriff</option><option value="view" ${s.permissions?.[u.id]==='view'?'selected':''}>ansehen</option><option value="edit" ${s.permissions?.[u.id]==='edit'?'selected':''}>bearbeiten</option><option value="manage" ${s.permissions?.[u.id]==='manage'?'selected':''}>verwalten</option></select></td>`).join('')}</tr>`).join('')}</table>`}
function openRelease(){showModal('Release Notes & Nutzeranleitung',`<div class="release-search"><input id="releaseSearch" placeholder="Suchen, z. B. Logo, Dropdown, Teilekatalog, Rechte …"><button id="releaseSearchBtn">Suchen</button></div><div class="release-tabs"><button data-rel="notes" class="active">Release Notes</button><button data-rel="guide">Nutzeranleitung</button><button data-rel="admin">Admin-Anleitung</button></div><div id="releaseContent"></div>`,[{text:'Schließen',secondary:true,action:closeModal}]); renderReleaseContent('notes'); $$('.release-tabs button').forEach(b=>b.onclick=()=>{ $$('.release-tabs button').forEach(x=>x.classList.remove('active')); b.classList.add('active'); renderReleaseContent(b.dataset.rel,$('#releaseSearch').value)}); $('#releaseSearchBtn').onclick=()=>renderReleaseContent($('.release-tabs button.active').dataset.rel,$('#releaseSearch').value); $('#releaseSearch').onkeydown=e=>{if(e.key==='Enter')$('#releaseSearchBtn').click()};}
function renderReleaseContent(tab,q=''){const list={notes:releaseNotes(),guide:guideSections(),admin:adminGuide() }[tab]; const needle=q.trim().toLowerCase(); const filtered=needle?list.filter(x=>(x.title+' '+x.body+' '+(x.jump||'')).toLowerCase().includes(needle)):list; $('#releaseContent').innerHTML=`<div class="release-list">${filtered.map(x=>`<div class="${tab==='notes'?'release-card':'guide-section'}"><h4>${esc(x.title)}</h4><p>${esc(x.body)}</p>${x.jump?`<button class="small secondary" onclick="jumpTo('${esc(x.jump)}')">Zur Stelle springen</button>`:''}</div>`).join('')||'<div class="release-card">Keine Treffer.</div>'}</div>`}
function releaseNotes(){return [
 {title:'V26.07.11 · Konsolidierung',body:'Design wieder an LH_Boden orientiert, Excel-Datenbasis bewahrt, Datenpool ausgeblendet, Dropdownverwaltung in Einstellungen, globaler Teilekatalog als Reiter, Datei-/Bilderbaum direkt untereinander, Logo inklusive Größe, Textfarben, Kilometerfeld und dynamische Formatierungen zusammengeführt.',jump:'settings:design'},
 {title:'V26.07.10 · Design-Reset',body:'LH-Topbar, Einstellungsmodal, Demo-Steuerung und Release-Modal wurden wieder stärker an die Beispieldatei angelehnt.',jump:'settings:profil'},
 {title:'V26.07.09 · Teilekatalog-Korrektur',body:'Ordner sollten direkt untereinander öffnen und der Teilekatalog als globaler Reiter erscheinen.',jump:'catalog'},
 {title:'V26.07.08 · Logo und Tabellenkopf',body:'Logo/Favicon, Farbtexte und Release-Format wurden ergänzt; spätere Fehler wurden in V26.07.11 zurückgenommen.',jump:'settings:design'},
 {title:'V26.07.07 · Versionsschema',body:'Versionsschema Jahr.Monat.fortlaufende Version eingeführt und Laufleistung als Fahrzeugfeld vorbereitet.',jump:'vehicle'},
 {title:'V26.07.06 · Katalog/Dateien',body:'Fahrzeug-Dateien und Teilekatalog wurden als gewünschte neue Bereiche spezifiziert.',jump:'files'},
 {title:'V26.07.05 · Kopfzeilen',body:'Mehrstufige Excel-Kopfzeilen wie Menge/Verfügbar und nächster Tausch wurden als App-Kopfstruktur übernommen.',jump:'table'},
 {title:'V26.07.04 · Dropdowns',body:'Datenpool wird nicht mehr als Reiter genutzt, sondern als Dropdown-Grundlage in Einstellungen verwaltet.',jump:'settings:dropdowns'},
 {title:'V26.07.03 · Rechte',body:'Nutzerrechte pro Reiter und Thema, unsichtbare Reiter ohne Berechtigung und Teilen zwischen Nutzern wurden aufgenommen.',jump:'settings:rechte'},
 {title:'V26.07.02 · Bearbeitung',body:'Zellen, Zeilen und Spalten wurden editierbar, sortierbar und per Griff verschiebbar geplant.',jump:'table'},
 {title:'V26.07.01 · Start',body:'Erste Web-App aus der Original-Excel mit Reitern, Themen, Sortierung, Notizen, Links und Bildern.',jump:'top'}];}
function guideSections(){return [
 {title:'1. Login und Demo-Steuerung',body:'Melde dich mit admin/admin, max/demo, lea/demo oder gast/demo an. Die Demo-Steuerung unten links dient nur zum schnellen Testen der Vorschau.',jump:'top'},
 {title:'2. Reiterarten',body:'Autoreiter besitzen ein Laufleistungsfeld und Fahrzeug-Dateien. Neutrale Reiter enthalten normale Themen/Tabellen. Der Teilekatalog ist global.',jump:'vehicle'},
 {title:'3. Laufleistung',body:'Im Autoreiter wird nur die Laufleistung oben eingetragen. 143430 wird automatisch zu 143.430 Km formatiert und für Berechnungen genutzt.',jump:'vehicle'},
 {title:'4. Tabellen bearbeiten',body:'Zellen und Kopfzellen per Doppelklick bearbeiten. Links bleiben per einfachem Klick anklickbar. Das Stiftsymbol öffnet Details wie Notiz, Link, Farbe und Bild.',jump:'table'},
 {title:'5. Zeilen hinzufügen',body:'Zwischen jeder Zeile befindet sich ein Einfüge-Button. Unten kann zusätzlich eine neue Zeile angelegt werden.',jump:'table'},
 {title:'6. Sortieren',body:'Jedes Thema besitzt eine eigene Sortierung. Sortiert wird nur dieses Thema, keine anderen Tabellen im selben Reiter.',jump:'table'},
 {title:'7. Dropdowns',body:'Dropdown-Inhalte und Spaltenzuordnungen liegen in den Einstellungen. Datenpool ist kein sichtbarer Reiter mehr.',jump:'settings:dropdowns'},
 {title:'8. Dateien und Bilder',body:'Autoreiter haben eine Ordnerstruktur. Ein Ordner öffnet direkt darunter seine Unterordner und Dateien.',jump:'files'},
 {title:'9. Teilekatalog',body:'Der globale Teilekatalog ist für alle Nutzer gleich. Admins verwalten Ordnerbilder, Icons und Dateitabelle.',jump:'catalog'},
 {title:'10. Rechte und Teilen',body:'Reiter können mit einzelnen Nutzern geteilt werden. Ohne Berechtigung sind sie komplett unsichtbar.',jump:'settings:rechte'},
 {title:'11. Design, Logo und Farben',body:'Admin kann Logo, Favicon, Logogröße, Appfarben, Textfarben und Standardfarben für neue Nutzer festlegen. Nutzer können eigene Farben für sich ändern.',jump:'settings:design'}];}
function adminGuide(){return [
 {title:'Admin · Nutzerverwaltung',body:'Admin sieht Systemreiter und verwaltet Nutzer, aber nicht automatisch private Nutzerreiter. Freigabe erfolgt explizit durch Teilen.',jump:'settings:nutzer'},
 {title:'Admin · Standardfarben',body:'Im Farben-Bereich können Standardfarben für neu registrierte Nutzer gespeichert werden. Spätere Nutzeränderungen gelten nur für den jeweiligen Nutzer.',jump:'settings:farben'},
 {title:'Admin · Logo',body:'Logo per Drag & Drop ändern, Größe per Slider festlegen und Favicon automatisch aktualisieren.',jump:'settings:design'},
 {title:'Admin · Teilekatalog',body:'Der Teilekatalog ist global. Ordner können Text/Icon oder Bild nutzen. Dateien springen zur Tabelle.',jump:'catalog'},
 {title:'Admin · Supabase später',body:'Nach Freigabe dieser HTML wird die echte Supabase-Version mit SQL, Auth, Storage und RLS-Regeln erstellt.',jump:'settings:daten'}];}
function jumpTo(j){closeModal(); if(j.startsWith('settings:')){openSettings(j.split(':')[1]);return} const sheet=ensureActive(); if(j==='catalog'){const cat=allSheetsForUser().find(s=>s.catalog); if(cat){state.activeSheetId=cat.id;render();}return} if(j==='files')document.getElementById('vehicleFiles')?.scrollIntoView({behavior:'smooth'}); else if(j==='vehicle')document.querySelector('.vehicle-panel')?.scrollIntoView({behavior:'smooth'}); else if(j==='table')document.querySelector('.topic-card')?.scrollIntoView({behavior:'smooth'}); else window.scrollTo({top:0,behavior:'smooth'});}
function addSheetDialog(){showModal('Neuen Reiter anlegen',`<div class="field"><label>Name</label><input id="newSheetName" value="Neues Auto"></div><div class="field"><label>Art</label><select id="newSheetType"><option value="vehicle">Auto anlegen</option><option value="neutral">Neutraler Reiter</option></select></div>`,[{text:'Anlegen',action:()=>{const name=$('#newSheetName').value.trim()||'Neuer Reiter'; const type=$('#newSheetType').value; const u=currentUser(); const target=u.id==='admin'?state.adminSheets:state.userSheets[u.id]; const orig=window.WORKBOOK_DATA.sheets.find(s=>s.name===(type==='vehicle'?'A3 8P':'Werkzeug')); const s=emptyFrom(orig,name,u.id,type==='vehicle'); target.push(s); state.activeSheetId=s.id; save(); closeModal(); render();}},{text:'Abbrechen',secondary:true,action:closeModal}]);}
function archiveSheet(sheet){if(sheet.global){toast('Globaler Teilekatalog kann nicht archiviert werden.');return} if(confirm('Reiter archivieren?')){sheet.archived=true; save(); render();}}
function deleteSheet(sheet){if(sheet.system){toast('Systemreiter werden in der Vorschau nicht gelöscht.');return} if(confirm('Reiter wirklich löschen?')&&confirm('Zweite Bestätigung: endgültig löschen?')){const arr=currentUser().id==='admin'?state.adminSheets:state.userSheets[currentUser().id]; const i=arr.findIndex(s=>s.id===sheet.id); if(i>=0)arr.splice(i,1); save(); render();}}
function moveTab(from,to){const arr=allSheetsForUser(); const order=arr.map(s=>s.id); const a=order.indexOf(from), b=order.indexOf(to); if(a<0||b<0)return; const [x]=order.splice(a,1); order.splice(b,0,x); const pool=currentUser().id==='admin'?state.adminSheets:state.userSheets[currentUser().id]; pool.sort((x,y)=>order.indexOf(x.id)-order.indexOf(y.id)); save(); render();}
function bindTabResize(){let resizing=null; document.querySelectorAll('[data-resize-tab]').forEach(h=>{h.onmousedown=e=>{e.stopPropagation(); resizing={id:h.dataset.resizeTab,start:e.clientX,w:sheetById(h.dataset.resizeTab)?.tabWidth||130}; document.body.style.cursor='ew-resize';};}); document.onmousemove=e=>{if(resizing){const s=sheetById(resizing.id); if(s){s.tabWidth=Math.max(90,Math.min(360,resizing.w+e.clientX-resizing.start)); renderTabs();}}}; document.onmouseup=()=>{if(resizing){resizing=null;document.body.style.cursor='';save();}};}
function login(username,password){const u=state.users.find(x=>(x.username===username||x.email===username)&&x.password===password); if(!u){toast('Login fehlgeschlagen');return} state.currentUserId=u.id; ensureActive(); save(); render();}
function logout(){state.currentUserId=null; save(); render();}
function register(){const name=$('#regName').value.trim(), mail=$('#regEmail').value.trim(), pass=$('#regPass').value.trim()||'demo'; if(!name||!mail){toast('Name und E-Mail fehlen');return} const id=slug(name)+'_'+uid('u'); state.users.push({id,name,username:slug(name),email:mail,password:pass,role:'Nutzer',abo:'Testphase'}); state.userSheets[id]=[]; const auto=emptyFrom(window.WORKBOOK_DATA.sheets.find(s=>s.name==='A3 8P'),'Auto',id,true); state.userSheets[id].push(auto); state.userSheets[id].push(emptyFrom(window.WORKBOOK_DATA.sheets.find(s=>s.name==='Teilebestand'),'Teilebestand',id,false)); state.userSheets[id].push(emptyFrom(window.WORKBOOK_DATA.sheets.find(s=>s.name==='Werkzeug'),'Werkzeug',id,false)); state.userSheets[id].push(state.adminSheets.find(s=>s.catalog)); if(state.settings.adminDefaultColors)state.userColorOverrides[id]=clone(state.settings.adminDefaultColors); state.currentUserId=id; save(); render();}
function renderDemoBar(){const bar=$('#demoBar'); if(!currentUser()){bar.innerHTML='';return} const sheets=allSheetsForUser(); bar.innerHTML=`<strong>Demo-Steuerung</strong><select id="demoUser">${state.users.map(u=>`<option value="${u.id}" ${u.id===state.currentUserId?'selected':''}>${esc(u.name)} · ${esc(u.role)}</option>`).join('')}</select><select id="demoSheet">${sheets.map(s=>`<option value="${s.id}" ${s.id===state.activeSheetId?'selected':''}>${esc(s.name)}</option>`).join('')}</select><button class="small secondary" onclick="openSettings('profil')">Einstellungen</button><button class="small secondary" onclick="openRelease()">Release/Anleitung</button>`; $('#demoUser').onchange=e=>{state.currentUserId=e.target.value; ensureActive(); save(); render()}; $('#demoSheet').onchange=e=>{state.activeSheetId=e.target.value; save(); render()};}
function download(name,content,type){const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([content],{type})); a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
function bindGlobal(){document.addEventListener('mouseover',e=>{const el=e.target.closest('[title]'); const title=el?.getAttribute('title'); if(title){if(!tooltipEl){tooltipEl=document.createElement('div'); tooltipEl.className='tooltip'; document.body.appendChild(tooltipEl);} tooltipEl.textContent=title; tooltipEl.style.left=(e.clientX+12)+'px'; tooltipEl.style.top=(e.clientY+12)+'px'; el.setAttribute('data-old-title',title); el.removeAttribute('title');}}); document.addEventListener('mouseout',e=>{const el=e.target.closest('[data-old-title]'); if(el){el.setAttribute('title',el.getAttribute('data-old-title')); el.removeAttribute('data-old-title');} if(tooltipEl)tooltipEl.remove(),tooltipEl=null;}); $('#loginBtn').onclick=()=>login($('#loginUser').value.trim(),$('#loginPass').value.trim()); $('#loginPass').onkeydown=e=>{if(e.key==='Enter')$('#loginBtn').click()}; $('#logoutBtn').onclick=logout; $('#registerBtn').onclick=register; $$('.quick-login').forEach(b=>b.onclick=()=>{const u=state.users.find(x=>x.id===b.dataset.user); login(u.username,u.password)}); $('#settingsGear').onclick=()=>openSettings('profil'); $('#releaseBtn').onclick=openRelease; $('#saveBtn').onclick=()=>{save();toast('Manuell gespeichert')}; $('#darkToggle').onclick=()=>{state.settings.dark=!state.settings.dark;save();render()};}


/* ---------------- V26.07.15: stabiler Neuaufbau aus V26.07.11, ohne rekursive Patch-Initialisierung ---------------- */
const DEMO_USERS_V260715 = [
  {id:'admin',name:'Admin',username:'admin',email:'admin@kfz.local',password:'admin',role:'Admin',abo:'Adminzugang'},
  {id:'max',name:'Max Mustermann',username:'max',email:'max@kfz.local',password:'demo',role:'Nutzer',abo:'Testphase'},
  {id:'lea',name:'Lea Schrauber',username:'lea',email:'lea@kfz.local',password:'demo',role:'Nutzer',abo:'Monatlich'},
  {id:'gast',name:'Gast Ansicht',username:'gast',email:'gast@kfz.local',password:'demo',role:'Gast',abo:'Gast/Test'}
];
function ensureDemoUsersV260715(){
  if(!state) return;
  state.users = Array.isArray(state.users) ? state.users : [];
  state.userSheets = state.userSheets || {};
  for(const du of DEMO_USERS_V260715){
    let u = state.users.find(x => String(x.id||'').toLowerCase()===du.id || String(x.username||'').toLowerCase()===du.username || String(x.email||'').toLowerCase()===du.email.toLowerCase());
    if(!u){ u = clone(du); state.users.push(u); }
    u.id = du.id; u.username = du.username; u.email = du.email; u.password = du.password; u.role = u.role || du.role; u.name = u.name || du.name; u.abo = u.abo || du.abo;
    if(du.id !== 'admin') state.userSheets[du.id] = Array.isArray(state.userSheets[du.id]) ? state.userSheets[du.id] : [];
  }
}
function demoLoginV260715(identifier, password){
  ensureDemoUsersV260715();
  const q = String(identifier||'').trim().toLowerCase();
  const p = String(password||'');
  const u = state.users.find(x => [x.id,x.username,x.email,x.name].some(v => String(v||'').trim().toLowerCase()===q));
  if(!u || !(p==='__quick_demo__' || String(u.password||'')===p)){
    toast('Login fehlgeschlagen. Demo-Zugänge: admin/admin, max/demo, lea/demo, gast/demo.');
    return;
  }
  state.currentUserId = u.id;
  try{ ensureActive(); } catch(e){ console.warn('ensureActive fehlgeschlagen', e); state.activeSheetId = null; }
  save();
  render();
  toast('Eingeloggt als ' + (u.name || u.username));
}
login = demoLoginV260715;
const bindGlobalBaseV260715 = bindGlobal;
bindGlobal = function(){
  bindGlobalBaseV260715();
  const lb = $('#loginBtn'); if(lb) lb.onclick = () => demoLoginV260715($('#loginUser').value.trim(), $('#loginPass').value.trim());
  const lp = $('#loginPass'); if(lp) lp.onkeydown = e => { if(e.key === 'Enter') lb?.click(); };
  $$('.quick-login').forEach(b => b.onclick = () => demoLoginV260715(b.dataset.user, '__quick_demo__'));
};
const initStateBaseV260715 = initState;
initState = function(){
  let s = initStateBaseV260715();
  s.version = VERSION;
  s.users = Array.isArray(s.users) ? s.users : clone(DEMO_USERS_V260715);
  s.adminSheets = Array.isArray(s.adminSheets) ? s.adminSheets : [];
  s.userSheets = s.userSheets || {};
  state = s;
  ensureDemoUsersV260715();
  state = null;
  return s;
};
const releaseNotesBaseV260715 = releaseNotes;
releaseNotes = function(){return [
  {title:'V26.07.15 · Demo wieder lauffähig',body:'Stabiler Neuaufbau aus der funktionierenden V26.07.11-Basis. Die rekursiven Patch-Initialisierungen aus V26.07.12 bis V26.07.14 wurden nicht übernommen. Demo-Login funktioniert wieder mit admin/admin, max/demo, lea/demo und gast/demo.',jump:'top'},
  {title:'V26.07.14 · verworfen',body:'Diese Zwischenversion enthielt einen fehlerhaften Login-Patch und wird nicht als Basis weiterverwendet.',jump:'top'},
  {title:'V26.07.13 · Funktionsabgleich vorgesehen',body:'LH-Funktionen und Excel-Abgleich bleiben als Ziel erhalten, werden aber künftig nur noch sauber und einzeln auf die stabile Basis übernommen.',jump:'settings:daten'},
  ...releaseNotesBaseV260715()
];};

/* ---------------- V26.07.16: vollständiger Funktions-Patch auf stabiler V26.07.15-Basis ---------------- */
(function(){
  const LIVE_STORE_KEY = 'kfz_app_live_state_preserved';
  const VERSION_16 = 'V26.07.16';
  const ICON_POOL = ['🚗','🚙','🛠️','🧰','🔧','⚙️','📁','🗂️','📄','🖼️','🧾','🔩','🛞','💡','🧪','🧼','🛢️','🔋','🪛','📦','⭐','✅','⚠️'];
  function safeClone(o){ try{return clone(o)}catch(e){return JSON.parse(JSON.stringify(o||{}))} }
  function globalCatalog(){ return (state?.adminSheets||[]).find(s=>s.catalog) || (state?.userSheets && Object.values(state.userSheets).flat().find(s=>s.catalog)); }
  function allStateSheets(){ return [...(state?.adminSheets||[]), ...Object.values(state?.userSheets||{}).flat()].filter(Boolean); }
  function normalizeColumnParts(col){
    if(!col.parts) col.parts = headerParts(col.label || col.letter || 'Spalte','');
    col.parts = col.parts.map(x=>String(x||'').trim()).filter(Boolean);
    if(!col.parts.length) col.parts=['Spalte'];
  }
  function normalizeVerbauliste(sheet){
    if(!sheet || !Array.isArray(sheet.themes)) return;
    const t = sheet.themes.find(x => /Verbauliste/i.test(String(x.name||x.displayName||'')));
    if(!t || !Array.isArray(t.columns)) return;
    t.displayName = 'Verbauliste';
    for(const col of t.columns) normalizeColumnParts(col);
    const byCol = new Map(t.columns.map(c=>[Number(c.col), c]));
    const setParts = (c,parts)=>{ const col=byCol.get(c); if(!col) return; col.parts=parts.slice(); col.label=parts.join(' / '); };
    setParts(2,['Menge','Verfügbar']);
    setParts(3,['Menge','verbaut']);
    setParts(4,['Beschreibung','Beschreibung']);
    setParts(5,['Beschreibung','Kategorie']);
    setParts(6,['Beschreibung','unter Kategorie']);
    setParts(7,['Artikel','Hersteller']);
    setParts(8,['Artikel','Teilenummer']);
    setParts(9,['Verarbeitung','Einbaudatum']);
    setParts(10,['Verarbeitung','Kilometerstand']);
    setParts(11,['Verarbeitung','nächster Tausch','Laufleistung']);
    setParts(12,['Verarbeitung','nächster Tausch','Datum']);
    setParts(13,['Verarbeitung','nächster Tausch','Tage bis']);
    setParts(14,['Verarbeitung','nächster Tausch','Anteil bis']);
  }
  function ensureTreeBasics(nodes){
    (nodes||[]).forEach(n=>{
      n.id = n.id || uid(n.type || 'node');
      n.type = n.type || 'folder';
      n.name = n.name || (n.type === 'folder' ? 'Ordner' : 'Datei');
      if(n.type === 'folder') n.children = Array.isArray(n.children) ? n.children : [];
      if(n.type === 'folder' && !n.icon && !n.thumb) n.icon='📁';
      if(n.type === 'file' && !n.icon && !n.thumb) n.icon='📄';
      ensureTreeBasics(n.children);
    });
  }
  function ensureCatalogV16(){
    let cat = globalCatalog();
    if(!cat){
      cat = createCatalogSheet();
      state.adminSheets = state.adminSheets || [];
      state.adminSheets.push(cat);
    }
    cat.id = 'global_parts_catalog';
    cat.name = 'Teilekatalog';
    cat.global = true; cat.catalog = true; cat.system = true;
    cat.permissions = cat.permissions || {admin:'manage','role:Nutzer':'view','role:Gast':'view'};
    cat.tree = Array.isArray(cat.tree) && cat.tree.length ? cat.tree : createCatalogSheet().tree;
    cat.rows = Array.isArray(cat.rows) && cat.rows.length ? cat.rows : createCatalogSheet().rows;
    ensureTreeBasics(cat.tree);
    state.catalogBrands = Array.isArray(state.catalogBrands) && state.catalogBrands.length ? state.catalogBrands : [
      {id:'audi',name:'Audi',logo:'🚗',models:[{id:'a3_8p',name:'A3 8P',logo:'🚙'},{id:'neutral',name:'Allgemein',logo:'🧩'}]},
      {id:'vw',name:'Volkswagen',logo:'🚘',models:[{id:'allgemein',name:'Allgemein',logo:'🧩'}]}
    ];
    state.catalogSelection = state.catalogSelection || {brand:state.catalogBrands[0].id,model:state.catalogBrands[0].models[0].id};
    state.catalogSelected = state.catalogSelected || findFirstFileId(cat.tree) || '';
    return cat;
  }
  function findFirstFileId(nodes){ for(const n of nodes||[]){ if(n.type==='file') return n.id; const v=findFirstFileId(n.children); if(v) return v; } return ''; }
  function migrateStateV16(s){
    state = s;
    s.version = VERSION_16;
    s.users = Array.isArray(s.users) && s.users.length ? s.users : safeClone(USERS);
    s.userSheets = s.userSheets || {};
    s.adminSheets = Array.isArray(s.adminSheets) ? s.adminSheets : [];
    s.settings = s.settings || {};
    s.settings.colors = Object.assign({primary:'#1f4e79',primaryDark:'#153a5a',secondary:'#d9eaf7',accent:'#fff2cc',bg:'#f4f7fb',surface:'#ffffff',text:'#19202a',tableText:'#111827',headText:'#12355b',border:'#d7dee8'}, s.settings.colors||{});
    s.settings.logo = s.settings.logo || DEFAULT_LOGO;
    s.settings.logoSize = s.settings.logoSize || 42;
    s.publishedAt = s.publishedAt || null;
    s.adminParameters = s.adminParameters || {excelSource:'Wartungsliste KFZ Details - 2(1).xlsm',lhReference:'LH_Boden_MTV_VTV_Online_Browser_App_V26.07.48_Aktualisiert_Kernblock_Bearbeitung_Netlify',preserveAdminChanges:true,preserveUserData:true};
    ensureDemoUsersV260715?.();
    for(const sh of allStateSheets()){
      (sh.themes||[]).forEach(t=>{ (t.columns||[]).forEach(normalizeColumnParts); });
      normalizeVerbauliste(sh);
      if(sh.vehicle && !Array.isArray(sh.files)) sh.files = defaultFileTree();
      ensureTreeBasics(sh.files);
      sh.icon = sh.icon || (sh.vehicle ? '🚗' : sh.catalog ? '🧩' : '📄');
    }
    const cat = ensureCatalogV16();
    for(const uidKey of Object.keys(s.userSheets||{})){
      s.userSheets[uidKey] = (s.userSheets[uidKey]||[]).filter(sh => !sh.catalog);
    }
    state = null;
    return s;
  }
  function loadPreservedState(){
    const keys = [LIVE_STORE_KEY,'kfz_app_V26.07.15_state','kfz_app_V26.07.14_state','kfz_app_V26.07.13_state','kfz_app_V26.07.12_state','kfz_app_V26.07.11_state'];
    for(const k of keys){
      try{ const raw=localStorage.getItem(k); if(!raw) continue; const parsed=JSON.parse(raw); if(parsed && parsed.users && (parsed.adminSheets||parsed.userSheets)) return parsed; }catch(e){}
    }
    return null;
  }
  const baseInitV16 = initState;
  initState = function(){
    const live = loadPreservedState();
    const base = live || baseInitV16();
    return migrateStateV16(base);
  };
  save = function(){
    if(!state) return;
    state.version = VERSION_16;
    state.savedAt = new Date().toISOString();
    try{
      localStorage.setItem(LIVE_STORE_KEY, JSON.stringify(state));
      localStorage.setItem('kfz_app_'+VERSION_16+'_state', JSON.stringify(state));
    }catch(e){ console.warn(e); }
    updateSaveInfo?.();
  };
  allSheetsForUser = function(){
    const u=currentUser(); if(!u) return [];
    const cat=globalCatalog();
    let arr = u.id==='admin' ? (state.adminSheets||[]).filter(s=>!s.catalog) : (state.userSheets[u.id]||[]).filter(s=>!s.catalog);
    if(cat && canAccess(accessFor(cat,u),'view')) arr = arr.concat([cat]);
    return arr.filter(s=>canAccess(accessFor(s,u),'view') && !s.archived);
  };
  accessFor = function(sheet,user=currentUser()){
    if(!sheet||!user)return '';
    if(sheet.catalog && user.role==='Admin') return 'manage';
    if(sheet.ownerId===user.id)return 'manage';
    if(sheet.permissions?.[user.id])return sheet.permissions[user.id];
    if(sheet.permissions?.['role:'+user.role])return sheet.permissions['role:'+user.role];
    if(sheet.system && user.id==='admin') return 'manage';
    return '';
  };
  const baseApplyThemeV16 = applyTheme;
  applyTheme = function(){ baseApplyThemeV16(); const c=state?.settings?.colors||{}; const root=document.documentElement; ['text','tableText','headText','surface','bg','border','accent'].forEach(k=>{ if(c[k]) root.style.setProperty('--'+k,c[k]); }); };
  function tabIconHtml(s){
    if(s.logo) return `<img class="tab-logo" src="${esc(s.logo)}" alt="">`;
    return `<span class="tab-icon">${esc(s.icon || (s.vehicle?'🚗':s.catalog?'🧩':'📄'))}</span>`;
  }
  renderTabs = function(){
    const tabs=$('#tabs'); const sheets=allSheetsForUser();
    tabs.innerHTML=sheets.map(s=>`<button class="tab ${s.id===state.activeSheetId?'active':''}" draggable="true" data-tab="${esc(s.id)}" style="--tabW:${s.tabWidth||140}px">${tabIconHtml(s)}<span>${esc(s.name)}</span><span class="type">${s.vehicle?'Auto':s.catalog?'global':'Reiter'}</span><span class="tab-resize" data-resize-tab="${esc(s.id)}"></span></button>`).join('')+`<button class="tab secondary" id="addTabBtn">+ Reiter</button>`;
    tabs.onclick=e=>{const b=e.target.closest('[data-tab]'); if(b){state.activeSheetId=b.dataset.tab; save(); render();} if(e.target.closest('#addTabBtn')) addSheetDialog();};
    tabs.ondragstart=e=>{const b=e.target.closest('[data-tab]'); if(b){dragged={type:'tab',id:b.dataset.tab};}};
    tabs.ondragover=e=>{if(dragged?.type==='tab'){e.preventDefault(); const b=e.target.closest('[data-tab]'); if(b)b.classList.add('drag-over')}};
    tabs.ondragleave=e=>e.target.closest('[data-tab]')?.classList.remove('drag-over');
    tabs.ondrop=e=>{const b=e.target.closest('[data-tab]'); $$('.tab').forEach(x=>x.classList.remove('drag-over')); if(b&&dragged?.type==='tab')moveTab(dragged.id,b.dataset.tab); dragged=null;};
    bindTabResize();
  };
  renderMain = function(){
    const sheet=ensureActive(); const main=$('#main');
    if(!sheet){main.innerHTML='<div class="login-panel"><h2>Keine Reiter verfügbar</h2></div>'; return;}
    if(sheet.catalog){renderCatalog(sheet); return;}
    const access=accessFor(sheet);
    main.innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>${esc(sheet.name)}</h2><div class="sheet-sub"><span class="pill ${sheet.vehicle?'ok':''}">${sheet.vehicle?'Autoreiter':'Neutraler Reiter'}</span><span class="pill">Zugriff: ${accessLabel(access)}</span>${sheet.ownerId!=='system'?`<span class="pill warn">Besitzer: ${esc(userName(sheet.ownerId))}</span>`:''}</div></div><div class="sheet-tools"><button class="secondary" id="editTabIconBtn">Reiterlogo/Icon</button><button class="secondary" id="shareTabBtn">Teilen/Rechte</button><button class="secondary" id="archiveTabBtn">Archivieren</button><button class="danger" id="deleteTabBtn">Löschen</button></div></div>${sheet.vehicle?vehicleHtml(sheet):''}<div class="global-search"><strong>🔎 Teilenummer / Partnummer</strong><input id="partSearch" placeholder="PN, OEM, SN, Teilenummer suchen …"><button id="partSearchBtn">Suchen</button></div><div id="partResults" class="search-results"></div>${sheet.vehicle?renderFileSection(sheet,access):''}<div class="topic-nav">${(sheet.themes||[]).map(t=>`<a href="#${esc(topicDomId(sheet,t))}">${esc(t.displayName||t.name)}</a>`).join('')}</div><div class="topic-stack">${(sheet.themes||[]).map(t=>renderTopic(sheet,t,access)).join('')}</div>`;
    bindSheetEvents(sheet,access);
    $('#editTabIconBtn')?.addEventListener('click',()=>editSheetIcon(sheet));
  };
  renderTopic = function(sheet,t,access){
    const canEdit=canAccess(access,'edit'); normalizeVerbauliste(sheet);
    return `<section class="topic-card" id="${topicDomId(sheet,t)}" data-topic="${esc(t.key||topicKey(t))}" style="width:${t.uiWidth||'auto'};height:${t.uiHeight||'auto'}"><div class="topic-head" draggable="${canEdit}" data-drag-topic="${esc(t.key||topicKey(t))}"><h3>${esc(t.displayName||t.name)}</h3><div class="topic-actions"><select class="sort-select" data-sort-select>${(t.columns||[]).map((c,i)=>`<option value="${i}">${esc(shortHeader(c.parts))}</option>`).join('')}</select><button class="small secondary" data-sort="asc">↑</button><button class="small secondary" data-sort="desc">↓</button><button class="small secondary" data-add-col>+ Spalte</button><button class="small secondary" data-collapse>Ein-/ausklappen</button></div></div><div class="topic-body"><div class="table-wrap"><table class="kfz-table">${renderHeader(sheet,t,canEdit)}<tbody>${renderRows(sheet,t,canEdit)}</tbody></table></div></div><div class="topic-foot"><button class="secondary" data-add-row-end>+ Zeile unten</button><span class="hint">Doppelklick bearbeitet · Kopfzellen können verbunden/geteilt werden · Links bleiben anklickbar</span></div></section>`;
  };
  renderHeader = function(sheet,t,canEdit){
    const cols=t.columns||[]; cols.forEach(normalizeColumnParts);
    const maxDepth=Math.max(1,...cols.map(c=>(c.parts||[]).length)); let html='<thead class="modern-head">';
    for(let tier=0;tier<maxDepth;tier++){
      html+='<tr>'; if(tier===0)html+=`<th class="row-tools" rowspan="${maxDepth}">↕</th>`;
      let i=0;
      while(i<cols.length){
        const parts=cols[i].parts||['Spalte']; if(tier>=parts.length){i++; continue;}
        const label=parts[tier]||''; const prefix=parts.slice(0,tier).join('|'); let span=1;
        while(i+span<cols.length){const p2=cols[i+span].parts||[]; if(tier>=p2.length)break; if((p2[tier]||'')!==label)break; if(p2.slice(0,tier).join('|')!==prefix)break; span++;}
        const rowspan=(tier===parts.length-1)?(maxDepth-tier):1;
        html+=`<th class="editable-head" draggable="${canEdit}" data-head-tier="${tier}" data-head-start="${i}" data-head-span="${span}" data-col-index="${i}" colspan="${span}" rowspan="${rowspan}" title="Doppelklick: Text bearbeiten / verbinden / teilen"><span>${esc(label)}</span>${canEdit?'<button class="head-tool" data-head-edit>✎</button>':''}</th>`;
        i+=span;
      }
      html+='</tr>';
    }
    return html+'</thead>';
  };
  function findFormulaCol(t,kind){
    return (t.columns||[]).find(cc=>{let l=(cc.parts||[]).join(' / ').toLowerCase(); return l.includes('nächster tausch') && l.includes(kind);});
  }
  recalcCellIfNeeded = function(sheet,t,r,col,cell){
    const h=(col.parts||[]).join(' / ').toLowerCase();
    if(!(h.includes('tage bis')||h.includes('anteil bis'))) return;
    const cols=t.columns||[]; const find=(s)=>cols.find(cc=>(cc.parts||[]).join(' / ').toLowerCase().includes(s));
    const installDate=find('einbaudatum'); const installKm=find('kilometerstand'); const nextKm=findFormulaCol(t,'laufleistung'); const nextDate=findFormulaCol(t,'datum');
    const get=cc=>cc?sheet.cells?.[`${r}:${cc.col}`]?.text:'';
    const base=parseNum(get(installKm)); const next=parseNum(get(nextKm)); const current=Number(sheet.mileage||0);
    let targetKm = NaN;
    if(!isNaN(next)) targetKm = (!isNaN(base) && next <= Math.max(base,100000)) ? base + next : next;
    if(h.includes('tage bis')){
      let days='-'; const d=parseDate(get(nextDate));
      if(d) days=String(daysBetween(TODAY,d));
      else if(!isNaN(targetKm) && current) days=String(Math.round((targetKm-current)/50));
      cell.text=days; cell.raw=days;
    }
    if(h.includes('anteil bis')){
      let pct='-'; const d0=parseDate(get(installDate)), d1=parseDate(get(nextDate));
      if(d0&&d1&&d1>d0){pct=(Math.max(0,Math.min(100,daysBetween(d0,TODAY)/daysBetween(d0,d1)*100))).toFixed(1).replace('.',',')+'%';}
      else if(!isNaN(base)&&!isNaN(targetKm)&&targetKm>base&&current){pct=(Math.max(0,Math.min(100,(current-base)/(targetKm-base)*100))).toFixed(1).replace('.',',')+'%';}
      cell.text=pct; cell.raw=pct;
    }
  };
  function highlightElement(el, cls='highlight-cell'){
    if(!el) return; document.querySelectorAll('.highlight-cell,.highlight-row').forEach(x=>x.classList.remove('highlight-cell','highlight-row'));
    el.classList.add(cls); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'});
    setTimeout(()=>el.classList.remove(cls),7000);
  }
  partSearch = function(sheet){
    const q=$('#partSearch')?.value.trim().toLowerCase(); const out=$('#partResults'); if(!q||!out)return;
    const hits=[];
    for(const t of sheet.themes||[]) for(const col of t.columns||[]){
      const h=(col.parts||[]).join(' ').toLowerCase();
      if(/teil|part|pn|oem|sn|nummer|hersteller/.test(h)){
        for(const g of t.rowGroups||[]){ const r=Array.isArray(g)?g[0]:g; const v=String(sheet.cells?.[`${r}:${col.col}`]?.text||''); if(v.toLowerCase().includes(q)) hits.push({theme:t,row:r,col,value:v}); }
      }
    }
    out.innerHTML=hits.length?hits.slice(0,80).map((h,i)=>`<div class="search-hit"><div><strong>${esc(h.value)}</strong><div class="hint">${esc(h.theme.displayName||h.theme.name)} · ${esc(shortHeader(h.col.parts))}</div></div><button class="small secondary" data-hit="${i}">Springen & markieren</button></div>`).join(''):'<div class="search-hit">Keine Treffer.</div>';
    out.querySelectorAll('[data-hit]').forEach(btn=>btn.onclick=()=>{const h=hits[+btn.dataset.hit]; const id=topicDomId(sheet,h.theme); render(); setTimeout(()=>highlightElement(document.querySelector(`#${CSS.escape(id)} td[data-r="${h.row}"][data-c="${h.col.col}"]`)),60);});
  };
  function treePathById(nodes,id,path=[]){
    for(const n of nodes||[]){ const p=path.concat(n.name); if(n.id===id) return p; const c=treePathById(n.children,id,p); if(c) return c; } return null;
  }
  renderFileSection = function(sheet,access){
    if(!sheet.files) sheet.files=defaultFileTree(); ensureTreeBasics(sheet.files);
    const selectedPath = treePathById(sheet.files, state.fileSelected) || [];
    return `<section class="file-section" id="vehicleFiles"><div class="section-head"><h3>Dateien & Bilder</h3><div>${canAccess(access,'edit')?'<button class="small secondary" data-file-add-root>+ Ordner</button> <button class="small secondary" data-file-upload-root>+ Datei/Bild</button>':''}</div></div><div class="path-line" id="filePath">Pfad: /${esc(selectedPath.join(' / '))}</div><div class="tree vertical-tree" id="fileTree">${renderTree(sheet.files,'file',access)}</div></section>`;
  };
  renderTree = function(nodes,kind,access,depth=0,path=[]){
    return (nodes||[]).map(n=>{
      const canEdit=kind==='catalog'?currentUser()?.role==='Admin':canAccess(access,'edit');
      const icon=n.thumb?`<img class="thumb" src="${n.thumb}" alt="">`:`<span class="thumb">${esc(n.icon || (n.type==='folder'?'📁':'📄'))}</span>`;
      return `<div class="tree-node" data-node="${esc(n.id)}" data-kind="${kind}"><div class="tree-row" style="margin-left:${depth*14}px" data-tree-click><span class="tree-toggle">${n.type==='folder'?(n.open?'▾':'▸'):'•'}</span>${icon}<strong>${esc(n.name)}</strong>${n.part?`<span class="pill">${esc(n.part)}</span>`:''}<span style="flex:1"></span>${canEdit?`<button class="small secondary" data-tree-edit>✎</button>${n.type==='folder'?'<button class="small secondary" data-tree-add-folder>+ Ordner</button><button class="small secondary" data-tree-add-file>+ Datei</button>':''}<button class="small danger" data-tree-delete>🗑</button>`:''}</div>${n.type==='folder'&&n.open?`<div class="tree-children">${renderTree(n.children||[],kind,access,depth+1,path.concat(n.name))}</div>`:''}</div>`;
    }).join('');
  };
  function removeNode(nodes,id){
    if(!Array.isArray(nodes)) return null;
    const idx=nodes.findIndex(n=>n.id===id); if(idx>=0) return nodes.splice(idx,1)[0];
    for(const n of nodes){ const r=removeNode(n.children,id); if(r) return r; }
    return null;
  }
  function parentNode(nodes,id){
    for(const n of nodes||[]){ if((n.children||[]).some(c=>c.id===id)) return n; const p=parentNode(n.children,id); if(p) return p; } return null;
  }
  bindTree = function(container,root,kind,access){
    if(!container) return;
    container.onclick=e=>{
      const nodeEl=e.target.closest('[data-node]'); if(!nodeEl)return;
      const node=findNode(root,nodeEl.dataset.node); if(!node)return;
      if(e.target.closest('[data-tree-delete]')){
        if(confirm(`${node.type==='folder'?'Ordner':'Datei'} „${node.name}“ löschen?`)){
          const removed=removeNode(root,node.id);
          if(kind==='catalog' && removed?.rowId){ const cat=globalCatalog(); const i=(cat.rows||[]).findIndex(r=>r.id===removed.rowId); if(i>=0) cat.rows.splice(i,1); }
          save(); render(); toast('Gelöscht');
        }
        return;
      }
      if(e.target.closest('[data-tree-edit]')) return editTreeNode(node,kind);
      if(e.target.closest('[data-tree-add-folder]')) return addTreeChild(node,'folder',kind);
      if(e.target.closest('[data-tree-add-file]')) return addTreeChild(node,'file',kind);
      if(e.target.closest('[data-tree-click]')){
        if(node.type==='folder'){
          node.open=!node.open; if(kind==='file') state.fileSelected=node.id; save(); render();
        } else {
          if(kind==='file'){ state.fileSelected=node.id; save(); render(); }
          if(kind==='catalog'){
            state.catalogSelected=node.id; save(); render();
            setTimeout(()=>highlightElement(document.querySelector(`[data-catalog-row="${node.rowId}"]`),'highlight-row'),70);
          }
        }
      }
    };
  };
  function iconPoolHtml(current=''){
    return `<div class="icon-pool">${ICON_POOL.map(i=>`<button class="small secondary ${current===i?'active':''}" data-pick-icon="${esc(i)}">${esc(i)}</button>`).join('')}</div>`;
  }
  editTreeNode = function(node,kind){
    showModal((node.type==='folder'?'Ordner':'Datei')+' bearbeiten',`<div class="grid2"><div class="field"><label>Name</label><input id="tnName" value="${esc(node.name)}"></div><div class="field"><label>Teilenummer / Bildtext</label><input id="tnPart" value="${esc(node.part||'')}"></div><div class="field full"><label>Icon-Pool</label>${iconPoolHtml(node.icon)}</div><div class="field"><label>Icon/Text manuell</label><input id="tnIcon" value="${esc(node.icon||'')}"></div><div class="field"><label>Vorschau/Bild</label><input id="tnFile" type="file" accept="image/*"></div></div>`,[
      {text:'Speichern',action:()=>{node.name=$('#tnName').value.trim()||node.name; node.part=$('#tnPart').value.trim(); node.icon=$('#tnIcon').value.trim(); const f=$('#tnFile').files[0]; const done=()=>{save(); closeModal(); render();}; if(f){const rd=new FileReader(); rd.onload=()=>{node.thumb=rd.result; done();}; rd.readAsDataURL(f);} else done();}},
      {text:'Löschen',secondary:true,action:()=>{closeModal(); const cat=globalCatalog(); const currentSheet=sheetById(state.activeSheetId); const root = kind==='catalog'?cat.tree:(currentSheet?.files||[]); const p=parentNode(root,node.id); if(confirm('Wirklich löschen?')){removeNode(root,node.id); save(); render();}}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    $$('#modalRoot [data-pick-icon]').forEach(b=>b.onclick=()=>{$('#tnIcon').value=b.dataset.pickIcon; $$('#modalRoot [data-pick-icon]').forEach(x=>x.classList.remove('active')); b.classList.add('active');});
  };
  function selectedBrand(){ return (state.catalogBrands||[]).find(b=>b.id===state.catalogSelection?.brand) || state.catalogBrands?.[0]; }
  function selectedModel(){ const b=selectedBrand(); return (b?.models||[]).find(m=>m.id===state.catalogSelection?.model) || b?.models?.[0]; }
  function catalogModelControls(isAdmin){
    const brands=state.catalogBrands||[]; const b=selectedBrand(); const models=b?.models||[];
    return `<section class="catalog-section catalog-selector"><div class="section-head"><h3>Marke & Modell auswählen</h3>${isAdmin?'<div><button class="small secondary" id="addBrandBtn">+ Marke</button> <button class="small secondary" id="addModelBtn">+ Modell</button> <button class="small secondary" id="editBrandModelBtn">Logo/Text bearbeiten</button></div>':''}</div><div class="brand-model-grid"><div class="field"><label>Marke</label><select id="catalogBrandSelect">${brands.map(x=>`<option value="${esc(x.id)}" ${x.id===b?.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>Modell</label><select id="catalogModelSelect">${models.map(x=>`<option value="${esc(x.id)}" ${x.id===selectedModel()?.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="catalog-current"><span class="big-logo">${esc(b?.logo||'🚗')}</span><strong>${esc(b?.name||'Marke')}</strong><span> / </span><span class="big-logo">${esc(selectedModel()?.logo||'🧩')}</span><strong>${esc(selectedModel()?.name||'Modell')}</strong></div></div></section>`;
  }
  renderCatalog = function(sheet){
    ensureCatalogV16(); sheet=globalCatalog(); const isAdmin=currentUser()?.role==='Admin'; const selectedNode=findNode(sheet.tree,state.catalogSelected);
    $('#main').innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>Teilekatalog</h2><div class="sheet-sub"><span class="pill ok">Global für alle Nutzer</span><span class="pill">${isAdmin?'Admin bearbeitet':'Nur ansehen'}</span></div></div><div class="sheet-tools">${isAdmin?'<button id="catAddRoot" class="secondary">+ Hauptordner</button> <button id="catPublish" class="secondary">Katalog veröffentlichen</button>':''}</div></div>${catalogModelControls(isAdmin)}<section class="catalog-section"><div class="section-head"><h3>Ordnerstruktur</h3><span class="hint">Ordner öffnen direkt untereinander. Datei anklicken springt zur markierten Tabellenzeile.</span></div><div class="catalog-grid"><div><div class="tree vertical-tree" id="catalogTree">${renderTree(sheet.tree,'catalog','manage')}</div></div><div><h3>Dateitabelle</h3><table class="catalog-table"><thead><tr><th>Bild</th><th>Bildtext / Datei</th><th>Teilenummer</th><th>Kategorie</th><th>Notiz</th></tr></thead><tbody>${(sheet.rows||[]).map(r=>`<tr data-catalog-row="${esc(r.id)}" class="${selectedNode?.rowId===r.id?'highlight-row':''}"><td>${r.image?`<img class="catalog-preview" src="${r.image}">`:'<span class="thumb">🧩</span>'}</td><td><span class="jump-link" data-jump-row="${esc(r.id)}">${esc(r.name)}</span></td><td>${esc(r.part)}</td><td>${esc(r.category)}</td><td>${esc(r.note)}</td></tr>`).join('')}</tbody></table></div></div></section>`;
    $('#catalogBrandSelect').onchange=e=>{state.catalogSelection.brand=e.target.value; const b=selectedBrand(); state.catalogSelection.model=b?.models?.[0]?.id||''; save(); render();};
    $('#catalogModelSelect').onchange=e=>{state.catalogSelection.model=e.target.value; save(); render();};
    $('#addBrandBtn')?.addEventListener('click',()=>{const name=prompt('Neue Marke?'); if(!name)return; const id=slug(name)+'_'+uid('b'); state.catalogBrands.push({id,name,logo:'🚗',models:[{id:'allgemein_'+uid('m'),name:'Allgemein',logo:'🧩'}]}); state.catalogSelection={brand:id,model:state.catalogBrands.at(-1).models[0].id}; save(); render();});
    $('#addModelBtn')?.addEventListener('click',()=>{const b=selectedBrand(); if(!b)return; const name=prompt('Neues Modell?'); if(!name)return; const id=slug(name)+'_'+uid('m'); b.models.push({id,name,logo:'🚙'}); state.catalogSelection.model=id; save(); render();});
    $('#editBrandModelBtn')?.addEventListener('click',editBrandModel);
    $('#catPublish')?.addEventListener('click',()=>{state.publishedAt=new Date().toISOString(); save(); toast('Teilekatalog veröffentlicht und für künftige Versionen als Live-Daten gespeichert.');});
    $('#catAddRoot')?.addEventListener('click',()=>{sheet.tree.push({id:uid('folder'),type:'folder',name:'Neuer Hauptordner',open:true,children:[],icon:'📁'}); save(); render();});
    $$('#main [data-jump-row]').forEach(el=>el.onclick=()=>highlightElement(document.querySelector(`[data-catalog-row="${el.dataset.jumpRow}"]`),'highlight-row'));
    bindTree($('#catalogTree'),sheet.tree,'catalog','manage');
  };
  function editBrandModel(){
    const b=selectedBrand(), m=selectedModel();
    showModal('Marke / Modell bearbeiten',`<div class="grid2"><div class="field"><label>Marke</label><input id="bmBrand" value="${esc(b?.name||'')}"></div><div class="field"><label>Markenlogo/Icon</label><input id="bmBrandLogo" value="${esc(b?.logo||'')}"></div><div class="field"><label>Modell</label><input id="bmModel" value="${esc(m?.name||'')}"></div><div class="field"><label>Modelllogo/Icon</label><input id="bmModelLogo" value="${esc(m?.logo||'')}"></div><div class="field full"><label>Icon-Pool</label>${iconPoolHtml('')}</div></div>`,[
      {text:'Speichern',action:()=>{ if(b){b.name=$('#bmBrand').value.trim()||b.name; b.logo=$('#bmBrandLogo').value.trim()||b.logo;} if(m){m.name=$('#bmModel').value.trim()||m.name; m.logo=$('#bmModelLogo').value.trim()||m.logo;} save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    $$('#modalRoot [data-pick-icon]').forEach(btn=>btn.onclick=()=>{$('#bmModelLogo').value=btn.dataset.pickIcon;});
  }
  function editSheetIcon(sheet){
    showModal('Reiterlogo / Icon bearbeiten',`<div class="grid2"><div class="field"><label>Icon aus Pool</label>${iconPoolHtml(sheet.icon)}</div><div class="field"><label>Icon/Text manuell</label><input id="sheetIcon" value="${esc(sheet.icon||'')}"></div><div class="field full"><label>Eigenes Bild/Logo</label><input id="sheetLogoFile" type="file" accept="image/*"></div></div>`,[
      {text:'Speichern',action:()=>{sheet.icon=$('#sheetIcon').value.trim()||sheet.icon||'📄'; const f=$('#sheetLogoFile').files[0]; const done=()=>{save(); closeModal(); render();}; if(f){const rd=new FileReader(); rd.onload=()=>{sheet.logo=rd.result; done();}; rd.readAsDataURL(f);} else done();}},
      {text:'Bild entfernen',secondary:true,action:()=>{delete sheet.logo; save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    $$('#modalRoot [data-pick-icon]').forEach(b=>b.onclick=()=>{$('#sheetIcon').value=b.dataset.pickIcon; $$('#modalRoot [data-pick-icon]').forEach(x=>x.classList.remove('active')); b.classList.add('active');});
  }
  editHeader = function(t,th){
    const tier=+th.dataset.headTier, start=+th.dataset.headStart, span=+th.dataset.headSpan; const old=th.innerText.replace('✎','').trim();
    showModal('Kopfzelle bearbeiten',`<div class="field"><label>Text</label><input id="headValue" value="${esc(old)}"></div><p class="hint">Verbinden bedeutet: Diese Kopfzelle bekommt denselben Überpunkt wie die rechte Nachbarspalte. Teilen trennt die aktuell zusammengefasste Kopfzelle wieder in einzelne Kopfzellen.</p>`,[
      {text:'Speichern',action:()=>{const val=$('#headValue').value.trim()||old; for(let i=start;i<start+span;i++){t.columns[i].parts[tier]=val; t.columns[i].label=t.columns[i].parts.join(' / ');} save(); closeModal(); render();}},
      {text:'Mit rechter Zelle verbinden',secondary:true,action:()=>{const val=$('#headValue').value.trim()||old; const end=start+span; if(t.columns[end]){for(let i=start;i<=end;i++){t.columns[i].parts[tier]=val; t.columns[i].label=t.columns[i].parts.join(' / ');}} save(); closeModal(); render();}},
      {text:'Verbund teilen',secondary:true,action:()=>{for(let i=start;i<start+span;i++){t.columns[i].parts[tier]=(i===start?old:(t.columns[i].parts[tier+1]||shortHeader(t.columns[i].parts)||old)); t.columns[i].label=t.columns[i].parts.join(' / ');} save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
  };
  const baseBindTopicV16 = bindTopic;
  bindTopic = function(card,sheet,access){
    baseBindTopicV16(card,sheet,access);
    card.addEventListener('click',e=>{ const b=e.target.closest('[data-head-edit]'); if(b){ const th=b.closest('th[data-head-tier]'); if(th && canAccess(access,'edit')) editHeader((sheet.themes||[]).find(x=>(x.key||topicKey(x))===card.dataset.topic), th); }});
  };
  function userHasSeenRelease(){ return state?.lastSeenRelease === VERSION_16; }
  const baseRenderV16 = render;
  render = function(){ baseRenderV16(); decorateReleaseButton(); };
  function decorateReleaseButton(){
    const btn=$('#releaseBtn'); if(!btn) return;
    btn.classList.toggle('unseen', !userHasSeenRelease());
    btn.innerHTML=`<span class="release-symbol">📝</span><span class="release-label">Release Notes</span>${!userHasSeenRelease()?'<span class="new-dot">Neu</span>':''}`;
  }
  const baseOpenReleaseV16 = openRelease;
  openRelease = function(){ state.lastSeenRelease = VERSION_16; save(); baseOpenReleaseV16(); decorateReleaseButton(); };
  function adminRawEditor(){
    return `<h3>Admin: Daten & Parameter</h3><p class="hint">Vorschau für die spätere Supabase-Persistenz: Admin-Parameter, Standardfarben, Katalog, Rechte und Nutzerdaten werden in der stabilen Live-Struktur gespeichert und bei neuen Versionen migriert.</p><div class="field"><label>Parameter als JSON</label><textarea id="adminParamJson" style="min-height:220px">${esc(JSON.stringify(state.adminParameters||{},null,2))}</textarea></div><button id="saveAdminParams">Parameter speichern</button> <button class="secondary" id="publishLiveState">Live veröffentlichen</button>`;
  }
  const baseRenderSettingsPanelV16 = renderSettingsPanel;
  renderSettingsPanel = function(panel){
    baseRenderSettingsPanelV16(panel);
    const el=$('#settingsPanel'); if(!el) return;
    if(panel==='daten'){
      el.innerHTML += `<hr style="border:0;border-top:1px solid var(--border);margin:18px 0">${adminRawEditor()}`;
      $('#saveAdminParams')?.addEventListener('click',()=>{try{state.adminParameters=JSON.parse($('#adminParamJson').value); save(); toast('Admin-Parameter gespeichert');}catch(e){toast('JSON ist ungültig');}});
      $('#publishLiveState')?.addEventListener('click',()=>{state.publishedAt=new Date().toISOString(); save(); toast('Live veröffentlicht: Admin- und Nutzerdaten bleiben für künftige Versionen erhalten.');});
    }
  };
  const baseBindGlobalV16 = bindGlobal;
  bindGlobal = function(){
    baseBindGlobalV16();
    const rb=$('#releaseBtn'); if(rb) rb.onclick=()=>openRelease();
  };
  const baseReleaseNotesV16 = releaseNotes;
  releaseNotes = function(){return [
    {title:'V26.07.16 · vollständiger Patch',body:'Übernimmt die geforderten Funktionen auf stabiler Basis: Release-Hover mit Neu-Hinweis, Verbauliste-Kopfstruktur, Kopfzellen verbinden/teilen, korrigierte Formelberechnung, Löschen in Dateiablage und Teilekatalog, Reiter-/Ordnerlogos mit Iconpool, Teilekatalog mit Marke-/Modell-Auswahl, markierte Suchsprünge und Live-Erhaltungslogik für Admin- und Nutzerdaten.',jump:'top'},
    ...baseReleaseNotesV16()
  ];};
})();

/* ---------------- V26.07.16 Zusatz: LH-Funktionsbereiche sichtbar übernommen ---------------- */
(function(){
  const prevOpenSettingsV16b = openSettings;
  const prevRenderSettingsPanelV16b = renderSettingsPanel;
  openSettings = function(panel='profil'){
    const nav=[['profil','Profil'],['design','Design & Logo'],['farben','Farben'],['dropdowns','Dropdowns'],['builder','Reiter-Builder'],['kernblock','Kernblock-Bearbeitung'],['rechte','Reiter & Rechte'],['nutzer','Nutzerverwaltung'],['abo','Konto & Abo'],['zahlungen','Bezahlmöglichkeiten'],['hinweise','Hinweise'],['supabase','Supabase'],['daten','Sicherung/Import'],['diagnose','Diagnose']];
    showModal('Appeinstellungen',`<div class="settings-layout"><nav class="settings-nav">${nav.map(([id,txt])=>`<button class="${id===panel?'active':''}" data-set-panel="${id}">${txt}</button>`).join('')}</nav><section class="settings-panel" id="settingsPanel"></section></div>`,[{text:'Schließen',secondary:true,action:closeModal}]);
    $('#modalRoot').addEventListener('click',e=>{const b=e.target.closest('[data-set-panel]'); if(b){openSettings(b.dataset.setPanel)}});
    renderSettingsPanel(panel);
  };
  function activeEditableSheet(){ return sheetById(state.activeSheetId) || allSheetsForUser()[0]; }
  renderSettingsPanel = function(panel){
    const el=$('#settingsPanel'); if(!el) return;
    if(panel==='builder'){
      const sheets=allSheetsForUser();
      el.innerHTML=`<h3>Reiter-Builder</h3><p class="hint">Wie in der LH-Datei: Reiter als Entwurf bearbeiten, Reihenfolge/Typ/Rechte prüfen und anschließend live veröffentlichen. Diese Demo speichert lokal, die spätere Version nutzt Supabase.</p><div class="builder-list">${sheets.map(s=>`<div class="release-card"><h4>${esc(s.icon||'📄')} ${esc(s.name)}</h4><p>${s.vehicle?'Autoreiter mit Laufleistung, Dateien/Bildern und Fahrzeugfunktionen.':s.catalog?'Globaler Teilekatalog für alle Nutzer.':'Neutraler Reiter.'}</p><button class="small secondary" onclick="state.activeSheetId='${esc(s.id)}';save();closeModal();render();">Öffnen</button> <button class="small secondary" onclick="openSettings('rechte')">Rechte</button></div>`).join('')}</div><div style="margin-top:12px"><button id="builderAddVehicle">+ Autoreiter</button> <button class="secondary" id="builderAddNeutral">+ neutraler Reiter</button> <button class="secondary" id="builderPublish">Live veröffentlichen</button></div>`;
      $('#builderAddVehicle')?.addEventListener('click',()=>{closeModal(); addSheetDialog();});
      $('#builderAddNeutral')?.addEventListener('click',()=>{closeModal(); addSheetDialog();});
      $('#builderPublish')?.addEventListener('click',()=>{state.publishedAt=new Date().toISOString(); save(); toast('Reiter-Builder live veröffentlicht.');});
      return;
    }
    if(panel==='kernblock'){
      const sh=activeEditableSheet();
      el.innerHTML=`<h3>Kernblock-Bearbeitung</h3><p class="hint">Struktur der aktuell geöffneten Reiter-/Tabellenbereiche. Kopfzellen können auch direkt in der Tabelle per Doppelklick bearbeitet, verbunden oder geteilt werden.</p>${!sh?'<p>Kein Reiter aktiv.</p>':`<div class="release-card"><h4>${esc(sh.name)}</h4><table class="permission-table"><tr><th>Thema</th><th>Spalten</th><th>Aktion</th></tr>${(sh.themes||[]).map((t,ti)=>`<tr><td>${esc(t.displayName||t.name)}</td><td>${(t.columns||[]).map(c=>`<span class="pill">${esc(shortHeader(c.parts))}</span>`).join(' ')}</td><td><button class="small secondary" data-open-topic="${ti}">In Tabelle öffnen</button></td></tr>`).join('')}</table></div>`}`;
      $$('[data-open-topic]').forEach(b=>b.onclick=()=>{const t=sh.themes[+b.dataset.openTopic]; closeModal(); setTimeout(()=>document.getElementById(topicDomId(sh,t))?.scrollIntoView({behavior:'smooth',block:'center'}),80);});
      return;
    }
    if(panel==='supabase'){
      el.innerHTML=`<h3>Supabase-Vorschau</h3><p class="hint">Die echte Supabase-Datei kommt erst, wenn die Vorschau passt. Die Funktionsbereiche sind vorbereitet.</p><div class="grid2"><div class="release-card"><h4>Auth</h4><p>Login, Registrierung, Profile, Rollen, Nutzerfreigaben.</p><span class="pill ok">vorgesehen</span></div><div class="release-card"><h4>Database / RLS</h4><p>Reiter-, Themen-, Zell-, Datei-, Katalog- und Rechte-Daten mit Row-Level-Security.</p><span class="pill ok">vorgesehen</span></div><div class="release-card"><h4>Storage</h4><p>App-Logo, Reiterlogos, Ordnerbilder, Fahrzeugdateien und Teilekatalog-Dateien.</p><span class="pill ok">vorgesehen</span></div><div class="release-card"><h4>Realtime</h4><p>Gemeinsame Bearbeitung geteilter Reiter; letzte Änderung aktualisiert den gemeinsamen Stand.</p><span class="pill ok">vorgesehen</span></div></div>`;
      return;
    }
    if(panel==='zahlungen'){
      el.innerHTML=`<h3>Bezahlmöglichkeiten</h3><p class="hint">Orientiert an der LH-Beispieldatei: Testphase, Folgeabo-Auswahl, PayPal, Stripe/Karte und manuelle Freischaltung.</p><div class="plan-grid"><div class="plan"><h4>Testphase</h4><div class="price">0 €</div><p>Startzugang mit Ablaufdatum.</p><button class="secondary">Testphase prüfen</button></div><div class="plan"><h4>Folgeabo vorauswählen</h4><div class="price">Monatlich</div><p>Nutzer kann vor Ablauf wählen, welches Abo danach gelten soll.</p><select><option>Monatlich</option><option>Jährlich</option><option>Manuelle Freischaltung</option></select></div><div class="plan"><h4>Zahlarten</h4><p><button>PayPal</button> <button>Stripe / Karte</button> <button class="secondary">manuell</button></p></div></div>`;
      return;
    }
    if(panel==='diagnose'){
      const sheets=allSheetsForUser(); const cells=sheets.reduce((a,s)=>a+Object.keys(s.cells||{}).length,0);
      el.innerHTML=`<h3>Diagnose & Erhaltung</h3><table class="permission-table"><tr><th>Bereich</th><th>Status</th></tr><tr><td>Demo-Version</td><td>V26.07.16</td></tr><tr><td>Excel-Quelle</td><td>${esc(state.adminParameters?.excelSource||'Wartungsliste KFZ Details - 2(1).xlsm')}</td></tr><tr><td>LH-Referenz</td><td>${esc(state.adminParameters?.lhReference||'LH_Boden … V26.07.48')}</td></tr><tr><td>Reiter sichtbar</td><td>${sheets.length}</td></tr><tr><td>Zellen im aktuellen Zugriff</td><td>${cells}</td></tr><tr><td>Live veröffentlicht</td><td>${state.publishedAt?new Date(state.publishedAt).toLocaleString('de-DE'):'noch nicht'}</td></tr><tr><td>Daten-Erhalt</td><td>Live-Schlüssel + Versionsschlüssel aktiv; neue Versionen sollen gespeicherte Admin-/Nutzerdaten migrieren.</td></tr></table>`;
      return;
    }
    prevRenderSettingsPanelV16b(panel);
  };
})();


/* ---------------- V26.07.18: stabile Umsetzung der gesammelten Änderungen auf V26.07.16-Basis ---------------- */
(function(){
  'use strict';
  const VERSION_18 = 'V26.07.18';
  const LIVE_STORE_KEY_18 = 'kfz_app_live_state_preserved';
  const STORE_KEY_18 = 'kfz_app_' + VERSION_18 + '_state';
  const ICON_POOL_18 = ['🚗','🚙','🚘','🛻','🏎️','🧰','🛠️','🔧','⚙️','📁','🗂️','📄','🖼️','🧾','🔩','🛞','💡','🧪','🧼','🛢️','🔋','🪛','📦','⭐','✅','⚠️'];
  const deep = v => JSON.parse(JSON.stringify(v || (Array.isArray(v)?[]:{})));
  const allRawSheets18 = () => [...(state?.adminSheets||[]), ...Object.values(state?.userSheets||{}).flat()].filter(Boolean);
  function setVersionVisible18(){
    document.title = 'KFZ-APP ' + VERSION_18;
    const lp = document.querySelector('.login-panel h2'); if(lp) lp.textContent = 'KFZ-APP Vorschau · ' + VERSION_18;
    document.querySelectorAll('.version-chip').forEach((el,i)=>{ if(i===0 || /V26\.07\./.test(el.textContent)) el.textContent = VERSION_18; });
  }
  function normalizeParts18(col){
    col.parts = (col.parts && col.parts.length ? col.parts : headerParts(col.label || col.letter || 'Spalte','')).map(x=>String(x||'').trim()).filter(Boolean);
    if(!col.parts.length) col.parts = ['Spalte'];
    col.label = col.parts.join(' / ');
  }
  function byCol18(t){ return new Map((t.columns||[]).map(c=>[Number(c.col),c])); }
  function setParts18(t, colNumber, parts){ const c=byCol18(t).get(Number(colNumber)); if(c){ c.parts=parts.slice(); c.label=parts.join(' / '); } }
  function normalizeVerbauliste18(sheet){
    const t=(sheet.themes||[]).find(x=>/Verbauliste/i.test(String(x.name||x.displayName||''))); if(!t) return;
    t.displayName='Verbauliste'; (t.columns||[]).forEach(normalizeParts18);
    setParts18(t,2,['Menge','Verfügbar']); setParts18(t,3,['Menge','verbaut']);
    setParts18(t,4,['Beschreibung','Beschreibung']); setParts18(t,5,['Beschreibung','Kategorie']); setParts18(t,6,['Beschreibung','unter Kategorie']);
    setParts18(t,7,['Artikel','Hersteller']); setParts18(t,8,['Artikel','Teilenummer']);
    setParts18(t,9,['Verarbeitung','Einbaudatum']); setParts18(t,10,['Verarbeitung','Kilometerstand']);
    setParts18(t,11,['Verarbeitung','nächster Tausch','Laufleistung']); setParts18(t,12,['Verarbeitung','nächster Tausch','Datum']);
    setParts18(t,13,['Laufleistung','Tage bis']); setParts18(t,14,['Laufleistung','Anteil bis']);
  }
  function normalizeWartung18(sheet){
    const t=(sheet.themes||[]).find(x=>/Wartungsarbeiten|Verbrauchsmat/i.test(String(x.name||x.displayName||''))); if(!t) return;
    t.displayName='Wartungsarbeiten / Verbrauchsmaterial'; (t.columns||[]).forEach(normalizeParts18);
    setParts18(t,26,['Menge','Zahl']); setParts18(t,27,['Menge','Einheit']);
    setParts18(t,28,['Beschreibung','Beschreibung']); setParts18(t,29,['Beschreibung','Kategorie']); setParts18(t,30,['Beschreibung','unter Kategorie']);
    setParts18(t,31,['Artikel','Hersteller']); setParts18(t,32,['Artikel','Teilenummer']);
    setParts18(t,33,['Verarbeitung','Einbaudatum']); setParts18(t,34,['Verarbeitung','Kilometerstand']); setParts18(t,35,['Verarbeitung','nächster Tausch']);
  }
  function normalizeOEM18(sheet){
    const t=(sheet.themes||[]).find(x=>/OEM installed Parts/i.test(String(x.name||x.displayName||''))); if(!t) return;
    t.displayName='OEM installed Parts'; (t.columns||[]).forEach(normalizeParts18);
    setParts18(t,16,['Beschreibung','Position']); setParts18(t,17,['Beschreibung','Kategorie']); setParts18(t,18,['Beschreibung','unter Kategorie']);
    setParts18(t,19,['Artikel','Hersteller']); setParts18(t,20,['Artikel','Teilenummer - PN']); setParts18(t,21,['Artikel','OEM Nummer ^ SN']);
    setParts18(t,22,['Verarbeitung','Ausbaudatum']); setParts18(t,23,['Verarbeitung','ersetzt durch']);
    if(!(t.columns||[]).some(c=>(c.parts||[]).join(' ').toLowerCase().includes('hinweis'))){
      const col=24; t.columns.push({col,letter:colLetter(col),label:'Hinweise',parts:['Hinweise']}); sheet.maxCol=Math.max(sheet.maxCol||0,col);
      for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; sheet.cells=sheet.cells||{}; if(!sheet.cells[`${r}:${col}`]) sheet.cells[`${r}:${col}`]={r,c:col,a1:colLetter(col)+r,text:'',raw:'',style:{}};}
    }
  }
  function normalizeFahrzeugInfo18(sheet){
    const t=(sheet.themes||[]).find(x=>/Fahrzeuginformationen/i.test(String(x.name||x.displayName||''))); if(!t) return;
    t.displayName='Fahrzeuginformationen'; (t.columns||[]).forEach(normalizeParts18);
    t.columns=(t.columns||[]).filter(c=>{ const h=(c.parts||[]).join(' ').toLowerCase(); return !(h.includes('hersteller') || h.includes('teilenummer')); });
    setParts18(t,37,['Menge']); setParts18(t,38,['Beschreibung','Beschreibung']); setParts18(t,39,['Beschreibung','Kategorie']); setParts18(t,40,['Beschreibung','unter Kategorie']); setParts18(t,43,['Info']);
  }
  function normalizeWerkzeug18(sheet){
    if(!/Werkzeug/i.test(sheet.name||'')) return; (sheet.themes||[]).forEach(t=>{t.compactFit=true; (t.columns||[]).forEach(normalizeParts18);});
  }
  function normalizeAllHeaders18(sheet){ normalizeVerbauliste18(sheet); normalizeWartung18(sheet); normalizeOEM18(sheet); normalizeFahrzeugInfo18(sheet); normalizeWerkzeug18(sheet); }
  function ensureTree18(nodes){ (nodes||[]).forEach(n=>{ n.id=n.id||uid(n.type||'node'); n.type=n.type||'folder'; n.name=n.name||(n.type==='folder'?'Ordner':'Datei'); if(n.type==='folder') n.children=Array.isArray(n.children)?n.children:[]; if(!n.icon&&!n.thumb) n.icon=n.type==='folder'?'📁':'📄'; ensureTree18(n.children); }); }
  function globalCatalog18(){ return (state?.adminSheets||[]).find(s=>s.catalog) || allRawSheets18().find(s=>s.catalog); }
  function ensureCatalog18(){
    let cat=globalCatalog18(); if(!cat){ cat=createCatalogSheet(); state.adminSheets=state.adminSheets||[]; state.adminSheets.push(cat); }
    cat.id='global_parts_catalog'; cat.name='Teilekatalog'; cat.catalog=true; cat.global=true; cat.system=true;
    cat.permissions=cat.permissions||{admin:'manage','role:Nutzer':'view','role:Gast':'view'};
    cat.tree=Array.isArray(cat.tree)&&cat.tree.length?cat.tree:createCatalogSheet().tree; cat.rows=Array.isArray(cat.rows)?cat.rows:createCatalogSheet().rows;
    ensureTree18(cat.tree); return cat;
  }
  function ensureCatalogBrands18(){
    state.catalogBrands = Array.isArray(state.catalogBrands)&&state.catalogBrands.length ? state.catalogBrands : [
      {id:'audi',name:'Audi',logo:'🚗',models:[{id:'a3_8p',name:'A3 8P',logo:'🚙',motors:[{id:'2_0_tdi_cr',name:'2.0 TDI CR',logo:'⚙️'},{id:'allgemein',name:'Allgemein',logo:'🧩'}]}]},
      {id:'vw',name:'Volkswagen',logo:'🚘',models:[{id:'allgemein',name:'Allgemein',logo:'🧩',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]}
    ];
    for(const b of state.catalogBrands){ b.models=Array.isArray(b.models)&&b.models.length?b.models:[{id:'allgemein',name:'Allgemein',logo:'🧩',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]; for(const m of b.models){m.motors=Array.isArray(m.motors)&&m.motors.length?m.motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}];} }
    state.catalogSelection=state.catalogSelection||{};
    const b=state.catalogBrands.find(x=>x.id===state.catalogSelection.brand)||state.catalogBrands[0]; state.catalogSelection.brand=b?.id||'';
    const m=b?.models.find(x=>x.id===state.catalogSelection.model)||b?.models[0]; state.catalogSelection.model=m?.id||'';
    const mot=m?.motors.find(x=>x.id===state.catalogSelection.motor)||m?.motors[0]; state.catalogSelection.motor=mot?.id||'allgemein';
    if(!('vin' in state.catalogSelection)) state.catalogSelection.vin='';
  }
  function selectedBrand18(){ ensureCatalogBrands18(); return state.catalogBrands.find(b=>b.id===state.catalogSelection.brand)||state.catalogBrands[0]; }
  function selectedModel18(){ const b=selectedBrand18(); return b?.models.find(m=>m.id===state.catalogSelection.model)||b?.models[0]; }
  function selectedMotor18(){ const m=selectedModel18(); return m?.motors.find(x=>x.id===state.catalogSelection.motor)||m?.motors[0]; }
  function catalogKey18(){ const s=state.catalogSelection||{}; return [s.brand||'allgemein',s.model||'allgemein',s.motor||'allgemein',String(s.vin||'').trim()||'ohne_vin'].join('__'); }
  function currentCatalogBundle18(){
    const cat=ensureCatalog18(); ensureCatalogBrands18(); cat.vehicleCatalogs=cat.vehicleCatalogs||{}; const key=catalogKey18();
    if(!cat.vehicleCatalogs[key]) cat.vehicleCatalogs[key]={tree:deep(cat.tree),rows:deep(cat.rows)};
    ensureTree18(cat.vehicleCatalogs[key].tree); cat.vehicleCatalogs[key].rows=Array.isArray(cat.vehicleCatalogs[key].rows)?cat.vehicleCatalogs[key].rows:[];
    return cat.vehicleCatalogs[key];
  }
  function migrate18(s){
    state=s; s.version=VERSION_18; s.adminParameters=s.adminParameters||{}; s.adminParameters.excelSource='Wartungsliste KFZ Details - 2(1).xlsm'; s.adminParameters.lhReference='LH_Boden_MTV_VTV_Online_Browser_App_V26.07.48_Aktualisiert_Kernblock_Bearbeitung_Netlify';
    for(const sh of allRawSheets18()){ normalizeAllHeaders18(sh); if(sh.vehicle){sh.files=Array.isArray(sh.files)?sh.files:defaultFileTree(); ensureTree18(sh.files);} if(sh.catalog) ensureTree18(sh.tree); }
    ensureCatalog18(); ensureCatalogBrands18(); currentCatalogBundle18(); state=null; return s;
  }
  const baseInit18=initState;
  initState=function(){ let s; try{s=JSON.parse(localStorage.getItem(STORE_KEY_18)||'null');}catch(e){s=null;} if(!(s&&s.users)) s=baseInit18(); return migrate18(s); };
  save=function(){ if(!state) return; state.version=VERSION_18; state.savedAt=new Date().toISOString(); try{localStorage.setItem(LIVE_STORE_KEY_18,JSON.stringify(state)); localStorage.setItem(STORE_KEY_18,JSON.stringify(state));}catch(e){console.warn(e);} updateSaveInfo?.(); };
  function deleteRow18(sheet,t,row){ if(!confirm('Diese Zeile löschen?')) return; const r=Number(row); t.rowGroups=(t.rowGroups||[]).filter(g=>(Array.isArray(g)?Number(g[0]):Number(g))!==r); for(const c of t.columns||[]) delete sheet.cells?.[`${r}:${c.col}`]; save(); render(); }
  function deleteColumn18(sheet,t,idx){ const i=Number(idx); if((t.columns||[]).length<=1){toast('Letzte Spalte kann nicht gelöscht werden.'); return;} const col=t.columns[i]; if(!col) return; if(!confirm(`Spalte „${shortHeader(col.parts)}“ löschen?`)) return; t.columns.splice(i,1); for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; delete sheet.cells?.[`${r}:${col.col}`];} save(); render(); }
  renderRows=function(sheet,t,canEdit){ let rows=(t.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g).filter(Boolean); if(!rows.length){for(let r=t.dataStart||5;r<=t.endRow;r++)rows.push(r)} let html=''; rows.forEach(r=>{ html+=`<tr class="insert-row"><td colspan="${(t.columns||[]).length+1}">${canEdit?`<button class="row-insert-plus secondary" title="Zeile hier einfügen" data-insert-before="${r}">+</button>`:''}</td></tr>`; html+=`<tr draggable="${canEdit}" data-row="${r}"><td class="row-tools"><button class="row-handle" ${canEdit?'':'disabled'} draggable="${canEdit}" data-row-handle="${r}" title="Zeile verschieben">↕</button>${canEdit?`<button class="row-delete small danger" data-delete-row="${r}" title="Zeile löschen">×</button>`:''}</td>`; for(const col of (t.columns||[])) html+=renderCell(sheet,t,r,col,canEdit); html+='</tr>'; }); html+=`<tr class="insert-row"><td colspan="${(t.columns||[]).length+1}">${canEdit?'<button class="row-insert-plus secondary" title="Zeile unten einfügen" data-insert-before="end">+</button>':''}</td></tr>`; return html; };
  renderHeader=function(sheet,t,canEdit){ normalizeAllHeaders18(sheet); const cols=t.columns||[]; cols.forEach(normalizeParts18); const maxDepth=Math.max(1,...cols.map(c=>(c.parts||[]).length)); let html='<thead class="modern-head">'; for(let tier=0;tier<maxDepth;tier++){ html+='<tr>'; if(tier===0) html+=`<th class="row-tools" rowspan="${maxDepth}">↕</th>`; let i=0; while(i<cols.length){ const parts=cols[i].parts||['Spalte']; if(tier>=parts.length){i++; continue;} const label=parts[tier]||''; const prefix=parts.slice(0,tier).join('|'); let span=1; while(i+span<cols.length){const p2=cols[i+span].parts||[]; if(tier>=p2.length)break; if((p2[tier]||'')!==label)break; if(p2.slice(0,tier).join('|')!==prefix)break; span++;} const rowspan=(tier===parts.length-1)?(maxDepth-tier):1; const leaf=(tier===parts.length-1); html+=`<th class="editable-head" draggable="${canEdit}" data-head-tier="${tier}" data-head-start="${i}" data-head-span="${span}" data-col-index="${i}" colspan="${span}" rowspan="${rowspan}" title="Doppelklick: Text bearbeiten / verbinden / teilen"><span>${esc(label)}</span>${canEdit?'<button class="head-tool" data-head-edit>✎</button>':''}${canEdit&&leaf?`<button class="col-delete" data-delete-col="${i}" title="Spalte löschen">×</button>`:''}</th>`; i+=span; } html+='</tr>'; } return html+'</thead>'; };
  renderTopic=function(sheet,t,access){ const canEdit=canAccess(access,'edit'); normalizeAllHeaders18(sheet); const fit=t.compactFit?' compact-fit':''; return `<section class="topic-card${fit}" id="${topicDomId(sheet,t)}" data-topic="${esc(t.key||topicKey(t))}" style="width:${t.uiWidth||'auto'};height:${t.uiHeight||'auto'}"><div class="topic-head" draggable="${canEdit}" data-drag-topic="${esc(t.key||topicKey(t))}"><h3>${esc(t.displayName||t.name)}</h3><div class="topic-actions"><select class="sort-select" data-sort-select>${(t.columns||[]).map((c,i)=>`<option value="${i}">${esc(shortHeader(c.parts))}</option>`).join('')}</select><button class="small secondary" data-sort="asc">↑</button><button class="small secondary" data-sort="desc">↓</button>${canEdit?'<button class="small secondary" data-add-col>+ Spalte</button>':''}<button class="small secondary" data-collapse>Ein-/ausklappen</button></div></div><div class="topic-body"><div class="table-wrap"><table class="kfz-table">${renderHeader(sheet,t,canEdit)}<tbody>${renderRows(sheet,t,canEdit)}</tbody></table></div></div><div class="topic-foot"><button class="secondary" data-add-row-end>+ Zeile unten</button><span class="hint">Doppelklick bearbeitet · + fügt ein · × löscht Zeilen/Spalten</span></div></section>`; };
  bindTopic=function(card,sheet,access){ const canEdit=canAccess(access,'edit'); const key=card.dataset.topic; const t=(sheet.themes||[]).find(x=>(x.key||topicKey(x))===key); if(!t)return; card.addEventListener('click',e=>{ const delR=e.target.closest('[data-delete-row]'); if(delR){e.stopPropagation(); return deleteRow18(sheet,t,delR.dataset.deleteRow);} const delC=e.target.closest('[data-delete-col]'); if(delC){e.stopPropagation(); return deleteColumn18(sheet,t,delC.dataset.deleteCol);} const headBtn=e.target.closest('[data-head-edit]'); if(headBtn){const th=headBtn.closest('th[data-head-tier]'); if(th&&canEdit) return editHeader(t,th);} if(e.target.matches('[data-collapse]'))card.classList.toggle('collapsed'); if(e.target.matches('[data-sort]'))sortTheme(sheet,t,e.target.dataset.sort); if(e.target.matches('[data-add-row-end]'))addRow(sheet,t,'end'); if(e.target.matches('[data-insert-before]'))addRow(sheet,t,e.target.dataset.insertBefore); if(e.target.matches('[data-add-col]'))addColumn(sheet,t); if(e.target.matches('[data-edit-cell]'))openCellEditor(sheet,t,e.target.closest('td')); }); card.addEventListener('dblclick',e=>{const th=e.target.closest('th[data-head-tier]'); if(th&&canEdit)return editHeader(t,th); const td=e.target.closest('td[data-r]'); if(td&&canEdit&&!e.target.closest('a'))return editCellInline(sheet,t,td);}); card.addEventListener('dragstart',e=>{const row=e.target.closest('[data-row]'); const th=e.target.closest('th[data-col-index]'); const head=e.target.closest('[data-drag-topic]'); if(row&&canEdit){dragged={type:'row',r:+row.dataset.row,topic:t.key,sheet:sheet.id};} else if(th&&canEdit){dragged={type:'col',idx:+th.dataset.colIndex,topic:t.key,sheet:sheet.id};} else if(head&&canEdit){dragged={type:'topic',key:t.key,sheet:sheet.id};}}); card.addEventListener('dragover',e=>{if(dragged)e.preventDefault();}); card.addEventListener('drop',e=>{if(!dragged)return; if(dragged.type==='row'){const tr=e.target.closest('tr[data-row]'); if(tr)moveRow(sheet,t,dragged.r,+tr.dataset.row);} if(dragged.type==='col'){const th=e.target.closest('th[data-col-index]'); if(th)moveCol(t,dragged.idx,+th.dataset.colIndex);} dragged=null; render(); save();}); };
  function highlight18(el,cls='highlight-cell'){ if(!el)return; document.querySelectorAll('.highlight-cell,.highlight-row').forEach(x=>x.classList.remove('highlight-cell','highlight-row')); el.classList.add(cls); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}); setTimeout(()=>el.classList.remove(cls),7000); }
  partSearch=function(sheet){ const q=$('#partSearch')?.value.trim().toLowerCase(); const out=$('#partResults'); if(!q||!out){if(out)out.innerHTML='';return;} const hits=[]; for(const t of sheet.themes||[]) for(const col of t.columns||[]){const h=(col.parts||[]).join(' ').toLowerCase(); if(/teil|part|pn|oem|sn|nummer/.test(h)){for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; const v=String(sheet.cells?.[`${r}:${col.col}`]?.text||''); if(v.toLowerCase().includes(q)) hits.push({theme:t,row:r,col,value:v});}}} out.innerHTML=hits.length?hits.slice(0,80).map((h,i)=>`<div class="search-hit"><div><strong>${esc(h.value)}</strong><div class="hint">${esc(h.theme.displayName||h.theme.name)} · ${esc(shortHeader(h.col.parts))}</div></div><button class="small secondary" data-hit="${i}">Springen & markieren</button></div>`).join(''):'<div class="search-hit">Keine Treffer.</div>'; out.querySelectorAll('[data-hit]').forEach(btn=>btn.onclick=()=>{const h=hits[+btn.dataset.hit]; render(); setTimeout(()=>highlight18(document.querySelector(`#${CSS.escape(topicDomId(sheet,h.theme))} td[data-r="${h.row}"][data-c="${h.col.col}"]`)),80);}); };
  function treePath18(nodes,id,path=[]){ for(const n of nodes||[]){const p=path.concat(n.name); if(n.id===id)return p; const c=treePath18(n.children,id,p); if(c)return c;} return null; }
  renderFileSection=function(sheet,access){ sheet.files=Array.isArray(sheet.files)?sheet.files:defaultFileTree(); ensureTree18(sheet.files); const path=treePath18(sheet.files,state.fileSelected)||[]; return `<section class="file-section" id="vehicleFiles"><div class="section-head"><h3>Dateien & Bilder</h3><div>${canAccess(access,'edit')?'<button class="small secondary" data-file-add-root>+ Ordner</button> <button class="small secondary" data-file-upload-root>+ Datei/Bild</button>':''}</div></div><div class="path-line" id="filePath">Pfad: /${esc(path.join(' / '))}</div><div class="tree vertical-tree" id="fileTree">${renderTree(sheet.files,'file',access)}</div></section>`; };
  renderMain=function(){ const sheet=ensureActive(); const main=$('#main'); if(!sheet){main.innerHTML='<div class="login-panel"><h2>Keine Reiter verfügbar</h2></div>';return;} if(sheet.catalog){renderCatalog(sheet);return;} normalizeAllHeaders18(sheet); const access=accessFor(sheet); main.innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>${esc(sheet.name)}</h2><div class="sheet-sub"><span class="pill ${sheet.vehicle?'ok':''}">${sheet.vehicle?'Autoreiter':'Neutraler Reiter'}</span><span class="pill">Zugriff: ${accessLabel(access)}</span>${sheet.ownerId!=='system'?`<span class="pill warn">Besitzer: ${esc(userName(sheet.ownerId))}</span>`:''}</div></div><div class="sheet-tools"><button class="secondary" id="editTabIconBtn">Reiterlogo/Icon</button><button class="secondary" id="shareTabBtn">Teilen/Rechte</button><button class="secondary" id="archiveTabBtn">Archivieren</button><button class="danger" id="deleteTabBtn">Löschen</button></div></div>${sheet.vehicle?vehicleHtml(sheet):''}<div class="global-search"><strong>🔎 Teilenummer / Partnummer</strong><input id="partSearch" placeholder="PN, OEM, SN, Teilenummer suchen …"><button id="partSearchBtn">Suchen</button></div><div id="partResults" class="search-results"></div><div class="topic-nav">${(sheet.themes||[]).map(t=>`<a href="#${esc(topicDomId(sheet,t))}">${esc(t.displayName||t.name)}</a>`).join('')}</div><div class="topic-stack">${(sheet.themes||[]).map(t=>renderTopic(sheet,t,access)).join('')}</div>${sheet.vehicle?renderFileSection(sheet,access):''}`; bindSheetEvents(sheet,access); $('#editTabIconBtn')?.addEventListener('click',()=>editSheetIcon(sheet)); };
  function removeNode18(nodes,id){ if(!Array.isArray(nodes))return null; const idx=nodes.findIndex(n=>n.id===id); if(idx>=0)return nodes.splice(idx,1)[0]; for(const n of nodes){const r=removeNode18(n.children,id); if(r)return r;} return null; }
  function addCatalogRowForNode18(node){ const bundle=currentCatalogBundle18(); const row={id:node.rowId||uid('catrow'),image:node.thumb||'',name:node.part||node.name,part:node.part||'',category:'',note:node.fileName||node.name}; node.rowId=row.id; bundle.rows.push(row); }
  function pickFile18(parent,kind){ const input=document.createElement('input'); input.type='file'; input.accept='image/*,.pdf,.txt,.doc,.docx,.xls,.xlsx,.xlsm'; input.onchange=()=>{const f=input.files&&input.files[0]; if(!f)return; parent.children=parent.children||[]; parent.open=true; const node={id:uid('file'),type:'file',name:f.name,fileName:f.name,fileType:f.type||'',fileSize:f.size||0,icon:f.type&&f.type.startsWith('image/')?'🖼️':'📄',thumb:'',part:'',rowId:kind==='catalog'?uid('catrow'):undefined}; const finish=()=>{parent.children.push(node); if(kind==='catalog')addCatalogRowForNode18(node); save(); render();}; if(f.type&&f.type.startsWith('image/')){const rd=new FileReader(); rd.onload=()=>{node.thumb=rd.result; finish();}; rd.readAsDataURL(f);} else finish();}; input.click(); }
  addTreeChild=function(parent,type,kind){ if(type==='file') return pickFile18(parent,kind); parent.children=parent.children||[]; parent.open=true; parent.children.push({id:uid('folder'),type:'folder',name:'Neuer Ordner',open:true,children:[],icon:'📁'}); save(); render(); };
  uploadVehicleFile=function(sheet){ sheet.files=Array.isArray(sheet.files)?sheet.files:defaultFileTree(); const root={children:sheet.files,open:true}; pickFile18(root,'file'); };
  renderTree=function(nodes,kind,access,depth=0){ return (nodes||[]).map(n=>{ const canEdit=kind==='catalog'?currentUser()?.role==='Admin':canAccess(access,'edit'); const icon=n.thumb?`<img class="thumb" src="${esc(n.thumb)}" alt="">`:`<span class="thumb">${esc(n.icon || (n.type==='folder'?'📁':'📄'))}</span>`; return `<div class="tree-node" data-node="${esc(n.id)}" data-kind="${kind}"><div class="tree-row" style="margin-left:${depth*14}px" data-tree-click><span class="tree-toggle">${n.type==='folder'?(n.open?'▾':'▸'):'•'}</span>${icon}<strong>${esc(n.name)}</strong>${n.part?`<span class="pill">${esc(n.part)}</span>`:''}<span style="flex:1"></span>${canEdit?`<button class="small secondary" data-tree-edit>✎</button>${n.type==='folder'?'<button class="small secondary" data-tree-add-folder>+ Ordner</button><button class="small secondary" data-tree-add-file>+ Datei</button>':''}<button class="small danger" data-tree-delete title="Löschen">🗑</button>`:''}</div>${n.type==='folder'&&n.open?`<div class="tree-children">${renderTree(n.children||[],kind,access,depth+1)}</div>`:''}</div>`; }).join(''); };
  bindTree=function(container,root,kind,access){ if(!container)return; container.onclick=e=>{ const nodeEl=e.target.closest('[data-node]'); if(!nodeEl)return; const node=findNode(root,nodeEl.dataset.node); if(!node)return; if(e.target.closest('[data-tree-delete]')){e.stopPropagation(); if(confirm(`${node.type==='folder'?'Ordner':'Datei'} „${node.name}“ löschen?`)){const removed=removeNode18(root,node.id); if(kind==='catalog'&&removed?.rowId){const bundle=currentCatalogBundle18(); const i=(bundle.rows||[]).findIndex(r=>r.id===removed.rowId); if(i>=0)bundle.rows.splice(i,1);} save(); render();} return;} if(e.target.closest('[data-tree-edit]')){e.stopPropagation(); return editTreeNode(node,kind);} if(e.target.closest('[data-tree-add-folder]')){e.stopPropagation(); return addTreeChild(node,'folder',kind);} if(e.target.closest('[data-tree-add-file]')){e.stopPropagation(); return addTreeChild(node,'file',kind);} if(e.target.closest('[data-tree-click]')){ if(node.type==='folder'){node.open=!node.open; if(kind==='file')state.fileSelected=node.id; save(); render();} else { if(kind==='file'){state.fileSelected=node.id; save(); render();} if(kind==='catalog'){state.catalogSelected=node.id; save(); render(); setTimeout(()=>highlight18(document.querySelector(`[data-catalog-row="${node.rowId}"]`),'highlight-row'),80);} } } }; };
  const baseEditTreeNode18=editTreeNode;
  editTreeNode=function(node,kind){ const iconButtons=ICON_POOL_18.map(i=>`<button class="small secondary" data-pick-icon="${esc(i)}">${esc(i)}</button>`).join(''); showModal((node.type==='folder'?'Ordner':'Datei')+' bearbeiten',`<div class="grid2"><div class="field"><label>Name</label><input id="tnName" value="${esc(node.name)}"></div><div class="field"><label>Teilenummer / Bildtext</label><input id="tnPart" value="${esc(node.part||'')}"></div><div class="field"><label>Icon/Text</label><input id="tnIcon" value="${esc(node.icon||'')}"></div><div class="field"><label>Vorschau/Bild</label><input id="tnFile" type="file" accept="image/*"></div><div class="field full"><label>Icon-Pool</label><div class="icon-pool">${iconButtons}</div></div></div>`,[{text:'Speichern',action:()=>{node.name=$('#tnName').value.trim()||node.name; node.part=$('#tnPart').value.trim(); node.icon=$('#tnIcon').value.trim()||node.icon; const f=$('#tnFile').files[0]; const finish=()=>{ if(kind==='catalog'&&node.rowId){const r=(currentCatalogBundle18().rows||[]).find(x=>x.id===node.rowId); if(r){r.name=node.part||node.name; r.part=node.part||r.part; r.image=node.thumb||r.image;}} save(); closeModal(); render(); }; if(f){const rd=new FileReader(); rd.onload=()=>{node.thumb=rd.result; finish();}; rd.readAsDataURL(f);} else finish();}},{text:'Abbrechen',secondary:true,action:closeModal}]); $$('#modalRoot [data-pick-icon]').forEach(b=>b.onclick=()=>{$('#tnIcon').value=b.dataset.pickIcon;}); };
  function catalogControls18(isAdmin){ ensureCatalogBrands18(); const b=selectedBrand18(), m=selectedModel18(), motors=m?.motors||[]; return `<section class="catalog-chooser"><div class="catalog-chooser-grid"><div class="field"><label>Marke</label><select id="catBrand">${state.catalogBrands.map(x=>`<option value="${esc(x.id)}" ${x.id===b?.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>Modell</label><select id="catModel">${(b?.models||[]).map(x=>`<option value="${esc(x.id)}" ${x.id===m?.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>Motor</label><select id="catMotor">${motors.map(x=>`<option value="${esc(x.id)}" ${x.id===selectedMotor18()?.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>VIN optional</label><input id="catVin" value="${esc(state.catalogSelection?.vin||'')}" placeholder="VIN optional"></div></div>${isAdmin?'<div class="admin-line"><button class="small secondary" id="catAddBrand">+ Marke</button><button class="small secondary" id="catAddModel">+ Modell</button><button class="small secondary" id="catAddMotor">+ Motor</button><button class="small secondary" id="catEditVehicle">Logos/Bezeichnungen</button></div>':''}</section>`; }
  function editVehicleSelector18(){ const b=selectedBrand18(), m=selectedModel18(), mot=selectedMotor18(); showModal('Marke / Modell / Motor bearbeiten',`<div class="grid2"><div class="field"><label>Marke</label><input id="evBrand" value="${esc(b?.name||'')}"></div><div class="field"><label>Markenlogo/Icon</label><input id="evBrandLogo" value="${esc(b?.logo||'')}"></div><div class="field"><label>Modell</label><input id="evModel" value="${esc(m?.name||'')}"></div><div class="field"><label>Modelllogo/Icon</label><input id="evModelLogo" value="${esc(m?.logo||'')}"></div><div class="field"><label>Motor</label><input id="evMotor" value="${esc(mot?.name||'')}"></div><div class="field"><label>Motorlogo/Icon</label><input id="evMotorLogo" value="${esc(mot?.logo||'')}"></div></div>`,[{text:'Speichern',action:()=>{if(b){b.name=$('#evBrand').value.trim()||b.name;b.logo=$('#evBrandLogo').value.trim()||b.logo;} if(m){m.name=$('#evModel').value.trim()||m.name;m.logo=$('#evModelLogo').value.trim()||m.logo;} if(mot){mot.name=$('#evMotor').value.trim()||mot.name;mot.logo=$('#evMotorLogo').value.trim()||mot.logo;} save(); closeModal(); render();}},{text:'Abbrechen',secondary:true,action:closeModal}]); }
  renderCatalog=function(sheet){ const cat=ensureCatalog18(); const isAdmin=currentUser()?.role==='Admin'; const bundle=currentCatalogBundle18(); const selectedNode=findNode(bundle.tree,state.catalogSelected); $('#main').innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>Teilekatalog</h2><div class="sheet-sub"><span class="pill ok">Globaler Teilekatalog</span><span class="pill">${isAdmin?'Admin bearbeitet':'Nur ansehen'}</span></div></div><div class="sheet-tools">${isAdmin?'<button id="catAddRoot" class="secondary">+ Hauptordner</button> <button id="catPublish" class="secondary">Katalog veröffentlichen</button>':''}</div></div>${catalogControls18(isAdmin)}<section class="catalog-section"><div class="section-head"><h3>Ordnerstruktur</h3><span class="hint">Ordner öffnen direkt untereinander. + Datei wählt eine Datei vom Desktop.</span></div><div class="catalog-grid"><div><div class="tree vertical-tree" id="catalogTree">${renderTree(bundle.tree,'catalog','manage')}</div><div class="file-picker-hint">Die Ordner/Dateien gehören zur gewählten Marke/Modell/Motor/VIN-Kombination.</div></div><div><h3>Dateitabelle</h3><table class="catalog-table"><thead><tr><th>Bild</th><th>Bildtext / Datei</th><th>Teilenummer</th><th>Kategorie</th><th>Notiz</th></tr></thead><tbody>${(bundle.rows||[]).map(r=>`<tr data-catalog-row="${esc(r.id)}" class="${selectedNode?.rowId===r.id?'highlight-row':''}"><td>${r.image?`<img class="catalog-preview" src="${esc(r.image)}">`:'<span class="thumb">🧩</span>'}</td><td><span class="jump-link" data-jump-row="${esc(r.id)}">${esc(r.name)}</span></td><td>${esc(r.part)}</td><td>${esc(r.category)}</td><td>${esc(r.note)}</td></tr>`).join('')}</tbody></table></div></div></section>`; $('#catBrand')?.addEventListener('change',e=>{state.catalogSelection.brand=e.target.value; const b=selectedBrand18(); state.catalogSelection.model=b?.models?.[0]?.id||''; const m=selectedModel18(); state.catalogSelection.motor=m?.motors?.[0]?.id||''; state.catalogSelected=''; save(); render();}); $('#catModel')?.addEventListener('change',e=>{state.catalogSelection.model=e.target.value; const m=selectedModel18(); state.catalogSelection.motor=m?.motors?.[0]?.id||''; state.catalogSelected=''; save(); render();}); $('#catMotor')?.addEventListener('change',e=>{state.catalogSelection.motor=e.target.value; state.catalogSelected=''; save(); render();}); $('#catVin')?.addEventListener('blur',e=>{state.catalogSelection.vin=e.target.value.trim(); state.catalogSelected=''; save(); render();}); $('#catAddRoot')?.addEventListener('click',()=>{bundle.tree.push({id:uid('folder'),type:'folder',name:'Neuer Hauptordner',open:true,children:[],icon:'📁'}); save(); render();}); $('#catPublish')?.addEventListener('click',()=>{state.publishedAt=new Date().toISOString(); save(); toast('Teilekatalog veröffentlicht.');}); $('#catAddBrand')?.addEventListener('click',()=>{const name=prompt('Neue Marke?'); if(!name)return; const id=slug(name)+'_'+uid('b'); state.catalogBrands.push({id,name,logo:'🚗',models:[{id:'allgemein_'+uid('m'),name:'Allgemein',logo:'🧩',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]}); state.catalogSelection={brand:id,model:state.catalogBrands[state.catalogBrands.length-1].models[0].id,motor:'allgemein',vin:''}; save(); render();}); $('#catAddModel')?.addEventListener('click',()=>{const b=selectedBrand18(); const name=prompt('Neues Modell?'); if(!b||!name)return; const id=slug(name)+'_'+uid('m'); b.models.push({id,name,logo:'🚙',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}); state.catalogSelection.model=id; state.catalogSelection.motor='allgemein'; save(); render();}); $('#catAddMotor')?.addEventListener('click',()=>{const m=selectedModel18(); const name=prompt('Neuer Motor?'); if(!m||!name)return; const id=slug(name)+'_'+uid('mot'); m.motors.push({id,name,logo:'⚙️'}); state.catalogSelection.motor=id; save(); render();}); $('#catEditVehicle')?.addEventListener('click',editVehicleSelector18); $$('#main [data-jump-row]').forEach(el=>el.onclick=()=>highlight18(document.querySelector(`[data-catalog-row="${el.dataset.jumpRow}"]`),'highlight-row')); bindTree($('#catalogTree'),bundle.tree,'catalog','manage'); };
  const baseBindSheet18=bindSheetEvents;
  bindSheetEvents=function(sheet,access){ baseBindSheet18(sheet,access); if(sheet.vehicle && $('#fileTree')) bindTree($('#fileTree'),sheet.files,'file',access); };
  function topActions18(){ const u=currentUser?.(); const box=document.querySelector('.top-actions'); if(!box||!u)return; box.classList.add('lh-like'); box.innerHTML=`<span class="version-chip">${VERSION_18}</span><span class="user-chip auth-user-box" id="userChip">👤 ${esc(u.name)} <span class="hint" style="color:#fff9">${esc(u.role)}</span></span><span class="version-chip" id="saveInfo">${state.savedAt?'Gespeichert · '+new Date(state.savedAt).toLocaleTimeString('de-DE'):'Vorschau nicht gespeichert'}</span><button id="topManageBtn" class="secondary">Verwaltung</button><button id="saveBtn" class="secondary">Cloud speichern</button><button id="darkToggle" class="secondary">🌙</button><button id="logoutBtn" class="secondary">Abmelden</button>`; $('#topManageBtn')?.addEventListener('click',()=>openSettings('profil')); $('#saveBtn')?.addEventListener('click',()=>{save();toast('Manuell gespeichert')}); $('#darkToggle')?.addEventListener('click',()=>{state.settings.dark=!state.settings.dark;save();render();}); $('#logoutBtn')?.addEventListener('click',logout); }
  const baseRender18=render;
  render=function(){ baseRender18(); setVersionVisible18(); topActions18(); };
  const baseReleaseNotes18=releaseNotes;
  releaseNotes=function(){ return [{title:'V26.07.18 · stabile Korrektur nach V26.07.17',body:'Repariert die lauffähige Demo auf Basis V26.07.16 und setzt die zuletzt gewünschten Änderungen sauber um: echte Dateiauswahl im Teilekatalog, Fahrzeug-/Motor-/VIN-bezogene Katalogbäume, Dateien & Bilder unten im Autoreiter, korrigierte Wartungs-/OEM-/Fahrzeuginformations-Tabellen, Löschfunktionen für Zeilen und Spalten, kompakte + Einfügepunkte und LH-ähnliche Kopfzeile.',jump:'catalog'}, ...baseReleaseNotes18()]; };
})();


/* ---------------- V26.07.19: gezielte Umsetzung der letzten Verbesserungsliste ---------------- */
(function(){
  const VERSION_19='V26.07.19';
  const STORE_KEY_19='kfz_app_'+VERSION_19+'_state';
  const LIVE_KEY='kfz_app_live_state_preserved';
  const ICONS_19=['🚗','🚙','🛠️','🧰','🔧','⚙️','🧩','📁','📄','🖼️','📦','🔌','💡','🛞','🧪','🧼','🛢️','🔋','🚦','🧾','⭐'];
  const baseInit19=initState;
  initState=function(){
    let s=null;
    try{s=JSON.parse(localStorage.getItem(STORE_KEY_19)||localStorage.getItem(LIVE_KEY)||localStorage.getItem('kfz_app_V26.07.18_state')||'null');}catch(e){s=null;}
    if(!(s&&s.users)) s=baseInit19();
    s.version=VERSION_19;
    migrate19(s);
    return s;
  };
  const baseSave19=save;
  save=function(){
    if(!state) return;
    state.version=VERSION_19;
    state.savedAt=new Date().toISOString();
    try{localStorage.setItem(STORE_KEY_19,JSON.stringify(state));localStorage.setItem(LIVE_KEY,JSON.stringify(state));}catch(e){console.warn(e);try{baseSave19();}catch(_){} }
    updateSaveInfo?.();
  };
  function migrate19(s){
    s.archivedSheets=s.archivedSheets||{};
    for(const sh of [...(s.adminSheets||[]), ...Object.values(s.userSheets||{}).flat()]){
      if(/Werkzeug/i.test(sh.name||'')) (sh.themes||[]).forEach(t=>{t.compactFit=true;t.fitMode='werkzeug';t.uiWidth='100%';});
      if(sh.logo && !sh.icon) sh.icon='';
      if(sh.archived){ const owner=sh.ownerId||'system'; s.archivedSheets[owner]=s.archivedSheets[owner]||[]; if(!s.archivedSheets[owner].some(x=>x.id===sh.id)) s.archivedSheets[owner].push(sh); }
      if(sh.vehicle && Array.isArray(sh.files)) ensureFileMeta19(sh.files);
    }
    const cat=(s.adminSheets||[]).find(x=>x.catalog);
    if(cat){ ensureFileMeta19(cat.tree||[]); Object.values(cat.vehicleCatalogs||{}).forEach(b=>ensureFileMeta19(b.tree||[])); }
  }
  function ensureFileMeta19(nodes){(nodes||[]).forEach(n=>{n.id=n.id||uid(n.type||'node'); if(n.type==='file'){n.icon=n.icon||'📄';n.name=n.name||n.fileName||'Datei';} if(n.type==='folder'){n.icon=n.icon||'📁';n.children=Array.isArray(n.children)?n.children:[]; ensureFileMeta19(n.children);}})}
  function allRawSheets19(){return [...(state.adminSheets||[]), ...Object.values(state.userSheets||{}).flat()];}
  function catalogSheet19(){return (state.adminSheets||[]).find(s=>s.catalog) || allRawSheets19().find(s=>s.catalog);}
  function catalogBundleForRoot19(root){const cat=catalogSheet19(); if(!cat) return null; if(cat.tree===root) return cat; for(const b of Object.values(cat.vehicleCatalogs||{})){ if(b.tree===root) return b; } return null;}
  function isAdmin19(){return currentUser()?.role==='Admin';}
  function setTitle19(){document.title='KFZ-APP '+VERSION_19; const h=document.querySelector('#loginView h2'); if(h) h.textContent='KFZ-APP Vorschau · '+VERSION_19;}

  // Reiterlogos tatsächlich anzeigen.
  renderTabs=function(){
    const tabs=$('#tabs'); const sheets=allSheetsForUser();
    tabs.innerHTML=sheets.map(s=>{
      const icon=s.logo?`<img class="tab-icon-img" src="${esc(s.logo)}" alt="">`:`<span class="tab-logo">${esc(s.icon || (s.vehicle?'🚗':s.catalog?'🧩':'📄'))}</span>`;
      return `<button class="tab ${s.id===state.activeSheetId?'active':''}" draggable="true" data-tab="${esc(s.id)}" style="--tabW:${s.tabWidth||130}px">${icon}<span>${esc(s.name)}</span><span class="type">${s.vehicle?'Auto':s.catalog?'global':'Reiter'}</span><span class="tab-resize" data-resize-tab="${esc(s.id)}"></span></button>`;
    }).join('')+`<button class="tab secondary" id="addTabBtn">+ Reiter</button>`;
    tabs.onclick=e=>{const b=e.target.closest('[data-tab]'); if(b){state.activeSheetId=b.dataset.tab; save(); render();} if(e.target.closest('#addTabBtn')) addSheetDialog();};
    tabs.ondragstart=e=>{const b=e.target.closest('[data-tab]'); if(b){dragged={type:'tab',id:b.dataset.tab};}};
    tabs.ondragover=e=>{if(dragged?.type==='tab'){e.preventDefault(); const b=e.target.closest('[data-tab]'); if(b)b.classList.add('drag-over')}};
    tabs.ondragleave=e=>e.target.closest('[data-tab]')?.classList.remove('drag-over');
    tabs.ondrop=e=>{const b=e.target.closest('[data-tab]'); $$('.tab').forEach(x=>x.classList.remove('drag-over')); if(b&&dragged?.type==='tab')moveTab(dragged.id,b.dataset.tab); dragged=null;};
    bindTabResize?.();
  };

  function editSheetIcon19(sheet){
    const pool=ICONS_19.map(i=>`<button class="small secondary ${sheet.icon===i?'active':''}" data-pick-icon="${esc(i)}">${esc(i)}</button>`).join('');
    showModal('Reiterlogo / Icon bearbeiten',`<div class="grid2"><div class="field"><label>Icon/Text</label><input id="sheetIcon19" value="${esc(sheet.icon||'')}"></div><div class="field"><label>Eigenes Bild/Logo</label><input id="sheetLogo19" type="file" accept="image/*"></div><div class="field full"><label>Icon-Pool</label><div class="icon-pool">${pool}</div></div></div>`,[
      {text:'Speichern',action:()=>{sheet.icon=$('#sheetIcon19').value.trim()||sheet.icon||'📄'; const f=$('#sheetLogo19').files[0]; const done=()=>{save();closeModal();render();}; if(f){const rd=new FileReader(); rd.onload=()=>{sheet.logo=rd.result;done();}; rd.readAsDataURL(f);} else done();}},
      {text:'Bild entfernen',secondary:true,action:()=>{delete sheet.logo; save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    $$('#modalRoot [data-pick-icon]').forEach(b=>b.onclick=()=>{$('#sheetIcon19').value=b.dataset.pickIcon; $$('#modalRoot [data-pick-icon]').forEach(x=>x.classList.remove('active')); b.classList.add('active');});
  }

  const baseRenderTopic19=renderTopic;
  renderTopic=function(sheet,t,access){
    if(/Werkzeug/i.test(sheet.name||'')){t.compactFit=true;t.fitMode='werkzeug';t.uiWidth='100%';}
    let html=baseRenderTopic19(sheet,t,access);
    if(/Werkzeug/i.test(sheet.name||'')) html=html.replace('topic-card compact-fit','topic-card compact-fit werkzeug-fit').replace('topic-card"','topic-card werkzeug-fit"');
    return html;
  };

  const baseBindSheet19=bindSheetEvents;
  bindSheetEvents=function(sheet,access){
    baseBindSheet19(sheet,access);
    const share=$('#shareTabBtn'); if(share){const fresh=share.cloneNode(true); share.replaceWith(fresh); fresh.addEventListener('click',()=>openShareModal19(sheet));}
    const iconBtn=$('#editTabIconBtn'); if(iconBtn){const fresh=iconBtn.cloneNode(true); iconBtn.replaceWith(fresh); fresh.addEventListener('click',()=>editSheetIcon19(sheet));}
    if(sheet.vehicle && $('#fileTree')) bindTree($('#fileTree'),sheet.files,'file',access);
  };

  function openShareModal19(sheet){
    const me=currentUser();
    if(!canAccess(accessFor(sheet,me),'manage')){toast('Du brauchst Verwaltungsrechte für diesen Reiter.'); return;}
    showModal('Reiter teilen',`<p class="hint">Gib den Namen der Person ein. Es werden nur passende Treffer angezeigt.</p><div class="grid2"><div class="field"><label>Person suchen</label><input id="shareSearch19" placeholder="Name oder Benutzername"></div><div class="field"><label>Recht</label><select id="shareLevel19"><option value="view">ansehen</option><option value="edit">bearbeiten</option><option value="manage">verwalten</option></select></div></div><div id="shareResults19" class="share-user-results"></div><h4>Aktuell freigegeben</h4><div class="share-user-results">${Object.entries(sheet.permissions||{}).filter(([id])=>!id.startsWith('role:') && id!==sheet.ownerId).map(([id,lvl])=>`<div class="share-user-result"><span>${esc(userName(id))}</span><span class="pill">${esc(accessLabel(lvl))}</span><button class="small danger" data-unshare="${esc(id)}">Entfernen</button></div>`).join('')||'<p class="hint">Noch keine direkte Freigabe.</p>'}</div>`,[{text:'Schließen',secondary:true,action:closeModal}]);
    function run(){const q=$('#shareSearch19').value.trim().toLowerCase(); const hits=q?state.users.filter(u=>u.id!==me.id && (u.name.toLowerCase().includes(q)||u.id.toLowerCase().includes(q)||String(u.email||'').toLowerCase().includes(q))).slice(0,8):[]; $('#shareResults19').innerHTML=hits.map(u=>`<div class="share-user-result"><span><strong>${esc(u.name)}</strong><br><span class="hint">${esc(u.role)}</span></span><button class="small secondary" data-share-user="${esc(u.id)}">Teilen</button></div>`).join('') || '<p class="hint">Tippe einen Namen ein.</p>'; $$('#shareResults19 [data-share-user]').forEach(b=>b.onclick=()=>{sheet.permissions=sheet.permissions||{}; sheet.permissions[b.dataset.shareUser]=$('#shareLevel19').value; save(); closeModal(); toast('Reiter geteilt.'); render();});}
    $('#shareSearch19').addEventListener('input',run); run();
    $$('#modalRoot [data-unshare]').forEach(b=>b.onclick=()=>{delete sheet.permissions[b.dataset.unshare]; save(); closeModal(); toast('Freigabe entfernt.'); render();});
  }

  // Admin-Bereiche nur für Admins; Archiv sichtbar machen.
  const baseRenderSettingsPanel19=renderSettingsPanel;
  openSettings=function(panel='profil'){
    const admin=isAdmin19();
    const userNav=[['profil','Profil'],['design','Design & Logo'],['farben','Farben'],['dropdowns','Dropdowns'],['archiv','Archivierte Reiter'],['abo','Konto & Abo']];
    const adminNav=[['profil','Profil'],['design','Design & Logo'],['farben','Farben'],['dropdowns','Dropdowns'],['archiv','Archivierte Reiter'],['builder','Reiter-Builder'],['kernblock','Kernblock-Bearbeitung'],['rechte','Reiter & Rechte'],['nutzer','Nutzerverwaltung'],['abo','Konto & Abo'],['zahlungen','Bezahlmöglichkeiten'],['hinweise','Hinweise'],['supabase','Supabase'],['daten','Sicherung/Import'],['diagnose','Diagnose']];
    const nav=admin?adminNav:userNav;
    if(!nav.some(([id])=>id===panel)) panel='profil';
    showModal('Appeinstellungen',`<div class="settings-layout"><nav class="settings-nav">${nav.map(([id,txt])=>`<button class="${id===panel?'active':''}" data-set-panel="${id}">${txt}</button>`).join('')}</nav><section class="settings-panel" id="settingsPanel"></section></div>`,[{text:'Schließen',secondary:true,action:closeModal}]);
    $('#modalRoot').addEventListener('click',e=>{const b=e.target.closest('[data-set-panel]'); if(b) openSettings(b.dataset.setPanel);});
    renderSettingsPanel(panel);
  };
  renderSettingsPanel=function(panel){
    if(panel==='archiv'){renderArchivePanel19(); return;}
    const adminPanels=['builder','kernblock','rechte','nutzer','zahlungen','hinweise','supabase','daten','diagnose'];
    if(adminPanels.includes(panel)&&!isAdmin19()){ $('#settingsPanel').innerHTML='<h3>Nicht verfügbar</h3><p>Dieser Bereich ist nur für Admins sichtbar.</p>'; return; }
    baseRenderSettingsPanel19(panel);
    const diag=document.querySelector('#settingsPanel');
    if(panel==='diagnose' && diag) diag.innerHTML=diag.innerHTML.replace(/V26\.07\.18|V26\.07\.16/g,VERSION_19);
  };
  function archivedSheets19(){
    const u=currentUser(); const arr=[];
    for(const s of allRawSheets19()) if(s.archived && (isAdmin19() || s.ownerId===u.id || canAccess(accessFor(s,u),'view'))) arr.push(s);
    for(const list of Object.values(state.archivedSheets||{})) for(const s of list||[]) if(!arr.some(x=>x.id===s.id) && (isAdmin19() || s.ownerId===u.id || canAccess(accessFor(s,u),'view'))) arr.push(s);
    return arr.filter(s=>!s.catalog);
  }
  function renderArchivePanel19(){
    const el=$('#settingsPanel'); const arr=archivedSheets19();
    el.innerHTML=`<h3>Archivierte Reiter</h3><p class="hint">Archivierte Reiter bleiben hier abgelegt und können wieder aktiviert werden.</p><div class="archive-list">${arr.map(s=>`<div class="archive-item"><div><strong>${esc(s.icon||'📄')} ${esc(s.name)}</strong><br><span class="hint">${s.vehicle?'Autoreiter':'Neutraler Reiter'} · Besitzer: ${esc(userName(s.ownerId||'system'))}</span></div><button class="small secondary" data-restore-sheet="${esc(s.id)}">Wieder aktivieren</button></div>`).join('')||'<p class="hint">Keine archivierten Reiter.</p>'}</div>`;
    $$('[data-restore-sheet]').forEach(b=>b.onclick=()=>{const s=allRawSheets19().find(x=>x.id===b.dataset.restoreSheet)||archivedSheets19().find(x=>x.id===b.dataset.restoreSheet); if(s){s.archived=false; state.activeSheetId=s.id; save(); closeModal(); render(); toast('Reiter wieder aktiviert.');}});
  }

  // Dateiauswahl + Vorschau + Bestand-/Internetlinks.
  function readFileData19(file, cb){const rd=new FileReader(); rd.onload=()=>cb(rd.result); rd.readAsDataURL(file);}
  function addTreeFile19(parent,kind){
    const input=document.createElement('input'); input.type='file'; input.accept='image/*,.pdf,.txt,.doc,.docx,.xls,.xlsx,.xlsm';
    input.onchange=()=>{const f=input.files&&input.files[0]; if(!f)return; readFileData19(f,(data)=>{ parent.children=parent.children||[]; parent.open=true; const isImg=(f.type||'').startsWith('image/'); const node={id:uid('file'),type:'file',name:f.name,fileName:f.name,fileType:f.type||'',fileSize:f.size||0,fileData:data,thumb:isImg?data:'',icon:isImg?'🖼️':'📄',part:'',url:'',rowId:kind==='catalog'?uid('catrow'):undefined}; parent.children.push(node); if(kind==='catalog'){ const b=catalogBundleForRoot19(findRootOfNode19(parent)||null) || catalogBundleForTreeByNode19(node); if(b){b.rows=b.rows||[]; b.rows.push({id:node.rowId,image:node.thumb||'',name:node.name,part:'',category:'',note:'Datei: '+node.fileName});} } save(); render(); setTimeout(()=>editTreeNode(node,kind),80); });};
    input.click();
  }
  function findRootOfNode19(parent){return null;}
  function catalogBundleForTreeByNode19(node){const cat=catalogSheet19(); if(!cat) return null; for(const b of Object.values(cat.vehicleCatalogs||{})){ if(findNode(b.tree,node.id)) return b; } if(findNode(cat.tree||[],node.id)) return cat; return null;}
  addTreeChild=function(parent,type,kind){ if(type==='file') return addTreeFile19(parent,kind); parent.children=parent.children||[]; parent.open=true; parent.children.push({id:uid('folder'),type:'folder',name:'Neuer Ordner',open:true,children:[],icon:'📁'}); save(); render(); };
  uploadVehicleFile=function(sheet){sheet.files=Array.isArray(sheet.files)?sheet.files:defaultFileTree(); const root={children:sheet.files,open:true}; addTreeFile19(root,'file');};

  function fileActions19(n,kind,canEdit){
    if(n.type!=='file') return '';
    const stock=`<button class="small secondary tree-file-action" data-tree-stock title="Im Teilebestand suchen/markieren">📦</button>`;
    const web=n.url?`<button class="small secondary tree-file-action" data-tree-web title="Internetlink öffnen">🌐</button>`:'';
    const preview=`<button class="small secondary tree-file-action" data-tree-preview title="Vorschau anzeigen">👁</button>`;
    return `<span class="tree-actions">${preview}${stock}${web}${canEdit?'' : ''}</span>`;
  }
  renderTree=function(nodes,kind,access,depth=0){
    return (nodes||[]).map(n=>{ const canEdit=kind==='catalog'?isAdmin19():canAccess(access,'edit'); const icon=n.thumb?`<img class="thumb" src="${esc(n.thumb)}" alt="">`:`<span class="thumb">${esc(n.icon || (n.type==='folder'?'📁':'📄'))}</span>`; return `<div class="tree-node" data-node="${esc(n.id)}" data-kind="${kind}"><div class="tree-row" style="margin-left:${depth*14}px" data-tree-click><span class="tree-toggle">${n.type==='folder'?(n.open?'▾':'▸'):'•'}</span>${icon}<strong>${esc(n.name)}</strong>${n.part?`<span class="pill">${esc(n.part)}</span>`:''}<span style="flex:1"></span>${fileActions19(n,kind,canEdit)}${canEdit?`<button class="small secondary" data-tree-edit>✎</button>${n.type==='folder'?'<button class="small secondary" data-tree-add-folder>+ Ordner</button><button class="small secondary" data-tree-add-file>+ Datei</button>':''}<button class="small danger" data-tree-delete title="Löschen">🗑</button>`:''}</div>${n.type==='folder'&&n.open?`<div class="tree-children">${renderTree(n.children||[],kind,access,depth+1)}</div>`:''}</div>`; }).join('');
  };
  function removeNodeDeep19(nodes,id){ if(!Array.isArray(nodes))return null; const idx=nodes.findIndex(n=>n.id===id); if(idx>=0)return nodes.splice(idx,1)[0]; for(const n of nodes){const r=removeNodeDeep19(n.children,id); if(r)return r;} return null; }
  bindTree=function(container,root,kind,access){
    if(!container) return;
    container.onclick=e=>{
      const nodeEl=e.target.closest('[data-node]'); if(!nodeEl)return; const node=findNode(root,nodeEl.dataset.node); if(!node)return;
      if(e.target.closest('[data-tree-preview]')){e.stopPropagation(); return openPreview19(root,node.id);}
      if(e.target.closest('[data-tree-stock]')){e.stopPropagation(); return jumpToStock19(node.part||node.name);}
      if(e.target.closest('[data-tree-web]')){e.stopPropagation(); if(node.url) window.open(node.url,'_blank','noopener'); return;}
      if(e.target.closest('[data-tree-delete]')){e.stopPropagation(); if(confirm(`${node.type==='folder'?'Ordner':'Datei'} „${node.name}“ löschen?`)){const removed=removeNodeDeep19(root,node.id); if(kind==='catalog'&&removed?.rowId){const b=catalogBundleForRoot19(root)||catalogBundleForTreeByNode19(removed); const i=(b?.rows||[]).findIndex(r=>r.id===removed.rowId); if(i>=0)b.rows.splice(i,1);} save(); render(); toast('Gelöscht');} return;}
      if(e.target.closest('[data-tree-edit]')){e.stopPropagation(); return editTreeNode(node,kind);}
      if(e.target.closest('[data-tree-add-folder]')){e.stopPropagation(); return addTreeChild(node,'folder',kind);}
      if(e.target.closest('[data-tree-add-file]')){e.stopPropagation(); return addTreeChild(node,'file',kind);}
      if(e.target.closest('[data-tree-click]')){ if(node.type==='folder'){node.open=!node.open; if(kind==='file')state.fileSelected=node.id; save(); render();} else { if(kind==='file'){state.fileSelected=node.id; save(); render(); openPreview19(root,node.id);} if(kind==='catalog'){state.catalogSelected=node.id; save(); render(); setTimeout(()=>{const el=document.querySelector(`[data-catalog-row="${CSS.escape(node.rowId||'')}"]`); highlight19(el,'highlight-row');},80);} } }
    };
  };
  function flatFiles19(nodes,parent=[]){let out=[]; for(const n of nodes||[]){ if(n.type==='file') out.push({...n,_siblings:parent}); if(n.type==='folder') out=out.concat(flatFiles19(n.children||[],n.children||[])); } return out;}
  function openPreview19(root,id){
    const all=flatFiles19(root); let idx=Math.max(0,all.findIndex(x=>x.id===id)); if(!all.length)return toast('Keine Dateien im Ordner.');
    function renderOne(){ const n=all[idx]; const content=n.fileData?(String(n.fileType||'').startsWith('image/')||String(n.fileData).startsWith('data:image/')?`<img src="${esc(n.fileData)}" alt="${esc(n.name)}">`:`<iframe class="preview-frame" src="${esc(n.fileData)}"></iframe>`):(n.thumb?`<img src="${esc(n.thumb)}" alt="${esc(n.name)}">`:`<div class="release-card"><h3>${esc(n.icon||'📄')} ${esc(n.name)}</h3><p>Für diese Datei ist noch keine eingebettete Vorschau gespeichert.</p></div>`); $('#previewBody19').innerHTML=`<div class="preview-meta"><strong>${esc(n.name)}</strong><span class="hint">${idx+1} / ${all.length}</span></div><div class="preview-stage"><button class="preview-arrow left" id="prevPreview19">‹</button>${content}<button class="preview-arrow right" id="nextPreview19">›</button></div><p class="hint">Vorschau im Browser. Bei mehreren Dateien im Ordner mit den Pfeilen wechseln.</p>`; $('#prevPreview19').onclick=()=>{idx=(idx-1+all.length)%all.length;renderOne();}; $('#nextPreview19').onclick=()=>{idx=(idx+1)%all.length;renderOne();}; }
    showModal('Dateivorschau',`<div id="previewBody19"></div>`,[{text:'Schließen',secondary:true,action:closeModal}]); renderOne();
  }
  editTreeNode=function(node,kind){
    const pool=ICONS_19.map(i=>`<button class="small secondary ${node.icon===i?'active':''}" data-pick-icon="${esc(i)}">${esc(i)}</button>`).join('');
    showModal((node.type==='folder'?'Ordner':'Datei')+' bearbeiten',`<div class="grid2"><div class="field"><label>Name</label><input id="tnName19" value="${esc(node.name)}"></div><div class="field"><label>Teilenummer / Bildtext</label><input id="tnPart19" value="${esc(node.part||'')}"></div><div class="field"><label>Internetlink</label><input id="tnUrl19" value="${esc(node.url||'')}" placeholder="https://…"></div><div class="field"><label>Icon/Text</label><input id="tnIcon19" value="${esc(node.icon||'')}"></div><div class="field"><label>Vorschau/Bild ändern</label><input id="tnFile19" type="file" accept="image/*,.pdf,.txt"></div><div class="field full"><label>Icon-Pool</label><div class="icon-pool">${pool}</div></div></div>`,[
      {text:'Speichern',action:()=>{node.name=$('#tnName19').value.trim()||node.name; node.part=$('#tnPart19').value.trim(); node.url=$('#tnUrl19').value.trim(); node.icon=$('#tnIcon19').value.trim()||node.icon; const f=$('#tnFile19').files[0]; const done=()=>{ if(kind==='catalog'&&node.rowId){const b=catalogBundleForTreeByNode19(node); const r=(b?.rows||[]).find(x=>x.id===node.rowId); if(r){r.name=node.part||node.name; r.part=node.part||r.part; r.image=node.thumb||r.image; r.note=node.url?('Internet: '+node.url):r.note;}} save(); closeModal(); render();}; if(f){readFileData19(f,data=>{node.fileData=data;node.fileType=f.type||'';node.fileName=f.name; if((f.type||'').startsWith('image/'))node.thumb=data; done();});} else done();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    $$('#modalRoot [data-pick-icon]').forEach(b=>b.onclick=()=>{$('#tnIcon19').value=b.dataset.pickIcon; $$('#modalRoot [data-pick-icon]').forEach(x=>x.classList.remove('active')); b.classList.add('active');});
  };

  function searchStock19(part){
    const q=String(part||'').trim().toLowerCase(); if(!q)return null;
    const sheets=allSheetsForUser().filter(s=>/Teilebestand/i.test(s.name||''));
    for(const sh of sheets){ for(const t of sh.themes||[]){ for(const col of t.columns||[]){const header=(col.parts||[]).join(' ').toLowerCase(); if(/teil|part|pn|oem|sn|nummer|artikel/.test(header)){for(const g of t.rowGroups||[]){const r=Array.isArray(g)?g[0]:g; const val=String(sh.cells?.[`${r}:${col.col}`]?.text||sh.cells?.[`${r}:${col.col}`]?.raw||'').toLowerCase(); if(val && (val.includes(q)||q.includes(val))) return {sheet:sh,theme:t,row:r,col};}}} } }
    return null;
  }
  function jumpToStock19(part){
    const hit=searchStock19(part);
    if(!hit){toast('Teil ist nicht auf Lager bzw. nicht im Teilebestand gefunden.'); return;}
    state.activeSheetId=hit.sheet.id; save(); render(); setTimeout(()=>highlight19(document.querySelector(`#${CSS.escape(topicDomId(hit.sheet,hit.theme))} td[data-r="${hit.row}"][data-c="${hit.col.col}"]`)),120);
  }
  function highlight19(el,cls='highlight-cell'){ if(!el)return; document.querySelectorAll('.highlight-cell,.highlight-row').forEach(x=>x.classList.remove('highlight-cell','highlight-row')); el.classList.add(cls); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}); setTimeout(()=>el.classList.remove(cls),7500); }

  const baseRenderCatalog19=renderCatalog;
  renderCatalog=function(sheet){
    baseRenderCatalog19(sheet);
    const bundle=(function(){const cat=catalogSheet19(); if(!cat)return null; const key=[state.catalogSelection?.brand||'allgemein',state.catalogSelection?.model||'allgemein',state.catalogSelection?.motor||'allgemein',String(state.catalogSelection?.vin||'').trim()||'ohne_vin'].join('__'); return cat.vehicleCatalogs?.[key]||cat;})();
    $$('#main [data-catalog-row]').forEach(row=>{const id=row.dataset.catalogRow; const r=(bundle?.rows||[]).find(x=>x.id===id); const actionCell=document.createElement('td');});
    // Zusätzliche Aktionsspalte an die Dateitabelle anhängen, ohne die bestehende Kataloglogik zu verändern.
    const table=$('#main .catalog-table');
    if(table && !table.dataset.v19){table.dataset.v19='1'; const hr=table.querySelector('thead tr'); if(hr) hr.insertAdjacentHTML('beforeend','<th>Links</th>'); table.querySelectorAll('tbody tr').forEach(tr=>{const id=tr.dataset.catalogRow; const row=(bundle?.rows||[]).find(x=>x.id===id)||{}; const node=findNode(bundle?.tree||[], state.catalogSelected)||null; tr.insertAdjacentHTML('beforeend',`<td><div class="catalog-actions"><button class="small secondary" data-stock-row="${esc(row.part||row.name||'')}">📦 Teilebestand</button>${row.note&&/^Internet:/.test(row.note)?`<a class="btn small secondary" target="_blank" rel="noopener" href="${esc(row.note.replace(/^Internet:\s*/,''))}">🌐 Internet</a>`:''}</div></td>`);}); table.querySelectorAll('[data-stock-row]').forEach(b=>b.onclick=()=>jumpToStock19(b.dataset.stockRow));}
  };

  const baseTopActions19=typeof topActions18==='function'?topActions18:null;
  function topActions19(){ const u=currentUser?.(); const box=document.querySelector('.top-actions'); if(!box||!u)return; const admin=isAdmin19(); box.classList.add('lh-like'); box.innerHTML=`<span class="version-chip">${VERSION_19}</span><span class="user-chip auth-user-box" id="userChip">👤 ${esc(u.name)} <span class="hint" style="color:#fff9">${esc(u.role)}</span></span><span class="version-chip" id="saveInfo">${state.savedAt?'Gespeichert · '+new Date(state.savedAt).toLocaleTimeString('de-DE'):'Vorschau nicht gespeichert'}</span><button id="topManageBtn" class="secondary">${admin?'Verwaltung':'Einstellungen'}</button>${admin?'<button id="saveBtn" class="secondary">Cloud speichern</button>':''}<button id="darkToggle" class="secondary">🌙</button><button id="logoutBtn" class="secondary">Abmelden</button>`; $('#topManageBtn')?.addEventListener('click',()=>openSettings('profil')); $('#saveBtn')?.addEventListener('click',()=>{save();toast('Manuell gespeichert')}); $('#darkToggle')?.addEventListener('click',()=>{state.settings.dark=!state.settings.dark;save();render();}); $('#logoutBtn')?.addEventListener('click',logout); }
  const baseRender19=render;
  render=function(){baseRender19(); setTitle19(); topActions19();};
  const baseReleaseNotes19=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.19 · Tabellen, Reiterlogo, Teilen, Archiv und Datei-Vorschau',body:'Werkzeug wird kompakter ohne unnötiges Rechts-Scrollen dargestellt, alle Tabellenzeilen sind niedriger, Reiterlogos werden in der Navigation angezeigt, Teilen nutzt eine Suchzeile statt einer vollständigen Nutzerliste, Admin-Bereiche sind nur für Admins sichtbar, archivierte Reiter können in den Einstellungen reaktiviert werden, Dateien im Teilekatalog haben Vorschau, Bestand-/Internetlinks und Bildnavigation.',jump:'settings:archiv'}, ...baseReleaseNotes19()];};
})();



/* ---------------- V26.07.20: Bestandslogik, Bestellhistorie, Rechtsklick ---------------- */
(function(){
  const VERSION_20='V26.07.20';
  const STORE_KEY_20='kfz_app_'+VERSION_20+'_state';
  const LIVE_KEY='kfz_app_live_state_preserved';
  const PREV_KEYS=['kfz_app_V26.07.19_state','kfz_app_V26.07.18_state','kfz_app_V26.07.16_state'];

  const css=document.createElement('style');
  css.textContent=`
    .context-menu20{position:fixed;z-index:9999;min-width:230px;background:var(--surface);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow);padding:8px;display:none;color:var(--text)}
    .context-menu20 button{display:flex;align-items:center;gap:8px;width:100%;background:transparent;color:var(--text);border:0;border-radius:10px;padding:8px 10px;text-align:left;cursor:pointer;font-weight:600}
    .context-menu20 button:hover{background:var(--secondary)}
    .inventory-match{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;border:1px solid var(--border);background:var(--surface2);border-radius:14px;padding:10px;margin:8px 0}
    .stock-warning{background:var(--danger);border:1px solid color-mix(in srgb,var(--danger),#9b1c1c 35%);border-radius:14px;padding:10px;margin:8px 0}
    .history-table{width:100%;border-collapse:separate;border-spacing:0 6px}.history-table th{font-size:.78rem;color:var(--muted);text-align:left}.history-table td{background:var(--surface2);padding:8px;border-top:1px solid var(--border);border-bottom:1px solid var(--border)}.history-table td:first-child{border-left:1px solid var(--border);border-radius:10px 0 0 10px}.history-table td:last-child{border-right:1px solid var(--border);border-radius:0 10px 10px 0}.history-summary{display:flex;gap:10px;flex-wrap:wrap;margin:8px 0}.history-summary .pill{font-weight:700}.linked-cell20{outline:2px solid var(--primary);outline-offset:-2px}.manual-stock20{box-shadow:inset 0 0 0 2px var(--warning)}
  `;
  document.head.appendChild(css);

  const baseInit20=initState;
  initState=function(){
    let s=null;
    try{
      s=JSON.parse(localStorage.getItem(STORE_KEY_20)||localStorage.getItem(LIVE_KEY)||PREV_KEYS.map(k=>localStorage.getItem(k)).find(Boolean)||'null');
    }catch(e){s=null;}
    if(!(s&&s.users)) s=baseInit20();
    s.version=VERSION_20;
    migrate20(s);
    return s;
  };
  const baseSave20=save;
  save=function(){
    if(!state) return;
    state.version=VERSION_20;
    state.savedAt=new Date().toISOString();
    try{localStorage.setItem(STORE_KEY_20,JSON.stringify(state));localStorage.setItem(LIVE_KEY,JSON.stringify(state));}
    catch(e){console.warn(e);try{baseSave20();}catch(_){}}
    updateSaveInfo?.();
  };
  function migrate20(s){
    s.purchaseHistory=s.purchaseHistory||{};
    s.stockMovements=s.stockMovements||[];
    s.installLinks=s.installLinks||{};
    s.stockIndex=s.stockIndex||{};
    for(const sh of allStateSheets20(s)){
      if(sh.themes) for(const t of sh.themes) for(const g of (t.rowGroups||[])){
        const r=Array.isArray(g)?g[0]:g;
        for(const col of (t.columns||[])){
          const cell=sh.cells?.[`${r}:${col.col}`];
          if(cell && cell.stockConsumption && !cell.installLink) cell.installLink=cell.stockConsumption;
        }
      }
    }
    rebuildStockIndex20(s);
  }
  function allStateSheets20(s=state){return [...(s.adminSheets||[]), ...Object.values(s.userSheets||{}).flat()];}
  function setTitle20(){document.title='KFZ-APP '+VERSION_20; const h=document.querySelector('#loginView h2'); if(h) h.textContent='KFZ-APP Vorschau · '+VERSION_20; const vc=document.querySelector('.version-chip'); if(vc) vc.textContent=VERSION_20;}
  function normPart20(v){return String(v||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');}
  function normMaker20(v){return String(v||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();}
  function colText20(col){return ((col?.parts||[]).join(' ')+' '+(col?.label||'')).toLowerCase();}
  function colHas20(col,rx){return rx.test(colText20(col));}
  function findCol20(t,rx,notRx){return (t?.columns||[]).find(c=>rx.test(colText20(c)) && !(notRx&&notRx.test(colText20(c))));}
  function rows20(t){let rs=(t?.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g).filter(Boolean); if(!rs.length&&t){for(let r=t.dataStart||5;r<=t.endRow;r++)rs.push(r)} return rs;}
  function getCell20(sheet,r,c){if(!sheet.cells)sheet.cells={}; return sheet.cells[`${r}:${c}`]||(sheet.cells[`${r}:${c}`]={r,c,a1:colLetter(c)+r,text:'',raw:'',style:{}});}
  function cellVal20(sheet,r,col){if(!col)return ''; const cell=sheet.cells?.[`${r}:${col.col}`]; return String(cell?.text??cell?.raw??'').trim();}
  function setCellVal20(sheet,r,col,v){if(!col)return; const cell=getCell20(sheet,r,col.col); cell.text=String(v??''); cell.raw=cell.text;}
  function parseQty20(v){const n=parseFloat(String(v??'').replace(/\./g,'').replace(',', '.').replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n;}
  function fmtQty20(n){return Number.isInteger(n)?String(n):String(Math.round(n*1000)/1000).replace('.', ',');}
  function formatMoney20(n){return (isFinite(n)?n:0).toLocaleString('de-DE',{minimumFractionDigits:2,maximumFractionDigits:2})+' €';}
  function stockSheet20(src=state){return allStateSheets20(src).find(s=>/^Teilebestand$/i.test(s.name||'')) || allStateSheets20(src).find(s=>/Teilebestand/i.test(s.name||''));}
  function stockTheme20(sheet=stockSheet20(state)){return (sheet?.themes||[]).find(t=>/gesamt|gesammt|teilebestand/i.test(t.name||t.displayName||'')) || sheet?.themes?.[0];}
  function stockCols20(t){return {avail:findCol20(t,/verfügbar|verfuegbar/),zulauf:findCol20(t,/zulauf/),bought:findCol20(t,/gekauft/),maker:findCol20(t,/hersteller/),part:findCol20(t,/teilenummer|part|pn|oem|sn/),price:findCol20(t,/preis/)};}
  function partCols20(t){return {maker:findCol20(t,/hersteller/),part:findCol20(t,/teilenummer|part|pn|oem|sn/),avail:findCol20(t,/verfügbar|verfuegbar/),built:findCol20(t,/verbaut/),zahl:findCol20(t,/\bzahl\b/),bought:findCol20(t,/gekauft/),zulauf:findCol20(t,/zulauf/),price:findCol20(t,/preis/)};}
  function rowKey20(maker,part){const m=normMaker20(maker), p=normPart20(part); return (m&&p)?m+'__'+p:'';}
  function rowPartKey20(sheet,t,r){const c=partCols20(t); return rowKey20(cellVal20(sheet,r,c.maker),cellVal20(sheet,r,c.part));}
  function rowLabel20(sheet,t,r){const c=partCols20(t); return [cellVal20(sheet,r,c.maker),cellVal20(sheet,r,c.part)].filter(Boolean).join(' · ')||'Bauteil';}
  function rebuildStockIndex20(s=state){
    const ss=stockSheet20(s); const st=stockTheme20(ss); const idx={}; if(!ss||!st)return idx;
    const c=stockCols20(st);
    for(const r of rows20(st)){
      const maker=cellVal20(ss,r,c.maker), part=cellVal20(ss,r,c.part); const key=rowKey20(maker,part); if(!key)continue;
      idx[key]=idx[key]||[]; idx[key].push({sheetId:ss.id,topicKey:st.key||topicKey(st),row:r,maker,part,available:parseQty20(cellVal20(ss,r,c.avail)),label:rowLabel20(ss,st,r)});
    }
    s.stockIndex=idx; return idx;
  }
  function findStockHits20(maker,part){const key=rowKey20(maker,part); if(!key)return []; const idx=rebuildStockIndex20(state); return (idx[key]||[]).map(h=>{const sh=sheetById(h.sheetId); const th=(sh?.themes||[]).find(t=>(t.key||topicKey(t))===h.topicKey); const c=stockCols20(th); return {...h,sheet:sh,theme:th,cols:c,available:parseQty20(cellVal20(sh,h.row,c.avail))};}).filter(h=>h.sheet&&h.theme);}
  function hitByKey20(stockKey){const arr=(state.stockIndex?.[stockKey]||[]); if(!arr.length)return null; const h=arr[0]; const sh=sheetById(h.sheetId); const th=(sh?.themes||[]).find(t=>(t.key||topicKey(t))===h.topicKey); if(!sh||!th)return null; return {...h,sheet:sh,theme:th,cols:stockCols20(th),available:parseQty20(cellVal20(sh,h.row,stockCols20(th).avail))};}
  function isInstallTopic20(t){const n=String(t?.displayName||t?.name||'').toLowerCase(); return /verbauliste/.test(n)||/wartungsarbeiten|verbrauchs/.test(n);}
  function installTriggerCol20(t,col){const n=String(t?.displayName||t?.name||'').toLowerCase(); const h=colText20(col); if(/verbauliste/.test(n)&&/verbaut/.test(h))return true; if(/wartungsarbeiten|verbrauchs/.test(n)&&/\bzahl\b/.test(h))return true; return false;}
  function syncVehicleAvailability20(){
    rebuildStockIndex20(state);
    for(const sh of allStateSheets20()){
      if(!sh.vehicle)continue;
      for(const t of (sh.themes||[])){
        if(!isInstallTopic20(t))continue;
        const c=partCols20(t); if(!c.avail||!c.maker||!c.part)continue;
        for(const r of rows20(t)){
          const hits=findStockHits20(cellVal20(sh,r,c.maker),cellVal20(sh,r,c.part));
          if(hits.length===1){setCellVal20(sh,r,c.avail,fmtQty20(hits[0].available)); getCell20(sh,r,c.avail.col).syncedFromStock=true;}
        }
      }
    }
  }
  function adjustStock20(hit,delta,reason){
    const cell=getCell20(hit.sheet,hit.row,hit.cols.avail.col); const before=parseQty20(cell.text??cell.raw); const after=before+delta;
    if(after<0)return false;
    cell.text=fmtQty20(after); cell.raw=cell.text;
    state.stockMovements=state.stockMovements||[];
    state.stockMovements.push({id:uid('move'),at:new Date().toISOString(),stockKey:rowKey20(hit.maker,hit.part),delta,before,after,reason:reason||'',userId:state.currentUserId});
    rebuildStockIndex20(state); syncVehicleAvailability20(); return true;
  }
  function restoreCell20(cell,oldValue){cell.text=String(oldValue??''); cell.raw=cell.text;}
  function confirmConsumption20(sheet,t,r,col,oldValue,newValue){
    if(!sheet.vehicle||!isInstallTopic20(t)||!installTriggerCol20(t,col))return false;
    const c=partCols20(t); const cell=getCell20(sheet,r,col.col); const newQty=parseQty20(newValue); const oldLinked=cell.installLink&&typeof cell.installLink.qty!=='undefined'?parseQty20(cell.installLink.qty):parseQty20(oldValue); const diff=newQty-oldLinked;
    if(Math.abs(diff)<0.000001){cell.installLink={...(cell.installLink||{}),qty:newQty}; return false;}
    const maker=cellVal20(sheet,r,c.maker), part=cellVal20(sheet,r,c.part); const key=rowKey20(maker,part);
    if(!key){toast('Hersteller und Teilenummer fehlen. Bestand kann nicht automatisch zugeordnet werden.'); cell.installLink={qty:newQty,manual:true}; return false;}
    let hits=findStockHits20(maker,part);
    if(diff<0 && cell.installLink?.stockKey){const oldHit=hitByKey20(cell.installLink.stockKey); if(oldHit) hits=[oldHit];}
    if(!hits.length){toast('Kein passendes Teil im Teilebestand gefunden. Bitte manuell austragen.'); cell.installLink={qty:newQty,manual:true}; return false;}
    const doWithHit=(hit)=>{
      if(diff>0 && hit.available<diff){restoreCell20(cell,oldLinked); showModal('Bestand reicht nicht',`<div class="stock-warning"><strong>Nicht möglich:</strong> Im Teilebestand sind nur ${esc(fmtQty20(hit.available))} verfügbar, benötigt werden ${esc(fmtQty20(diff))}.</div>`,[{text:'OK',action:()=>{closeModal();render();}}]); return true;}
      const action=diff>0?'aus dem Teilebestand entnehmen':'in den Teilebestand zurückbuchen';
      showModal('Bestand synchronisieren',`<p>Erkanntes Bauteil:</p><div class="inventory-match"><div><strong>${esc(hit.label)}</strong><br><span class="hint">Hersteller + normalisierte Teilenummer stimmen überein.</span><br><span class="hint">Aktuell verfügbar: ${esc(fmtQty20(hit.available))}</span></div><span class="pill">${diff>0?'-':'+'}${esc(fmtQty20(Math.abs(diff)))}</span></div><p>Soll diese Menge ${action} werden?</p>`,[
        {text:'Ja, bestätigen',action:()=>{const ok=adjustStock20(hit,diff>0?-diff:Math.abs(diff),'Verbau '+sheet.name+' · '+(t.displayName||t.name)); if(ok){cell.installLink={qty:newQty,stockKey:rowKey20(hit.maker,hit.part),manual:false}; closeModal(); save(); render(); toast('Teilebestand aktualisiert.');}else{restoreCell20(cell,oldLinked); closeModal(); render(); toast('Bestand konnte nicht aktualisiert werden.');}}},
        {text:'Nein, manuell austragen',secondary:true,action:()=>{cell.installLink={qty:newQty,manual:true,stockKey:null}; closeModal(); save(); render(); toast('Keine automatische Entnahme. Bitte manuell austragen.');}}
      ]); return true;
    };
    if(hits.length===1)return doWithHit(hits[0]);
    showModal('Mehrere passende Teile gefunden',`<p>Bitte wähle das richtige Bauteil aus.</p>${hits.map((h,i)=>`<div class="inventory-match"><div><strong>${esc(h.label)}</strong><br><span class="hint">Verfügbar: ${esc(fmtQty20(h.available))}</span></div><button data-pick-stock20="${i}">Dieses Teil</button></div>`).join('')}`,[{text:'Abbrechen',secondary:true,action:()=>{restoreCell20(cell,oldLinked);closeModal();render();}}]);
    $$('#modalRoot [data-pick-stock20]').forEach(b=>b.onclick=()=>doWithHit(hits[+b.dataset.pickStock20]));
    return true;
  }
  function afterCellChange20(sheet,t,r,col,oldValue,newValue){
    const handled=confirmConsumption20(sheet,t,r,col,oldValue,newValue);
    if(handled)return;
    const h=colText20(col); const oldN=parseQty20(oldValue), newN=parseQty20(newValue);
    if((/gekauft|zulauf/.test(h)) && newN>oldN){
      openPurchaseHistory20(sheet,t,r,{prefillQty:newN-oldN,prefillStatus:/zulauf/.test(h)?'im Zulauf':'gekauft',reason:'increase'});
      return;
    }
    if(/hersteller|teilenummer|part|pn|oem|sn|verfügbar|verfuegbar/.test(h)){rebuildStockIndex20(state); syncVehicleAvailability20();}
    save(); render();
  }

  editCellInline=function(sheet,t,td){
    const r=+td.dataset.r,c=+td.dataset.c; const col=t.columns.find(cc=>cc.col===c); const cell=getCell20(sheet,r,c); const old=String(cell.text??''); const dd=dropdownForColumn(col); const inner=td.querySelector('.cell-inner');
    if(dd&&dd.values.length){const sel=document.createElement('select'); sel.innerHTML=[''].concat(dd.values).map(v=>`<option ${v===cell.text?'selected':''}>${esc(v)}</option>`).join(''); inner.innerHTML=''; inner.appendChild(sel); sel.focus(); const finish=()=>{const val=sel.value; if(val===old){render();return;} cell.text=val; cell.raw=val; afterCellChange20(sheet,t,r,col,old,val);}; sel.onchange=sel.onblur=finish; return;}
    const span=td.querySelector('[data-cell-text]')||td.querySelector('a'); if(!span)return; span.classList.add('editing'); span.contentEditable='true'; span.focus(); document.execCommand?.('selectAll',false,null); let done=false; const finish=()=>{if(done)return; done=true; const val=span.textContent.trim(); cell.text=val; cell.raw=val; span.contentEditable='false'; span.classList.remove('editing'); if(val!==old) afterCellChange20(sheet,t,r,col,old,val); else render();}; span.onblur=finish; span.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault(); span.blur()} if(e.key==='Escape'){done=true; cell.text=old; span.textContent=old; span.blur(); render();}};
  };
  openCellEditor=function(sheet,t,td){
    const r=+td.dataset.r,c=+td.dataset.c; const col=t.columns.find(cc=>cc.col===c); const cell=getCell20(sheet,r,c); const old=String(cell.text??'');
    showModal('Zelle bearbeiten',`<div class="grid2"><div class="field"><label>Wert</label><input id="ceValue" value="${esc(cell.text||'')}"></div><div class="field"><label>Link</label><input id="ceLink" value="${esc(cell.hyperlink||'')}"></div><div class="field full"><label>Notiz / Hovertext</label><textarea id="ceNote">${esc(cell.comment||'')}</textarea></div><div class="field"><label>Hintergrundfarbe</label><input id="ceBg" type="color" value="${esc(cell.style?.backgroundColor||'#ffffff')}"></div><div class="field"><label>Textfarbe</label><input id="ceColor" type="color" value="${esc(cell.style?.color||'#111827')}"></div><div class="field full"><label>Bild hochladen</label><input id="ceImg" type="file" accept="image/*"></div></div>`,[
      {text:'Speichern',action:()=>{const val=$('#ceValue').value; cell.text=val; cell.raw=val; cell.hyperlink=$('#ceLink').value.trim(); cell.comment=$('#ceNote').value.trim(); cell.style=cell.style||{}; cell.style.backgroundColor=$('#ceBg').value; cell.style.color=$('#ceColor').value; const f=$('#ceImg').files[0]; const finish=()=>{closeModal(); afterCellChange20(sheet,t,r,col,old,val);}; if(f){const rd=new FileReader(); rd.onload=()=>{cell.image=rd.result; finish();}; rd.readAsDataURL(f);} else finish();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
  };

  function historyKey20(sheet,t,r){return rowPartKey20(sheet,t,r)||`${sheet.id}__${t.key||topicKey(t)}__${r}`;}
  function historyEntries20(sheet,t,r){const key=historyKey20(sheet,t,r); state.purchaseHistory=state.purchaseHistory||{}; state.purchaseHistory[key]=state.purchaseHistory[key]||[]; return state.purchaseHistory[key];}
  function updateAveragePrice20(sheet,t,r){const entries=historyEntries20(sheet,t,r).filter(e=>e.status!=='storniert' && parseQty20(e.qty)>0 && parseQty20(e.unitPrice)>0); const qty=entries.reduce((a,e)=>a+parseQty20(e.qty),0); const total=entries.reduce((a,e)=>a+parseQty20(e.qty)*parseQty20(e.unitPrice),0); const avg=qty?total/qty:0; const col=partCols20(t).price; if(col&&avg)setCellVal20(sheet,r,col,formatMoney20(avg)); return avg;}
  function openPurchaseHistory20(sheet,t,r,opts={}){
    const entries=historyEntries20(sheet,t,r); const avg=updateAveragePrice20(sheet,t,r); const label=rowLabel20(sheet,t,r); const list=entries.length?`<table class="history-table"><thead><tr><th>Datum</th><th>Status</th><th>Menge</th><th>Einzelpreis</th><th>Händler</th><th>Rechnung</th><th>Datei</th><th></th></tr></thead><tbody>${entries.map(e=>`<tr><td>${esc(e.date||'')}</td><td>${esc(e.status||'')}</td><td>${esc(e.qty||'')}</td><td>${esc(e.unitPrice||'')}</td><td>${esc(e.vendor||'')}</td><td>${esc(e.invoice||'')}</td><td>${e.fileData?`<button class="small secondary" data-open-invoice20="${esc(e.id)}">Ansehen</button>`:''}</td><td><button class="small danger" data-del-history20="${esc(e.id)}">×</button></td></tr>`).join('')}</tbody></table>`:'<p class="hint">Noch keine Bestellhistorie für diese Zeile.</p>';
    showModal('Bestellhistorie · '+label,`<div class="history-summary"><span class="pill">Durchschnittspreis: ${esc(formatMoney20(avg))}</span><span class="pill">Einträge: ${entries.length}</span></div>${list}<hr><h3>Historieneintrag hinzufügen</h3><div class="grid2"><div class="field"><label>Kauf-/Bestelldatum</label><input id="phDate20" type="date" value="${new Date().toISOString().slice(0,10)}"></div><div class="field"><label>Status</label><select id="phStatus20"><option ${opts.prefillStatus==='gekauft'?'selected':''}>gekauft</option><option ${opts.prefillStatus==='im Zulauf'?'selected':''}>im Zulauf</option><option>angekommen</option><option>storniert</option></select></div><div class="field"><label>Menge</label><input id="phQty20" inputmode="decimal" value="${esc(opts.prefillQty||'')}"></div><div class="field"><label>Einzelpreis</label><input id="phPrice20" inputmode="decimal" placeholder="z. B. 12,50"></div><div class="field"><label>Händler</label><input id="phVendor20"></div><div class="field"><label>Rechnungsnummer / Suchtext</label><input id="phInvoice20"></div><div class="field full"><label>Zusatztext / Notiz</label><textarea id="phNote20" placeholder="z. B. Hinweis zur Rechnung oder Bestellung"></textarea></div><div class="field full"><label>Rechnung/Datei ablegen</label><input id="phFile20" type="file" accept="image/*,.pdf,.txt,.doc,.docx,.xls,.xlsx,.xlsm"></div><div class="field full"><label>Rechnung suchen</label><input id="phSearch20" placeholder="Rechnungsnummer oder Text suchen"><div id="phSearchResults20" class="search-results"></div></div></div>`,[
      {text:'Eintrag speichern',action:()=>{const f=$('#phFile20').files[0]; const saveEntry=(data,type,name)=>{const e={id:uid('hist'),date:$('#phDate20').value,status:$('#phStatus20').value,qty:$('#phQty20').value,unitPrice:$('#phPrice20').value,vendor:$('#phVendor20').value.trim(),invoice:$('#phInvoice20').value.trim(),note:$('#phNote20').value.trim(),fileData:data||'',fileType:type||'',fileName:name||'',userId:state.currentUserId,createdAt:new Date().toISOString()}; entries.push(e); updateAveragePrice20(sheet,t,r); save(); closeModal(); render(); toast('Bestellhistorie gespeichert.');}; if(f){const rd=new FileReader(); rd.onload=()=>saveEntry(rd.result,f.type||'',f.name); rd.readAsDataURL(f);} else saveEntry();}},
      {text:'Schließen',secondary:true,action:()=>{closeModal();save();render();}}
    ]);
    const runSearch=()=>{const q=String($('#phSearch20')?.value||'').toLowerCase(); const hits=!q?[]:entries.filter(e=>[e.invoice,e.note,e.vendor,e.fileName].join(' ').toLowerCase().includes(q)); $('#phSearchResults20').innerHTML=hits.map(e=>`<div class="search-hit"><span><strong>${esc(e.invoice||e.fileName||'Eintrag')}</strong><br><span class="hint">${esc(e.date||'')} · ${esc(e.note||e.vendor||'')}</span></span>${e.fileData?`<button class="small secondary" data-open-invoice20="${esc(e.id)}">Ansehen</button>`:''}</div>`).join('') || (q?'<div class="search-hit">Keine passende Rechnung gefunden.</div>':''); bindHistoryButtons20(entries);};
    $('#phSearch20')?.addEventListener('input',runSearch); bindHistoryButtons20(entries);
  }
  function bindHistoryButtons20(entries){
    $$('#modalRoot [data-del-history20]').forEach(b=>b.onclick=()=>{const i=entries.findIndex(e=>e.id===b.dataset.delHistory20); if(i>=0&&confirm('Historieneintrag löschen?')){entries.splice(i,1); closeModal(); save(); render(); toast('Historieneintrag gelöscht.');}});
    $$('#modalRoot [data-open-invoice20]').forEach(b=>b.onclick=()=>{const e=entries.find(x=>x.id===b.dataset.openInvoice20); if(!e)return; const content=e.fileData?(String(e.fileType||'').startsWith('image/')?`<img src="${esc(e.fileData)}" style="max-width:100%;max-height:70vh;object-fit:contain">`:`<iframe class="preview-frame" src="${esc(e.fileData)}"></iframe>`):'<p>Keine Datei hinterlegt.</p>'; showModal('Rechnung / Datei',`<h3>${esc(e.fileName||e.invoice||'Datei')}</h3>${content}`,[{text:'Zurück',secondary:true,action:()=>{closeModal();}}]);});
  }

  const baseBindTopic20=bindTopic;
  bindTopic=function(card,sheet,access){
    baseBindTopic20(card,sheet,access);
    card.addEventListener('contextmenu',e=>{
      const td=e.target.closest('td[data-r]'); if(!td)return; e.preventDefault(); const key=card.dataset.topic; const t=(sheet.themes||[]).find(x=>(x.key||topicKey(x))===key); if(!t)return; const r=+td.dataset.r; const c=+td.dataset.c; const col=t.columns.find(cc=>cc.col===c); openContextMenu20(e.clientX,e.clientY,sheet,t,r,col,td,access);
    });
  };
  function openContextMenu20(x,y,sheet,t,r,col,td,access){
    let menu=document.getElementById('contextMenu20'); if(!menu){menu=document.createElement('div'); menu.id='contextMenu20'; menu.className='context-menu20'; document.body.appendChild(menu); document.addEventListener('click',()=>menu.style.display='none'); window.addEventListener('scroll',()=>menu.style.display='none',true);}
    const canEdit=canAccess(access,'edit'); const isBought=/gekauft/.test(colText20(col)); const isZulauf=/zulauf/.test(colText20(col));
    menu.innerHTML=`${canEdit?'<button data-cm="edit">✎ Zelle bearbeiten</button><button data-cm="rowdel">🗑 Zeile löschen</button><button data-cm="coldel">🧹 Spalte löschen</button>':''}${(isBought||isZulauf)?'<button data-cm="history">🧾 Bestellhistorie</button>':''}${isInstallTopic20(t)?'<button data-cm="resync">🔁 Bestand synchronisieren</button>':''}`;
    menu.style.left=Math.min(x,window.innerWidth-260)+'px'; menu.style.top=Math.min(y,window.innerHeight-220)+'px'; menu.style.display='block';
    menu.onclick=ev=>{const b=ev.target.closest('[data-cm]'); if(!b)return; menu.style.display='none'; const cmd=b.dataset.cm; if(cmd==='edit')openCellEditor(sheet,t,td); if(cmd==='rowdel')deleteRow18?.(sheet,t,r); if(cmd==='coldel')deleteColumn18?.(sheet,t,(t.columns||[]).findIndex(cc=>cc.col===col.col)); if(cmd==='history')openPurchaseHistory20(sheet,t,r); if(cmd==='resync'){syncVehicleAvailability20();save();render();toast('Verfügbare Bestände synchronisiert.');}};
  }

  partSearch=function(sheet){
    const qRaw=$('#partSearch').value.trim(); const q=normPart20(qRaw); const qText=qRaw.toLowerCase(); const out=$('#partResults'); if(!qRaw){out.innerHTML='';return}
    const hits=[];
    for(const t of sheet.themes||[]) for(const col of t.columns||[]){const h=colText20(col); if(/teil|part|pn|oem|sn|nummer|hersteller/.test(h)){for(const r of rows20(t)){const v=String(sheet.cells?.[`${r}:${col.col}`]?.text||''); if((q&&normPart20(v).includes(q)) || v.toLowerCase().includes(qText)) hits.push({theme:t,row:r,col,value:v});}}}
    out.innerHTML=hits.length?hits.slice(0,50).map(h=>`<div class="search-hit"><div><strong>${esc(h.value)}</strong><div class="hint">${esc(h.theme.displayName||h.theme.name)} · ${esc(shortHeader(h.col.parts))}</div></div><button class="small secondary" data-part-hit-r="${h.row}" data-part-hit-c="${h.col.col}" data-part-hit-t="${esc(h.theme.key||topicKey(h.theme))}">Springen</button></div>`).join(''):'<div class="search-hit">Keine Treffer.</div>';
    $$('#partResults [data-part-hit-r]').forEach(b=>b.onclick=()=>{const t=(sheet.themes||[]).find(x=>(x.key||topicKey(x))===b.dataset.partHitT); document.getElementById(topicDomId(sheet,t))?.scrollIntoView({behavior:'smooth'}); setTimeout(()=>{const el=document.querySelector(`#${CSS.escape(topicDomId(sheet,t))} td[data-r="${b.dataset.partHitR}"][data-c="${b.dataset.partHitC}"]`); if(el){el.classList.add('highlight-cell'); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}); setTimeout(()=>el.classList.remove('highlight-cell'),7500);}},250);});
  };
  function jumpToStock20(part){
    const q=normPart20(part); if(!q){toast('Keine Teilenummer hinterlegt.');return;}
    const ss=stockSheet20(), st=stockTheme20(ss); if(!ss||!st){toast('Teilebestand nicht gefunden.');return;}
    const cols=stockCols20(st); let hit=null;
    for(const r of rows20(st)){const val=cellVal20(ss,r,cols.part); const nv=normPart20(val); if(nv&&(nv.includes(q)||q.includes(nv))){hit={sheet:ss,theme:st,row:r,col:cols.part};break;}}
    if(!hit){toast('Teil ist nicht auf Lager bzw. nicht im Teilebestand gefunden.');return;}
    state.activeSheetId=ss.id; save(); render(); setTimeout(()=>{const el=document.querySelector(`#${CSS.escape(topicDomId(ss,st))} td[data-r="${hit.row}"][data-c="${hit.col.col}"]`); if(el){el.classList.add('highlight-cell'); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}); setTimeout(()=>el.classList.remove('highlight-cell'),7500);}},200);
  }
  const baseRenderCatalog20=renderCatalog;
  renderCatalog=function(sheet){
    baseRenderCatalog20(sheet);
    $$('#main [data-stock-row]').forEach(b=>{b.onclick=()=>jumpToStock20(b.dataset.stockRow);});
  };

  const baseRender20=render;
  render=function(){syncVehicleAvailability20(); baseRender20(); setTitle20();};
  const baseReleaseNotes20=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.20 · Bestandsabzug, Bestellhistorie und Rechtsklick-Funktionen',body:'Beim Eintragen von Verbaut bzw. Zahl in Autoreitern erkennt die App Bauteile über Hersteller und normalisierte Teilenummer, fragt vor der Entnahme aus dem Teilebestand nach, verhindert negative Bestände, bucht Differenzen und Rücknahmen, synchronisiert Verfügbar-Werte und ergänzt eine Bestellhistorie per Rechtsklick auf Gekauft/Zulauf mit Durchschnittspreis und Rechnungsdateien.',jump:'settings:daten'}, ...baseReleaseNotes20()];};
})();


/* ---------------- V26.07.22: Teilekatalog Ordner-Vorschau + Hotspots ---------------- */
(function(){
  const VERSION_21='V26.07.22';
  const STORE_KEY_21='kfz_app_'+VERSION_21+'_state';
  const LIVE_KEY_21='kfz_app_live_state_preserved';

  function css21(){
    if(document.getElementById('styleV2207')) return;
    const st=document.createElement('style'); st.id='styleV2207';
    st.textContent=`
      .catalog21-layout{display:grid;grid-template-columns:minmax(270px,380px) 1fr;gap:16px;align-items:start}
      .catalog21-treebox{position:sticky;top:154px;align-self:start;max-height:calc(100vh - 190px);overflow:auto;padding-right:4px}
      .catalog21-preview-card{background:var(--surface);border:1px solid var(--border);border-radius:20px;padding:12px;box-shadow:var(--softShadow);margin-bottom:14px}
      .catalog21-preview-head{display:flex;justify-content:space-between;gap:10px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
      .catalog21-preview-stage{position:relative;min-height:210px;max-height:360px;display:grid;place-items:center;background:linear-gradient(180deg,var(--surface2),var(--surface));border:1px dashed var(--border);border-radius:18px;overflow:hidden}
      .catalog21-preview-stage img{max-width:100%;max-height:360px;object-fit:contain;display:block}
      .catalog21-preview-stage iframe{width:100%;height:340px;border:0;background:#fff;border-radius:12px}
      .catalog21-preview-empty{color:var(--muted);font-weight:900;text-align:center;padding:24px}
      .catalog21-hotspot-layer{position:absolute;inset:0;pointer-events:none}
      .catalog21-hotspot{position:absolute;border:2px solid #f2c94c;background:rgba(242,201,76,.24);border-radius:6px;cursor:pointer;pointer-events:auto;box-shadow:0 0 0 1px rgba(0,0,0,.08)}
      .catalog21-hotspot:hover{background:rgba(242,201,76,.45)}
      .catalog21-hotspot .tag{position:absolute;left:2px;top:2px;font-size:11px;font-weight:900;background:#f2c94c;color:#453000;border-radius:999px;padding:1px 6px;max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .catalog21-draw{position:absolute;border:2px dashed var(--primary);background:color-mix(in srgb,var(--primary) 18%, transparent);pointer-events:none;border-radius:5px;display:none}
      .catalog21-preview-nav{position:absolute;top:50%;transform:translateY(-50%);z-index:4;width:34px;height:44px;border-radius:999px;background:rgba(255,255,255,.84);border:1px solid var(--border);box-shadow:var(--softShadow);font-weight:900;color:var(--primary)}
      body.dark .catalog21-preview-nav{background:rgba(20,26,38,.86);color:var(--text)}
      .catalog21-preview-nav.prev{left:8px}.catalog21-preview-nav.next{right:8px}
      .catalog21-table-wrap{overflow:auto;border:1px solid var(--border);border-radius:16px;background:var(--surface)}
      .catalog21-table .active-file td{background:#fff3b0!important;outline:2px solid #f2c94c;outline-offset:-2px}
      body.dark .catalog21-table .active-file td{background:#5b4c13!important}
      .catalog21-folder-title{display:flex;align-items:center;gap:8px;margin:0 0 10px;color:var(--primary)}
      .catalog21-table-actions{display:flex;gap:6px;flex-wrap:wrap}
      .catalog21-filechip{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--border);border-radius:999px;padding:3px 8px;background:var(--surface2);font-size:12px;font-weight:900}.catalog21-folder-preview-badge{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--border);border-radius:999px;padding:4px 9px;background:var(--surface2);font-size:12px;font-weight:900;color:var(--muted)}
      @media(max-width:980px){.catalog21-layout{grid-template-columns:1fr}.catalog21-treebox{position:relative;top:auto;max-height:none}}
    `;
    document.head.appendChild(st);
  }

  const baseInit21=initState;
  initState=function(){
    let s=null;
    try{s=JSON.parse(localStorage.getItem(STORE_KEY_21)||localStorage.getItem(LIVE_KEY_21)||localStorage.getItem('kfz_app_V26.07.20_state')||'null');}catch(e){s=null;}
    if(!(s&&s.users)) s=baseInit21();
    migrate21(s);
    return s;
  };
  save=function(){
    if(!state) return;
    state.version=VERSION_21; state.savedAt=new Date().toISOString();
    try{localStorage.setItem(STORE_KEY_21,JSON.stringify(state));localStorage.setItem(LIVE_KEY_21,JSON.stringify(state));}catch(e){console.warn(e)}
    updateSaveInfo?.();
  };
  function migrate21(s){
    s.version=VERSION_21;
    s.catalogFolderSelected=s.catalogFolderSelected||{};
    const cat=catalogSheet21(s);
    if(cat){
      cat.vehicleCatalogs=cat.vehicleCatalogs||{};
      for(const b of Object.values(cat.vehicleCatalogs)) prepBundle21(b);
      prepBundle21(cat);
    }
  }
  function catalogSheet21(src=state){return (src.adminSheets||[]).find(s=>s.catalog)||[...(src.adminSheets||[]),...Object.values(src.userSheets||{}).flat()].find(s=>s.catalog)}
  function ensureBrands21(){
    state.catalogBrands=Array.isArray(state.catalogBrands)&&state.catalogBrands.length?state.catalogBrands:[{id:'audi',name:'Audi',logo:'🚗',models:[{id:'a3_8p',name:'A3 8P',logo:'🚙',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]}];
    for(const b of state.catalogBrands){b.models=Array.isArray(b.models)&&b.models.length?b.models:[{id:'allgemein',name:'Allgemein',logo:'🧩',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]; for(const m of b.models){m.motors=Array.isArray(m.motors)&&m.motors.length?m.motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}];}}
    state.catalogSelection=state.catalogSelection||{};
    const b=state.catalogBrands.find(x=>x.id===state.catalogSelection.brand)||state.catalogBrands[0]; state.catalogSelection.brand=b.id;
    const m=b.models.find(x=>x.id===state.catalogSelection.model)||b.models[0]; state.catalogSelection.model=m.id;
    const mot=m.motors.find(x=>x.id===state.catalogSelection.motor)||m.motors[0]; state.catalogSelection.motor=mot.id;
    if(!('vin' in state.catalogSelection)) state.catalogSelection.vin='';
  }
  function brand21(){ensureBrands21(); return state.catalogBrands.find(x=>x.id===state.catalogSelection.brand)||state.catalogBrands[0]}
  function model21(){const b=brand21(); return b.models.find(x=>x.id===state.catalogSelection.model)||b.models[0]}
  function motor21(){const m=model21(); return m.motors.find(x=>x.id===state.catalogSelection.motor)||m.motors[0]}
  function key21(){const s=state.catalogSelection||{}; return [s.brand||'allgemein',s.model||'allgemein',s.motor||'allgemein',String(s.vin||'').trim()||'ohne_vin'].join('__')}
  function bundle21(){
    const cat=catalogSheet21(); ensureBrands21(); cat.vehicleCatalogs=cat.vehicleCatalogs||{}; const k=key21();
    if(!cat.vehicleCatalogs[k]) cat.vehicleCatalogs[k]={tree:clone(cat.tree||[]),rows:clone(cat.rows||[])};
    prepBundle21(cat.vehicleCatalogs[k]); return cat.vehicleCatalogs[k];
  }
  function prepBundle21(b){b.tree=Array.isArray(b.tree)?b.tree:[]; b.rows=Array.isArray(b.rows)?b.rows:[]; walk21(b.tree,n=>{n.id=n.id||uid(n.type||'node'); n.type=n.type||'folder'; if(n.type==='folder'){n.children=Array.isArray(n.children)?n.children:[]; n.icon=n.icon||'📁'; n.previewHotspots=Array.isArray(n.previewHotspots)?n.previewHotspots:[]; if(n.previewFile&&!n.previewFile.id)n.previewFile.id=uid('preview');} else {n.icon=n.icon||'📄'; if(!n.rowId)n.rowId=uid('catrow'); if(!b.rows.some(r=>r.id===n.rowId)) b.rows.push({id:n.rowId,image:n.thumb||'',name:n.name||n.fileName||'Datei',part:n.part||'',category:'',note:n.fileName?('Datei: '+n.fileName):''}); n.hotspots=Array.isArray(n.hotspots)?n.hotspots:[];}});}
  function walk21(nodes,fn){(nodes||[]).forEach(n=>{fn(n); if(n.children)walk21(n.children,fn)})}
  function nodeInfo21(nodes,id,parent=null){for(const n of nodes||[]){if(n.id===id)return {node:n,parent}; const r=nodeInfo21(n.children,id,n); if(r)return r;} return null}
  function firstFolder21(nodes){for(const n of nodes||[]){if(n.type==='folder')return n; const r=firstFolder21(n.children); if(r)return r;} return null}
  function firstFile21(nodes){for(const n of nodes||[]){if(n.type==='file')return n; const r=firstFile21(n.children); if(r)return r;} return null}
  function currentFolder21(b){
    const k=key21(); let id=state.catalogFolderSelected?.[k]; let info=id?nodeInfo21(b.tree,id):null;
    if(info?.node?.type==='folder') return info.node;
    const sel=nodeInfo21(b.tree,state.catalogSelected); if(sel?.node?.type==='file'&&sel.parent){state.catalogFolderSelected=state.catalogFolderSelected||{}; state.catalogFolderSelected[k]=sel.parent.id; return sel.parent;}
    const f=firstFolder21(b.tree); if(f){state.catalogFolderSelected=state.catalogFolderSelected||{}; state.catalogFolderSelected[k]=f.id; return f;}
    return {id:'root',name:'Katalog',type:'folder',open:true,children:b.tree};
  }
  function directFiles21(folder){return (folder.children||[]).filter(n=>n.type==='file')}
  function rowForFile21(b,file){prepBundle21(b); let r=(b.rows||[]).find(x=>x.id===file.rowId); if(!r){r={id:file.rowId||uid('catrow'),image:file.thumb||'',name:file.name||file.fileName||'Datei',part:file.part||'',category:'',note:file.fileName?('Datei: '+file.fileName):''}; file.rowId=r.id; b.rows.push(r);} return r;}
  function rowName21(b,rowId){const r=(b.rows||[]).find(x=>x.id===rowId); return r?.name||r?.part||rowId||'Eintrag'}
  function fileContent21(file){
    if(!file) return '<div class="catalog21-preview-empty">Keine Datei im aktuellen Ordner ausgewählt.</div>';
    const data=file.fileData||file.thumb||''; const type=String(file.fileType||'');
    if(data && (type.startsWith('image/') || String(data).startsWith('data:image'))) return `<img src="${esc(data)}" alt="${esc(file.name||'Vorschau')}">`;
    if(data && (type.includes('pdf') || String(data).startsWith('data:application/pdf'))) return `<iframe src="${esc(data)}"></iframe>`;
    if(file.thumb) return `<img src="${esc(file.thumb)}" alt="${esc(file.name||'Vorschau')}">`;
    return `<div class="catalog21-preview-empty"><div style="font-size:54px">${esc(file.icon||'📄')}</div><strong>${esc(file.name||file.fileName||'Datei')}</strong><br><span class="hint">Für diese Datei ist keine Bildvorschau hinterlegt.</span></div>`;
  }
  function renderCatalogTree21(nodes,depth=0){return (nodes||[]).map(n=>{const isFolder=n.type==='folder'; const selected=(isFolder && state.catalogFolderSelected?.[key21()]===n.id) || (!isFolder && state.catalogSelected===n.id); const icon=n.thumb?`<img class="thumb" src="${esc(n.thumb)}" alt="">`:`<span class="thumb">${esc(n.icon||(isFolder?'📁':'📄'))}</span>`; return `<div class="tree-node" data-cat-node="${esc(n.id)}"><div class="tree-row ${selected?'active-file':''}" style="margin-left:${depth*14}px" data-cat-open><span class="tree-toggle">${isFolder?(n.open?'▾':'▸'):'•'}</span>${icon}<strong>${esc(n.name||'')}</strong>${n.part?`<span class="pill">${esc(n.part)}</span>`:''}<span style="flex:1"></span>${currentUser()?.role==='Admin'?`${isFolder?'<button class="small secondary" data-cat-add-folder>+ Ordner</button><button class="small secondary" data-cat-add-file>+ Datei</button>':''}<button class="small secondary" data-cat-edit>✎</button><button class="small danger" data-cat-del>🗑</button>`:''}</div>${isFolder&&n.open?`<div class="tree-children">${renderCatalogTree21(n.children||[],depth+1)}</div>`:''}</div>`}).join('')}
  function removeNode21(nodes,id){const i=(nodes||[]).findIndex(n=>n.id===id); if(i>=0)return nodes.splice(i,1)[0]; for(const n of nodes||[]){const r=removeNode21(n.children,id); if(r)return r;} return null}
  function editNode21(node,b){
    showModal((node.type==='folder'?'Ordner':'Datei')+' bearbeiten',`<div class="grid2"><div class="field"><label>Name</label><input id="n21Name" value="${esc(node.name||'')}"></div><div class="field"><label>Teilenummer / Bildtext</label><input id="n21Part" value="${esc(node.part||'')}"></div><div class="field"><label>Icon/Text</label><input id="n21Icon" value="${esc(node.icon||'')}"></div><div class="field"><label>Vorschau/Bild</label><input id="n21Img" type="file" accept="image/*"></div><div class="field"><label>Internetlink</label><input id="n21Url" value="${esc(node.url||'')}"></div></div>`,[
      {text:'Speichern',action:()=>{node.name=$('#n21Name').value.trim()||node.name; node.part=$('#n21Part').value.trim(); node.icon=$('#n21Icon').value.trim()||node.icon; node.url=$('#n21Url').value.trim(); const f=$('#n21Img').files[0]; const done=()=>{if(node.type==='file'){const r=rowForFile21(b,node); r.name=node.name; r.part=node.part; r.image=node.thumb||r.image||'';} save(); closeModal(); render();}; if(f){const rd=new FileReader(); rd.onload=()=>{node.thumb=rd.result; done();}; rd.readAsDataURL(f);} else done();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
  }
  function addFolder21(parent){parent.children=parent.children||[]; parent.open=true; parent.children.push({id:uid('folder'),type:'folder',name:'Neuer Ordner',open:true,children:[],icon:'📁'}); save(); render()}
  function addFile21(parent,b){
    const input=document.createElement('input'); input.type='file'; input.accept='image/*,.pdf,.txt,.doc,.docx,.xls,.xlsx,.xlsm';
    input.onchange=()=>{const f=input.files&&input.files[0]; if(!f)return; const rd=new FileReader(); rd.onload=()=>{parent.children=parent.children||[]; parent.open=true; const isImg=String(f.type||'').startsWith('image/'); const file={id:uid('file'),type:'file',name:f.name,fileName:f.name,fileType:f.type||'',fileSize:f.size||0,fileData:rd.result,thumb:isImg?rd.result:'',icon:isImg?'🖼️':'📄',part:'',url:'',rowId:uid('catrow'),hotspots:[]}; parent.children.push(file); b.rows=b.rows||[]; b.rows.push({id:file.rowId,image:file.thumb||'',name:file.name,part:'',category:'',note:'Datei: '+file.fileName}); state.catalogSelected=file.id; state.catalogFolderSelected=state.catalogFolderSelected||{}; state.catalogFolderSelected[key21()]=parent.id; save(); render(); setTimeout(()=>editNode21(file,b),80);}; rd.readAsDataURL(f);};
    input.click();
  }
  function bindCatalogTree21(b){
    const el=$('#catalogTree21'); if(!el)return;
    el.onclick=e=>{const row=e.target.closest('[data-cat-node]'); if(!row)return; const info=nodeInfo21(b.tree,row.dataset.catNode); if(!info)return; const n=info.node;
      if(e.target.closest('[data-cat-add-folder]')) return addFolder21(n);
      if(e.target.closest('[data-cat-add-file]')) return addFile21(n,b);
      if(e.target.closest('[data-cat-edit]')) return editNode21(n,b);
      if(e.target.closest('[data-cat-del]')){ if(confirm('Wirklich löschen?')){const removed=removeNode21(b.tree,n.id); if(removed?.rowId)b.rows=(b.rows||[]).filter(r=>r.id!==removed.rowId); save(); render();} return; }
      if(n.type==='folder'){n.open=!n.open; state.catalogFolderSelected=state.catalogFolderSelected||{}; state.catalogFolderSelected[key21()]=n.id; const f=firstFile21(n.children||[]); if(f)state.catalogSelected=f.id; save(); render(); return;}
      state.catalogSelected=n.id; if(info.parent){state.catalogFolderSelected=state.catalogFolderSelected||{}; state.catalogFolderSelected[key21()]=info.parent.id;} save(); render(); setTimeout(()=>highlight21(`[data-catalog-row="${CSS.escape(n.rowId||'')}" ]`),80);
    };
  }
  function selectedFile21(b,folder){let info=state.catalogSelected?nodeInfo21(b.tree,state.catalogSelected):null; if(!(info?.node?.type==='file') || (info.parent&&info.parent.id!==folder.id)){const f=directFiles21(folder)[0]||firstFile21(folder.children||[]); if(f){state.catalogSelected=f.id; info=nodeInfo21(b.tree,f.id);} }
    return info?.node?.type==='file'?info.node:null;
  }
  function highlight21(selector){const el=document.querySelector(selector); if(!el)return; el.classList.add('highlight-row'); el.scrollIntoView({behavior:'smooth',block:'center'}); setTimeout(()=>el.classList.remove('highlight-row'),6500)}
  function addPreviewFile21(folder){
    const input=document.createElement('input'); input.type='file'; input.accept='image/*,.pdf';
    input.onchange=()=>{const f=input.files&&input.files[0]; if(!f)return; const rd=new FileReader(); rd.onload=()=>{const isImg=String(f.type||'').startsWith('image/'); folder.previewFile={id:uid('preview'),type:'preview',name:f.name,fileName:f.name,fileType:f.type||'',fileSize:f.size||0,fileData:rd.result,thumb:isImg?rd.result:'',icon:isImg?'🖼️':'📄'}; folder.previewHotspots=Array.isArray(folder.previewHotspots)?folder.previewHotspots:[]; save(); render(); toast('Vorschaudatei für diesen Ordner gespeichert.');}; rd.readAsDataURL(f);};
    input.click();
  }
  function bindPreview21(b,folder,previewFile,files){
    $$('#main [data-hotspot-row]').forEach(h=>h.onclick=e=>{e.stopPropagation(); const row=h.dataset.hotspotRow; const tr=document.querySelector(`[data-catalog-row="${CSS.escape(row)}"]`); if(tr){tr.classList.add('highlight-row'); tr.scrollIntoView({behavior:'smooth',block:'center'}); setTimeout(()=>tr.classList.remove('highlight-row'),7000);}});
    $$('#main [data-hotspot-row]').forEach(h=>h.oncontextmenu=e=>{if(currentUser()?.role!=='Admin')return; e.preventDefault(); if(confirm('Hotspot löschen?')){folder.previewHotspots=(folder.previewHotspots||[]).filter(x=>x.id!==h.dataset.hotspotId); save(); render();}});
    $('#addHotspot21')?.addEventListener('click',()=>startHotspotDraw21(b,folder));
    $('#catAddFileToFolder21')?.addEventListener('click',()=>addFile21(folder,b));
    $('#catSetFolderPreview21')?.addEventListener('click',()=>addPreviewFile21(folder));
    $('#catClearFolderPreview21')?.addEventListener('click',()=>{if(confirm('Vorschaudatei dieses Ordners entfernen?')){folder.previewFile=null; folder.previewHotspots=[]; save(); render();}});
  }
  function startHotspotDraw21(b,folder){
    if(!folder.previewFile){toast('Bitte zuerst eine Vorschaudatei für diesen Ordner wählen.');return;}
    const files=directFiles21(folder); if(!files.length){toast('Dieser Ordner enthält noch keine Teiledateien zum Verknüpfen.');return;}
    const stage=$('#catalogPreviewStage21'); const draw=$('#drawBox21'); if(!stage||!draw)return;
    toast('Bereich auf der Ordner-Vorschau ziehen. Danach Datei/Eintrag verknüpfen.');
    stage.classList.add('drawing'); let start=null;
    const rect=()=>stage.getBoundingClientRect();
    function pos(ev){const r=rect(); return {x:Math.max(0,Math.min(r.width,ev.clientX-r.left)),y:Math.max(0,Math.min(r.height,ev.clientY-r.top)),w:r.width,h:r.height};}
    function down(ev){if(ev.target.closest('.catalog21-hotspot')||ev.target.closest('button'))return; ev.preventDefault(); start=pos(ev); Object.assign(draw.style,{display:'block',left:start.x+'px',top:start.y+'px',width:'0px',height:'0px'}); stage.addEventListener('pointermove',move); stage.addEventListener('pointerup',up,{once:true});}
    function move(ev){if(!start)return; const p=pos(ev); const x=Math.min(start.x,p.x),y=Math.min(start.y,p.y),w=Math.abs(p.x-start.x),h=Math.abs(p.y-start.y); Object.assign(draw.style,{left:x+'px',top:y+'px',width:w+'px',height:h+'px'});}
    function up(ev){stage.removeEventListener('pointermove',move); stage.removeEventListener('pointerdown',down); draw.style.display='none'; const p=pos(ev); const x=Math.min(start.x,p.x),y=Math.min(start.y,p.y),w=Math.abs(p.x-start.x),h=Math.abs(p.y-start.y); start=null; if(w<12||h<12){toast('Bereich zu klein.'); return render();} const rowOptions=directFiles21(folder).map(f=>{const r=rowForFile21(b,f); return `<option value="${esc(r.id)}">${esc(r.name||f.name||r.part)}</option>`}).join(''); showModal('Hotspot verknüpfen',`<p>Welcher Teileeintrag soll gelb aufleuchten, wenn dieser Bereich in der Ordner-Vorschau angeklickt wird?</p><div class="field"><label>Datei / Teileeintrag</label><select id="hotspotRow21">${rowOptions}</select></div>`,[
      {text:'Speichern',action:()=>{folder.previewHotspots=folder.previewHotspots||[]; folder.previewHotspots.push({id:uid('hot'),x:x/p.w*100,y:y/p.h*100,w:w/p.w*100,h:h/p.h*100,rowId:$('#hotspotRow21').value}); closeModal(); save(); render();}},
      {text:'Abbrechen',secondary:true,action:()=>{closeModal();render();}}
    ]);}
    stage.addEventListener('pointerdown',down);
  }
  function renderChooser21(isAdmin){const b=brand21(), m=model21(), mot=motor21(); return `<section class="catalog-chooser"><div class="section-head"><h3>Fahrzeugauswahl</h3>${isAdmin?'<div><button class="small secondary" id="catAddBrand21">+ Marke</button> <button class="small secondary" id="catAddModel21">+ Modell</button> <button class="small secondary" id="catAddMotor21">+ Motor</button></div>':''}</div><div class="catalog-chooser-grid"><div class="field"><label>Marke</label><select id="catBrand21">${state.catalogBrands.map(x=>`<option value="${esc(x.id)}" ${x.id===b.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>Modell</label><select id="catModel21">${b.models.map(x=>`<option value="${esc(x.id)}" ${x.id===m.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>Motor</label><select id="catMotor21">${m.motors.map(x=>`<option value="${esc(x.id)}" ${x.id===mot.id?'selected':''}>${esc(x.logo||'')} ${esc(x.name)}</option>`).join('')}</select></div><div class="field"><label>VIN optional</label><input id="catVin21" value="${esc(state.catalogSelection.vin||'')}" placeholder="optional"></div></div></section>`}
  function bindChooser21(){
    $('#catBrand21')?.addEventListener('change',e=>{state.catalogSelection.brand=e.target.value; const b=brand21(); state.catalogSelection.model=b.models[0]?.id||''; state.catalogSelection.motor=b.models[0]?.motors?.[0]?.id||''; state.catalogSelected=''; save(); render();});
    $('#catModel21')?.addEventListener('change',e=>{state.catalogSelection.model=e.target.value; const m=model21(); state.catalogSelection.motor=m.motors[0]?.id||''; state.catalogSelected=''; save(); render();});
    $('#catMotor21')?.addEventListener('change',e=>{state.catalogSelection.motor=e.target.value; state.catalogSelected=''; save(); render();});
    $('#catVin21')?.addEventListener('blur',e=>{state.catalogSelection.vin=e.target.value.trim(); state.catalogSelected=''; save(); render();});
    $('#catAddBrand21')?.addEventListener('click',()=>{const name=prompt('Neue Marke?'); if(!name)return; const id=slug(name)+'_'+uid('b'); state.catalogBrands.push({id,name,logo:'🚗',models:[{id:'allgemein_'+uid('m'),name:'Allgemein',logo:'🧩',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}]}); state.catalogSelection={brand:id,model:state.catalogBrands.at(-1).models[0].id,motor:'allgemein',vin:''}; save(); render();});
    $('#catAddModel21')?.addEventListener('click',()=>{const b=brand21(); const name=prompt('Neues Modell?'); if(!name)return; const id=slug(name)+'_'+uid('m'); b.models.push({id,name,logo:'🚙',motors:[{id:'allgemein',name:'Allgemein',logo:'⚙️'}]}); state.catalogSelection.model=id; state.catalogSelection.motor='allgemein'; save(); render();});
    $('#catAddMotor21')?.addEventListener('click',()=>{const m=model21(); const name=prompt('Neuer Motor?'); if(!name)return; const id=slug(name)+'_'+uid('mot'); m.motors.push({id,name,logo:'⚙️'}); state.catalogSelection.motor=id; save(); render();});
  }
  function normPart21(v){return String(v||'').toLowerCase().replace(/[\s\.\-_\/\\]+/g,'').replace(/[^a-z0-9äöüß]/gi,'');}
  function visibleSheets21(){return allSheetsForUser?.() || [...(state.adminSheets||[]),...Object.values(state.userSheets||{}).flat()];}
  function colText21(col){return (col?.parts||[col?.label||'']).join(' ').toLowerCase();}
  function rowsOf21(t){return (t?.rowGroups||[]).map(g=>Array.isArray(g)?g[0]:g).filter(Boolean);}
  function jumpToStock21(part){
    const q=normPart21(part); if(!q){toast('Keine Teilenummer hinterlegt.');return;}
    const sheets=visibleSheets21(); const sh=sheets.find(s=>/^Teilebestand$/i.test(s.name||'')) || sheets.find(s=>/Teilebestand/i.test(s.name||''));
    if(!sh){toast('Teilebestand nicht gefunden.');return;}
    let hit=null;
    for(const t of sh.themes||[]){
      const partCols=(t.columns||[]).filter(c=>/teilenummer|part|pn|oem|sn/.test(colText21(c)));
      for(const c of partCols){
        for(const r of rowsOf21(t)){const v=String(sh.cells?.[`${r}:${c.col}`]?.text||sh.cells?.[`${r}:${c.col}`]?.raw||''); const nv=normPart21(v); if(nv&&(nv.includes(q)||q.includes(nv))){hit={sheet:sh,theme:t,row:r,col:c};break;}}
        if(hit)break;
      }
      if(hit)break;
    }
    if(!hit){toast('Teil ist nicht auf Lager bzw. nicht im Teilebestand gefunden.');return;}
    state.activeSheetId=hit.sheet.id; save(); render();
    setTimeout(()=>{const id=topicDomId(hit.sheet,hit.theme); const el=document.querySelector(`#${CSS.escape(id)} td[data-r="${hit.row}"][data-c="${hit.col.col}"]`); if(el){el.classList.add('highlight-cell'); el.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}); setTimeout(()=>el.classList.remove('highlight-cell'),7500);}},240);
  }

  renderCatalog=function(sheet){
    css21(); ensureBrands21(); const isAdmin=currentUser()?.role==='Admin'; const cat=catalogSheet21(); const b=bundle21(); prepBundle21(b); const folder=currentFolder21(b); folder.previewHotspots=Array.isArray(folder.previewHotspots)?folder.previewHotspots:[]; const files=directFiles21(folder); const rows=files.map(f=>({file:f,row:rowForFile21(b,f)})); const previewFile=folder.previewFile||null;
    const hotspots=(folder.previewHotspots||[]).map(h=>`<button class="catalog21-hotspot" data-hotspot-id="${esc(h.id)}" data-hotspot-row="${esc(h.rowId)}" style="left:${h.x}%;top:${h.y}%;width:${h.w}%;height:${h.h}%"><span class="tag">${esc(rowName21(b,h.rowId))}</span></button>`).join('');
    $('#main').innerHTML=`<div class="sheet-top"><div class="sheet-title"><h2>Teilekatalog</h2><div class="sheet-sub"><span class="pill ok">Global für alle Nutzer</span><span class="pill">${isAdmin?'Admin bearbeitet':'Nur ansehen'}</span></div></div><div class="sheet-tools">${isAdmin?'<button id="catAddRoot21" class="secondary">+ Hauptordner</button> <button id="catPublish21" class="secondary">Katalog veröffentlichen</button>':''}</div></div>${renderChooser21(isAdmin)}<section class="catalog-section"><div class="catalog21-layout"><aside class="catalog21-treebox"><div class="section-head"><h3>Ordner</h3></div><div class="tree vertical-tree" id="catalogTree21">${renderCatalogTree21(b.tree)}</div></aside><div><div class="catalog21-preview-card"><div class="catalog21-preview-head"><div><h3 style="margin:0">Ordner-Vorschau · ${esc(folder.name||'Ordner')}</h3><span class="hint">Die Vorschaudatei gehört zum Ordner. Die Teiledateien stehen darunter und werden per Hotspot verknüpft.</span></div><div class="catalog21-table-actions">${previewFile?`<span class="catalog21-folder-preview-badge">${esc(previewFile.icon||'🖼️')} ${esc(previewFile.name||previewFile.fileName||'Vorschau')}</span>`:''}${isAdmin?'<button class="small secondary" id="catSetFolderPreview21">Vorschaudatei wählen</button><button class="small secondary" id="addHotspot21">Bereich markieren</button>'+ (previewFile?'<button class="small danger" id="catClearFolderPreview21">Vorschau löschen</button>':''):''}</div></div><div class="catalog21-preview-stage" id="catalogPreviewStage21">${fileContent21(previewFile)}<div class="catalog21-hotspot-layer">${hotspots}</div><div id="drawBox21" class="catalog21-draw"></div></div></div><div class="catalog21-preview-card"><h3 class="catalog21-folder-title">${esc(folder.icon||'📁')} ${esc(folder.name||'Aktueller Ordner')} <span class="catalog21-filechip">${rows.length} Teiledatei(en)</span>${isAdmin?'<button class="small secondary" id="catAddFileToFolder21">+ Datei</button>':''}</h3><div class="catalog21-table-wrap"><table class="catalog-table catalog21-table"><thead><tr><th>Vorschau</th><th>Datei / Bildtext</th><th>Teilenummer</th><th>Kategorie</th><th>Notiz</th><th>Aktionen</th></tr></thead><tbody>${rows.length?rows.map(({file:f,row:r})=>`<tr data-catalog-row="${esc(r.id)}" class="${state.catalogSelected===f.id?'active-file':''}"><td>${(f.thumb||r.image)?`<img class="catalog-preview" src="${esc(f.thumb||r.image)}">`:`<span class="thumb">${esc(f.icon||'📄')}</span>`}</td><td><span class="jump-link" data-select-file21="${esc(f.id)}">${esc(r.name||f.name)}</span></td><td>${esc(r.part||f.part||'')}</td><td>${esc(r.category||'')}</td><td>${esc(r.note||'')}</td><td><button class="small secondary" data-stock21="${esc(r.part||f.part||'')}">📦</button>${f.url?`<button class="small secondary" data-web21="${esc(f.url)}">🌐</button>`:''}${isAdmin?` <button class="small secondary" data-edit-file21="${esc(f.id)}">✎</button><button class="small danger" data-del-file21="${esc(f.id)}">🗑</button>`:''}</td></tr>`).join(''):`<tr><td colspan="6"><span class="hint">Dieser Ordner enthält noch keine Teiledateien. Die Ordner-Vorschau ist davon getrennt.</span></td></tr>`}</tbody></table></div></div></div></div></section>`;
    bindChooser21(); bindCatalogTree21(b); bindPreview21(b,folder,previewFile,files);
    $('#catAddRoot21')?.addEventListener('click',()=>{b.tree.push({id:uid('folder'),type:'folder',name:'Neuer Hauptordner',open:true,children:[],icon:'📁',previewFile:null,previewHotspots:[]}); save(); render();});
    $('#catPublish21')?.addEventListener('click',()=>{state.publishedAt=new Date().toISOString(); save(); toast('Teilekatalog veröffentlicht.');});
    $$('#main [data-select-file21]').forEach(el=>el.onclick=()=>{state.catalogSelected=el.dataset.selectFile21; save(); render(); setTimeout(()=>highlight21(`[data-catalog-row="${CSS.escape(nodeInfo21(b.tree,state.catalogSelected)?.node?.rowId||'')}"]`),80);});
    $$('#main [data-edit-file21]').forEach(el=>el.onclick=()=>{const n=nodeInfo21(b.tree,el.dataset.editFile21)?.node; if(n)editNode21(n,b);});
    $$('#main [data-del-file21]').forEach(el=>el.onclick=()=>{if(confirm('Datei wirklich löschen?')){const removed=removeNode21(b.tree,el.dataset.delFile21); if(removed?.rowId)b.rows=(b.rows||[]).filter(r=>r.id!==removed.rowId); save(); render();}});
    $$('#main [data-stock21]').forEach(el=>el.onclick=()=>{jumpToStock21(el.dataset.stock21);});
    $$('#main [data-web21]').forEach(el=>el.onclick=()=>{window.open(el.dataset.web21,'_blank','noopener');});
  };

  const baseRender21=render;
  render=function(){baseRender21(); css21(); document.title='KFZ-APP '+VERSION_21; const chip=document.querySelector('.version-chip'); if(chip) chip.textContent=VERSION_21;};
  const baseReleaseNotes21=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.22 · Ordner-Vorschau im Teilekatalog',body:'Im Teilekatalog gehört die Vorschaudatei jetzt zum aktuell geöffneten Ordner. Die Teiledateien bleiben darunter als eigene Dateitabelle des Ordners. Admins können eine Ordner-Vorschau wählen und Hotspots auf dieser Vorschau mit Teiledateien aus demselben Ordner verknüpfen.',jump:'catalog'}, ...baseReleaseNotes21()];};
})();


/* ---------------- V26.07.23: eindeutige Teilekatalog-Markierung + Logo ohne weißen Rahmen ---------------- */
(function(){
  const VERSION_23='V26.07.23';
  const STORE_KEY_23='kfz_app_'+VERSION_23+'_state';
  const LIVE_KEY_23='kfz_app_live_state_preserved';

  function css23(){
    if(document.getElementById('styleV2307')) return;
    const st=document.createElement('style');
    st.id='styleV2307';
    st.textContent=`
      .brand img#appLogo,.brand img{background:transparent!important;padding:0!important;border:0!important;box-shadow:none!important;border-radius:10px!important;}
      .catalog21-table .highlight-row td,.catalog21-table tr.highlight-row td{background:#fff3b0!important;outline:2px solid #f2c94c;outline-offset:-2px;}
      body.dark .catalog21-table .highlight-row td,body.dark .catalog21-table tr.highlight-row td{background:#5b4c13!important;}
      .catalog21-hotspot.is-active{background:rgba(242,201,76,.56)!important;box-shadow:0 0 0 3px rgba(242,201,76,.35)!important;}
    `;
    document.head.appendChild(st);
  }

  const baseInit23=initState;
  initState=function(){
    let s=null;
    try{s=JSON.parse(localStorage.getItem(STORE_KEY_23)||localStorage.getItem(LIVE_KEY_23)||localStorage.getItem('kfz_app_V26.07.22_state')||localStorage.getItem('kfz_app_V26.07.20_state')||'null');}catch(e){s=null;}
    if(!(s&&s.users)) s=baseInit23();
    s.version=VERSION_23;
    return s;
  };

  const baseSave23=save;
  save=function(){
    if(!state) return;
    state.version=VERSION_23; state.savedAt=new Date().toISOString();
    try{localStorage.setItem(STORE_KEY_23,JSON.stringify(state));localStorage.setItem(LIVE_KEY_23,JSON.stringify(state));}catch(e){console.warn(e)}
    updateSaveInfo?.();
  };

  function clearCatalogMarks23(opts={}){
    const root=document.querySelector('#main');
    if(!root) return;
    root.querySelectorAll('.highlight-row,.highlight-cell,.active-file,.is-active').forEach(el=>{
      el.classList.remove('highlight-row','highlight-cell','active-file','is-active');
    });
    if(opts.clearSelected && state){state.catalogSelected=''; try{save()}catch(e){}}
  }

  function highlightCatalogRow23(rowId){
    clearCatalogMarks23({clearSelected:false});
    if(!rowId) return;
    const row=document.querySelector(`[data-catalog-row="${CSS.escape(rowId)}"]`);
    if(row){row.classList.add('highlight-row','active-file'); row.scrollIntoView({behavior:'smooth',block:'center'});}
    document.querySelectorAll(`[data-hotspot-row="${CSS.escape(rowId)}"]`).forEach(h=>h.classList.add('is-active'));
  }

  function bindCatalogMarking23(){
    css23();
    const main=document.querySelector('#main');
    const section=main?.querySelector('.catalog-section');
    if(!section || section.dataset.v23MarkingBound==='1') return;
    section.dataset.v23MarkingBound='1';

    // Vor jeder neuen gezielten Auswahl alte Markierungen entfernen, damit nie zwei Ziele gelb bleiben.
    section.addEventListener('click',ev=>{
      const hotspot=ev.target.closest('[data-hotspot-row]');
      const fileLink=ev.target.closest('[data-select-file21]');
      const stockLink=ev.target.closest('[data-stock21]');
      const treeRow=ev.target.closest('[data-cat-open]');
      const buttonOrInput=ev.target.closest('button,input,select,textarea,label');

      if(hotspot){
        ev.stopPropagation();
        highlightCatalogRow23(hotspot.dataset.hotspotRow||'');
        return;
      }

      if(fileLink){
        clearCatalogMarks23({clearSelected:false});
        const id=fileLink.dataset.selectFile21;
        setTimeout(()=>{
          const node=document.querySelector(`[data-select-file21="${CSS.escape(id)}"]`)?.closest('tr');
          const rowId=node?.dataset.catalogRow;
          if(rowId) highlightCatalogRow23(rowId);
        },0);
        return;
      }

      if(stockLink){
        clearCatalogMarks23({clearSelected:false});
        return;
      }

      // Klick auf Ordner oder leere Fläche: alte Datei-/Hotspot-Markierung löschen.
      if(treeRow){
        const toggle=treeRow.querySelector('.tree-toggle')?.textContent?.trim();
        const isFolder=toggle && toggle!=='•';
        if(isFolder){
          clearCatalogMarks23({clearSelected:true});
          setTimeout(()=>{state.catalogSelected=''; clearCatalogMarks23({clearSelected:false});},0);
        }
        return;
      }

      // Admin-Buttons/Inputs sollen keine Auswahl markieren, aber auch nichts doppelt stehen lassen.
      if(buttonOrInput){
        clearCatalogMarks23({clearSelected:false});
        return;
      }

      // Klick ins Leere / auf sonstige Katalogfläche löscht sichtbare Markierungen und aktive Datei.
      clearCatalogMarks23({clearSelected:true});
    },true);
  }

  const baseRenderCatalog23=renderCatalog;
  renderCatalog=function(sheet){
    baseRenderCatalog23(sheet);
    bindCatalogMarking23();
  };

  const baseRender23=render;
  render=function(){
    baseRender23();
    css23();
    document.title='KFZ-APP '+VERSION_23;
    const chip=document.querySelector('.version-chip'); if(chip) chip.textContent=VERSION_23;
  };

  const baseApplyTheme23=applyTheme;
  applyTheme=function(){
    baseApplyTheme23();
    css23();
  };

  const baseReleaseNotes23=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.23 · eindeutige Katalog-Markierung und Logo ohne Rahmen',body:'Im Teilekatalog verschwindet eine Markierung jetzt sofort, wenn in einen anderen Bereich oder ins Leere geklickt wird. Es bleibt immer nur eine Datei/Zeile aktiv markiert. Außerdem hat das App-Logo oben links keinen weißen Hintergrundrahmen und kein Padding mehr.',jump:'catalog'}, ...baseReleaseNotes23()];};
})();


/* ---------------- V26.07.24: Hotspot-Bereiche ohne sichtbare Benennung ---------------- */
(function(){
  const VERSION_24='V26.07.24';

  function css24(){
    let st=document.getElementById('v24-style');
    if(!st){st=document.createElement('style'); st.id='v24-style'; document.head.appendChild(st);}
    st.textContent=`
      /* Hotspot-Rechtecke bleiben sichtbar und klickbar, aber ohne eingeblendeten Namen/Text. */
      .catalog21-hotspot .tag{display:none!important;}
      .catalog21-hotspot{font-size:0!important;line-height:0!important;}
    `;
  }

  function stripHotspotLabels24(){
    document.querySelectorAll('.catalog21-hotspot .tag').forEach(el=>el.remove());
    document.querySelectorAll('.catalog21-hotspot').forEach(el=>{
      el.removeAttribute('title');
      el.setAttribute('aria-label','Hotspot');
    });
  }

  const baseRenderCatalog24=renderCatalog;
  renderCatalog=function(sheet){
    baseRenderCatalog24(sheet);
    css24();
    stripHotspotLabels24();
  };

  const baseRender24=render;
  render=function(){
    baseRender24();
    css24();
    stripHotspotLabels24();
    document.title='KFZ-APP '+VERSION_24;
    const chip=document.querySelector('.version-chip'); if(chip) chip.textContent=VERSION_24;
  };

  const baseApplyTheme24=applyTheme;
  applyTheme=function(){
    baseApplyTheme24();
    css24();
  };

  const baseReleaseNotes24=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.24 · Hotspots ohne sichtbare Beschriftung',body:'Die markierten Bereiche in der Ordner-Vorschau des Teilekatalogs bleiben sichtbar und anklickbar, zeigen aber keine Benennung mehr im Rechteck an.',jump:'catalog'}, ...baseReleaseNotes24()];};
})();


/* ---------------- V26.07.26: Hotspots verschieben/bearbeiten + dezente/Admin-wählbare Optik ---------------- */
(function(){
  const VERSION_25='V26.07.26';
  const STORE_KEY_25='kfz_app_'+VERSION_25+'_state';
  const LIVE_KEY_25='kfz_app_live_state_preserved';

  function css25(){
    let st=document.getElementById('v25-style');
    if(!st){st=document.createElement('style'); st.id='v25-style'; document.head.appendChild(st);}
    st.textContent=`
      :root{
        --hotspot-bg: rgba(242,201,76,.14);
        --hotspot-border: rgba(242,201,76,.58);
        --hotspot-border-width: 1px;
        --hotspot-active-bg: rgba(242,201,76,.30);
        --hotspot-active-border: rgba(242,201,76,.86);
      }
      .catalog21-hotspot{
        background:var(--hotspot-bg)!important;
        border:var(--hotspot-border-width) solid var(--hotspot-border)!important;
        box-shadow:none!important;
        outline:0!important;
        border-radius:7px!important;
        transition:background .12s ease,border-color .12s ease,box-shadow .12s ease;
      }
      .catalog21-hotspot:hover{background:var(--hotspot-active-bg)!important;border-color:var(--hotspot-active-border)!important;}
      .catalog21-hotspot.is-active{background:var(--hotspot-active-bg)!important;border-color:var(--hotspot-active-border)!important;box-shadow:0 0 0 1px var(--hotspot-active-border)!important;}
      .catalog21-hotspot .tag{display:none!important;}
      .catalog21-hotspot.is-admin-editable{cursor:move;}
      .catalog21-hotspot.is-admin-editable::after{
        content:'✥';
        position:absolute;
        left:5px;
        top:4px;
        font-size:10px;
        line-height:1;
        color:var(--hotspot-active-border);
        opacity:.55;
        pointer-events:none;
      }
      .hotspot-resize-handle{
        position:absolute;
        right:-4px;
        bottom:-4px;
        width:9px;
        height:9px;
        border-radius:999px;
        background:var(--hotspot-active-border);
        opacity:.62;
        cursor:nwse-resize;
        border:1px solid rgba(255,255,255,.65);
      }
      .catalog21-hotspot:not(.is-admin-editable) .hotspot-resize-handle{display:none!important;}
      .catalog21-hotspot.is-admin-editable:hover .hotspot-resize-handle{opacity:.95;}
      .hotspot-style-button{margin-left:6px;}
      .hotspot-edit-hint{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);margin-left:8px;}
      body.dark .hotspot-resize-handle{border-color:rgba(20,24,35,.8);}
    `;
  }

  function hexToRgba25(hex,alpha){
    hex=String(hex||'#f2c94c').trim();
    if(/^rgba?\(/i.test(hex)) return hex;
    if(hex[0]==='#') hex=hex.slice(1);
    if(hex.length===3) hex=hex.split('').map(c=>c+c).join('');
    const n=parseInt(hex,16);
    if(Number.isNaN(n)) return `rgba(242,201,76,${alpha})`;
    const r=(n>>16)&255,g=(n>>8)&255,b=n&255;
    return `rgba(${r},${g},${b},${alpha})`;
  }
  function ensureHotspotStyle25(){
    state.catalogHotspotStyle=state.catalogHotspotStyle||{};
    const hs=state.catalogHotspotStyle;
    if(hs.fillColor==null) hs.fillColor='#f2c94c';
    if(hs.borderColor==null) hs.borderColor='#f2c94c';
    if(hs.opacity==null) hs.opacity=0.14;
    if(hs.activeOpacity==null) hs.activeOpacity=0.30;
    if(hs.borderWidth==null) hs.borderWidth=1;
    return hs;
  }
  function applyHotspotStyle25(){
    if(!state) return;
    css25();
    const hs=ensureHotspotStyle25();
    const root=document.documentElement;
    root.style.setProperty('--hotspot-bg',hexToRgba25(hs.fillColor,Number(hs.opacity)||0.14));
    root.style.setProperty('--hotspot-border',hexToRgba25(hs.borderColor,Math.min(1,(Number(hs.opacity)||0.14)+0.32)));
    root.style.setProperty('--hotspot-active-bg',hexToRgba25(hs.fillColor,Number(hs.activeOpacity)||0.30));
    root.style.setProperty('--hotspot-active-border',hexToRgba25(hs.borderColor,0.86));
    root.style.setProperty('--hotspot-border-width',(Number(hs.borderWidth)||1)+'px');
  }

  function catalogSheet25(src=state){return (src.adminSheets||[]).find(s=>s.catalog)||[...(src.adminSheets||[]),...Object.values(src.userSheets||{}).flat()].find(s=>s.catalog)}
  function brandKey25(){
    const s=state.catalogSelection||{};
    return [s.brand||'allgemein',s.model||'allgemein',s.motor||'allgemein',String(s.vin||'').trim()||'ohne_vin'].join('__');
  }
  function bundle25(){
    const cat=catalogSheet25(); if(!cat) return null;
    cat.vehicleCatalogs=cat.vehicleCatalogs||{};
    return cat.vehicleCatalogs[brandKey25()]||cat;
  }
  function walk25(nodes,fn,parent=null){(nodes||[]).forEach(n=>{fn(n,parent); if(n.children)walk25(n.children,fn,n);});}
  function nodeInfo25(nodes,id,parent=null){for(const n of nodes||[]){if(n.id===id)return {node:n,parent}; const r=nodeInfo25(n.children,id,n); if(r)return r;} return null;}
  function currentFolder25(b){
    if(!b) return null;
    const id=state.catalogFolderSelected?.[brandKey25()];
    const f=id?nodeInfo25(b.tree,id):null;
    if(f?.node?.type==='folder') return f.node;
    let first=null; walk25(b.tree,n=>{if(!first&&n.type==='folder') first=n;});
    return first;
  }
  function directFiles25(folder){return (folder?.children||[]).filter(n=>n.type==='file');}
  function rowForFile25(b,file){return (b?.rows||[]).find(r=>r.id===file?.rowId)||{id:file?.rowId||'',name:file?.name||'',part:file?.part||''};}
  function rowName25(b,rowId){const r=(b?.rows||[]).find(x=>x.id===rowId); return r?.name||r?.part||rowId||'Eintrag';}
  function hotspotById25(folder,id){return (folder?.previewHotspots||[]).find(h=>h.id===id);}

  function enhanceHotspots25(){
    const isAdmin=currentUser?.()?.role==='Admin';
    document.querySelectorAll('.catalog21-hotspot').forEach(h=>{
      h.querySelectorAll('.tag').forEach(x=>x.remove());
      h.setAttribute('aria-label','Hotspot');
      h.title=isAdmin?'Hotspot verschieben; Doppelklick/Rechtsklick zum Bearbeiten':'';
      h.classList.toggle('is-admin-editable',!!isAdmin);
      if(isAdmin && !h.querySelector('.hotspot-resize-handle')) h.insertAdjacentHTML('beforeend','<span class="hotspot-resize-handle" aria-hidden="true"></span>');
    });
  }

  function insertHotspotStyleButton25(){
    if(currentUser?.()?.role!=='Admin') return;
    const actions=document.querySelector('.catalog21-preview-head .catalog21-table-actions');
    if(!actions || document.getElementById('hotspotStyle25')) return;
    const btn=document.createElement('button');
    btn.id='hotspotStyle25'; btn.type='button'; btn.className='small secondary hotspot-style-button'; btn.textContent='Hotspot-Optik';
    btn.addEventListener('click',openHotspotStyle25);
    actions.appendChild(btn);
    const hint=document.createElement('span');
    hint.className='hotspot-edit-hint';
    hint.textContent='Hotspots ziehen, Ecke zum Vergrößern, Doppelklick zum Bearbeiten';
    actions.appendChild(hint);
  }

  function openHotspotStyle25(){
    const hs=ensureHotspotStyle25();
    showModal('Hotspot-Optik',`<div class="grid2">
      <div class="field"><label>Füllfarbe</label><input id="hsFill25" type="color" value="${esc(hs.fillColor)}"></div>
      <div class="field"><label>Rahmenfarbe</label><input id="hsBorder25" type="color" value="${esc(hs.borderColor)}"></div>
      <div class="field"><label>Transparenz</label><input id="hsOpacity25" type="range" min="0.04" max="0.45" step="0.01" value="${Number(hs.opacity)||0.14}"><span id="hsOpacityOut25">${Math.round((Number(hs.opacity)||0.14)*100)}%</span></div>
      <div class="field"><label>Rahmenstärke</label><input id="hsBorderWidth25" type="range" min="0" max="4" step="1" value="${Number(hs.borderWidth)||1}"><span id="hsBorderOut25">${Number(hs.borderWidth)||1}px</span></div>
    </div><p class="hint">Gilt für die Hotspot-Markierungen in der Ordner-Vorschau des Teilekatalogs.</p>`,[
      {text:'Speichern',action:()=>{const s=ensureHotspotStyle25(); s.fillColor=$('#hsFill25').value; s.borderColor=$('#hsBorder25').value; s.opacity=Number($('#hsOpacity25').value); s.activeOpacity=Math.min(.7,s.opacity+.16); s.borderWidth=Number($('#hsBorderWidth25').value); applyHotspotStyle25(); save(); closeModal(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
    const upd=()=>{const op=$('#hsOpacity25'),bw=$('#hsBorderWidth25'); if(op)$('#hsOpacityOut25').textContent=Math.round(Number(op.value)*100)+'%'; if(bw)$('#hsBorderOut25').textContent=Number(bw.value)+'px';};
    $('#hsOpacity25')?.addEventListener('input',upd); $('#hsBorderWidth25')?.addEventListener('input',upd);
  }

  function openHotspotEditor25(hot,folder,b){
    const files=directFiles25(folder);
    const options=files.map(f=>{const r=rowForFile25(b,f); return `<option value="${esc(r.id)}" ${r.id===hot.rowId?'selected':''}>${esc(r.name||f.name||r.part||f.fileName||'Eintrag')}</option>`;}).join('');
    showModal('Hotspot bearbeiten',`<div class="field"><label>Verknüpfter Teileeintrag</label><select id="hotRow25">${options}</select></div><div class="grid2">
      <div class="field"><label>X in %</label><input id="hotX25" type="number" min="0" max="100" step="0.1" value="${Number(hot.x||0).toFixed(1)}"></div>
      <div class="field"><label>Y in %</label><input id="hotY25" type="number" min="0" max="100" step="0.1" value="${Number(hot.y||0).toFixed(1)}"></div>
      <div class="field"><label>Breite in %</label><input id="hotW25" type="number" min="1" max="100" step="0.1" value="${Number(hot.w||10).toFixed(1)}"></div>
      <div class="field"><label>Höhe in %</label><input id="hotH25" type="number" min="1" max="100" step="0.1" value="${Number(hot.h||10).toFixed(1)}"></div>
    </div><p class="hint">Du kannst den Bereich auch direkt in der Vorschau ziehen. Unten rechts lässt er sich vergrößern/verkleinern.</p>`,[
      {text:'Speichern',action:()=>{hot.rowId=$('#hotRow25').value; hot.x=clamp25(Number($('#hotX25').value),0,99); hot.y=clamp25(Number($('#hotY25').value),0,99); hot.w=clamp25(Number($('#hotW25').value),1,100-hot.x); hot.h=clamp25(Number($('#hotH25').value),1,100-hot.y); closeModal(); save(); render();}},
      {text:'Löschen',danger:true,action:()=>{folder.previewHotspots=(folder.previewHotspots||[]).filter(x=>x.id!==hot.id); closeModal(); save(); render();}},
      {text:'Abbrechen',secondary:true,action:closeModal}
    ]);
  }
  function clamp25(v,min,max){v=Number(v); if(!Number.isFinite(v))v=min; return Math.max(min,Math.min(max,v));}

  function bindHotspotEditing25(){
    const isAdmin=currentUser?.()?.role==='Admin';
    const stage=document.getElementById('catalogPreviewStage21');
    if(!stage || stage.dataset.v25Bound==='1') return;
    stage.dataset.v25Bound='1';
    if(!isAdmin) return;

    stage.addEventListener('contextmenu',ev=>{
      const el=ev.target.closest('.catalog21-hotspot');
      if(!el) return;
      ev.preventDefault(); ev.stopPropagation(); ev.stopImmediatePropagation();
      const b=bundle25(), folder=currentFolder25(b), hot=hotspotById25(folder,el.dataset.hotspotId);
      if(hot) openHotspotEditor25(hot,folder,b);
    },true);
    stage.addEventListener('dblclick',ev=>{
      const el=ev.target.closest('.catalog21-hotspot');
      if(!el) return;
      ev.preventDefault(); ev.stopPropagation();
      const b=bundle25(), folder=currentFolder25(b), hot=hotspotById25(folder,el.dataset.hotspotId);
      if(hot) openHotspotEditor25(hot,folder,b);
    },true);

    stage.addEventListener('pointerdown',ev=>{
      const el=ev.target.closest('.catalog21-hotspot');
      if(!el || ev.button!==0) return;
      if(ev.target.closest('.hotspot-resize-handle')) return startResize25(ev,el,stage);
      return startMove25(ev,el,stage);
    },true);
  }

  function startMove25(ev,el,stage){
    const b=bundle25(), folder=currentFolder25(b), hot=hotspotById25(folder,el.dataset.hotspotId);
    if(!hot) return;
    const rect=stage.getBoundingClientRect();
    const startX=ev.clientX, startY=ev.clientY, origX=Number(hot.x)||0, origY=Number(hot.y)||0, w=Number(hot.w)||10, h=Number(hot.h)||10;
    el.setPointerCapture?.(ev.pointerId);
    let moved=false;
    function move(e){
      const dx=(e.clientX-startX)/rect.width*100, dy=(e.clientY-startY)/rect.height*100;
      if(!moved && (Math.abs(dx)>0.35||Math.abs(dy)>0.35)) moved=true;
      if(!moved) return;
      e.preventDefault(); e.stopPropagation();
      hot.x=clamp25(origX+dx,0,100-w); hot.y=clamp25(origY+dy,0,100-h);
      el.style.left=hot.x+'%'; el.style.top=hot.y+'%';
    }
    function up(e){
      el.removeEventListener('pointermove',move); el.removeEventListener('pointerup',up); el.releasePointerCapture?.(ev.pointerId);
      if(moved){e.preventDefault(); e.stopPropagation(); save(); el.dataset.justMoved='1'; setTimeout(()=>delete el.dataset.justMoved,300);}
    }
    el.addEventListener('pointermove',move); el.addEventListener('pointerup',up,{once:true});
  }
  function startResize25(ev,el,stage){
    ev.preventDefault(); ev.stopPropagation();
    const b=bundle25(), folder=currentFolder25(b), hot=hotspotById25(folder,el.dataset.hotspotId);
    if(!hot) return;
    const rect=stage.getBoundingClientRect();
    const startX=ev.clientX, startY=ev.clientY, origW=Number(hot.w)||10, origH=Number(hot.h)||10, x=Number(hot.x)||0, y=Number(hot.y)||0;
    el.setPointerCapture?.(ev.pointerId);
    function move(e){
      const dw=(e.clientX-startX)/rect.width*100, dh=(e.clientY-startY)/rect.height*100;
      hot.w=clamp25(origW+dw,1,100-x); hot.h=clamp25(origH+dh,1,100-y);
      el.style.width=hot.w+'%'; el.style.height=hot.h+'%';
    }
    function up(e){el.removeEventListener('pointermove',move); el.removeEventListener('pointerup',up); el.releasePointerCapture?.(ev.pointerId); save(); el.dataset.justMoved='1'; setTimeout(()=>delete el.dataset.justMoved,250);}
    el.addEventListener('pointermove',move); el.addEventListener('pointerup',up,{once:true});
  }

  const baseInit25=initState;
  initState=function(){
    let s=null;
    try{s=JSON.parse(localStorage.getItem(STORE_KEY_25)||localStorage.getItem(LIVE_KEY_25)||localStorage.getItem('kfz_app_V26.07.23_state')||localStorage.getItem('kfz_app_V26.07.22_state')||localStorage.getItem('kfz_app_V26.07.20_state')||'null');}catch(e){s=null;}
    if(!(s&&s.users)) s=baseInit25();
    s.version=VERSION_25;
    s.catalogHotspotStyle=s.catalogHotspotStyle||{fillColor:'#f2c94c',borderColor:'#f2c94c',opacity:0.14,activeOpacity:0.30,borderWidth:1};
    return s;
  };

  const baseSave25=save;
  save=function(){
    if(!state) return;
    state.version=VERSION_25; state.savedAt=new Date().toISOString();
    try{localStorage.setItem(STORE_KEY_25,JSON.stringify(state));localStorage.setItem(LIVE_KEY_25,JSON.stringify(state));}catch(e){console.warn(e)}
    updateSaveInfo?.();
  };

  const baseRenderCatalog25=renderCatalog;
  renderCatalog=function(sheet){
    baseRenderCatalog25(sheet);
    css25(); applyHotspotStyle25(); enhanceHotspots25(); insertHotspotStyleButton25(); bindHotspotEditing25();
  };

  const baseRender25=render;
  render=function(){
    baseRender25();
    css25(); applyHotspotStyle25(); enhanceHotspots25();
    document.title='KFZ-APP '+VERSION_25;
    const chip=document.querySelector('.version-chip'); if(chip) chip.textContent=VERSION_25;
  };

  const baseApplyTheme25=applyTheme;
  applyTheme=function(){baseApplyTheme25(); css25(); applyHotspotStyle25();};

  const baseReleaseNotes25=releaseNotes;
  releaseNotes=function(){return [{title:'V26.07.26 · Hotspots verschieben und dezenter gestalten',body:'Bereits markierte Bereiche in der Teilekatalog-Vorschau können vom Admin jetzt verschoben, über die Ecke vergrößert/verkleinert und per Doppelklick oder Rechtsklick bearbeitet werden. Die Markierungen sind dezenter/transparenter und die Hotspot-Optik kann vom Admin angepasst werden.',jump:'catalog'}, ...baseReleaseNotes25()];};
})();

document.addEventListener('DOMContentLoaded',()=>{state=initState(); bindGlobal(); render(); setInterval(()=>{if(currentUser())save()},25000);});