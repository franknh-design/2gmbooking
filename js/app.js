// ============================================================
// 2GM Booking v14.5.7 — app.js (Core)
// Auth, Graph API, Data, Rendering, Bookings
// ============================================================

// --- CONFIG ---
const msalConfig={auth:{clientId:'f8e2259d-c440-41d3-94e3-3a2dce095817',authority:'https://login.microsoftonline.com/2b495272-f733-47a8-a771-bb744309fa17',redirectUri:'https://franknh-design.github.io/2gmbooking/'},cache:{cacheLocation:'localStorage'}};
const msalInstance=new msal.PublicClientApplication(msalConfig);
const SITE_HOST='2gmeiendom.sharepoint.com';
const SITE_PATH='/sites/2GMBooking';
const LIST_IDS={Properties:'d842d574-f238-442a-be3d-77334727e89f',Rooms:'bfa962a0-5eb2-416c-abe8-adba06558c11',Bookings:'fe1dfe34-23df-4864-b0b1-b01bf60bfb75',Persons:'ebbe517d-83f8-4169-9423-70c63a3f8c07',Cleaning_Log:'6b1bd5f9-c54f-42ee-892f-d50c79481375',Hours:'9db53c54-70dd-483d-ad1d-565d0e4ac7ac',Users:'1b9b866f-0944-4f43-a80d-2a630e1e7c25',Rates:'a604493f-e879-48a0-bcab-cdeb9ae2195e'};

// --- PERMISSIONS LIST ---
const ALL_PERMS=[
  {key:'view_bookings',label:'View bookings'},
  {key:'edit_bookings',label:'Create/edit bookings'},
  {key:'checkin_out',label:'Check in/out'},
  {key:'cancel_bookings',label:'Cancel bookings'},
  {key:'cleaning',label:'Change cleaning status'},
  {key:'doortag',label:'Change door tag status'},
  {key:'print_doortag',label:'Print door tags'},
  {key:'view_hours',label:'View hours'},
  {key:'edit_hours',label:'Register hours'},
  {key:'edit_others_hours',label:'Register hours for others'},
  {key:'view_all_hours',label:'View all workers\' hours'},
  {key:'archive',label:'View archive'},
  {key:'import_export',label:'Import/Export'},
  {key:'view_prices',label:'View prices'},
  {key:'manage_rates',label:'Manage rates'},
  {key:'manage_companies',label:'Manage companies'},
  {key:'hours_reminder',label:'Daily hours reminder'},
  {key:'view_efficiency',label:'View cleaning efficiency analysis'},
  {key:'admin',label:'User administration'}
];

// --- STATE ---
let accessToken=null,siteId=null;
let currentUser={email:'',displayName:'',permissions:[]};
let properties=[],rooms=[],allRooms=[],bookings=[],allBookings=[],allUsers=[],allPersons=[],allRates=[],allCompanies=[];
let selectedProperty=null,selectedRoom=null,selectedBooking=null;
let editingBookingId=null,checkoutBookingId=null;
let activeFilter=null;
let currentView='main'; // 'main' or 'hours'
let _lastRefreshTime=Date.now();
let _knownBookingIds=new Set();
let _knownBookingModifiedMax='';
let _pollInterval=null;

// --- AUTH ---
async function signIn(){
  if(!msalReady){alert('Please wait...');return}
  try{
    await msalInstance.loginPopup({scopes:['Sites.ReadWrite.All','Mail.Send']});
    await getToken();await loadCurrentUser();
    showApp();applyPermissions();
    await loadProperties();await loadData();
    checkHoursReminder();
  }catch(e){console.error('Login failed:',e)}
}
async function getToken(){
  const a=msalInstance.getAllAccounts();if(!a.length)return null;
  try{
    const r=await msalInstance.acquireTokenSilent({scopes:['Sites.ReadWrite.All','Mail.Send'],account:a[0]});
    if(!r||!r.accessToken)throw new Error('Silent returned empty token');
    accessToken=r.accessToken;return accessToken;
  }catch(e){
    console.warn('[Auth] Silent token failed:',e.message,'— trying popup');
    try{
      const r=await msalInstance.acquireTokenPopup({scopes:['Sites.ReadWrite.All','Mail.Send']});
      if(!r||!r.accessToken)throw new Error('Popup returned empty token');
      accessToken=r.accessToken;return accessToken;
    }catch(popupErr){
      console.error('[Auth] Popup token failed:',popupErr.message);
      accessToken=null;
      alert('Sesjonen har utløpt. Last siden på nytt (F5) og logg inn på nytt.');
      throw new Error('Token unavailable');
    }
  }
}
function signOut(){msalInstance.logoutPopup();document.getElementById('app').style.display='none';document.getElementById('loginScreen').style.display='block'}
function showApp(){document.getElementById('loginScreen').style.display='none';document.getElementById('app').style.display='block'}

// --- GRAPH API ---
async function graphGet(ep){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{headers:{Authorization:'Bearer '+accessToken,Accept:'application/json'}});if(!r.ok)throw new Error('Graph error '+r.status+': '+await r.text());return r.json()}
async function graphPatch(ep,body){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'PATCH',headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},body:JSON.stringify(body)});if(!r.ok)throw new Error('Graph error '+r.status);return r.json()}
async function graphPost(ep,body){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'POST',headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},body:JSON.stringify(body)});if(!r.ok){const t=await r.text();throw new Error('Graph error '+r.status+': '+t)}return r.json()}
async function graphDelete(ep){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'DELETE',headers:{Authorization:'Bearer '+accessToken}});if(!r.ok)throw new Error('Graph error '+r.status);return true}

async function getSiteId(){if(siteId)return siteId;const r=await graphGet('/sites/'+SITE_HOST+':'+SITE_PATH);siteId=r.id;return siteId}
// Cache for dynamically resolved list IDs (lists not in LIST_IDS hardcoded map)
const _dynamicListIds={};
async function getListId(name){
  if(LIST_IDS[name])return LIST_IDS[name];
  if(_dynamicListIds[name])return _dynamicListIds[name];
  // Fall back to looking up by display name via Graph API
  const s=await getSiteId();
  try{
    const r=await graphGet('/sites/'+s+'/lists?$filter=displayName eq \''+name+'\'&$select=id,displayName');
    if(r.value&&r.value.length){
      _dynamicListIds[name]=r.value[0].id;
      return r.value[0].id;
    }
  }catch(e){console.error('Failed to lookup list '+name+':',e)}
  throw new Error('List not found: '+name);
}
async function getListItems(listName){const s=await getSiteId();const lid=await getListId(listName);let all=[];let url='/sites/'+s+'/lists/'+lid+'/items?$expand=fields&$top=500';while(url){const r=await graphGet(url);all=all.concat(r.value.map(i=>({id:i.id,...i.fields})));url=r['@odata.nextLink']?r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0',''):null}return all}

// Fetch a text file from a SharePoint document library.
// If pathInLibrary starts with a library name that's not the default, tries that library specifically.
// Example: 'Batteristatus/RoomBattery.csv' — first tries default lib with that path, then tries 'Batteristatus' as its own library.
async function fetchSiteFileText(pathInLibrary){
  const s=await getSiteId();
  await getToken();
  const errors=[];
  // Attempt 1: Default library (Shared Documents) with the full path
  try{
    const url='https://graph.microsoft.com/v1.0/sites/'+s+'/drive/root:/'+encodeURI(pathInLibrary);
    const r=await fetch(url,{headers:{Authorization:'Bearer '+accessToken}});
    if(r.ok){
      const item=await r.json();
      if(item['@microsoft.graph.downloadUrl']){
        const c=await fetch(item['@microsoft.graph.downloadUrl']);
        if(c.ok)return c.text();
      }
    }
    errors.push('Default library: '+r.status);
  }catch(e){errors.push('Default library: '+e.message)}
  // Attempt 2: Parse path as "LibraryName/file/path" — try as separate library
  const firstSlash=pathInLibrary.indexOf('/');
  if(firstSlash>0){
    const libName=pathInLibrary.substring(0,firstSlash);
    const remaining=pathInLibrary.substring(firstSlash+1);
    try{
      // Find the drive with matching name
      const drives=await graphGet('/sites/'+s+'/drives');
      const lib=drives.value.find(d=>d.name===libName||d.name.toLowerCase()===libName.toLowerCase());
      if(lib){
        const url='https://graph.microsoft.com/v1.0/drives/'+lib.id+'/root:/'+encodeURI(remaining);
        const r=await fetch(url,{headers:{Authorization:'Bearer '+accessToken}});
        if(r.ok){
          const item=await r.json();
          if(item['@microsoft.graph.downloadUrl']){
            const c=await fetch(item['@microsoft.graph.downloadUrl']);
            if(c.ok)return c.text();
          }
        }
        errors.push('Library "'+libName+'": '+r.status);
      }else{
        errors.push('Library "'+libName+'" not found among: '+drives.value.map(d=>d.name).join(', '));
      }
    }catch(e){errors.push('Library search: '+e.message)}
  }
  throw new Error('File not found. Tried:\n'+errors.join('\n'));
}
// Cache of known columns per list. Populated lazily on first save attempt.
const _knownColumnsByList={};
const _unknownFieldsByList={};

async function _discoverColumns(listName){
  if(_knownColumnsByList[listName])return _knownColumnsByList[listName];
  try{
    const s=await getSiteId();const lid=await getListId(listName);
    const res=await graphGet('/sites/'+s+'/lists/'+lid+'/columns?$select=name,displayName');
    const cols=new Set();
    (res.value||[]).forEach(c=>{if(c.name)cols.add(c.name)});
    // Also add common system fields that should always be allowed even if not in schema
    ['Title'].forEach(k=>cols.add(k));
    _knownColumnsByList[listName]=cols;
    console.log('[SharePoint] Discovered '+cols.size+' columns for '+listName+':',[...cols].sort().join(', '));
    return cols;
  }catch(e){
    console.warn('Could not discover columns for '+listName+':',e.message);
    _knownColumnsByList[listName]=new Set();
    return _knownColumnsByList[listName];
  }
}

async function _stripUnknownFieldsAsync(listName,fields){
  const cols=await _discoverColumns(listName);
  if(!cols||!cols.size)return fields; // discovery failed — let SharePoint reject as before
  const cleaned={};
  const skipped=[];
  Object.keys(fields).forEach(k=>{
    // Always allow Lookup-prefixed fields (e.g. RoomLookupId) — SharePoint resolves these
    if(k.endsWith('LookupId')||cols.has(k)){
      let v=fields[k];
      // PRAGMATIC: Yes/No fields cause 500 errors via Graph API.
      // Skip them entirely — SharePoint default value will be used.
      // TODO: figure out correct format. For now this gets bookings working.
      const isBool=(typeof v==='boolean'||v===0||v===1);
      const isYesNoField=(k==='Include_Checkout_Fee'||k==='Continuation');
      if(isBool&&isYesNoField){
        console.log('[SharePoint] Skipping Yes/No field "'+k+'" with value '+v+' (Graph API issue — using SharePoint default)');
        return; // skip this field
      }
      cleaned[k]=v;
    }else{
      skipped.push(k);
      if(!_unknownFieldsByList[listName])_unknownFieldsByList[listName]=new Set();
      _unknownFieldsByList[listName].add(k);
    }
  });
  if(skipped.length){
    console.warn('[SharePoint] Skipping unknown columns in '+listName+': '+skipped.join(', ')+'. Create these in SharePoint to enable.');
  }
  return cleaned;
}

async function createListItem(listName,fields){
  const cleaned=await _stripUnknownFieldsAsync(listName,fields);
  // Strip null/undefined values for create — SharePoint can throw 500 on unexpected null
  const final={};
  Object.keys(cleaned).forEach(k=>{if(cleaned[k]!==null&&cleaned[k]!==undefined)final[k]=cleaned[k]});
  const s=await getSiteId();const lid=await getListId(listName);
  console.log('[SharePoint] CREATE '+listName+' payload:',JSON.parse(JSON.stringify(final)));
  console.log('[SharePoint] CREATE '+listName+' field names:',Object.keys(final).join(', '));
  console.log('[SharePoint] Known columns:',[..._knownColumnsByList[listName]||[]].sort().join(', '));
  try{
    return await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:final});
  }catch(e){
    if(String(e.message||'').indexOf('500')<0&&String(e.message||'').indexOf('General exception')<0)throw e;
    // 500 with no useful info → systematic bisect
    console.warn('[BISECT] 500 received. Building payload up from minimal to find the broken field combination...');
    const keys=Object.keys(final);
    // Phase 1: try absolute minimal — just Title (or empty)
    const startMinimal={};
    if(final.Title)startMinimal.Title=final.Title;
    else startMinimal.Title='_BISECT_TEST_'+Date.now();
    let lastWorking=null;
    let lastWorkingItemId=null;
    try{
      console.log('[BISECT] Phase 1: minimal payload',startMinimal);
      const r=await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:startMinimal});
      console.log('[BISECT] ✓ Minimal succeeded with id='+r.id);
      lastWorking={...startMinimal};
      lastWorkingItemId=r.id;
    }catch(e2){
      console.warn('[BISECT] ✗ Even minimal payload failed:',e2.message);
      throw new Error('Save failed. Even a minimal payload (just Title) fails. This is a list-level problem in SharePoint, not a field problem. Original error: '+e.message);
    }
    // Phase 2: add fields one at a time
    let breakingField=null;
    let breakingValue=null;
    for(let i=0;i<keys.length;i++){
      const k=keys[i];
      if(k in lastWorking)continue;
      const testFields={...lastWorking,[k]:final[k]};
      try{
        console.log('[BISECT] Adding "'+k+'"='+JSON.stringify(final[k])+'...');
        // Delete previous test item before creating new one
        if(lastWorkingItemId){try{await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+lastWorkingItemId)}catch(e3){}}
        const r=await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:testFields});
        lastWorking=testFields;
        lastWorkingItemId=r.id;
        console.log('[BISECT] ✓ OK with "'+k+'"');
      }catch(e2){
        console.warn('[BISECT] ✗ FAILED when adding "'+k+'"='+JSON.stringify(final[k])+':',e2.message);
        breakingField=k;
        breakingValue=final[k];
        break;
      }
    }
    // Cleanup last test item
    if(lastWorkingItemId){try{await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+lastWorkingItemId)}catch(e3){console.warn('[BISECT] Could not delete test item '+lastWorkingItemId+' — please remove manually')}}
    if(breakingField){
      throw new Error('Save failed. Adding field "'+breakingField+'" with value '+JSON.stringify(breakingValue)+' broke the request. Check SharePoint column type/required. Last working set: '+Object.keys(lastWorking).join(', '));
    }
    throw new Error('Save failed unexpectedly. Bisect added all fields without breaking but original payload still failed. Strange. Original error: '+e.message);
  }
}
async function updateListItem(listName,itemId,fields){
  const cleaned=await _stripUnknownFieldsAsync(listName,fields);
  const s=await getSiteId();const lid=await getListId(listName);
  return graphPatch('/sites/'+s+'/lists/'+lid+'/items/'+itemId+'/fields',cleaned);
}

// --- USER & PERMISSIONS ---
async function loadCurrentUser(){
  const accounts=msalInstance.getAllAccounts();if(!accounts.length)return;
  const email=(accounts[0].username||'').toLowerCase();
  currentUser.email=email;
  currentUser.displayName=email;
  try{
    allUsers=await getListItems('Users');
    const match=allUsers.find(u=>(u.Epost||'').toLowerCase()===email&&u.Active!==false);
    if(match){
      currentUser.displayName=match.DisplayName||email;
      currentUser.userId=match.id;
      // Parse permissions from comma-separated string, or use defaults
      if(match.Permissions){
        currentUser.permissions=match.Permissions.split(',').map(s=>s.trim());
      }else{
        // Legacy: map old Role to permissions
        const roleMap={
          SuperAdmin:ALL_PERMS.map(p=>p.key),
          Admin:ALL_PERMS.filter(p=>p.key!=='admin').map(p=>p.key),
          Cleaner:['cleaning','print_doortag','view_hours','edit_hours'],
          ReadOnly:['view_bookings']
        };
        currentUser.permissions=roleMap[match.Role]||['view_bookings'];
      }
    }else{
      currentUser.permissions=['view_bookings'];
    }
  }catch(e){console.error('Failed to load user:',e);currentUser.permissions=['view_bookings']}
  console.log('User:',currentUser.email,'Perms:',currentUser.permissions);
}

function can(perm){return currentUser.permissions.includes(perm)}

// Returns the effective billing company for a booking (falls back to Company if no Billing_Company set)
function getEffectiveCompany(b){
  if(!b)return '';
  const bc=(b.Billing_Company||'').trim();
  if(bc)return bc;
  return (b.Company||'').trim();
}

// Check if a name exists in the Persons/Guests list (fuzzy match)
function isKnownGuest(name){
  if(!name||!allPersons||!allPersons.length)return false;
  const lower=name.toLowerCase().trim();
  if(!lower)return false;
  const words=lower.split(/[\s,]+/).filter(w=>w.length>1);
  return allPersons.some(p=>{
    const pn=(p.Name||p.Title||'').toLowerCase().trim();
    if(!pn)return false;
    if(pn===lower)return true;
    if(words.length<2)return false;
    const pwords=pn.split(/[\s,]+/).filter(w=>w.length>1);
    if(pwords.length<2)return false;
    return words.every(w=>pn.indexOf(w)>=0)||pwords.every(w=>lower.indexOf(w)>=0);
  });
}

// Wrap a name in known-guest indicator if applicable
function guestMarkedName(name){
  if(!name)return '';
  const escaped=name.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]));
  return isKnownGuest(name)?'<span class="known-guest">'+escaped+'</span>':escaped;
}

// Update which nav button is highlighted as active
function updateNavActiveState(){
  const btns={btnUpcoming:'incomingPanel',btnArchive:'archivePanel',btnPersons:'personsPanel'};
  Object.keys(btns).forEach(bid=>{
    const btn=document.getElementById(bid);if(!btn)return;
    const panel=document.getElementById(btns[bid]);
    const isActive=panel&&panel.classList.contains('open');
    btn.classList.toggle('active-nav',isActive);
  });
  // Hours is special — it's a view, not a panel
  const hBtn=document.getElementById('btnHours');
  if(hBtn)hBtn.classList.toggle('active-nav',currentView==='hours');
  // Invoicing menu item
  const invPanel=document.getElementById('invoicingPanel');
  const invActive=invPanel&&invPanel.classList.contains('open');
  const mi=document.getElementById('menuBtnInvoicing');if(mi)mi.classList.toggle('active-nav',invActive);
  // More-button gets highlight if any menu item is active
  const mb=document.getElementById('btnMoreMenu');if(mb)mb.classList.toggle('active-nav',invActive);
}

function applyPermissions(){
  const el=id=>document.getElementById(id);
  const show=(id,vis)=>{const e=el(id);if(e)e.style.display=vis?'':'none'};
  const showBlock=(id,vis)=>{const e=el(id);if(e)e.style.display=vis?'block':'none'};
  // Header buttons
  show('btnNewBooking',can('edit_bookings'));
  show('btnNewGuest',can('edit_bookings'));
  show('menuBtnCompanies',can('manage_companies')||can('admin'));
  show('menuBtnBackup',can('admin'));
  show('menuBtnRestore',can('admin'));
  show('menuBtnTemplates',can('admin')||can('manage_properties'));
  show('menuBtnMassSMS',can('edit_bookings'));
  show('menuBtnMassEmail',can('edit_bookings'));
  show('btnArchive',can('archive')||can('view_bookings'));
  show('btnUpcoming',can('view_bookings'));
  show('btnHours',can('view_hours')||can('edit_hours'));
  show('efficiencyBtn',can('view_efficiency'));
  showBlock('adminBar',can('admin')||can('manage_rates'));
  // Rates button only if manage_rates
  const rb=el('ratesBtn');if(rb)rb.style.display=can('manage_rates')||can('admin')?'':'none';
  // Property select
  const ps=el('propertySelect');if(ps)ps.style.display='';
  // Sign out label
  const so=el('btnSignOut');if(so)so.textContent=currentUser.displayName+' — Sign out';
  // Stats: hide if no view permission
  if(!can('view_bookings')){
    const sb=el('statsBar');if(sb)sb.style.display='none';
    const fl=document.querySelector('.floors');if(fl)fl.style.display='none';
  }
}

// --- DATA LOADING ---
function resetViewStateForPropertyChange(){
  // Clear active filter
  activeFilter=null;
  const fb=document.getElementById('filterBar');if(fb)fb.classList.remove('active');
  // Close detail panel
  const dp=document.getElementById('detailPanel');if(dp)dp.classList.remove('open');
  selectedRoom=null;selectedBooking=null;
  // Close side panels
  ['incomingPanel','archivePanel','personsPanel','invoicingPanel'].forEach(id=>{
    const el=document.getElementById(id);if(el)el.classList.remove('open');
  });
  // Remove panel-mode on main view
  const mv=document.getElementById('mainView');if(mv)mv.classList.remove('panel-mode');
  // If user was in Hours view, go back to main for the new property
  if(currentView==='hours'){
    currentView='main';
    const hv=document.getElementById('hoursView');if(hv)hv.style.display='none';
    if(mv)mv.style.display='';
    const ps=document.getElementById('propertySelect');if(ps)ps.style.display='';
  }
  // Close More-menu if open
  const mm=document.getElementById('moreMenu');if(mm)mm.style.display='none';
  // Update nav-active state (all inactive after reset)
  if(typeof updateNavActiveState==='function')updateNavActiveState();
}

async function loadProperties(){
  try{
    const allProps=await getListItems('Properties');
    // Filter properties based on user's AssignedProperties (if set)
    const user=allUsers.find(u=>(u.Epost||'').toLowerCase()===currentUser.email);
    const assigned=user&&user.AssignedProperties?user.AssignedProperties.split(',').map(s=>s.trim()):[];
    if(assigned.length>0&&!can('admin')){
      properties=allProps.filter(p=>assigned.includes(p.Title));
    }else{
      properties=allProps;
    }
    const sel=document.getElementById('propertySelect');
    sel.innerHTML='<option value="__ALL__">⭐ All properties</option>'+properties.map(p=>'<option value="'+p.id+'">'+p.Title+'</option>').join('');
    sel.onchange=()=>{
      if(sel.value==='__ALL__'){
        selectedProperty=null; // null means "all properties" mode
      }else{
        selectedProperty=properties.find(p=>p.id===sel.value);
      }
      resetViewStateForPropertyChange();
      loadData();
    };
    selectedProperty=null; // v14.5.7: default to "All properties" instead of first property
  }catch(e){console.error('Error loading properties:',e)}
}

async function loadData(){
  _lastRefreshTime=Date.now();
  const isAll=selectedProperty===null;
  document.getElementById('headerTitle').textContent='2GM Eiendom AS – Booking'+(isAll?' — All properties':(selectedProperty?' — '+selectedProperty.Title:''));
  document.getElementById('floor1Body').innerHTML='<tr><td colspan="7" class="loading">Loading...</td></tr>';
  document.getElementById('floor2Body').innerHTML='<tr><td colspan="7" class="loading">Loading...</td></tr>';
  closeDetail();
  try{
    allRooms=await getListItems('Rooms');
    allBookings=await getListItems('Bookings');
    try{allPersons=await getListItems('Persons')}catch(e){allPersons=[]}
    try{allRates=await getListItems('Rates')}catch(e){allRates=[]}
    try{allCompanies=await getListItems('Companies')}catch(e){allCompanies=[]}
    if(isAll){
      // Show rooms from all properties the user has access to
      const assignedPropIds=new Set(properties.map(p=>String(p.id)));
      rooms=allRooms.filter(r=>assignedPropIds.has(String(r.PropertyLookupId)));
    }else{
      rooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(selectedProperty.id));
      if(rooms.length===0){rooms=allRooms.filter(r=>r.Active!==false)}
    }
    filterBookingsForView();
    renderFloors();updateStats();
    refreshPersonDatalists();
    // Re-render Upcoming if open (otherwise it shows bookings from previous property)
    if(document.getElementById('incomingPanel').classList.contains('open'))renderIncoming();
    if(document.getElementById('personsPanel').classList.contains('open'))renderPersons();
  }catch(e){console.error('Error:',e);document.getElementById('floor1Body').innerHTML='<tr><td colspan="7" class="error">Error: '+e.message+'</td></tr>'}
}

function filterBookingsForView(){
  const roomIds=new Set(rooms.map(r=>r.id));
  bookings=allBookings.filter(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!roomIds.has(rid))return false;
    if(b.Status==='Active')return true;
    if(b.Status==='Upcoming'){
      // v14.5.7: Show Upcoming in main list as soon as Check_In <= today (no kl-12 rule)
      const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
      const today=new Date();today.setHours(0,0,0,0);
      if(ci.getTime()<=today.getTime())return true;
    }
    return false;
  });
}

function refreshLocal(){
  filterBookingsForView();
  renderFloors();updateStats();
}

// --- UTILS ---
function formatDate(d){if(!d)return'';const dt=new Date(d);return String(dt.getDate()).padStart(2,'0')+'.'+String(dt.getMonth()+1).padStart(2,'0')+'.'+dt.getFullYear()}
function toISODate(d){if(!d)return'';const dt=new Date(d);return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0')}

function getNextWeekday(date){const d=new Date(date);const day=d.getDay();if(day===0)d.setDate(d.getDate()+1);else if(day===6)d.setDate(d.getDate()+2);return d}

function calcWashDates(checkInDate,checkOutDate){
  const ci=new Date(checkInDate);ci.setHours(0,0,0,0);
  const co=checkOutDate?new Date(checkOutDate):null;if(co)co.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  const washes=[];let week=1;
  while(week<=52){
    const rawDate=new Date(ci);rawDate.setDate(rawDate.getDate()+week*7);
    const washDate=getNextWeekday(rawDate);
    if(co&&washDate>=co)break;
    const type=(week%2===1)?'Towels':'Towels + Beddings';
    const isPast=washDate<today;const isToday=washDate.getTime()===today.getTime();
    const isNext=!isPast&&!isToday&&washes.every(w=>w.isPast||w.isToday);
    washes.push({date:washDate,type,week,isPast,isToday,isNext});week++;
  }
  return washes;
}

function getWashScheduleHtml(booking){
  if(!booking||!booking.Check_In||!(booking.Status==='Active'||booking.Status==='Upcoming'))return'';
  const washes=calcWashDates(booking.Check_In,booking.Check_Out);
  const show=washes.filter(w=>!w.isPast).slice(0,6);if(!show.length)return'';
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  let html='<div style="margin-top:14px"><div style="font-size:12px;color:var(--text-secondary);margin-bottom:4px;font-weight:500">Wash schedule</div><table style="font-size:12px;width:auto">';
  show.forEach(w=>{
    let s='',badge='';
    if(w.isToday){s='color:var(--text-danger);font-weight:500';badge=' <span class="pill danger">Today</span>'}
    else if(w.isNext){s='color:var(--accent);font-weight:500';badge=' <span class="pill" style="background:var(--bg-success);color:var(--text-success)">Next</span>'}
    html+='<tr style="'+s+'"><td style="padding:2px 12px 2px 0">'+days[w.date.getDay()]+' '+formatDate(w.date)+badge+'</td><td style="padding:2px 0">'+w.type+'</td></tr>';
  });
  return html+'</table></div>';
}

function getNextWashDate(booking){
  if(!booking||!booking.Check_In||booking.Status!=='Active')return'';
  const washes=calcWashDates(booking.Check_In,booking.Check_Out);
  const next=washes.find(w=>!w.isPast);if(!next)return'<span class="muted">—</span>';
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  if(next.isToday)return'<span class="pill danger">Today — '+next.type+'</span>';
  return days[next.date.getDay()]+' '+formatDate(next.date)+' — '+next.type;
}

// --- RENDERING ---
function getBookingForRoom(roomId){
  return bookings.find(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Active')
    ||bookings.find(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Upcoming');
}

// --- PRICING ---
function _nameMatch(a,b){
  // Exact case-insensitive
  const la=(a||'').toLowerCase().trim();const lb=(b||'').toLowerCase().trim();
  if(!la||!lb)return false;
  if(la===lb)return true;
  // All words match (regardless of order) — handles "Marek Filas" vs "Filas, Marek"
  const wa=la.split(/[\s,]+/).filter(w=>w.length>1);
  const wb=lb.split(/[\s,]+/).filter(w=>w.length>1);
  if(wa.length<2||wb.length<2)return false;
  // a contains all words of b OR b contains all words of a
  return wa.every(w=>lb.indexOf(w)>=0)||wb.every(w=>la.indexOf(w)>=0);
}

function getDailyRate(personName,company,propertyTitle,roomId){
  // Priority: 1) Person+Property  2) Person (any)  3) Company+Property  4) Company (any)  5) Room rate  6) Property default
  const pn=(personName||'').toLowerCase();
  const co=(company||'').toLowerCase();
  const pt=(propertyTitle||'').toLowerCase();

  // 1. Person + specific property (fuzzy name, exact property)
  // Exclude rates with FeeType=Checkout (those are one-time fees, not nightly)
  const isNightly=r=>(r.FeeType||'').toLowerCase()!=='checkout';
  let rate=allRates.find(r=>isNightly(r)&&_nameMatch(r.Person_Name,personName)&&(r.Property||'').toLowerCase()===pt&&r.DailyRate);
  if(rate)return{rate:rate.DailyRate,source:'Person+Property',matchedName:rate.Person_Name};

  // 2. Person any property (fuzzy)
  rate=allRates.find(r=>isNightly(r)&&_nameMatch(r.Person_Name,personName)&&!(r.Property)&&r.DailyRate);
  if(rate)return{rate:rate.DailyRate,source:'Person',matchedName:rate.Person_Name};

  // 3. Company + specific property
  if(co){
    rate=allRates.find(r=>isNightly(r)&&(r.Company||'').toLowerCase()===co&&(r.Property||'').toLowerCase()===pt&&r.DailyRate);
    if(rate)return{rate:rate.DailyRate,source:'Company+Property'};
  }

  // 4. Company any property
  if(co){
    rate=allRates.find(r=>isNightly(r)&&(r.Company||'').toLowerCase()===co&&!(r.Property)&&r.DailyRate);
    if(rate)return{rate:rate.DailyRate,source:'Company'};
  }

  // 5. Room rate
  if(roomId){
    const room=allRooms.find(r=>r.id===String(roomId));
    if(room&&room.DailyRate)return{rate:room.DailyRate,source:'Room rate'};
  }

  // 6. Property default
  const prop=properties.find(p=>(p.Title||'').toLowerCase()===pt);
  if(prop&&prop.DailyRate)return{rate:prop.DailyRate,source:'Property default'};

  // No rate found — but check for near-misses and flag them
  const nearMiss=_findRateNearMiss(personName,company,propertyTitle);
  return{rate:0,source:'No rate set',nearMiss:nearMiss};
}

// Look up checkout fee (one-time cleaning fee at end of stay)
// Priority: 1) Company+Property  2) Company  3) Property  4) 0 (no fee)
function getCheckoutFee(company,propertyTitle){
  const co=(company||'').toLowerCase().trim();
  const pt=(propertyTitle||'').toLowerCase().trim();
  // Only consider rates explicitly marked as Checkout fee
  const checkoutRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='checkout'&&r.DailyRate);
  if(!checkoutRates.length)return 0;
  // 1. Company + specific property
  if(co){
    const r=checkoutRates.find(rt=>(rt.Company||'').toLowerCase()===co&&(rt.Property||'').toLowerCase()===pt);
    if(r)return Number(r.DailyRate)||0;
  }
  // 2. Company any property
  if(co){
    const r=checkoutRates.find(rt=>(rt.Company||'').toLowerCase()===co&&!(rt.Property));
    if(r)return Number(r.DailyRate)||0;
  }
  // 3. Property default
  const r=checkoutRates.find(rt=>(rt.Property||'').toLowerCase()===pt&&!(rt.Company));
  if(r)return Number(r.DailyRate)||0;
  return 0;
}

// Look up percent-based fee for a company (e.g. Jobzone 10% of month).
// Priority: Company+Property > Company. Returns percent as decimal (0.10 for 10%) or 0 if not configured.
function getPercentFeeRate(company,propertyTitle){
  const co=(company||'').toLowerCase().trim();
  if(!co)return 0;
  const pt=(propertyTitle||'').toLowerCase().trim();
  const percentRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='percent'&&r.DailyRate);
  if(!percentRates.length)return 0;
  // 1. Company + specific property
  const r1=percentRates.find(rt=>(rt.Company||'').toLowerCase()===co&&(rt.Property||'').toLowerCase()===pt);
  if(r1)return (Number(r1.DailyRate)||0)/100;
  // 2. Company any property
  const r2=percentRates.find(rt=>(rt.Company||'').toLowerCase()===co&&!(rt.Property));
  if(r2)return (Number(r2.DailyRate)||0)/100;
  return 0;
}

// Does this company have a percent-based fee configured? (used to skip flat checkout fee)
function hasPercentFee(company,propertyTitle){
  return getPercentFeeRate(company,propertyTitle)>0;
}

// Detect rate config issues: rate exists for this name but property mismatch, or fuzzy company name
function _findRateNearMiss(personName,company,propertyTitle){
  const pn=(personName||'').toLowerCase().trim();
  const pt=(propertyTitle||'').toLowerCase().trim();
  const co=(company||'').toLowerCase().trim();
  // Does a rate exist with this name but a different property set?
  if(pn){
    const r=allRates.find(rt=>_nameMatch(rt.Person_Name,personName)&&rt.Property&&(rt.Property||'').toLowerCase()!==pt&&rt.DailyRate);
    if(r)return'Rate exists for "'+r.Person_Name+'" but only for property "'+r.Property+'" (this booking is at "'+propertyTitle+'")';
  }
  // Does a rate exist where Person_Name appears in the rate's Company field? (possible data entry mistake)
  if(pn){
    const r=allRates.find(rt=>(rt.Company||'').toLowerCase().includes(pn)&&rt.DailyRate);
    if(r)return'A rate with "'+personName+'" appears in the Company field of another rate row — possible data entry mistake';
  }
  return null;
}

function calcBookingNights(booking){
  if(!booking||!booking.Check_In)return 0;
  const ci=new Date(booking.Check_In);ci.setHours(0,0,0,0);
  const co=booking.Check_Out?new Date(booking.Check_Out):new Date();co.setHours(0,0,0,0);
  return Math.max(0,Math.round((co-ci)/864e5));
}

function calcBookingCost(booking,propertyTitle){
  const nights=calcBookingNights(booking);
  // Rate follows billing company (if set), otherwise guest's own company
  const effectiveCompany=getEffectiveCompany(booking);
  const rateInfo=getDailyRate(booking.Person_Name,effectiveCompany,propertyTitle,booking.RoomLookupId);
  return{nights,rate:rateInfo.rate,total:nights*rateInfo.rate,source:rateInfo.source,matchedName:rateInfo.matchedName||null,nearMiss:rateInfo.nearMiss||null};
}

function doorTagBtn(b){
  if(!b)return'<button class="status-btn" disabled></button>';
  const s=b.Door_Tag_Status||'None';
  if(s==='Needs-print')return'<button class="status-btn needs-print" onclick="cycleDT(event,\''+b.id+'\')">✕</button>';
  if(s==='Printed')return'<button class="status-btn printed" onclick="cycleDT(event,\''+b.id+'\')">✓</button>';
  return'<button class="status-btn" onclick="cycleDT(event,\''+b.id+'\')"></button>';
}
function cleanBtn(b){
  if(!b)return'<button class="clean-btn" disabled></button>';
  const s=b.Cleaning_Status||'None';
  if(s==='Dirty')return'<button class="clean-btn dirty" onclick="cycleCS(event,\''+b.id+'\')"></button>';
  if(s==='Clean')return'<button class="clean-btn clean" onclick="cycleCS(event,\''+b.id+'\')"></button>';
  return'<button class="clean-btn" onclick="cycleCS(event,\''+b.id+'\')"></button>';
}
function batCell(l){if(l==null)return'<span class="muted">—</span>';if(l<30)return'<span class="pill danger">'+l+'%</span>';if(l<60)return'<span class="pill warning">'+l+'%</span>';return'<span>'+l+'%</span>'}
function datesCell(b){
  if(!b)return'<span class="empty-text">Empty</span>';
  const ci=formatDate(b.Check_In);const co=b.Check_Out?formatDate(b.Check_Out):'Open-ended';
  const today=new Date();today.setHours(0,0,0,0);const ind=new Date(b.Check_In);ind.setHours(0,0,0,0);
  const days=Math.round((ind-today)/864e5);let s='';
  if(b.Status==='Upcoming'||days>0){if(days>=0&&days<=4)s='color:var(--accent);font-weight:500;';else if(days>4&&days<=30)s='color:#EF9F27;font-weight:500;'}
  // v14.5.7: overdue badges
  let overdueBadge='';
  if(isOverdueCheckIn(b)){
    const d=daysOverdueCheckIn(b);
    overdueBadge=' <span style="background:rgba(209,67,67,.12);color:#A32D2D;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500" title="Should have been checked in '+d+' day'+(d===1?'':'s')+' ago">⚠ Check-in '+d+'d</span>';
  }else if(isOverdueCheckOut(b)){
    const d=daysOverdueCheckOut(b);
    overdueBadge=' <span style="background:rgba(209,67,67,.12);color:#A32D2D;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500" title="Should have been checked out '+d+' day'+(d===1?'':'s')+' ago">⚠ Check-out '+d+'d</span>';
  }
  return'<span style="'+s+'">'+ci+'</span> — '+co+overdueBadge;
}

// Returns the full-tenant company for a property if active on the given date (default today), else null
function getActiveFullTenant(property,date){
  if(!property)return null;
  const company=(property.FullTenant_Company||'').trim();
  if(!company)return null;
  const rate=Number(property.FullTenant_RatePerRoom)||0;
  if(!rate)return null;
  const checkDate=date||new Date();
  const start=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):null;
  const end=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):null;
  if(start&&checkDate<start)return null;
  if(end&&checkDate>end)return null;
  return{company,rate,start,end};
}

// Checks if the property containing a given room is on full-tenant lease today
function getRoomFullTenant(room,date){
  if(!room)return null;
  const prop=properties.find(p=>String(p.id)===String(room.PropertyLookupId));
  return prop?getActiveFullTenant(prop,date):null;
}

// Returns active long-term contract for a single room on the given date, or null.
// Long-term = a fixed contract on a specific room (e.g. SalMar leases Leilighet 1A monthly).
// Distinct from full-tenant which applies to ALL rooms on a property uniformly.
function getActiveLongTermContract(room,date){
  if(!room)return null;
  const company=(room.LongTerm_Company||'').trim();
  if(!company)return null;
  const price=Number(room.LongTerm_Price)||0;
  if(!price)return null;
  const checkDate=date||new Date();
  const start=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):null;
  const end=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):null;
  if(start&&checkDate<start)return null;
  if(end&&checkDate>end)return null;
  const rateUnitRaw=(room.LongTerm_RateUnit||'Per day').toString().toLowerCase().trim();
  const isMonthly=rateUnitRaw.indexOf('month')>=0;
  return{company,price,start,end,isMonthly};
}

// Computes long-term contract amount for a single room over a period (handles pro-rata).
function computeLongTermForRoomPeriod(room,fromDate,toDate){
  const c=getActiveLongTermContract(room);
  if(!c)return null;
  // Recompute date overlap with agreement bounds (use start/end from contract)
  const agreementStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):new Date(1970,0,1);
  const agreementEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  const days=Math.floor((effTo-effFrom)/86400000)+1;
  let total,unitLabel,detailLabel;
  if(c.isMonthly){
    let monthFraction=0;
    let cursor=new Date(effFrom.getFullYear(),effFrom.getMonth(),1);
    while(cursor<=effTo){
      const monthStart=new Date(cursor.getFullYear(),cursor.getMonth(),1);
      const monthEnd=new Date(cursor.getFullYear(),cursor.getMonth()+1,0,23,59,59);
      const periodInMonthStart=new Date(Math.max(monthStart.getTime(),effFrom.getTime()));
      const periodInMonthEnd=new Date(Math.min(monthEnd.getTime(),effTo.getTime()));
      if(periodInMonthStart<=periodInMonthEnd){
        const daysInMonth=monthEnd.getDate();
        const periodDaysInMonth=Math.floor((periodInMonthEnd-periodInMonthStart)/86400000)+1;
        monthFraction+=periodDaysInMonth/daysInMonth;
      }
      cursor=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
    }
    total=Math.round(c.price*monthFraction*100)/100;
    detailLabel=c.price.toLocaleString('nb-NO')+' kr/mnd × '+monthFraction.toFixed(3)+' mnd';
  }else{
    total=Math.round(c.price*days*100)/100;
    detailLabel=c.price.toLocaleString('nb-NO')+' kr/dag × '+days+' dager';
  }
  return{room,company:c.company,price:c.price,isMonthly:c.isMonthly,days,total,detailLabel};
}

// Splits a long-term room's invoicing period into segments based on actual bookings
// in the room. Each segment is either a guest stay or an empty period (where the
// company still pays). Total of all segments matches the room's contract amount,
// with rounding adjustment on the last segment.
//
// Returns array of: {fromDate, toDate, days, name, isEmpty, total, dailyRate}
function segmentLongTermRoom(room,fromDate,toDate){
  const c=getActiveLongTermContract(room);
  if(!c)return null;
  // Determine the effective period (overlap of contract and invoicing period)
  const agreementStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):new Date(1970,0,1);
  const agreementEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  const totalDays=Math.floor((effTo-effFrom)/86400000)+1;
  // Compute the total contract amount for this period (with pro-rata)
  const contractCalc=computeLongTermForRoomPeriod(room,fromDate,toDate);
  if(!contractCalc)return null;
  const contractTotal=contractCalc.total;
  const dailyRate=contractTotal/totalDays;
  // Find all bookings on this room that overlap the period
  const roomBookings=allBookings.filter(b=>{
    if(String(b.RoomLookupId)!==String(room.id))return false;
    if(b.Status==='Cancelled')return false;
    if(!b.Check_In)return false;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):effTo;co.setHours(0,0,0,0);
    if(co<effFrom||ci>effTo)return false;
    return true;
  }).map(b=>{
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):effTo;co.setHours(0,0,0,0);
    // Clip to effective period
    const sFrom=ci<effFrom?effFrom:ci;
    const sTo=co>effTo?effTo:co;
    return {bookingId:b.id,name:b.Person_Name||'(uten navn)',from:sFrom,to:sTo};
  }).sort((a,b)=>a.from-b.from);
  // Build segments: walk from effFrom forward, alternate between booking and empty
  const segments=[];
  let cursor=new Date(effFrom);
  for(let i=0;i<roomBookings.length;i++){
    const bk=roomBookings[i];
    // Empty segment before this booking?
    if(bk.from>cursor){
      const segTo=new Date(bk.from.getTime()-86400000);// day before booking starts
      segTo.setHours(0,0,0,0);
      if(segTo>=cursor){
        const days=Math.floor((segTo-cursor)/86400000)+1;
        segments.push({fromDate:new Date(cursor),toDate:segTo,days,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
      }
    }
    // The booking segment
    const bSegFrom=bk.from>cursor?bk.from:cursor;
    const bSegTo=bk.to;
    if(bSegTo>=bSegFrom){
      const days=Math.floor((bSegTo-bSegFrom)/86400000)+1;
      segments.push({fromDate:new Date(bSegFrom),toDate:new Date(bSegTo),days,name:bk.name,isEmpty:false,bookingId:bk.bookingId});
    }
    // Move cursor to day after booking ends
    const nextCursor=new Date(bSegTo.getTime()+86400000);
    nextCursor.setHours(0,0,0,0);
    if(nextCursor>cursor)cursor=nextCursor;
  }
  // Trailing empty segment after last booking
  if(cursor<=effTo){
    const days=Math.floor((effTo-cursor)/86400000)+1;
    segments.push({fromDate:new Date(cursor),toDate:new Date(effTo),days,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
  }
  // No bookings at all → one big empty segment
  if(segments.length===0){
    segments.push({fromDate:new Date(effFrom),toDate:new Date(effTo),days:totalDays,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
  }
  // Compute totals per segment, with rounding adjustment on last segment
  let runningTotal=0;
  segments.forEach((s,idx)=>{
    if(idx===segments.length-1){
      // Last segment gets the rounding diff
      s.total=Math.round((contractTotal-runningTotal)*100)/100;
    }else{
      s.total=Math.round(s.days*dailyRate*100)/100;
      runningTotal+=s.total;
    }
    s.dailyRate=dailyRate;
  });
  return{
    room,company:c.company,price:c.price,isMonthly:c.isMonthly,
    contractTotal,totalDays,dailyRate,segments,
    contractDetailLabel:contractCalc.detailLabel
  };
}

// Compute full-tenant lease amount for a property within a date period.
// Handles pro-rata (partial overlap between period and agreement dates).
// Returns {days,rooms,rate,total,company,effectiveFrom,effectiveTo} or null if not applicable.
function computeFullTenantForPeriod(property,fromDate,toDate){
  if(!property)return null;
  const company=(property.FullTenant_Company||'').trim();
  if(!company)return null;
  const propertyRate=Number(property.FullTenant_RatePerRoom)||0;
  // Rate unit: 'Per day' (default, legacy) or 'Per month'
  const rateUnitRaw=(property.FullTenant_RateUnit||'Per day').toString().toLowerCase().trim();
  const isMonthly=rateUnitRaw.indexOf('month')>=0;
  const agreementStart=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):new Date(1970,0,1);
  const agreementEnd=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  // Effective overlap between period and agreement
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  // Count days (inclusive)
  const days=Math.floor((effTo-effFrom)/86400000)+1;
  // Rooms on this property
  const propRooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(property.id));
  const rooms=propRooms.length;
  if(rooms===0)return null;
  // Two pricing models:
  // A) Property has FullTenant_RatePerRoom set → uniform rate × rooms (Rigg 44 style)
  // B) Property rate is empty → sum each room's FullTenant_RoomPrice (Strandveien style, per-room)
  const useUniformRate=propertyRate>0;
  const sumRoomFullTenantPrices=propRooms.reduce((s,r)=>s+(Number(r.FullTenant_RoomPrice)||0),0);
  const usePerRoomRates=!useUniformRate&&sumRoomFullTenantPrices>0;
  if(!useUniformRate&&!usePerRoomRates)return null;
  let total,unitLabel,detailLabel,rate;
  // Pre-compute month fraction for both monthly modes
  let monthFraction=0;
  const breakdown=[];
  if(isMonthly){
    let cursor=new Date(effFrom.getFullYear(),effFrom.getMonth(),1);
    while(cursor<=effTo){
      const monthStart=new Date(cursor.getFullYear(),cursor.getMonth(),1);
      const monthEnd=new Date(cursor.getFullYear(),cursor.getMonth()+1,0,23,59,59);
      const periodInMonthStart=new Date(Math.max(monthStart.getTime(),effFrom.getTime()));
      const periodInMonthEnd=new Date(Math.min(monthEnd.getTime(),effTo.getTime()));
      if(periodInMonthStart<=periodInMonthEnd){
        const daysInMonth=monthEnd.getDate();
        const periodDaysInMonth=Math.floor((periodInMonthEnd-periodInMonthStart)/86400000)+1;
        monthFraction+=periodDaysInMonth/daysInMonth;
        breakdown.push(periodDaysInMonth+'/'+daysInMonth);
      }
      cursor=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
    }
  }
  if(useUniformRate){
    rate=propertyRate;
    if(isMonthly){
      total=Math.round(rate*rooms*monthFraction*100)/100;
      unitLabel='/mnd';
      detailLabel=rooms+' rom × '+rate.toLocaleString('nb-NO')+' kr/mnd × '+monthFraction.toFixed(3)+' mnd ('+breakdown.join(' + ')+')';
    }else{
      total=Math.round(rate*rooms*days*100)/100;
      unitLabel='/dag';
      detailLabel=rooms+' rom × '+rate.toLocaleString('nb-NO')+' kr/dag × '+days+' dager';
    }
  }else{
    // Per-room pricing model: sum each room's FullTenant_RoomPrice
    rate=sumRoomFullTenantPrices; // total per unit (day or month) for all rooms combined
    const roomsWithRate=propRooms.filter(r=>Number(r.FullTenant_RoomPrice)>0).length;
    if(isMonthly){
      total=Math.round(sumRoomFullTenantPrices*monthFraction*100)/100;
      unitLabel='/mnd (per rom)';
      detailLabel='Sum '+roomsWithRate+'/'+rooms+' rom-priser ('+sumRoomFullTenantPrices.toLocaleString('nb-NO')+' kr/mnd) × '+monthFraction.toFixed(3)+' mnd';
    }else{
      total=Math.round(sumRoomFullTenantPrices*days*100)/100;
      unitLabel='/dag (per rom)';
      detailLabel='Sum '+roomsWithRate+'/'+rooms+' rom-priser ('+sumRoomFullTenantPrices.toLocaleString('nb-NO')+' kr/dag) × '+days+' dager';
    }
  }
  return{days,rooms,rate,total,company,effectiveFrom:effFrom,effectiveTo:effTo,isMonthly,unitLabel,detailLabel,usePerRoomRates};
}

function renderRow(room,booking){
  const n=booking?booking.Person_Name:'';const c=booking?(booking.Company||''):'';
  // For empty rooms: find next upcoming booking, full-tenant, or long-term contract
  let emptyCell='<span class="empty-text">—</span>';
  if(!booking){
    const fullTenant=getRoomFullTenant(room);
    const longTerm=getActiveLongTermContract(room);
    const reserveLabel=fullTenant?fullTenant.company:(longTerm?longTerm.company:null);
    if(reserveLabel){
      emptyCell='<span style="color:#EF9F27;font-style:italic">🔒 Reservert '+escapeHtml(reserveLabel)+'</span>';
    }
    const upcoming=findNextUpcomingForRoom(room.id);
    if(upcoming){
      emptyCell=(reserveLabel?'<span style="color:#EF9F27;font-style:italic">🔒 '+escapeHtml(reserveLabel)+'</span>':'<span class="empty-text">—</span>')+' <span style="font-size:10px;color:#2C7A7B;font-style:italic" title="Upcoming booking">📅 '+escapeHtml(upcoming.Person_Name||'')+(upcoming.Check_In?' · '+formatDate(upcoming.Check_In):'')+'</span>';
    }
  }
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+doorTagBtn(booking)+'</td><td>'+cleanBtn(booking)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+'</td>'
    +'<td>'+(n?guestMarkedName(n):emptyCell)+(booking&&booking.Notes?'<span class="note-dot"></span>':'')+'</td>'
    +'<td class="muted">'+c+'</td>'
    +'<td style="text-align:right;font-variant-numeric:tabular-nums">'+batCell(room.Door_Battery_Level)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums">'+datesCell(booking)+'</td></tr>';
}

// Find the soonest Upcoming booking for a given room
function findNextUpcomingForRoom(roomId){
  const now=new Date();now.setHours(0,0,0,0);
  const ups=allBookings.filter(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Upcoming'&&b.Check_In);
  ups.sort((a,b)=>new Date(a.Check_In)-new Date(b.Check_In));
  return ups.find(b=>{const d=new Date(b.Check_In);d.setHours(0,0,0,0);return d>=now})||null;
}

function renderRowWithProperty(room,booking,propName){
  const n=booking?booking.Person_Name:'';const washNext=booking?getNextWashDate(booking):'';
  let emptyCell='<span class="empty-text">—</span>';
  if(!booking){
    const fullTenant=getRoomFullTenant(room);
    const longTerm=getActiveLongTermContract(room);
    const reserveLabel=fullTenant?fullTenant.company:(longTerm?longTerm.company:null);
    if(reserveLabel){
      emptyCell='<span style="color:#EF9F27;font-style:italic">🔒 Reservert '+escapeHtml(reserveLabel)+'</span>';
    }
    const upcoming=findNextUpcomingForRoom(room.id);
    if(upcoming){
      emptyCell=(reserveLabel?'<span style="color:#EF9F27;font-style:italic">🔒 '+escapeHtml(reserveLabel)+'</span>':'<span class="empty-text">—</span>')+' <span style="font-size:10px;color:#2C7A7B;font-style:italic">📅 '+escapeHtml(upcoming.Person_Name||'')+(upcoming.Check_In?' · '+formatDate(upcoming.Check_In):'')+'</span>';
    }
  }
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+cleanBtn(booking)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+'</td>'
    +'<td class="muted" style="font-size:11px">'+propName+'</td>'
    +'<td>'+(n?guestMarkedName(n):emptyCell)+'</td>'
    +'<td>'+washNext+'</td>'
    +'<td style="font-variant-numeric:tabular-nums">'+(booking?datesCell(booking):'')+'</td></tr>';
}

function renderFloors(){
  const sourceBk=(activeFilter==='dirty')?allBookings:bookings;
  const bMap={};
  sourceBk.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&(!bMap[rid]||b.Status==='Active'))bMap[rid]=b});
  const cols=7;
  const f1=getFilteredRoomsForFloor(1).sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  const f2=getFilteredRoomsForFloor(2).sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  const allF1=rooms.filter(r=>r.Floor===1||String(r.Floor)==='1');
  const allF2=rooms.filter(r=>r.Floor===2||String(r.Floor)==='2');
  const noMatch='<tr><td colspan="'+cols+'" class="loading">No matching rooms</td></tr>';

  const isAllProps=selectedProperty===null;

  const renderFn=(r)=>{
    const b=bMap[r.id];
    if(activeFilter==='dirty'||isAllProps){
      const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
      return renderRowWithProperty(r,b,prop?prop.Title:'');
    }
    return renderRow(r,b);
  };

  document.getElementById('floor1Body').innerHTML=f1.length?f1.map(renderFn).join(''):noMatch;
  document.getElementById('floor2Body').innerHTML=f2.length?f2.map(renderFn).join(''):noMatch;

  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut'].includes(activeFilter);
  if(isStatFilter||isAllProps){
    document.getElementById('floor1Sub').textContent=f1.length+' rooms — all properties';
    document.getElementById('floor2Sub').textContent=f2.length+' rooms — all properties';
  }else{
    const f1range=allF1.length?'Rooms '+allF1.sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}))[0].Title+'–'+allF1[allF1.length-1].Title:'';
    const f2range=allF2.length?'Rooms '+allF2.sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}))[0].Title+'–'+allF2[allF2.length-1].Title:'';
    document.getElementById('floor1Sub').textContent=activeFilter?f1.length+' of '+allF1.length+' rooms':f1range;
    document.getElementById('floor2Sub').textContent=activeFilter?f2.length+' of '+allF2.length+' rooms':f2range;
  }
}

function updateStats(){
  // v14.5.7: All stat cards count across ALL assigned properties (regardless of selected property)
  const assignedPropIds=new Set(properties.map(p=>String(p.id)));
  const allAssignedRooms=allRooms.filter(r=>assignedPropIds.has(String(r.PropertyLookupId)));
  const allAssignedRoomIds=new Set(allAssignedRooms.map(r=>r.id));
  const tr=allAssignedRooms.length;
  // Active bookings (current) across all
  const today=new Date();today.setHours(0,0,0,0);
  const occupiedRoomIds=new Set();
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!allAssignedRoomIds.has(rid))return;
    // Active bookings: Status='Active' OR (Status='Upcoming' with Check_In <= today)
    if(b.Status==='Active'){occupiedRoomIds.add(rid);return}
    if(b.Status==='Upcoming'&&b.Check_In){
      const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
      if(ci.getTime()<=today.getTime())occupiedRoomIds.add(rid);
    }
  });
  document.getElementById('statCheckedIn').textContent=occupiedRoomIds.size+' / '+tr;
  document.getElementById('statEmpty').textContent=tr-occupiedRoomIds.size;
  // Dirty: across all
  const allDirtyRoomIds=new Set();
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!allAssignedRoomIds.has(rid))return;
    if(b.Cleaning_Status==='Dirty'&&(b.Status==='Active'||b.Status==='Upcoming'))allDirtyRoomIds.add(rid);
    if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out);if(w.some(x=>x.isToday))allDirtyRoomIds.add(rid)}
  });
  document.getElementById('statDirty').textContent=allDirtyRoomIds.size;
  // Door tag: across all
  document.getElementById('statDoorTag').textContent=allBookings.filter(b=>{const rid=String(b.RoomLookupId||'');return allAssignedRoomIds.has(rid)&&b.Door_Tag_Status==='Needs-print'&&(b.Status==='Active'||b.Status==='Upcoming')}).length;
  // Battery: across all
  document.getElementById('statBattery').textContent=allAssignedRooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30).length;
  // Overdue: across all
  const overdueBookings=allBookings.filter(b=>{const rid=String(b.RoomLookupId||'');return allAssignedRoomIds.has(rid)});
  const overdueCheckInCount=overdueBookings.filter(b=>isOverdueCheckIn(b)).length;
  const overdueCheckOutCount=overdueBookings.filter(b=>isOverdueCheckOut(b)).length;
  document.getElementById('statOverdueCheckIn').textContent=overdueCheckInCount;
  document.getElementById('statOverdueCheckOut').textContent=overdueCheckOutCount;
  const ciBox=document.getElementById('statOverdueCheckInBox');
  if(ciBox)ciBox.style.cssText=overdueCheckInCount>0?'background:rgba(209,67,67,.10);border-color:#D14343':'';
  const coBox=document.getElementById('statOverdueCheckOutBox');
  if(coBox)coBox.style.cssText=overdueCheckOutCount>0?'background:rgba(209,67,67,.10);border-color:#D14343':'';

  // v14.5.7: Occupancy across ALL properties — month-to-date
  // Counts a room as "occupied" on a day if any of:
  //  1. An Active or Completed booking covers that day
  //  2. The property has an active full-tenant agreement that day
  //  3. The room has an active long-term contract that day
  const now=new Date();const curMonth=now.getMonth();const curYear=now.getFullYear();
  const todayDate=now.getDate();
  const monthStart=new Date(curYear,curMonth,1);monthStart.setHours(0,0,0,0);
  const todayEnd=new Date(curYear,curMonth,todayDate);todayEnd.setHours(0,0,0,0);
  const totalDaysSoFar=todayDate;
  let occupiedRoomDays=0;
  allAssignedRooms.forEach(room=>{
    const property=properties.find(p=>String(p.id)===String(room.PropertyLookupId));
    // For each day in the month so far, check if room is occupied
    for(let d=1;d<=todayDate;d++){
      const checkDate=new Date(curYear,curMonth,d);checkDate.setHours(0,0,0,0);
      let occupied=false;
      // 1. Full-tenant agreement
      if(property){
        const ftCompany=(property.FullTenant_Company||'').trim();
        const ftRate=Number(property.FullTenant_RatePerRoom)||0;
        if(ftCompany&&ftRate>0){
          const ftStart=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):null;
          const ftEnd=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):null;
          if((!ftStart||checkDate>=ftStart)&&(!ftEnd||checkDate<=ftEnd))occupied=true;
        }
      }
      // 2. Long-term contract on the room
      if(!occupied){
        const ltCompany=(room.LongTerm_Company||'').trim();
        const ltPrice=Number(room.LongTerm_Price)||0;
        if(ltCompany&&ltPrice>0){
          const ltStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):null;
          const ltEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):null;
          if((!ltStart||checkDate>=ltStart)&&(!ltEnd||checkDate<=ltEnd))occupied=true;
        }
      }
      // 3. Actual booking
      if(!occupied){
        for(let i=0;i<allBookings.length;i++){
          const b=allBookings[i];
          if(String(b.RoomLookupId)!==String(room.id))continue;
          if(b.Status==='Cancelled')continue;
          if(!b.Check_In)continue;
          const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
          const co=b.Check_Out?new Date(b.Check_Out):now;
          if(co instanceof Date)co.setHours(0,0,0,0);
          if(checkDate>=ci&&checkDate<=co){occupied=true;break}
        }
      }
      if(occupied)occupiedRoomDays++;
    }
  });
  const totalPossible=tr*totalDaysSoFar;
  const occPct=totalPossible>0?Math.round(occupiedRoomDays/totalPossible*100):0;
  document.getElementById('statOccupancy').textContent=occPct+'%';
}

// --- DETAIL PANEL ---
function showDetail(roomId){
  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut'].includes(activeFilter);
  const room=isStatFilter?allRooms.find(r=>r.id===roomId):rooms.find(r=>r.id===roomId);
  if(!room)return;
  const sourceBk=isStatFilter?allBookings:bookings;
  const booking=sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Active')
    ||sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Upcoming');
  selectedRoom=room;selectedBooking=booking;
  const p=document.getElementById('detailPanel');
  const prop=properties.find(pr=>String(pr.id)===String(room.PropertyLookupId));
  const propName=prop?prop.Title:'';

  if(!booking){
    // Check if room has an Upcoming booking (future)
    const upcoming=findNextUpcomingForRoom(room.id);
    let subHtml='Empty — '+propName;
    if(upcoming){
      const ci=upcoming.Check_In?formatDate(upcoming.Check_In):'?';
      const name=upcoming.Person_Name||'(unnamed)';
      const company=upcoming.Company?' · '+escapeHtml(upcoming.Company):'';
      subHtml='<div>Empty — '+propName+'</div>'
        +'<div style="margin-top:8px;padding:8px 10px;background:rgba(123,97,255,.08);border-left:3px solid #7B61FF;border-radius:4px;font-size:12px">'
        +'📅 <strong>Upcoming booking:</strong> '+escapeHtml(name)+company+' · Check-in <strong>'+ci+'</strong>'
        +'</div>';
    }
    let actions=(can('edit_bookings')?'<button class="primary" onclick="openNewBooking(\''+room.id+'\')">Create booking</button>':'');
    if(upcoming&&can('edit_bookings')){
      actions+='<button onclick="openEditBooking(\''+upcoming.id+'\')" style="background:#7B61FF;color:#fff;border-color:#7B61FF">See Upcoming</button>';
    }
    if(upcoming&&can('cancel_bookings')){
      actions+='<button class="danger" onclick="cancelBookingConfirmed(\''+upcoming.id+'\')">Cancel Upcoming</button>';
    }
    actions+='<button onclick="closeDetail()">Close</button>';
    p.innerHTML='<div class="detail-grid"><div class="detail-main"><div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">'+subHtml+'</div></div><div class="detail-actions">'+actions+'</div></div>';
  }else{
    const dt={'None':'—','Needs-print':'✕ Needs print','Printed':'✓ Printed'}[booking.Door_Tag_Status]||'—';
    const cl={'None':'—','Dirty':'● Needs cleaning','Clean':'● Clean'}[booking.Cleaning_Status]||'—';
    const washHtml=getWashScheduleHtml(booking);
    let infoHtml='';
    if(can('view_bookings')){
      // Look up contact info from Persons list
      const person=allPersons.find(p=>(p.Title||'').toLowerCase()===(booking.Person_Name||'').toLowerCase()
        ||(p.Name||'').toLowerCase()===(booking.Person_Name||'').toLowerCase());
      const phone=person?(person.Mobile||person.Phone||person.Telefon||''):'';
      const email=person?(person.Email||''):'';
      const addr=person?(person.Address||''):'';
      // v14.5.7: Overdue banner at top of detail
      let overdueBanner='';
      if(isOverdueCheckIn(booking)){
        const d=daysOverdueCheckIn(booking);
        overdueBanner='<div style="background:rgba(209,67,67,.10);border-left:3px solid #D14343;padding:10px 14px;margin-bottom:12px;border-radius:6px;display:flex;align-items:center;gap:10px"><div style="flex:1;font-size:13px;color:#A32D2D"><strong>⚠ Overdue check-in:</strong> This booking should have been checked in '+d+' day'+(d===1?'':'s')+' ago.</div>'+(can('checkin_out')?'<button onclick="checkIn(\''+booking.id+'\')" style="padding:6px 14px;background:#1D9E75;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-family:inherit;font-weight:500">✓ Check in now</button>':'')+'</div>';
      }else if(isOverdueCheckOut(booking)){
        const d=daysOverdueCheckOut(booking);
        overdueBanner='<div style="background:rgba(209,67,67,.10);border-left:3px solid #D14343;padding:10px 14px;margin-bottom:12px;border-radius:6px;display:flex;align-items:center;gap:10px"><div style="flex:1;font-size:13px;color:#A32D2D"><strong>⚠ Overdue check-out:</strong> This booking should have been checked out '+d+' day'+(d===1?'':'s')+' ago.</div>'+(can('checkin_out')?'<button onclick="openCheckoutModal(\''+booking.id+'\')" style="padding:6px 14px;background:#1D9E75;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-family:inherit;font-weight:500">✓ Mark as Completed</button>':'')+'</div>';
      }
      infoHtml=overdueBanner+'<div class="detail-name">'+booking.Person_Name+'</div>'
        +'<div class="detail-sub">Room '+room.Title+' · '+(booking.Company||'')+' · '+propName+'</div>'
        +'<table class="detail-info">'
        +(phone?'<tr><td>Mobile</td><td><a href="tel:'+phone+'" style="color:var(--accent)">'+phone+'</a></td></tr>':'')
        +(email?'<tr><td>Email</td><td><a href="mailto:'+email+'" style="color:var(--accent)">'+email+'</a></td></tr>':'')
        +(addr?'<tr><td>Address</td><td style="white-space:pre-line">'+addr+'</td></tr>':'')
        +'<tr><td>Check-in</td><td>'+formatDate(booking.Check_In)+'</td></tr>'
        +'<tr><td>Check-out</td><td>'+(booking.Check_Out?formatDate(booking.Check_Out):'Open-ended')+'</td></tr>'
        +'<tr><td>Status</td><td>'+booking.Status+'</td></tr>'
        +'<tr><td>Door tag</td><td>'+dt+'</td></tr>'
        +'<tr><td>Cleaning</td><td>'+cl+'</td></tr>'
        +(booking.Notes?'<tr><td>Notes</td><td>'+booking.Notes+'</td></tr>':'')
        +'<tr><td>Battery</td><td>'+(typeof renderBatteryStatusHtml==='function'?renderBatteryStatusHtml(room):'(n/a)')+'</td></tr>'
        +((booking.Continuation===true||booking.Continuation==='true'||booking.Continuation===1)?'<tr><td>🔗 Continuation</td><td><span style="color:#7B61FF;font-weight:500">Yes — utvask skipped</span></td></tr>':'')
        +((booking.Billing_Company||'').trim()&&(booking.Billing_Company||'').trim()!==(booking.Company||'').trim()?'<tr><td>💳 Billing</td><td><span style="color:var(--accent);font-weight:500">'+escapeHtml(booking.Billing_Company)+'</span> <span style="color:var(--text-tertiary);font-size:11px">(rate &amp; invoice follow billing company)</span></td></tr>':'')
        +'</table>'
        +(can('view_prices')?(function(){
          const cost=calcBookingCost(booking,propName);
          // Always show the pricing block, even when no rate is set, so user sees the near-miss warning
          let extra='';
          if(cost.matchedName&&cost.matchedName.toLowerCase()!==(booking.Person_Name||'').toLowerCase()){
            extra='<div style="font-size:11px;color:var(--text-tertiary);margin-top:4px">Matched rate row: "'+cost.matchedName+'"</div>';
          }
          if(cost.nearMiss){
            extra+='<div style="margin-top:6px;padding:6px 8px;background:var(--bg-warning);border:1px solid #EF9F27;border-radius:4px;color:var(--text-warning);font-size:11px">⚠ '+cost.nearMiss+'</div>';
          }
          if(!cost.rate){
            return'<div style="margin-top:10px;padding:10px;background:var(--bg-secondary);border-radius:var(--radius-md);font-size:12px">'
              +'<strong>Pricing</strong> <span class="muted">(no rate set)</span>'
              +extra+'</div>';
          }
          return'<div style="margin-top:10px;padding:10px;background:var(--bg-secondary);border-radius:var(--radius-md);font-size:12px">'
            +'<strong>Pricing</strong> <span class="muted">('+cost.source+')</span><br>'
            +'Rate: <strong>'+cost.rate+' kr/night</strong> × '+cost.nights+' nights = <strong>'+cost.total.toLocaleString('nb-NO')+' kr</strong>'
            +extra+'</div>';
        })():'')
        +washHtml;
    }else{
      infoHtml='<div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">'+cl+'</div>'+(typeof renderBatteryStatusHtml==='function'?renderBatteryStatusHtml(room):'')+washHtml;
    }
    let btns='';
    if(can('edit_bookings'))btns+='<button onclick="openEditBooking(\''+booking.id+'\')">Edit booking</button>';
    // "Add to Guests" button — only if not already in Persons list (fuzzy match)
    if(can('edit_bookings')&&booking.Person_Name){
      const inList=allPersons.some(p=>{
        const pn=(p.Name||p.Title||'').toLowerCase().trim();
        const bn=(booking.Person_Name||'').toLowerCase().trim();
        if(pn===bn)return true;
        const wa=pn.split(/[\s,]+/).filter(w=>w.length>1);
        const wb=bn.split(/[\s,]+/).filter(w=>w.length>1);
        if(wa.length<2||wb.length<2)return false;
        return wa.every(w=>bn.indexOf(w)>=0)||wb.every(w=>pn.indexOf(w)>=0);
      });
      if(!inList){
        btns+='<button onclick="addBookingToGuests(\''+booking.id+'\')" style="background:var(--bg-success);color:var(--text-success);border-color:var(--accent)">+ Add to Guests</button>';
      }
    }
    if(can('print_doortag'))btns+='<button onclick="printDoorTag(\''+booking.id+'\')">Print door tag</button>';
    // Door code buttons (Phase 1: per-room, manual)
    btns+='<button onclick="window._currentDoorCodeBookingId=\''+booking.id+'\';showRoomDoorCode(\''+booking.id+'\')" style="background:rgba(239,159,39,.1);color:#a76800;border-color:#EF9F27" title="Vis nåværende dørkode for rommet">🔑 Vis kode</button>';
    btns+='<button onclick="window._currentDoorCodeBookingId=\''+booking.id+'\';generateRoomDoorCode(\''+booking.id+'\')" style="background:rgba(239,159,39,.1);color:#a76800;border-color:#EF9F27" title="Generer ny 6-sifret dørkode for rommet">🔑 Generer kode</button>';
    // Messaging buttons
    btns+='<button onclick="copyBookingSMS(\''+booking.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Kopier SMS-tekst til utklippstavle">📱 Kopier SMS</button>';
    btns+='<button onclick="openBookingSMS(\''+booking.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Åpne SMS-app med ferdig tekst">📱 Send SMS</button>';
    btns+='<button onclick="copyBookingEmail(\''+booking.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Kopier e-post-tekst til utklippstavle">📧 Kopier e-post</button>';
    btns+='<button onclick="openBookingEmail(\''+booking.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Åpne e-postklient med ferdig tekst">📧 Send e-post</button>';
    if(booking.Status==='Upcoming'&&can('checkin_out'))btns+='<button class="primary" onclick="checkIn(\''+booking.id+'\')">Check in</button>';
    if(booking.Status==='Active'&&can('checkin_out'))btns+='<button class="primary" style="background:#EF9F27;border-color:#EF9F27" onclick="checkOut(\''+booking.id+'\')">Check out</button>';
    if(can('cleaning')){
      if(booking.Cleaning_Status==='Dirty')btns+='<button class="primary" onclick="markClean(\''+booking.id+'\')">Mark as cleaned ✓</button>';
      else btns+='<button onclick="markDirty(\''+booking.id+'\')">Mark as dirty</button>';
    }
    if(can('cancel_bookings'))btns+='<button class="danger" onclick="cancelBooking(\''+booking.id+'\')">Cancel booking</button>';
    btns+='<button onclick="closeDetail()">Close</button>';
    p.innerHTML='<div class="detail-grid"><div class="detail-main">'+infoHtml+'</div><div class="detail-actions">'+btns+'</div></div>';
  }
  p.classList.add('open');
  // Scroll to bring detail panel into view (it's now above the floor tables)
  setTimeout(()=>{p.scrollIntoView({behavior:'smooth',block:'nearest'})},50);
}
function closeDetail(){document.getElementById('detailPanel').classList.remove('open');selectedRoom=null;selectedBooking=null}

// --- STATUS TOGGLES ---
async function cycleDT(e,id){
  e.stopPropagation();if(!can('doortag'))return;
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const c={'None':'Needs-print','Needs-print':'Printed','Printed':'None'};const ns=c[b.Door_Tag_Status||'None'];
  try{await updateListItem('Bookings',id,{Door_Tag_Status:ns});b.Door_Tag_Status=ns;renderFloors();updateStats()}catch(er){console.error(er);alert('Failed')}
}
async function cycleCS(e,id){
  e.stopPropagation();if(!can('cleaning'))return;
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const c={'None':'Dirty','Dirty':'Clean','Clean':'None'};const ns=c[b.Cleaning_Status||'None'];
  try{await updateListItem('Bookings',id,{Cleaning_Status:ns});b.Cleaning_Status=ns;renderFloors();updateStats()}catch(er){console.error(er);alert('Failed')}
}
async function markClean(id){try{await updateListItem('Bookings',id,{Cleaning_Status:'Clean'});const l=allBookings.find(x=>x.id===id);if(l)l.Cleaning_Status='Clean';closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}}
async function markDirty(id){try{await updateListItem('Bookings',id,{Cleaning_Status:'Dirty'});const l=allBookings.find(x=>x.id===id);if(l)l.Cleaning_Status='Dirty';closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}}

// --- CHECK IN/OUT/CANCEL ---
async function checkIn(id){
  if(!confirm('Check in this guest?'))return;
  try{const now=new Date().toISOString();await updateListItem('Bookings',id,{Status:'Active',Check_In:now});const l=allBookings.find(x=>x.id===id);if(l){l.Status='Active';l.Check_In=now}closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}
}
function checkOut(id){
  const b=allBookings.find(x=>x.id===id);if(!b)return;checkoutBookingId=id;
  document.getElementById('fCheckOutDate').value=new Date().toISOString().split('T')[0];
  document.getElementById('checkoutGuestName').textContent=b.Person_Name||'Guest';
  document.getElementById('checkoutModal').classList.add('open');
}
function closeCheckoutModal(){document.getElementById('checkoutModal').classList.remove('open');checkoutBookingId=null}
async function confirmCheckout(){
  if(!checkoutBookingId)return;const dateVal=document.getElementById('fCheckOutDate').value;if(!dateVal){alert('Select a date');return}
  const btn=document.getElementById('checkoutConfirmBtn');btn.disabled=true;btn.textContent='Processing...';
  try{await updateListItem('Bookings',checkoutBookingId,{Status:'Completed',Cleaning_Status:'Dirty',Check_Out:dateVal+'T12:00:00Z'});const l=allBookings.find(x=>x.id===checkoutBookingId);if(l){l.Status='Completed';l.Cleaning_Status='Dirty';l.Check_Out=dateVal+'T12:00:00Z'}closeCheckoutModal();closeDetail();refreshLocal();loadData()}catch(e){alert('Failed: '+e.message)}finally{btn.disabled=false;btn.textContent='Confirm check-out'}
}
async function cancelBooking(id){
  return cancelBookingConfirmed(id);
}

// Detailed cancel confirmation — shows guest name and dates
async function cancelBookingConfirmed(id){
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const name=b.Person_Name||'(unnamed)';
  const ci=b.Check_In?formatDate(b.Check_In):'?';
  const co=b.Check_Out?formatDate(b.Check_Out):'Open-ended';
  const company=b.Company?' ('+b.Company+')':'';
  const msg='Are you sure you want to cancel this booking?\n\n'
    +'Guest: '+name+company+'\n'
    +'Check-in: '+ci+'\n'
    +'Check-out: '+co+'\n'
    +'Status: '+b.Status+'\n\n'
    +'This cannot be undone from the app — you would need to edit the booking manually to reactivate it.';
  if(!confirm(msg))return;
  try{
    await updateListItem('Bookings',id,{Status:'Cancelled'});
    const l=allBookings.find(x=>x.id===id);if(l)l.Status='Cancelled';
    closeDetail();refreshLocal();loadData();
  }catch(e){alert('Failed: '+e.message)}
}

// --- BOOKING MODAL ---
function populateRoomSelect(preselectedRoomId){
  const sel=document.getElementById('fRoom');
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  sel.innerHTML=sorted.map(r=>'<option value="'+r.id+'"'+(r.id===preselectedRoomId?' selected':'')+'>'+r.Title+' (Floor '+r.Floor+')</option>').join('');
  sel.onchange=()=>{const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:''};
  const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:'';
}

// Returns id of first available room for given check-in/check-out, or null
function findFirstAvailableRoomId(checkInStr,checkOutStr){
  if(!checkInStr)return null;
  const newIn=new Date(checkInStr+'T00:00:00');newIn.setHours(0,0,0,0);
  const newOut=checkOutStr?new Date(checkOutStr+'T00:00:00'):null;
  if(newOut)newOut.setHours(0,0,0,0);
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  for(const room of sorted){
    const hasConflict=allBookings.some(b=>{
      if(b.Status==='Cancelled'||b.Status==='Completed')return false;
      if(String(b.RoomLookupId)!==String(room.id))return false;
      const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);
      const bOut=b.Check_Out?new Date(b.Check_Out):null;if(bOut)bOut.setHours(0,0,0,0);
      if(!bOut)return newIn>=bIn||(newOut?newOut>bIn:true);
      if(!newOut)return bOut>newIn||bIn>=newIn;
      return newIn<bOut&&newOut>bIn;
    });
    if(!hasConflict)return room.id;
  }
  return null;
}
function openNewBooking(preselectedRoomId){
  ensureMainView();
  editingBookingId=null;document.getElementById('bookingModalTitle').textContent='New booking';
  document.getElementById('bookingSaveBtn').textContent='Create booking';
  const todayStr=toISODate(new Date());
  // If no room pre-selected, find first available for today
  let roomToSelect=preselectedRoomId||'';
  let autoSelected=false;
  if(!roomToSelect){
    const auto=findFirstAvailableRoomId(todayStr,'');
    if(auto){roomToSelect=auto;autoSelected=true}
  }
  populateRoomSelect(roomToSelect);
  document.getElementById('fName').value='';document.getElementById('fCompany').value='';
  document.getElementById('fBillingCompany').value='';
  const cw=document.getElementById('fCompanyWarn');if(cw)cw.innerHTML='';
  const bw=document.getElementById('fBillingCompanyWarn');if(bw)bw.innerHTML='';
  document.getElementById('fCheckIn').value=todayStr;document.getElementById('fCheckOut').value='';
  // Default to Active if check-in is today, Upcoming otherwise
  document.getElementById('fStatus').value='Active';
  document.getElementById('fNotes').value='';
  document.getElementById('fIncludeCheckoutFee').checked=true;
  document.getElementById('fContinuation').checked=false;
  document.getElementById('fNameInfo').innerHTML='';
  // Show auto-select hint in room-info area
  const roomInfo=document.getElementById('fRoomInfo');
  if(autoSelected&&roomToSelect){
    const r=rooms.find(rm=>rm.id===roomToSelect);
    if(r){
      roomInfo.textContent='✓ Auto-selected first available: Room '+r.Title;
      roomInfo.style.color='var(--text-success)';
    }
  }else if(!roomToSelect){
    roomInfo.textContent='';
  }else{
    roomInfo.textContent='';
  }
  document.getElementById('fOverlapWarning').style.display='none';
  attachOverlapListeners();
  attachStatusAutoSelect();
  checkBookingOverlap();
  const modal=document.getElementById('bookingModal');
  modal.classList.add('open');
  const modalContent=modal.querySelector('.modal');
  if(modalContent)modalContent.scrollTop=0;
  modal.scrollTop=0;
}
function openEditBooking(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;editingBookingId=bookingId;
  document.getElementById('bookingModalTitle').textContent='Edit booking';
  document.getElementById('bookingSaveBtn').textContent='Save changes';
  populateRoomSelect(String(b.RoomLookupId));
  document.getElementById('fName').value=b.Person_Name||'';document.getElementById('fCompany').value=b.Company||'';
  document.getElementById('fBillingCompany').value=b.Billing_Company||'';
  if(typeof checkCompanyRegistration==='function'){
    checkCompanyRegistration(b.Company||'','fCompanyWarn');
    checkCompanyRegistration(b.Billing_Company||'','fBillingCompanyWarn');
  }
  document.getElementById('fCheckIn').value=b.Check_In?toISODate(b.Check_In):'';
  document.getElementById('fCheckOut').value=b.Check_Out?toISODate(b.Check_Out):'';
  document.getElementById('fStatus').value=b.Status||'Upcoming';document.getElementById('fNotes').value=b.Notes||'';
  // Checkout fee: default true if not explicitly stored as false
  const fee=b.Include_Checkout_Fee;
  document.getElementById('fIncludeCheckoutFee').checked=(fee===undefined||fee===null||fee===true||fee==='true'||fee===1);
  const cont=b.Continuation;
  document.getElementById('fContinuation').checked=(cont===true||cont==='true'||cont===1);
  document.getElementById('fNameInfo').innerHTML='';
  document.getElementById('fOverlapWarning').style.display='none';
  attachOverlapListeners();
  checkBookingOverlap();
  const modal=document.getElementById('bookingModal');
  modal.classList.add('open');
  const modalContent=modal.querySelector('.modal');
  if(modalContent)modalContent.scrollTop=0;
  modal.scrollTop=0;
}
function closeBookingModal(){document.getElementById('bookingModal').classList.remove('open');editingBookingId=null}

// Attach change-listeners to room/date fields to check for overlap in real time
let _overlapAttached=false;
function attachOverlapListeners(){
  if(_overlapAttached)return;_overlapAttached=true;
  ['fRoom','fCheckIn','fCheckOut'].forEach(id=>{
    const el=document.getElementById(id);
    if(el)el.addEventListener('change',checkBookingOverlap);
  });
}

// Auto-set Status based on Check-in date (only for NEW bookings, not edits)
let _statusAutoAttached=false;
function attachStatusAutoSelect(){
  if(_statusAutoAttached)return;_statusAutoAttached=true;
  const ciEl=document.getElementById('fCheckIn');
  if(!ciEl)return;
  ciEl.addEventListener('change',()=>{
    if(editingBookingId)return; // don't override status when editing existing
    const val=ciEl.value;if(!val)return;
    const sel=document.getElementById('fStatus');if(!sel)return;
    const today=new Date();today.setHours(0,0,0,0);
    const picked=new Date(val+'T00:00:00');picked.setHours(0,0,0,0);
    // If check-in is today or in the past → Active. If in future → Upcoming.
    sel.value=picked<=today?'Active':'Upcoming';
  });
}

// Check if current modal values overlap with other bookings on same room
function checkBookingOverlap(){
  const warn=document.getElementById('fOverlapWarning');if(!warn)return;
  const roomId=document.getElementById('fRoom').value;
  const ciStr=document.getElementById('fCheckIn').value;
  const coStr=document.getElementById('fCheckOut').value;
  if(!roomId||!ciStr){warn.style.display='none';return}
  const ci=new Date(ciStr+'T00:00:00');const co=coStr?new Date(coStr+'T23:59:59'):null;
  // Find other bookings on this room (not this one if editing)
  const conflicts=allBookings.filter(b=>{
    if(b.id===editingBookingId)return false;
    if(String(b.RoomLookupId)!==String(roomId))return false;
    if(b.Status==='Cancelled'||b.Status==='Completed')return false;
    if(!b.Check_In)return false;
    const bi=new Date(b.Check_In);const bo=b.Check_Out?new Date(b.Check_Out):null;
    // Overlap: ci <= bo AND (co >= bi OR co is null)
    if(co){
      if(bo)return ci<=bo&&co>=bi;
      return co>=bi; // other is open-ended, overlap if our end is after its start
    }else{
      // our booking is open-ended
      if(bo)return ci<=bo;
      return true; // both open-ended = definite overlap
    }
  });
  if(!conflicts.length){warn.style.display='none';return}
  const lines=conflicts.map(c=>{
    const name=c.Person_Name||'(unnamed)';
    const period=formatDate(c.Check_In)+' → '+(c.Check_Out?formatDate(c.Check_Out):'Open');
    return '• <strong>'+escapeHtml(name)+'</strong> ('+c.Status+') · '+period;
  });
  warn.innerHTML='<strong>⚠ Double-booking warning</strong> — this room already has '+conflicts.length+' overlapping booking'+(conflicts.length!==1?'s':'')+':<br>'+lines.join('<br>');
  warn.style.display='block';
}

function findAvailableRoom(){
  const checkIn=document.getElementById('fCheckIn').value;
  const checkOut=document.getElementById('fCheckOut').value;
  const info=document.getElementById('fRoomInfo');

  if(!checkIn){info.textContent='Set check-in date first';info.style.color='var(--text-danger)';return}

  const newIn=new Date(checkIn);newIn.setHours(0,0,0,0);
  const newOut=checkOut?new Date(checkOut):null;
  if(newOut)newOut.setHours(0,0,0,0);

  // Find rooms that are NOT occupied during the given dates
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));

  for(const room of sorted){
    const hasConflict=allBookings.some(b=>{
      if(b.Status==='Cancelled'||b.Status==='Completed')return false;
      if(String(b.RoomLookupId)!==String(room.id))return false;

      const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);
      const bOut=b.Check_Out?new Date(b.Check_Out):null;
      if(bOut)bOut.setHours(0,0,0,0);

      // Check overlap
      if(!bOut){return newIn>=bIn||(newOut?newOut>bIn:true)}
      if(!newOut){return bOut>newIn||bIn>=newIn}
      return newIn<bOut&&newOut>bIn;
    });

    if(!hasConflict){
      // Found available room — select it and trigger change event
      const sel=document.getElementById('fRoom');
      sel.value=room.id;
      const rm=rooms.find(r=>r.id===room.id);
      document.getElementById('fFloor').value=rm?rm.Floor:'';
      // Dispatch change event so overlap warning and other listeners update
      sel.dispatchEvent(new Event('change',{bubbles:true}));
      info.textContent='✓ Room '+room.Title+' (Floor '+room.Floor+') — selected';
      info.style.color='var(--text-success)';
      return;
    }
  }

  // No room available
  info.textContent='✕ No rooms available for these dates';
  info.style.color='var(--text-danger)';
}

async function saveBooking(){
  const roomId=document.getElementById('fRoom').value;
  const name=document.getElementById('fName').value.trim();
  const company=document.getElementById('fCompany').value.trim();
  const billingCompany=document.getElementById('fBillingCompany').value.trim();
  const checkIn=document.getElementById('fCheckIn').value;
  const checkOut=document.getElementById('fCheckOut').value;
  const status=document.getElementById('fStatus').value;
  const notes=document.getElementById('fNotes').value.trim();
  const room=rooms.find(r=>r.id===roomId);
  if(!name){alert('Guest name is required');return}
  if(!checkIn){alert('Check-in date is required');return}

  // Collision check
  const newIn=new Date(checkIn);newIn.setHours(0,0,0,0);
  const newOut=checkOut?new Date(checkOut):null;if(newOut)newOut.setHours(0,0,0,0);
  const conflicts=allBookings.filter(b=>{
    if(editingBookingId&&b.id===editingBookingId)return false;
    if(b.Status==='Cancelled'||b.Status==='Completed')return false;
    if(String(b.RoomLookupId)!==String(roomId))return false;
    const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);const bOut=b.Check_Out?new Date(b.Check_Out):null;if(bOut)bOut.setHours(0,0,0,0);
    if(!bOut)return newIn>=bIn||(newOut?newOut>bIn:true);
    if(!newOut)return bOut>newIn||bIn>=newIn;
    return newIn<bOut&&newOut>bIn;
  });
  if(conflicts.length>0){
    const c=conflicts[0];
    if(!confirm('Room already booked:\n'+c.Person_Name+' ('+c.Status+')\n'+formatDate(c.Check_In)+' — '+(c.Check_Out?formatDate(c.Check_Out):'Open-ended')+'\n\nContinue anyway?'))return;
  }

  // Property_Name: find from room's property (works even in "All properties" mode)
  const roomProp=room?properties.find(pr=>String(pr.id)===String(room.PropertyLookupId)):null;
  const propNameForSave=roomProp?roomProp.Title:(selectedProperty?selectedProperty.Title:'');
  const fields={Person_Name:name,Company:company,Billing_Company:billingCompany||null,Check_In:checkIn+'T15:00:00Z',Status:status,Door_Tag_Status:'Needs-print',Cleaning_Status:'None',Property_Name:propNameForSave,Floor:room?room.Floor:1,Notes:notes||null};
  fields.Include_Checkout_Fee=document.getElementById('fIncludeCheckoutFee').checked;
  fields.Continuation=document.getElementById('fContinuation').checked;
  if(checkOut)fields.Check_Out=checkOut+'T12:00:00Z';else fields.Check_Out=null;
  fields.RoomLookupId=parseInt(roomId);
  const btn=document.getElementById('bookingSaveBtn');btn.disabled=true;btn.textContent='Saving...';
  try{
    if(editingBookingId){delete fields.Door_Tag_Status;delete fields.Cleaning_Status;await updateListItem('Bookings',editingBookingId,fields);const l=allBookings.find(x=>x.id===editingBookingId);if(l){Object.assign(l,fields);l.Check_Out=fields.Check_Out}closeBookingModal();closeDetail();refreshLocal();loadData()}
    else{await createListItem('Bookings',fields);closeBookingModal();closeDetail();await loadData()}
  }catch(e){alert('Failed: '+e.message)}finally{btn.disabled=false;btn.textContent=editingBookingId?'Save changes':'Create booking'}
}

// --- DOOR TAG PRINT ---
function printDoorTag(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));const roomTitle=room?room.Title:'?';
  const html='<div style="font-family:Arial,sans-serif;padding:40px;max-width:600px;margin:0 auto">'
    +'<div style="text-align:center;margin-bottom:30px"><div style="font-size:72px;font-weight:700;letter-spacing:2px">'+roomTitle+'</div><div style="font-size:14px;color:#888;margin-top:4px">'+(selectedProperty?selectedProperty.Title:'')+'</div></div>'
    +'<div style="border-top:2px solid #2C2C2A;padding-top:20px"><h2 style="font-size:18px;margin:0 0 16px">Welcome, '+b.Person_Name+'</h2>'
    +'<table style="font-size:14px;width:100%"><tr><td style="padding:6px 0;color:#888;width:120px">Company</td><td>'+(b.Company||'—')+'</td></tr>'
    +'<tr><td style="padding:6px 0;color:#888">Check-in</td><td>After 15:00 — '+formatDate(b.Check_In)+'</td></tr>'
    +'<tr><td style="padding:6px 0;color:#888">Check-out</td><td>Before 12:00 — '+(b.Check_Out?formatDate(b.Check_Out):'Open-ended')+'</td></tr></table></div>'
    +'<div style="margin-top:24px;padding:16px;background:#f5f4ef;border-radius:8px;font-size:13px"><strong>Room information</strong><br>The room will be washed once a week.<br>New towels every week, and new beddings biweekly.</div>'
    +'<div style="margin-top:16px;padding:16px;background:#f5f4ef;border-radius:8px;font-size:13px"><strong>Contact</strong><br>Questions? Contact Frank: +47 99 10 10 41 · frank@2gm.no</div>'
    +'<div style="text-align:center;margin-top:40px;font-size:16px;color:#888">Have a nice stay :)</div></div>';
  const w=window.open('','_blank','width=700,height=900');w.document.write('<!DOCTYPE html><html><head><title>Door Tag — Room '+roomTitle+'</title></head><body style="margin:0">'+html+'</body></html>');w.document.close();setTimeout(()=>w.print(),500);
  // Mark as printed
  const bk=allBookings.find(x=>x.id===bookingId);if(bk){updateListItem('Bookings',bookingId,{Door_Tag_Status:'Printed'}).then(()=>{bk.Door_Tag_Status='Printed';renderFloors();updateStats()}).catch(console.error)}
}

// --- VIEW SWITCHING ---
function showMainView(){currentView='main';document.getElementById('mainView').style.display='';document.getElementById('hoursView').style.display='none';document.getElementById('propertySelect').style.display='';if(selectedProperty)document.getElementById('headerTitle').textContent='2GM Eiendom AS – Booking — '+selectedProperty.Title;updateNavActiveState()}
function showHoursView(){currentView='hours';document.getElementById('mainView').style.display='none';document.getElementById('mainView').classList.remove('panel-mode');document.getElementById('incomingPanel').classList.remove('open');document.getElementById('archivePanel').classList.remove('open');const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
  document.getElementById('hoursView').style.display='';document.getElementById('propertySelect').style.display='none';document.getElementById('headerTitle').textContent='2GM Eiendom AS – Booking — Hours';updateNavActiveState()}
function ensureMainView(){if(currentView==='hours')showMainView()}

// --- FILTER ---
function toggleFilter(filter){
  if(activeFilter===filter){clearFilter();return}
  activeFilter=filter;
  document.querySelectorAll('.stat').forEach((el,i)=>{const f=['checkedIn','empty','dirty','doorTag','battery'];el.classList.toggle('active',f[i]===filter)});
  const labels={checkedIn:'Showing: Checked-in rooms',empty:'Showing: Empty rooms',dirty:'Showing: Rooms needing cleaning',doorTag:'Showing: Door tags needing print',battery:'Showing: Low battery rooms (<30%)'};
  document.getElementById('filterLabel').textContent=labels[filter]||'';
  document.getElementById('filterBar').classList.add('open');renderFloors();
  // Show action button for door-tag filter
  const actionBtn=document.getElementById('filterActionBtn');
  if(filter==='doorTag'&&can('print_doortag')){
    actionBtn.style.display='';
    actionBtn.textContent='🖨 Print all';
    actionBtn.onclick=printAllPendingDoorTags;
  }else{
    actionBtn.style.display='none';
    actionBtn.onclick=null;
  }
}
function clearFilter(){
  activeFilter=null;
  document.querySelectorAll('.stat').forEach(el=>el.classList.remove('active'));
  document.getElementById('filterBar').classList.remove('open');
  const actionBtn=document.getElementById('filterActionBtn');if(actionBtn)actionBtn.style.display='none';
  renderFloors();
}

// Print all door tags for bookings that need printing (current property or all if in All mode)
function printAllPendingDoorTags(){
  // Find rooms currently visible in the doorTag filter
  const floor1=getFilteredRoomsForFloor(1);
  const floor2=getFilteredRoomsForFloor(2);
  const visibleRoomIds=new Set([...floor1,...floor2].map(r=>r.id));
  // Find bookings with Door_Tag_Status=Needs-print on visible rooms
  const pending=allBookings.filter(b=>
    b.Door_Tag_Status==='Needs-print'&&
    (b.Status==='Active'||b.Status==='Upcoming')&&
    visibleRoomIds.has(String(b.RoomLookupId))
  );
  if(!pending.length){alert('No door tags need printing.');return}
  if(!confirm('Print '+pending.length+' door tag'+(pending.length!==1?'s':'')+'?\n\nThey will open in a new window as a single printable document.'))return;
  // Build a single combined document with page-breaks between each tag
  let html='<html><head><title>Door tags</title><style>@media print{.tag{page-break-after:always}}body{margin:0;font-family:Arial,sans-serif}.tag{padding:40px;max-width:600px;margin:0 auto}</style></head><body>';
  pending.forEach(b=>{
    const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
    const roomTitle=room?room.Title:'?';
    const prop=properties.find(p=>String(p.id)===String(room?room.PropertyLookupId:''));
    const propTitle=prop?prop.Title:'';
    html+='<div class="tag">'
      +'<div style="text-align:center;margin-bottom:30px"><div style="font-size:72px;font-weight:700;letter-spacing:2px">'+roomTitle+'</div><div style="font-size:14px;color:#888;margin-top:4px">'+propTitle+'</div></div>'
      +'<div style="border-top:2px solid #2C2C2A;padding-top:20px"><h2 style="font-size:18px;margin:0 0 16px">Welcome, '+(b.Person_Name||'')+'</h2>'
      +'<table style="font-size:14px;width:100%"><tr><td style="padding:6px 0;color:#888;width:120px">Company</td><td>'+(b.Company||'—')+'</td></tr>'
      +'<tr><td style="padding:6px 0;color:#888">Check-in</td><td>After 15:00 — '+formatDate(b.Check_In)+'</td></tr>'
      +'<tr><td style="padding:6px 0;color:#888">Check-out</td><td>'+(b.Check_Out?'Before 12:00 — '+formatDate(b.Check_Out):'Open-ended')+'</td></tr></table></div>'
      +'</div>';
  });
  html+='</body></html>';
  const w=window.open('','_blank');
  w.document.write(html);
  w.document.close();
  setTimeout(()=>w.print(),300);
  // Mark them all as printed (in background, with throttling)
  (async()=>{
    for(let i=0;i<pending.length;i++){
      try{
        await updateListItem('Bookings',pending[i].id,{Door_Tag_Status:'Printed'});
        const l=allBookings.find(x=>x.id===pending[i].id);if(l)l.Door_Tag_Status='Printed';
      }catch(e){console.error(e)}
      if(i%10===9)await new Promise(r=>setTimeout(r,300));
    }
    refreshLocal();updateStats();
  })();
}

function getFilteredRoomsForFloor(floor){
  // v14.5.7: All stat filters now show across ALL assigned properties (not just selected)
  const assignedPropIds=new Set(properties.map(p=>p.id));
  const allAssignedRooms=allRooms.filter(r=>assignedPropIds.has(String(r.PropertyLookupId)));
  // For stat filters, use cross-property source. For non-filter view, use selected property
  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut'].includes(activeFilter);
  const sourceRooms=isStatFilter?allAssignedRooms:rooms;
  let floorRooms=sourceRooms.filter(r=>r.Floor===floor||String(r.Floor)===String(floor));
  if(!activeFilter)return floorRooms;
  const sourceBookings=isStatFilter?allBookings:bookings;
  const bMap={};sourceBookings.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&(!bMap[rid]||b.Status==='Active'))bMap[rid]=b});
  switch(activeFilter){
    case 'checkedIn':return floorRooms.filter(r=>!!bMap[r.id]);
    case 'empty':return floorRooms.filter(r=>!bMap[r.id]);
    case 'dirty':return floorRooms.filter(r=>{const b=bMap[r.id];if(!b)return false;if(b.Cleaning_Status==='Dirty')return true;if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out);if(w.some(x=>x.isToday))return true}return false});
    case 'doorTag':return floorRooms.filter(r=>bMap[r.id]&&bMap[r.id].Door_Tag_Status==='Needs-print');
    case 'battery':return floorRooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30);
    case 'overdueCheckIn':return floorRooms.filter(r=>{const b=bMap[r.id];return b&&isOverdueCheckIn(b)});
    case 'overdueCheckOut':return floorRooms.filter(r=>{const b=bMap[r.id];return b&&isOverdueCheckOut(b)});
    default:return floorRooms;
  }
}

// --- COLUMN RESIZE ---
function initResize(){
  document.querySelectorAll('.th-resize').forEach(handle=>{
    handle.addEventListener('mousedown',function(e){
      e.preventDefault();e.stopPropagation();const th=this.parentElement;const table=th.closest('table');
      const startX=e.pageX;const startW=th.offsetWidth;const colIndex=Array.from(th.parentElement.children).indexOf(th);
      const col=table.querySelector('colgroup')?table.querySelector('colgroup').children[colIndex]:null;
      function onMove(ev){const diff=ev.pageX-startX;const newW=Math.max(20,startW+diff);th.style.width=newW+'px';if(col)col.style.width=newW+'px';
        const otherId=table.id==='table1'?'table2':'table1';const other=document.getElementById(otherId);
        if(other){const otherTh=other.querySelectorAll('thead th')[colIndex];const otherCol=other.querySelector('colgroup')?other.querySelector('colgroup').children[colIndex]:null;if(otherTh)otherTh.style.width=newW+'px';if(otherCol)otherCol.style.width=newW+'px'}}
      function onUp(){document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp)}
      document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
    });
  });
}

// --- HOURS REMINDER ---
async function checkHoursReminder(){
  if(!can('hours_reminder'))return;
  try{
    const hours=await getListItems('Hours');
    // Check yesterday (or last workday if today is Monday)
    const today=new Date();today.setHours(0,0,0,0);
    let checkDate=new Date(today);
    if(today.getDay()===1){
      // Monday: check Friday
      checkDate.setDate(checkDate.getDate()-3);
    }else if(today.getDay()===0){
      // Sunday: don't remind
      return;
    }else if(today.getDay()===6){
      // Saturday: check Friday
      checkDate.setDate(checkDate.getDate()-1);
    }else{
      // Tue-Fri: check yesterday
      checkDate.setDate(checkDate.getDate()-1);
    }

    const checkDateStr=checkDate.toISOString().split('T')[0];
    const hasHours=hours.some(h=>{
      if(!h.Date)return false;
      if((h.Worker||'').toLowerCase()!==currentUser.email)return false;
      const hDate=new Date(h.Date);hDate.setHours(0,0,0,0);
      return hDate.getTime()===checkDate.getTime();
    });

    if(!hasHours){
      const days=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
      const dayName=days[checkDate.getDay()];
      document.getElementById('hoursReminderText').textContent='⏰ You have not registered hours for '+dayName+' '+formatDate(checkDate)+'. Please add your hours.';
      document.getElementById('hoursReminder').style.display='flex';
    }
  }catch(e){console.error('Reminder check failed:',e)}
}

function dismissReminder(){document.getElementById('hoursReminder').style.display='none'}

// --- INIT ---
let msalReady=false;
msalInstance.initialize().then(()=>{
  msalReady=true;initResize();
  const a=msalInstance.getAllAccounts();
  if(a.length>0){getToken().then(async()=>{if(accessToken){await loadCurrentUser();showApp();applyPermissions();await loadProperties();await loadData();checkHoursReminder()}})}
});

// ============================================================
// AUTO-REFRESH (v14.5.7)
// ============================================================

// Build a fingerprint that tells us if data has changed without full reload
async function _checkBookingChanges(){
  // Skip if any modal is open (don't disturb the user)
  if(document.querySelector('.modal-overlay.open'))return;
  try{
    const s=await getSiteId();
    const lid=await getListId('Bookings');
    // Minimal fields to detect changes — get only id + Modified
    const r=await graphGet('/sites/'+s+'/lists/'+lid+'/items?$expand=fields($select=Modified,Status)&$top=500&$orderby=fields/Modified desc');
    const currentIds=new Set(r.value.map(i=>i.id));
    let maxModified='';
    r.value.forEach(i=>{if(i.fields.Modified>maxModified)maxModified=i.fields.Modified});
    // First time — just record state
    if(!_knownBookingIds.size){
      _knownBookingIds=currentIds;
      _knownBookingModifiedMax=maxModified;
      return;
    }
    // Detect new or modified bookings
    const newIds=[...currentIds].filter(id=>!_knownBookingIds.has(id));
    const modifiedChanged=maxModified>_knownBookingModifiedMax;
    if(newIds.length>0||modifiedChanged){
      _showRefreshBanner(newIds.length);
    }
  }catch(e){/* silent fail on polling */}
}

function _showRefreshBanner(newCount){
  let banner=document.getElementById('refreshBanner');
  if(!banner){
    banner=document.createElement('div');
    banner.id='refreshBanner';
    banner.style.cssText='position:fixed;top:80px;left:50%;transform:translateX(-50%);background:#EF9F27;color:#fff;padding:10px 20px;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.2);z-index:2000;cursor:pointer;font-size:13px;font-weight:500;font-family:inherit;display:flex;align-items:center;gap:10px';
    banner.onclick=()=>{banner.remove();loadData();_knownBookingIds=new Set();_knownBookingModifiedMax=''};
    document.body.appendChild(banner);
  }
  const label=newCount>0?'⚡ '+newCount+' new booking'+(newCount!==1?'s':'')+' — click to refresh':'⚡ Data updated — click to refresh';
  banner.innerHTML=label+'<span style="opacity:.7;font-size:11px;margin-left:8px">✕</span>';
}

function _startAutoRefresh(){
  // Refresh when tab regains focus
  document.addEventListener('visibilitychange',()=>{
    if(document.visibilityState==='visible'){
      const sinceLastRefresh=Date.now()-_lastRefreshTime;
      // If away for more than 60 seconds, force refresh
      if(sinceLastRefresh>60000){
        _lastRefreshTime=Date.now();
        _knownBookingIds=new Set();_knownBookingModifiedMax='';
        const banner=document.getElementById('refreshBanner');if(banner)banner.remove();
        loadData();
      }
    }
  });
  // Poll every 5 minutes for changes (non-intrusive — just detects)
  if(_pollInterval)clearInterval(_pollInterval);
  _pollInterval=setInterval(_checkBookingChanges,5*60*1000);
}

// Start the auto-refresh watchers after initial load
window.addEventListener('DOMContentLoaded',()=>{
  // Delay startup so initial loadData completes first
  setTimeout(()=>{_startAutoRefresh()},5000);
});

// Show current user's permissions (right-click Sign Out button to trigger)
function showMyPermissions(){
  const perms=(currentUser.permissions||[]).sort();
  const hasEdit=perms.includes('edit_bookings');
  const msg='Logged in as: '+currentUser.email
    +'\nDisplay name: '+currentUser.displayName
    +'\n\nPermissions ('+perms.length+'):\n• '+perms.join('\n• ')
    +'\n\nCan edit bookings/guests: '+(hasEdit?'YES':'NO')
    +(hasEdit?'\n\nIf this user should be read-only, their Permissions field in the Users list needs to be corrected. It should only contain "view_bookings".':'');
  alert(msg);
}

// Debug helper for full-tenant calculation. Right-click "More" button to trigger.
function showFullTenantDebug(){
  if(!selectedProperty){alert('Please select a specific property first (not All).');return}
  const p=selectedProperty;
  const lines=[];
  lines.push('=== FULL-TENANT DEBUG: '+p.Title+' ===\n');
  lines.push('Property ID: '+p.id);
  lines.push('FullTenant_Company: "'+(p.FullTenant_Company||'')+'" (type: '+typeof p.FullTenant_Company+')');
  lines.push('FullTenant_RatePerRoom: '+p.FullTenant_RatePerRoom+' (type: '+typeof p.FullTenant_RatePerRoom+')');
  lines.push('FullTenant_RateUnit: '+(p.FullTenant_RateUnit||'(default: Per day)'));
  lines.push('FullTenant_StartDate: '+p.FullTenant_StartDate);
  lines.push('FullTenant_EndDate: '+(p.FullTenant_EndDate||'(empty = no end)'));
  lines.push('');
  // Count rooms matching this property
  const matching=allRooms.filter(r=>String(r.PropertyLookupId)===String(p.id));
  lines.push('Rooms with PropertyLookupId === "'+p.id+'": '+matching.length);
  // Look for Rigg 44 rooms by name pattern
  const allByPattern=allRooms.filter(r=>{
    const t=(r.Title||'').toString();
    return t.startsWith('70')||t.startsWith('80'); // adjust if room numbering differs
  });
  lines.push('Rooms with Title starting 70x or 80x: '+allByPattern.length);
  // Check for any rooms where PropertyLookupId could be a problem
  const allRoomsForReference=allRooms.length;
  lines.push('TOTAL rooms in system: '+allRoomsForReference);
  lines.push('');
  // Show a few rooms with their PropertyLookupId
  lines.push('Sample of rooms (first 5):');
  allRooms.slice(0,5).forEach(r=>{
    lines.push('  Room "'+r.Title+'" → PropertyLookupId="'+r.PropertyLookupId+'" (type: '+typeof r.PropertyLookupId+')');
  });
  lines.push('');
  // Test the actual computation for current month
  const now=new Date();
  const fromDate=new Date(now.getFullYear(),now.getMonth(),1);
  const toDate=new Date(now.getFullYear(),now.getMonth()+1,0,23,59,59);
  lines.push('Test period: '+formatDate(fromDate)+' to '+formatDate(toDate));
  const result=computeFullTenantForPeriod(p,fromDate,toDate);
  if(!result){
    lines.push('\n⚠ computeFullTenantForPeriod returned NULL (no full-tenant — that may be OK)');
  }else{
    lines.push('\nFull-tenant Result:');
    lines.push('  days: '+result.days);
    lines.push('  rooms: '+result.rooms);
    lines.push('  rate: '+result.rate);
    lines.push('  TOTAL: '+result.total+' kr');
  }
  // Long-term contracts on individual rooms
  lines.push('\n--- LONG-TERM CONTRACTS (per-room) ---');
  let ltTotal=0;
  let ltCount=0;
  matching.forEach(r=>{
    const lt=computeLongTermForRoomPeriod(r,fromDate,toDate);
    if(lt){
      ltCount++;
      ltTotal+=lt.total;
      lines.push('  '+r.Title+' → '+lt.company+' · '+lt.detailLabel+' = '+lt.total.toLocaleString('nb-NO')+' kr');
    }
  });
  if(ltCount===0){
    lines.push('  (no rooms with active LongTerm contracts on this property)');
    // Show rooms that have LongTerm_Company but no active contract — likely misconfigured
    const partialConfig=matching.filter(r=>(r.LongTerm_Company||'').trim()&&!computeLongTermForRoomPeriod(r,fromDate,toDate));
    if(partialConfig.length){
      lines.push('\n⚠ Rooms with LongTerm_Company but inactive contract:');
      partialConfig.forEach(r=>{
        lines.push('  '+r.Title+': Company="'+(r.LongTerm_Company||'')+'", Price='+r.LongTerm_Price+', Start='+r.LongTerm_StartDate+', End='+r.LongTerm_EndDate);
      });
    }
    // Show all field names on first room — to diagnose field-name mismatches
    if(matching.length){
      const sample=matching[0];
      const allKeys=Object.keys(sample).sort();
      const longTermKeys=allKeys.filter(k=>k.toLowerCase().indexOf('long')>=0||k.toLowerCase().indexOf('tenant')>=0);
      lines.push('\nAll fields on first room ('+sample.Title+'):');
      lines.push(allKeys.join(', '));
      if(longTermKeys.length){
        lines.push('\nLongTerm/Tenant-related fields found:');
        longTermKeys.forEach(k=>{
          lines.push('  '+k+' = '+JSON.stringify(sample[k]));
        });
      }else{
        lines.push('\n⚠ No fields containing "long" or "tenant" found on this room');
      }
    }
  }else{
    lines.push('  Total: '+ltCount+' rooms · '+ltTotal.toLocaleString('nb-NO')+' kr');
  }
  console.log(lines.join('\n'));
  // Use a popup window for long output (alert() truncates above ~1024 chars in Chrome)
  const txt=lines.join('\n');
  if(txt.length>800){
    const w=window.open('','_blank','width=700,height=600');
    if(w){
      w.document.write('<pre style="font-family:Consolas,monospace;font-size:12px;white-space:pre-wrap;padding:20px">'+txt.replace(/</g,'&lt;')+'</pre>');
      w.document.close();
    }else{
      alert(txt);
    }
  }else{
    alert(txt);
  }
}

// ============================================================
// OVERDUE CHECK-IN / CHECK-OUT (v14.5.7)
// ============================================================
function isOverdueCheckIn(b){
  if(!b||b.Status!=='Upcoming'||!b.Check_In)return false;
  const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  return ci.getTime()<today.getTime();
}
function isOverdueCheckOut(b){
  if(!b||b.Status!=='Active'||!b.Check_Out)return false;
  const co=new Date(b.Check_Out);co.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  return co.getTime()<today.getTime();
}
function daysOverdueCheckIn(b){
  if(!isOverdueCheckIn(b))return 0;
  const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  return Math.round((today-ci)/864e5);
}
function daysOverdueCheckOut(b){
  if(!isOverdueCheckOut(b))return 0;
  const co=new Date(b.Check_Out);co.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  return Math.round((today-co)/864e5);
}
