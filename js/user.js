// ============================================================
// 2GM Booking v14.7.0 — user.js
// Brukerprofil, tillatelser, nav-tilstand
// ============================================================

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
  // v14.6.0: Cleaning calendar — admin or cleaning permission
  show('btnCleaningCalendar',can('admin')||can('cleaning'));
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
    selectedProperty=null; // v14.5.10: default to "All properties" instead of first property
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
    // v14.5.21: Load wash overrides
    try{allWashOverrides=await getListItems('WashOverrides')}catch(e){allWashOverrides=[];console.warn('[WashOverrides] load failed:',e.message)}
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
      // v14.5.10: Show Upcoming in main list as soon as Check_In <= today (no kl-12 rule)
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
