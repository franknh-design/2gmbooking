// ============================================================
// 2GM Booking v10.7 — app.js (Core)
// Auth, Graph API, Data, Rendering, Bookings
// ============================================================

// --- CONFIG ---
const msalConfig={auth:{clientId:'f8e2259d-c440-41d3-94e3-3a2dce095817',authority:'https://login.microsoftonline.com/2b495272-f733-47a8-a771-bb744309fa17',redirectUri:'https://franknh-design.github.io/2gmbooking/'},cache:{cacheLocation:'localStorage'}};
const msalInstance=new msal.PublicClientApplication(msalConfig);
const SITE_HOST='2gmeiendom.sharepoint.com';
const SITE_PATH='/sites/2GMBooking';
const LIST_IDS={Properties:'d842d574-f238-442a-be3d-77334727e89f',Rooms:'bfa962a0-5eb2-416c-abe8-adba06558c11',Bookings:'fe1dfe34-23df-4864-b0b1-b01bf60bfb75',Persons:'ebbe517d-83f8-4169-9423-70c63a3f8c07',Cleaning_Log:'6b1bd5f9-c54f-42ee-892f-d50c79481375',Hours:'9db53c54-70dd-483d-ad1d-565d0e4ac7ac',Users:'1b9b866f-0944-4f43-a80d-2a630e1e7c25'};

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
  {key:'admin',label:'User administration'}
];

// --- STATE ---
let accessToken=null,siteId=null;
let currentUser={email:'',displayName:'',permissions:[]};
let properties=[],rooms=[],allRooms=[],bookings=[],allBookings=[],allUsers=[];
let selectedProperty=null,selectedRoom=null,selectedBooking=null;
let editingBookingId=null,checkoutBookingId=null;
let activeFilter=null;
let currentView='main'; // 'main' or 'hours'

// --- AUTH ---
async function signIn(){
  if(!msalReady){alert('Please wait...');return}
  try{
    await msalInstance.loginPopup({scopes:['Sites.ReadWrite.All']});
    await getToken();await loadCurrentUser();
    showApp();applyPermissions();
    await loadProperties();await loadData();
  }catch(e){console.error('Login failed:',e)}
}
async function getToken(){
  const a=msalInstance.getAllAccounts();if(!a.length)return null;
  try{const r=await msalInstance.acquireTokenSilent({scopes:['Sites.ReadWrite.All'],account:a[0]});accessToken=r.accessToken;return accessToken}
  catch(e){const r=await msalInstance.acquireTokenPopup({scopes:['Sites.ReadWrite.All']});accessToken=r.accessToken;return accessToken}
}
function signOut(){msalInstance.logoutPopup();document.getElementById('app').style.display='none';document.getElementById('loginScreen').style.display='block'}
function showApp(){document.getElementById('loginScreen').style.display='none';document.getElementById('app').style.display='block'}

// --- GRAPH API ---
async function graphGet(ep){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{headers:{Authorization:'Bearer '+accessToken,Accept:'application/json'}});if(!r.ok)throw new Error('Graph error '+r.status+': '+await r.text());return r.json()}
async function graphPatch(ep,body){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'PATCH',headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},body:JSON.stringify(body)});if(!r.ok)throw new Error('Graph error '+r.status);return r.json()}
async function graphPost(ep,body){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'POST',headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},body:JSON.stringify(body)});if(!r.ok){const t=await r.text();throw new Error('Graph error '+r.status+': '+t)}return r.json()}
async function graphDelete(ep){await getToken();const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'DELETE',headers:{Authorization:'Bearer '+accessToken}});if(!r.ok)throw new Error('Graph error '+r.status);return true}

async function getSiteId(){if(siteId)return siteId;const r=await graphGet('/sites/'+SITE_HOST+':'+SITE_PATH);siteId=r.id;return siteId}
async function getListId(name){if(LIST_IDS[name])return LIST_IDS[name];throw new Error('List not found: '+name)}
async function getListItems(listName){const s=await getSiteId();const lid=await getListId(listName);let all=[];let url='/sites/'+s+'/lists/'+lid+'/items?$expand=fields&$top=500';while(url){const r=await graphGet(url);all=all.concat(r.value.map(i=>({id:i.id,...i.fields})));url=r['@odata.nextLink']?r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0',''):null}return all}
async function createListItem(listName,fields){const s=await getSiteId();const lid=await getListId(listName);return graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields})}
async function updateListItem(listName,itemId,fields){const s=await getSiteId();const lid=await getListId(listName);return graphPatch('/sites/'+s+'/lists/'+lid+'/items/'+itemId+'/fields',fields)}

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

function applyPermissions(){
  const el=id=>document.getElementById(id);
  // Header buttons
  el('btnNewBooking').style.display=can('edit_bookings')?'':'none';
  el('adminBtn').style.display=can('admin')?'':'none';
  el('btnArchive').style.display=can('archive')||can('view_bookings')?'':'none';
  el('btnUpcoming').style.display=can('view_bookings')?'':'none';
  el('btnHours').style.display=can('view_hours')||can('edit_hours')?'':'none';
  // Sign out label
  el('btnSignOut').textContent=currentUser.displayName+' — Sign out';
  // Stats: hide if no view permission
  if(!can('view_bookings')){
    el('statsBar').style.display='none';
    document.querySelector('.floors').style.display='none';
  }
}

// --- DATA LOADING ---
async function loadProperties(){
  try{
    properties=await getListItems('Properties');
    const sel=document.getElementById('propertySelect');
    sel.innerHTML=properties.map(p=>'<option value="'+p.id+'">'+p.Title+'</option>').join('');
    sel.onchange=()=>{selectedProperty=properties.find(p=>p.id===sel.value);loadData()};
    selectedProperty=properties[0];
  }catch(e){console.error('Error loading properties:',e)}
}

async function loadData(){
  if(!selectedProperty)return;
  document.getElementById('headerTitle').textContent='2GM Booking — '+selectedProperty.Title;
  document.getElementById('floor1Body').innerHTML='<tr><td colspan="7" class="loading">Loading...</td></tr>';
  document.getElementById('floor2Body').innerHTML='<tr><td colspan="7" class="loading">Loading...</td></tr>';
  closeDetail();
  try{
    allRooms=await getListItems('Rooms');
    allBookings=await getListItems('Bookings');
    rooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(selectedProperty.id));
    if(rooms.length===0){rooms=allRooms.filter(r=>r.Active!==false)}
    filterBookingsForView();
    renderFloors();updateStats();
  }catch(e){console.error('Error:',e);document.getElementById('floor1Body').innerHTML='<tr><td colspan="7" class="error">Error: '+e.message+'</td></tr>'}
}

function filterBookingsForView(){
  const roomIds=new Set(rooms.map(r=>r.id));
  const now=new Date();
  bookings=allBookings.filter(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!roomIds.has(rid))return false;
    if(b.Status==='Active')return true;
    if(b.Status==='Upcoming'){
      const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
      const today=new Date();today.setHours(0,0,0,0);
      if(ci.getTime()<today.getTime())return true;
      if(ci.getTime()===today.getTime()&&now.getHours()>=12)return true;
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
function toISODate(d){if(!d)return'';return new Date(d).toISOString().split('T')[0]}

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
  return'<span style="'+s+'">'+ci+'</span> — '+co;
}

function renderRow(room,booking){
  const n=booking?booking.Person_Name:'';const c=booking?(booking.Company||''):'';
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+doorTagBtn(booking)+'</td><td>'+cleanBtn(booking)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+'</td>'
    +'<td>'+(n||'<span class="empty-text">—</span>')+(booking&&booking.Notes?'<span class="note-dot"></span>':'')+'</td>'
    +'<td class="muted">'+c+'</td>'
    +'<td style="text-align:right;font-variant-numeric:tabular-nums">'+batCell(room.Door_Battery_Level)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums">'+datesCell(booking)+'</td></tr>';
}

function renderRowWithProperty(room,booking,propName){
  const n=booking?booking.Person_Name:'';const washNext=booking?getNextWashDate(booking):'';
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+cleanBtn(booking)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+'</td>'
    +'<td class="muted" style="font-size:11px">'+propName+'</td>'
    +'<td>'+(n||'<span class="empty-text">—</span>')+'</td>'
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

  const renderFn=(r)=>{
    const b=bMap[r.id];
    if(activeFilter==='dirty'){
      const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
      return renderRowWithProperty(r,b,prop?prop.Title:'');
    }
    return renderRow(r,b);
  };

  document.getElementById('floor1Body').innerHTML=f1.length?f1.map(renderFn).join(''):noMatch;
  document.getElementById('floor2Body').innerHTML=f2.length?f2.map(renderFn).join(''):noMatch;

  if(activeFilter==='dirty'){
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
  const tr=rooms.length;
  const occupiedRoomIds=new Set();
  bookings.forEach(b=>occupiedRoomIds.add(String(b.RoomLookupId||'')));
  document.getElementById('statCheckedIn').textContent=occupiedRoomIds.size+' / '+tr;
  document.getElementById('statEmpty').textContent=tr-occupiedRoomIds.size;
  // Dirty: all properties
  const allDirtyRoomIds=new Set();
  allBookings.forEach(b=>{
    if(b.Cleaning_Status==='Dirty'&&(b.Status==='Active'||b.Status==='Upcoming'))allDirtyRoomIds.add(String(b.RoomLookupId));
    if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out);if(w.some(x=>x.isToday))allDirtyRoomIds.add(String(b.RoomLookupId))}
  });
  document.getElementById('statDirty').textContent=allDirtyRoomIds.size;
  document.getElementById('statDoorTag').textContent=bookings.filter(b=>b.Door_Tag_Status==='Needs-print').length;
  document.getElementById('statBattery').textContent=rooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30).length;
}

// --- DETAIL PANEL ---
function showDetail(roomId){
  const room=(activeFilter==='dirty'?allRooms:rooms).find(r=>r.id===roomId);
  if(!room)return;
  const sourceBk=(activeFilter==='dirty')?allBookings:bookings;
  const booking=sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Active')
    ||sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Upcoming');
  selectedRoom=room;selectedBooking=booking;
  const p=document.getElementById('detailPanel');
  const prop=properties.find(pr=>String(pr.id)===String(room.PropertyLookupId));
  const propName=prop?prop.Title:'';

  if(!booking){
    p.innerHTML='<div class="detail-grid"><div class="detail-main"><div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">Empty — '+propName+'</div></div><div class="detail-actions">'
      +(can('edit_bookings')?'<button class="primary" onclick="openNewBooking(\''+room.id+'\')">Create booking</button>':'')
      +'<button onclick="closeDetail()">Close</button></div></div>';
  }else{
    const dt={'None':'—','Needs-print':'✕ Needs print','Printed':'✓ Printed'}[booking.Door_Tag_Status]||'—';
    const cl={'None':'—','Dirty':'● Needs cleaning','Clean':'● Clean'}[booking.Cleaning_Status]||'—';
    const washHtml=getWashScheduleHtml(booking);
    let infoHtml='';
    if(can('view_bookings')){
      infoHtml='<div class="detail-name">'+booking.Person_Name+'</div>'
        +'<div class="detail-sub">Room '+room.Title+' · '+(booking.Company||'')+' · '+propName+'</div>'
        +'<table class="detail-info">'
        +'<tr><td>Check-in</td><td>'+formatDate(booking.Check_In)+'</td></tr>'
        +'<tr><td>Check-out</td><td>'+(booking.Check_Out?formatDate(booking.Check_Out):'Open-ended')+'</td></tr>'
        +'<tr><td>Status</td><td>'+booking.Status+'</td></tr>'
        +'<tr><td>Door tag</td><td>'+dt+'</td></tr>'
        +'<tr><td>Cleaning</td><td>'+cl+'</td></tr>'
        +(booking.Notes?'<tr><td>Notes</td><td>'+booking.Notes+'</td></tr>':'')
        +'</table>'+washHtml;
    }else{
      infoHtml='<div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">'+cl+'</div>'+washHtml;
    }
    let btns='';
    if(can('edit_bookings'))btns+='<button onclick="openEditBooking(\''+booking.id+'\')">Edit booking</button>';
    if(can('print_doortag'))btns+='<button onclick="printDoorTag(\''+booking.id+'\')">Print door tag</button>';
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
  if(!confirm('Cancel this booking?'))return;
  try{await updateListItem('Bookings',id,{Status:'Cancelled'});const l=allBookings.find(x=>x.id===id);if(l)l.Status='Cancelled';closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}
}

// --- BOOKING MODAL ---
function populateRoomSelect(preselectedRoomId){
  const sel=document.getElementById('fRoom');
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  sel.innerHTML=sorted.map(r=>'<option value="'+r.id+'"'+(r.id===preselectedRoomId?' selected':'')+'>'+r.Title+' (Floor '+r.Floor+')</option>').join('');
  sel.onchange=()=>{const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:''};
  const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:'';
}
function openNewBooking(preselectedRoomId){
  ensureMainView();
  editingBookingId=null;document.getElementById('bookingModalTitle').textContent='New booking';
  document.getElementById('bookingSaveBtn').textContent='Create booking';
  populateRoomSelect(preselectedRoomId||'');
  document.getElementById('fName').value='';document.getElementById('fCompany').value='';
  document.getElementById('fCheckIn').value=toISODate(new Date());document.getElementById('fCheckOut').value='';
  document.getElementById('fStatus').value='Upcoming';document.getElementById('fNotes').value='';
  document.getElementById('bookingModal').classList.add('open');
}
function openEditBooking(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;editingBookingId=bookingId;
  document.getElementById('bookingModalTitle').textContent='Edit booking';
  document.getElementById('bookingSaveBtn').textContent='Save changes';
  populateRoomSelect(String(b.RoomLookupId));
  document.getElementById('fName').value=b.Person_Name||'';document.getElementById('fCompany').value=b.Company||'';
  document.getElementById('fCheckIn').value=b.Check_In?toISODate(b.Check_In):'';
  document.getElementById('fCheckOut').value=b.Check_Out?toISODate(b.Check_Out):'';
  document.getElementById('fStatus').value=b.Status||'Upcoming';document.getElementById('fNotes').value=b.Notes||'';
  document.getElementById('bookingModal').classList.add('open');
}
function closeBookingModal(){document.getElementById('bookingModal').classList.remove('open');editingBookingId=null}

async function saveBooking(){
  const roomId=document.getElementById('fRoom').value;
  const name=document.getElementById('fName').value.trim();
  const company=document.getElementById('fCompany').value.trim();
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

  const fields={Person_Name:name,Company:company,Check_In:checkIn+'T15:00:00Z',Status:status,Door_Tag_Status:'Needs-print',Cleaning_Status:'None',Property_Name:selectedProperty.Title,Floor:room?room.Floor:1,Notes:notes||null};
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
function showMainView(){currentView='main';document.getElementById('mainView').style.display='';document.getElementById('hoursView').style.display='none';document.getElementById('propertySelect').style.display='';if(selectedProperty)document.getElementById('headerTitle').textContent='2GM Booking — '+selectedProperty.Title}
function showHoursView(){currentView='hours';document.getElementById('mainView').style.display='none';document.getElementById('hoursView').style.display='';document.getElementById('propertySelect').style.display='none';document.getElementById('headerTitle').textContent='2GM Booking — Hours'}
function ensureMainView(){if(currentView==='hours')showMainView()}

// --- FILTER ---
function toggleFilter(filter){
  if(activeFilter===filter){clearFilter();return}
  activeFilter=filter;
  document.querySelectorAll('.stat').forEach((el,i)=>{const f=['checkedIn','empty','dirty','doorTag','battery'];el.classList.toggle('active',f[i]===filter)});
  const labels={checkedIn:'Showing: Checked-in rooms',empty:'Showing: Empty rooms',dirty:'Showing: Rooms needing cleaning',doorTag:'Showing: Door tags needing print',battery:'Showing: Low battery rooms (<30%)'};
  document.getElementById('filterLabel').textContent=labels[filter]||'';
  document.getElementById('filterBar').classList.add('open');renderFloors();
}
function clearFilter(){activeFilter=null;document.querySelectorAll('.stat').forEach(el=>el.classList.remove('active'));document.getElementById('filterBar').classList.remove('open');renderFloors()}

function getFilteredRoomsForFloor(floor){
  const sourceRooms=(activeFilter==='dirty')?allRooms:rooms;
  let floorRooms=sourceRooms.filter(r=>r.Floor===floor||String(r.Floor)===String(floor));
  if(!activeFilter)return floorRooms;
  const sourceBookings=(activeFilter==='dirty')?allBookings:bookings;
  const bMap={};sourceBookings.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&(!bMap[rid]||b.Status==='Active'))bMap[rid]=b});
  switch(activeFilter){
    case 'checkedIn':return floorRooms.filter(r=>!!bMap[r.id]);
    case 'empty':return floorRooms.filter(r=>!bMap[r.id]);
    case 'dirty':return floorRooms.filter(r=>{const b=bMap[r.id];if(!b)return false;if(b.Cleaning_Status==='Dirty')return true;if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out);if(w.some(x=>x.isToday))return true}return false});
    case 'doorTag':return floorRooms.filter(r=>bMap[r.id]&&bMap[r.id].Door_Tag_Status==='Needs-print');
    case 'battery':return floorRooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30);
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

// --- INIT ---
let msalReady=false;
msalInstance.initialize().then(()=>{
  msalReady=true;initResize();
  const a=msalInstance.getAllAccounts();
  if(a.length>0){getToken().then(async()=>{if(accessToken){await loadCurrentUser();showApp();applyPermissions();await loadProperties();await loadData()}})}
});
