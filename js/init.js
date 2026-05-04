// ============================================================
// 2GM Booking v14.7.0 — init.js
// Bootstrap, MSAL-init, auto-refresh, resize, debug
// ============================================================

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
  const labels={checkedIn:'Showing: Checked-in rooms',empty:'Showing: Empty rooms',dirty:'Showing: Rooms needing cleaning',doorTag:'Showing: Door tags needing print',battery:'Showing: Low battery rooms (<30%)',overdueCheckIn:'Showing: Overdue check-in',overdueCheckOut:'Showing: Overdue check-out',needsAttention:'Showing: Bookings needing attention'};
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
  // v14.5.10: All stat filters now show across ALL assigned properties (not just selected)
  const assignedPropIds=new Set(properties.map(p=>p.id));
  const allAssignedRooms=allRooms.filter(r=>assignedPropIds.has(String(r.PropertyLookupId)));
  // For stat filters, use cross-property source. For non-filter view, use selected property
  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut','needsAttention'].includes(activeFilter);
  const sourceRooms=isStatFilter?allAssignedRooms:rooms;
  let floorRooms=sourceRooms.filter(r=>r.Floor===floor||String(r.Floor)===String(floor));
  if(!activeFilter)return floorRooms;
  const sourceBookings=isStatFilter?allBookings:bookings;
  const bMap={};sourceBookings.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&(!bMap[rid]||b.Status==='Active'))bMap[rid]=b});
  switch(activeFilter){
    case 'checkedIn':return floorRooms.filter(r=>!!bMap[r.id]);
    case 'empty':return floorRooms.filter(r=>!bMap[r.id]);
    case 'dirty':return floorRooms.filter(r=>{const b=bMap[r.id];if(!b)return false;if(b.Cleaning_Status==='Dirty')return true;if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out,b.id);if(w.some(x=>x.isToday))return true}return false});
    case 'doorTag':return floorRooms.filter(r=>bMap[r.id]&&bMap[r.id].Door_Tag_Status==='Needs-print');
    case 'battery':return floorRooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30);
    case 'overdueCheckIn':return floorRooms.filter(r=>{const b=bMap[r.id];return b&&isOverdueCheckIn(b)});
    case 'overdueCheckOut':return floorRooms.filter(r=>{const b=bMap[r.id];return b&&isOverdueCheckOut(b)});
    // v14.5.19: needsAttention-filter must check ALL bookings on the room, not just the one in bMap.
    // bMap only keeps the "active" booking per room (preferring Active over Upcoming), but a
    // problematic booking may be hidden behind a newer one. Without this fix, the count and
    // filter results disagree (count says 1, filter returns 0).
    case 'needsAttention':return floorRooms.filter(r=>{
      return allBookings.some(b=>String(b.RoomLookupId||'')===r.id&&bookingNeedsAttention(b)!==null);
    });
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

// --- INIT (v14.6.0 — handleRedirectPromise consumes redirect-login response) ---
msalInstance.initialize().then(async()=>{
  // Consume any redirect response sitting in URL (#code=...) — MSAL otherwise leaves it
  // hanging which causes silent token failures later. v14.6.0: this is now the PRIMARY login path.
  let redirectLoginCompleted=false;
  try{
    const resp=await msalInstance.handleRedirectPromise();
    if(resp&&resp.accessToken){
      accessToken=resp.accessToken;
      _tokenExpiresAt=resp.expiresOn?resp.expiresOn.getTime():(Date.now()+50*60*1000);
      redirectLoginCompleted=true;
      console.log('[Auth] Redirect login completed for',resp.account&&resp.account.username);
    }
  }catch(e){console.warn('[Auth] handleRedirectPromise:',e.message)}
  msalReady=true;initResize();
  const a=msalInstance.getAllAccounts();
  if(a.length>0){
    try{
      // After redirect-login, accessToken is already set above — getToken(false) should find it cached.
      // For returning users (not redirect-login), getToken(false) silently refreshes if needed.
      const tok=await getToken(false);
      if(tok){await loadCurrentUser();showApp();applyPermissions();await loadProperties();await loadData();checkHoursReminder()}
    }catch(e){
      console.warn('[Init] startup token failed:',e.message);
      // If we just completed a redirect-login but still failed → something is wrong, surface it
      if(redirectLoginCompleted)alert('Innlogging fullført, men kunne ikke laste data: '+e.message);
    }
  }
});

// ============================================================
// AUTO-REFRESH (v14.5.10)
// ============================================================

// Build a fingerprint that tells us if data has changed without full reload
async function _checkBookingChanges(){
  // Skip if any modal is open (don't disturb the user)
  if(document.querySelector('.modal-overlay.open'))return;
  try{
    // v14.5.11: silent mode — never trigger popup from background polling
    const tok=await getToken(false);
    if(!tok)return; // no token, skip this cycle quietly
    const s=await getSiteId();
    const lid=await getListId('Bookings');
    // Minimal fields to detect changes — get only id + Modified
    const r=await graphGet('/sites/'+s+'/lists/'+lid+'/items?$expand=fields($select=Modified,Status)&$top=500&$orderby=fields/Modified desc',true);
    if(!r)return; // silent fail
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
  // visibilitychange-refresh fjernet (v15) — forstyrret ved fanebytte.
  // Endringer fanges av 5-minutters polling nedenfor.
  if(_pollInterval)clearInterval(_pollInterval);
  _pollInterval=setInterval(_checkBookingChanges,5*60*1000);
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
// OVERDUE CHECK-IN / CHECK-OUT (v14.5.10)
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

// v14.5.18: Detect bookings with logically inconsistent state.
// Returns null if booking is OK, otherwise an object describing the issue.
// Two issue types:
//   - 'invalid_status': Check_Out has passed but Status is still Upcoming/Active
//   - 'extreme_overdue_in': Status=Upcoming but Check_In was >30 days ago (forgotten booking)
function bookingNeedsAttention(b){
  if(!b)return null;
  if(b.Status==='Completed'||b.Status==='Cancelled')return null;
  const today=new Date();today.setHours(0,0,0,0);
  // 1. Invalid status: Check_Out is in the past but status is still active/upcoming
  if(b.Check_Out){
    const co=new Date(b.Check_Out);co.setHours(0,0,0,0);
    if(co.getTime()<today.getTime()){
      const days=Math.round((today-co)/864e5);
      return{type:'invalid_status',label:'Should be Completed',daysSinceCheckOut:days};
    }
  }
  // v15.1: Upcoming uten check-in dato — gjest har ikke gitt ankomstdato
  if(b.Status==='Upcoming'&&!b.Check_In){
    return{type:'no_date',label:'Dato ikke satt'};
  }
  // 2. Extreme overdue check-in: Upcoming but Check_In was >30 days ago
  if(b.Status==='Upcoming'&&b.Check_In){
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const daysSince=Math.round((today-ci)/864e5);
    if(daysSince>30){
      return{type:'extreme_overdue_in',label:'Never checked in',daysSinceCheckIn:daysSince};
    }
  }
  return null;
}

// Convenience: return all bookings that need attention from given list
function getBookingsNeedingAttention(bookings){
  return (bookings||[]).filter(b=>bookingNeedsAttention(b)!==null);
}
