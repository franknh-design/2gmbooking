// ============================================================
// 2GM Booking v14.7.0 — wash.js
// Vaskeplan-logikk, overrides, vaskekalender
// ============================================================

// ============================================================
// WASH OVERRIDES (v14.5.21) — Iteration 1: data layer (CRUD)
// SP list: WashOverrides (id 626a9546-60b2-4203-91fe-ca28a1a77e94)
// Each row: BookingLookupId, Action (Add/Remove/Move), OriginalDate, NewDate,
//           ChangedBy (Person, optional), ChangedAt (DateTime), Reason, Status (Active/Reverted)
// ============================================================

// Get all overrides for one booking, sorted by ChangedAt ascending (oldest first).
// Order matters in iteration 2 because each Move shifts the baseline for following weeks.
function getWashOverridesForBooking(bookingId){
  if(!bookingId)return[];
  const bid=String(bookingId);
  return allWashOverrides
    .filter(o=>String(o.BookingLookupId||'')===bid&&(o.Status||'Active')!=='Reverted')
    .sort((a,b)=>{
      const ta=a.ChangedAt?new Date(a.ChangedAt).getTime():0;
      const tb=b.ChangedAt?new Date(b.ChangedAt).getTime():0;
      return ta-tb;
    });
}

// Save a new override to SharePoint.
// action: 'Add' | 'Remove' | 'Move'
// originalDate / newDate: Date objects or ISO strings (depending on action — see WashOverrides spec)
// reason: optional free text
// Returns the new override item from SP cache (with id), or throws on failure.
async function saveWashOverride(bookingId,action,originalDate,newDate,reason){
  if(!bookingId)throw new Error('bookingId required');
  if(!['Add','Remove','Move'].includes(action))throw new Error('invalid action: '+action);
  if((action==='Remove'||action==='Move')&&!originalDate)throw new Error(action+' requires originalDate');
  if((action==='Add'||action==='Move')&&!newDate)throw new Error(action+' requires newDate');

  const fields={
    BookingLookupId:parseInt(bookingId),
    Action:action,
    Status:'Active',
    ChangedAt:new Date().toISOString()
  };
  // Date Only fields — store as YYYY-MM-DD at midnight UTC to avoid timezone drift
  if(originalDate){
    const d=originalDate instanceof Date?originalDate:new Date(originalDate);
    fields.OriginalDate=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())).toISOString();
  }
  if(newDate){
    const d=newDate instanceof Date?newDate:new Date(newDate);
    fields.NewDate=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())).toISOString();
  }
  if(reason&&String(reason).trim())fields.Reason=String(reason).trim();
  // ChangedBy is a Person field. Setting it via Graph requires the SP user id (not email),
  // and ensureUser is awkward through Graph. For iteration 1 we set ChangedBy_Email instead
  // (custom column) — if you'd rather use the SP Person column, we can refine in iteration 4.
  if(currentUser&&currentUser.email)fields.ChangedBy_Email=currentUser.email;

  const created=await createListItem('WashOverrides',fields);
  // Append to local cache so UI updates immediately
  if(created){
    const enriched={
      id:created.id,
      BookingLookupId:fields.BookingLookupId,
      Action:fields.Action,
      Status:fields.Status,
      OriginalDate:fields.OriginalDate||null,
      NewDate:fields.NewDate||null,
      ChangedAt:fields.ChangedAt,
      Reason:fields.Reason||null,
      ChangedBy_Email:fields.ChangedBy_Email||null
    };
    allWashOverrides.push(enriched);
    return enriched;
  }
  return null;
}

// Soft-delete by setting Status to Reverted, preserving audit trail.
async function revertWashOverride(overrideId){
  if(!overrideId)throw new Error('overrideId required');
  await updateListItem('WashOverrides',overrideId,{Status:'Reverted'});
  const local=allWashOverrides.find(o=>String(o.id)===String(overrideId));
  if(local)local.Status='Reverted';
  return true;
}

// Hard delete — admin only, mainly for cleaning up test data.
async function _hardDeleteWashOverride(overrideId){
  if(!overrideId)throw new Error('overrideId required');
  const s=await getSiteId();const lid=await getListId('WashOverrides');
  await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+overrideId);
  allWashOverrides=allWashOverrides.filter(o=>String(o.id)!==String(overrideId));
  return true;
}

// v14.5.22: Calculate wash dates with optional overrides applied.
// Overrides are applied in chronological order (by ChangedAt) — each Move shifts the
// baseline weekday for ALL following weeks, per Frank's spec ("alle påfølgende vasker
// følger den nye ukedagen").
//
// Algorithm:
//   1. Generate baseline schedule starting from Check_In, weekly (skipping weekends).
//   2. For each override in ChangedAt order:
//      - Move(O, N): Find baseline date matching O. From that date forward, regenerate
//                    using N as new anchor (N + 7d, N + 14d, ...). Discard any baseline
//                    dates that came strictly after O. Type cycle (towels/beddings)
//                    continues from the moved week's parity.
//      - Add(N):     Insert N as a one-off date. Does NOT shift baseline. Marked custom.
//      - Remove(O):  Drop the date matching O from current schedule. Does NOT shift baseline.
//   3. Sort by date, recompute isPast/isToday/isNext flags.
//
// Returns: array of {date, type, week, isPast, isToday, isNext, custom?, overrideId?}
function calcWashDates(checkInDate,checkOutDate,bookingId){
  const ci=new Date(checkInDate);ci.setHours(0,0,0,0);
  const co=checkOutDate?new Date(checkOutDate):null;if(co)co.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);

  // Helper: generate weekly schedule from given anchor date, starting at weekOffset
  // (anchor itself is at weekOffset, then anchor+7, anchor+14, ...).
  // The 'parityStart' param controls towels/beddings cycle (so it continues across moves).
  function generateFrom(anchorDate,weekOffset,parityStart){
    const out=[];
    let w=weekOffset;
    while(w<=52){
      const raw=new Date(anchorDate);raw.setDate(raw.getDate()+(w-weekOffset)*7);
      const d=getNextWorkingDay(raw);
      if(co&&d>=co)break;
      const type=((w-1+parityStart)%2===0)?'Towels':'Towels + Beddings';
      out.push({date:d,type,week:w});
      w++;
    }
    return out;
  }

  // 1. Baseline: anchor = Check_In + 7 days, weeks 1..N
  // parityStart=0 means w=1 → 'Towels', w=2 → 'Towels + Beddings' (matches legacy logic)
  const firstAnchor=new Date(ci);firstAnchor.setDate(firstAnchor.getDate()+7);
  let washes=generateFrom(getNextWorkingDay(firstAnchor),1,0);

  // 2. Apply overrides in chronological order
  const overrides=bookingId?getWashOverridesForBooking(bookingId):[];
  for(const ov of overrides){
    const action=ov.Action;
    const origDate=ov.OriginalDate?new Date(ov.OriginalDate):null;
    const newDate=ov.NewDate?new Date(ov.NewDate):null;
    if(origDate)origDate.setHours(0,0,0,0);
    if(newDate)newDate.setHours(0,0,0,0);

    if(action==='Move'&&origDate&&newDate){
      // Find the wash matching origDate
      const idx=washes.findIndex(w=>w.date.getTime()===origDate.getTime());
      if(idx<0)continue; // override references a date not in current schedule — skip silently
      const movedWeek=washes[idx].week;
      // Drop washes from idx onwards and regenerate from newDate.
      // parityStart=0 keeps the towel/bedding cycle aligned with the original week numbers.
      washes=washes.slice(0,idx);
      const regen=generateFrom(newDate,movedWeek,0);
      regen.forEach(r=>{r.overrideId=ov.id;r.custom=true});
      washes=washes.concat(regen);
    }else if(action==='Remove'&&origDate){
      const idx=washes.findIndex(w=>w.date.getTime()===origDate.getTime());
      if(idx>=0)washes.splice(idx,1);
    }else if(action==='Add'&&newDate){
      // Add as one-off — does not shift baseline
      // Type defaults to Towels for ad-hoc; could be smarter but keep simple for now
      washes.push({date:newDate,type:'Towels (custom)',week:0,custom:true,overrideId:ov.id});
    }
  }

  // 3. Sort by date and recompute flags
  washes.sort((a,b)=>a.date-b.date);
  washes.forEach((w,i)=>{
    w.isPast=w.date<today;
    w.isToday=w.date.getTime()===today.getTime();
  });
  // isNext = first non-past, non-today wash
  let foundNext=false;
  washes.forEach(w=>{
    if(!foundNext&&!w.isPast&&!w.isToday){w.isNext=true;foundNext=true}
    else w.isNext=false;
  });

  return washes;
}

function getWashScheduleHtml(booking){
  if(!booking||!booking.Check_In||!(booking.Status==='Active'||booking.Status==='Upcoming'))return'';
  // v14.5.22: pass booking.id so overrides are applied
  const washes=calcWashDates(booking.Check_In,booking.Check_Out,booking.id);
  const show=washes.filter(w=>!w.isPast).slice(0,6);if(!show.length)return'';
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  // v14.6.0: Manage button for cleaners/admin
  const manageBtn=canManageWashSchedule()
    ?' <button onclick="openWashScheduleModal(\''+booking.id+'\')" style="margin-left:8px;padding:2px 8px;border:1px solid var(--accent);border-radius:4px;background:rgba(29,158,117,.1);color:var(--accent);cursor:pointer;font-size:10px;font-family:inherit">Manage</button>'
    :'';
  let html='<div style="margin-top:14px"><div style="font-size:12px;color:var(--text-secondary);margin-bottom:4px;font-weight:500">Wash schedule'+manageBtn+'</div><table style="font-size:12px;width:auto">';
  show.forEach(w=>{
    let s='',badge='';
    if(w.isToday){s='color:var(--text-danger);font-weight:500';badge=' <span class="pill danger">Today</span>'}
    else if(w.isNext){s='color:var(--accent);font-weight:500';badge=' <span class="pill" style="background:var(--bg-success);color:var(--text-success)">Next</span>'}
    if(w.custom)badge+=' <span class="pill" style="background:rgba(239,159,39,.15);color:#854F0B;font-size:10px">custom</span>';
    html+='<tr style="'+s+'"><td style="padding:2px 12px 2px 0">'+days[w.date.getDay()]+' '+formatDate(w.date)+badge+'</td><td style="padding:2px 0">'+w.type+'</td></tr>';
  });
  return html+'</table></div>';
}

function getNextWashDate(booking){
  if(!booking||!booking.Check_In||booking.Status!=='Active')return'';
  const washes=calcWashDates(booking.Check_In,booking.Check_Out,booking.id);
  const next=washes.find(w=>!w.isPast);if(!next)return'<span class="muted">—</span>';
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  if(next.isToday)return'<span class="pill danger">Today — '+next.type+'</span>';
  return days[next.date.getDay()]+' '+formatDate(next.date)+' — '+next.type;
}

// ============================================================
// WASH SCHEDULE MODAL (v14.6.0) — Iteration 3: UI for cleaners/admin
// Allows Move / Remove / Add operations on wash dates with audit trail.
// ============================================================
let _washScheduleBookingId=null;
let _washScheduleWashes=[]; // v14.6.0: cached washes-array used by inline Move toggle
let _washScheduleMinDate='';
let _washScheduleMaxDate='';

function canManageWashSchedule(){return can('admin')||can('cleaning')}

function openWashScheduleModal(bookingId){
  if(!canManageWashSchedule()){alert('No permission to manage wash schedule.');return}
  const b=allBookings.find(x=>String(x.id)===String(bookingId));
  if(!b){alert('Booking not found.');return}
  if(!b.Check_In){alert('Booking has no Check-in date.');return}
  _washScheduleBookingId=bookingId;
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  document.getElementById('washScheduleTitle').textContent='Manage wash schedule — '+(b.Person_Name||'(no name)')+', Room '+(room?room.Title:'?');
  renderWashScheduleModal();
  document.getElementById('washScheduleModal').classList.add('open');
}

function closeWashScheduleModal(){
  const hadId=_washScheduleBookingId;
  _washScheduleBookingId=null;
  document.getElementById('washScheduleModal').classList.remove('open');
  // v14.6.0: If cleaning calendar / day modal is still open underneath, refresh them
  // so the user sees their changes immediately when returning to the calendar context.
  if(document.getElementById('cleaningCalendarModal').classList.contains('open')){
    renderCleaningCalendar();
  }
  if(document.getElementById('cleaningDayModal').classList.contains('open')){
    // Re-render the currently-open day modal by re-extracting its date from title
    // — simpler: just close the day modal so user sees fresh calendar
    document.getElementById('cleaningDayModal').classList.remove('open');
  }
  // Also refresh main floors view if it's currently visible
  if(typeof renderFloors==='function'&&document.getElementById('mainPanel')){
    try{renderFloors()}catch(e){}
  }
}

function renderWashScheduleModal(){
  const bid=_washScheduleBookingId;if(!bid)return;
  const b=allBookings.find(x=>String(x.id)===String(bid));if(!b)return;
  const washes=calcWashDates(b.Check_In,b.Check_Out,b.id);
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  // Min/max for date pickers — entire stay period (no future-only restriction)
  const minDate=toISODate(b.Check_In);
  const maxDate=b.Check_Out?toISODate(b.Check_Out):'';

  // v14.6.0: Cache for inline Move toggle
  _washScheduleWashes=washes;
  _washScheduleMinDate=minDate;
  _washScheduleMaxDate=maxDate;

  // Section 1: Wash list with Move/Skip buttons per row
  // v14.6.0: Date picker is hidden by default — appears inline only when "Move" is clicked.
  let listHtml='<table style="width:100%;font-size:13px;margin-bottom:16px"><thead><tr><th style="text-align:left">Date</th><th style="text-align:left">Type</th><th style="width:280px"></th></tr></thead><tbody>';
  if(!washes.length){
    listHtml+='<tr><td colspan="3" class="muted" style="padding:12px 0">No wash dates in current schedule.</td></tr>';
  }else{
    washes.forEach((w,idx)=>{
      const dateStr=days[w.date.getDay()]+' '+formatDate(w.date);
      let stylings='';
      if(w.isPast)stylings='color:var(--text-tertiary)';
      else if(w.isToday)stylings='color:var(--text-danger);font-weight:500';
      else if(w.isNext)stylings='color:var(--accent);font-weight:500';
      const customBadge=w.custom?' <span class="pill" style="background:rgba(239,159,39,.15);color:#854F0B;font-size:10px">custom</span>':'';
      const moveDateInputId='wsMoveDate_'+idx;
      const moveCellId='wsMoveCell_'+idx;
      const moveBtn='<button onclick="washToggleMove('+idx+')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:inherit">Move</button>';
      const skipBtn='<button onclick="washRemove(\''+w.date.toISOString()+'\')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:inherit">Skip</button>';
      // The cell starts with just the buttons. When Move is clicked, replaceContent shows date picker + confirm/cancel.
      listHtml+='<tr style="'+stylings+'">'
        +'<td style="padding:6px 0">'+dateStr+customBadge+'</td>'
        +'<td>'+w.type+'</td>'
        +'<td id="'+moveCellId+'" style="text-align:right">'+moveBtn+' '+skipBtn+'</td>'
        +'</tr>';
      // Stash the inputs we'll need later in window state — reusable via washToggleMove()
    });
  }
  listHtml+='</tbody></table>';

  // Section 2: Add extra wash
  const addHtml='<div style="border-top:1px solid var(--border-tertiary);padding-top:12px;margin-bottom:16px">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:6px">Add extra wash</div>'
    +'<div style="display:flex;gap:8px;align-items:center">'
    +'<input type="date" id="wsAddDate" min="'+minDate+'" '+(maxDate?'max="'+maxDate+'"':'')+' style="padding:4px 8px;font-size:12px">'
    +'<button onclick="washAdd()" style="padding:4px 12px;border:1px solid var(--accent);border-radius:4px;background:var(--accent);color:#fff;cursor:pointer;font-size:12px;font-family:inherit">+ Add</button>'
    +'</div></div>';

  // Section 3: Reason (optional, applies to next action)
  const reasonHtml='<div style="margin-bottom:16px">'
    +'<label style="display:block;font-size:12px;font-weight:500;margin-bottom:4px">Reason <span class="muted" style="font-weight:normal">(optional, applied to next change)</span></label>'
    +'<input type="text" id="wsReason" placeholder="e.g. Guest declined, holiday, sick day..." style="width:100%;padding:6px 10px;font-size:13px;border:1px solid var(--border-tertiary);border-radius:4px">'
    +'</div>';

  // Section 4: Recent changes (last 5 overrides)
  const overrides=getWashOverridesForBooking(bid).slice(-5).reverse();
  let historyHtml='<div style="border-top:1px solid var(--border-tertiary);padding-top:12px"><div style="font-size:12px;font-weight:500;margin-bottom:6px">Recent changes</div>';
  if(!overrides.length){
    historyHtml+='<div class="muted" style="font-size:12px">No changes yet.</div>';
  }else{
    historyHtml+='<table style="width:100%;font-size:11px"><tbody>';
    overrides.forEach(o=>{
      const ts=o.ChangedAt?new Date(o.ChangedAt):null;
      const tsStr=ts?formatDate(ts)+' '+String(ts.getHours()).padStart(2,'0')+':'+String(ts.getMinutes()).padStart(2,'0'):'?';
      let desc='';
      if(o.Action==='Move')desc='Moved '+(o.OriginalDate?formatDate(o.OriginalDate):'?')+' → '+(o.NewDate?formatDate(o.NewDate):'?');
      else if(o.Action==='Remove')desc='Removed '+(o.OriginalDate?formatDate(o.OriginalDate):'?');
      else if(o.Action==='Add')desc='Added '+(o.NewDate?formatDate(o.NewDate):'?');
      const who=o.ChangedBy_Email||'?';
      const reason=o.Reason?' — <em>'+escapeHtml(o.Reason)+'</em>':'';
      historyHtml+='<tr><td style="padding:3px 0;color:var(--text-secondary);width:130px">'+tsStr+'</td><td>'+desc+reason+'</td><td style="color:var(--text-tertiary);text-align:right">'+escapeHtml(who)+'</td></tr>';
    });
    historyHtml+='</tbody></table>';
  }
  historyHtml+='</div>';

  document.getElementById('washScheduleBody').innerHTML=listHtml+addHtml+reasonHtml+historyHtml;
}

// v14.6.0: Toggle inline date picker for moving a specific wash
function washToggleMove(idx){
  const cellId='wsMoveCell_'+idx;
  const cell=document.getElementById(cellId);
  if(!cell)return;
  const w=_washScheduleWashes[idx];
  if(!w)return;
  const inputId='wsMoveDate_'+idx;
  // If already showing the date picker, close it back to buttons
  if(document.getElementById(inputId)){
    const moveBtn='<button onclick="washToggleMove('+idx+')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:inherit">Move</button>';
    const skipBtn='<button onclick="washRemove(\''+w.date.toISOString()+'\')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:inherit">Skip</button>';
    cell.innerHTML=moveBtn+' '+skipBtn;
    return;
  }
  // Show date picker + Confirm/Cancel buttons
  const maxAttr=_washScheduleMaxDate?'max="'+_washScheduleMaxDate+'"':'';
  cell.innerHTML='<input type="date" id="'+inputId+'" min="'+_washScheduleMinDate+'" '+maxAttr+' style="padding:2px 4px;font-size:11px;width:130px"> '
    +'<button onclick="washMove('+idx+')" style="padding:3px 8px;border:1px solid var(--accent);border-radius:4px;background:var(--accent);color:#fff;cursor:pointer;font-size:11px;font-family:inherit">Confirm</button> '
    +'<button onclick="washToggleMove('+idx+')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:inherit">Cancel</button>';
  // Auto-focus the date picker for quick entry
  setTimeout(()=>{const inp=document.getElementById(inputId);if(inp)inp.focus()},0);
}

async function washMove(idx){
  const w=_washScheduleWashes[idx];
  if(!w){alert('Internal error: wash not found.');return}
  const newDateInput=document.getElementById('wsMoveDate_'+idx);
  if(!newDateInput){alert('Date picker not visible.');return}
  const newDateStr=newDateInput.value;
  if(!newDateStr){alert('Pick a new date first.');newDateInput.focus();return}
  const reason=(document.getElementById('wsReason')||{}).value||'';
  const newDate=new Date(newDateStr+'T12:00:00');
  try{
    await saveWashOverride(_washScheduleBookingId,'Move',w.date,newDate,reason);
    document.getElementById('wsReason').value='';
    renderWashScheduleModal();
  }catch(e){console.error(e);alert('Failed to save: '+e.message)}
}

async function washRemove(originalDateIso){
  if(!confirm('Skip this wash date?'))return;
  const reason=(document.getElementById('wsReason')||{}).value||'';
  try{
    await saveWashOverride(_washScheduleBookingId,'Remove',new Date(originalDateIso),null,reason);
    document.getElementById('wsReason').value='';
    renderWashScheduleModal();
  }catch(e){console.error(e);alert('Failed to save: '+e.message)}
}

async function washAdd(){
  const dateStr=document.getElementById('wsAddDate').value;
  if(!dateStr){alert('Pick a date first.');return}
  const reason=(document.getElementById('wsReason')||{}).value||'';
  const newDate=new Date(dateStr+'T12:00:00');
  try{
    await saveWashOverride(_washScheduleBookingId,'Add',null,newDate,reason);
    document.getElementById('wsAddDate').value='';
    document.getElementById('wsReason').value='';
    renderWashScheduleModal();
  }catch(e){console.error(e);alert('Failed to save: '+e.message)}
}

// ============================================================
// CLEANING CALENDAR (v14.6.0) — Visualize cleaning load over the next 4 weeks
// Helps spot clustering days (30-40 rooms on one day) before they happen.
// ============================================================

function toggleCleaningCalendar(){
  const m=document.getElementById('cleaningCalendarModal');
  if(m.classList.contains('open')){closeCleaningCalendar();return}
  renderCleaningCalendar();
  m.classList.add('open');
}

function closeCleaningCalendar(){
  document.getElementById('cleaningCalendarModal').classList.remove('open');
}

// Build map: ISO-date-string -> array of {booking, room, washType}
function buildCleaningLoadMap(weeksAhead){
  const map={};
  const today=new Date();today.setHours(0,0,0,0);
  const endDate=new Date(today);endDate.setDate(endDate.getDate()+weeksAhead*7);

  // Iterate active + upcoming bookings, compute their wash dates within window
  allBookings.forEach(b=>{
    if(!b.Check_In)return;
    if(b.Status!=='Active'&&b.Status!=='Upcoming')return;
    const washes=calcWashDates(b.Check_In,b.Check_Out,b.id);
    washes.forEach(w=>{
      if(w.date<today||w.date>endDate)return;
      const key=toISODate(w.date);
      if(!map[key])map[key]=[];
      const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
      map[key].push({booking:b,room:room,washType:w.type,custom:!!w.custom});
    });

    // Also include checkout (utvask) within window — that's also cleaning load
    if(b.Check_Out&&(b.Status==='Active'||b.Status==='Upcoming')){
      const co=new Date(b.Check_Out);co.setHours(0,0,0,0);
      if(co>=today&&co<=endDate){
        const key=toISODate(co);
        if(!map[key])map[key]=[];
        const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
        map[key].push({booking:b,room:room,washType:'Checkout (utvask)',isCheckout:true});
      }
    }
  });
  return map;
}

function renderCleaningCalendar(){
  const weeks=4;
  const loadMap=buildCleaningLoadMap(weeks);

  // Determine relative color thresholds: greatest count = red, ≤1/3 of that = green, between = yellow
  const counts=Object.values(loadMap).map(arr=>arr.length);
  const maxCount=counts.length?Math.max(...counts):0;
  const greenThreshold=Math.max(2,Math.floor(maxCount/3));
  const yellowThreshold=Math.max(4,Math.floor(maxCount*2/3));

  // Find Monday of the week containing today
  const today=new Date();today.setHours(0,0,0,0);
  const dayOfWeek=today.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  const offsetToMon=(dayOfWeek===0)?-6:(1-dayOfWeek);
  const startMon=new Date(today);startMon.setDate(startMon.getDate()+offsetToMon);

  const dayLabels=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  let html='<div style="margin-bottom:12px;font-size:12px;color:var(--text-secondary)">'
    +'Showing cleaning load (towel/bedding washes + checkouts) for the next '+weeks+' weeks. '
    +'Colors are relative to busiest day. Click a day to see rooms.</div>';

  html+='<div style="display:flex;gap:16px;margin-bottom:12px;font-size:11px;flex-wrap:wrap"><span><span style="display:inline-block;width:10px;height:10px;background:#D4F4E2;border:1px solid #1D9E75;border-radius:2px"></span> Light (≤'+greenThreshold+')</span> '
    +'<span><span style="display:inline-block;width:10px;height:10px;background:#FCE4B6;border:1px solid #EF9F27;border-radius:2px"></span> Medium</span> '
    +'<span><span style="display:inline-block;width:10px;height:10px;background:#F7CACA;border:1px solid #D14343;border-radius:2px"></span> Heavy ('+yellowThreshold+'+)</span> '
    +'<span><span style="display:inline-block;width:10px;height:10px;background:#E8DCF5;border:1px solid #7B5FBF;border-radius:2px"></span> Holiday</span> '
    +'<span style="color:var(--text-tertiary)">Max: '+maxCount+' rooms</span></div>';

  // Calendar grid: 4 weeks × 7 days
  html+='<table style="width:100%;table-layout:fixed;border-collapse:separate;border-spacing:4px"><thead><tr>';
  dayLabels.forEach(d=>{html+='<th style="font-size:11px;color:var(--text-secondary);padding:4px;text-align:center;font-weight:500">'+d+'</th>'});
  html+='</tr></thead><tbody>';

  for(let wk=0;wk<weeks;wk++){
    html+='<tr>';
    for(let d=0;d<7;d++){
      const cellDate=new Date(startMon);cellDate.setDate(cellDate.getDate()+wk*7+d);
      const isWeekend=(d===5||d===6);
      const isToday=cellDate.getTime()===today.getTime();
      const isPast=cellDate<today;
      const key=toISODate(cellDate);
      const items=loadMap[key]||[];
      const count=items.length;
      // v14.6.0: Norwegian holidays — distinct visual signal
      const holidayName=getHolidayName(cellDate);

      // Color
      let bg='#fff',border='var(--border-tertiary)',textColor='var(--text-primary)';
      if(isPast){bg='#f5f5f5';textColor='var(--text-tertiary)'}
      else if(holidayName){bg='#E8DCF5';border='#7B5FBF';textColor='#4A2D8C'}
      else if(isWeekend){bg='#fafafa';textColor='var(--text-tertiary)'}
      else if(count>=yellowThreshold){bg='#F7CACA';border='#D14343'}
      else if(count>=greenThreshold){bg='#FCE4B6';border='#EF9F27'}
      else if(count>0){bg='#D4F4E2';border='#1D9E75'}

      const todayRing=isToday?'box-shadow:0 0 0 2px var(--accent);':'';
      const cursor=(count>0&&!isPast)?'cursor:pointer;':'';
      const onclick=(count>0&&!isPast)?'onclick="openCleaningDayModal(\''+key+'\')"':'';
      const dateStr=cellDate.getDate()+'. '+months[cellDate.getMonth()];
      const titleAttr=holidayName?' title="'+holidayName+'"':'';

      let cellContent='<div style="font-size:11px;color:'+textColor+';font-weight:500">'+dateStr+'</div>';
      if(holidayName){
        // Show holiday name and any wash count (should normally be 0 since calcWashDates avoids holidays)
        cellContent+='<div style="font-size:9px;color:'+textColor+';margin-top:2px;line-height:1.2">🎌 '+holidayName+'</div>';
        if(count>0){
          cellContent+='<div style="font-size:14px;color:#A32D2D;font-weight:600;margin-top:2px">⚠ '+count+'</div>';
        }
      }else{
        cellContent+='<div style="font-size:18px;color:'+textColor+';font-weight:600;margin-top:4px">'+(count||(isPast||isWeekend?'':'·'))+'</div>';
      }

      html+='<td style="background:'+bg+';border:1px solid '+border+';border-radius:6px;padding:8px 4px;text-align:center;height:70px;'+todayRing+cursor+'"'+titleAttr+' '+onclick+'>'+cellContent+'</td>';
    }
    html+='</tr>';
  }
  html+='</tbody></table>';

  document.getElementById('cleaningCalendarBody').innerHTML=html;
}

// Day detail: show all rooms scheduled for one specific date with Manage links
function openCleaningDayModal(isoDate){
  const loadMap=buildCleaningLoadMap(4);
  const items=loadMap[isoDate]||[];
  const date=new Date(isoDate+'T00:00:00');
  const days=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  document.getElementById('cleaningDayTitle').textContent=days[date.getDay()]+' '+formatDate(date)+' — '+items.length+' room'+(items.length===1?'':'s');

  let html='';
  if(!items.length){
    html='<div class="muted">No cleaning scheduled this day.</div>';
  }else{
    // Sort: checkouts first (more time-critical), then by room number
    const sorted=[...items].sort((a,b)=>{
      if(a.isCheckout&&!b.isCheckout)return -1;
      if(b.isCheckout&&!a.isCheckout)return 1;
      const ra=a.room?a.room.Title:'';
      const rb=b.room?b.room.Title:'';
      return String(ra).localeCompare(String(rb),undefined,{numeric:true});
    });
    html='<table style="width:100%;font-size:13px"><thead><tr><th style="text-align:left">Room</th><th style="text-align:left">Guest</th><th style="text-align:left">Type</th><th style="text-align:right;width:90px"></th></tr></thead><tbody>';
    sorted.forEach(it=>{
      const room=it.room?it.room.Title:'?';
      const name=it.booking.Person_Name||'(no name)';
      const typeStyle=it.isCheckout?'color:#EF9F27;font-weight:500':(it.custom?'color:#854F0B':'');
      const customMark=it.custom?' <span class="pill" style="background:rgba(239,159,39,.15);color:#854F0B;font-size:10px">custom</span>':'';
      const manageBtn=canManageWashSchedule()&&!it.isCheckout
        ?'<button onclick="openWashScheduleModal(\''+it.booking.id+'\')" style="padding:3px 8px;border:1px solid var(--accent);border-radius:4px;background:rgba(29,158,117,.1);color:var(--accent);cursor:pointer;font-size:11px;font-family:inherit">Manage</button>'
        :'';
      html+='<tr><td style="padding:5px 0">'+escapeHtml(room)+'</td><td>'+escapeHtml(name)+'</td><td style="'+typeStyle+'">'+escapeHtml(it.washType)+customMark+'</td><td style="text-align:right">'+manageBtn+'</td></tr>';
    });
    html+='</tbody></table>';
  }

  document.getElementById('cleaningDayBody').innerHTML=html;
  document.getElementById('cleaningDayModal').classList.add('open');
}

function closeCleaningDayModal(){
  document.getElementById('cleaningDayModal').classList.remove('open');
}

// --- RENDERING ---
