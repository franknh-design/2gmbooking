// ============================================================
// 2GM Booking v14.7.0 — hours.js
// Timeregistrering, import, eksport
// ============================================================

function toggleHours(){
  if(currentView==='hours'){showMainView();return}
  showHoursView();initHoursSelectors();loadHoursData();
}
function initHoursSelectors(){
  const now=new Date();
  const monthSel=document.getElementById('hoursMonth');const yearSel=document.getElementById('hoursYear');
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  if(!monthSel.children.length){
    monthSel.innerHTML=months.map((m,i)=>'<option value="'+i+'"'+(i===now.getMonth()?' selected':'')+'>'+m+'</option>').join('');
    const y=now.getFullYear();yearSel.innerHTML=[y-1,y,y+1].map(yr=>'<option value="'+yr+'"'+(yr===y?' selected':'')+'>'+yr+'</option>').join('');
  }
  // Worker filter dropdown
  const wf=document.getElementById('hoursWorkerFilter');
  if(can('view_all_hours')){
    const workers=[...new Set(allUsers.map(u=>u.Epost).filter(Boolean))];
    wf.innerHTML='<option value="all">All workers</option>'+workers.map(w=>{
      const u=allUsers.find(x=>(x.Epost||'').toLowerCase()===w.toLowerCase());
      return'<option value="'+w+'">'+(u?userDisplayName(u):w)+'</option>';
    }).join('');
    wf.style.display='';
  }else{
    wf.innerHTML='<option value="'+currentUser.email+'">'+currentUser.displayName+'</option>';
    wf.style.display='none';
  }
  // Location datalist
  const dl=document.getElementById('locationList');
  const locs=[...new Set(properties.map(p=>p.Title))];
  // Add custom locations
  const extraLocs=['Kontor','Diverse'];
  const allLocs=[...new Set([...locs,...extraLocs])];
  dl.innerHTML=allLocs.map(l=>'<option value="'+l+'">').join('');
}

async function loadHoursData(){try{allHours=await getListItems('Hours');renderHours()}catch(e){console.error('Failed to load hours:',e)}}

function onHoursMonthChange(){
  // Clear date range when user changes month/year
  const from=document.getElementById('hoursFrom');const to=document.getElementById('hoursTo');
  if(from)from.value='';if(to)to.value='';
  renderHours();
}
function clearHoursDateRange(){
  document.getElementById('hoursFrom').value='';document.getElementById('hoursTo').value='';
  renderHours();
}

function renderHours(){
  const month=parseInt(document.getElementById('hoursMonth').value);
  const year=parseInt(document.getElementById('hoursYear').value);
  const workerFilter=document.getElementById('hoursWorkerFilter').value;
  const fromVal=document.getElementById('hoursFrom').value;
  const toVal=document.getElementById('hoursTo').value;
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];

  // If either from or to is set, use date range. Otherwise use month/year.
  const useRange=!!(fromVal||toVal);
  const fromDate=fromVal?new Date(fromVal+'T00:00:00'):null;
  const toDate=toVal?new Date(toVal+'T23:59:59'):null;

  const filtered=allHours.filter(h=>{
    if(!h.Date)return false;const d=new Date(h.Date);
    if(useRange){
      if(fromDate&&d<fromDate)return false;
      if(toDate&&d>toDate)return false;
    }else{
      if(d.getMonth()!==month||d.getFullYear()!==year)return false;
    }
    if(workerFilter!=='all'&&(h.Worker||'').toLowerCase()!==workerFilter.toLowerCase())return false;
    return true;
  }).sort((a,b)=>new Date(a.Date)-new Date(b.Date));

  const workerName=workerFilter==='all'?'All workers':userDisplayName(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase()))||workerFilter;
  let periodLabel;
  if(useRange){
    const f=fromVal?formatDate(fromVal):'…';const t=toVal?formatDate(toVal):'…';
    periodLabel=f+' – '+t;
  }else{
    periodLabel=months[month]+' '+year;
  }
  document.getElementById('hoursTitle').textContent='Hours — '+periodLabel+' — '+workerName;

  // Update charts if visible
  const chartsContainer=document.getElementById('hoursChartsContainer');
  if(chartsContainer&&chartsContainer.style.display!=='none'){renderHoursCharts(filtered)}
  // Update efficiency if visible
  const effContainer=document.getElementById('efficiencyContainer');
  if(effContainer&&effContainer.style.display!=='none'){renderEfficiency()}

  const body=document.getElementById('hoursBody');
  if(!filtered.length){body.innerHTML='<tr><td colspan="8" class="loading">No hours registered</td></tr>';document.getElementById('hoursTotal').textContent='0.00';return}

  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];let total=0;
  body.innerHTML=filtered.map(h=>{
    const hrs=calcHoursDiff(h.Time_From,h.Time_To);total+=hrs;
    const d=new Date(h.Date);
    const workerUser=allUsers.find(u=>(u.Epost||'').toLowerCase()===(h.Worker||'').toLowerCase());
    const wName=workerUser?userDisplayName(workerUser):(h.Worker||'');
    return'<tr data-hours-id="'+h.id+'" style="cursor:pointer"><td>'+days[d.getDay()]+' '+formatDate(h.Date)+'</td><td>'+(h.Location||'')+'</td><td>'+wName+'</td><td>'+(h.Time_From||'')+'</td><td>'+(h.Time_To||'')+'</td><td style="text-align:right">'+hrs.toFixed(2)+'</td>'
      +'<td class="muted" style="font-size:11px">'+_hoursNotes(h)+'</td>'
      +'<td style="text-align:right"><button data-hours-delete="'+h.id+'" style="width:20px;height:20px;border-radius:50%;border:1px solid var(--border-tertiary);background:var(--bg-primary);color:var(--text-danger);cursor:pointer;font-size:11px;line-height:1;padding:0" title="Delete">✕</button></td></tr>';
  }).join('');
  document.getElementById('hoursTotal').textContent=total.toFixed(2);

  // Attach click handlers via delegation
  document.getElementById('hoursBody').onclick=function(e){
    // Check if delete button was clicked
    const delBtn=e.target.closest('[data-hours-delete]');
    if(delBtn){e.stopPropagation();deleteHoursEntry(delBtn.dataset.hoursDelete);return}
    // Otherwise check if row was clicked for edit
    const row=e.target.closest('tr[data-hours-id]');
    if(row)openEditHours(row.dataset.hoursId);
  };
}

function calcHoursDiff(from,to){
  if(!from||!to)return 0;const[fh,fm]=(from||'00:00').split(':').map(Number);const[th,tm]=(to||'00:00').split(':').map(Number);
  let diff=(th*60+tm)-(fh*60+fm);if(diff<0)diff+=24*60;return diff/60;
}

function openAddHours(){
  editingHoursId=null;
  document.getElementById('hoursModal').querySelector('h2').textContent='Add hours';
  document.getElementById('hoursSaveBtn').textContent='Save';
  // Default date: last day of selected month, or today if current month
  const selMonth=parseInt(document.getElementById('hoursMonth').value);
  const selYear=parseInt(document.getElementById('hoursYear').value);
  const now=new Date();
  let defaultDate;
  if(selMonth===now.getMonth()&&selYear===now.getFullYear()){
    defaultDate=now.toISOString().split('T')[0];
  }else{
    const lastDay=new Date(selYear,selMonth+1,0).getDate();
    defaultDate=selYear+'-'+String(selMonth+1).padStart(2,'0')+'-'+String(lastDay).padStart(2,'0');
  }
  document.getElementById('hDate').value=defaultDate;
  document.getElementById('hFrom').value='08:00';document.getElementById('hTo').value='16:00';
  document.getElementById('hLocation').value='';
  document.getElementById('hNotes').value='';
  populateHoursWorkerSelect();
  document.getElementById('hoursModal').classList.add('open');
}

function openEditHours(id){
  const h=allHours.find(x=>x.id===id);if(!h)return;
  editingHoursId=id;
  document.getElementById('hoursModal').querySelector('h2').textContent='Edit hours';
  document.getElementById('hoursSaveBtn').textContent='Save changes';
  document.getElementById('hDate').value=h.Date?toISODate(h.Date):'';
  document.getElementById('hFrom').value=h.Time_From||'08:00';
  document.getElementById('hTo').value=h.Time_To||'16:00';
  document.getElementById('hLocation').value=h.Location||'';
  document.getElementById('hNotes').value=_hoursNotes(h);
  populateHoursWorkerSelect(h.Worker);
  document.getElementById('hoursModal').classList.add('open');
}

function populateHoursWorkerSelect(preselectedWorker){
  const ws=document.getElementById('hWorker');
  const selected=(preselectedWorker||currentUser.email).toLowerCase();
  if(can('edit_others_hours')){
    const workers=allUsers.filter(u=>u.Active!==false);
    ws.innerHTML=workers.map(u=>'<option value="'+(u.Epost||'')+'"'+((u.Epost||'').toLowerCase()===selected?' selected':'')+'>'+userDisplayName(u)+'</option>').join('');
    document.getElementById('hWorkerGroup').style.display='';
  }else{
    ws.innerHTML='<option value="'+currentUser.email+'">'+currentUser.displayName+'</option>';
    document.getElementById('hWorkerGroup').style.display='none';
  }
}

function closeHoursModal(){document.getElementById('hoursModal').classList.remove('open');editingHoursId=null}

async function saveHours(){
  const date=document.getElementById('hDate').value;const location=document.getElementById('hLocation').value;
  const from=document.getElementById('hFrom').value;const to=document.getElementById('hTo').value;
  const worker=document.getElementById('hWorker').value;
  const notes=document.getElementById('hNotes').value.trim();
  if(!date){alert('Date required');return}if(!from||!to){alert('From/To required');return}if(!location){alert('Location required');return}
  const workerUser=allUsers.find(u=>(u.Epost||'').toLowerCase()===worker.toLowerCase());
  const workerName=workerUser?userDisplayName(workerUser):worker;
  const btn=document.getElementById('hoursSaveBtn');btn.disabled=true;btn.textContent='Saving...';
  const fields={Title:workerName+' — '+location+' — '+date,Date:date+'T00:00:00Z',Location:location,Time_From:from,Time_To:to,Worker:worker};
  if(notes)fields.Notes=notes;else fields.Notes='';
  try{
    if(editingHoursId){
      await updateListItem('Hours',editingHoursId,fields);
      const local=allHours.find(x=>x.id===editingHoursId);
      if(local)Object.assign(local,fields);
    }else{
      await createListItem('Hours',fields);
    }
    closeHoursModal();await loadHoursData();
  }catch(e){alert('Failed: '+e.message)}finally{btn.disabled=false;btn.textContent=editingHoursId?'Save changes':'Save'}
}

async function deleteHoursEntry(id){
  if(!confirm('Delete this entry?'))return;
  try{const s=await getSiteId();const lid=await getListId('Hours');await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+id);allHours=allHours.filter(h=>h.id!==id);renderHours()}catch(e){alert('Failed')}
}

function _hoursNotes(h){return h.Notes||h.Note||h.Merknad||h.Comments||''}

function exportHoursExcel(){
  const month=parseInt(document.getElementById('hoursMonth').value);const year=parseInt(document.getElementById('hoursYear').value);
  const workerFilter=document.getElementById('hoursWorkerFilter').value;
  const fromVal=document.getElementById('hoursFrom').value;const toVal=document.getElementById('hoursTo').value;
  const useRange=!!(fromVal||toVal);
  const fromDate=fromVal?new Date(fromVal+'T00:00:00'):null;
  const toDate=toVal?new Date(toVal+'T23:59:59'):null;
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const filtered=allHours.filter(h=>{
    if(!h.Date)return false;const d=new Date(h.Date);
    if(useRange){if(fromDate&&d<fromDate)return false;if(toDate&&d>toDate)return false}
    else{if(d.getMonth()!==month||d.getFullYear()!==year)return false}
    if(workerFilter!=='all'&&(h.Worker||'').toLowerCase()!==workerFilter.toLowerCase())return false;
    return true;
  }).sort((a,b)=>new Date(a.Date)-new Date(b.Date));
  const headers=['Date','Day','Location','Worker','From','To','Hours','Notes'];let total=0;
  const rows=filtered.map(h=>{
    const hrs=calcHoursDiff(h.Time_From,h.Time_To);total+=hrs;
    const d=new Date(h.Date);
    const wu=allUsers.find(u=>(u.Epost||'').toLowerCase()===(h.Worker||'').toLowerCase());
    return[formatDate(h.Date),days[d.getDay()],h.Location||'',wu?userDisplayName(wu):h.Worker||'',h.Time_From||'',h.Time_To||'',hrs.toFixed(2),_hoursNotes(h)];
  });
  // Total row: push to Hours column only, leave Notes empty
  rows.push(['','','','','','Total',total.toFixed(2),'']);
  const workerName=workerFilter==='all'?'All':userDisplayName(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase()))||workerFilter;
  const periodStr=useRange?((fromVal||'start')+'_to_'+(toVal||'end')):(months[month]+'_'+year);
  downloadCSV('Hours_'+workerName.replace(/\s+/g,'_')+'_'+periodStr,headers,rows);
}

async function archiveHoursMonth(){
  const month=parseInt(document.getElementById('hoursMonth').value);const year=parseInt(document.getElementById('hoursYear').value);
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const filtered=allHours.filter(h=>{if(!h.Date)return false;const d=new Date(h.Date);return d.getMonth()===month&&d.getFullYear()===year});
  if(!filtered.length){alert('No hours for '+months[month]+' '+year);return}
  if(!confirm('Archive '+filtered.length+' entries for '+months[month]+' '+year+'?'))return;
  let done=0;for(const h of filtered){try{await updateListItem('Hours',h.id,{Archived:'Yes'});h.Archived='Yes';done++}catch(e){}}
  alert(done+' entries archived.');renderHours();
}

// --- CSV HELPER ---
function downloadCSV(filename,headers,rows){
  const escape=v=>{const s=String(v);if(s.includes(';')||s.includes('"')||s.includes('\n'))return'"'+s.replace(/"/g,'""')+'"';return s};
  const csv='\uFEFF'+headers.join(';')+'\n'+rows.map(r=>r.map(escape).join(';')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=filename+'.csv';a.click();URL.revokeObjectURL(url);
}

// --- ADMIN ---
function addBookingToGuests(bookingId){
  if(!can('edit_bookings')){alert('You do not have permission to add guests.');return}
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking not found');return}
  // Open Guest-editor with booking data pre-filled
  editingPersonId=null;
  document.getElementById('personModalTitle').textContent='New guest (from booking)';
  document.getElementById('pName').value=b.Person_Name||'';
  document.getElementById('pMobile').value='';
  document.getElementById('pEmail').value='';
  document.getElementById('pAddress').value='';
  document.getElementById('pCompany').value=b.Company||'';
  document.getElementById('pCompanyOrgNr').value='';
  document.getElementById('pCompanyEmail').value='';
  document.getElementById('pCompanyAddress').value='';
  document.getElementById('pNotes').value='';
  document.getElementById('pDeleteRow').style.display='none';
  document.getElementById('personModal').classList.add('open');
  // Focus on mobile field since name+company are already filled
  setTimeout(()=>{const el=document.getElementById('pMobile');if(el)el.focus()},100);
}

// ============================================================
// GUEST BOOKINGS HISTORY (v14.5.10)
// ============================================================
function showGuestBookings(name){
  if(!name)return;
  const lower=name.toLowerCase().trim();
  // Find all bookings matching this name (fuzzy)
  const words=lower.split(/[\s,]+/).filter(w=>w.length>1);
  const matching=allBookings.filter(b=>{
    const bn=(b.Person_Name||'').toLowerCase().trim();
    if(bn===lower)return true;
    if(words.length>=2){
      const bwords=bn.split(/[\s,]+/).filter(w=>w.length>1);
      if(bwords.length>=2){
        return words.every(w=>bn.indexOf(w)>=0)||bwords.every(w=>lower.indexOf(w)>=0);
      }
    }
    return false;
  }).sort((a,b)=>new Date(b.Check_In||0)-new Date(a.Check_In||0));
  document.getElementById('guestBookingsTitle').textContent=name+' — '+matching.length+' booking'+(matching.length!==1?'s':'');
  const body=document.getElementById('guestBookingsBody');
  if(!matching.length){
    body.innerHTML='<div style="text-align:center;padding:20px;color:var(--text-secondary)">No bookings found.</div>';
  }else{
    // Summary
    let totalNights=0,totalRevenue=0;
    matching.forEach(b=>{
      if(b.Status==='Cancelled')return;
      if(!b.Check_In)return;
      const ci=new Date(b.Check_In);const co=b.Check_Out?new Date(b.Check_Out):new Date();
      const n=Math.max(0,Math.round((co-ci)/864e5));
      totalNights+=n;
      const cost=calcBookingCost(b,b.Property_Name||'');
      totalRevenue+=n*(cost.rate||0);
    });
    let html='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:14px">'
      +'<div style="background:var(--bg-secondary);padding:10px;border-radius:8px"><div style="font-size:11px;color:var(--text-tertiary)">Bookings</div><div style="font-size:18px;font-weight:500">'+matching.length+'</div></div>'
      +'<div style="background:var(--bg-secondary);padding:10px;border-radius:8px"><div style="font-size:11px;color:var(--text-tertiary)">Guest-nights</div><div style="font-size:18px;font-weight:500">'+totalNights+'</div></div>'
      +'<div style="background:var(--bg-secondary);padding:10px;border-radius:8px"><div style="font-size:11px;color:var(--text-tertiary)">Total revenue</div><div style="font-size:18px;font-weight:500;color:var(--text-success)">'+totalRevenue.toLocaleString('nb-NO')+' kr</div></div>'
      +'</div>';
    html+='<table style="width:100%;font-size:12px"><thead><tr style="background:var(--bg-secondary)">'
      +'<th style="padding:6px 10px;text-align:left">Room</th>'
      +'<th style="padding:6px 10px;text-align:left">Property</th>'
      +'<th style="padding:6px 10px;text-align:left">Company</th>'
      +'<th style="padding:6px 10px;text-align:left">Check-in</th>'
      +'<th style="padding:6px 10px;text-align:left">Check-out</th>'
      +'<th style="padding:6px 10px;text-align:right">Nights</th>'
      +'<th style="padding:6px 10px;text-align:left">Status</th>'
      +'</tr></thead><tbody>';
    matching.forEach(b=>{
      const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
      const roomTitle=room?room.Title:'?';
      const ci=b.Check_In?new Date(b.Check_In):null;
      const co=b.Check_Out?new Date(b.Check_Out):(b.Status==='Active'?new Date():null);
      const nights=ci&&co?Math.max(0,Math.round((co-ci)/864e5)):0;
      const statusColor={Completed:'background:var(--bg-secondary);color:var(--text-secondary)',Cancelled:'background:var(--bg-danger);color:var(--text-danger)',Active:'background:var(--bg-success);color:var(--text-success)',Upcoming:'background:var(--bg-warning);color:var(--text-warning)'}[b.Status]||'';
      html+='<tr onclick="document.getElementById(\'guestBookingsModal\').classList.remove(\'open\');openEditBooking(\''+b.id+'\')" style="border-top:.5px solid var(--border-tertiary);cursor:pointer" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
        +'<td style="padding:6px 10px;font-weight:500">'+escapeHtml(roomTitle)+'</td>'
        +'<td style="padding:6px 10px">'+escapeHtml(b.Property_Name||'')+'</td>'
        +'<td style="padding:6px 10px">'+escapeHtml(b.Company||'')+'</td>'
        +'<td style="padding:6px 10px">'+(ci?formatDate(b.Check_In):'—')+'</td>'
        +'<td style="padding:6px 10px">'+(b.Check_Out?formatDate(b.Check_Out):(b.Status==='Active'?'<span class="muted">Open</span>':'—'))+'</td>'
        +'<td style="padding:6px 10px;text-align:right">'+nights+'</td>'
        +'<td style="padding:6px 10px"><span class="pill" style="'+statusColor+'">'+b.Status+'</span></td>'
        +'</tr>';
    });
    html+='</tbody></table>';
    body.innerHTML=html;
  }
  document.getElementById('guestBookingsModal').classList.add('open');
}

// ============================================================
// HOURS IMPORT (v14.5.10)
// ============================================================
let importHoursData=[];

function closeImportHoursModal(){
  document.getElementById('importHoursModal').classList.remove('open');
  document.getElementById('importHoursFile').value='';
  document.getElementById('importHoursPreview').style.display='none';
  document.getElementById('importHoursProgress').style.display='none';
  document.getElementById('importHoursResult').style.display='none';
  document.getElementById('importHoursRunBtn').disabled=true;
  importHoursData=[];
}

// Parse HH:MM time strings, validate
function _parseTime(t){
  if(!t)return null;t=String(t).trim();
  const m=t.match(/^(\d{1,2}):(\d{2})$/);
  if(!m)return null;
  const h=parseInt(m[1]),min=parseInt(m[2]);
  if(h<0||h>23||min<0||min>59)return null;
  return String(h).padStart(2,'0')+':'+String(min).padStart(2,'0');
}

function parseImportHoursFile(){
  const file=document.getElementById('importHoursFile').files[0];
  if(!file){alert('Select a CSV file');return}
  const reader=new FileReader();
  reader.onload=function(e){
    const lines=e.target.result.split(/\r?\n/).filter(l=>l.trim());
    if(lines.length<2){alert('Need header row + at least one data row');return}
    const sep=lines[0].includes(';')?';':',';
    const headers=lines[0].split(sep).map(h=>h.trim().toLowerCase().replace(/['"]/g,''));
    // Map column indices (accept Norwegian + English variants)
    const colDate=headers.findIndex(h=>h==='date'||h==='dato');
    const colWorker=headers.findIndex(h=>h==='worker'||h==='arbeider'||h==='ansatt'||h==='email'||h==='epost');
    const colLocation=headers.findIndex(h=>h==='location'||h==='lokasjon'||h==='sted'||h==='property');
    const colFrom=headers.findIndex(h=>h==='time_from'||h==='from'||h==='fra'||h==='start');
    const colTo=headers.findIndex(h=>h==='time_to'||h==='to'||h==='til'||h==='end'||h==='slutt');
    const colNotes=headers.findIndex(h=>h==='notes'||h==='note'||h==='merknad'||h==='kommentar');
    if(colDate===-1||colWorker===-1||colLocation===-1||colFrom===-1||colTo===-1){
      alert('Missing required columns (Date, Worker, Location, Time_From, Time_To). Found: '+headers.join(', '));return;
    }
    importHoursData=[];
    for(let i=1;i<lines.length;i++){
      const cols=lines[i].split(sep).map(c=>c.trim().replace(/^['"]|['"]$/g,''));
      const dateRaw=cols[colDate]||'';
      const worker=(cols[colWorker]||'').toLowerCase();
      const location=cols[colLocation]||'';
      const timeFrom=cols[colFrom]||'';
      const timeTo=cols[colTo]||'';
      const notes=colNotes>=0?(cols[colNotes]||''):'';
      if(!dateRaw&&!worker&&!location)continue; // skip blank rows
      const date=parseDate(dateRaw);
      const tFrom=_parseTime(timeFrom);
      const tTo=_parseTime(timeTo);
      // Worker validation: should match a known user's email
      const userMatch=worker?allUsers.find(u=>(u.Epost||'').toLowerCase()===worker):null;
      let error=null;
      if(!date)error='Invalid date';
      else if(!worker)error='Worker required';
      else if(!userMatch)error='Worker email not found in Users';
      else if(!location)error='Location required';
      else if(!tFrom)error='Invalid From time (HH:MM)';
      else if(!tTo)error='Invalid To time (HH:MM)';
      else if(calcHoursDiff(tFrom,tTo)<=0)error='Time_To must be after Time_From';
      const hrs=(tFrom&&tTo)?calcHoursDiff(tFrom,tTo).toFixed(2):'—';
      importHoursData.push({date,worker,location,timeFrom:tFrom,timeTo:tTo,notes,hrs,error,workerName:userMatch?userDisplayName(userMatch):worker});
    }
    const valid=importHoursData.filter(r=>!r.error).length;
    const issues=importHoursData.filter(r=>r.error).length;
    document.getElementById('importHoursPreviewTitle').textContent=importHoursData.length+' rows — '+valid+' OK'+(issues?', '+issues+' with issues (will be skipped)':'');
    document.getElementById('importHoursPreviewBody').innerHTML=importHoursData.slice(0,100).map(r=>
      '<tr style="'+(r.error?'background:var(--bg-danger);color:var(--text-danger)':'')+'">'
      +'<td>'+(r.date||'—')+'</td>'
      +'<td>'+(r.workerName||r.worker||'—')+'</td>'
      +'<td>'+(r.location||'—')+'</td>'
      +'<td>'+(r.timeFrom||'—')+'</td>'
      +'<td>'+(r.timeTo||'—')+'</td>'
      +'<td style="text-align:right">'+r.hrs+'</td>'
      +'<td style="font-size:10px">'+escapeHtml(r.notes||'')+'</td>'
      +'<td style="font-size:10px">'+(r.error||'')+'</td>'
      +'</tr>'
    ).join('');
    if(importHoursData.length>100){
      document.getElementById('importHoursPreviewBody').innerHTML+='<tr><td colspan="8" style="text-align:center;font-style:italic;color:var(--text-tertiary);padding:8px">…and '+(importHoursData.length-100)+' more rows (preview limited to 100)</td></tr>';
    }
    document.getElementById('importHoursPreview').style.display='block';
    document.getElementById('importHoursRunBtn').disabled=valid===0;
  };
  reader.readAsText(file,'UTF-8');
}

async function runImportHours(){
  const valid=importHoursData.filter(r=>!r.error);
  if(!valid.length)return;
  // Check for potential duplicates against existing data
  const dupes=valid.filter(r=>{
    return allHours.some(h=>{
      const hd=h.Date?toISODate(h.Date):'';
      return hd===r.date&&(h.Worker||'').toLowerCase()===r.worker&&(h.Location||'')===r.location&&(h.Time_From||'')===r.timeFrom;
    });
  });
  let confirmMsg='Import '+valid.length+' hours entries?';
  if(dupes.length){
    confirmMsg+='\n\n⚠ '+dupes.length+' look like duplicates of existing entries (same date/worker/location/start time). Continue anyway?';
  }
  if(!confirm(confirmMsg))return;
  const bar=document.getElementById('importHoursProgressBar');
  const text=document.getElementById('importHoursProgressText');
  document.getElementById('importHoursProgress').style.display='block';
  document.getElementById('importHoursRunBtn').disabled=true;
  let success=0,failed=0;
  const errors=[];
  for(let i=0;i<valid.length;i++){
    const r=valid[i];
    text.textContent='Importing '+(i+1)+'/'+valid.length+'…';
    bar.style.width=Math.round((i+1)/valid.length*100)+'%';
    const fields={
      Title:r.workerName+' — '+r.location+' — '+r.date,
      Date:r.date+'T00:00:00Z',
      Location:r.location,
      Time_From:r.timeFrom,
      Time_To:r.timeTo,
      Worker:r.worker,
      Notes:r.notes||''
    };
    try{
      await createListItem('Hours',fields);
      success++;
    }catch(e){
      failed++;
      errors.push('Row '+(i+2)+': '+e.message);
    }
    // Throttle every 10 to avoid rate limiting
    if(i%10===9)await new Promise(res=>setTimeout(res,500));
  }
  document.getElementById('importHoursProgress').style.display='none';
  const result=document.getElementById('importHoursResult');
  result.style.display='block';
  result.style.background=failed?'var(--bg-warning)':'var(--bg-success)';
  result.style.color=failed?'var(--text-warning)':'var(--text-success)';
  let msg='Done! '+success+' imported'+(failed?', '+failed+' failed':'')+'.';
  if(errors.length&&errors.length<=5)msg+='\n\nErrors: '+errors.join('; ');
  result.textContent=msg;
  // Reload Hours data so the new entries show up
  if(success>0)await loadHoursData();
}

