// ============================================================
// 2GM Booking v11.7 — modules.js
// Hours, Archive, Import/Export, Admin (checkbox permissions)
// ============================================================

// --- UPCOMING ---
function toggleIncoming(){
  ensureMainView();
  document.getElementById('archivePanel').classList.remove('open');
  document.getElementById('incomingPanel').classList.toggle('open');
  if(document.getElementById('incomingPanel').classList.contains('open'))renderIncoming();
}
function renderIncoming(){
  const today=new Date();today.setHours(0,0,0,0);
  const in30=new Date(today);in30.setDate(in30.getDate()+30);
  const tomorrow=new Date(today);tomorrow.setDate(tomorrow.getDate()+1);
  const roomIds=new Set(rooms.map(r=>r.id));
  const upcoming=allBookings.filter(b=>{
    if(b.Status!=='Upcoming')return false;
    const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    return ci>=tomorrow&&ci<=in30;
  }).sort((a,b)=>new Date(a.Check_In)-new Date(b.Check_In));
  const body=document.getElementById('incomingBody');
  if(!upcoming.length){body.innerHTML='<tr><td colspan="7" class="loading">No upcoming bookings</td></tr>';return}
  body.innerHTML=upcoming.map(b=>{
    const room=rooms.find(r=>r.id===String(b.RoomLookupId));const roomTitle=room?room.Title:'?';
    const daysUntil=Math.round((new Date(b.Check_In)-today)/864e5);let badge='';
    if(daysUntil<=3)badge='<span class="pill danger">'+daysUntil+'d</span>';
    else if(daysUntil<=7)badge='<span class="pill warning">'+daysUntil+'d</span>';
    return'<tr onclick="showDetail(\''+(room?room.id:'')+'\')">'
      +'<td style="font-weight:500">'+roomTitle+'</td><td>'+b.Person_Name+'</td><td class="muted">'+(b.Company||'')+'</td>'
      +'<td>'+formatDate(b.Check_In)+' '+badge+'</td><td>'+(b.Check_Out?formatDate(b.Check_Out):'Open-ended')+'</td>'
      +'<td><span class="pill" style="background:var(--bg-warning);color:var(--text-warning)">Upcoming</span></td>'
      +'<td><button onclick="event.stopPropagation();openEditBooking(\''+b.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit">Edit</button></td></tr>';
  }).join('');
}

// --- ARCHIVE ---
function toggleArchive(){
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.toggle('open');
  if(document.getElementById('archivePanel').classList.contains('open'))renderArchive();
}
function renderArchive(){
  const search=(document.getElementById('archiveSearch').value||'').toLowerCase();
  const statusFilter=document.getElementById('archiveStatus').value;
  const roomIds=new Set(rooms.map(r=>r.id));
  let archived=allBookings.filter(b=>{
    const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;
    if(statusFilter!=='all'&&b.Status!==statusFilter)return false;
    if(search){if(!((b.Person_Name||'')+(b.Company||'')+getRoomTitle(b)).toLowerCase().includes(search))return false}
    return true;
  }).sort((a,b)=>new Date(b.Check_In||0)-new Date(a.Check_In||0));
  const limited=archived.slice(0,100);
  document.getElementById('archiveTitle').textContent='Archive — '+archived.length+' booking'+(archived.length!==1?'s':'');
  const body=document.getElementById('archiveBody');
  if(!limited.length){body.innerHTML='<tr><td colspan="7" class="loading">No bookings found</td></tr>';return}
  body.innerHTML=limited.map(b=>{
    const sc={Completed:'background:var(--bg-secondary);color:var(--text-secondary)',Cancelled:'background:var(--bg-danger);color:var(--text-danger)',Active:'background:var(--bg-success);color:var(--text-success)',Upcoming:'background:var(--bg-warning);color:var(--text-warning)'}[b.Status]||'';
    return'<tr><td style="font-weight:500">'+getRoomTitle(b)+'</td><td>'+(b.Person_Name||'—')+'</td><td class="muted">'+(b.Company||'')+'</td><td>'+formatDate(b.Check_In)+'</td><td>'+(b.Check_Out?formatDate(b.Check_Out):'Open-ended')+'</td><td><span class="pill" style="'+sc+'">'+b.Status+'</span></td>'
      +'<td>'+(b.Status==='Completed'||b.Status==='Cancelled'?'<button onclick="reopenBooking(\''+b.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit">Reopen</button>':'')+'</td></tr>';
  }).join('')+(archived.length>100?'<tr><td colspan="7" class="loading">Showing 100 of '+archived.length+'</td></tr>':'');
}
function getRoomTitle(b){const r=allRooms.find(rm=>rm.id===String(b.RoomLookupId));return r?r.Title:'?'}
async function reopenBooking(id){
  const b=allBookings.find(x=>x.id===id);
  if(!b)return;
  // If check-in is today or earlier → Active, otherwise Upcoming
  const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
  const today=new Date();today.setHours(0,0,0,0);
  const newStatus=ci<=today?'Active':'Upcoming';
  if(!confirm('Reopen as '+newStatus+'?\n\n'+b.Person_Name+'\nCheck-in: '+formatDate(b.Check_In)+(b.Check_Out?'\nCheck-out: '+formatDate(b.Check_Out):'\nOpen-ended')))return;
  try{
    // Keep original dates, just change status and reset cleaning/doortag
    await updateListItem('Bookings',id,{Status:newStatus,Cleaning_Status:'None',Door_Tag_Status:'Needs-print'});
    await loadData();renderArchive();
  }catch(e){alert('Failed: '+e.message)}
}

// --- EXPORT ARCHIVE ---
function exportArchiveExcel(){
  const search=(document.getElementById('archiveSearch').value||'').toLowerCase();
  const statusFilter=document.getElementById('archiveStatus').value;
  const roomIds=new Set(rooms.map(r=>r.id));
  let archived=allBookings.filter(b=>{const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;if(statusFilter!=='all'&&b.Status!==statusFilter)return false;if(search){if(!((b.Person_Name||'')+(b.Company||'')+getRoomTitle(b)).toLowerCase().includes(search))return false}return true}).sort((a,b)=>new Date(b.Check_In||0)-new Date(a.Check_In||0));
  const showPrices=can('view_prices');
  const headers=['Room','Name','Company','Check-in','Check-out','Nights','Status','Door Tag','Cleaning','Notes'];
  if(showPrices)headers.push('Rate/night','Total','Rate source');
  const rows=archived.map(b=>{
    const propTitle=b.Property_Name||selectedProperty.Title||'';
    const nights=calcBookingNights(b);
    const row=[getRoomTitle(b),b.Person_Name||'',b.Company||'',formatDate(b.Check_In),b.Check_Out?formatDate(b.Check_Out):'Open-ended',nights,b.Status||'',b.Door_Tag_Status||'',b.Cleaning_Status||'',(b.Notes||'').replace(/[\r\n]+/g,' ')];
    if(showPrices){const cost=calcBookingCost(b,propTitle);row.push(cost.rate,cost.total,cost.source)}
    return row;
  });
  downloadCSV('Archive_'+(selectedProperty?selectedProperty.Title:'2GM').replace(/\s+/g,'_')+'_'+new Date().toISOString().split('T')[0],headers,rows);
}

// --- IMPORT ---
let importData=[];
function closeImportModal(){document.getElementById('importModal').classList.remove('open');document.getElementById('importFile').value='';document.getElementById('importPreview').style.display='none';document.getElementById('importProgress').style.display='none';document.getElementById('importResult').style.display='none';document.getElementById('importRunBtn').disabled=true;importData=[]}
function parseDate(str){
  if(!str||!str.trim())return null;str=str.trim();
  let m=str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
  m=str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);if(m){const y=parseInt(m[3]);return(y>50?'19':'20')+m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0')}
  m=str.match(/^(\d{4})-(\d{2})-(\d{2})$/);if(m)return str;
  m=str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
  return null;
}
function parseImportFile(){
  const file=document.getElementById('importFile').files[0];if(!file){alert('Select a CSV file');return}
  const reader=new FileReader();
  reader.onload=function(e){
    const lines=e.target.result.split(/\r?\n/).filter(l=>l.trim());if(lines.length<2){alert('Need header + data');return}
    const sep=lines[0].includes(';')?';':',';
    const headers=lines[0].split(sep).map(h=>h.trim().toLowerCase().replace(/['"]/g,''));
    const colRoom=headers.findIndex(h=>h==='room'||h==='rom');const colName=headers.findIndex(h=>h==='name'||h==='navn'||h==='person_name');
    const colCompany=headers.findIndex(h=>h==='company'||h==='firma');const colCheckIn=headers.findIndex(h=>h.includes('check-in')||h.includes('checkin')||h==='inn'||h==='from');
    const colCheckOut=headers.findIndex(h=>h.includes('check-out')||h.includes('checkout')||h==='ut'||h==='to');const colStatus=headers.findIndex(h=>h==='status');
    if(colRoom===-1||colName===-1||colCheckIn===-1){alert('Missing columns. Found: '+headers.join(', '));return}
    const defaultStatus=document.getElementById('importStatus').value;importData=[];
    for(let i=1;i<lines.length;i++){
      const cols=lines[i].split(sep).map(c=>c.trim().replace(/^['"]|['"]$/g,''));
      const roomTitle=cols[colRoom]||'';const name=cols[colName]||'';if(!roomTitle||!name)continue;
      const checkIn=parseDate(cols[colCheckIn]||'');const checkOut=colCheckOut>=0?parseDate(cols[colCheckOut]||''):null;
      const status=(colStatus>=0&&cols[colStatus])?cols[colStatus]:defaultStatus;
      const room=allRooms.find(r=>r.Title===roomTitle);
      importData.push({roomTitle,roomId:room?room.id:null,name,company:cols[colCompany]||'',checkIn,checkOut,status,error:!room?'Room not found':(!checkIn?'Invalid date':null)});
    }
    document.getElementById('importPreviewTitle').textContent=importData.length+' rows'+(importData.filter(r=>r.error).length?' — '+importData.filter(r=>r.error).length+' issues':'');
    document.getElementById('importPreviewBody').innerHTML=importData.slice(0,50).map(r=>'<tr style="'+(r.error?'background:var(--bg-danger)':'')+'"><td>'+r.roomTitle+'</td><td>'+r.name+'</td><td>'+r.company+'</td><td>'+(r.checkIn||'—')+'</td><td>'+(r.checkOut||'—')+'</td><td>'+r.status+'</td></tr>').join('');
    document.getElementById('importPreview').style.display='block';
    document.getElementById('importRunBtn').disabled=importData.filter(r=>!r.error).length===0;
  };reader.readAsText(file,'UTF-8');
}
async function runImport(){
  const valid=importData.filter(r=>!r.error);if(!valid.length)return;
  if(!confirm('Import '+valid.length+' bookings?'))return;
  const bar=document.getElementById('importProgressBar');const text=document.getElementById('importProgressText');
  document.getElementById('importProgress').style.display='block';document.getElementById('importRunBtn').disabled=true;
  let success=0,failed=0;
  for(let i=0;i<valid.length;i++){
    const r=valid[i];text.textContent='Importing '+(i+1)+'/'+valid.length;bar.style.width=Math.round((i+1)/valid.length*100)+'%';
    const fields={Person_Name:r.name,Company:r.company,Check_In:r.checkIn+'T15:00:00Z',Status:r.status,Door_Tag_Status:'None',Cleaning_Status:'None',Property_Name:selectedProperty.Title,RoomLookupId:parseInt(r.roomId)};
    if(r.checkOut)fields.Check_Out=r.checkOut+'T12:00:00Z';
    const room=allRooms.find(rm=>rm.id===r.roomId);if(room)fields.Floor=room.Floor;
    try{await createListItem('Bookings',fields);success++}catch(e){failed++}
    if(i%10===9)await new Promise(res=>setTimeout(res,500));
  }
  document.getElementById('importProgress').style.display='none';
  const result=document.getElementById('importResult');result.style.display='block';
  result.style.background=failed?'var(--bg-warning)':'var(--bg-success)';result.style.color=failed?'var(--text-warning)':'var(--text-success)';
  result.textContent='Done! '+success+' imported'+(failed?', '+failed+' failed':'')+'.';
  await loadData();renderArchive();
}

// --- HOURS ---
let allHours=[],editingHoursId=null;

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
      return'<option value="'+w+'">'+(u?u.DisplayName:w)+'</option>';
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

function renderHours(){
  const month=parseInt(document.getElementById('hoursMonth').value);
  const year=parseInt(document.getElementById('hoursYear').value);
  const workerFilter=document.getElementById('hoursWorkerFilter').value;
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];

  const filtered=allHours.filter(h=>{
    if(!h.Date)return false;const d=new Date(h.Date);
    if(d.getMonth()!==month||d.getFullYear()!==year)return false;
    if(workerFilter!=='all'&&(h.Worker||'').toLowerCase()!==workerFilter.toLowerCase())return false;
    return true;
  }).sort((a,b)=>new Date(a.Date)-new Date(b.Date));

  const workerName=workerFilter==='all'?'All workers':(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase())||{}).DisplayName||workerFilter;
  document.getElementById('hoursTitle').textContent='Hours — '+months[month]+' '+year+' — '+workerName;

  const body=document.getElementById('hoursBody');
  if(!filtered.length){body.innerHTML='<tr><td colspan="8" class="loading">No hours registered</td></tr>';document.getElementById('hoursTotal').textContent='0.00';return}

  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];let total=0;
  body.innerHTML=filtered.map(h=>{
    const hrs=calcHoursDiff(h.Time_From,h.Time_To);total+=hrs;
    const d=new Date(h.Date);
    const workerUser=allUsers.find(u=>(u.Epost||'').toLowerCase()===(h.Worker||'').toLowerCase());
    const wName=workerUser?workerUser.DisplayName:(h.Worker||'');
    return'<tr data-hours-id="'+h.id+'" style="cursor:pointer"><td>'+days[d.getDay()]+' '+formatDate(h.Date)+'</td><td>'+(h.Location||'')+'</td><td>'+wName+'</td><td>'+(h.Time_From||'')+'</td><td>'+(h.Time_To||'')+'</td><td style="text-align:right">'+hrs.toFixed(2)+'</td>'
      +'<td class="muted" style="font-size:11px">'+(h.Notes||'')+'</td>'
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
  document.getElementById('hNotes').value=h.Notes||'';
  populateHoursWorkerSelect(h.Worker);
  document.getElementById('hoursModal').classList.add('open');
}

function populateHoursWorkerSelect(preselectedWorker){
  const ws=document.getElementById('hWorker');
  const selected=(preselectedWorker||currentUser.email).toLowerCase();
  if(can('edit_others_hours')){
    const workers=allUsers.filter(u=>u.Active!==false);
    ws.innerHTML=workers.map(u=>'<option value="'+(u.Epost||'')+'"'+((u.Epost||'').toLowerCase()===selected?' selected':'')+'>'+u.DisplayName+'</option>').join('');
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
  const workerName=workerUser?workerUser.DisplayName:worker;
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

function exportHoursExcel(){
  const month=parseInt(document.getElementById('hoursMonth').value);const year=parseInt(document.getElementById('hoursYear').value);
  const workerFilter=document.getElementById('hoursWorkerFilter').value;
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const filtered=allHours.filter(h=>{if(!h.Date)return false;const d=new Date(h.Date);if(d.getMonth()!==month||d.getFullYear()!==year)return false;if(workerFilter!=='all'&&(h.Worker||'').toLowerCase()!==workerFilter.toLowerCase())return false;return true}).sort((a,b)=>new Date(a.Date)-new Date(b.Date));
  const headers=['Date','Day','Location','Worker','From','To','Hours','Notes'];let total=0;
  const rows=filtered.map(h=>{const hrs=calcHoursDiff(h.Time_From,h.Time_To);total+=hrs;const d=new Date(h.Date);const wu=allUsers.find(u=>(u.Epost||'').toLowerCase()===(h.Worker||'').toLowerCase());return[formatDate(h.Date),days[d.getDay()],h.Location||'',wu?wu.DisplayName:h.Worker||'',h.Time_From||'',h.Time_To||'',hrs.toFixed(2),h.Notes||'']});
  rows.push(['','','','','','','Total',total.toFixed(2)]);
  const workerName=workerFilter==='all'?'All':(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase())||{}).DisplayName||workerFilter;
  downloadCSV('Hours_'+workerName.replace(/\s+/g,'_')+'_'+months[month]+'_'+year,headers,rows);
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
function openAdminPanel(){if(!can('admin'))return;renderAdminUsers();document.getElementById('adminModal').classList.add('open')}
function closeAdminPanel(){document.getElementById('adminModal').classList.remove('open')}

function renderAdminUsers(){
  const list=document.getElementById('adminUserList');
  list.innerHTML=allUsers.map(u=>{
    const isSelf=(u.Epost||'').toLowerCase()===currentUser.email;
    const perms=u.Permissions?u.Permissions.split(',').map(s=>s.trim()):[];
    // Map old Role to permissions if no Permissions field
    if(!u.Permissions&&u.Role){
      const roleMap={SuperAdmin:ALL_PERMS.map(p=>p.key),Admin:ALL_PERMS.filter(p=>p.key!=='admin').map(p=>p.key),Cleaner:['cleaning','print_doortag','view_hours','edit_hours'],ReadOnly:['view_bookings']};
      perms.push(...(roleMap[u.Role]||['view_bookings']));
    }
    const assigned=u.AssignedProperties?u.AssignedProperties.split(',').map(s=>s.trim()):[];

    let html='<div style="padding:12px;border-bottom:1px solid var(--border-tertiary)">';
    html+='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">';
    html+='<div><strong>'+(u.DisplayName||'—')+'</strong> <span class="muted" style="font-size:12px">'+(u.Epost||'')+'</span></div>';
    html+='<div style="display:flex;gap:8px;align-items:center">';
    html+='<label style="font-size:12px;display:flex;align-items:center;gap:4px"><input type="checkbox"'+(u.Active!==false?' checked':'')+(isSelf?' disabled':'')+' onchange="toggleUserActive(\''+u.id+'\',this.checked)"> Active</label>';
    html+='<button onclick="sendInviteEmail(\''+u.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit" title="Send login invite email">✉ Invite</button>';
    if(!isSelf)html+='<button onclick="deleteUser(\''+u.id+'\')" style="border:0;background:0 0;color:var(--text-danger);cursor:pointer;font-size:12px">Delete</button>';
    else html+='<span style="font-size:11px;color:var(--text-tertiary)">You</span>';
    html+='</div></div>';
    // Permission checkboxes
    html+='<div class="perm-grid">';
    ALL_PERMS.forEach(p=>{
      const checked=perms.includes(p.key);
      const disabled=isSelf&&p.key==='admin';
      html+='<label><input type="checkbox"'+(checked?' checked':'')+(disabled?' disabled':'')+' onchange="toggleUserPerm(\''+u.id+'\',\''+p.key+'\',this.checked)"> '+p.label+'</label>';
    });
    html+='</div>';
    // Assigned properties
    html+='<div style="margin-top:8px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">';
    html+='<span style="font-size:12px;color:var(--text-secondary)">Locations:</span>';
    properties.forEach(p=>{
      const checked=assigned.includes(p.Title);
      html+='<label style="font-size:12px;display:flex;align-items:center;gap:3px"><input type="checkbox"'+(checked?' checked':'')+' onchange="toggleUserProperty(\''+u.id+'\',\''+p.Title.replace(/'/g,"\\'")+'\',this.checked)"> '+p.Title+'</label>';
    });
    if(assigned.length===0)html+='<span style="font-size:11px;color:var(--text-tertiary);font-style:italic">All (no restriction)</span>';
    html+='</div>';
    html+='</div>';
    return html;
  }).join('');
}

async function toggleUserPerm(userId,perm,enabled){
  const u=allUsers.find(x=>x.id===userId);if(!u)return;
  let perms=u.Permissions?u.Permissions.split(',').map(s=>s.trim()):[];
  // If no Permissions field, derive from Role
  if(!u.Permissions&&u.Role){
    const roleMap={SuperAdmin:ALL_PERMS.map(p=>p.key),Admin:ALL_PERMS.filter(p=>p.key!=='admin').map(p=>p.key),Cleaner:['cleaning','print_doortag','view_hours','edit_hours'],ReadOnly:['view_bookings']};
    perms=roleMap[u.Role]||['view_bookings'];
  }
  if(enabled&&!perms.includes(perm))perms.push(perm);
  if(!enabled)perms=perms.filter(p=>p!==perm);
  const permStr=perms.join(',');
  try{await updateListItem('Users',userId,{Permissions:permStr});u.Permissions=permStr}catch(e){alert('Failed');renderAdminUsers()}
}

async function toggleUserActive(userId,active){
  try{await updateListItem('Users',userId,{Active:active});const u=allUsers.find(x=>x.id===userId);if(u)u.Active=active}catch(e){alert('Failed')}
}

async function toggleUserProperty(userId,propTitle,enabled){
  const u=allUsers.find(x=>x.id===userId);if(!u)return;
  let assigned=u.AssignedProperties?u.AssignedProperties.split(',').map(s=>s.trim()).filter(Boolean):[];
  if(enabled&&!assigned.includes(propTitle))assigned.push(propTitle);
  if(!enabled)assigned=assigned.filter(p=>p!==propTitle);
  const assignedStr=assigned.join(',');
  try{
    await updateListItem('Users',userId,{AssignedProperties:assignedStr});
    u.AssignedProperties=assignedStr;
  }catch(e){alert('Failed — have you created the AssignedProperties column in the Users list?');renderAdminUsers()}
}

async function addUser(){
  const name=document.getElementById('aName').value.trim();const email=document.getElementById('aEmail').value.trim().toLowerCase();
  if(!name||!email){alert('Name and email required');return}
  if(allUsers.find(u=>(u.Epost||'').toLowerCase()===email)){alert('User already exists');return}
  const defaultPerms='view_bookings';
  try{
    const result=await createListItem('Users',{Title:name,DisplayName:name,Epost:email,Permissions:defaultPerms,Active:true});
    allUsers.push({id:result.id||'new',Title:name,DisplayName:name,Epost:email,Permissions:defaultPerms,Active:true});
    document.getElementById('aName').value='';document.getElementById('aEmail').value='';
    renderAdminUsers();
  }catch(e){alert('Failed: '+e.message)}
}

async function deleteUser(userId){
  const u=allUsers.find(x=>x.id===userId);if(!u)return;
  if(!confirm('Delete '+u.DisplayName+'?'))return;
  try{const s=await getSiteId();const lid=await getListId('Users');await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+userId);allUsers=allUsers.filter(x=>x.id!==userId);renderAdminUsers()}catch(e){alert('Failed')}
}

// --- INVITE EMAILS ---
async function sendInviteEmail(userId){
  const u=allUsers.find(x=>x.id===userId);
  if(!u||!u.Epost){alert('No email for this user');return}
  try{
    await sendInviteEmailSilent(u);
    alert('Invite sent to '+u.DisplayName+' ('+u.Epost+')');
  }catch(e){
    console.error('Send mail failed:',e);
    alert('Failed to send invite: '+e.message);
  }
}

async function sendAllInvites(){
  const active=allUsers.filter(u=>u.Active!==false&&u.Epost);
  if(!active.length){alert('No active users to invite');return}
  if(!confirm('Send invite email to '+active.length+' users?\n\n'+active.map(u=>u.DisplayName+' ('+u.Epost+')').join('\n')))return;

  let sent=0,failed=0;
  for(const u of active){
    try{
      await sendInviteEmailSilent(u);
      sent++;
    }catch(e){failed++}
    await new Promise(res=>setTimeout(res,500));
  }
  alert('Done! '+sent+' invites sent'+(failed?', '+failed+' failed':'')+'.');
}

async function sendInviteEmailSilent(u){
  const appUrl='https://franknh-design.github.io/2gmbooking/';
  const body=buildInviteHtml(u);
  await getToken();
  const r=await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
    method:'POST',
    headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
    body:JSON.stringify({message:{subject:'2GM Booking — You have been invited',body:{contentType:'HTML',content:body},toRecipients:[{emailAddress:{address:u.Epost}}]},saveToSentItems:true})
  });
  if(!r.ok)throw new Error('Failed');
}

function buildInviteHtml(u){
  const appUrl='https://franknh-design.github.io/2gmbooking/';
  return'<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px">'
    +'<h2 style="color:#2C2C2A;margin-bottom:4px">Welcome to 2GM Booking</h2>'
    +'<p style="color:#5F5E5A">Hi '+(u.DisplayName||'')+',</p>'
    +'<p>You have been given access to the 2GM Booking system. Click the button below to sign in with your Microsoft account.</p>'
    +'<div style="margin:24px 0"><a href="'+appUrl+'" style="display:inline-block;padding:12px 32px;background:#1D9E75;color:#fff;text-decoration:none;border-radius:8px;font-size:16px;font-weight:500">Open 2GM Booking</a></div>'
    +'<p style="font-size:13px;color:#888">Sign in using your email: <strong>'+(u.Epost||'')+'</strong></p>'
    +'<p style="font-size:13px;color:#888">If you have any questions, contact Frank at frank@2gm.no or +47 99 10 10 41.</p>'
    +'<hr style="border:none;border-top:1px solid #eee;margin:24px 0">'
    +'<p style="font-size:11px;color:#aaa">This email was sent from the 2GM Booking system.</p></div>';
}

// --- OCCUPANCY REPORT ---
function showOccupancyReport(){
  const now=new Date();
  const yearSel=document.getElementById('occYear');
  const propSel=document.getElementById('occProperty');
  if(!yearSel.children.length){
    const curYear=now.getFullYear();
    const years=[];
    allBookings.forEach(b=>{if(b.Check_In){const y=new Date(b.Check_In).getFullYear();if(!years.includes(y))years.push(y)}});
    if(!years.includes(curYear))years.push(curYear);
    years.sort((a,b)=>b-a);
    yearSel.innerHTML=years.map(y=>'<option value="'+y+'"'+(y===curYear?' selected':'')+'>'+y+'</option>').join('');
    propSel.innerHTML='<option value="all">All properties</option>'+properties.map(p=>'<option value="'+p.id+'">'+p.Title+'</option>').join('');
  }
  renderOccupancyReport();
  document.getElementById('occupancyModal').classList.add('open');
}

function renderOccupancyReport(){
  const year=parseInt(document.getElementById('occYear').value);
  const propFilter=document.getElementById('occProperty').value;
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const now=new Date();

  // Get rooms for selected property
  let reportRooms=propFilter==='all'?allRooms.filter(r=>{const pids=new Set(properties.map(p=>p.id));return pids.has(String(r.PropertyLookupId))}):allRooms.filter(r=>String(r.PropertyLookupId)===propFilter);
  const roomCount=reportRooms.length;
  const roomIds=new Set(reportRooms.map(r=>r.id));

  // Relevant bookings
  const yearBookings=allBookings.filter(b=>{
    if(b.Status!=='Active'&&b.Status!=='Completed')return false;
    const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;
    if(!b.Check_In)return false;
    const ci=new Date(b.Check_In);const co=b.Check_Out?new Date(b.Check_Out):now;
    return ci.getFullYear()===year||co.getFullYear()===year||(ci.getFullYear()<year&&co.getFullYear()>year);
  });

  let html='<table style="width:100%;font-size:13px;border-collapse:collapse"><thead><tr style="border-bottom:2px solid var(--border-secondary)"><th style="text-align:left;padding:6px">Month</th><th style="text-align:right;padding:6px">Room nights</th><th style="text-align:right;padding:6px">Possible</th><th style="text-align:right;padding:6px">Occupancy</th></tr></thead><tbody>';

  let totalOccupied=0,totalPossible=0;

  for(let m=0;m<12;m++){
    const monthStart=new Date(year,m,1);
    const monthEnd=new Date(year,m+1,0);// last day
    const daysInMonth=monthEnd.getDate();

    // Don't count future months
    if(monthStart>now){html+='<tr style="color:var(--text-tertiary)"><td style="padding:4px 6px">'+months[m]+'</td><td style="text-align:right;padding:4px 6px">—</td><td style="text-align:right;padding:4px 6px">—</td><td style="text-align:right;padding:4px 6px">—</td></tr>';continue}

    const lastDay=monthEnd<now?daysInMonth:now.getDate();
    const possible=roomCount*lastDay;

    let occupied=0;
    yearBookings.forEach(b=>{
      const ci=new Date(b.Check_In);const co=b.Check_Out?new Date(b.Check_Out):now;
      const start=ci>monthStart?ci:monthStart;
      const end=co<new Date(year,m,lastDay+1)?co:new Date(year,m,lastDay+1);
      const nights=Math.max(0,Math.round((end-start)/864e5));
      occupied+=nights;
    });

    // Cap at possible (can't have more than 100%)
    occupied=Math.min(occupied,possible);
    totalOccupied+=occupied;totalPossible+=possible;
    const pct=possible>0?Math.round(occupied/possible*100):0;
    const barColor=pct>=80?'var(--accent)':pct>=50?'#EF9F27':'var(--text-danger)';

    html+='<tr style="border-bottom:1px solid var(--border-tertiary)"><td style="padding:6px">'+months[m]+'</td><td style="text-align:right;padding:6px">'+occupied+'</td><td style="text-align:right;padding:6px">'+possible+'</td><td style="text-align:right;padding:6px"><div style="display:flex;align-items:center;justify-content:flex-end;gap:8px"><div style="width:80px;height:8px;background:var(--bg-secondary);border-radius:4px;overflow:hidden"><div style="height:100%;width:'+pct+'%;background:'+barColor+';border-radius:4px"></div></div><strong>'+pct+'%</strong></div></td></tr>';
  }

  const totalPct=totalPossible>0?Math.round(totalOccupied/totalPossible*100):0;
  html+='</tbody><tfoot><tr style="border-top:2px solid var(--border-secondary);font-weight:500"><td style="padding:6px">Total '+year+'</td><td style="text-align:right;padding:6px">'+totalOccupied+'</td><td style="text-align:right;padding:6px">'+totalPossible+'</td><td style="text-align:right;padding:6px"><strong>'+totalPct+'%</strong></td></tr></tfoot></table>';
  html+='<div style="margin-top:8px;font-size:12px;color:var(--text-tertiary)">Based on '+roomCount+' rooms. Active and completed bookings only.</div>';

  document.getElementById('occReport').innerHTML=html;
}

function exportOccupancyReport(){
  const year=parseInt(document.getElementById('occYear').value);
  const propFilter=document.getElementById('occProperty').value;
  const propName=propFilter==='all'?'All':(properties.find(p=>p.id===propFilter)||{}).Title||'Unknown';
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const now=new Date();
  let reportRooms=propFilter==='all'?allRooms.filter(r=>{const pids=new Set(properties.map(p=>p.id));return pids.has(String(r.PropertyLookupId))}):allRooms.filter(r=>String(r.PropertyLookupId)===propFilter);
  const roomCount=reportRooms.length;const roomIds=new Set(reportRooms.map(r=>r.id));
  const yearBookings=allBookings.filter(b=>{if(b.Status!=='Active'&&b.Status!=='Completed')return false;const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;if(!b.Check_In)return false;const ci=new Date(b.Check_In);const co=b.Check_Out?new Date(b.Check_Out):now;return ci.getFullYear()===year||co.getFullYear()===year||(ci.getFullYear()<year&&co.getFullYear()>year)});
  const headers=['Month','Room nights','Possible','Occupancy %'];
  const rows=[];let tO=0,tP=0;
  for(let m=0;m<12;m++){
    const monthStart=new Date(year,m,1);const monthEnd=new Date(year,m+1,0);const daysInMonth=monthEnd.getDate();
    if(monthStart>now){rows.push([months[m],'','','']);continue}
    const lastDay=monthEnd<now?daysInMonth:now.getDate();const possible=roomCount*lastDay;
    let occupied=0;yearBookings.forEach(b=>{const ci=new Date(b.Check_In);const co=b.Check_Out?new Date(b.Check_Out):now;const start=ci>monthStart?ci:monthStart;const end=co<new Date(year,m,lastDay+1)?co:new Date(year,m,lastDay+1);occupied+=Math.max(0,Math.round((end-start)/864e5))});
    occupied=Math.min(occupied,possible);tO+=occupied;tP+=possible;
    rows.push([months[m],occupied,possible,possible>0?Math.round(occupied/possible*100)+'%':'']);
  }
  rows.push(['Total',tO,tP,tP>0?Math.round(tO/tP*100)+'%':'']);
  downloadCSV('Occupancy_'+propName.replace(/\s+/g,'_')+'_'+year,headers,rows);
}

// --- RATES MANAGEMENT ---
function openRatesPanel(){
  if(!can('manage_rates')&&!can('admin')){alert('Access denied');return}
  renderRatesPanel();
  document.getElementById('ratesModal').classList.add('open');
}

function renderRatesPanel(){
  // Property default rates
  const propList=document.getElementById('ratesPropertyList');
  propList.innerHTML='<table style="font-size:13px;width:100%"><thead><tr><th>Property</th><th style="width:120px">Daily rate (kr)</th></tr></thead><tbody>'
    +properties.map(p=>{
      return'<tr><td>'+p.Title+'</td><td><input type="number" value="'+(p.DailyRate||'')+'" onchange="updatePropertyRate(\''+p.id+'\',this.value)" style="width:100%;padding:4px 6px;border:1px solid var(--border-tertiary);border-radius:4px;font-size:13px;text-align:right" placeholder="0"></td></tr>';
    }).join('')+'</tbody></table>';

  // Custom rates
  const customList=document.getElementById('ratesCustomList');
  if(!allRates.length){
    customList.innerHTML='<div class="muted" style="font-size:13px;padding:8px">No custom rates set</div>';
  }else{
    customList.innerHTML='<table style="font-size:13px;width:100%"><thead><tr><th>Company</th><th>Person</th><th>Property</th><th style="width:80px">Rate</th><th style="width:30px"></th></tr></thead><tbody>'
      +allRates.map(r=>{
        return'<tr><td>'+(r.Company||'<span class="muted">—</span>')+'</td><td>'+(r.Person_Name||'<span class="muted">—</span>')+'</td><td>'+(r.Property||'<span class="muted">All</span>')+'</td><td style="text-align:right">'+(r.DailyRate||0)+' kr</td>'
          +'<td><button onclick="deleteRate(\''+r.id+'\')" style="width:20px;height:20px;border-radius:50%;border:1px solid var(--border-tertiary);background:var(--bg-primary);color:var(--text-danger);cursor:pointer;font-size:11px;padding:0">✕</button></td></tr>';
      }).join('')+'</tbody></table>';
  }

  // Property select for new rate
  const propSel=document.getElementById('rProperty');
  propSel.innerHTML='<option value="">All properties</option>'+properties.map(p=>'<option value="'+p.Title+'">'+p.Title+'</option>').join('');
}

async function updatePropertyRate(propId,value){
  const rate=parseFloat(value)||0;
  try{
    await updateListItem('Properties',propId,{DailyRate:rate});
    const p=properties.find(x=>x.id===propId);if(p)p.DailyRate=rate;
  }catch(e){alert('Failed: '+e.message)}
}

async function addCustomRate(){
  const company=document.getElementById('rCompany').value.trim();
  const person=document.getElementById('rPerson').value.trim();
  const propName=document.getElementById('rProperty').value;
  const rate=parseFloat(document.getElementById('rRate').value);

  if(!company&&!person){alert('Company or person name required');return}
  if(!rate||rate<=0){alert('Rate must be a positive number');return}

  const fields={
    Title:(person||company)+' — '+(propName||'All')+' — '+rate+'kr',
    Company:company||'',
    Person_Name:person||'',
    Property:propName||'',
    DailyRate:rate
  };

  try{
    const result=await createListItem('Rates',fields);
    allRates.push({id:result.id||'new',...fields});
    document.getElementById('rCompany').value='';
    document.getElementById('rPerson').value='';
    document.getElementById('rRate').value='';
    renderRatesPanel();
  }catch(e){alert('Failed: '+e.message)}
}

async function deleteRate(id){
  if(!confirm('Delete this rate?'))return;
  try{
    const s=await getSiteId();const lid=await getListId('Rates');
    await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+id);
    allRates=allRates.filter(r=>r.id!==id);
    renderRatesPanel();
  }catch(e){alert('Failed')}
}
