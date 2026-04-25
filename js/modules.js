// ============================================================
// 2GM Booking v14.0.10 — modules.js
// Hours, Archive, Import/Export, Admin (checkbox permissions)
// ============================================================

// --- UPCOMING ---
async function checkInFromUpcoming(id){
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  if(!confirm('Check in '+b.Person_Name+' now?\n\nThis will mark the booking as Active with today\'s date.'))return;
  try{
    const now=new Date().toISOString();
    await updateListItem('Bookings',id,{Status:'Active',Check_In:now});
    b.Status='Active';b.Check_In=now;
    refreshLocal();renderIncoming();loadData();
  }catch(e){console.error(e);alert('Failed to check in: '+e.message)}
}
function toggleIncoming(){
  ensureMainView();
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const panel=document.getElementById('incomingPanel');
  panel.classList.toggle('open');
  const isOpen=panel.classList.contains('open');
  document.getElementById('mainView').classList.toggle('panel-mode',isOpen);
  if(isOpen)renderIncoming();
  updateNavActiveState();
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
      +'<td style="font-weight:500">'+roomTitle+'</td><td>'+guestMarkedName(b.Person_Name||'')+'</td><td class="muted">'+(b.Company||'')+'</td>'
      +'<td>'+formatDate(b.Check_In)+' '+badge+'</td><td>'+(b.Check_Out?formatDate(b.Check_Out):'Open-ended')+'</td>'
      +'<td><span class="pill" style="background:var(--bg-warning);color:var(--text-warning)">Upcoming</span></td>'
      +'<td><button onclick="event.stopPropagation();checkInFromUpcoming(\''+b.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--accent);color:#fff;cursor:pointer;font-size:11px;font-family:inherit;margin-right:4px" title="Check in now">Check in</button>'
      +'<button onclick="event.stopPropagation();openEditBooking(\''+b.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit">Edit</button></td></tr>';
  }).join('');
}

// --- ARCHIVE ---
function toggleArchive(){
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const panel=document.getElementById('archivePanel');
  panel.classList.toggle('open');
  const isOpen=panel.classList.contains('open');
  document.getElementById('mainView').classList.toggle('panel-mode',isOpen);
  if(isOpen)renderArchive();
  updateNavActiveState();
}
let archivePage=0; // 0-indexed
const ARCHIVE_PAGE_SIZE=50;

function renderArchive(){
  const search=(document.getElementById('archiveSearch').value||'').toLowerCase();
  const statusFilter=document.getElementById('archiveStatus').value;
  const fromVal=document.getElementById('archiveFrom').value;
  const toVal=document.getElementById('archiveTo').value;
  const fromDate=fromVal?new Date(fromVal+'T00:00:00'):null;
  const toDate=toVal?new Date(toVal+'T23:59:59'):null;
  const roomIds=new Set(rooms.map(r=>r.id));
  let archived=allBookings.filter(b=>{
    const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;
    if(statusFilter!=='all'&&b.Status!==statusFilter)return false;
    if(fromDate||toDate){
      if(!b.Check_In)return false;
      const ci=new Date(b.Check_In);
      if(fromDate&&ci<fromDate)return false;
      if(toDate&&ci>toDate)return false;
    }
    if(search){if(!((b.Person_Name||'')+(b.Company||'')+getRoomTitle(b)).toLowerCase().includes(search))return false}
    return true;
  }).sort((a,b)=>new Date(b.Check_In||0)-new Date(a.Check_In||0));

  const totalPages=Math.max(1,Math.ceil(archived.length/ARCHIVE_PAGE_SIZE));
  if(archivePage>=totalPages)archivePage=totalPages-1;
  if(archivePage<0)archivePage=0;
  const start=archivePage*ARCHIVE_PAGE_SIZE;
  const pageItems=archived.slice(start,start+ARCHIVE_PAGE_SIZE);

  let titleSuffix='';
  if(fromVal||toVal){titleSuffix=' — '+(fromVal?formatDate(fromVal):'…')+' → '+(toVal?formatDate(toVal):'…')}
  document.getElementById('archiveTitle').textContent='Archive — '+archived.length+' booking'+(archived.length!==1?'s':'')+titleSuffix;
  const body=document.getElementById('archiveBody');
  // Re-render charts if open
  if(document.getElementById('archiveChartsContainer').style.display!=='none'){renderArchiveCharts(archived)}
  if(!pageItems.length){body.innerHTML='<tr><td colspan="7" class="loading">No bookings found</td></tr>';renderArchivePagination(0,1);return}
  body.innerHTML=pageItems.map(b=>{
    const sc={Completed:'background:var(--bg-secondary);color:var(--text-secondary)',Cancelled:'background:var(--bg-danger);color:var(--text-danger)',Active:'background:var(--bg-success);color:var(--text-success)',Upcoming:'background:var(--bg-warning);color:var(--text-warning)'}[b.Status]||'';
    return'<tr><td style="font-weight:500">'+getRoomTitle(b)+'</td><td>'+(b.Person_Name?guestMarkedName(b.Person_Name):'—')+'</td><td class="muted">'+(b.Company||'')+'</td><td>'+formatDate(b.Check_In)+'</td><td>'+(b.Check_Out?formatDate(b.Check_Out):'Open-ended')+'</td><td><span class="pill" style="'+sc+'">'+b.Status+'</span></td>'
      +'<td>'+(b.Status==='Completed'||b.Status==='Cancelled'?'<button onclick="reopenBooking(\''+b.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit">Reopen</button>':'')+'</td></tr>';
  }).join('');
  renderArchivePagination(archived.length,totalPages);
}

function renderArchivePagination(totalItems,totalPages){
  let pager=document.getElementById('archivePagination');
  if(!pager){
    pager=document.createElement('div');pager.id='archivePagination';
    pager.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:10px 14px;border-top:1px solid var(--border-tertiary);background:var(--bg-secondary);font-size:12px';
    const bodyWrap=document.getElementById('archiveBody').closest('div[style*="overflow-y:auto"]');
    if(bodyWrap&&bodyWrap.parentNode)bodyWrap.parentNode.appendChild(pager);
  }
  if(totalItems<=ARCHIVE_PAGE_SIZE){pager.style.display='none';return}
  pager.style.display='flex';
  const start=archivePage*ARCHIVE_PAGE_SIZE+1;
  const end=Math.min((archivePage+1)*ARCHIVE_PAGE_SIZE,totalItems);
  // Build page numbers — show first, last, current +-2
  const pageNums=new Set([0,totalPages-1,archivePage]);
  for(let i=Math.max(0,archivePage-2);i<=Math.min(totalPages-1,archivePage+2);i++)pageNums.add(i);
  const sortedPages=[...pageNums].sort((a,b)=>a-b);
  let btns='';
  sortedPages.forEach((p,i)=>{
    if(i>0&&p-sortedPages[i-1]>1)btns+='<span style="color:var(--text-tertiary);margin:0 4px">…</span>';
    const isActive=p===archivePage;
    btns+='<button onclick="goArchivePage('+p+')" style="padding:4px 10px;margin:0 2px;border:1px solid var(--border-tertiary);border-radius:4px;background:'+(isActive?'var(--accent)':'var(--bg-primary)')+';color:'+(isActive?'#fff':'var(--text-primary)')+';cursor:pointer;font-size:12px;min-width:28px">'+(p+1)+'</button>';
  });
  pager.innerHTML='<div>Showing <strong>'+start+'–'+end+'</strong> of <strong>'+totalItems+'</strong></div>'
    +'<div style="display:flex;align-items:center;gap:4px">'
    +'<button onclick="goArchivePage('+(archivePage-1)+')" '+(archivePage===0?'disabled':'')+' style="padding:4px 10px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-primary);cursor:pointer;font-size:12px;'+(archivePage===0?'opacity:.4;cursor:not-allowed':'')+'">← Prev</button>'
    +btns
    +'<button onclick="goArchivePage('+(archivePage+1)+')" '+(archivePage>=totalPages-1?'disabled':'')+' style="padding:4px 10px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-primary);cursor:pointer;font-size:12px;'+(archivePage>=totalPages-1?'opacity:.4;cursor:not-allowed':'')+'">Next →</button>'
    +'</div>';
}

function goArchivePage(p){archivePage=p;renderArchive();
  // Scroll archive body to top so user sees new page
  const bodyWrap=document.getElementById('archiveBody').closest('div[style*="overflow-y:auto"]');
  if(bodyWrap)bodyWrap.scrollTop=0;
}
function resetArchivePage(){archivePage=0;renderArchive()}

function clearArchiveDateRange(){
  document.getElementById('archiveFrom').value='';document.getElementById('archiveTo').value='';
  archivePage=0;
  renderArchive();
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

  const workerName=workerFilter==='all'?'All workers':(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase())||{}).DisplayName||workerFilter;
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
    const wName=workerUser?workerUser.DisplayName:(h.Worker||'');
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
    return[formatDate(h.Date),days[d.getDay()],h.Location||'',wu?wu.DisplayName:h.Worker||'',h.Time_From||'',h.Time_To||'',hrs.toFixed(2),_hoursNotes(h)];
  });
  // Total row: push to Hours column only, leave Notes empty
  rows.push(['','','','','','Total',total.toFixed(2),'']);
  const workerName=workerFilter==='all'?'All':(allUsers.find(u=>(u.Epost||'').toLowerCase()===workerFilter.toLowerCase())||{}).DisplayName||workerFilter;
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
    html+='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;flex-wrap:wrap;gap:8px">';
    html+='<div><strong>'+(u.DisplayName||'—')+'</strong> <span class="muted" style="font-size:12px">'+(u.Epost||'')+'</span></div>';
    html+='<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">';
    // Role dropdown
    const roles=['SuperAdmin','Admin','Cleaner','ReadOnly','Custom'];
    const currentRole=u.Role||'Custom';
    html+='<label style="font-size:12px;display:flex;align-items:center;gap:4px">Role: <select onchange="setUserRole(\''+u.id+'\',this.value)" style="padding:3px 6px;border:1px solid var(--border-tertiary);border-radius:4px;font-size:12px;font-family:inherit"'+(isSelf?' disabled':'')+'>'
      +roles.map(r=>'<option value="'+r+'"'+(r===currentRole?' selected':'')+'>'+r+'</option>').join('')+'</select></label>';
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

async function setUserRole(userId,role){
  const u=allUsers.find(x=>x.id===userId);if(!u)return;
  const roleMap={SuperAdmin:ALL_PERMS.map(p=>p.key),Admin:ALL_PERMS.filter(p=>p.key!=='admin').map(p=>p.key),Cleaner:['cleaning','print_doortag','view_hours','edit_hours'],ReadOnly:['view_bookings']};
  const fields={Role:role};
  // If picking a real role, also set permissions. For "Custom", just save role and keep current perms.
  if(roleMap[role]){fields.Permissions=roleMap[role].join(',')}
  try{
    await updateListItem('Users',userId,fields);
    u.Role=role;
    if(fields.Permissions)u.Permissions=fields.Permissions;
    renderAdminUsers();
  }catch(e){alert('Failed to update role: '+e.message);renderAdminUsers()}
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
  // Filter out properties without a Title (corrupt/undefined entries)
  const validProperties=properties.filter(p=>p.Title&&p.Title.trim());

  // Property default rates
  const propList=document.getElementById('ratesPropertyList');
  propList.innerHTML='<table style="font-size:13px;width:100%"><thead><tr><th>Property</th><th style="width:120px">Daily rate (kr)</th></tr></thead><tbody>'
    +validProperties.map(p=>{
      return'<tr><td>'+p.Title+'</td><td><input type="number" value="'+(p.DailyRate||'')+'" onchange="updatePropertyRate(\''+p.id+'\',this.value)" style="width:100%;padding:4px 6px;border:1px solid var(--border-tertiary);border-radius:4px;font-size:13px;text-align:right" placeholder="0"></td></tr>';
    }).join('')+'</tbody></table>';

  // Room property selector — always populate
  const rpSel=document.getElementById('rRoomProperty');
  const prevVal=rpSel.value;
  rpSel.innerHTML=validProperties.map(p=>'<option value="'+p.id+'">'+p.Title+'</option>').join('');
  if(prevVal&&[...rpSel.options].some(o=>o.value===prevVal))rpSel.value=prevVal;
  renderRoomRates();

  // Split custom rates into Nightly, Checkout, Percent
  const nightlyRates=allRates.filter(r=>{const t=(r.FeeType||'').toLowerCase();return t!=='checkout'&&t!=='percent'});
  const checkoutRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='checkout');
  const percentRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='percent');

  const customList=document.getElementById('ratesCustomList');
  let customHtml='';
  // Nightly section
  if(!nightlyRates.length){
    customHtml+='<div class="muted" style="font-size:13px;padding:8px">No custom nightly rates set</div>';
  }else{
    customHtml+='<table style="font-size:13px;width:100%"><thead><tr><th>Company</th><th>Person</th><th>Property</th><th style="width:80px">Rate/night</th><th style="width:30px"></th></tr></thead><tbody>'
      +nightlyRates.map(r=>{
        return'<tr><td>'+(r.Company||'<span class="muted">—</span>')+'</td><td>'+(r.Person_Name||'<span class="muted">—</span>')+'</td><td>'+(r.Property||'<span class="muted">All</span>')+'</td><td style="text-align:right">'+(r.DailyRate||0)+' kr</td>'
          +'<td><button onclick="deleteRate(\''+r.id+'\')" style="width:20px;height:20px;border-radius:50%;border:1px solid var(--border-tertiary);background:var(--bg-primary);color:var(--text-danger);cursor:pointer;font-size:11px;padding:0">✕</button></td></tr>';
      }).join('')+'</tbody></table>';
  }
  // Checkout section
  customHtml+='<div style="margin-top:18px;padding-top:14px;border-top:1px solid var(--border-tertiary)"><strong style="font-size:13px">🧹 Checkout fees (utvask)</strong> <span class="muted" style="font-size:11px">one-time fee added at checkout</span></div>';
  if(!checkoutRates.length){
    customHtml+='<div class="muted" style="font-size:13px;padding:8px">No checkout fees configured.</div>';
  }else{
    customHtml+='<table style="font-size:13px;width:100%;margin-top:6px"><thead><tr><th>Company</th><th>Property</th><th style="width:100px">Fee</th><th style="width:30px"></th></tr></thead><tbody>'
      +checkoutRates.map(r=>{
        return'<tr style="background:rgba(123,97,255,.04)"><td>'+(r.Company||'<span class="muted">All companies</span>')+'</td><td>'+(r.Property||'<span class="muted">All</span>')+'</td><td style="text-align:right;font-weight:500">'+(r.DailyRate||0)+' kr</td>'
          +'<td><button onclick="deleteRate(\''+r.id+'\')" style="width:20px;height:20px;border-radius:50%;border:1px solid var(--border-tertiary);background:var(--bg-primary);color:var(--text-danger);cursor:pointer;font-size:11px;padding:0">✕</button></td></tr>';
      }).join('')+'</tbody></table>';
  }
  // Percent section (month-based, replaces Checkout for that company)
  customHtml+='<div style="margin-top:18px;padding-top:14px;border-top:1px solid var(--border-tertiary)"><strong style="font-size:13px">📊 Percent of month</strong> <span class="muted" style="font-size:11px">% of company\'s monthly revenue, replaces flat checkout fee</span></div>';
  if(!percentRates.length){
    customHtml+='<div class="muted" style="font-size:13px;padding:8px">No percent-based fees configured.</div>';
  }else{
    customHtml+='<table style="font-size:13px;width:100%;margin-top:6px"><thead><tr><th>Company</th><th>Property</th><th style="width:100px">%</th><th style="width:30px"></th></tr></thead><tbody>'
      +percentRates.map(r=>{
        return'<tr style="background:rgba(239,159,39,.06)"><td>'+(r.Company||'<span class="muted">All companies</span>')+'</td><td>'+(r.Property||'<span class="muted">All</span>')+'</td><td style="text-align:right;font-weight:500">'+(r.DailyRate||0)+' %</td>'
          +'<td><button onclick="deleteRate(\''+r.id+'\')" style="width:20px;height:20px;border-radius:50%;border:1px solid var(--border-tertiary);background:var(--bg-primary);color:var(--text-danger);cursor:pointer;font-size:11px;padding:0">✕</button></td></tr>';
      }).join('')+'</tbody></table>';
  }
  customList.innerHTML=customHtml;

  // Property select for new rate
  const propSel=document.getElementById('rProperty');
  propSel.innerHTML='<option value="">All properties</option>'+validProperties.map(p=>'<option value="'+p.Title+'">'+p.Title+'</option>').join('');

  // Company datalist — from bookings + existing rates
  const companies=new Set();
  allBookings.forEach(b=>{if(b.Company)companies.add(b.Company)});
  allRates.forEach(r=>{if(r.Company)companies.add(r.Company)});
  document.getElementById('rCompanyList').innerHTML=[...companies].sort().map(c=>'<option value="'+c+'">').join('');

  // Person datalist — from bookings + persons list
  const persons=new Set();
  allBookings.forEach(b=>{if(b.Person_Name)persons.add(b.Person_Name)});
  allPersons.forEach(p=>{if(p.Title)persons.add(p.Title);if(p.Name)persons.add(p.Name)});
  document.getElementById('rPersonList').innerHTML=[...persons].sort().map(p=>'<option value="'+p+'">').join('');
}

function renderRoomRates(){
  const propId=document.getElementById('rRoomProperty').value;
  const propRooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(propId))
    .sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));

  // Only show rooms that have a rate set, or all if few rooms (<30)
  const roomsWithRate=propRooms.filter(r=>r.DailyRate);
  const showAll=propRooms.length<=30;
  const displayRooms=showAll?propRooms:roomsWithRate;

  const container=document.getElementById('ratesRoomList');
  if(!displayRooms.length){
    container.innerHTML='<div class="muted" style="font-size:13px;padding:8px">'+(showAll?'No rooms in this property':'No rooms with individual rates set')+'</div>';
    return;
  }

  container.innerHTML='<table style="font-size:13px;width:100%"><thead><tr><th>Room</th><th>Floor</th><th style="width:120px">Daily rate (kr)</th></tr></thead><tbody>'
    +displayRooms.map(r=>{
      const hasRate=r.DailyRate?'':'color:var(--text-tertiary)';
      return'<tr><td style="font-weight:500">'+r.Title+'</td><td class="muted">Floor '+(r.Floor||'?')+'</td>'
        +'<td><input type="number" value="'+(r.DailyRate||'')+'" onchange="updateRoomRate(\''+r.id+'\',this.value)" style="width:100%;padding:4px 6px;border:1px solid var(--border-tertiary);border-radius:4px;font-size:13px;text-align:right;'+hasRate+'" placeholder="Use property default"></td></tr>';
    }).join('')+'</tbody></table>'
    +(!showAll&&propRooms.length>roomsWithRate.length?'<div style="font-size:11px;color:var(--text-tertiary);margin-top:4px">'+roomsWithRate.length+' of '+propRooms.length+' rooms have individual rates. Rooms without rate use property default.</div>':'');
}

async function updatePropertyRate(propId,value){
  const rate=parseFloat(value)||0;
  try{
    await updateListItem('Properties',propId,{DailyRate:rate});
    const p=properties.find(x=>x.id===propId);if(p)p.DailyRate=rate;
  }catch(e){alert('Failed: '+e.message)}
}

async function updateRoomRate(roomId,value){
  const rate=value?parseFloat(value):null;
  try{
    await updateListItem('Rooms',roomId,{DailyRate:rate||0});
    const r=allRooms.find(x=>x.id===roomId);if(r)r.DailyRate=rate||0;
  }catch(e){alert('Failed: '+e.message)}
}

async function addCustomRate(){
  const company=document.getElementById('rCompany').value.trim();
  const person=document.getElementById('rPerson').value.trim();
  const propName=document.getElementById('rProperty').value;
  const feeType=document.getElementById('rFeeType').value||'Nightly';
  const rate=parseFloat(document.getElementById('rRate').value);

  // Validation rules per fee type
  if(feeType==='Checkout'){
    if(!propName){alert('Property is required for checkout fees');return}
  }else if(feeType==='Percent'){
    if(!company){alert('Company is required for percent-based fees');return}
    if(rate<=0||rate>100){alert('Percent must be between 0 and 100');return}
  }else{
    if(!company&&!person){alert('Company or person name required');return}
  }
  if(!rate||rate<=0){alert('Rate must be a positive number');return}

  const titleSuffix=feeType==='Percent'?rate+'%':rate+'kr'+(feeType==='Checkout'?' (utvask)':'');
  const fields={
    Title:(person||company||'Checkout')+' — '+(propName||'All')+' — '+titleSuffix,
    Company:company||'',
    Person_Name:(feeType==='Checkout'||feeType==='Percent')?'':(person||''),
    Property:propName||'',
    DailyRate:rate,
    FeeType:feeType
  };

  try{
    const result=await createListItem('Rates',fields);
    allRates.push({id:result.id||'new',...fields});
    document.getElementById('rCompany').value='';
    document.getElementById('rPerson').value='';
    document.getElementById('rRate').value='';
    document.getElementById('rFeeType').value='Nightly';
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

// ============================================================
// PERSONS / CUSTOMERS (v14.0.10)
// ============================================================
let editingPersonId=null;

function togglePersons(){
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const panel=document.getElementById('personsPanel');
  panel.classList.toggle('open');
  const isOpen=panel.classList.contains('open');
  document.getElementById('mainView').classList.toggle('panel-mode',isOpen);
  if(isOpen)renderPersons();
  updateNavActiveState();
}

function _personName(p){return p.Name||p.Title||''}
function _personCompany(p){return p.Company||''}

function renderPersons(){
  const body=document.getElementById('personsBody');
  if(!body)return;
  const q=(document.getElementById('personsSearch').value||'').toLowerCase().trim();
  let list=allPersons.slice();
  if(q){
    list=list.filter(p=>{
      const hay=[_personName(p),_personCompany(p),p.Email||'',p.Company_Email||'',p.Mobile||p.Phone||p.Telefon||'',p.Address||'',p.Company_OrgNr||''].join(' ').toLowerCase();
      return hay.indexOf(q)>=0;
    });
  }
  list.sort((a,b)=>_personName(a).localeCompare(_personName(b),'nb'));
  if(!list.length){body.innerHTML='<tr><td colspan="8" class="loading">No guests found. Click "+ New guest" to add one.</td></tr>';return}
  // Count bookings per person name (case-insensitive)
  const bookingCount={};
  allBookings.forEach(b=>{
    const n=(b.Person_Name||'').toLowerCase();
    if(n)bookingCount[n]=(bookingCount[n]||0)+1;
  });
  body.innerHTML=list.map(p=>{
    const name=_personName(p);
    const mobile=p.Mobile||p.Phone||p.Telefon||'';
    const company=_personCompany(p);
    const bookings=bookingCount[name.toLowerCase()]||0;
    const addr=(p.Address||'').replace(/\n/g,', ');
    // Find active booking for this person (fuzzy name match)
    const active=findActiveBookingForPerson(name);
    let activeCell='<span class="muted">—</span>';
    if(active){
      const room=allRooms.find(r=>r.id===String(active.RoomLookupId));
      const roomTitle=room?room.Title:'?';
      const propName=active.Property_Name||'';
      // Check for today's wash
      let todayBadge='';
      if(active.Check_In){
        const washes=calcWashDates(active.Check_In,active.Check_Out);
        const todayWash=washes.find(w=>w.isToday);
        if(todayWash)todayBadge=' <span class="pill danger" style="font-size:10px">Today — '+todayWash.type+'</span>';
        else if(active.Cleaning_Status==='Dirty')todayBadge=' <span class="pill danger" style="font-size:10px">Needs cleaning</span>';
      }
      activeCell='<span style="display:inline-flex;align-items:center;gap:4px;flex-wrap:wrap"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--accent)"></span>'
        +'<strong style="color:var(--text-success)">Room '+escapeHtml(roomTitle)+'</strong>'
        +(propName?' <span class="muted" style="font-size:11px">('+escapeHtml(propName)+')</span>':'')
        +todayBadge
        +'</span>';
    }
    const nameCell=active
      ?'<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--accent);vertical-align:middle;margin-right:6px"></span>'+escapeHtml(name)
      :escapeHtml(name);
    // Bookings count clickable if > 0
    const bookingsCell=bookings
      ?'<span class="pill" style="background:var(--bg-success);color:var(--text-success);cursor:pointer;text-decoration:underline" onclick="event.stopPropagation();showGuestBookings(\''+escapeHtml(name).replace(/'/g,"\\'")+'\')" title="Show bookings">'+bookings+'</span>'
      :'<span class="muted">0</span>';
    const canEdit=can('edit_bookings');
    const rowStyle=canEdit?'cursor:pointer':'';
    const rowOnclick=canEdit?'onclick="openPersonEdit(\''+p.id+'\')"':'';
    const rowHover=canEdit?'onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'"':'';
    const editBtn=canEdit
      ?'<td onclick="event.stopPropagation()"><button onclick="openPersonEdit(\''+p.id+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit">Edit</button></td>'
      :'<td></td>';
    return '<tr '+rowOnclick+' style="'+rowStyle+'" '+rowHover+'>'
      +'<td style="font-weight:500">'+nameCell+'</td>'
      +'<td>'+activeCell+'</td>'
      +'<td>'+escapeHtml(company)+'</td>'
      +'<td onclick="event.stopPropagation()">'+(mobile?'<a href="tel:'+escapeHtml(mobile)+'" style="color:var(--accent)">'+escapeHtml(mobile)+'</a>':'<span class="muted">—</span>')+'</td>'
      +'<td onclick="event.stopPropagation()">'+(p.Email?'<a href="mailto:'+escapeHtml(p.Email)+'" style="color:var(--accent)">'+escapeHtml(p.Email)+'</a>':'<span class="muted">—</span>')+'</td>'
      +'<td class="muted" style="font-size:11px">'+escapeHtml(addr)+'</td>'
      +'<td>'+bookingsCell+'</td>'
      +editBtn
      +'</tr>';
  }).join('');
}

// Find active booking matching person name (fuzzy: exact or reverse-order match)
function findActiveBookingForPerson(name){
  if(!name)return null;
  const lower=name.toLowerCase().trim();
  // First: exact case-insensitive match, Active status
  let m=allBookings.find(b=>b.Status==='Active'&&(b.Person_Name||'').toLowerCase().trim()===lower);
  if(m)return m;
  // Second: match all words regardless of order (handles "Marek Filas" vs "Filas, Marek" etc.)
  const words=lower.split(/[\s,]+/).filter(w=>w.length>1);
  if(words.length<2)return null;
  m=allBookings.find(b=>{
    if(b.Status!=='Active')return false;
    const bn=(b.Person_Name||'').toLowerCase();
    return words.every(w=>bn.indexOf(w)>=0);
  });
  return m||null;
}

function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]))}

function openPersonEdit(personId){
  if(!can('edit_bookings')){alert('You do not have permission to edit guests.');return}
  editingPersonId=personId||null;
  const p=personId?allPersons.find(x=>x.id===personId):null;
  document.getElementById('personModalTitle').textContent=p?'Edit guest':'New guest';
  document.getElementById('pName').value=p?_personName(p):'';
  document.getElementById('pMobile').value=p?(p.Mobile||p.Phone||p.Telefon||''):'';
  document.getElementById('pEmail').value=p?(p.Email||''):'';
  document.getElementById('pAddress').value=p?(p.Address||''):'';
  document.getElementById('pCompany').value=p?_personCompany(p):'';
  document.getElementById('pCompanyOrgNr').value=p?(p.Company_OrgNr||''):'';
  document.getElementById('pCompanyEmail').value=p?(p.Company_Email||''):'';
  document.getElementById('pCompanyAddress').value=p?(p.Company_Address||''):'';
  document.getElementById('pNotes').value=p?(p.Notes||''):'';
  document.getElementById('pDeleteRow').style.display=p?'':'none';
  // Render booking history for this person
  renderPersonHistory(p);
  const modal=document.getElementById('personModal');
  modal.classList.add('open');
  const mc=modal.querySelector('.modal');if(mc)mc.scrollTop=0;
  modal.scrollTop=0;
}

function renderPersonHistory(p){
  const section=document.getElementById('pHistorySection');
  const body=document.getElementById('pHistoryBody');
  const summary=document.getElementById('pHistorySummary');
  if(!p||!section){if(section)section.style.display='none';return}
  const name=_personName(p);
  if(!name){section.style.display='none';return}
  // Fuzzy-match all bookings for this person
  const lower=name.toLowerCase().trim();
  const words=lower.split(/[\s,]+/).filter(w=>w.length>1);
  const matching=allBookings.filter(b=>{
    const bn=(b.Person_Name||'').toLowerCase().trim();
    if(bn===lower)return true;
    if(words.length<2)return false;
    const bwords=bn.split(/[\s,]+/).filter(w=>w.length>1);
    if(bwords.length<2)return false;
    return words.every(w=>bn.indexOf(w)>=0)||bwords.every(w=>lower.indexOf(w)>=0);
  }).sort((a,b)=>new Date(b.Check_In||0)-new Date(a.Check_In||0));
  if(!matching.length){section.style.display='none';return}
  section.style.display='';
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
  summary.textContent=matching.length+' booking'+(matching.length!==1?'s':'')+' · '+totalNights+' nights · '+totalRevenue.toLocaleString('nb-NO')+' kr';
  let html='<table style="width:100%;font-size:11px"><thead><tr style="background:var(--bg-tertiary)">'
    +'<th style="padding:5px 8px;text-align:left">Room</th>'
    +'<th style="padding:5px 8px;text-align:left">Property</th>'
    +'<th style="padding:5px 8px;text-align:left">Period</th>'
    +'<th style="padding:5px 8px;text-align:right">Nights</th>'
    +'<th style="padding:5px 8px;text-align:left">Status</th>'
    +'</tr></thead><tbody>';
  matching.forEach(b=>{
    const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
    const roomTitle=room?room.Title:'?';
    const ci=b.Check_In?new Date(b.Check_In):null;
    const co=b.Check_Out?new Date(b.Check_Out):(b.Status==='Active'?new Date():null);
    const nights=ci&&co?Math.max(0,Math.round((co-ci)/864e5)):0;
    const statusColor={Completed:'background:var(--bg-secondary);color:var(--text-secondary)',Cancelled:'background:var(--bg-danger);color:var(--text-danger)',Active:'background:var(--bg-success);color:var(--text-success)',Upcoming:'background:var(--bg-warning);color:var(--text-warning)'}[b.Status]||'';
    const period=(ci?formatDate(b.Check_In):'—')+' → '+(b.Check_Out?formatDate(b.Check_Out):(b.Status==='Active'?'Open':'—'));
    html+='<tr onclick="document.getElementById(\'personModal\').classList.remove(\'open\');openEditBooking(\''+b.id+'\')" style="border-top:.5px solid var(--border-tertiary);cursor:pointer" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
      +'<td style="padding:5px 8px;font-weight:500">'+escapeHtml(roomTitle)+'</td>'
      +'<td style="padding:5px 8px">'+escapeHtml(b.Property_Name||'')+'</td>'
      +'<td style="padding:5px 8px">'+period+'</td>'
      +'<td style="padding:5px 8px;text-align:right">'+nights+'</td>'
      +'<td style="padding:5px 8px"><span class="pill" style="'+statusColor+';font-size:10px">'+b.Status+'</span></td>'
      +'</tr>';
  });
  html+='</tbody></table>';
  body.innerHTML=html;
}

async function savePerson(){
  if(!can('edit_bookings')){alert('You do not have permission to save guests.');return}
  const name=document.getElementById('pName').value.trim();
  if(!name){alert('Name is required');return}
  const fields={
    Title:name,
    Name:name,
    Mobile:document.getElementById('pMobile').value.trim()||null,
    Email:document.getElementById('pEmail').value.trim()||null,
    Address:document.getElementById('pAddress').value.trim()||null,
    Company:document.getElementById('pCompany').value.trim()||null,
    Company_OrgNr:document.getElementById('pCompanyOrgNr').value.trim()||null,
    Company_Email:document.getElementById('pCompanyEmail').value.trim()||null,
    Company_Address:document.getElementById('pCompanyAddress').value.trim()||null,
    Notes:document.getElementById('pNotes').value.trim()||null
  };
  const btn=document.getElementById('personSaveBtn');
  btn.disabled=true;btn.textContent='Saving...';
  try{
    if(editingPersonId){
      await updateListItem('Persons',editingPersonId,fields);
      const l=allPersons.find(x=>x.id===editingPersonId);
      if(l)Object.assign(l,fields);
    }else{
      const r=await createListItem('Persons',fields);
      if(r&&r.id)allPersons.push({id:r.id,...fields});
      else{try{allPersons=await getListItems('Persons')}catch(e){}}
    }
    document.getElementById('personModal').classList.remove('open');
    renderPersons();
    refreshPersonDatalists();
  }catch(e){console.error(e);alert('Failed to save: '+e.message)}
  finally{btn.disabled=false;btn.textContent='Save'}
}

async function deletePerson(){
  if(!can('edit_bookings')){alert('You do not have permission to delete guests.');return}
  if(!editingPersonId)return;
  const p=allPersons.find(x=>x.id===editingPersonId);
  if(!p)return;
  if(!confirm('Delete '+_personName(p)+'?\n\nThis does NOT delete their bookings. The bookings will still show the name as free text.'))return;
  try{
    const s=await getSiteId();const lid=await getListId('Persons');
    await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+editingPersonId);
    allPersons=allPersons.filter(x=>x.id!==editingPersonId);
    document.getElementById('personModal').classList.remove('open');
    renderPersons();refreshPersonDatalists();
  }catch(e){alert('Failed to delete: '+e.message)}
}

// --- Autocomplete integration in booking modal ---
function refreshPersonDatalists(){
  const nameList=document.getElementById('fNameList');
  const compList=document.getElementById('fCompanyList');
  if(nameList){
    const names=[...new Set(allPersons.map(_personName).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'nb'));
    nameList.innerHTML=names.map(n=>'<option value="'+escapeHtml(n)+'">').join('');
  }
  if(compList){
    // Merge active Companies + historical company names from bookings/persons
    const fromCompanies=allCompanies.filter(c=>c.Active!==false).map(c=>c.Title).filter(Boolean);
    const fromPersons=allPersons.map(_personCompany).filter(Boolean);
    const fromBookings=allBookings.map(b=>[b.Company,b.Billing_Company]).flat().filter(Boolean);
    const companies=[...new Set([...fromCompanies,...fromPersons,...fromBookings])].sort((a,b)=>a.localeCompare(b,'nb'));
    compList.innerHTML=companies.map(c=>'<option value="'+escapeHtml(c)+'">').join('');
  }
}

// Called when user types/selects in the guest name field
function onPersonNameInput(){
  const val=(document.getElementById('fName').value||'').trim();
  const info=document.getElementById('fNameInfo');
  if(!val){info.textContent='';return}
  const match=allPersons.find(p=>_personName(p).toLowerCase()===val.toLowerCase());
  // Check for existing active booking with this name (independent of person-card match)
  const activeBooking=findActiveBookingForPerson(val);
  // If we're editing an existing booking, don't warn about that same booking
  let activeWarning='';
  if(activeBooking&&activeBooking.id!==editingBookingId){
    const room=allRooms.find(r=>r.id===String(activeBooking.RoomLookupId));
    const roomTitle=room?room.Title:'?';
    const propName=activeBooking.Property_Name||'';
    activeWarning='<div style="margin-top:4px;padding:6px 8px;background:var(--bg-warning);border:1px solid #EF9F27;border-radius:4px;color:var(--text-warning);font-size:11px">'
      +'⚠ <strong>Already has an active booking</strong> in Room '+escapeHtml(roomTitle)+(propName?' ('+escapeHtml(propName)+')':'')
      +' — check-in '+formatDate(activeBooking.Check_In)+'</div>';
  }
  if(match){
    // Auto-fill company only if empty (don't overwrite user's manual entry)
    const compField=document.getElementById('fCompany');
    if(!compField.value.trim()&&_personCompany(match)){
      compField.value=_personCompany(match);
    }
    const mobile=match.Mobile||match.Phone||match.Telefon||'';
    const parts=[];
    if(mobile)parts.push('📱 '+mobile);
    if(match.Email)parts.push('✉ '+match.Email);
    info.innerHTML='<span style="color:var(--text-success)">✓ Existing person</span>'+(parts.length?' · '+escapeHtml(parts.join(' · ')):'')+activeWarning;
  }else{
    info.innerHTML='<span class="muted">New name — will be saved as free text (create person card in Persons panel to enable autofill next time)</span>'+activeWarning;
  }
}

// ============================================================
// CHARTS (v14.0.10) — pure SVG, no dependencies
// ============================================================

// Reusable bar chart: data = [{label, value, subtitle?}]
function svgBarChart(data, opts){
  opts=opts||{};
  const width=opts.width||640;
  const barHeight=opts.barHeight||24;
  const gap=opts.gap||6;
  const labelW=opts.labelW||140;
  const rightPad=opts.rightPad||60;
  const maxV=Math.max(1,...data.map(d=>d.value||0));
  const height=data.length*(barHeight+gap)+10;
  const color=opts.color||'#1D9E75';
  let svg='<svg width="100%" viewBox="0 0 '+width+' '+height+'" xmlns="http://www.w3.org/2000/svg" style="font-family:-apple-system,Segoe UI,sans-serif;font-size:12px">';
  data.forEach((d,i)=>{
    const y=i*(barHeight+gap)+5;
    const barW=((d.value||0)/maxV)*(width-labelW-rightPad);
    const label=escapeHtml(d.label||'');
    const val=opts.formatValue?opts.formatValue(d.value):d.value;
    svg+='<text x="'+(labelW-8)+'" y="'+(y+barHeight/2+4)+'" text-anchor="end" fill="#2C2C2A">'+label+'</text>';
    svg+='<rect x="'+labelW+'" y="'+y+'" width="'+barW+'" height="'+barHeight+'" fill="'+color+'" rx="3"/>';
    svg+='<text x="'+(labelW+barW+6)+'" y="'+(y+barHeight/2+4)+'" fill="#5F5E5A">'+val+'</text>';
    if(d.subtitle){svg+='<text x="'+(labelW-8)+'" y="'+(y+barHeight/2+16)+'" text-anchor="end" fill="#888780" font-size="10">'+escapeHtml(d.subtitle)+'</text>'}
  });
  svg+='</svg>';
  return svg;
}

// Reusable line-like area chart for time series
// series = [{date: Date, value: number}] already sorted ascending
function svgTimeSeries(series, opts){
  opts=opts||{};
  const width=opts.width||800;
  const height=opts.height||180;
  const padL=36, padR=12, padT=12, padB=28;
  const innerW=width-padL-padR, innerH=height-padT-padB;
  const color=opts.color||'#1D9E75';
  if(!series.length)return '<svg width="100%" viewBox="0 0 '+width+' '+height+'"><text x="'+width/2+'" y="'+height/2+'" text-anchor="middle" fill="#888780" font-family="sans-serif" font-size="12">No data</text></svg>';
  const maxV=Math.max(1,...series.map(s=>s.value));
  // Y-axis ticks: 4 steps
  let svg='<svg width="100%" viewBox="0 0 '+width+' '+height+'" xmlns="http://www.w3.org/2000/svg" style="font-family:-apple-system,Segoe UI,sans-serif;font-size:10px">';
  for(let i=0;i<=4;i++){
    const y=padT+(innerH/4)*i;
    const v=Math.round(maxV*(1-i/4)*100)/100;
    svg+='<line x1="'+padL+'" y1="'+y+'" x2="'+(width-padR)+'" y2="'+y+'" stroke="#e5e4df" stroke-width="1"/>';
    svg+='<text x="'+(padL-6)+'" y="'+(y+3)+'" text-anchor="end" fill="#888780">'+v+'</text>';
  }
  // Bars (easier to read than line for sparse daily data)
  const barW=innerW/series.length*0.7;
  const step=innerW/series.length;
  series.forEach((s,i)=>{
    const h=(s.value/maxV)*innerH;
    const x=padL+i*step+(step-barW)/2;
    const y=padT+innerH-h;
    svg+='<rect x="'+x+'" y="'+y+'" width="'+barW+'" height="'+h+'" fill="'+color+'" rx="2"><title>'+escapeHtml(s.label||'')+': '+s.value+'</title></rect>';
  });
  // X-axis labels: show first, last, middle
  const showIdx=new Set();
  if(series.length)showIdx.add(0);
  if(series.length>1)showIdx.add(series.length-1);
  if(series.length>4)showIdx.add(Math.floor(series.length/2));
  if(series.length>8){showIdx.add(Math.floor(series.length/4));showIdx.add(Math.floor(3*series.length/4))}
  series.forEach((s,i)=>{
    if(!showIdx.has(i))return;
    const x=padL+i*step+step/2;
    svg+='<text x="'+x+'" y="'+(height-8)+'" text-anchor="middle" fill="#5F5E5A">'+escapeHtml(s.label||'')+'</text>';
  });
  svg+='</svg>';
  return svg;
}

// --- ARCHIVE CHARTS ---
function toggleArchiveCharts(){
  const c=document.getElementById('archiveChartsContainer');
  const isOpen=c.style.display!=='none';
  c.style.display=isOpen?'none':'block';
  const btn=document.getElementById('archiveChartsBtn');
  btn.style.background=isOpen?'var(--bg-secondary)':'var(--accent)';
  btn.style.color=isOpen?'':'#fff';
  if(!isOpen)renderArchive();
}

function renderArchiveCharts(archived){
  const c=document.getElementById('archiveChartsContainer');
  if(!archived||!archived.length){c.innerHTML='<div style="text-align:center;padding:20px;color:var(--text-secondary);font-size:13px">No data to chart — adjust filters above</div>';return}

  // --- Card layout ---
  let html='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:14px">';

  // Chart 1: Bookings per month (check-in date)
  const byMonth={};
  archived.forEach(b=>{
    if(!b.Check_In)return;
    const d=new Date(b.Check_In);
    const k=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
    byMonth[k]=(byMonth[k]||0)+1;
  });
  const monthKeys=Object.keys(byMonth).sort();
  const monthNames=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const monthSeries=monthKeys.map(k=>{
    const [y,m]=k.split('-');
    return {label:monthNames[parseInt(m)-1]+' '+y.slice(2),value:byMonth[k]};
  });
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Bookings per month (by check-in)</div>'
    +svgTimeSeries(monthSeries,{height:160,color:'#1D9E75'})
    +'</div>';

  // Chart 2: Guest-nights per month — actual occupancy load
  const nightsPerMonth={};
  archived.forEach(b=>{
    if(!b.Check_In)return;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):new Date();co.setHours(0,0,0,0);
    // Walk each night
    const cur=new Date(ci);
    while(cur<co){
      const k=cur.getFullYear()+'-'+String(cur.getMonth()+1).padStart(2,'0');
      nightsPerMonth[k]=(nightsPerMonth[k]||0)+1;
      cur.setDate(cur.getDate()+1);
    }
  });
  const nightSeries=Object.keys(nightsPerMonth).sort().map(k=>{
    const [y,m]=k.split('-');
    return {label:monthNames[parseInt(m)-1]+' '+y.slice(2),value:nightsPerMonth[k]};
  });
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Guest-nights per month</div>'
    +svgTimeSeries(nightSeries,{height:160,color:'#EF9F27'})
    +'</div>';

  // Chart 3: Top companies
  const byCompany={};
  archived.forEach(b=>{
    const c=(b.Company||'(no company)').trim()||'(no company)';
    if(!byCompany[c])byCompany[c]={bookings:0,nights:0};
    byCompany[c].bookings++;
    if(b.Check_In){
      const ci=new Date(b.Check_In);
      const co=b.Check_Out?new Date(b.Check_Out):new Date();
      const nights=Math.max(1,Math.round((co-ci)/864e5));
      byCompany[c].nights+=nights;
    }
  });
  const topCompanies=Object.entries(byCompany)
    .sort((a,b)=>b[1].nights-a[1].nights).slice(0,10)
    .map(([name,d])=>({label:name.length>20?name.slice(0,19)+'…':name,value:d.nights,subtitle:d.bookings+' booking'+(d.bookings!==1?'s':'')}));
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Top 10 companies (by guest-nights)</div>'
    +svgBarChart(topCompanies,{barHeight:22,gap:10,labelW:150,color:'#1D9E75'})
    +'</div>';

  // Chart 4: Stay length distribution
  const buckets={'1-3 nights':0,'4-7 nights':0,'8-14 nights':0,'15-30 nights':0,'31-90 nights':0,'90+ nights':0};
  archived.forEach(b=>{
    if(!b.Check_In)return;
    const ci=new Date(b.Check_In);
    const co=b.Check_Out?new Date(b.Check_Out):new Date();
    const n=Math.max(1,Math.round((co-ci)/864e5));
    if(n<=3)buckets['1-3 nights']++;
    else if(n<=7)buckets['4-7 nights']++;
    else if(n<=14)buckets['8-14 nights']++;
    else if(n<=30)buckets['15-30 nights']++;
    else if(n<=90)buckets['31-90 nights']++;
    else buckets['90+ nights']++;
  });
  const stayData=Object.entries(buckets).map(([k,v])=>({label:k,value:v}));
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Stay length distribution</div>'
    +svgBarChart(stayData,{barHeight:22,gap:8,labelW:110,color:'#5B8AC4'})
    +'</div>';

  // Summary numbers
  const totalNights=Object.values(nightsPerMonth).reduce((a,b)=>a+b,0);
  const totalBookings=archived.length;
  const avgStay=totalBookings?Math.round(totalNights/totalBookings*10)/10:0;
  const uniqueGuests=new Set(archived.map(b=>(b.Person_Name||'').toLowerCase()).filter(Boolean)).size;
  const uniqueCompanies=new Set(archived.map(b=>(b.Company||'').toLowerCase()).filter(Boolean)).size;
  const completed=archived.filter(b=>b.Status==='Completed').length;
  const cancelled=archived.filter(b=>b.Status==='Cancelled').length;
  const cancelRate=totalBookings?Math.round(cancelled/totalBookings*1000)/10:0;

  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);grid-column:1/-1">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Summary</div>'
    +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px">'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Bookings</div><div style="font-size:20px;font-weight:500">'+totalBookings+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Guest-nights</div><div style="font-size:20px;font-weight:500">'+totalNights+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Avg. stay (nights)</div><div style="font-size:20px;font-weight:500">'+avgStay+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Unique guests</div><div style="font-size:20px;font-weight:500">'+uniqueGuests+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Companies</div><div style="font-size:20px;font-weight:500">'+uniqueCompanies+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Cancel rate</div><div style="font-size:20px;font-weight:500">'+cancelRate+'%</div></div>'
    +'</div></div>';

  html+='</div>';
  c.innerHTML=html;
}

// --- HOURS CHARTS ---
function toggleHoursCharts(){
  const c=document.getElementById('hoursChartsContainer');
  const isOpen=c.style.display!=='none';
  c.style.display=isOpen?'none':'block';
  const btn=document.getElementById('hoursChartsBtn');
  btn.style.background=isOpen?'var(--bg-secondary)':'var(--accent)';
  btn.style.color=isOpen?'':'#fff';
  // Close efficiency to avoid clash
  if(!isOpen){
    const ec=document.getElementById('efficiencyContainer');
    if(ec&&ec.style.display!=='none'){
      ec.style.display='none';
      const eb=document.getElementById('efficiencyBtn');
      if(eb){eb.style.background='var(--bg-secondary)';eb.style.color=''}
    }
    renderHours();
  }
}

function renderHoursCharts(filtered){
  const c=document.getElementById('hoursChartsContainer');
  if(!c)return;
  if(!filtered||!filtered.length){c.innerHTML='<div style="text-align:center;padding:20px;color:var(--text-secondary);font-size:13px">No hours data to chart — adjust filters above</div>';return}

  // Helper: hours for one entry
  const hoursOf=h=>calcHoursDiff(h.Time_From,h.Time_To);

  let html='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:14px">';

  // Chart 1: Hours per day (time series in visible period)
  const byDay={};
  filtered.forEach(h=>{
    if(!h.Date)return;
    const d=new Date(h.Date);
    const k=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
    byDay[k]=(byDay[k]||0)+hoursOf(h);
  });
  const dayKeys=Object.keys(byDay).sort();
  const daySeries=dayKeys.map(k=>{
    const [y,m,d]=k.split('-');
    return {label:d+'.'+m,value:Math.round(byDay[k]*100)/100};
  });
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);grid-column:1/-1">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Hours per day</div>'
    +svgTimeSeries(daySeries,{height:180,color:'#1D9E75'})
    +'</div>';

  // Chart 2: Hours per location
  const byLoc={};
  filtered.forEach(h=>{
    const l=h.Location||'(no location)';
    byLoc[l]=(byLoc[l]||0)+hoursOf(h);
  });
  const locData=Object.entries(byLoc)
    .sort((a,b)=>b[1]-a[1])
    .map(([k,v])=>({label:k.length>18?k.slice(0,17)+'…':k,value:Math.round(v*100)/100}));
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Hours per location</div>'
    +svgBarChart(locData,{barHeight:22,gap:8,labelW:130,color:'#EF9F27',formatValue:v=>v+' h'})
    +'</div>';

  // Chart 3: Hours per worker (only if multiple workers visible)
  const byWorker={};
  filtered.forEach(h=>{
    const w=h.Worker||'(unknown)';
    const u=allUsers.find(x=>(x.Epost||'').toLowerCase()===w.toLowerCase());
    const name=u?u.DisplayName:w;
    byWorker[name]=(byWorker[name]||0)+hoursOf(h);
  });
  if(Object.keys(byWorker).length>1){
    const workerData=Object.entries(byWorker)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>({label:k.length>18?k.slice(0,17)+'…':k,value:Math.round(v*100)/100}));
    html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
      +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Hours per worker</div>'
      +svgBarChart(workerData,{barHeight:22,gap:8,labelW:130,color:'#5B8AC4',formatValue:v=>v+' h'})
      +'</div>';
  }

  // Chart 4: Weekday distribution
  const byDow={0:0,1:0,2:0,3:0,4:0,5:0,6:0};
  filtered.forEach(h=>{if(!h.Date)return;byDow[new Date(h.Date).getDay()]+=hoursOf(h)});
  const dowNames=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  // Re-order Mon-Sun
  const dowData=[1,2,3,4,5,6,0].map(i=>({label:dowNames[i],value:Math.round(byDow[i]*100)/100}));
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:8px;color:var(--text-secondary)">Hours per weekday</div>'
    +svgBarChart(dowData,{barHeight:22,gap:8,labelW:50,color:'#9B7EC4',formatValue:v=>v+' h'})
    +'</div>';

  // Summary
  const totalHours=Math.round(Object.values(byLoc).reduce((a,b)=>a+b,0)*100)/100;
  const uniqueDays=Object.keys(byDay).length;
  const avgPerDay=uniqueDays?Math.round(totalHours/uniqueDays*100)/100:0;
  const nWorkers=Object.keys(byWorker).length;
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary)">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Summary</div>'
    +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:10px">'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Total hours</div><div style="font-size:20px;font-weight:500">'+totalHours+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Days worked</div><div style="font-size:20px;font-weight:500">'+uniqueDays+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Avg/day</div><div style="font-size:20px;font-weight:500">'+avgPerDay+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Entries</div><div style="font-size:20px;font-weight:500">'+filtered.length+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Workers</div><div style="font-size:20px;font-weight:500">'+nWorkers+'</div></div>'
    +'</div></div>';

  html+='</div>';
  c.innerHTML=html;
}

// ============================================================
// CLEANING EFFICIENCY ANALYSIS (v14.0.10)
// ============================================================
// Compares cleaner hours against guest-nights per property, per week/month.
// USE WITH CAUTION: Hours include breaks, transport, repairs — not just cleaning.
// Use this to spot trends and big deviations, not as absolute performance metric.

let efficiencyMode='month'; // 'week' or 'month'

function toggleEfficiency(){
  if(!can('view_efficiency')){alert('You do not have permission to view efficiency analysis.');return}
  const c=document.getElementById('efficiencyContainer');
  const isOpen=c.style.display!=='none';
  c.style.display=isOpen?'none':'block';
  const btn=document.getElementById('efficiencyBtn');
  btn.style.background=isOpen?'var(--bg-secondary)':'var(--accent)';
  btn.style.color=isOpen?'':'#fff';
  // Close charts view to avoid visual clash
  if(!isOpen){
    const cc=document.getElementById('hoursChartsContainer');
    if(cc&&cc.style.display!=='none'){toggleHoursCharts()}
    renderEfficiency();
  }
}

function setEfficiencyMode(m){efficiencyMode=m;renderEfficiency()}

// Find all users with Cleaner role (from SharePoint Users list)
function _getCleanerEmails(){
  return new Set(allUsers.filter(u=>u.Role==='Cleaner').map(u=>(u.Epost||'').toLowerCase()).filter(Boolean));
}

// ISO week number (Monday-start) for a date
function _isoWeek(d){
  const x=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  const day=x.getUTCDay()||7;
  x.setUTCDate(x.getUTCDate()+4-day);
  const yearStart=new Date(Date.UTC(x.getUTCFullYear(),0,1));
  const wk=Math.ceil(((x-yearStart)/864e5+1)/7);
  return x.getUTCFullYear()+'-W'+String(wk).padStart(2,'0');
}
function _monthKey(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')}

// Distribute a stay into day-buckets (returns {bucketKey: nights})
function _distributeNights(checkIn,checkOut,bucketFn){
  const buckets={};
  if(!checkIn)return buckets;
  const ci=new Date(checkIn);ci.setHours(0,0,0,0);
  const co=checkOut?new Date(checkOut):new Date();co.setHours(0,0,0,0);
  if(co<=ci)return buckets;
  const cur=new Date(ci);
  while(cur<co){
    const k=bucketFn(cur);
    buckets[k]=(buckets[k]||0)+1;
    cur.setDate(cur.getDate()+1);
  }
  return buckets;
}

function renderEfficiency(){
  const c=document.getElementById('efficiencyContainer');
  if(!c)return;
  const cleanerEmails=_getCleanerEmails();
  if(!cleanerEmails.size){
    c.innerHTML='<div style="padding:20px;text-align:center;color:var(--text-secondary);font-size:13px">'
      +'<strong>No cleaners configured</strong><br>'
      +'Efficiency analysis requires at least one user with role "Cleaner" in the admin panel.'
      +'</div>';
    return;
  }

  // Determine active period (from Hours filter)
  const fromVal=document.getElementById('hoursFrom').value;
  const toVal=document.getElementById('hoursTo').value;
  const month=parseInt(document.getElementById('hoursMonth').value);
  const year=parseInt(document.getElementById('hoursYear').value);
  const useRange=!!(fromVal||toVal);
  const fromDate=useRange?(fromVal?new Date(fromVal+'T00:00:00'):new Date(1970,0,1)):new Date(year,month,1);
  const toDate=useRange?(toVal?new Date(toVal+'T23:59:59'):new Date(2100,0,1)):new Date(year,month+1,0,23,59,59);

  const bucketFn=efficiencyMode==='week'?_isoWeek:_monthKey;

  // --- Collect cleaner hours by property + bucket ---
  // hoursByPropBucket[propertyTitle][bucketKey] = hours
  const hoursByPropBucket={};
  const totalCleanerHours={}; // per bucket, across all properties
  allHours.forEach(h=>{
    if(!h.Date||!h.Time_From||!h.Time_To)return;
    if(!cleanerEmails.has((h.Worker||'').toLowerCase()))return;
    const d=new Date(h.Date);
    if(d<fromDate||d>toDate)return;
    const loc=h.Location||'(unknown)';
    const hrs=calcHoursDiff(h.Time_From,h.Time_To);
    if(!hrs)return;
    const bk=bucketFn(d);
    if(!hoursByPropBucket[loc])hoursByPropBucket[loc]={};
    hoursByPropBucket[loc][bk]=(hoursByPropBucket[loc][bk]||0)+hrs;
    totalCleanerHours[bk]=(totalCleanerHours[bk]||0)+hrs;
  });

  // --- Collect guest-nights by property + bucket ---
  // nightsByPropBucket[propertyTitle][bucketKey] = nights
  const nightsByPropBucket={};
  const totalNights={};
  allBookings.forEach(b=>{
    if(!b.Check_In)return;
    const propName=b.Property_Name;
    if(!propName)return;
    const buckets=_distributeNights(b.Check_In,b.Check_Out,bucketFn);
    Object.entries(buckets).forEach(([bk,n])=>{
      // Filter by date window — only count buckets that overlap our period
      // Parse bucket back to rough date for filtering
      let bucketDate;
      if(efficiencyMode==='week'){const [y,w]=bk.split('-W');bucketDate=_dateFromIsoWeek(parseInt(y),parseInt(w))}
      else{const [y,m]=bk.split('-');bucketDate=new Date(parseInt(y),parseInt(m)-1,15)}
      if(bucketDate<fromDate||bucketDate>toDate)return;
      if(!nightsByPropBucket[propName])nightsByPropBucket[propName]={};
      nightsByPropBucket[propName][bk]=(nightsByPropBucket[propName][bk]||0)+n;
      totalNights[bk]=(totalNights[bk]||0)+n;
    });
  });

  // --- Build property list (union of both sources, sorted) ---
  const propNames=[...new Set([...Object.keys(hoursByPropBucket),...Object.keys(nightsByPropBucket)])].sort();
  const allBuckets=[...new Set([...Object.keys(totalCleanerHours),...Object.keys(totalNights)])].sort();

  // --- Render ---
  let html='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px">'
    +'<div><strong style="font-size:14px">🧹 Cleaning efficiency</strong>'
    +'<div style="font-size:11px;color:var(--text-tertiary);margin-top:2px">Only counts hours from users with role "Cleaner". Period follows Hours filter above.</div></div>'
    +'<div style="display:flex;gap:4px">'
    +'<button onclick="setEfficiencyMode(\'week\')" style="padding:5px 12px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:12px;background:'+(efficiencyMode==='week'?'var(--accent)':'var(--bg-primary)')+';color:'+(efficiencyMode==='week'?'#fff':'var(--text-primary)')+';cursor:pointer">Per week</button>'
    +'<button onclick="setEfficiencyMode(\'month\')" style="padding:5px 12px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:12px;background:'+(efficiencyMode==='month'?'var(--accent)':'var(--bg-primary)')+';color:'+(efficiencyMode==='month'?'#fff':'var(--text-primary)')+';cursor:pointer">Per month</button>'
    +'</div></div>';

  if(!allBuckets.length){
    html+='<div style="padding:20px;text-align:center;color:var(--text-secondary);font-size:13px">No cleaner hours or guest-nights found in this period.</div>';
    c.innerHTML=html;return;
  }

  // --- Summary cards ---
  const totH=Object.values(totalCleanerHours).reduce((a,b)=>a+b,0);
  const totN=Object.values(totalNights).reduce((a,b)=>a+b,0);
  const overallMinPerNight=totN?Math.round(totH*60/totN):0;
  html+='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px">'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Cleaner hours</div><div style="font-size:20px;font-weight:500">'+Math.round(totH*10)/10+'</div></div>'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Guest-nights</div><div style="font-size:20px;font-weight:500">'+totN+'</div></div>'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Minutes / guest-night</div><div style="font-size:20px;font-weight:500">'+overallMinPerNight+'</div></div>'
    +'</div>';

  // --- Combined bar chart: hours vs nights per bucket (across all properties) ---
  // Show as overlapping bars so user sees correlation
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);margin-bottom:14px">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Overall — hours vs guest-nights '+(efficiencyMode==='week'?'(per week)':'(per month)')+'</div>';
  // Render dual-axis chart manually
  const chartW=800,chartH=180,padL=45,padR=45,padT=10,padB=28;
  const innerW=chartW-padL-padR,innerH=chartH-padT-padB;
  const maxH=Math.max(1,...Object.values(totalCleanerHours));
  const maxN=Math.max(1,...Object.values(totalNights));
  let svg='<svg width="100%" viewBox="0 0 '+chartW+' '+chartH+'" xmlns="http://www.w3.org/2000/svg" style="font-family:-apple-system,Segoe UI,sans-serif;font-size:10px">';
  // Grid
  for(let i=0;i<=4;i++){
    const y=padT+(innerH/4)*i;
    svg+='<line x1="'+padL+'" y1="'+y+'" x2="'+(chartW-padR)+'" y2="'+y+'" stroke="#e5e4df" stroke-width="1"/>';
    svg+='<text x="'+(padL-6)+'" y="'+(y+3)+'" text-anchor="end" fill="#1D9E75">'+Math.round(maxH*(1-i/4)*10)/10+'</text>';
    svg+='<text x="'+(chartW-padR+6)+'" y="'+(y+3)+'" text-anchor="start" fill="#EF9F27">'+Math.round(maxN*(1-i/4))+'</text>';
  }
  const step=innerW/allBuckets.length;
  const barW=step*0.35;
  allBuckets.forEach((bk,i)=>{
    const h=totalCleanerHours[bk]||0;
    const n=totalNights[bk]||0;
    const xCenter=padL+i*step+step/2;
    // Left bar: hours
    const barHH=(h/maxH)*innerH;
    svg+='<rect x="'+(xCenter-barW-1)+'" y="'+(padT+innerH-barHH)+'" width="'+barW+'" height="'+barHH+'" fill="#1D9E75" rx="2"><title>'+bk+': '+Math.round(h*10)/10+' h</title></rect>';
    // Right bar: nights
    const barNH=(n/maxN)*innerH;
    svg+='<rect x="'+(xCenter+1)+'" y="'+(padT+innerH-barNH)+'" width="'+barW+'" height="'+barNH+'" fill="#EF9F27" rx="2"><title>'+bk+': '+n+' nights</title></rect>';
    // X label
    if(allBuckets.length<=12||i%Math.ceil(allBuckets.length/10)===0){
      const shortLbl=efficiencyMode==='week'?bk.slice(5):bk.slice(5);
      svg+='<text x="'+xCenter+'" y="'+(chartH-10)+'" text-anchor="middle" fill="#5F5E5A">'+shortLbl+'</text>';
    }
  });
  svg+='</svg>';
  html+=svg
    +'<div style="display:flex;gap:16px;font-size:11px;margin-top:6px;color:var(--text-secondary)"><span><span style="display:inline-block;width:10px;height:10px;background:#1D9E75;border-radius:2px;vertical-align:middle"></span> Cleaner hours (left axis)</span><span><span style="display:inline-block;width:10px;height:10px;background:#EF9F27;border-radius:2px;vertical-align:middle"></span> Guest-nights (right axis)</span></div>'
    +'</div>';

  // --- Per-property table: nights, hours, minutes per night per bucket ---
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);margin-bottom:14px">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Per property — minutes cleaning per guest-night</div>'
    +'<div style="overflow-x:auto"><table style="font-size:12px;min-width:500px"><thead><tr>'
    +'<th style="text-align:left;padding:6px 10px 6px 0">Property</th>'
    +'<th style="text-align:right;padding:6px 10px">Cleaner hours</th>'
    +'<th style="text-align:right;padding:6px 10px">Guest-nights</th>'
    +'<th style="text-align:right;padding:6px 10px">Min/night</th>'
    +'<th style="text-align:left;padding:6px 10px">Trend</th>'
    +'</tr></thead><tbody>';
  propNames.forEach(pn=>{
    const hSum=Object.values(hoursByPropBucket[pn]||{}).reduce((a,b)=>a+b,0);
    const nSum=Object.values(nightsByPropBucket[pn]||{}).reduce((a,b)=>a+b,0);
    const mpn=nSum?Math.round(hSum*60/nSum):0;
    // Inline sparkline of min/night per bucket
    const sparkVals=allBuckets.map(bk=>{
      const h=(hoursByPropBucket[pn]||{})[bk]||0;
      const n=(nightsByPropBucket[pn]||{})[bk]||0;
      return n?(h*60/n):0;
    });
    const maxSpark=Math.max(1,...sparkVals);
    const sparkW=200,sparkH=24;
    const barW2=Math.max(2,sparkW/sparkVals.length-1);
    let spark='<svg width="'+sparkW+'" height="'+sparkH+'" xmlns="http://www.w3.org/2000/svg">';
    sparkVals.forEach((v,i)=>{
      if(v===0)return;
      const bh=(v/maxSpark)*sparkH;
      const x=i*(sparkW/sparkVals.length);
      spark+='<rect x="'+x+'" y="'+(sparkH-bh)+'" width="'+barW2+'" height="'+bh+'" fill="#5B8AC4" rx="1"><title>'+allBuckets[i]+': '+Math.round(v)+' min/night</title></rect>';
    });
    spark+='</svg>';
    // Color-code min/night — a rough flag, but warn for suspicious values
    let mpnStyle='';
    if(mpn>0){
      if(mpn>90)mpnStyle='color:var(--text-danger);font-weight:500';
      else if(mpn>60)mpnStyle='color:var(--text-warning);font-weight:500';
      else mpnStyle='color:var(--text-success);font-weight:500';
    }
    html+='<tr style="border-top:.5px solid var(--border-tertiary)">'
      +'<td style="padding:8px 10px 8px 0;font-weight:500">'+escapeHtml(pn)+'</td>'
      +'<td style="text-align:right;padding:8px 10px">'+Math.round(hSum*10)/10+'</td>'
      +'<td style="text-align:right;padding:8px 10px">'+nSum+'</td>'
      +'<td style="text-align:right;padding:8px 10px;'+mpnStyle+'">'+(nSum?mpn:'—')+'</td>'
      +'<td style="padding:4px 10px">'+spark+'</td>'
      +'</tr>';
  });
  html+='</tbody></table></div></div>';

  // --- Share of time spent on room cleaning (vs other work) ---
  // Standard: assume MIN_PER_TURNOVER minutes per room turnover. Compare expected vs actual.
  const MIN_PER_TURNOVER=30; // default assumption
  // Count turnovers (check-ins) in period per property
  const turnoversByProp={};
  let totalTurnovers=0;
  allBookings.forEach(b=>{
    if(!b.Check_In)return;
    const ci=new Date(b.Check_In);
    if(ci<fromDate||ci>toDate)return;
    const pn=b.Property_Name;if(!pn)return;
    turnoversByProp[pn]=(turnoversByProp[pn]||0)+1;
    totalTurnovers++;
  });
  const expectedRoomCleanHours=totalTurnovers*MIN_PER_TURNOVER/60;
  const shareRoomCleaning=totH>0?Math.round(expectedRoomCleanHours/totH*100):0;
  html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);margin-bottom:14px">'
    +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Share of cleaner time on room cleaning</div>'
    +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px">'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Check-ins (turnovers)</div><div style="font-size:18px;font-weight:500">'+totalTurnovers+'</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Expected room cleaning</div><div style="font-size:18px;font-weight:500">'+Math.round(expectedRoomCleanHours*10)/10+' h</div><div style="font-size:10px;color:var(--text-tertiary)">at '+MIN_PER_TURNOVER+' min/turnover</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Actual cleaner hours</div><div style="font-size:18px;font-weight:500">'+Math.round(totH*10)/10+' h</div></div>'
    +'<div><div style="font-size:11px;color:var(--text-tertiary)">Room-cleaning share</div><div style="font-size:24px;font-weight:500;color:'+(shareRoomCleaning>100?'var(--text-warning)':shareRoomCleaning<50?'var(--text-danger)':'var(--text-success)')+'">'+shareRoomCleaning+'%</div></div>'
    +'</div>'
    +'<div style="font-size:11px;color:var(--text-tertiary);margin-top:8px">Over 100% = cleaners worked faster than standard (or quality slipped). Under 50% = much time went to other work (maintenance, transport, deep cleaning).</div>'
    +'</div>';

  // --- Per-cleaner breakdown ---
  const cleanerUsers=allUsers.filter(u=>u.Role==='Cleaner'&&u.Epost);
  if(cleanerUsers.length){
    // For each cleaner: hours worked in period, and their own per-bucket data
    // Estimated room cleaning per cleaner: split total turnovers proportionally by hours
    html+='<div style="background:#fff;padding:12px;border-radius:8px;border:.5px solid var(--border-tertiary);margin-bottom:14px">'
      +'<div style="font-size:12px;font-weight:500;margin-bottom:10px;color:var(--text-secondary)">Per cleaner</div>'
      +'<div style="overflow-x:auto"><table style="font-size:12px;min-width:550px"><thead><tr>'
      +'<th style="text-align:left;padding:6px 10px 6px 0">Cleaner</th>'
      +'<th style="text-align:right;padding:6px 10px">Hours</th>'
      +'<th style="text-align:right;padding:6px 10px">Est. turnovers</th>'
      +'<th style="text-align:right;padding:6px 10px">Room-clean share</th>'
      +'<th style="text-align:left;padding:6px 10px">Hours trend ('+(efficiencyMode==='week'?'per week':'per month')+')</th>'
      +'</tr></thead><tbody>';
    // Compute per-cleaner hours per bucket
    const hoursByCleanerBucket={};
    allHours.forEach(h=>{
      if(!h.Date||!h.Time_From||!h.Time_To)return;
      const email=(h.Worker||'').toLowerCase();
      if(!cleanerEmails.has(email))return;
      const d=new Date(h.Date);
      if(d<fromDate||d>toDate)return;
      const hrs=calcHoursDiff(h.Time_From,h.Time_To);
      if(!hrs)return;
      const bk=bucketFn(d);
      if(!hoursByCleanerBucket[email])hoursByCleanerBucket[email]={};
      hoursByCleanerBucket[email][bk]=(hoursByCleanerBucket[email][bk]||0)+hrs;
    });
    cleanerUsers.forEach(u=>{
      const email=u.Epost.toLowerCase();
      const perBucket=hoursByCleanerBucket[email]||{};
      const cleanerHours=Object.values(perBucket).reduce((a,b)=>a+b,0);
      // Split turnovers proportionally to hours
      const estTurnovers=totH>0?Math.round(totalTurnovers*cleanerHours/totH):0;
      const expectedHours=estTurnovers*MIN_PER_TURNOVER/60;
      const shareCleaner=cleanerHours>0?Math.round(expectedHours/cleanerHours*100):0;
      // Sparkline of hours per bucket
      const sparkVals=allBuckets.map(bk=>perBucket[bk]||0);
      const maxSpark=Math.max(1,...sparkVals);
      const sparkW=220,sparkH=24;
      const barW2=Math.max(2,sparkW/Math.max(sparkVals.length,1)-1);
      let spark='<svg width="'+sparkW+'" height="'+sparkH+'" xmlns="http://www.w3.org/2000/svg">';
      sparkVals.forEach((v,i)=>{
        if(v===0)return;
        const bh=(v/maxSpark)*sparkH;
        const x=i*(sparkW/sparkVals.length);
        spark+='<rect x="'+x+'" y="'+(sparkH-bh)+'" width="'+barW2+'" height="'+bh+'" fill="#1D9E75" rx="1"><title>'+allBuckets[i]+': '+Math.round(v*10)/10+' h</title></rect>';
      });
      spark+='</svg>';
      let shareStyle='';
      if(cleanerHours>0){
        if(shareCleaner>100)shareStyle='color:var(--text-warning);font-weight:500';
        else if(shareCleaner<50)shareStyle='color:var(--text-danger);font-weight:500';
        else shareStyle='color:var(--text-success);font-weight:500';
      }
      html+='<tr style="border-top:.5px solid var(--border-tertiary)">'
        +'<td style="padding:8px 10px 8px 0;font-weight:500">'+escapeHtml(u.DisplayName||u.Epost)+'</td>'
        +'<td style="text-align:right;padding:8px 10px">'+Math.round(cleanerHours*10)/10+'</td>'
        +'<td style="text-align:right;padding:8px 10px">'+estTurnovers+'</td>'
        +'<td style="text-align:right;padding:8px 10px;'+shareStyle+'">'+(cleanerHours>0?shareCleaner+'%':'—')+'</td>'
        +'<td style="padding:4px 10px">'+spark+'</td>'
        +'</tr>';
    });
    html+='</tbody></table></div>'
      +'<div style="font-size:11px;color:var(--text-tertiary);margin-top:8px">Turnovers are split across cleaners proportionally to hours worked (we cannot tell who cleaned which room). "Room-clean share" uses '+MIN_PER_TURNOVER+' min/turnover as standard.</div>'
      +'</div>';
  }

  // --- Interpretation guide ---
  html+='<div style="background:#FAEEDA;border:1px solid #EF9F27;border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text-warning)">'
    +'<strong>⚠ How to interpret</strong><br>'
    +'Cleaner hours include breaks, transport between rigs, and small repairs — not just room cleaning. '
    +'Low guest-nights (e.g. empty month) will inflate "min/night" disproportionately. '
    +'Use trend sparklines to spot <strong>deviations within the same property over time</strong>, not to compare properties directly. '
    +'A sudden jump of 40%+ in one rig is worth a conversation — not a conclusion.'
    +'</div>';

  c.innerHTML=html;
}

// Helper: approximate date from ISO week
function _dateFromIsoWeek(year,week){
  const jan4=new Date(year,0,4);
  const day=jan4.getDay()||7;
  const mondayOfWeek1=new Date(year,0,4-day+1);
  return new Date(mondayOfWeek1.getFullYear(),mondayOfWeek1.getMonth(),mondayOfWeek1.getDate()+(week-1)*7);
}

// ============================================================
// MORE MENU (v14.0.10)
// ============================================================
function toggleMoreMenu(e){
  if(e){e.stopPropagation();e.preventDefault()}
  const m=document.getElementById('moreMenu');
  const wasOpen=m.style.display!=='none'&&m.style.display!=='';
  // Close first (reset state)
  m.style.display='none';
  if(!wasOpen){
    m.style.display='block';
    // Install outside-click closer AFTER this click completes
    requestAnimationFrame(()=>{document.addEventListener('mousedown',_moreMenuOutsideClick)});
  }
}
function _moreMenuOutsideClick(e){
  const wrap=document.getElementById('moreMenuWrap');
  if(wrap&&!wrap.contains(e.target)){
    closeMoreMenu();
  }
}
function closeMoreMenu(){
  const m=document.getElementById('moreMenu');
  if(m)m.style.display='none';
  document.removeEventListener('mousedown',_moreMenuOutsideClick);
}

// ============================================================
// FAKTURAGRUNNLAG / INVOICING (v14.0.10)
// ============================================================
let invoicingInitialized=false;

function toggleInvoicing(){
  ensureMainView();
  // Close other panels
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const panel=document.getElementById('invoicingPanel');
  panel.classList.toggle('open');
  const isOpen=panel.classList.contains('open');
  document.getElementById('mainView').classList.toggle('panel-mode',isOpen);
  if(isOpen){
    initInvoicingSelectors();
    renderInvoicing();
    // Highlight menu item
    const mi=document.getElementById('menuBtnInvoicing');if(mi)mi.classList.add('active-nav');
  }else{
    const mi=document.getElementById('menuBtnInvoicing');if(mi)mi.classList.remove('active-nav');
  }
  updateNavActiveState();
}

function initInvoicingSelectors(){
  if(invoicingInitialized)return;
  const now=new Date();
  const monthSel=document.getElementById('invMonth');const yearSel=document.getElementById('invYear');
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  monthSel.innerHTML=months.map((m,i)=>'<option value="'+i+'"'+(i===now.getMonth()?' selected':'')+'>'+m+'</option>').join('');
  const y=now.getFullYear();yearSel.innerHTML=[y-2,y-1,y,y+1].map(yr=>'<option value="'+yr+'"'+(yr===y?' selected':'')+'>'+yr+'</option>').join('');
  invoicingInitialized=true;
}

function clearInvoicingDateRange(){
  document.getElementById('invFrom').value='';document.getElementById('invTo').value='';
  renderInvoicing();
}

// Compute nights that fall WITHIN the given period (pro-rata for bookings that span across)
function _nightsInPeriod(booking,fromDate,toDate){
  if(!booking.Check_In)return 0;
  const ci=new Date(booking.Check_In);ci.setHours(0,0,0,0);
  const co=booking.Check_Out?new Date(booking.Check_Out):new Date();co.setHours(0,0,0,0);
  const start=ci>fromDate?ci:fromDate;
  const end=co<toDate?co:toDate;
  return Math.max(0,Math.round((end-start)/864e5));
}

function renderInvoicing(){
  const body=document.getElementById('invoicingBody');if(!body)return;
  const monthVal=document.getElementById('invMonth').value;
  const yearVal=document.getElementById('invYear').value;
  const fromVal=document.getElementById('invFrom').value;
  const toVal=document.getElementById('invTo').value;
  const groupBy=document.getElementById('invGroupBy').value;
  const useRange=!!(fromVal||toVal);
  let fromDate,toDate,periodLabel;
  if(useRange){
    fromDate=fromVal?new Date(fromVal+'T00:00:00'):new Date(1970,0,1);
    toDate=toVal?new Date(toVal+'T23:59:59'):new Date(2100,0,1);
    periodLabel=(fromVal?formatDate(fromVal):'…')+' → '+(toVal?formatDate(toVal):'…');
  }else{
    const m=parseInt(monthVal),y=parseInt(yearVal);
    fromDate=new Date(y,m,1);
    toDate=new Date(y,m+1,0,23,59,59);
    const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
    periodLabel=months[m]+' '+y;
  }

  document.getElementById('invoicingTitle').textContent='Fakturagrunnlag — '+periodLabel+' — '+(selectedProperty?selectedProperty.Title:'');

  // Filter bookings that overlap with the period, on current property
  const currentRoomIds=new Set(rooms.map(r=>r.id));
  // Identify which properties on this view have full-tenant lease active during the period
  const fullTenantByPropId={};
  const viewedProps=selectedProperty?[selectedProperty]:properties;
  viewedProps.forEach(p=>{
    const ft=computeFullTenantForPeriod(p,fromDate,toDate);
    if(ft)fullTenantByPropId[p.id]=ft;
  });
  // Room IDs that belong to a full-tenant property in this period
  const fullTenantRoomIds=new Set(
    allRooms.filter(r=>fullTenantByPropId[r.PropertyLookupId]).map(r=>r.id)
  );
  let fullTenantExcludedBookings={}; // propId -> count of individual bookings hidden

  // Identify rooms with active long-term contracts in this period
  const longTermByRoomId={};
  let longTermExcludedBookingCount=0;
  allRooms.forEach(r=>{
    if(!currentRoomIds.has(r.id))return;
    if(fullTenantRoomIds.has(r.id))return; // already covered by full-tenant
    const lt=computeLongTermForRoomPeriod(r,fromDate,toDate);
    if(lt)longTermByRoomId[r.id]=lt;
  });
  const longTermRoomIds=new Set(Object.keys(longTermByRoomId));

  const items=[];
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!currentRoomIds.has(rid))return;
    if(!b.Check_In)return;
    // Only bookings that have been active at any point during period
    if(b.Status==='Cancelled')return;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):new Date();co.setHours(0,0,0,0);
    // Check overlap
    if(co<fromDate||ci>toDate)return;
    // If room is in a full-tenant property, exclude from regular invoicing (tracked separately)
    if(fullTenantRoomIds.has(rid)){
      const room=allRooms.find(r=>r.id===rid);
      const pid=room?room.PropertyLookupId:'';
      fullTenantExcludedBookings[pid]=(fullTenantExcludedBookings[pid]||0)+1;
      return;
    }
    // If room has active long-term contract, exclude from regular invoicing
    if(longTermRoomIds.has(rid)){
      longTermExcludedBookingCount++;
      return;
    }
    const nights=_nightsInPeriod(b,fromDate,toDate);
    const cost=calcBookingCost(b,selectedProperty?selectedProperty.Title:'');
    const room=allRooms.find(r=>r.id===rid);
    const effectiveCo=getEffectiveCompany(b);
    const origCo=(b.Company||'').trim();
    const hasBillingOverride=(b.Billing_Company||'').trim()&&(b.Billing_Company||'').trim()!==origCo;
    if(nights>0){
      items.push({
        booking:b,
        room:room?room.Title:'?',
        name:b.Person_Name||'',
        company:effectiveCo,
        guestCompany:origCo,
        hasBillingOverride,
        nights,
        rate:cost.rate,
        total:nights*cost.rate,
        source:cost.source,
        nearMiss:cost.nearMiss,
        lineType:'nights'
      });
    }
    // CHECKOUT FEE: only for Completed bookings where Check_Out falls within period
    // and Include_Checkout_Fee is not explicitly false
    // Skip if this company has a Percent-based fee (handled per-company below)
    // Skip if Continuation=true (mid-stay room change — only one utvask per logical stay)
    const isContinuation=(b.Continuation===true||b.Continuation==='true'||b.Continuation===1);
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuation&&!hasPercentFee(effectiveCo,selectedProperty?selectedProperty.Title:'')){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,selectedProperty?selectedProperty.Title:'');
        if(fee>0){
          items.push({
            booking:b,
            room:room?room.Title:'?',
            name:b.Person_Name||'',
            company:effectiveCo,
            guestCompany:origCo,
            hasBillingOverride,
            nights:1,
            rate:fee,
            total:fee,
            source:'Checkout fee',
            nearMiss:null,
            lineType:'checkout',
            checkoutDate:b.Check_Out
          });
        }
      }
    }
  });

  // FULL-TENANT LEASE: add one line per property with active full-tenant agreement
  Object.keys(fullTenantByPropId).forEach(pid=>{
    const ft=fullTenantByPropId[pid];
    const prop=properties.find(p=>String(p.id)===String(pid));
    const propName=prop?prop.Title:'';
    const excluded=fullTenantExcludedBookings[pid]||0;
    items.push({
      booking:{id:''},
      room:propName,
      name:ft.company+' (full-tenant lease)',
      company:ft.company,
      guestCompany:'',
      hasBillingOverride:false,
      nights:ft.days,
      rate:ft.rate*ft.rooms,
      total:ft.total,
      source:ft.detailLabel+(excluded?' · '+excluded+' individual bookings hidden':''),
      nearMiss:null,
      lineType:'fulltenant',
      checkoutDate:null
    });
  });

  // LONG-TERM CONTRACTS (per-room): one line per room with active contract
  Object.keys(longTermByRoomId).forEach(rid=>{
    const lt=longTermByRoomId[rid];
    items.push({
      booking:{id:''},
      room:lt.room.Title||'',
      name:lt.company+' — '+(lt.room.Title||''),
      company:lt.company,
      guestCompany:'',
      hasBillingOverride:false,
      nights:lt.days,
      rate:lt.price,
      total:lt.total,
      source:lt.detailLabel,
      nearMiss:null,
      lineType:'longterm',
      checkoutDate:null
    });
  });

  // PERCENT-BASED FEES: compute per-company total of nights × rate and add a % fee line
  const propTitleForPercent=selectedProperty?selectedProperty.Title:'';
  const companyNightSum={};
  items.filter(i=>i.lineType==='nights').forEach(i=>{
    const c=(i.company||'').trim();
    if(!c)return;
    if(!companyNightSum[c])companyNightSum[c]=0;
    companyNightSum[c]+=i.total;
  });
  Object.keys(companyNightSum).forEach(c=>{
    const pct=getPercentFeeRate(c,propTitleForPercent);
    if(pct>0){
      const feeAmount=Math.round(companyNightSum[c]*pct);
      // Attach to a synthetic booking-like object so row-click still works with first company booking
      const firstBooking=items.find(i=>i.company===c);
      items.push({
        booking:firstBooking?firstBooking.booking:{id:''},
        room:'—',
        name:c+' (monthly fee)',
        company:c,
        nights:1,
        rate:feeAmount,
        total:feeAmount,
        source:(pct*100)+'% of '+companyNightSum[c].toLocaleString('nb-NO')+' kr',
        nearMiss:null,
        lineType:'percent',
        checkoutDate:null
      });
    }
  });

  // Apply company filter (if set)
  const companyFilter=document.getElementById('invCompanyFilter').value||'__ALL__';
  const filteredItems=companyFilter==='__ALL__'?items:items.filter(i=>i.company===companyFilter);

  // Populate company filter dropdown (preserve current selection)
  const allCompanies=[...new Set(items.map(i=>i.company||'(no company)'))].sort();
  const cfSel=document.getElementById('invCompanyFilter');
  const prevVal=cfSel.value;
  cfSel.innerHTML='<option value="__ALL__">All companies</option>'+allCompanies.map(c=>'<option value="'+c+'">'+c+'</option>').join('');
  if(prevVal&&[...cfSel.options].some(o=>o.value===prevVal))cfSel.value=prevVal;
  const finalItems=(cfSel.value==='__ALL__')?items:items.filter(i=>i.company===cfSel.value);

  // Warnings for missing rates
  const missingRate=finalItems.filter(i=>!i.rate&&i.lineType==='nights');
  const warnings=missingRate.length?'<div style="margin-bottom:12px;padding:8px 12px;background:var(--bg-warning);border:1px solid #EF9F27;border-radius:6px;font-size:12px;color:var(--text-warning)">⚠ '+missingRate.length+' booking'+(missingRate.length!==1?'s':'')+' without rates — these are not included in totals. Check rate configuration.</div>':'';

  if(!finalItems.length){
    body.innerHTML='<div style="text-align:center;padding:40px;color:var(--text-secondary)">No bookings in this period'+(cfSel.value!=='__ALL__'?' for '+escapeHtml(cfSel.value):'')+' on '+(selectedProperty?selectedProperty.Title:'selected property')+'.</div>';
    return;
  }

  let html=warnings;

  // Grand totals — separate nights from checkout/percent fees and full-tenant leases
  const nightItems=finalItems.filter(i=>i.lineType==='nights');
  const feeItems=finalItems.filter(i=>i.lineType==='checkout'||i.lineType==='percent'||i.lineType==='fulltenant'||i.lineType==='longterm');
  const totalNights=nightItems.reduce((a,i)=>a+i.nights,0);
  const nightRevenue=nightItems.reduce((a,i)=>a+i.total,0);
  const feeRevenue=feeItems.reduce((a,i)=>a+i.total,0);
  const totalRevenue=nightRevenue+feeRevenue;
  const totalBookings=new Set(nightItems.map(i=>i.booking.id)).size;
  html+='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px">'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Bookings</div><div style="font-size:20px;font-weight:500">'+totalBookings+'</div></div>'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Guest-nights</div><div style="font-size:20px;font-weight:500">'+totalNights+'</div></div>'
    +'<div style="background:#fff;padding:10px;border-radius:8px;border:.5px solid var(--border-tertiary)"><div style="font-size:11px;color:var(--text-tertiary)">Total revenue</div><div style="font-size:20px;font-weight:500;color:var(--text-success)">'+totalRevenue.toLocaleString('nb-NO')+' kr</div>'+(feeRevenue>0?'<div style="font-size:10px;color:var(--text-tertiary);margin-top:2px">Nights: '+nightRevenue.toLocaleString('nb-NO')+' kr + Utvask: '+feeRevenue.toLocaleString('nb-NO')+' kr</div>':'')+'</div>'
    +'</div>';

  // Grouped rendering
  if(groupBy==='none'){
    html+=_renderInvoicingFlat(finalItems);
  }else{
    const keyFn=groupBy==='company'?(i=>i.company||'(no company)'):(i=>i.name||'(no name)');
    const groups={};
    finalItems.forEach(i=>{const k=keyFn(i);if(!groups[k])groups[k]=[];groups[k].push(i)});
    const sortedKeys=Object.keys(groups).sort((a,b)=>{
      const tA=groups[a].reduce((s,i)=>s+i.total,0);
      const tB=groups[b].reduce((s,i)=>s+i.total,0);
      return tB-tA;
    });
    html+='<div style="background:#fff;border-radius:8px;border:.5px solid var(--border-tertiary);overflow:hidden">';
    sortedKeys.forEach(k=>{
      const grp=groups[k];
      const gNightsItems=grp.filter(i=>i.lineType==='nights');
      const gFees=grp.filter(i=>i.lineType==='checkout');
      const gPercent=grp.filter(i=>i.lineType==='percent');
      const gNights=gNightsItems.reduce((s,i)=>s+i.nights,0);
      const gTotal=grp.reduce((s,i)=>s+i.total,0);
      const gBookings=new Set(gNightsItems.map(i=>i.booking.id)).size;
      const feeSuffix=(gFees.length?' · '+gFees.length+' utvask':'')+(gPercent.length?' · % fee':'');
      const exportBtn=(groupBy==='company'&&k!=='(no company)')
        ?'<button onclick="event.stopPropagation();exportInvoicingCSV(\''+k.replace(/'/g,"\\'")+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit;margin-left:10px" title="Export CSV for '+escapeHtml(k)+'">↓ CSV</button>'
         +'<button onclick="event.stopPropagation();exportInvoicingPDF(\''+k.replace(/'/g,"\\'")+'\')" style="padding:3px 10px;border:1px solid #1D9E75;border-radius:4px;background:rgba(29,158,117,.1);color:#1D9E75;cursor:pointer;font-size:11px;font-family:inherit;margin-left:4px" title="Export PDF for '+escapeHtml(k)+'">📄 PDF</button>'
        :'';
      html+='<div style="padding:10px 14px;background:var(--bg-secondary);border-bottom:1px solid var(--border-tertiary);display:flex;justify-content:space-between;align-items:center">'
        +'<div><strong>'+escapeHtml(k)+'</strong> <span class="muted" style="font-size:11px;margin-left:8px">'+gBookings+' booking'+(gBookings!==1?'s':'')+feeSuffix+'</span></div>'
        +'<div style="font-size:13px;display:flex;align-items:center"><strong>'+gNights+'</strong>&nbsp;nights · <strong style="color:var(--text-success);margin-left:4px">'+gTotal.toLocaleString('nb-NO')+' kr</strong>'+exportBtn+'</div>'
        +'</div>';
      html+='<table style="width:100%;font-size:12px"><thead><tr style="background:var(--bg-tertiary)"><th style="padding:6px 10px;text-align:left">Guest</th><th style="padding:6px 10px;text-align:left">Company</th><th style="padding:6px 10px;text-align:left">Room</th><th style="padding:6px 10px;text-align:left">Period</th><th style="padding:6px 10px;text-align:right">Nights</th><th style="padding:6px 10px;text-align:right">Rate</th><th style="padding:6px 10px;text-align:right">Total</th><th style="padding:6px 10px;text-align:left">Rate source</th></tr></thead><tbody>';
      grp.forEach(i=>{
        const isCheckout=i.lineType==='checkout';
        const isPercent=i.lineType==='percent';
        const isFullTenant=i.lineType==='fulltenant';
        const isLongTerm=i.lineType==='longterm';
        const ci=i.booking.Check_In?formatDate(i.booking.Check_In):'';
        const co=i.booking.Check_Out?formatDate(i.booking.Check_Out):'Open';
        let period;
        if(isFullTenant)period='🔒 Full-tenant lease';
        else if(isLongTerm)period='🔑 Långtidsleie';
        else if(isPercent)period='📊 Monthly percent fee';
        else if(isCheckout)period='🧹 Checkout '+formatDate(i.checkoutDate);
        else period=ci+' → '+co;
        const nightsCell=(isCheckout||isPercent)?'—':((isFullTenant||isLongTerm)?i.nights+' days':i.nights);
        let rateCell;
        if(isFullTenant)rateCell='<em style="color:var(--text-tertiary)">Full-tenant</em>';
        else if(isLongTerm)rateCell='<em style="color:var(--text-tertiary)">Långtid</em>';
        else if(isPercent)rateCell='<em style="color:var(--text-tertiary)">%-basert</em>';
        else if(isCheckout)rateCell='<em style="color:var(--text-tertiary)">Utvask</em>';
        else rateCell=(i.rate?i.rate.toLocaleString('nb-NO')+' kr':'<span style="color:var(--text-danger)">— missing</span>');
        const totalCell=i.total?i.total.toLocaleString('nb-NO')+' kr':'—';
        let sourceCell;
        if(isFullTenant)sourceCell='<span style="color:#1D9E75">🔒 '+escapeHtml(i.source)+'</span>';
        else if(isLongTerm)sourceCell='<span style="color:#0EA5A5">🔑 '+escapeHtml(i.source)+'</span>';
        else if(isPercent)sourceCell='<span style="color:#EF9F27">📊 '+escapeHtml(i.source)+'</span>';
        else if(isCheckout)sourceCell='<span style="color:#7B61FF">🧹 Checkout fee</span>';
        else sourceCell=(i.nearMiss?'<span title="'+escapeHtml(i.nearMiss)+'" style="color:var(--text-warning)">⚠ '+escapeHtml(i.source)+'</span>':escapeHtml(i.source));
        let rowStyle;
        if(isFullTenant)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(29,158,117,.08)';
        else if(isLongTerm)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(14,165,165,.07)';
        else if(isPercent)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(239,159,39,.06)';
        else if(isCheckout)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:pointer;background:rgba(123,97,255,.04)';
        else rowStyle='border-top:.5px solid var(--border-tertiary);cursor:pointer';
        const hoverBg=isFullTenant?'rgba(29,158,117,.16)':(isLongTerm?'rgba(14,165,165,.14)':(isPercent?'rgba(239,159,39,.12)':(isCheckout?'rgba(123,97,255,.12)':'var(--bg-secondary)')));
        const restBg=isFullTenant?'rgba(29,158,117,.08)':(isLongTerm?'rgba(14,165,165,.07)':(isPercent?'rgba(239,159,39,.06)':(isCheckout?'rgba(123,97,255,.04)':'')));
        // Flag company mismatch when grouping by company: if company field differs from group key, highlight
        const groupKey=k;
        const actualCompany=i.company||'(no company)';
        const companyMismatch=groupBy==='company'&&actualCompany!==groupKey;
        let companyCell;
        if(companyMismatch){
          companyCell='<span style="color:var(--text-danger);font-weight:500" title="Mismatch with group">⚠ '+escapeHtml(actualCompany)+'</span>';
        }else if(i.hasBillingOverride&&i.guestCompany){
          companyCell=escapeHtml(actualCompany)+' <span style="color:var(--text-tertiary);font-size:11px" title="Guest works for '+escapeHtml(i.guestCompany)+'">← '+escapeHtml(i.guestCompany)+'</span>';
        }else{
          companyCell=escapeHtml(actualCompany);
        }
        let nameCell;
        if(isFullTenant)nameCell='<span style="color:#1D9E75;font-weight:500">'+escapeHtml(i.name)+'</span>';
        else if(isLongTerm)nameCell='<span style="color:#0EA5A5;font-weight:500">🔑 '+escapeHtml(i.name)+'</span>';
        else if(isPercent)nameCell='<span style="color:var(--text-warning);font-weight:500">'+escapeHtml(i.name)+'</span>';
        else if(isCheckout)nameCell='<span style="color:var(--text-tertiary)">↳ '+guestMarkedName(i.name)+'</span>';
        else nameCell=guestMarkedName(i.name);
        // Full-tenant, long-term and percent rows are not clickable
        const clickAttr=(isPercent||isFullTenant||isLongTerm)?'':'onclick="openEditBooking(\''+i.booking.id+'\')"';
        html+='<tr '+clickAttr+' style="'+rowStyle+'" onmouseover="this.style.background=\''+hoverBg+'\'" onmouseout="this.style.background=\''+restBg+'\'">'
          +'<td style="padding:6px 10px">'+nameCell+'</td>'
          +'<td style="padding:6px 10px">'+companyCell+'</td>'
          +'<td style="padding:6px 10px;font-weight:500">'+escapeHtml(i.room)+'</td>'
          +'<td style="padding:6px 10px">'+period+'</td>'
          +'<td style="padding:6px 10px;text-align:right">'+nightsCell+'</td>'
          +'<td style="padding:6px 10px;text-align:right">'+rateCell+'</td>'
          +'<td style="padding:6px 10px;text-align:right;font-weight:500">'+totalCell+'</td>'
          +'<td style="padding:6px 10px;font-size:11px;color:var(--text-tertiary)">'+sourceCell+'</td>'
          +'</tr>';
      });
      html+='</tbody></table>';
    });
    html+='</div>';
  }

  body.innerHTML=html;
}

function _renderInvoicingFlat(items){
  let html='<div style="background:#fff;border-radius:8px;border:.5px solid var(--border-tertiary);overflow:hidden">';
  html+='<table style="width:100%;font-size:12px"><thead><tr style="background:var(--bg-secondary)">'
    +'<th style="padding:8px 10px;text-align:left">Guest</th>'
    +'<th style="padding:8px 10px;text-align:left">Company</th>'
    +'<th style="padding:8px 10px;text-align:left">Room</th>'
    +'<th style="padding:8px 10px;text-align:left">Period</th>'
    +'<th style="padding:8px 10px;text-align:right">Nights</th>'
    +'<th style="padding:8px 10px;text-align:right">Rate</th>'
    +'<th style="padding:8px 10px;text-align:right">Total</th>'
    +'</tr></thead><tbody>';
  items.sort((a,b)=>new Date(a.booking.Check_In)-new Date(b.booking.Check_In));
  items.forEach(i=>{
    const ci=formatDate(i.booking.Check_In);
    const co=i.booking.Check_Out?formatDate(i.booking.Check_Out):'Open';
    const rateCell=i.rate?i.rate.toLocaleString('nb-NO')+' kr':'<span style="color:var(--text-danger)">— missing</span>';
    const totalCell=i.total?i.total.toLocaleString('nb-NO')+' kr':'—';
    html+='<tr onclick="openEditBooking(\''+i.booking.id+'\')" style="border-top:.5px solid var(--border-tertiary);cursor:pointer" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
      +'<td style="padding:6px 10px">'+guestMarkedName(i.name)+'</td>'
      +'<td style="padding:6px 10px">'+escapeHtml(i.company)+'</td>'
      +'<td style="padding:6px 10px;font-weight:500">'+escapeHtml(i.room)+'</td>'
      +'<td style="padding:6px 10px">'+ci+' → '+co+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+i.nights+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+rateCell+'</td>'
      +'<td style="padding:6px 10px;text-align:right;font-weight:500">'+totalCell+'</td>'
      +'</tr>';
  });
  html+='</tbody></table></div>';
  return html;
}

function exportInvoicingCSV(companyFilterName){
  const monthVal=document.getElementById('invMonth').value;
  const yearVal=document.getElementById('invYear').value;
  const fromVal=document.getElementById('invFrom').value;
  const toVal=document.getElementById('invTo').value;
  const useRange=!!(fromVal||toVal);
  let fromDate,toDate,periodStr;
  if(useRange){
    fromDate=fromVal?new Date(fromVal+'T00:00:00'):new Date(1970,0,1);
    toDate=toVal?new Date(toVal+'T23:59:59'):new Date(2100,0,1);
    periodStr=(fromVal||'start')+'_to_'+(toVal||'end');
  }else{
    const m=parseInt(monthVal),y=parseInt(yearVal);
    fromDate=new Date(y,m,1);
    toDate=new Date(y,m+1,0,23,59,59);
    periodStr=y+'_'+String(m+1).padStart(2,'0');
  }
  // If not given explicitly, fall back to current dropdown filter (or __ALL__)
  if(companyFilterName===undefined){
    const cf=document.getElementById('invCompanyFilter');
    companyFilterName=cf&&cf.value!=='__ALL__'?cf.value:null;
  }
  const currentRoomIds=new Set(rooms.map(r=>r.id));
  const rows=[];
  const companyNightSum={};
  const propTitleForPercent=selectedProperty?selectedProperty.Title:'';
  // Identify full-tenant properties
  const fullTenantByPropId={};
  const viewedProps=selectedProperty?[selectedProperty]:properties;
  viewedProps.forEach(p=>{
    const ft=computeFullTenantForPeriod(p,fromDate,toDate);
    if(ft)fullTenantByPropId[p.id]=ft;
  });
  const fullTenantRoomIds=new Set(
    allRooms.filter(r=>fullTenantByPropId[r.PropertyLookupId]).map(r=>r.id)
  );
  // LongTerm contracts per room
  const longTermByRoomIdCsv={};
  allRooms.forEach(r=>{
    if(!currentRoomIds.has(r.id))return;
    if(fullTenantRoomIds.has(r.id))return;
    const lt=computeLongTermForRoomPeriod(r,fromDate,toDate);
    if(lt)longTermByRoomIdCsv[r.id]=lt;
  });
  const longTermRoomIdsCsv=new Set(Object.keys(longTermByRoomIdCsv));
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!currentRoomIds.has(rid))return;
    if(!b.Check_In)return;
    if(b.Status==='Cancelled')return;
    // Skip bookings on full-tenant rooms (covered by lease)
    if(fullTenantRoomIds.has(rid))return;
    // Skip bookings on long-term contract rooms
    if(longTermRoomIdsCsv.has(rid))return;
    const effectiveCo=getEffectiveCompany(b);
    // Apply company filter — filter against effective (billing) company
    if(companyFilterName&&effectiveCo!==companyFilterName)return;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):new Date();co.setHours(0,0,0,0);
    if(co<fromDate||ci>toDate)return;
    const nights=_nightsInPeriod(b,fromDate,toDate);
    const cost=calcBookingCost(b,selectedProperty?selectedProperty.Title:'');
    const room=allRooms.find(r=>r.id===rid);
    const origCo=(b.Company||'').trim();
    const billingCo=effectiveCo!==origCo?effectiveCo:'';
    if(nights>0){
      rows.push([
        b.Person_Name||'',
        origCo,
        billingCo,
        room?room.Title:'',
        formatDate(b.Check_In),
        b.Check_Out?formatDate(b.Check_Out):'Open',
        nights,
        cost.rate||0,
        nights*(cost.rate||0),
        cost.source||''
      ]);
      // Track for percent calculation by effective company
      if(effectiveCo){companyNightSum[effectiveCo]=(companyNightSum[effectiveCo]||0)+nights*(cost.rate||0)}
    }
    // Checkout fee line (skip if company has Percent fee, skip if Continuation)
    const isContinuationExp=(b.Continuation===true||b.Continuation==='true'||b.Continuation===1);
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuationExp&&!hasPercentFee(effectiveCo,propTitleForPercent)){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,propTitleForPercent);
        if(fee>0){
          rows.push([
            b.Person_Name||'',
            origCo,
            billingCo,
            room?room.Title:'',
            'Checkout '+formatDate(b.Check_Out),
            '',
            0,
            fee,
            fee,
            'Utvask'
          ]);
        }
      }
    }
  });
  // Full-tenant lease lines
  Object.keys(fullTenantByPropId).forEach(pid=>{
    const ft=fullTenantByPropId[pid];
    // Apply company filter if set
    if(companyFilterName&&ft.company!==companyFilterName)return;
    const prop=properties.find(p=>String(p.id)===String(pid));
    rows.push([
      ft.company+' (full-tenant lease)',
      '',                           // guest company
      ft.company,                   // billing company
      prop?prop.Title:'',           // room column = property
      '',                           // check-in
      '',                           // check-out
      ft.days,                      // days in period
      ft.rate,                      // rate (per day or per month — see source label)
      ft.total,                     // total
      ft.detailLabel
    ]);
  });

  // Long-term per-room contracts
  Object.keys(longTermByRoomIdCsv).forEach(rid=>{
    const lt=longTermByRoomIdCsv[rid];
    if(companyFilterName&&lt.company!==companyFilterName)return;
    rows.push([
      lt.company+' — '+(lt.room.Title||''),
      '',
      lt.company,
      lt.room.Title||'',
      '',
      '',
      lt.days,
      lt.price,
      lt.total,
      'Långtid: '+lt.detailLabel
    ]);
  });

  // Percent-based fee lines (by effective/billing company)
  Object.keys(companyNightSum).forEach(c=>{
    const pct=getPercentFeeRate(c,propTitleForPercent);
    if(pct>0){
      const feeAmount=Math.round(companyNightSum[c]*pct);
      rows.push([
        c+' (monthly fee)',
        '',            // guest company
        c,             // billing company
        '',            // room
        periodStr,
        '',
        0,
        feeAmount,
        feeAmount,
        (pct*100)+'% of '+companyNightSum[c]+' kr'
      ]);
    }
  });
  // Sort by billing company, then guest name
  rows.sort((a,b)=>((a[2]||a[1])+'').localeCompare(((b[2]||b[1])+''),'nb')||(a[0]+'').localeCompare((b[0]+''),'nb'));
  // Totals (nights col moved from [5] to [6], total col from [7] to [8])
  const totalN=rows.reduce((s,r)=>s+(typeof r[6]==='number'?r[6]:0),0);
  const totalT=rows.reduce((s,r)=>s+(typeof r[8]==='number'?r[8]:0),0);
  rows.push(['','','','','','Total',totalN,'',totalT,'']);
  const headers=['Guest','Company (guest)','Billing company','Room','Check-in','Check-out','Nights','Rate','Total','Rate source'];
  const propName=(selectedProperty?selectedProperty.Title:'').replace(/\s+/g,'_');
  const companyPart=companyFilterName?'_'+companyFilterName.replace(/\s+/g,'_'):'';
  downloadCSV('Fakturagrunnlag_'+propName+companyPart+'_'+periodStr,headers,rows);
}

// ============================================================
// ADD GUEST FROM BOOKING (v14.0.10)
// ============================================================
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
// GUEST BOOKINGS HISTORY (v14.0.10)
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
// HOURS IMPORT (v14.0.10)
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
      importHoursData.push({date,worker,location,timeFrom:tFrom,timeTo:tTo,notes,hrs,error,workerName:userMatch?userMatch.DisplayName:worker});
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

// ============================================================
// CLEANING DIAGNOSTICS (v14.0.10)
// ============================================================
function showCleaningDiagnostics(){
  const today=new Date();today.setHours(0,0,0,0);
  const todayStr=formatDate(today);
  const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const dayName=days[today.getDay()];

  // Group all bookings by property and analyze
  const propGroups={};
  allBookings.forEach(b=>{
    const propName=b.Property_Name||'(unknown)';
    if(!propGroups[propName])propGroups[propName]={active:0,upcoming:0,completed:0,cancelled:0,noStatus:0,washToday:0,noCheckIn:0,noRoomLink:0,bookings:[]};
    const g=propGroups[propName];
    if(b.Status==='Active')g.active++;
    else if(b.Status==='Upcoming')g.upcoming++;
    else if(b.Status==='Completed')g.completed++;
    else if(b.Status==='Cancelled')g.cancelled++;
    else g.noStatus++;
    if(!b.Check_In)g.noCheckIn++;
    if(!b.RoomLookupId)g.noRoomLink++;
    if(b.Status==='Active'&&b.Check_In){
      const w=calcWashDates(b.Check_In,b.Check_Out);
      if(w.some(x=>x.isToday)){g.washToday++;g.bookings.push(b)}
    }
  });

  let html='<div style="background:var(--bg-warning);padding:10px 12px;border-radius:6px;margin-bottom:14px;color:var(--text-warning);font-size:12px">'
    +'<strong>Today is '+todayStr+' ('+dayName+').</strong> This shows what the system thinks needs cleaning today, and why bookings might be excluded.</div>';

  // Per-property summary
  html+='<table style="width:100%;font-size:12px;margin-bottom:18px"><thead><tr style="background:var(--bg-secondary)">'
    +'<th style="padding:8px 10px;text-align:left">Property</th>'
    +'<th style="padding:8px 10px;text-align:right">Active</th>'
    +'<th style="padding:8px 10px;text-align:right">Upcoming</th>'
    +'<th style="padding:8px 10px;text-align:right">Completed</th>'
    +'<th style="padding:8px 10px;text-align:right">No status</th>'
    +'<th style="padding:8px 10px;text-align:right">No Check-in</th>'
    +'<th style="padding:8px 10px;text-align:right">No room link</th>'
    +'<th style="padding:8px 10px;text-align:right;color:var(--text-success)">Wash today</th>'
    +'</tr></thead><tbody>';
  Object.keys(propGroups).sort().forEach(pn=>{
    const g=propGroups[pn];
    html+='<tr style="border-top:.5px solid var(--border-tertiary)">'
      +'<td style="padding:6px 10px;font-weight:500">'+escapeHtml(pn)+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+g.active+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+g.upcoming+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+g.completed+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+(g.noStatus?'<span style="color:var(--text-danger)">'+g.noStatus+'</span>':'0')+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+(g.noCheckIn?'<span style="color:var(--text-danger)">'+g.noCheckIn+'</span>':'0')+'</td>'
      +'<td style="padding:6px 10px;text-align:right">'+(g.noRoomLink?'<span style="color:var(--text-danger)">'+g.noRoomLink+'</span>':'0')+'</td>'
      +'<td style="padding:6px 10px;text-align:right;font-weight:500;color:var(--text-success)">'+g.washToday+'</td>'
      +'</tr>';
  });
  html+='</tbody></table>';

  // For current property: detailed analysis of every Active booking
  if(selectedProperty){
    const propName=selectedProperty.Title;
    const currentRoomIds=new Set(rooms.map(r=>r.id));
    const activeOnCurrentProp=allBookings.filter(b=>b.Status==='Active'&&currentRoomIds.has(String(b.RoomLookupId||'')));
    html+='<h3 style="margin:0 0 8px;font-size:14px">Active bookings on '+escapeHtml(propName)+' — detailed wash analysis</h3>';
    if(!activeOnCurrentProp.length){
      html+='<p style="color:var(--text-secondary)">No active bookings on this property.</p>';
    }else{
      html+='<table style="width:100%;font-size:11px"><thead><tr style="background:var(--bg-secondary)">'
        +'<th style="padding:6px 8px;text-align:left">Room</th>'
        +'<th style="padding:6px 8px;text-align:left">Guest</th>'
        +'<th style="padding:6px 8px;text-align:left">Check-in</th>'
        +'<th style="padding:6px 8px;text-align:right">Days since</th>'
        +'<th style="padding:6px 8px;text-align:left">Last wash</th>'
        +'<th style="padding:6px 8px;text-align:left">Next wash</th>'
        +'<th style="padding:6px 8px;text-align:left">Status today</th>'
        +'</tr></thead><tbody>';
      activeOnCurrentProp.sort((a,b)=>{
        const ra=allRooms.find(r=>r.id===String(a.RoomLookupId));
        const rb=allRooms.find(r=>r.id===String(b.RoomLookupId));
        return (ra?ra.Title:'').localeCompare(rb?rb.Title:'',undefined,{numeric:true});
      }).forEach(b=>{
        const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
        const roomTitle=room?room.Title:'(no room)';
        const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
        const daysSince=Math.round((today-ci)/864e5);
        const washes=calcWashDates(b.Check_In,b.Check_Out);
        const past=washes.filter(w=>w.isPast);
        const lastWash=past.length?past[past.length-1]:null;
        const todayWash=washes.find(w=>w.isToday);
        const nextWash=washes.find(w=>!w.isPast&&!w.isToday);
        let statusToday='<span style="color:var(--text-tertiary)">—</span>';
        if(todayWash)statusToday='<span style="color:var(--text-success);font-weight:500">✓ Wash today ('+todayWash.type+')</span>';
        else if(b.Cleaning_Status==='Dirty')statusToday='<span style="color:var(--text-warning)">⚠ Marked Dirty</span>';
        const nextStr=nextWash?formatDate(nextWash.date)+' ('+days[nextWash.date.getDay()]+') — '+nextWash.type:'<span class="muted">none</span>';
        const lastStr=lastWash?formatDate(lastWash.date)+' — '+lastWash.type:'<span class="muted">none yet</span>';
        html+='<tr style="border-top:.5px solid var(--border-tertiary)">'
          +'<td style="padding:5px 8px;font-weight:500">'+escapeHtml(roomTitle)+'</td>'
          +'<td style="padding:5px 8px">'+escapeHtml(b.Person_Name||'')+'</td>'
          +'<td style="padding:5px 8px">'+formatDate(b.Check_In)+' ('+days[ci.getDay()]+')</td>'
          +'<td style="padding:5px 8px;text-align:right">'+daysSince+'</td>'
          +'<td style="padding:5px 8px">'+lastStr+'</td>'
          +'<td style="padding:5px 8px">'+nextStr+'</td>'
          +'<td style="padding:5px 8px">'+statusToday+'</td>'
          +'</tr>';
      });
      html+='</tbody></table>';
    }
  }

  document.getElementById('diagBody').innerHTML=html;
  document.getElementById('diagModal').classList.add('open');
}

// ============================================================
// BATTERY REFRESH (v14.0.10)
// ============================================================
const BATTERY_FILE_PATH='Batteristatus/RoomBattery.csv';

async function refreshBatteryStatus(){
  const btn=document.querySelector('[data-battery-refresh-btn]');
  if(btn){btn.disabled=true;btn.textContent='⏳ Loading CSV...'}
  try{
    // 1. Fetch CSV content
    const csvText=await fetchSiteFileText(BATTERY_FILE_PATH);
    if(!csvText||!csvText.trim())throw new Error('CSV file is empty');

    // 2. Parse CSV — accept comma, semicolon, or tab as separator
    const lines=csvText.split(/\r?\n/).filter(l=>l.trim());
    if(lines.length<2)throw new Error('CSV has no data rows (only header or blank)');
    const firstLine=lines[0];
    // Detect separator: priority semicolon, then tab, then comma
    let sep=';';
    if(firstLine.indexOf(';')>=0)sep=';';
    else if(firstLine.indexOf('\t')>=0)sep='\t';
    else if(firstLine.indexOf(',')>=0)sep=',';
    else throw new Error('Could not detect CSV separator');

    const headers=lines[0].split(sep).map(h=>h.trim().toLowerCase().replace(/['"]/g,''));
    // Find room column (RomNr/Room/Rom) and battery column (Verdi/Battery/%)
    const colRoom=headers.findIndex(h=>h==='romnr'||h==='room'||h==='rom'||h==='roomnr');
    const colBat=headers.findIndex(h=>h==='verdi'||h==='battery'||h==='battery_level'||h==='%'||h==='batteri');
    if(colRoom===-1||colBat===-1){
      throw new Error('CSV must have columns RomNr and Verdi (or Room/Battery). Found: '+headers.join(', '));
    }

    // 3. Parse each row into {roomTitle, batteryPct}
    const entries=[];
    for(let i=1;i<lines.length;i++){
      const cols=lines[i].split(sep).map(c=>c.trim().replace(/^['"]|['"]$/g,''));
      const roomTitle=cols[colRoom]||'';
      let batRaw=cols[colBat]||'';
      // Handle Norwegian decimal: "92,5" → 92.5, but we expect integer percent
      batRaw=batRaw.replace(',','.');
      const bat=parseFloat(batRaw);
      if(!roomTitle||isNaN(bat))continue;
      entries.push({roomTitle,bat:Math.round(bat)});
    }
    if(!entries.length)throw new Error('No valid rows parsed from CSV');

    // 4. Match against allRooms (by Title, exact string match — trimmed)
    if(btn)btn.textContent='⏳ Updating rooms ('+entries.length+')...';
    let updated=0,skipped=0,unchanged=0;
    const notFound=[];
    for(let i=0;i<entries.length;i++){
      const e=entries[i];
      const room=allRooms.find(r=>(r.Title||'').toString().trim()===e.roomTitle.trim());
      if(!room){notFound.push(e.roomTitle);skipped++;continue}
      // Skip if value hasn't changed
      if(Number(room.Door_Battery_Level)===e.bat){unchanged++;continue}
      try{
        await updateListItem('Rooms',room.id,{Door_Battery_Level:e.bat});
        room.Door_Battery_Level=e.bat;
        updated++;
      }catch(err){console.error('Failed to update room '+e.roomTitle+':',err);skipped++}
      // Throttle every 10 to avoid rate limiting
      if(i%10===9)await new Promise(res=>setTimeout(res,300));
    }

    // 5. Show summary + re-render
    let summary='✓ Battery status updated: '+updated+' changed, '+unchanged+' unchanged';
    if(skipped)summary+=', '+skipped+' skipped';
    if(notFound.length)summary+='\n\nRooms not found in system: '+notFound.slice(0,20).join(', ')+(notFound.length>20?' (and '+(notFound.length-20)+' more)':'');
    alert(summary);
    if(typeof renderFloors==='function')renderFloors();
    if(typeof updateStats==='function')updateStats();
  }catch(e){
    alert('Battery refresh failed:\n\n'+e.message+'\n\nExpected file location: Default document library > '+BATTERY_FILE_PATH);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='🔋 Refresh battery'}
  }
}

// ============================================================
// COMPANIES MANAGEMENT (v14.0.10)
// ============================================================
let editingCompanyId=null;

function openCompaniesPanel(){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  document.getElementById('coSearch').value='';
  renderCompaniesList();
  document.getElementById('companiesModal').classList.add('open');
}

function renderCompaniesList(){
  const list=document.getElementById('companiesList');
  const stats=document.getElementById('companiesStats');
  const q=(document.getElementById('coSearch').value||'').toLowerCase().trim();
  // Count usage per company: scan bookings + rates
  const usage={};
  allBookings.forEach(b=>{
    const c=(b.Company||'').trim();if(c){usage[c]=usage[c]||{bookings:0,billing:0,rates:0};usage[c].bookings++}
    const bc=(b.Billing_Company||'').trim();if(bc){usage[bc]=usage[bc]||{bookings:0,billing:0,rates:0};usage[bc].billing++}
  });
  allRates.forEach(r=>{
    const c=(r.Company||'').trim();if(c){usage[c]=usage[c]||{bookings:0,billing:0,rates:0};usage[c].rates++}
  });
  // Build list
  const rows=[...allCompanies].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'','nb',{sensitivity:'base'}));
  const filtered=q?rows.filter(r=>(r.Title||'').toLowerCase().includes(q)||(r.OrgNr||'').includes(q)||(r.InvoiceEmail||'').toLowerCase().includes(q)):rows;
  const activeCount=rows.filter(r=>r.Active!==false).length;
  stats.textContent=rows.length+' companies ('+activeCount+' active)'+(q?' — '+filtered.length+' matching':'');
  if(!filtered.length){
    list.innerHTML='<div style="text-align:center;padding:30px;color:var(--text-tertiary);font-size:13px">No companies yet. Click "+ New company" to add one, or "🔍 Find unlinked" to scan your existing bookings.</div>';
    return;
  }
  list.innerHTML='<table style="width:100%;font-size:13px"><thead><tr style="background:var(--bg-secondary)"><th style="padding:6px 10px;text-align:left">Company</th><th style="padding:6px 10px;text-align:left">Org.nr</th><th style="padding:6px 10px;text-align:left">Invoice email</th><th style="padding:6px 10px;text-align:right">Bookings</th><th style="padding:6px 10px;text-align:right">Billing</th><th style="padding:6px 10px;text-align:right">Rates</th><th style="padding:6px 10px;text-align:left">Status</th><th style="padding:6px 10px;width:30px"></th></tr></thead><tbody>'
    +filtered.map(r=>{
      const u=usage[r.Title]||{bookings:0,billing:0,rates:0};
      const isInactive=r.Active===false;
      return '<tr onclick="openCompanyEdit(\''+r.id+'\')" style="cursor:pointer;border-top:.5px solid var(--border-tertiary);'+(isInactive?'opacity:.5':'')+'" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
        +'<td style="padding:6px 10px;font-weight:500">'+escapeHtml(r.Title||'')+'</td>'
        +'<td style="padding:6px 10px;font-family:monospace;font-size:12px">'+escapeHtml(r.OrgNr||'')+'</td>'
        +'<td style="padding:6px 10px;font-size:12px">'+escapeHtml(r.InvoiceEmail||'')+'</td>'
        +'<td style="padding:6px 10px;text-align:right">'+(u.bookings||'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:6px 10px;text-align:right">'+(u.billing||'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:6px 10px;text-align:right">'+(u.rates||'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:6px 10px">'+(isInactive?'<span class="pill" style="background:var(--bg-tertiary)">Inactive</span>':'<span class="pill" style="background:var(--bg-success);color:var(--text-success)">Active</span>')+'</td>'
        +'<td style="padding:6px 10px"><button onclick="event.stopPropagation();openCompanyEdit(\''+r.id+'\')" style="padding:3px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-primary);cursor:pointer;font-size:11px">Edit</button></td>'
        +'</tr>';
    }).join('')+'</tbody></table>';
}

function openCompanyEdit(companyId){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  editingCompanyId=companyId||null;
  const c=companyId?allCompanies.find(x=>x.id===companyId):null;
  document.getElementById('companyEditModalTitle').textContent=c?'Edit company':'New company';
  document.getElementById('coName').value=c?c.Title||'':'';
  document.getElementById('coOrgNr').value=c?c.OrgNr||'':'';
  document.getElementById('coInvoiceEmail').value=c?c.InvoiceEmail||'':'';
  document.getElementById('coInvoiceAddress').value=c?c.InvoiceAddress||'':'';
  document.getElementById('coContactName').value=c?c.ContactName||'':'';
  document.getElementById('coContactPhone').value=c?c.ContactPhone||'':'';
  document.getElementById('coNotes').value=c?c.Notes||'':'';
  document.getElementById('coActive').value=c?(c.Active===false?'false':'true'):'true';
  const brregStatus=document.getElementById('coBrregStatus');if(brregStatus)brregStatus.innerHTML='';
  document.getElementById('coDeleteBtn').style.display=c?'':'none';
  // Show usage info
  if(c){
    const name=c.Title||'';
    const bookings=allBookings.filter(b=>(b.Company||'').trim()===name).length;
    const billing=allBookings.filter(b=>(b.Billing_Company||'').trim()===name).length;
    const rates=allRates.filter(r=>(r.Company||'').trim()===name).length;
    document.getElementById('coUsageInfo').textContent='Used in: '+bookings+' bookings as Company · '+billing+' bookings as Billing · '+rates+' rates';
  }else{
    document.getElementById('coUsageInfo').textContent='';
  }
  document.getElementById('companyEditModal').classList.add('open');
}

async function saveCompany(){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  const name=document.getElementById('coName').value.trim();
  if(!name){alert('Company name is required');return}
  // Check for duplicate (unless editing this one)
  const dup=allCompanies.find(c=>(c.Title||'').toLowerCase()===name.toLowerCase()&&c.id!==editingCompanyId);
  if(dup){alert('A company with this name already exists: '+dup.Title);return}
  const fields={
    Title:name,
    OrgNr:document.getElementById('coOrgNr').value.trim()||null,
    InvoiceEmail:document.getElementById('coInvoiceEmail').value.trim()||null,
    InvoiceAddress:document.getElementById('coInvoiceAddress').value.trim()||null,
    ContactName:document.getElementById('coContactName').value.trim()||null,
    ContactPhone:document.getElementById('coContactPhone').value.trim()||null,
    Notes:document.getElementById('coNotes').value.trim()||null,
    Active:document.getElementById('coActive').value==='true'
  };
  try{
    if(editingCompanyId){
      await updateListItem('Companies',editingCompanyId,fields);
      const c=allCompanies.find(x=>x.id===editingCompanyId);
      if(c)Object.assign(c,fields);
    }else{
      const result=await createListItem('Companies',fields);
      allCompanies.push({id:result.id||'new',...fields});
    }
    document.getElementById('companyEditModal').classList.remove('open');
    renderCompaniesList();
    refreshPersonDatalists();
  }catch(e){alert('Failed: '+e.message)}
}

async function deleteCompany(){
  if(!editingCompanyId)return;
  const c=allCompanies.find(x=>x.id===editingCompanyId);
  if(!c)return;
  const name=c.Title;
  const usedIn=allBookings.filter(b=>(b.Company||'').trim()===name||(b.Billing_Company||'').trim()===name).length;
  const inRates=allRates.filter(r=>(r.Company||'').trim()===name).length;
  let msg='Delete company "'+name+'"?';
  if(usedIn||inRates){
    msg+='\n\nWarning: this company is used in '+usedIn+' bookings and '+inRates+' rates.\nThe text references will remain but will no longer appear in dropdowns.';
  }
  msg+='\n\nConsider setting it to Inactive instead.';
  if(!confirm(msg))return;
  try{
    await deleteListItem('Companies',editingCompanyId);
    allCompanies=allCompanies.filter(x=>x.id!==editingCompanyId);
    editingCompanyId=null;
    document.getElementById('companyEditModal').classList.remove('open');
    renderCompaniesList();
    refreshPersonDatalists();
  }catch(e){alert('Failed: '+e.message)}
}

function findUnlinkedCompanies(){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  // Scan bookings + rates for company names not in Companies list
  const existing=new Set(allCompanies.map(c=>(c.Title||'').toLowerCase().trim()));
  const found={};
  allBookings.forEach(b=>{
    [b.Company,b.Billing_Company].forEach(name=>{
      const n=(name||'').trim();if(!n)return;
      if(!existing.has(n.toLowerCase())){
        found[n]=(found[n]||0)+1;
      }
    });
  });
  allRates.forEach(r=>{
    const n=(r.Company||'').trim();if(!n)return;
    if(!existing.has(n.toLowerCase())){
      found[n]=(found[n]||0)+1;
    }
  });
  const names=Object.keys(found).sort((a,b)=>found[b]-found[a]);
  if(!names.length){alert('✓ All companies used in bookings and rates are already registered in the Companies list.');return}
  const preview=names.slice(0,20).map(n=>'• '+n+' ('+found[n]+' uses)').join('\n');
  const extra=names.length>20?'\n\n...and '+(names.length-20)+' more':'';
  if(!confirm('Found '+names.length+' unlinked companies in your bookings/rates:\n\n'+preview+extra+'\n\nCreate Company records for all of them now? (You can edit them individually afterwards.)'))return;
  (async()=>{
    let created=0,failed=0;
    for(let i=0;i<names.length;i++){
      try{
        const result=await createListItem('Companies',{Title:names[i],Active:true});
        allCompanies.push({id:result.id||'new_'+i,Title:names[i],Active:true});
        created++;
      }catch(e){console.error('Failed to create '+names[i]+':',e);failed++}
      if(i%10===9)await new Promise(r=>setTimeout(r,300));
    }
    alert('✓ Created '+created+' companies'+(failed?', '+failed+' failed':'')+'.');
    renderCompaniesList();
    refreshPersonDatalists();
  })();
}

// Check if company name exists in Companies list; show inline warning with quick-add link
function checkCompanyRegistration(name,warnElId){
  const el=document.getElementById(warnElId);if(!el)return;
  const trimmed=(name||'').trim();
  if(!trimmed){el.innerHTML='';return}
  const match=allCompanies.find(c=>(c.Title||'').toLowerCase()===trimmed.toLowerCase());
  if(match){
    if(match.Active===false){
      el.innerHTML='<span style="color:var(--text-warning)">⚠ "'+escapeHtml(match.Title)+'" is marked inactive.</span>';
    }else{
      el.innerHTML='';
    }
    return;
  }
  // Not registered
  if(can('manage_companies')||can('admin')){
    el.innerHTML='<span style="color:var(--text-warning)">⚠ "'+escapeHtml(trimmed)+'" not in Companies list. <a href="javascript:void(0)" onclick="quickAddCompany(\''+trimmed.replace(/'/g,"\\'")+'\')" style="color:var(--accent)">Add now</a></span>';
  }else{
    el.innerHTML='<span style="color:var(--text-tertiary)">"'+escapeHtml(trimmed)+'" is a new company (not in registry).</span>';
  }
}

// Quickly add a company with just the name, then re-check
async function quickAddCompany(name){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  try{
    const result=await createListItem('Companies',{Title:name,Active:true});
    allCompanies.push({id:result.id||'new',Title:name,Active:true});
    refreshPersonDatalists();
    // Clear warnings on both fields if they match
    ['fCompanyWarn','fBillingCompanyWarn'].forEach(id=>{
      const el=document.getElementById(id);if(el)el.innerHTML='<span style="color:var(--text-success)">✓ Added "'+escapeHtml(name)+'" to Companies</span>';
    });
  }catch(e){alert('Failed to add company: '+e.message)}
}

// ============================================================
// BRREG LOOKUP (v14.0.10)
// ============================================================
// Fetches company information from Brønnøysundregistrene open API.
// https://data.brreg.no/enhetsregisteret/api/enheter/{orgnr}
async function lookupBrreg(){
  const rawNr=(document.getElementById('coOrgNr').value||'').trim();
  const status=document.getElementById('coBrregStatus');
  // Strip spaces, dashes, dots
  const orgNr=rawNr.replace(/[\s\-.]/g,'');
  if(!/^\d{9}$/.test(orgNr)){
    status.innerHTML='<span style="color:var(--text-danger)">⚠ Ugyldig org.nr (må være 9 sifre)</span>';
    return;
  }
  status.innerHTML='<span style="color:var(--text-tertiary)">Henter fra brreg.no...</span>';
  try{
    const r=await fetch('https://data.brreg.no/enhetsregisteret/api/enheter/'+orgNr,{headers:{Accept:'application/json'}});
    if(r.status===404){
      status.innerHTML='<span style="color:var(--text-danger)">✕ Fant ikke org.nr '+orgNr+' i Enhetsregisteret</span>';
      return;
    }
    if(!r.ok){
      status.innerHTML='<span style="color:var(--text-danger)">✕ Feil fra brreg.no: '+r.status+'</span>';
      return;
    }
    const data=await r.json();
    // Warnings for problematic entities
    const warnings=[];
    if(data.slettedato)warnings.push('⚠ SLETTET '+data.slettedato);
    if(data.konkurs)warnings.push('⚠ KONKURS');
    if(data.underAvvikling)warnings.push('⚠ UNDER AVVIKLING');
    if(data.underTvangsavviklingEllerTvangsopplosning)warnings.push('⚠ UNDER TVANGSAVVIKLING');
    // Offer to pre-fill; confirm first if user already has data
    const name=data.navn||'';
    const addr=data.forretningsadresse||data.postadresse||{};
    const addrLines=(addr.adresse||[]).join(', ');
    const postalCode=addr.postnummer||'';
    const city=addr.poststed||'';
    const fullAddress=[addrLines,postalCode+' '+city].filter(x=>x.trim()).join(', ');
    const email=data.epostadresse||'';
    const phone=data.telefon||data.mobil||'';
    // Check if fields have existing data
    const existingName=document.getElementById('coName').value.trim();
    const existingAddr=document.getElementById('coInvoiceAddress').value.trim();
    const existingEmail=document.getElementById('coInvoiceEmail').value.trim();
    const existingPhone=document.getElementById('coContactPhone').value.trim();
    const hasExisting=existingName||existingAddr||existingEmail||existingPhone;
    if(hasExisting){
      const preview='Navn: '+name+'\nAdresse: '+fullAddress+(email?'\nEpost: '+email:'')+(phone?'\nTelefon: '+phone:'');
      if(!confirm('Overskriv eksisterende felter med data fra brreg.no?\n\n'+preview))return;
    }
    // Populate fields
    if(name)document.getElementById('coName').value=name;
    if(fullAddress)document.getElementById('coInvoiceAddress').value=fullAddress;
    if(email)document.getElementById('coInvoiceEmail').value=email;
    if(phone)document.getElementById('coContactPhone').value=phone;
    // Show status with warnings
    const bransje=data.naeringskode1?data.naeringskode1.beskrivelse:'';
    let msg='<span style="color:var(--text-success)">✓ Hentet: '+escapeHtml(name)+(bransje?' · '+escapeHtml(bransje):'')+'</span>';
    if(warnings.length){
      msg+='<div style="color:var(--text-danger);font-weight:500;margin-top:3px">'+warnings.join(' · ')+'</div>';
    }
    status.innerHTML=msg;
  }catch(e){
    status.innerHTML='<span style="color:var(--text-danger)">✕ Nettverksfeil: '+escapeHtml(e.message)+'</span>';
  }
}

// ============================================================
// PDF EXPORT VIA PRINT (v14.0.10)
// ============================================================
// Opens a print-friendly window containing the same data as exportInvoicingCSV.
// Browser's print dialog allows "Save as PDF" as the destination.
function exportInvoicingPDF(companyFilterName){
  const monthVal=document.getElementById('invMonth').value;
  const yearVal=document.getElementById('invYear').value;
  const fromVal=document.getElementById('invFrom').value;
  const toVal=document.getElementById('invTo').value;
  const useRange=!!(fromVal||toVal);
  let fromDate,toDate,periodLabel;
  const monthNames=['januar','februar','mars','april','mai','juni','juli','august','september','oktober','november','desember'];
  if(useRange){
    fromDate=fromVal?new Date(fromVal+'T00:00:00'):new Date(1970,0,1);
    toDate=toVal?new Date(toVal+'T23:59:59'):new Date(2100,0,1);
    periodLabel=(fromVal?formatDate(fromVal):'start')+' → '+(toVal?formatDate(toVal):'slutt');
  }else{
    const m=parseInt(monthVal),y=parseInt(yearVal);
    fromDate=new Date(y,m,1);
    toDate=new Date(y,m+1,0,23,59,59);
    periodLabel=monthNames[m]+' '+y;
  }
  if(companyFilterName===undefined){
    const cf=document.getElementById('invCompanyFilter');
    companyFilterName=cf&&cf.value!=='__ALL__'?cf.value:null;
  }
  const propTitle=selectedProperty?selectedProperty.Title:'Alle eiendommer';
  const currentRoomIds=new Set(rooms.map(r=>r.id));
  const propTitleForPercent=selectedProperty?selectedProperty.Title:'';

  // Identify full-tenant properties
  const fullTenantByPropId={};
  const viewedProps=selectedProperty?[selectedProperty]:properties;
  viewedProps.forEach(p=>{
    const ft=computeFullTenantForPeriod(p,fromDate,toDate);
    if(ft)fullTenantByPropId[p.id]=ft;
  });
  const fullTenantRoomIds=new Set(
    allRooms.filter(r=>fullTenantByPropId[r.PropertyLookupId]).map(r=>r.id)
  );
  // LongTerm contracts per room
  const longTermByRoomIdPdf={};
  allRooms.forEach(r=>{
    if(!currentRoomIds.has(r.id))return;
    if(fullTenantRoomIds.has(r.id))return;
    const lt=computeLongTermForRoomPeriod(r,fromDate,toDate);
    if(lt)longTermByRoomIdPdf[r.id]=lt;
  });
  const longTermRoomIdsPdf=new Set(Object.keys(longTermByRoomIdPdf));

  // Collect line items — grouped by effective billing company
  const groups={};
  const companyNightSum={};
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!currentRoomIds.has(rid))return;
    if(!b.Check_In)return;
    if(b.Status==='Cancelled')return;
    if(fullTenantRoomIds.has(rid))return;
    if(longTermRoomIdsPdf.has(rid))return;
    const effectiveCo=getEffectiveCompany(b);
    if(companyFilterName&&effectiveCo!==companyFilterName)return;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):new Date();co.setHours(0,0,0,0);
    if(co<fromDate||ci>toDate)return;
    const nights=_nightsInPeriod(b,fromDate,toDate);
    const cost=calcBookingCost(b,propTitleForPercent);
    const room=allRooms.find(r=>r.id===rid);
    const origCo=(b.Company||'').trim();
    const key=effectiveCo||'(uten firma)';
    if(!groups[key])groups[key]={nights:[],fees:[],percent:null,fullTenant:null,longTerm:[]};
    if(nights>0){
      groups[key].nights.push({
        name:b.Person_Name||'',guestCompany:origCo,effectiveCo,
        room:room?room.Title:'?',checkIn:b.Check_In,checkOut:b.Check_Out,
        nightsCount:nights,rate:cost.rate||0,total:nights*(cost.rate||0),source:cost.source||''
      });
      if(effectiveCo){companyNightSum[effectiveCo]=(companyNightSum[effectiveCo]||0)+nights*(cost.rate||0)}
    }
    const isContinuation=(b.Continuation===true||b.Continuation==='true'||b.Continuation===1);
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuation&&!hasPercentFee(effectiveCo,propTitleForPercent)){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,propTitleForPercent);
        if(fee>0){
          groups[key].fees.push({
            name:b.Person_Name||'',room:room?room.Title:'?',
            checkoutDate:b.Check_Out,fee
          });
        }
      }
    }
  });
  // Percent fees
  Object.keys(companyNightSum).forEach(c=>{
    if(companyFilterName&&c!==companyFilterName)return;
    const pct=getPercentFeeRate(c,propTitleForPercent);
    if(pct>0){
      const feeAmount=Math.round(companyNightSum[c]*pct);
      if(!groups[c])groups[c]={nights:[],fees:[],percent:null,fullTenant:null,longTerm:[]};
      groups[c].percent={rate:pct,base:companyNightSum[c],amount:feeAmount};
    }
  });
  // Full tenant
  Object.keys(fullTenantByPropId).forEach(pid=>{
    const ft=fullTenantByPropId[pid];
    if(companyFilterName&&ft.company!==companyFilterName)return;
    const prop=properties.find(p=>String(p.id)===String(pid));
    const key=ft.company;
    if(!groups[key])groups[key]={nights:[],fees:[],percent:null,fullTenant:null,longTerm:[]};
    groups[key].fullTenant={property:prop?prop.Title:'',rooms:ft.rooms,days:ft.days,rate:ft.rate,total:ft.total,detailLabel:ft.detailLabel};
  });
  // Long-term per-room contracts
  Object.keys(longTermByRoomIdPdf).forEach(rid=>{
    const lt=longTermByRoomIdPdf[rid];
    if(companyFilterName&&lt.company!==companyFilterName)return;
    const key=lt.company;
    if(!groups[key])groups[key]={nights:[],fees:[],percent:null,fullTenant:null,longTerm:[]};
    if(!groups[key].longTerm)groups[key].longTerm=[];
    groups[key].longTerm.push({
      roomTitle:lt.room.Title||'',
      price:lt.price,
      total:lt.total,
      detailLabel:lt.detailLabel,
      isMonthly:lt.isMonthly
    });
  });

  const groupKeys=Object.keys(groups).sort();
  if(!groupKeys.length){
    alert('Ingen data å eksportere for denne perioden.');
    return;
  }

  // Generate HTML
  const now=new Date();
  const nowStr=formatDate(now)+' '+String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0');
  const fmtKr=n=>(n||0).toLocaleString('nb-NO',{minimumFractionDigits:0,maximumFractionDigits:2})+' kr';

  let grandTotal=0;
  let bodyHtml='';

  groupKeys.forEach(key=>{
    const g=groups[key];
    let groupTotal=0;
    let tableRows='';

    // Full tenant section (first, stands out)
    if(g.fullTenant){
      const ft=g.fullTenant;
      groupTotal+=ft.total;
      tableRows+='<tr class="ft-row"><td colspan="4"><strong>🔒 Full-tenant lease — '+escapeHtml(ft.property)+'</strong><br><small>'+escapeHtml(ft.detailLabel||'')+'</small></td><td class="num"><strong>'+fmtKr(ft.total)+'</strong></td></tr>';
    }

    // Long-term per-room contracts: summary + collapsible detail rows
    if(g.longTerm&&g.longTerm.length){
      const ltTotal=g.longTerm.reduce((s,lt)=>s+lt.total,0);
      groupTotal+=ltTotal;
      const sectionId='lt-'+escapeHtml(key).replace(/[^a-zA-Z0-9]/g,'_');
      // Summary row — clickable for screen, always shows on print
      tableRows+='<tr class="lt-row lt-summary" onclick="document.querySelectorAll(\'.'+sectionId+'\').forEach(el=>el.classList.toggle(\'lt-hidden\'))">'
        +'<td colspan="4"><strong>🔑 Långtidsleie ('+g.longTerm.length+' rom)</strong> <span class="muted no-print">▼ klikk for detaljer</span></td>'
        +'<td class="num"><strong>'+fmtKr(ltTotal)+'</strong></td>'
        +'</tr>';
      // Detail rows — hidden by default on screen, always shown on print
      g.longTerm.forEach(lt=>{
        tableRows+='<tr class="lt-row lt-detail '+sectionId+' lt-hidden">'
          +'<td style="padding-left:24px"><small>↳ '+escapeHtml(lt.roomTitle)+'</small></td>'
          +'<td><small>'+escapeHtml(lt.roomTitle)+'</small></td>'
          +'<td><small>'+escapeHtml(lt.detailLabel||'')+'</small></td>'
          +'<td class="num"><small>'+fmtKr(lt.price)+(lt.isMonthly?'/mnd':'/dag')+'</small></td>'
          +'<td class="num"><small>'+fmtKr(lt.total)+'</small></td>'
          +'</tr>';
      });
    }

    // Night bookings
    g.nights.forEach(n=>{
      groupTotal+=n.total;
      const billingInfo=n.guestCompany&&n.guestCompany!==n.effectiveCo?'<br><small class="muted">Gjest jobber for: '+escapeHtml(n.guestCompany)+'</small>':'';
      tableRows+='<tr>'
        +'<td>'+escapeHtml(n.name)+billingInfo+'</td>'
        +'<td>'+escapeHtml(n.room)+'</td>'
        +'<td>'+formatDate(n.checkIn)+' → '+(n.checkOut?formatDate(n.checkOut):'Åpen')+'</td>'
        +'<td class="num">'+n.nightsCount+' × '+fmtKr(n.rate)+'</td>'
        +'<td class="num">'+fmtKr(n.total)+'</td>'
        +'</tr>';
    });

    // Checkout fees
    g.fees.forEach(f=>{
      groupTotal+=f.fee;
      tableRows+='<tr class="fee-row">'
        +'<td>↳ Utvask: '+escapeHtml(f.name)+'</td>'
        +'<td>'+escapeHtml(f.room)+'</td>'
        +'<td>'+formatDate(f.checkoutDate)+'</td>'
        +'<td class="num">—</td>'
        +'<td class="num">'+fmtKr(f.fee)+'</td>'
        +'</tr>';
    });

    // Percent fee
    if(g.percent){
      groupTotal+=g.percent.amount;
      tableRows+='<tr class="pct-row"><td colspan="4">📊 Månedsgebyr ('+(g.percent.rate*100)+'% av '+fmtKr(g.percent.base)+')</td><td class="num"><strong>'+fmtKr(g.percent.amount)+'</strong></td></tr>';
    }

    grandTotal+=groupTotal;

    bodyHtml+='<section class="company-section">'
      +'<h2>'+escapeHtml(key)+'</h2>'
      +'<table>'
      +'<thead><tr><th>Gjest</th><th>Rom</th><th>Periode</th><th class="num">Netter × sats</th><th class="num">Sum</th></tr></thead>'
      +'<tbody>'+tableRows+'</tbody>'
      +'<tfoot><tr><td colspan="4" class="num"><strong>Sum '+escapeHtml(key)+'</strong></td><td class="num"><strong>'+fmtKr(groupTotal)+'</strong></td></tr></tfoot>'
      +'</table>'
      +'</section>';
  });

  const title='Fakturagrunnlag — '+propTitle+(companyFilterName?' — '+companyFilterName:'')+' — '+periodLabel;

  const html='<!DOCTYPE html><html><head><meta charset="UTF-8"><title>'+escapeHtml(title)+'</title>'
    +'<style>'
    +'*{box-sizing:border-box}'
    +'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#1a1a1a;margin:20mm;font-size:11pt;line-height:1.4}'
    +'header{border-bottom:2px solid #1a1a1a;padding-bottom:12px;margin-bottom:20px}'
    +'header h1{font-size:18pt;margin:0 0 4px}'
    +'header .subtitle{color:#555;font-size:10pt}'
    +'header .meta{color:#888;font-size:9pt;margin-top:6px}'
    +'.company-section{margin-bottom:30px;page-break-inside:avoid}'
    +'.company-section h2{font-size:13pt;margin:0 0 8px;padding:6px 10px;background:#f0f0f0;border-left:3px solid #1D9E75}'
    +'table{width:100%;border-collapse:collapse;font-size:10pt}'
    +'th{text-align:left;padding:6px 8px;background:#fafafa;border-bottom:1px solid #ccc;font-weight:600}'
    +'td{padding:5px 8px;border-bottom:.5px solid #e5e5e5;vertical-align:top}'
    +'.num{text-align:right;font-variant-numeric:tabular-nums}'
    +'.muted{color:#888;font-size:9pt}'
    +'.ft-row{background:rgba(29,158,117,.08)}'
    +'.lt-row{background:rgba(14,165,165,.07)}'
    +'.lt-summary{cursor:pointer}'
    +'.lt-summary:hover{background:rgba(14,165,165,.14)}'
    +'.lt-hidden{display:none}'
    +'@media print{.lt-hidden{display:table-row !important}}'
    +'.fee-row{background:rgba(123,97,255,.04);color:#555}'
    +'.pct-row{background:rgba(239,159,39,.08)}'
    +'tfoot td{border-top:1.5px solid #1a1a1a;border-bottom:0;padding-top:8px;font-size:11pt}'
    +'.grand-total{margin-top:30px;padding:12px;background:#1D9E75;color:#fff;text-align:right;font-size:14pt;font-weight:600;page-break-inside:avoid}'
    +'footer{margin-top:40px;padding-top:10px;border-top:.5px solid #ccc;color:#888;font-size:9pt;text-align:center}'
    +'@media print{'
    +'  body{margin:15mm}'
    +'  @page{size:A4;margin:0}'
    +'  .no-print{display:none}'
    +'  header{margin-bottom:14px}'
    +'  .company-section{margin-bottom:18px}'
    +'}'
    +'.print-btn{position:fixed;top:20px;right:20px;padding:10px 20px;background:#1D9E75;color:#fff;border:0;border-radius:6px;cursor:pointer;font-size:14px;box-shadow:0 2px 8px rgba(0,0,0,.2)}'
    +'</style>'
    +'</head><body>'
    +'<button class="print-btn no-print" onclick="window.print()">🖨 Skriv ut / Lagre som PDF</button>'
    +'<header>'
    +'<h1>Fakturagrunnlag</h1>'
    +'<div class="subtitle">'+escapeHtml(propTitle)+(companyFilterName?' · '+escapeHtml(companyFilterName):'')+' · '+escapeHtml(periodLabel)+'</div>'
    +'<div class="meta">Generert '+nowStr+' · 2GM Booking</div>'
    +'</header>'
    +bodyHtml
    +'<div class="grand-total">Totalt: '+fmtKr(grandTotal)+'</div>'
    +'<footer>2GM Booking · genereret '+nowStr+'</footer>'
    +'</body></html>';

  const w=window.open('','_blank');
  if(!w){alert('Popup blocked. Allow popups for this site to export PDF.');return}
  w.document.write(html);
  w.document.close();
  // Auto-trigger print dialog after render
  setTimeout(()=>{try{w.focus();w.print()}catch(e){console.error(e)}},500);
}

// ============================================================
// PRICING TABS — Full-tenant + Long-term editors (v14.0.10)
// ============================================================
function switchPricingTab(tab){
  document.querySelectorAll('.pricing-tab').forEach(b=>{
    if(b.dataset.tab===tab){
      b.classList.add('pricing-tab-active');
      b.style.borderBottom='2px solid var(--accent)';
      b.style.color='';
      b.style.fontWeight='500';
    }else{
      b.classList.remove('pricing-tab-active');
      b.style.borderBottom='2px solid transparent';
      b.style.color='var(--text-secondary)';
      b.style.fontWeight='';
    }
  });
  document.querySelectorAll('.pricing-tab-content').forEach(d=>d.style.display='none');
  document.getElementById('tab-'+tab).style.display='';
  if(tab==='fulltenant')renderFullTenantList();
  else if(tab==='longterm'){populateLongTermPropertyFilter();renderLongTermList()}
}

// --- FULL-TENANT TAB ---
function renderFullTenantList(){
  const list=document.getElementById('fullTenantList');
  if(!properties.length){list.innerHTML='<div class="muted" style="padding:20px;text-align:center">No properties found.</div>';return}
  const fmtKr=n=>(n||0).toLocaleString('nb-NO');
  list.innerHTML='<table style="width:100%;font-size:13px"><thead><tr style="background:var(--bg-secondary)"><th style="padding:8px;text-align:left">Property</th><th style="padding:8px;text-align:left">Company</th><th style="padding:8px;text-align:right">Price/room</th><th style="padding:8px;text-align:left">Unit</th><th style="padding:8px;text-align:left">Start</th><th style="padding:8px;text-align:left">End</th><th style="padding:8px;text-align:left">Status</th><th style="padding:8px;width:60px"></th></tr></thead><tbody>'
    +properties.map(p=>{
      const company=(p.FullTenant_Company||'').trim();
      const rate=Number(p.FullTenant_RatePerRoom)||0;
      const unit=p.FullTenant_RateUnit||'Per day';
      const start=p.FullTenant_StartDate?formatDate(p.FullTenant_StartDate):'';
      const end=p.FullTenant_EndDate?formatDate(p.FullTenant_EndDate):'';
      const today=new Date();
      const startD=p.FullTenant_StartDate?new Date(p.FullTenant_StartDate):null;
      const endD=p.FullTenant_EndDate?new Date(p.FullTenant_EndDate):null;
      let status,statusClr;
      if(!company){status='—';statusClr='var(--text-tertiary)'}
      else if(startD&&today<startD){status='Upcoming';statusClr='var(--accent)'}
      else if(endD&&today>endD){status='Expired';statusClr='var(--text-tertiary)'}
      else{status='Active';statusClr='var(--text-success)'}
      return '<tr style="border-top:.5px solid var(--border-tertiary);cursor:pointer" onclick="openFullTenantEdit(\''+p.id+'\')" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
        +'<td style="padding:8px;font-weight:500">'+escapeHtml(p.Title||'')+'</td>'
        +'<td style="padding:8px">'+(company?escapeHtml(company):'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:8px;text-align:right">'+(rate?fmtKr(rate)+' kr':'<span class="muted">empty</span>')+'</td>'
        +'<td style="padding:8px"><small>'+escapeHtml(unit)+'</small></td>'
        +'<td style="padding:8px">'+(start||'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:8px">'+(end||'<span class="muted">—</span>')+'</td>'
        +'<td style="padding:8px"><span style="color:'+statusClr+';font-weight:500">'+status+'</span></td>'
        +'<td style="padding:8px"><button onclick="event.stopPropagation();openFullTenantEdit(\''+p.id+'\')" style="padding:3px 10px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-primary);cursor:pointer;font-size:11px">Edit</button></td>'
        +'</tr>';
    }).join('')+'</tbody></table>';
}

let editingFullTenantPropId=null;
function openFullTenantEdit(propId){
  editingFullTenantPropId=propId;
  const p=properties.find(x=>String(x.id)===String(propId));
  if(!p)return;
  document.getElementById('ftEditPropertyLabel').textContent=p.Title;
  document.getElementById('ftEditCompany').value=p.FullTenant_Company||'';
  document.getElementById('ftEditRate').value=p.FullTenant_RatePerRoom||'';
  document.getElementById('ftEditRateUnit').value=p.FullTenant_RateUnit||'Per day';
  document.getElementById('ftEditStartDate').value=p.FullTenant_StartDate?p.FullTenant_StartDate.substring(0,10):'';
  document.getElementById('ftEditEndDate').value=p.FullTenant_EndDate?p.FullTenant_EndDate.substring(0,10):'';
  document.getElementById('fullTenantEditModal').classList.add('open');
}

async function saveFullTenantAgreement(){
  if(!editingFullTenantPropId)return;
  const company=document.getElementById('ftEditCompany').value.trim();
  const rate=parseFloat(document.getElementById('ftEditRate').value)||null;
  const rateUnit=document.getElementById('ftEditRateUnit').value;
  const startDate=document.getElementById('ftEditStartDate').value;
  const endDate=document.getElementById('ftEditEndDate').value;
  const fields={
    FullTenant_Company:company||null,
    FullTenant_RatePerRoom:rate,
    FullTenant_RateUnit:rateUnit,
    FullTenant_StartDate:startDate||null,
    FullTenant_EndDate:endDate||null
  };
  try{
    await updateListItem('Properties',editingFullTenantPropId,fields);
    const p=properties.find(x=>String(x.id)===String(editingFullTenantPropId));
    if(p)Object.assign(p,fields);
    document.getElementById('fullTenantEditModal').classList.remove('open');
    renderFullTenantList();
  }catch(e){alert('Save failed: '+e.message)}
}

// --- LONG-TERM TAB ---
function populateLongTermPropertyFilter(){
  const sel=document.getElementById('ltFilterProperty');
  const current=sel.value;
  sel.innerHTML='<option value="__ALL__">All properties</option>'+properties.map(p=>'<option value="'+p.id+'">'+escapeHtml(p.Title)+'</option>').join('');
  if(current)sel.value=current;
}

function renderLongTermList(){
  const list=document.getElementById('longTermList');
  const filter=document.getElementById('ltFilterProperty').value;
  let rooms=allRooms;
  if(filter&&filter!=='__ALL__')rooms=rooms.filter(r=>String(r.PropertyLookupId)===String(filter));
  // Sort by property then room title
  rooms=[...rooms].sort((a,b)=>{
    const pa=String(a.PropertyLookupId||'');
    const pb=String(b.PropertyLookupId||'');
    if(pa!==pb)return pa.localeCompare(pb);
    return (a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true});
  });
  if(!rooms.length){list.innerHTML='<div class="muted" style="padding:20px;text-align:center">No rooms in this property.</div>';return}
  const fmtKr=n=>(n||0).toLocaleString('nb-NO');
  let html='<table style="width:100%;font-size:13px"><thead><tr style="background:var(--bg-secondary)"><th style="padding:8px;text-align:left">Property</th><th style="padding:8px;text-align:left">Room</th><th style="padding:8px;text-align:left">Company</th><th style="padding:8px;text-align:right">Price</th><th style="padding:8px;text-align:left">Unit</th><th style="padding:8px;text-align:left">Period</th><th style="padding:8px;text-align:left">Status</th><th style="padding:8px;width:60px"></th></tr></thead><tbody>';
  rooms.forEach(r=>{
    const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
    const propName=prop?prop.Title:'';
    const company=(r.LongTerm_Company||'').trim();
    const price=Number(r.LongTerm_Price)||0;
    const unit=r.LongTerm_RateUnit||'Per day';
    const start=r.LongTerm_StartDate?formatDate(r.LongTerm_StartDate):'';
    const end=r.LongTerm_EndDate?formatDate(r.LongTerm_EndDate):'(open)';
    const today=new Date();
    const startD=r.LongTerm_StartDate?new Date(r.LongTerm_StartDate):null;
    const endD=r.LongTerm_EndDate?new Date(r.LongTerm_EndDate):null;
    let status,statusClr;
    if(!company){status='—';statusClr='var(--text-tertiary)'}
    else if(!price){status='⚠ No price';statusClr='var(--text-warning)'}
    else if(startD&&today<startD){status='Upcoming';statusClr='var(--accent)'}
    else if(endD&&today>endD){status='Expired';statusClr='var(--text-tertiary)'}
    else{status='Active';statusClr='var(--text-success)'}
    const periodText=company?(start+' → '+end):'<span class="muted">—</span>';
    html+='<tr style="border-top:.5px solid var(--border-tertiary);cursor:pointer'+(company?'':';opacity:.7')+'" onclick="openLongTermEdit(\''+r.id+'\')" onmouseover="this.style.background=\'var(--bg-secondary)\'" onmouseout="this.style.background=\'\'">'
      +'<td style="padding:8px"><small>'+escapeHtml(propName)+'</small></td>'
      +'<td style="padding:8px;font-weight:500">'+escapeHtml(r.Title||'')+'</td>'
      +'<td style="padding:8px">'+(company?escapeHtml(company):'<span class="muted">—</span>')+'</td>'
      +'<td style="padding:8px;text-align:right">'+(price?fmtKr(price)+' kr':'<span class="muted">—</span>')+'</td>'
      +'<td style="padding:8px"><small>'+escapeHtml(unit)+'</small></td>'
      +'<td style="padding:8px"><small>'+periodText+'</small></td>'
      +'<td style="padding:8px"><span style="color:'+statusClr+';font-weight:500">'+status+'</span></td>'
      +'<td style="padding:8px"><button onclick="event.stopPropagation();openLongTermEdit(\''+r.id+'\')" style="padding:3px 10px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-primary);cursor:pointer;font-size:11px">Edit</button></td>'
      +'</tr>';
  });
  html+='</tbody></table>';
  // Summary footer
  const activeRooms=rooms.filter(r=>(r.LongTerm_Company||'').trim()&&Number(r.LongTerm_Price)>0);
  const totalSum=activeRooms.reduce((s,r)=>s+(Number(r.LongTerm_Price)||0),0);
  if(activeRooms.length){
    html+='<div style="padding:10px 8px;margin-top:10px;background:rgba(14,165,165,.07);border-radius:6px;font-size:12px"><strong>'+activeRooms.length+' active contracts</strong> · Total: <strong>'+fmtKr(totalSum)+' kr</strong> <span class="muted">(per unit shown — does not factor pro-rata)</span></div>';
  }
  list.innerHTML=html;
}

let editingLongTermRoomId=null;
function openLongTermEdit(roomId){
  editingLongTermRoomId=roomId;
  const r=allRooms.find(x=>x.id===roomId);
  if(!r)return;
  const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
  document.getElementById('ltEditTitle').textContent='Long-term contract — '+(r.Title||'');
  document.getElementById('ltEditRoomLabel').textContent=(prop?prop.Title:'?')+' · '+(r.Title||'');
  document.getElementById('ltEditCompany').value=r.LongTerm_Company||'';
  document.getElementById('ltEditPrice').value=r.LongTerm_Price||'';
  document.getElementById('ltEditRateUnit').value=r.LongTerm_RateUnit||'Per month';
  document.getElementById('ltEditStartDate').value=r.LongTerm_StartDate?r.LongTerm_StartDate.substring(0,10):'';
  document.getElementById('ltEditEndDate').value=r.LongTerm_EndDate?r.LongTerm_EndDate.substring(0,10):'';
  document.getElementById('ltEditClearBtn').style.display=(r.LongTerm_Company||'').trim()?'':'none';
  document.getElementById('longTermEditModal').classList.add('open');
}

async function saveLongTermContract(){
  if(!editingLongTermRoomId)return;
  const company=document.getElementById('ltEditCompany').value.trim();
  const price=parseFloat(document.getElementById('ltEditPrice').value)||null;
  const rateUnit=document.getElementById('ltEditRateUnit').value;
  const startDate=document.getElementById('ltEditStartDate').value;
  const endDate=document.getElementById('ltEditEndDate').value;
  if(company){
    if(!price){alert('Price is required');return}
    if(!startDate){alert('Start date is required');return}
  }
  const fields={
    LongTerm_Company:company||null,
    LongTerm_Price:price,
    LongTerm_RateUnit:rateUnit,
    LongTerm_StartDate:startDate||null,
    LongTerm_EndDate:endDate||null
  };
  try{
    await updateListItem('Rooms',editingLongTermRoomId,fields);
    const r=allRooms.find(x=>x.id===editingLongTermRoomId);
    if(r)Object.assign(r,fields);
    document.getElementById('longTermEditModal').classList.remove('open');
    renderLongTermList();
  }catch(e){alert('Save failed: '+e.message)}
}

async function clearLongTermContract(){
  if(!editingLongTermRoomId)return;
  if(!confirm('Clear long-term contract for this room? This will make it a regular bookable room again.'))return;
  const fields={LongTerm_Company:null,LongTerm_Price:null,LongTerm_RateUnit:null,LongTerm_StartDate:null,LongTerm_EndDate:null};
  try{
    await updateListItem('Rooms',editingLongTermRoomId,fields);
    const r=allRooms.find(x=>x.id===editingLongTermRoomId);
    if(r)Object.assign(r,fields);
    document.getElementById('longTermEditModal').classList.remove('open');
    renderLongTermList();
  }catch(e){alert('Clear failed: '+e.message)}
}

function openLongTermBulkAdd(){
  const propSel=document.getElementById('ltBulkProperty');
  propSel.innerHTML=properties.map(p=>'<option value="'+p.id+'">'+escapeHtml(p.Title)+'</option>').join('');
  document.getElementById('ltBulkCompany').value='';
  document.getElementById('ltBulkStartDate').value='';
  document.getElementById('ltBulkRateUnit').value='Per month';
  renderLongTermBulkRoomList();
  document.getElementById('longTermBulkModal').classList.add('open');
}

function renderLongTermBulkRoomList(){
  const propId=document.getElementById('ltBulkProperty').value;
  const list=document.getElementById('ltBulkRoomList');
  const rooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(propId)).sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  if(!rooms.length){list.innerHTML='<div class="muted">No rooms.</div>';return}
  list.innerHTML='<div style="margin-bottom:6px"><label style="font-size:11px"><input type="checkbox" onchange="document.querySelectorAll(\'.ltBulkRoom\').forEach(cb=>cb.checked=this.checked)"> Select all</label></div>'+rooms.map(r=>{
    const has=(r.LongTerm_Company||'').trim();
    return '<label style="display:block;font-size:12px;padding:3px 0"><input type="checkbox" class="ltBulkRoom" value="'+r.id+'"'+(has?' disabled title="Already has contract"':'')+'> '+escapeHtml(r.Title||'')+(has?' <span class="muted" style="font-size:10px">(already: '+escapeHtml(has)+')</span>':'')+'</label>';
  }).join('');
}

async function bulkApplyLongTermContract(){
  const company=document.getElementById('ltBulkCompany').value.trim();
  const rateUnit=document.getElementById('ltBulkRateUnit').value;
  const startDate=document.getElementById('ltBulkStartDate').value;
  if(!company||!startDate){alert('Company and start date are required');return}
  const ids=[...document.querySelectorAll('.ltBulkRoom:checked')].map(cb=>cb.value);
  if(!ids.length){alert('Select at least one room');return}
  if(!confirm('Apply contract "'+company+'" to '+ids.length+' rooms? You will need to set price per room afterwards.'))return;
  let success=0,failed=0;
  for(let i=0;i<ids.length;i++){
    try{
      await updateListItem('Rooms',ids[i],{LongTerm_Company:company,LongTerm_RateUnit:rateUnit,LongTerm_StartDate:startDate});
      const r=allRooms.find(x=>x.id===ids[i]);
      if(r){r.LongTerm_Company=company;r.LongTerm_RateUnit=rateUnit;r.LongTerm_StartDate=startDate}
      success++;
    }catch(e){console.error(e);failed++}
    if(i%10===9)await new Promise(res=>setTimeout(res,300));
  }
  alert('Applied to '+success+' rooms'+(failed?', '+failed+' failed':'')+'. Now set the individual prices in the Long-term tab.');
  document.getElementById('longTermBulkModal').classList.remove('open');
  renderLongTermList();
}

// ============================================================
// BACKUP & RESTORE (v14.0.10)
// ============================================================
const BACKUP_LISTS=['Properties','Rooms','Bookings','Persons','Cleaning_Log','Hours','Users','Rates','Companies'];

async function exportBackup(){
  if(!can('admin')){alert('Admin permission required');return}
  const btn=document.querySelector('[data-backup-btn]');
  if(btn){btn.disabled=true;btn.textContent='⏳ Backing up...'}
  try{
    const data={
      meta:{
        appVersion:'v14.0.10',
        timestamp:new Date().toISOString(),
        exportedBy:currentUser.email||'unknown',
        siteId:siteId
      },
      lists:{}
    };
    for(let i=0;i<BACKUP_LISTS.length;i++){
      const name=BACKUP_LISTS[i];
      if(btn)btn.textContent='⏳ '+name+' ('+(i+1)+'/'+BACKUP_LISTS.length+')...';
      try{
        const items=await getListItems(name);
        data.lists[name]={count:items.length,items};
      }catch(e){
        data.lists[name]={count:0,items:[],error:e.message};
        console.warn('Skipping '+name+':',e.message);
      }
    }
    // Build summary
    const summary=Object.keys(data.lists).map(k=>k+': '+data.lists[k].count).join(', ');
    data.meta.summary=summary;
    // Download as JSON
    const json=JSON.stringify(data,null,2);
    const blob=new Blob([json],{type:'application/json'});
    const dateStr=new Date().toISOString().substring(0,10);
    const filename='2gmbooking_backup_'+dateStr+'.json';
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=filename;
    document.body.appendChild(a);a.click();
    setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url)},100);
    alert('✓ Backup ready: '+filename+'\n\n'+summary);
  }catch(e){
    alert('Backup failed: '+e.message);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='💾 Backup data'}
  }
}

// --- RESTORE: inspect only ---
let _backupInspectData=null;

function openRestoreInspect(){
  if(!can('admin')){alert('Admin permission required');return}
  document.getElementById('restoreFileInput').click();
}

function onRestoreFileSelected(event){
  const file=event.target.files[0];
  if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const data=JSON.parse(e.target.result);
      if(!data.meta||!data.lists){throw new Error('Not a valid 2GM Booking backup file')}
      _backupInspectData=data;
      renderRestoreInspect();
      document.getElementById('restoreInspectModal').classList.add('open');
    }catch(err){
      alert('Could not read file: '+err.message);
    }
  };
  reader.readAsText(file);
  // Reset input so same file can be selected again
  event.target.value='';
}

function renderRestoreInspect(){
  const data=_backupInspectData;
  if(!data)return;
  const meta=data.meta;
  const exportedDate=new Date(meta.timestamp);
  const ageDays=Math.floor((new Date()-exportedDate)/86400000);
  let html='<div style="background:var(--bg-secondary);padding:10px;border-radius:6px;margin-bottom:12px;font-size:12px">'
    +'<strong>Backup info</strong><br>'
    +'Exported: '+formatDate(meta.timestamp)+' ('+ageDays+' days ago) by '+escapeHtml(meta.exportedBy||'?')+'<br>'
    +'App version: '+escapeHtml(meta.appVersion||'?')+'<br>'
    +'Summary: '+escapeHtml(meta.summary||'')
    +'</div>';
  // List selector
  html+='<div style="margin-bottom:8px"><label style="font-size:12px">Pick a list: </label>'
    +'<select id="restoreListPicker" onchange="renderRestoreItems()" style="padding:5px 8px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:13px;font-family:inherit">'
    +'<option value="">— choose —</option>'
    +Object.keys(data.lists).map(k=>'<option value="'+k+'">'+k+' ('+data.lists[k].count+')</option>').join('')
    +'</select>'
    +' <input id="restoreSearch" type="text" placeholder="Search..." oninput="renderRestoreItems()" style="margin-left:8px;padding:5px 8px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:13px;font-family:inherit;width:240px">'
    +'</div>';
  html+='<div id="restoreItemsContainer" style="border:1px solid var(--border-tertiary);border-radius:6px;padding:8px;min-height:200px;max-height:400px;overflow:auto;font-size:12px"><div class="muted" style="text-align:center;padding:30px">Pick a list to inspect items</div></div>';
  document.getElementById('restoreInspectBody').innerHTML=html;
}

function renderRestoreItems(){
  const data=_backupInspectData;
  if(!data)return;
  const listName=document.getElementById('restoreListPicker').value;
  const search=(document.getElementById('restoreSearch').value||'').toLowerCase().trim();
  const container=document.getElementById('restoreItemsContainer');
  if(!listName){container.innerHTML='<div class="muted" style="text-align:center;padding:30px">Pick a list to inspect items</div>';return}
  const list=data.lists[listName];
  if(!list||!list.items||!list.items.length){
    container.innerHTML='<div class="muted" style="text-align:center;padding:30px">No items in this list.</div>';
    return;
  }
  // Filter by search
  let items=list.items;
  if(search){
    items=items.filter(item=>{
      const blob=JSON.stringify(item).toLowerCase();
      return blob.indexOf(search)>=0;
    });
  }
  if(!items.length){container.innerHTML='<div class="muted" style="text-align:center;padding:30px">No items match search.</div>';return}
  // Show as expandable rows
  let html='<div class="muted" style="font-size:11px;margin-bottom:6px">Showing '+items.length+(search?' / '+list.items.length:'')+' items. Click to expand.</div>';
  items.slice(0,200).forEach((item,idx)=>{
    const titleField=item.Title||item.Person_Name||item.Name||item.Email||('item #'+(item.id||idx));
    const itemJson=JSON.stringify(item,null,2);
    const elId='restore-item-'+listName+'-'+idx;
    html+='<div style="border-bottom:.5px solid var(--border-tertiary);padding:4px 0">'
      +'<div style="display:flex;justify-content:space-between;align-items:center;cursor:pointer" onclick="document.getElementById(\''+elId+'\').classList.toggle(\'restore-collapsed\')">'
      +'<span style="font-weight:500">'+escapeHtml(String(titleField))+'</span>'
      +'<span><button onclick="event.stopPropagation();restoreSingleItem(\''+listName+'\','+idx+')" style="padding:2px 8px;background:#1D9E75;color:#fff;border:0;border-radius:4px;font-size:11px;cursor:pointer">↻ Restore</button></span>'
      +'</div>'
      +'<pre id="'+elId+'" class="restore-collapsed" style="font-family:Consolas,monospace;font-size:11px;background:var(--bg-secondary);padding:6px;border-radius:4px;margin:4px 0 0 0;white-space:pre-wrap;max-height:200px;overflow:auto">'+escapeHtml(itemJson)+'</pre>'
      +'</div>';
  });
  if(items.length>200)html+='<div class="muted" style="text-align:center;padding:8px;font-size:11px">Showing first 200. Use search to narrow.</div>';
  container.innerHTML=html;
  // Make CSS work
  if(!document.getElementById('restoreCSS')){
    const css=document.createElement('style');
    css.id='restoreCSS';
    css.textContent='.restore-collapsed{display:none}';
    document.head.appendChild(css);
  }
}

async function restoreSingleItem(listName,idx){
  if(!can('admin')){alert('Admin permission required');return}
  const data=_backupInspectData;
  if(!data)return;
  const item=data.lists[listName].items[idx];
  if(!item){alert('Item not found in backup');return}
  // Build a clean fields object — strip system fields
  const skipFields={id:1,_id:1,'@odata.etag':1,Created:1,Modified:1,Author:1,Editor:1,AuthorLookupId:1,EditorLookupId:1,ContentType:1,_UIVersionString:1,_ColorTag:1,Attachments:1,LinkTitle:1,LinkTitleNoMenu:1,ItemChildCount:1,FolderChildCount:1,_ComplianceFlags:1,_ComplianceTag:1,_ComplianceTagWrittenTime:1,_ComplianceTagUserId:1,AppAuthor:1,AppEditor:1};
  const fields={};
  Object.keys(item).forEach(k=>{if(!skipFields[k]&&!k.startsWith('OData_'))fields[k]=item[k]});
  // Confirm
  const titleField=item.Title||item.Person_Name||item.Name||'item';
  const summary='Restore "'+titleField+'" to '+listName+'?\n\n'
    +'A NEW item will be created with this data. The original ID ('+item.id+') will not be reused — SharePoint assigns a new one.\n\n'
    +'Note: Lookup-references (e.g. RoomLookupId, PropertyLookupId) will be copied as-is. If the referenced room/property no longer exists, the item may not display correctly.\n\n'
    +'Proceed?';
  if(!confirm(summary))return;
  try{
    const result=await createListItem(listName,fields);
    alert('✓ Restored as new item with id '+result.id+' in '+listName);
    // Reload data so it shows up
    await loadData();
  }catch(e){
    alert('Restore failed: '+e.message);
  }
}
