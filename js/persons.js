// ============================================================
// 2GM Booking v14.7.0 — persons.js
// Gjestekort, personliste, historikk
// ============================================================

function togglePersons(){
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
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
        const washes=calcWashDates(active.Check_In,active.Check_Out,active.id);
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
// CHARTS (v14.5.10) — pure SVG, no dependencies
// ============================================================

// Reusable bar chart: data = [{label, value, subtitle?}]
