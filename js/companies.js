// ============================================================
// 2GM Booking v14.7.0 — companies.js
// Firmahåndtering, Brreg-oppslag, sammenslåing
// ============================================================

function openCompaniesPanel(){
  if(!can('manage_companies')&&!can('admin')){alert('Access denied');return}
  document.getElementById('coSearch').value='';
  // Close other panels (single-panel mode)
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  document.getElementById('invoicingPanel').classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
  document.getElementById('mainView').classList.add('panel-mode');
  document.getElementById('companiesPanel').classList.add('open');
  renderCompaniesList();
}

function toggleCompaniesPanel(){
  const p=document.getElementById('companiesPanel');
  if(p.classList.contains('open')){
    p.classList.remove('open');
    document.getElementById('mainView').classList.remove('panel-mode');
  }else{
    openCompaniesPanel();
  }
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
// BRREG LOOKUP (v14.5.10)
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

