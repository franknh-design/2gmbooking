// ============================================================
// 2GM Booking v14.7.0 — admin.js
// Brukerstyring, backup/restore, rapporter, diagnostikk
// ============================================================

function openAdminPanel(){
  if(!can('admin'))return;
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
  document.getElementById('mainView').classList.add('panel-mode');
  document.getElementById('adminPanel').classList.add('open');
  renderAdminUsers();
}
function closeAdminPanel(){
  document.getElementById('adminPanel').classList.remove('open');
  document.getElementById('mainView').classList.remove('panel-mode');
}
function toggleAdminPanel(){
  const p=document.getElementById('adminPanel');
  if(p.classList.contains('open'))closeAdminPanel();else openAdminPanel();
}

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
  ensureMainView();
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const ip=document.getElementById('invoicingPanel');if(ip)ip.classList.remove('open');
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
  document.getElementById('mainView').classList.add('panel-mode');
  document.getElementById('pricingPanel').classList.add('open');
  renderRatesPanel();
}

function togglePricingPanel(){
  const p=document.getElementById('pricingPanel');
  if(p.classList.contains('open')){
    p.classList.remove('open');
    document.getElementById('mainView').classList.remove('panel-mode');
  }else{
    openRatesPanel();
  }
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
// PERSONS / CUSTOMERS (v14.5.10)
// ============================================================
let editingPersonId=null;

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
// CLEANING EFFICIENCY ANALYSIS (v14.5.10)
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
// MORE MENU (v14.5.10)
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
      const w=calcWashDates(b.Check_In,b.Check_Out,b.id);
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
        const washes=calcWashDates(b.Check_In,b.Check_Out,b.id);
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
// BATTERY REFRESH (v14.5.10)
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
        const nowIso=new Date().toISOString();
        await updateListItem('Rooms',room.id,{Door_Battery_Level:e.bat,Door_Battery_Updated:nowIso});
        room.Door_Battery_Level=e.bat;
        room.Door_Battery_Updated=nowIso;
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
    // Show low-battery alert (v14.5.10) — locks under 30%
    showLowBatteryAlert();
    if(typeof updateStats==='function')updateStats();
  }catch(e){
    alert('Battery refresh failed:\n\n'+e.message+'\n\nExpected file location: Default document library > '+BATTERY_FILE_PATH);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='🔋 Refresh battery'}
  }
}

// ============================================================
// COMPANIES MANAGEMENT (v14.5.10)
// ============================================================
let editingCompanyId=null;

async function exportBackup(){
  if(!can('admin')){alert('Admin permission required');return}
  const btn=document.querySelector('[data-backup-btn]');
  if(btn){btn.disabled=true;btn.textContent='⏳ Backing up...'}
  try{
    const data={
      meta:{
        appVersion:'v14.5.10',
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

// ============================================================
// COMPANY MERGE (v14.5.10)
// ============================================================
function openMergeCompanies(){
  if(!can('manage_companies')&&!can('admin')){alert('Permission required');return}
  // Populate datalist with all known company names
  const allCos=new Set();
  allCompanies.forEach(c=>{if(c.Title)allCos.add(c.Title)});
  allBookings.forEach(b=>{if(b.Company)allCos.add(b.Company);if(b.Billing_Company)allCos.add(b.Billing_Company)});
  allPersons.forEach(p=>{if(p.Company)allCos.add(p.Company)});
  document.getElementById('mergeCanonicalList').innerHTML=[...allCos].sort().map(c=>'<option value="'+escapeHtml(c)+'">').join('');
  document.getElementById('mergeCanonical').value='';
  document.getElementById('mergeAliases').value='';
  document.getElementById('mergePreview').innerHTML='<span class="muted">Skriv inn kanonisk navn og minst ett alias for å se forhåndsvisning</span>';
  document.getElementById('mergeConfirmBtn').disabled=true;
  document.getElementById('mergeCompaniesModal').classList.add('open');
}

function _parseAliases(text){
  return text.split(/[\n,]/).map(s=>s.trim()).filter(s=>s.length>0);
}

function _findItemsForCompany(name){
  const lc=name.toLowerCase();
  // Bookings (Company)
  const bookingsCompany=allBookings.filter(b=>(b.Company||'').toLowerCase()===lc);
  // Bookings (Billing_Company)
  const bookingsBilling=allBookings.filter(b=>(b.Billing_Company||'').toLowerCase()===lc);
  // Persons
  const persons=allPersons.filter(p=>(p.Company||'').toLowerCase()===lc);
  // Rates
  const rates=allRates.filter(r=>(r.Company||'').toLowerCase()===lc);
  // Properties (FullTenant_Company)
  const propsFT=properties.filter(p=>(p.FullTenant_Company||'').toLowerCase()===lc);
  // Rooms (LongTerm_Company)
  const roomsLT=allRooms.filter(r=>(r.LongTerm_Company||'').toLowerCase()===lc);
  // Companies-list itself
  const cosEntries=allCompanies.filter(c=>(c.Title||'').toLowerCase()===lc);
  return {bookingsCompany,bookingsBilling,persons,rates,propsFT,roomsLT,cosEntries};
}

function renderMergePreview(){
  const canonical=document.getElementById('mergeCanonical').value.trim();
  const aliasText=document.getElementById('mergeAliases').value;
  const aliases=_parseAliases(aliasText);
  const preview=document.getElementById('mergePreview');
  const btn=document.getElementById('mergeConfirmBtn');
  if(!canonical||!aliases.length){
    preview.innerHTML='<span class="muted">Skriv inn kanonisk navn og minst ett alias for å se forhåndsvisning</span>';
    btn.disabled=true;
    return;
  }
  // Filter out aliases that match canonical (case-insensitive)
  const realAliases=aliases.filter(a=>a.toLowerCase()!==canonical.toLowerCase());
  if(!realAliases.length){
    preview.innerHTML='<span style="color:var(--text-warning)">⚠ Alle alias er like det kanoniske navnet — ingenting å slå sammen</span>';
    btn.disabled=true;
    return;
  }
  // Compute totals
  let totals={bookings:0,billing:0,persons:0,rates:0,propsFT:0,roomsLT:0,cosEntries:0};
  const perAlias=[];
  realAliases.forEach(a=>{
    const found=_findItemsForCompany(a);
    perAlias.push({alias:a,counts:{
      bookings:found.bookingsCompany.length,
      billing:found.bookingsBilling.length,
      persons:found.persons.length,
      rates:found.rates.length,
      propsFT:found.propsFT.length,
      roomsLT:found.roomsLT.length,
      cosEntries:found.cosEntries.length
    }});
    totals.bookings+=found.bookingsCompany.length;
    totals.billing+=found.bookingsBilling.length;
    totals.persons+=found.persons.length;
    totals.rates+=found.rates.length;
    totals.propsFT+=found.propsFT.length;
    totals.roomsLT+=found.roomsLT.length;
    totals.cosEntries+=found.cosEntries.length;
  });
  const totalChanges=totals.bookings+totals.billing+totals.persons+totals.rates+totals.propsFT+totals.roomsLT+totals.cosEntries;
  let html='<div style="margin-bottom:8px"><strong>Vil endres til "'+escapeHtml(canonical)+'":</strong></div>';
  html+='<table style="width:100%;font-size:12px;border-collapse:collapse">'
    +'<thead><tr style="background:rgba(239,159,39,.08)"><th style="padding:6px 8px;text-align:left">Alias</th><th style="padding:6px 8px;text-align:right">Bookings</th><th style="padding:6px 8px;text-align:right">Billing</th><th style="padding:6px 8px;text-align:right">Persons</th><th style="padding:6px 8px;text-align:right">Rates</th><th style="padding:6px 8px;text-align:right">Properties</th><th style="padding:6px 8px;text-align:right">Rooms</th><th style="padding:6px 8px;text-align:right">Co list</th></tr></thead><tbody>';
  perAlias.forEach(a=>{
    const c=a.counts;
    const total=c.bookings+c.billing+c.persons+c.rates+c.propsFT+c.roomsLT+c.cosEntries;
    const styleSuffix=total===0?';color:var(--text-tertiary)':'';
    html+='<tr style="border-top:.5px solid var(--border-tertiary)'+styleSuffix+'">'
      +'<td style="padding:6px 8px"><strong>'+escapeHtml(a.alias)+'</strong>'+(total===0?' <small class="muted">(ikke funnet)</small>':'')+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.bookings+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.billing+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.persons+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.rates+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.propsFT+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.roomsLT+'</td>'
      +'<td style="padding:6px 8px;text-align:right">'+c.cosEntries+'</td>'
      +'</tr>';
  });
  html+='<tr style="border-top:1.5px solid var(--border-secondary);font-weight:600;background:rgba(239,159,39,.04)">'
    +'<td style="padding:6px 8px">Totalt</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.bookings+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.billing+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.persons+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.rates+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.propsFT+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.roomsLT+'</td>'
    +'<td style="padding:6px 8px;text-align:right">'+totals.cosEntries+'</td>'
    +'</tr></tbody></table>';
  if(totalChanges===0){
    html+='<div style="color:var(--text-warning);margin-top:8px;font-size:11px">⚠ Ingen items vil endres — sjekk skrivemåten på alias-navnene</div>';
    btn.disabled=true;
  }else{
    html+='<div style="margin-top:10px;padding:8px;background:rgba(123,97,255,.06);border-radius:4px;font-size:11px;color:#5949c4"><strong>'+totalChanges+' items vil oppdateres + '+totals.cosEntries+' Companies-rader vil slettes.</strong> Operasjonen er ikke reversibel — ta backup først (More → 💾 Backup data).</div>';
    btn.disabled=false;
  }
  preview.innerHTML=html;
}

async function confirmMergeCompanies(){
  if(!can('manage_companies')&&!can('admin')){alert('Permission required');return}
  const canonical=document.getElementById('mergeCanonical').value.trim();
  const aliases=_parseAliases(document.getElementById('mergeAliases').value).filter(a=>a.toLowerCase()!==canonical.toLowerCase());
  if(!canonical||!aliases.length)return;
  if(!confirm('Slå sammen '+aliases.length+' alias til "'+canonical+'"?\n\nDette kan ikke angres. Husk backup først (More → 💾 Backup data).\n\nTrykk OK for å fortsette.'))return;
  const btn=document.getElementById('mergeConfirmBtn');
  btn.disabled=true;btn.textContent='⏳ Slår sammen...';
  let success=0,failed=0;
  const errors=[];
  // Helper: bulk-update one list with throttling
  async function updateMany(listName,items,fields){
    for(let i=0;i<items.length;i++){
      try{
        await updateListItem(listName,items[i].id,fields);
        // Also update local cache so UI re-renders correctly
        Object.assign(items[i],fields);
        success++;
      }catch(e){
        console.error('Failed update '+listName+' #'+items[i].id,e);
        errors.push(listName+' #'+items[i].id+': '+e.message);
        failed++;
      }
      if(i%10===9)await new Promise(r=>setTimeout(r,300));
    }
  }
  for(let aIdx=0;aIdx<aliases.length;aIdx++){
    const a=aliases[aIdx];
    btn.textContent='⏳ Behandler "'+a+'" ('+(aIdx+1)+'/'+aliases.length+')...';
    const found=_findItemsForCompany(a);
    // Bookings: Company field
    await updateMany('Bookings',found.bookingsCompany,{Company:canonical});
    // Bookings: Billing_Company field
    await updateMany('Bookings',found.bookingsBilling,{Billing_Company:canonical});
    // Persons
    await updateMany('Persons',found.persons,{Company:canonical});
    // Rates
    await updateMany('Rates',found.rates,{Company:canonical});
    // Properties (FullTenant_Company)
    await updateMany('Properties',found.propsFT,{FullTenant_Company:canonical});
    // Rooms (LongTerm_Company)
    await updateMany('Rooms',found.roomsLT,{LongTerm_Company:canonical});
    // Companies-list: delete alias entries (only after confirming canonical exists in Companies list)
    const canonicalExists=allCompanies.some(c=>(c.Title||'').toLowerCase()===canonical.toLowerCase());
    if(canonicalExists){
      for(let i=0;i<found.cosEntries.length;i++){
        try{
          await deleteListItem('Companies',found.cosEntries[i].id);
          // Remove from local cache
          const idx=allCompanies.indexOf(found.cosEntries[i]);
          if(idx>=0)allCompanies.splice(idx,1);
          success++;
        }catch(e){console.error('Delete Company '+found.cosEntries[i].id,e);errors.push('Companies #'+found.cosEntries[i].id+': '+e.message);failed++}
        if(i%5===4)await new Promise(r=>setTimeout(r,300));
      }
    }else{
      // Canonical not in Companies list yet — rename first alias to canonical instead of deleting
      if(found.cosEntries.length){
        try{
          await updateListItem('Companies',found.cosEntries[0].id,{Title:canonical});
          found.cosEntries[0].Title=canonical;
          success++;
          // Delete the rest
          for(let i=1;i<found.cosEntries.length;i++){
            try{
              await deleteListItem('Companies',found.cosEntries[i].id);
              const idx=allCompanies.indexOf(found.cosEntries[i]);
              if(idx>=0)allCompanies.splice(idx,1);
              success++;
            }catch(e){failed++;errors.push('Companies #'+found.cosEntries[i].id+': '+e.message)}
          }
        }catch(e){failed++;errors.push('Companies rename: '+e.message)}
      }
    }
  }
  btn.disabled=false;btn.textContent='🔀 Slå sammen';
  let msg='✓ Merge ferdig.\n\n'+success+' items oppdatert/slettet.';
  if(failed)msg+='\n\n⚠ '+failed+' feil:\n'+errors.slice(0,5).join('\n')+(errors.length>5?'\n...':'');
  alert(msg);
  document.getElementById('mergeCompaniesModal').classList.remove('open');
  // Re-render
  if(typeof renderCompaniesList==='function')renderCompaniesList();
  refreshLocal();
}

// Helper: deleteListItem

// ============================================================
// MESSAGING — SMS & E-post (v14.5.10)
// ============================================================
// ============================================================
// 2GM Booking v14.6.0 — modules.js
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
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
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
  const roomIds=new Set(rooms.map(r=>r.id));
  const upcoming=allBookings.filter(b=>{
    if(b.Status!=='Upcoming')return false;
    const rid=String(b.RoomLookupId||'');if(!roomIds.has(rid))return false;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    // v14.5.10: include today (was: ci>=tomorrow)
    return ci>=today&&ci<=in30;
  }).sort((a,b)=>new Date(a.Check_In)-new Date(b.Check_In));
  const body=document.getElementById('incomingBody');
  if(!upcoming.length){body.innerHTML='<tr><td colspan="7" class="loading">No upcoming bookings</td></tr>';return}
  body.innerHTML=upcoming.map(b=>{
    const room=rooms.find(r=>r.id===String(b.RoomLookupId));const roomTitle=room?room.Title:'?';
    const ciDate=new Date(b.Check_In);ciDate.setHours(0,0,0,0);
    const daysUntil=Math.round((ciDate-today)/864e5);let badge='';
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
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
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

