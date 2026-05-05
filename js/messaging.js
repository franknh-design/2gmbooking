// ============================================================
// 2GM Booking v14.7.0 — messaging.js
// SMS, e-post, maler, massemelding, toast
// ============================================================

async function deleteListItem(listName,itemId){
  const s=await getSiteId();const lid=await getListId(listName);
  return graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+itemId);
}

// ============================================================
// MESSAGING — SMS & E-post (v14.5.10)
// ============================================================
const DEFAULT_SMS_TEMPLATE=`Hello {first_name},
Welcome to {property}.
Room: {room}, door code: {room_door_code}
WiFi: {wifi_ssid} / {wifi_password}
{floor_info}
{welcome_message}
Best regards, Frank — 2GM`;

const DEFAULT_EMAIL_TEMPLATE=`Dear {first_name},

Welcome to {property}.

Room: {room}
Door code: {room_door_code}
WiFi: {wifi_ssid}
WiFi password: {wifi_password}
Check-in date: {check_in_date}

{floor_info}

{welcome_message}

We hope you enjoy your stay.

Best regards,
{my_name}
{my_phone}
{my_email}`;

const DEFAULT_EMAIL_SUBJECT='Welcome to {property} — room {room}';

// v15.3: Door-tag template (per-property HTML in Properties.DoorTag_Template).
// Same placeholders as SMS/Email plus {company} and {check_out_date}.
const DEFAULT_DOORTAG_TEMPLATE=`<div style="font-family:Arial,sans-serif;padding:40px;max-width:600px;margin:0 auto">
  <div style="text-align:center;margin-bottom:30px">
    <div style="font-size:72px;font-weight:700;letter-spacing:2px">{room}</div>
    <div style="font-size:14px;color:#888;margin-top:4px">{property}</div>
  </div>
  <div style="border-top:2px solid #2C2C2A;padding-top:20px">
    <h2 style="font-size:18px;margin:0 0 16px">Welcome, {guest_name}</h2>
    <table style="font-size:14px;width:100%">
      <tr><td style="padding:6px 0;color:#888;width:120px">Company</td><td>{company}</td></tr>
      <tr><td style="padding:6px 0;color:#888">Check-in</td><td>After 15:00 — {check_in_date}</td></tr>
      <tr><td style="padding:6px 0;color:#888">Check-out</td><td>Before 12:00 — {check_out_date}</td></tr>
    </table>
  </div>
  <div style="margin-top:24px;padding:16px;background:#f5f4ef;border-radius:8px;font-size:13px">
    <strong>Room information</strong><br>The room will be washed once a week.<br>New towels every week, and new beddings biweekly.
  </div>
  <div style="margin-top:16px;padding:16px;background:#f5f4ef;border-radius:8px;font-size:13px">
    <strong>Contact</strong><br>Questions? Contact {my_name}: {my_phone} · {my_email}
  </div>
  <div style="text-align:center;margin-top:40px;font-size:16px;color:#888">Have a nice stay :)</div>
</div>`;

function _renderTemplate(template,vars){
  let out=template||'';
  Object.keys(vars).forEach(k=>{
    const re=new RegExp('\\{'+k+'\\}','g');
    out=out.replace(re,vars[k]||'');
  });
  return out;
}

function _buildMessageVars(booking){
  const room=allRooms.find(r=>r.id===String(booking.RoomLookupId));
  const property=room?properties.find(p=>String(p.id)===String(room.PropertyLookupId)):null;
  const fullName=booking.Person_Name||'';
  const firstName=fullName.split(/\s+/)[0]||fullName;
  const checkIn=booking.Check_In?formatDate(booking.Check_In):'';
  // Find person record to get phone/email
  const person=allPersons.find(p=>(p.Name||p.Title||'').toLowerCase()===fullName.toLowerCase());
  // Floor-specific info (v14.5.10): pick Floor1_Info or Floor2_Info based on room.Floor
  let floorInfo='';
  if(property&&room){
    const floor=String(room.Floor||'').trim();
    if(floor==='1')floorInfo=property.Floor1_Info||'';
    else if(floor==='2')floorInfo=property.Floor2_Info||'';
  }
  const checkOut=booking.Check_Out?formatDate(booking.Check_Out):'Open-ended';
  return{
    guest_name:fullName,
    first_name:firstName,
    property:property?property.Title:'',
    room:room?room.Title:'',
    room_door_code:room?room.Door_Code||'(not set)':'',
    wifi_ssid:property?property.WiFi_SSID||'(not set)':'',
    wifi_password:property?property.WiFi_Password||'(not set)':'',
    welcome_message:property?property.Welcome_Message||'':'',
    floor_info:floorInfo,
    check_in_date:checkIn,
    check_out_date:checkOut,                  // v15.3: brukes på door tag
    company:(booking.Company||'—').toString(),// v15.3: brukes på door tag
    my_name:'Frank Haugan',
    my_phone:'+47 99 10 10 41',
    my_email:'frank@2gm.no',
    _person_phone:person?person.Mobile||booking.Mobile||'':booking.Mobile||'',
    _person_email:person?person.Email||booking.Email||'':booking.Email||''
  };
}

// v15.3: Returnerer ferdig HTML for én booking sin dørmerke.
// Bruker DoorTag_Template på rommets eiendom hvis satt, ellers DEFAULT_DOORTAG_TEMPLATE.
function _renderDoorTagHtml(booking){
  const room=allRooms.find(r=>r.id===String(booking.RoomLookupId));
  const property=room?properties.find(p=>String(p.id)===String(room.PropertyLookupId)):null;
  const tmpl=(property&&property.DoorTag_Template)||DEFAULT_DOORTAG_TEMPLATE;
  return _renderTemplate(tmpl,_buildMessageVars(booking));
}

// v15.4: Hjelpere for live preview + variabel-innsetting i Templates-modalen.
const _DOORTAG_VARS=['first_name','guest_name','property','room','room_door_code','wifi_ssid','wifi_password','welcome_message','floor_info','check_in_date','check_out_date','company','my_name','my_phone','my_email'];

function _previewVarsForDoorTag(){
  const sel=document.getElementById('tmplPropSel');if(!sel)return null;
  const p=properties.find(x=>String(x.id)===String(sel.value));if(!p)return null;
  const propRooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(p.id));
  const ids=new Set(propRooms.map(r=>r.id));
  const sample=allBookings.find(b=>ids.has(String(b.RoomLookupId))&&b.Person_Name);
  if(sample)return _buildMessageVars(sample);
  // Mock-data så preview fungerer selv uten reelle bookinger på eiendommen
  const room=propRooms[0];
  return{guest_name:'Ola Nordmann',first_name:'Ola',property:p.Title||'',room:room?room.Title:'101',room_door_code:room?(room.Door_Code||''):'',wifi_ssid:p.WiFi_SSID||'',wifi_password:p.WiFi_Password||'',welcome_message:p.Welcome_Message||'',floor_info:p.Floor1_Info||'',check_in_date:'2026-05-10',check_out_date:'2026-05-17',company:'Eksempel AS',my_name:'Frank Haugan',my_phone:'+47 99 10 10 41',my_email:'frank@2gm.no'};
}

function _renderDoorTagVarBar(){
  const bar=document.getElementById('doorTagVarBar');if(!bar)return;
  bar.innerHTML='<span style="font-size:11px;color:var(--text-secondary);margin-right:4px">Insert variable:</span>'+_DOORTAG_VARS.map(v=>'<button type="button" onclick="_insertDoorTagVar(\''+v+'\')" style="padding:2px 8px;border:1px solid var(--border-tertiary);border-radius:4px;background:var(--bg-secondary);cursor:pointer;font-size:11px;font-family:Consolas,monospace">{'+v+'}</button>').join('');
}

function _insertDoorTagVar(name){
  const ta=document.getElementById('tmplDoorTag');if(!ta)return;
  const s=ta.selectionStart,e=ta.selectionEnd,tag='{'+name+'}';
  ta.value=ta.value.slice(0,s)+tag+ta.value.slice(e);
  ta.selectionStart=ta.selectionEnd=s+tag.length;
  ta.focus();
  _updateDoorTagPreview();
}

function _updateDoorTagPreview(){
  const ta=document.getElementById('tmplDoorTag'),ifr=document.getElementById('tmplDoorTagPreview');
  if(!ta||!ifr)return;
  const vars=_previewVarsForDoorTag()||{};
  const html=_renderTemplate(ta.value||'',vars);
  ifr.srcdoc='<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:8px;background:#f0f0f0"><div style="background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.1);max-width:600px;margin:0 auto">'+html+'</div></body></html>';
}

function _getTemplate(booking,kind){
  const room=allRooms.find(r=>r.id===String(booking.RoomLookupId));
  const property=room?properties.find(p=>String(p.id)===String(room.PropertyLookupId)):null;
  if(kind==='sms')return (property&&property.SMS_Template)||DEFAULT_SMS_TEMPLATE;
  if(kind==='email')return (property&&property.Email_Template)||DEFAULT_EMAIL_TEMPLATE;
  if(kind==='subject')return (property&&property.Email_Subject_Template)||DEFAULT_EMAIL_SUBJECT;
  return '';
}

function _toast(msg){
  const t=document.createElement('div');
  t.textContent=msg;
  t.style.cssText='position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#1D9E75;color:#fff;padding:10px 20px;border-radius:6px;font-size:14px;z-index:10000;box-shadow:0 4px 12px rgba(0,0,0,.2)';
  document.body.appendChild(t);
  setTimeout(()=>{t.style.transition='opacity .4s';t.style.opacity='0';setTimeout(()=>document.body.removeChild(t),400)},1800);
}

async function _copyToClipboard(text){
  try{
    await navigator.clipboard.writeText(text);
    _toast('✓ Kopiert! Lim inn i Mobil-app eller e-post');
    return true;
  }catch(e){
    // Fallback
    const ta=document.createElement('textarea');
    ta.value=text;document.body.appendChild(ta);ta.select();
    try{document.execCommand('copy');_toast('✓ Kopiert!');document.body.removeChild(ta);return true}
    catch(e2){document.body.removeChild(ta);alert('Kunne ikke kopiere. Tekst:\n\n'+text);return false}
  }
}

function copyBookingSMS(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const vars=_buildMessageVars(b);
  const text=_renderTemplate(_getTemplate(b,'sms'),vars);
  _copyToClipboard(text);
}

function copyBookingEmail(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const vars=_buildMessageVars(b);
  const subject=_renderTemplate(_getTemplate(b,'subject'),vars);
  const body=_renderTemplate(_getTemplate(b,'email'),vars);
  _copyToClipboard('Emne: '+subject+'\n\n'+body);
}

function openBookingSMS(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const vars=_buildMessageVars(b);
  const text=_renderTemplate(_getTemplate(b,'sms'),vars);
  const phone=vars._person_phone||'';
  const url='sms:'+encodeURIComponent(phone)+'?body='+encodeURIComponent(text);
  window.location.href=url;
  setTimeout(()=>_toast('Hvis SMS-app ikke åpnet seg, bruk Kopier-knappen i stedet'),1500);
}

function openBookingEmail(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const vars=_buildMessageVars(b);
  const subject=_renderTemplate(_getTemplate(b,'subject'),vars);
  const body=_renderTemplate(_getTemplate(b,'email'),vars);
  const email=vars._person_email||'';
  const url='mailto:'+encodeURIComponent(email)+'?subject='+encodeURIComponent(subject)+'&body='+encodeURIComponent(body);
  window.location.href=url;
}

// --- MASS MESSAGING ---
function openMassMessage(kind){
  const m=document.getElementById('massMessageModal');
  document.getElementById('massMsgKind').value=kind;
  document.getElementById('massMsgTitle').textContent=kind==='sms'?'📱 Mass SMS':'📧 Mass e-post';
  // Populate property and company filters
  const propSel=document.getElementById('massPropFilter');
  propSel.innerHTML='<option value="__ALL__">Alle eiendommer</option>'+properties.map(p=>'<option value="'+p.id+'">'+escapeHtml(p.Title)+'</option>').join('');
  // Get all unique companies from active/upcoming bookings
  const today=new Date();today.setHours(0,0,0,0);
  const activeBookings=allBookings.filter(b=>{
    if(b.Status==='Cancelled'||b.Status==='Completed')return false;
    if(!b.Check_In)return false;
    return true;
  });
  const cos=[...new Set(activeBookings.map(b=>b.Company||'').filter(Boolean))].sort();
  const coSel=document.getElementById('massCoFilter');
  coSel.innerHTML='<option value="__ALL__">Alle firma</option>'+cos.map(c=>'<option value="'+escapeHtml(c)+'">'+escapeHtml(c)+'</option>').join('');
  renderMassMessageList();
  m.classList.add('open');
}

function renderMassMessageList(){
  const propFilter=document.getElementById('massPropFilter').value;
  const coFilter=document.getElementById('massCoFilter').value;
  const today=new Date();today.setHours(0,0,0,0);
  const items=allBookings.filter(b=>{
    if(b.Status==='Cancelled'||b.Status==='Completed')return false;
    if(!b.Check_In)return false;
    if(propFilter!=='__ALL__'){
      const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
      if(!room||String(room.PropertyLookupId)!==String(propFilter))return false;
    }
    if(coFilter!=='__ALL__'&&b.Company!==coFilter)return false;
    return true;
  }).sort((a,b)=>{
    const ca=(a.Company||'').localeCompare(b.Company||'');
    if(ca!==0)return ca;
    return (a.Person_Name||'').localeCompare(b.Person_Name||'');
  });
  const list=document.getElementById('massMsgList');
  if(!items.length){list.innerHTML='<div class="muted" style="text-align:center;padding:20px">Ingen aktive bookinger matcher filtrene.</div>';return}
  let html='<div style="margin-bottom:6px"><label style="font-size:11px"><input type="checkbox" onchange="document.querySelectorAll(\'.massMsgRow\').forEach(cb=>cb.checked=this.checked)"> Velg alle</label></div>';
  items.forEach(b=>{
    const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
    const prop=room?properties.find(p=>String(p.id)===String(room.PropertyLookupId)):null;
    const kind=document.getElementById('massMsgKind').value;
    const contact=kind==='sms'?(b.Mobile||''):(b.Email||'');
    const hasContact=contact.trim().length>0;
    const checkInStr=b.Check_In?formatDate(b.Check_In):'';
    html+='<label style="display:flex;align-items:center;gap:8px;padding:6px 4px;border-bottom:.5px solid var(--border-tertiary);font-size:12px;'+(hasContact?'':'opacity:.5')+'">'
      +'<input type="checkbox" class="massMsgRow" value="'+b.id+'"'+(hasContact?' checked':' disabled')+'>'
      +'<div style="flex:1"><strong>'+escapeHtml(b.Person_Name||'')+'</strong> · '+escapeHtml(b.Company||'(uten firma)')+'</div>'
      +'<div style="color:var(--text-tertiary);font-size:11px">'+escapeHtml(prop?prop.Title:'')+' / '+escapeHtml(room?room.Title:'')+'</div>'
      +'<div style="color:var(--text-tertiary);font-size:11px;min-width:90px">'+checkInStr+'</div>'
      +'<div style="color:'+(hasContact?'var(--text-success)':'var(--text-warning)')+';font-size:11px;min-width:130px">'+(hasContact?escapeHtml(contact):'⚠ mangler')+'</div>'
      +'</label>';
  });
  list.innerHTML=html;
}

async function executeMassMessage(){
  const kind=document.getElementById('massMsgKind').value;
  const ids=[...document.querySelectorAll('.massMsgRow:checked')].map(cb=>cb.value);
  if(!ids.length){alert('Velg minst én gjest');return}
  const action=document.querySelector('input[name="massAction"]:checked').value;
  const bookings=ids.map(id=>allBookings.find(b=>b.id===id)).filter(Boolean);
  if(action==='copy_all'){
    // Build one big text with all messages, separated by lines
    const parts=bookings.map(b=>{
      const vars=_buildMessageVars(b);
      const txt=_renderTemplate(_getTemplate(b,kind==='sms'?'sms':'email'),vars);
      const contact=kind==='sms'?vars._person_phone:vars._person_email;
      return '═══ '+vars.guest_name+' ('+(contact||'mangler kontakt')+') ═══\n'+txt;
    });
    await _copyToClipboard(parts.join('\n\n'));
    alert('Kopiert '+bookings.length+' meldinger til utklippstavle. Hver melding er separert med ═══');
  }else if(action==='open_each'){
    if(!confirm('Dette åpner '+kind.toUpperCase()+'-app for hver av '+bookings.length+' gjester. Du må sende hver manuelt. Fortsett?'))return;
    for(let i=0;i<bookings.length;i++){
      if(kind==='sms')openBookingSMS(bookings[i].id);
      else openBookingEmail(bookings[i].id);
      if(i<bookings.length-1)await new Promise(r=>setTimeout(r,2000));
    }
  }
  document.getElementById('massMessageModal').classList.remove('open');
}

// --- TEMPLATES EDITOR (per property) ---
function openTemplatesEditor(){
  if(!can('admin')&&!can('manage_properties')){alert('Permission required');return}
  const m=document.getElementById('templatesModal');
  const sel=document.getElementById('tmplPropSel');
  sel.innerHTML=properties.map(p=>'<option value="'+p.id+'">'+escapeHtml(p.Title)+'</option>').join('');
  loadTemplateForProperty();
  m.classList.add('open');
}

function loadTemplateForProperty(){
  const propId=document.getElementById('tmplPropSel').value;
  const p=properties.find(x=>String(x.id)===String(propId));
  if(!p)return;
  document.getElementById('tmplWifiSsid').value=p.WiFi_SSID||'';
  document.getElementById('tmplWifiPwd').value=p.WiFi_Password||'';
  document.getElementById('tmplWelcome').value=p.Welcome_Message||'';
  const f1=document.getElementById('tmplFloor1');if(f1)f1.value=p.Floor1_Info||'';
  const f2=document.getElementById('tmplFloor2');if(f2)f2.value=p.Floor2_Info||'';
  document.getElementById('tmplSms').value=p.SMS_Template||DEFAULT_SMS_TEMPLATE;
  document.getElementById('tmplEmailSubj').value=p.Email_Subject_Template||DEFAULT_EMAIL_SUBJECT;
  document.getElementById('tmplEmail').value=p.Email_Template||DEFAULT_EMAIL_TEMPLATE;
  // v15.3: dørmerke-mal
  const dt=document.getElementById('tmplDoorTag');if(dt)dt.value=p.DoorTag_Template||DEFAULT_DOORTAG_TEMPLATE;
  // v15.4: render variabel-knapper og live preview
  _renderDoorTagVarBar();
  _updateDoorTagPreview();
}

// v14.5.10: Reset-knapper for hver mal
function resetTemplateField(field){
  if(field==='sms'){
    if(!confirm('Reset SMS template to default for this property?\n\n(Click "Lagre for valgt eiendom" after to save.)'))return;
    document.getElementById('tmplSms').value=DEFAULT_SMS_TEMPLATE;
  }else if(field==='subject'){
    if(!confirm('Reset Email subject to default for this property?\n\n(Click "Lagre for valgt eiendom" after to save.)'))return;
    document.getElementById('tmplEmailSubj').value=DEFAULT_EMAIL_SUBJECT;
  }else if(field==='email'){
    if(!confirm('Reset Email template to default for this property?\n\n(Click "Lagre for valgt eiendom" after to save.)'))return;
    document.getElementById('tmplEmail').value=DEFAULT_EMAIL_TEMPLATE;
  }else if(field==='doortag'){
    if(!confirm('Reset Door tag template to default for this property?\n\n(Click "Lagre for valgt eiendom" after to save.)'))return;
    document.getElementById('tmplDoorTag').value=DEFAULT_DOORTAG_TEMPLATE;
    _updateDoorTagPreview();
  }
}

async function saveTemplateForProperty(){
  const propId=document.getElementById('tmplPropSel').value;
  const p=properties.find(x=>String(x.id)===String(propId));
  if(!p)return;
  const f1=document.getElementById('tmplFloor1');
  const f2=document.getElementById('tmplFloor2');
  const dt=document.getElementById('tmplDoorTag');
  const fields={
    WiFi_SSID:document.getElementById('tmplWifiSsid').value.trim()||null,
    WiFi_Password:document.getElementById('tmplWifiPwd').value.trim()||null,
    Welcome_Message:document.getElementById('tmplWelcome').value||null,
    Floor1_Info:f1?(f1.value||null):null,
    Floor2_Info:f2?(f2.value||null):null,
    SMS_Template:document.getElementById('tmplSms').value||null,
    Email_Subject_Template:document.getElementById('tmplEmailSubj').value.trim()||null,
    Email_Template:document.getElementById('tmplEmail').value||null,
    DoorTag_Template:dt?(dt.value||null):null   // v15.3
  };
  try{
    await updateListItem('Properties',p.id,fields);
    Object.assign(p,fields);
    _toast('✓ Template saved for '+p.Title);
  }catch(e){alert('Save failed: '+e.message)}
}

function previewTemplate(kind){
  const propId=document.getElementById('tmplPropSel').value;
  const p=properties.find(x=>String(x.id)===String(propId));
  if(!p){alert('Velg eiendom først');return}
  // Find a sample booking from this property
  const propRooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(p.id));
  const propRoomIds=new Set(propRooms.map(r=>r.id));
  const sample=allBookings.find(b=>propRoomIds.has(String(b.RoomLookupId))&&b.Person_Name);
  if(!sample){alert('Ingen booking funnet på '+p.Title+' for forhåndsvisning. Bruk en gjest som er booket der.');return}
  const vars=_buildMessageVars(sample);
  if(kind==='sms'){
    const tmpl=document.getElementById('tmplSms').value;
    alert('SMS-forhåndsvisning (basert på '+vars.guest_name+'):\n\n'+_renderTemplate(tmpl,vars));
  }else if(kind==='email'){
    const subj=_renderTemplate(document.getElementById('tmplEmailSubj').value,vars);
    const body=_renderTemplate(document.getElementById('tmplEmail').value,vars);
    alert('E-post-forhåndsvisning (basert på '+vars.guest_name+'):\n\nEmne: '+subj+'\n\n'+body);
  }else if(kind==='doortag'){
    // v15.3: Åpne forhåndsvisning av dørmerke i nytt vindu (rendret HTML, ikke alert).
    const tmpl=document.getElementById('tmplDoorTag').value||DEFAULT_DOORTAG_TEMPLATE;
    const html=_renderTemplate(tmpl,vars);
    const w=window.open('','_blank','width=700,height=900');
    if(!w){alert('Popup blokkert — tillat popups for å forhåndsvise dørmerket.');return}
    w.document.write('<!DOCTYPE html><html><head><title>Door tag preview — '+(vars.guest_name||'')+' / '+vars.room+'</title></head><body style="margin:0;background:#f0f0f0;padding:20px"><div style="background:#fff;box-shadow:0 2px 8px rgba(0,0,0,.1);max-width:680px;margin:0 auto">'+html+'</div></body></html>');
    w.document.close();
  }
}

