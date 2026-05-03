// ============================================================
// 2GM Booking v14.7.0 — tuya.js
// Tuya Smart Lock integrasjon via Azure Function proxy
// ============================================================


function _tuyaUrl(endpoint){
  return `${TUYA_PROXY_BASE}/tuya/${endpoint}?code=${TUYA_FUNCTION_KEY}`;
}

async function _tuyaPost(endpoint, body){
  const r = await fetch(_tuyaUrl(endpoint), {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body),
  });
  if(!r.ok) throw new Error(`HTTP ${r.status} fra Tuya-proxy`);
  const data = await r.json();
  if(!data.success) throw new Error(data.error || 'Ukjent Tuya-feil');
  return data.result;
}

async function _tuyaGet(endpoint, params={}){
  const qs = new URLSearchParams({code: TUYA_FUNCTION_KEY, ...params}).toString();
  const r = await fetch(`${TUYA_PROXY_BASE}/tuya/${endpoint}?${qs}`);
  if(!r.ok) throw new Error(`HTTP ${r.status} fra Tuya-proxy`);
  const data = await r.json();
  if(!data.success) throw new Error(data.error || 'Ukjent Tuya-feil');
  return data.result;
}

function _generate6DigitCode(){
  const weakPatterns=['000000','111111','222222','333333','444444','555555','666666','777777','888888','999999','123456','654321','012345','543210'];
  let code, attempts=0;
  do{
    code=String(Math.floor(Math.random()*1000000)).padStart(6,'0');
    attempts++;
  }while(weakPatterns.includes(code)&&attempts<10);
  return code;
}

// ---- OPPRETT PIN PÅ LÅS ----
async function tuyaCreatePin(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking ikke funnet');return}
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  if(!room){alert('Rom ikke funnet');return}

  const deviceId=room.Tuya_Device_ID;
  if(!deviceId){
    alert('Rom «'+(room.Title||'?')+'» har ingen Tuya_Device_ID.\n\nLegg inn enhets-ID i SharePoint-listen Rooms.');
    return;
  }

  // Generer ny kode (forskjellig fra eksisterende)
  let newPin=_generate6DigitCode();
  let attempts=0;
  while(newPin===room.Door_Code&&attempts<5){newPin=_generate6DigitCode();attempts++}

  // Varighet: fra check-in til check-out + 2 timer buffer
  const ciRaw=b.Check_In||new Date().toISOString();
  const coRaw=b.Check_Out;
  const validFrom=Math.floor(new Date(ciRaw).getTime()/1000);
  const validTo=coRaw
    ? Math.floor(new Date(coRaw).getTime()/1000) + 7200
    : validFrom + (30*24*60*60);

  const guestName=(b.Person_Name||'Gjest').substring(0,20);

  if(!confirm(
    '🔑 Opprett PIN på lås for '+guestName+'?\n\n'
    +'Rom: '+(room.Title||'?')+'\n'
    +'PIN: '+newPin+'\n'
    +'Gyldig: '+new Date(validFrom*1000).toLocaleDateString('nb-NO')+' → '+new Date(validTo*1000).toLocaleDateString('nb-NO')
  ))return;

  const btn=document.getElementById('btnTuyaCreate_'+bookingId);
  if(btn){btn.disabled=true;btn.textContent='⏳ Oppretter...'}
  try{
    const result=await _tuyaPost('create_pin',{
      device_id: deviceId,
      pin: newPin,
      name: guestName,
      valid_from: validFrom,
      valid_to: validTo,
    });

    // Lagre PIN og password_id i SharePoint
    await updateListItem('Rooms',room.id,{Door_Code:newPin});
    room.Door_Code=newPin;
    await updateListItem('Bookings',bookingId,{Tuya_Password_ID:String(result.password_id)});
    b.Tuya_Password_ID=String(result.password_id);

    _toast('✓ PIN '+newPin+' aktivert på låsen (ID: '+result.password_id+')');
    showTuyaPinDisplay(room, newPin, result.password_id, bookingId);

    if(typeof showBookingDetail==='function')showBookingDetail(bookingId);
  }catch(e){
    alert('❌ Tuya feilet:\n\n'+e.message);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='🔑 Opprett PIN på lås'}
  }
}

// ---- SLETT PIN FRA LÅS ----
async function tuyaDeletePin(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking ikke funnet');return}
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  if(!room){alert('Rom ikke funnet');return}

  const deviceId=room.Tuya_Device_ID;
  const passwordId=b.Tuya_Password_ID;

  if(!deviceId){
    alert('Rom «'+(room.Title||'?')+'» har ingen Tuya_Device_ID.');return;
  }
  if(!passwordId){
    alert('Ingen aktiv PIN registrert for denne bookingen.\n\nSjekk Bookings.Tuya_Password_ID i SharePoint.');return;
  }

  if(!confirm(
    '🗑️ Slett PIN fra lås?\n\n'
    +'Gjest: '+(b.Person_Name||'?')+'\n'
    +'Rom: '+(room.Title||'?')+'\n'
    +'Password ID: '+passwordId+'\n\n'
    +'Adgang blir fjernet umiddelbart.'
  ))return;

  const btn=document.getElementById('btnTuyaDelete_'+bookingId);
  if(btn){btn.disabled=true;btn.textContent='⏳ Sletter...'}
  try{
    await _tuyaPost('delete_pin',{
      device_id: deviceId,
      password_id: parseInt(passwordId,10),
    });

    // Nullstill i SharePoint
    await updateListItem('Bookings',bookingId,{Tuya_Password_ID:null});
    b.Tuya_Password_ID=null;

    _toast('✓ PIN slettet fra låsen');
    if(typeof showBookingDetail==='function')showBookingDetail(bookingId);
  }catch(e){
    alert('❌ Tuya feilet:\n\n'+e.message);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='🗑️ Slett PIN fra lås'}
  }
}

// ---- LIST AKTIVE PINS ----
async function tuyaListPins(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking ikke funnet');return}
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  if(!room||!room.Tuya_Device_ID){alert('Ingen Tuya_Device_ID på rommet');return}

  const btn=document.getElementById('btnTuyaList_'+bookingId);
  if(btn){btn.disabled=true;btn.textContent='⏳ Henter...'}
  try{
    const pins=await _tuyaGet('list_pins',{device_id:room.Tuya_Device_ID});
    showTuyaPinList(room, pins, bookingId);
  }catch(e){
    alert('❌ Klarte ikke hente PIN-liste:\n\n'+e.message);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='📋 List PINs på lås'}
  }
}

// ---- VISNINGSMODAL: ENKELT PIN ----
function showTuyaPinDisplay(room, pin, passwordId, bookingId){
  let modal=document.getElementById('tuyaPinModal');
  if(!modal){
    modal=document.createElement('div');
    modal.id='tuyaPinModal';
    modal.className='modal-overlay';
    modal.innerHTML='<div class="modal" style="max-width:480px"><div class="modal-header"><h2>🔑 PIN aktivert</h2><button onclick="document.getElementById(\'tuyaPinModal\').classList.remove(\'open\')" style="padding:5px 14px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:14px;font-family:inherit;background:var(--bg-secondary);cursor:pointer;font-weight:500">✕ Lukk</button></div><div class="modal-body"><div id="tuyaPinModalBody"></div></div></div>';
    document.body.appendChild(modal);
  }
  const b=allBookings.find(x=>x.id===bookingId);
  const guestName=b?(b.Person_Name||'Gjest'):'';
  document.getElementById('tuyaPinModalBody').innerHTML=
    '<div style="text-align:center;padding:20px 0">'
    +(guestName?'<div style="font-size:16px;font-weight:600;color:var(--text-primary);margin-bottom:4px">'+escapeHtml(guestName)+'</div>':'')
    +'<div style="font-size:13px;color:var(--text-tertiary);margin-bottom:4px">Rom: '+escapeHtml(room.Title||'')+'</div>'
    +'<div style="font-size:64px;font-weight:700;letter-spacing:8px;color:var(--text-primary);margin:16px 0;font-family:Consolas,monospace;background:var(--bg-secondary);padding:20px;border-radius:8px;user-select:all;cursor:text">'+escapeHtml(pin)+'</div>'
    +'<div style="font-size:12px;color:var(--text-tertiary);margin-bottom:16px">Tuya password ID: '+escapeHtml(String(passwordId))+'</div>'
    +'<div style="display:flex;gap:8px;justify-content:center">'
    +'<button onclick="navigator.clipboard.writeText(\''+pin+'\').then(()=>_toast(\'✓ PIN kopiert\'))" style="padding:8px 16px;background:var(--accent);color:#fff;border:none;border-radius:var(--radius-md);cursor:pointer;font-size:14px">📋 Kopier PIN</button>'
    +'</div>'
    +'<div style="margin-top:18px;padding:10px 12px;background:var(--bg-success);border-left:3px solid var(--accent);border-radius:6px;text-align:left;font-size:12px;color:var(--text-success)">'
    +'✓ PIN er aktivert direkte på låsen via Tuya API.'
    +'</div>'
    +'</div>';
  modal.classList.add('open');
}

// ---- VISNINGSMODAL: LISTE OVER PINS ----
function showTuyaPinList(room, pins, bookingId){
  let modal=document.getElementById('tuyaPinListModal');
  if(!modal){
    modal=document.createElement('div');
    modal.id='tuyaPinListModal';
    modal.className='modal-overlay';
    modal.innerHTML='<div class="modal" style="max-width:600px"><div class="modal-header"><h2>📋 Aktive PINs</h2><button onclick="document.getElementById(\'tuyaPinListModal\').classList.remove(\'open\')" style="padding:5px 14px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:14px;font-family:inherit;background:var(--bg-secondary);cursor:pointer;font-weight:500">✕ Lukk</button></div><div class="modal-body"><div id="tuyaPinListBody"></div></div></div>';
    document.body.appendChild(modal);
  }
  const fmt=ts=>new Date(ts*1000).toLocaleDateString('nb-NO',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
  let rows='';
  if(!pins||!pins.length){
    rows='<div style="color:var(--text-tertiary);padding:16px 0">Ingen aktive PINs på denne låsen.</div>';
  }else{
    rows='<table style="width:100%;font-size:12px;border-collapse:collapse">'
      +'<thead><tr style="border-bottom:1px solid var(--border-primary)"><th style="text-align:left;padding:6px 8px">Navn</th><th style="padding:6px 8px">Fra</th><th style="padding:6px 8px">Til</th><th style="padding:6px 8px">Status</th><th style="padding:6px 8px">ID</th></tr></thead><tbody>';
    for(const p of pins){
      const status=p.phase===2?'<span style="color:var(--text-success)">Aktiv</span>':'<span style="color:var(--text-tertiary)">Phase '+p.phase+'</span>';
      rows+='<tr style="border-bottom:1px solid var(--border-secondary)">'
        +'<td style="padding:6px 8px">'+escapeHtml(p.name||'')+'</td>'
        +'<td style="padding:6px 8px;text-align:center">'+fmt(p.effective_time)+'</td>'
        +'<td style="padding:6px 8px;text-align:center">'+fmt(p.invalid_time)+'</td>'
        +'<td style="padding:6px 8px;text-align:center">'+status+'</td>'
        +'<td style="padding:6px 8px;text-align:center;color:var(--text-tertiary)">'+p.id+'</td>'
        +'</tr>';
    }
    rows+='</tbody></table>';
  }
  document.getElementById('tuyaPinListBody').innerHTML=
    '<div style="margin-bottom:8px;font-size:13px;color:var(--text-tertiary)">Lås: '+escapeHtml(room.Title||'')+(room.Tuya_Device_ID?' ('+room.Tuya_Device_ID+')':'')+'</div>'
    +rows;
  modal.classList.add('open');
}

// ============================================================
// BATTERY DISPLAY (v14.5.10)
// ============================================================
function _formatRelativeTime(iso){
  if(!iso)return '(ukjent)';
  const d=new Date(iso);
  if(isNaN(d.getTime()))return '(ukjent)';
  const now=new Date();
  const diffMs=now-d;
  const diffMin=Math.floor(diffMs/60000);
  const diffHr=Math.floor(diffMs/3600000);
  const diffDay=Math.floor(diffMs/86400000);
  if(diffMs<0)return formatDate(d);// future, just show date
  if(diffMin<1)return 'akkurat nå';
  if(diffMin<60)return 'for '+diffMin+' min siden';
  if(diffHr<24)return 'for '+diffHr+' time'+(diffHr===1?'':'r')+' siden';
  if(diffDay<7)return 'for '+diffDay+' dag'+(diffDay===1?'':'er')+' siden';
  return formatDate(d);
}

function _formatExactDateTime(iso){
  if(!iso)return '';
  const d=new Date(iso);
  if(isNaN(d.getTime()))return '';
  const dd=String(d.getDate()).padStart(2,'0');
  const mm=String(d.getMonth()+1).padStart(2,'0');
  const yyyy=d.getFullYear();
  const hh=String(d.getHours()).padStart(2,'0');
  const mi=String(d.getMinutes()).padStart(2,'0');
  return dd+'.'+mm+'.'+yyyy+' '+hh+':'+mi;
}

function renderBatteryStatusHtml(room){
  if(!room)return '';
  const bat=room.Door_Battery_Level;
  if(bat==null)return '<div style="font-size:12px;color:var(--text-tertiary);margin-top:6px">🔋 Batteri: <em>ikke registrert</em></div>';
  const updated=room.Door_Battery_Updated;
  // Color coding
  let color,bg;
  if(bat>=50){color='#1D9E75';bg='rgba(29,158,117,.08)'}
  else if(bat>=30){color='#EF9F27';bg='rgba(239,159,39,.10)'}
  else{color='#D14343';bg='rgba(209,67,67,.10)'}
  const icon=bat>=80?'🔋':bat>=30?'🔋':'🪫';
  const relTime=updated?_formatRelativeTime(updated):'';
  const exactTime=updated?_formatExactDateTime(updated):'';
  const timeHtml=updated
    ?'<span style="color:var(--text-tertiary);font-size:11px" title="'+exactTime+'"> · sist oppdatert '+relTime+'</span>'
    :'<span style="color:var(--text-tertiary);font-size:11px"> · timestamp mangler</span>';
  return '<div style="display:inline-block;padding:4px 10px;border-radius:6px;background:'+bg+';color:'+color+';font-size:12px;margin-top:6px;font-weight:500">'
    +icon+' '+bat+'%'+timeHtml
    +'</div>';
}

function showLowBatteryAlert(){
  const lowRooms=allRooms.filter(r=>r.Door_Battery_Level!=null&&Number(r.Door_Battery_Level)<30);
  if(!lowRooms.length)return;
  // Build modal once, reuse
  let modal=document.getElementById('lowBatteryModal');
  if(!modal){
    modal=document.createElement('div');
    modal.id='lowBatteryModal';
    modal.className='modal-overlay';
    modal.innerHTML='<div class="modal" style="max-width:600px"><div class="modal-header"><h2>🪫 Lavt batteri</h2><button onclick="document.getElementById(\'lowBatteryModal\').classList.remove(\'open\')" style="padding:5px 14px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:14px;font-family:inherit;background:var(--bg-secondary);cursor:pointer;font-weight:500">✕ Lukk</button></div><div class="modal-body"><div id="lowBatteryBody"></div></div></div>';
    document.body.appendChild(modal);
  }
  // Sort by battery ascending (lowest first)
  lowRooms.sort((a,b)=>Number(a.Door_Battery_Level)-Number(b.Door_Battery_Level));
  let html='<div style="background:rgba(209,67,67,.08);border-left:3px solid #D14343;padding:10px 14px;margin-bottom:14px;font-size:13px;border-radius:6px">'
    +'<strong>'+lowRooms.length+' lås'+(lowRooms.length===1?'':'er')+' har batteri under 30%</strong> — bytt batteri snart for å unngå at gjester blir låst ute.'
    +'</div>';
  html+='<div style="max-height:400px;overflow-y:auto">';
  lowRooms.forEach(r=>{
    const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
    const bat=Number(r.Door_Battery_Level);
    const color=bat<15?'#D14343':'#EF9F27';
    const icon=bat<15?'🪫':'🔋';
    html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 10px;border-bottom:.5px solid var(--border-tertiary)">'
      +'<div><strong>'+escapeHtml(r.Title||'')+'</strong> <span style="color:var(--text-tertiary);font-size:11px">'+escapeHtml(prop?prop.Title:'')+'</span></div>'
      +'<div style="color:'+color+';font-weight:600;font-size:14px">'+icon+' '+bat+'%</div>'
      +'</div>';
  });
  html+='</div>';
  document.getElementById('lowBatteryBody').innerHTML=html;
  modal.classList.add('open');
}

// Manual trigger from menu
function showLowBatteryStatus(){
  const lowRooms=allRooms.filter(r=>r.Door_Battery_Level!=null&&Number(r.Door_Battery_Level)<30);
  if(!lowRooms.length){
    _toast('✓ Ingen låser med lavt batteri');
    return;
  }
  showLowBatteryAlert();
}

// ============================================================
// MANUAL DOOR CODE (used when room has no Tuya_Device_ID — render.js calls these)
// Ported from v14.6.0 monolith (modules.js:4725-4793).
// ============================================================
async function generateRoomDoorCode(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking ikke funnet');return}
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  if(!room){alert('Rom ikke funnet');return}
  let newCode=_generate6DigitCode();
  let attempts=0;
  while(newCode===room.Door_Code&&attempts<5){
    newCode=_generate6DigitCode();
    attempts++;
  }
  const oldCode=room.Door_Code||'(ingen)';
  if(!confirm('Generer ny dørkode for '+(room.Title||'rom')+'?\n\nGammel kode: '+oldCode+'\nNy kode: '+newCode+'\n\nObs: Du må også oppdatere koden i Tuya-appen manuelt.'))return;
  try{
    await updateListItem('Rooms',room.id,{Door_Code:newCode});
    room.Door_Code=newCode;
    _toast('✓ Ny dørkode lagret: '+newCode);
    if(typeof showBookingDetail==='function'&&typeof currentBookingDetailId!=='undefined'&&currentBookingDetailId===bookingId){
      showBookingDetail(bookingId);
    }
    showDoorCodeDisplay(room,newCode);
  }catch(e){
    alert('Lagring feilet: '+e.message);
  }
}

function showRoomDoorCode(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);
  if(!b){alert('Booking ikke funnet');return}
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));
  if(!room){alert('Rom ikke funnet');return}
  if(!room.Door_Code){
    if(confirm('Ingen kode satt for '+(room.Title||'rom')+'.\n\nGenerer ny nå?')){
      generateRoomDoorCode(bookingId);
    }
    return;
  }
  showDoorCodeDisplay(room,room.Door_Code);
}

function showDoorCodeDisplay(room,code){
  let modal=document.getElementById('doorCodeDisplayModal');
  if(!modal){
    modal=document.createElement('div');
    modal.id='doorCodeDisplayModal';
    modal.className='modal-overlay';
    modal.innerHTML='<div class="modal" style="max-width:480px"><div class="modal-header"><h2>🔑 Dørkode</h2><button onclick="document.getElementById(\'doorCodeDisplayModal\').classList.remove(\'open\')" style="padding:5px 14px;border:1px solid var(--border-tertiary);border-radius:var(--radius-md);font-size:14px;font-family:inherit;background:var(--bg-secondary);cursor:pointer;font-weight:500">✕ Lukk</button></div><div class="modal-body"><div id="doorCodeDisplayBody"></div></div></div>';
    document.body.appendChild(modal);
  }
  const body=document.getElementById('doorCodeDisplayBody');
  body.innerHTML=
    '<div style="text-align:center;padding:20px 0">'
    +'<div style="font-size:14px;color:var(--text-tertiary);margin-bottom:8px">Rom: '+escapeHtml(room.Title||'')+'</div>'
    +'<div id="doorCodeBig" style="font-size:64px;font-weight:700;letter-spacing:8px;color:var(--text-primary);margin:20px 0;font-family:Consolas,monospace;background:var(--bg-secondary);padding:24px;border-radius:8px;user-select:all;cursor:text">'+code+'</div>'
    +'<div style="display:flex;gap:8px;justify-content:center;margin-top:14px">'
    +'<button onclick="navigator.clipboard.writeText(\''+code+'\').then(()=>_toast(\'✓ Kode kopiert\'))" style="padding:8px 16px;background:var(--accent);color:#fff;border:none;border-radius:var(--radius-md);cursor:pointer;font-size:14px">📋 Kopier kode</button>'
    +'<button onclick="generateRoomDoorCode(\''+(window._currentDoorCodeBookingId||'')+'\');document.getElementById(\'doorCodeDisplayModal\').classList.remove(\'open\')" style="padding:8px 16px;background:#EF9F27;color:#fff;border:none;border-radius:var(--radius-md);cursor:pointer;font-size:14px">🔄 Generer ny</button>'
    +'</div>'
    +'<div style="margin-top:20px;padding:12px;background:rgba(239,159,39,.08);border-left:3px solid #EF9F27;border-radius:6px;text-align:left;font-size:12px">'
    +'<strong>⚠ Husk:</strong> Du må også oppdatere koden i Tuya-appen manuelt. Tuya-integrasjon kommer i fase 2.'
    +'</div>'
    +'</div>';
  modal.classList.add('open');
}
