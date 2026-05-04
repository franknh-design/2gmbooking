// ============================================================
// 2GM Booking v14.7.0 — bookings.js
// Check-in, check-out, opprett/rediger booking, avbooking
// ============================================================

async function cycleDT(e,id){
  e.stopPropagation();if(!can('doortag'))return;
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const c={'None':'Needs-print','Needs-print':'Printed','Printed':'None'};const ns=c[b.Door_Tag_Status||'None'];
  try{await updateListItem('Bookings',id,{Door_Tag_Status:ns});b.Door_Tag_Status=ns;renderFloors();updateStats()}catch(er){console.error(er);alert('Failed')}
}
async function cycleCS(e,id){
  e.stopPropagation();if(!can('cleaning'))return;
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const c={'None':'Dirty','Dirty':'Clean','Clean':'None'};const ns=c[b.Cleaning_Status||'None'];
  try{await updateListItem('Bookings',id,{Cleaning_Status:ns});b.Cleaning_Status=ns;renderFloors();updateStats()}catch(er){console.error(er);alert('Failed')}
}
// v14.5.10: Cycle cleaning status on empty rooms (Room.Cleaning_Status)
async function cycleRoomCS(e,roomId){
  e.stopPropagation();if(!can('cleaning'))return;
  const r=allRooms.find(x=>x.id===roomId);if(!r)return;
  const c={'None':'Dirty','Dirty':'Clean','Clean':'None'};const ns=c[r.Cleaning_Status||'None'];
  try{await updateListItem('Rooms',roomId,{Cleaning_Status:ns});r.Cleaning_Status=ns;renderFloors();updateStats()}catch(er){console.error(er);alert('Failed: '+er.message)}
}
async function markClean(id){try{await updateListItem('Bookings',id,{Cleaning_Status:'Clean'});const l=allBookings.find(x=>x.id===id);if(l)l.Cleaning_Status='Clean';closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}}
async function markDirty(id){try{await updateListItem('Bookings',id,{Cleaning_Status:'Dirty'});const l=allBookings.find(x=>x.id===id);if(l)l.Cleaning_Status='Dirty';closeDetail();refreshLocal();loadData()}catch(e){alert('Failed')}}

// --- CHECK IN/OUT/CANCEL ---
async function checkIn(id){
  if(!confirm('Check in this guest?'))return;
  try{
    const now=new Date().toISOString();
    const b=allBookings.find(x=>x.id===id);
    const r=b?allRooms.find(x=>x.id===String(b.RoomLookupId)):null;
    const fields={Status:'Active',Check_In:now};
    // v15: Inherit room cleaning status if booking has none yet — so a clean (green) room stays green at check-in.
    if(r&&r.Cleaning_Status&&(!b.Cleaning_Status||b.Cleaning_Status==='None'))fields.Cleaning_Status=r.Cleaning_Status;
    await updateListItem('Bookings',id,fields);
    if(b){b.Status='Active';b.Check_In=now;if(fields.Cleaning_Status)b.Cleaning_Status=fields.Cleaning_Status;}
    closeDetail();refreshLocal();loadData();
  }catch(e){alert('Failed')}
}
function checkOut(id){
  const b=allBookings.find(x=>x.id===id);if(!b)return;checkoutBookingId=id;
  document.getElementById('fCheckOutDate').value=new Date().toISOString().split('T')[0];
  document.getElementById('checkoutGuestName').textContent=b.Person_Name||'Guest';
  document.getElementById('checkoutModal').classList.add('open');
}
function closeCheckoutModal(){document.getElementById('checkoutModal').classList.remove('open');checkoutBookingId=null}
async function confirmCheckout(){
  if(!checkoutBookingId)return;const dateVal=document.getElementById('fCheckOutDate').value;if(!dateVal){alert('Select a date');return}
  const btn=document.getElementById('checkoutConfirmBtn');btn.disabled=true;btn.textContent='Processing...';
  try{
    await updateListItem('Bookings',checkoutBookingId,{Status:'Completed',Cleaning_Status:'Dirty',Check_Out:dateVal+'T12:00:00Z'});
    const l=allBookings.find(x=>x.id===checkoutBookingId);
    if(l){l.Status='Completed';l.Cleaning_Status='Dirty';l.Check_Out=dateVal+'T12:00:00Z';
      // v14.5.10: copy cleaning status to room
      const r=allRooms.find(x=>x.id===String(l.RoomLookupId));
      if(r){try{await updateListItem('Rooms',r.id,{Cleaning_Status:'Dirty'});r.Cleaning_Status='Dirty'}catch(_){}}
      // v14.7.0: slett Tuya PIN ved checkout (ikke-blokkerende — checkout går igjennom selv om Tuya feiler)
      if(l.Tuya_Password_ID&&r&&r.Tuya_Device_ID){
        btn.textContent='Sletter PIN fra lås...';
        try{
          await _tuyaPost('delete_pin',{device_id:r.Tuya_Device_ID,password_id:parseInt(l.Tuya_Password_ID,10)});
          await updateListItem('Bookings',checkoutBookingId,{Tuya_Password_ID:null});
          l.Tuya_Password_ID=null;
          _toast('✓ PIN slettet fra låsen');
        }catch(tuyaErr){
          // Tuya-feil stopper ikke checkout — vis advarsel, la saksbehandler rydde manuelt
          console.warn('[Tuya] PIN-sletting feilet ved checkout:',tuyaErr);
          setTimeout(()=>alert('⚠ Check-out fullført, men Tuya PIN-sletting feilet:\n\n'+tuyaErr.message+'\n\nSlett PIN manuelt fra booking-panelet.'),300);
        }
      }
    }
    closeCheckoutModal();closeDetail();refreshLocal();loadData()
  }catch(e){alert('Failed: '+e.message)}finally{btn.disabled=false;btn.textContent='Confirm check-out'}
}
async function cancelBooking(id){
  return cancelBookingConfirmed(id);
}

// Detailed cancel confirmation — shows guest name and dates
async function cancelBookingConfirmed(id){
  const b=allBookings.find(x=>x.id===id);if(!b)return;
  const name=b.Person_Name||'(unnamed)';
  const ci=b.Check_In?formatDate(b.Check_In):'?';
  const co=b.Check_Out?formatDate(b.Check_Out):'Open-ended';
  const company=b.Company?' ('+b.Company+')':'';
  const msg='Are you sure you want to cancel this booking?\n\n'
    +'Guest: '+name+company+'\n'
    +'Check-in: '+ci+'\n'
    +'Check-out: '+co+'\n'
    +'Status: '+b.Status+'\n\n'
    +'This cannot be undone from the app — you would need to edit the booking manually to reactivate it.';
  if(!confirm(msg))return;
  try{
    await updateListItem('Bookings',id,{Status:'Cancelled'});
    const l=allBookings.find(x=>x.id===id);if(l)l.Status='Cancelled';
    closeDetail();refreshLocal();loadData();
  }catch(e){alert('Failed: '+e.message)}
}

// --- BOOKING MODAL ---
function populateRoomSelect(preselectedRoomId){
  const sel=document.getElementById('fRoom');
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  sel.innerHTML=sorted.map(r=>'<option value="'+r.id+'"'+(r.id===preselectedRoomId?' selected':'')+'>'+r.Title+' (Floor '+r.Floor+')</option>').join('');
  sel.onchange=()=>{const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:''};
  const rm=rooms.find(r=>r.id===sel.value);document.getElementById('fFloor').value=rm?rm.Floor:'';
}

// Returns id of first available room for given check-in/check-out, or null
// v15.1: Foretrekk rom med Cleaning_Status === 'Clean' eller 'None' fremfor Dirty rom.
function findFirstAvailableRoomId(checkInStr,checkOutStr){
  if(!checkInStr)return null;
  const newIn=new Date(checkInStr+'T00:00:00');newIn.setHours(0,0,0,0);
  const newOut=checkOutStr?new Date(checkOutStr+'T00:00:00'):null;
  if(newOut)newOut.setHours(0,0,0,0);
  // Sort: Dirty rom havner sist; innen samme renhetsgruppe sorteres på Title (numerisk).
  const cleanlinessRank=r=>(r.Cleaning_Status==='Dirty'?1:0);
  const sorted=[...rooms].sort((a,b)=>{
    const cr=cleanlinessRank(a)-cleanlinessRank(b);
    if(cr!==0)return cr;
    return (a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true});
  });
  for(const room of sorted){
    const hasConflict=allBookings.some(b=>{
      if(b.Status==='Cancelled'||b.Status==='Completed')return false;
      if(String(b.RoomLookupId)!==String(room.id))return false;
      if(!b.Check_In)return false; // v15.1: bookinger uten dato kan ikke kollidere
      const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);
      const bOut=b.Check_Out?new Date(b.Check_Out):null;if(bOut)bOut.setHours(0,0,0,0);
      if(!bOut)return newIn>=bIn||(newOut?newOut>bIn:true);
      if(!newOut)return bOut>newIn||bIn>=newIn;
      return newIn<bOut&&newOut>bIn;
    });
    if(!hasConflict)return room.id;
  }
  return null;
}
function openNewBooking(preselectedRoomId){
  ensureMainView();
  editingBookingId=null;document.getElementById('bookingModalTitle').textContent='New booking';
  document.getElementById('bookingSaveBtn').textContent='Create booking';
  const todayStr=toISODate(new Date());
  // If no room pre-selected, find first available for today
  let roomToSelect=preselectedRoomId||'';
  let autoSelected=false;
  if(!roomToSelect){
    const auto=findFirstAvailableRoomId(todayStr,'');
    if(auto){roomToSelect=auto;autoSelected=true}
  }
  populateRoomSelect(roomToSelect);
  document.getElementById('fName').value='';document.getElementById('fCompany').value='';
  document.getElementById('fBillingCompany').value='';
  const cw=document.getElementById('fCompanyWarn');if(cw)cw.innerHTML='';
  const bw=document.getElementById('fBillingCompanyWarn');if(bw)bw.innerHTML='';
  document.getElementById('fCheckIn').value=todayStr;document.getElementById('fCheckOut').value='';
  // Default to Active if check-in is today, Upcoming otherwise
  document.getElementById('fStatus').value='Active';
  document.getElementById('fNotes').value='';
  document.getElementById('fIncludeCheckoutFee').checked=true;
  document.getElementById('fContinuation').checked=false;
  document.getElementById('fNameInfo').innerHTML='';
  // Show auto-select hint in room-info area
  const roomInfo=document.getElementById('fRoomInfo');
  if(autoSelected&&roomToSelect){
    const r=rooms.find(rm=>rm.id===roomToSelect);
    if(r){
      roomInfo.textContent='✓ Auto-selected first available: Room '+r.Title;
      roomInfo.style.color='var(--text-success)';
    }
  }else if(!roomToSelect){
    roomInfo.textContent='';
  }else{
    roomInfo.textContent='';
  }
  document.getElementById('fOverlapWarning').style.display='none';
  attachOverlapListeners();
  attachStatusAutoSelect();
  checkBookingOverlap();
  const modal=document.getElementById('bookingModal');
  modal.classList.add('open');
  const modalContent=modal.querySelector('.modal');
  if(modalContent)modalContent.scrollTop=0;
  modal.scrollTop=0;
}
function openEditBooking(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;editingBookingId=bookingId;
  document.getElementById('bookingModalTitle').textContent='Edit booking';
  document.getElementById('bookingSaveBtn').textContent='Save changes';
  populateRoomSelect(String(b.RoomLookupId));
  document.getElementById('fName').value=b.Person_Name||'';document.getElementById('fCompany').value=b.Company||'';
  document.getElementById('fBillingCompany').value=b.Billing_Company||'';
  if(typeof checkCompanyRegistration==='function'){
    checkCompanyRegistration(b.Company||'','fCompanyWarn');
    checkCompanyRegistration(b.Billing_Company||'','fBillingCompanyWarn');
  }
  document.getElementById('fCheckIn').value=b.Check_In?toISODate(b.Check_In):'';
  document.getElementById('fCheckOut').value=b.Check_Out?toISODate(b.Check_Out):'';
  document.getElementById('fStatus').value=b.Status||'Upcoming';document.getElementById('fNotes').value=b.Notes||'';
  // Checkout fee: default true if not explicitly stored as false
  const fee=b.Include_Checkout_Fee;
  document.getElementById('fIncludeCheckoutFee').checked=(fee===undefined||fee===null||fee===true||fee==='true'||fee===1);
  const cont=b.Continuation;
  document.getElementById('fContinuation').checked=(cont===true||cont==='true'||cont===1);
  document.getElementById('fNameInfo').innerHTML='';
  document.getElementById('fOverlapWarning').style.display='none';
  attachOverlapListeners();
  checkBookingOverlap();
  const modal=document.getElementById('bookingModal');
  modal.classList.add('open');
  const modalContent=modal.querySelector('.modal');
  if(modalContent)modalContent.scrollTop=0;
  modal.scrollTop=0;
}
function closeBookingModal(){document.getElementById('bookingModal').classList.remove('open');editingBookingId=null}

// Attach change-listeners to room/date fields to check for overlap in real time
let _overlapAttached=false;
function attachOverlapListeners(){
  if(_overlapAttached)return;_overlapAttached=true;
  ['fRoom','fCheckIn','fCheckOut'].forEach(id=>{
    const el=document.getElementById(id);
    if(el)el.addEventListener('change',checkBookingOverlap);
  });
}

// Auto-set Status based on Check-in date (only for NEW bookings, not edits)
let _statusAutoAttached=false;
function attachStatusAutoSelect(){
  if(_statusAutoAttached)return;_statusAutoAttached=true;
  const ciEl=document.getElementById('fCheckIn');
  if(!ciEl)return;
  ciEl.addEventListener('change',()=>{
    if(editingBookingId)return; // don't override status when editing existing
    const val=ciEl.value;if(!val)return;
    const sel=document.getElementById('fStatus');if(!sel)return;
    const today=new Date();today.setHours(0,0,0,0);
    const picked=new Date(val+'T00:00:00');picked.setHours(0,0,0,0);
    // If check-in is today or in the past → Active. If in future → Upcoming.
    sel.value=picked<=today?'Active':'Upcoming';
  });
}

// Check if current modal values overlap with other bookings on same room
function checkBookingOverlap(){
  const warn=document.getElementById('fOverlapWarning');if(!warn)return;
  const roomId=document.getElementById('fRoom').value;
  const ciStr=document.getElementById('fCheckIn').value;
  const coStr=document.getElementById('fCheckOut').value;
  if(!roomId||!ciStr){warn.style.display='none';return}
  const ci=new Date(ciStr+'T00:00:00');const co=coStr?new Date(coStr+'T23:59:59'):null;
  // Find other bookings on this room (not this one if editing)
  const conflicts=allBookings.filter(b=>{
    if(b.id===editingBookingId)return false;
    if(String(b.RoomLookupId)!==String(roomId))return false;
    if(b.Status==='Cancelled'||b.Status==='Completed')return false;
    if(!b.Check_In)return false;
    const bi=new Date(b.Check_In);const bo=b.Check_Out?new Date(b.Check_Out):null;
    // Overlap: ci <= bo AND (co >= bi OR co is null)
    if(co){
      if(bo)return ci<=bo&&co>=bi;
      return co>=bi; // other is open-ended, overlap if our end is after its start
    }else{
      // our booking is open-ended
      if(bo)return ci<=bo;
      return true; // both open-ended = definite overlap
    }
  });
  if(!conflicts.length){warn.style.display='none';return}
  const lines=conflicts.map(c=>{
    const name=c.Person_Name||'(unnamed)';
    const period=formatDate(c.Check_In)+' → '+(c.Check_Out?formatDate(c.Check_Out):'Open');
    return '• <strong>'+escapeHtml(name)+'</strong> ('+c.Status+') · '+period;
  });
  warn.innerHTML='<strong>⚠ Double-booking warning</strong> — this room already has '+conflicts.length+' overlapping booking'+(conflicts.length!==1?'s':'')+':<br>'+lines.join('<br>');
  warn.style.display='block';
}

function findAvailableRoom(){
  const checkIn=document.getElementById('fCheckIn').value;
  const checkOut=document.getElementById('fCheckOut').value;
  const info=document.getElementById('fRoomInfo');

  if(!checkIn){info.textContent='Set check-in date first';info.style.color='var(--text-danger)';return}

  const newIn=new Date(checkIn);newIn.setHours(0,0,0,0);
  const newOut=checkOut?new Date(checkOut):null;
  if(newOut)newOut.setHours(0,0,0,0);

  // Find rooms that are NOT occupied during the given dates
  const sorted=[...rooms].sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));

  for(const room of sorted){
    const hasConflict=allBookings.some(b=>{
      if(b.Status==='Cancelled'||b.Status==='Completed')return false;
      if(String(b.RoomLookupId)!==String(room.id))return false;

      const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);
      const bOut=b.Check_Out?new Date(b.Check_Out):null;
      if(bOut)bOut.setHours(0,0,0,0);

      // Check overlap
      if(!bOut){return newIn>=bIn||(newOut?newOut>bIn:true)}
      if(!newOut){return bOut>newIn||bIn>=newIn}
      return newIn<bOut&&newOut>bIn;
    });

    if(!hasConflict){
      // Found available room — select it and trigger change event
      const sel=document.getElementById('fRoom');
      sel.value=room.id;
      const rm=rooms.find(r=>r.id===room.id);
      document.getElementById('fFloor').value=rm?rm.Floor:'';
      // Dispatch change event so overlap warning and other listeners update
      sel.dispatchEvent(new Event('change',{bubbles:true}));
      info.textContent='✓ Room '+room.Title+' (Floor '+room.Floor+') — selected';
      info.style.color='var(--text-success)';
      return;
    }
  }

  // No room available
  info.textContent='✕ No rooms available for these dates';
  info.style.color='var(--text-danger)';
}

async function saveBooking(){
  const roomId=document.getElementById('fRoom').value;
  const name=document.getElementById('fName').value.trim();
  const company=document.getElementById('fCompany').value.trim();
  const billingCompany=document.getElementById('fBillingCompany').value.trim();
  const checkIn=document.getElementById('fCheckIn').value;
  const checkOut=document.getElementById('fCheckOut').value;
  const status=document.getElementById('fStatus').value;
  const notes=document.getElementById('fNotes').value.trim();
  const room=rooms.find(r=>r.id===roomId);
  if(!name){alert('Guest name is required');return}

  // Collision check (v15.1: hopp over hvis ingen check-in dato)
  if(checkIn){
    const newIn=new Date(checkIn);newIn.setHours(0,0,0,0);
    const newOut=checkOut?new Date(checkOut):null;if(newOut)newOut.setHours(0,0,0,0);
    const conflicts=allBookings.filter(b=>{
      if(editingBookingId&&b.id===editingBookingId)return false;
      if(b.Status==='Cancelled'||b.Status==='Completed')return false;
      if(String(b.RoomLookupId)!==String(roomId))return false;
      if(!b.Check_In)return false;
      const bIn=new Date(b.Check_In);bIn.setHours(0,0,0,0);const bOut=b.Check_Out?new Date(b.Check_Out):null;if(bOut)bOut.setHours(0,0,0,0);
      if(!bOut)return newIn>=bIn||(newOut?newOut>bIn:true);
      if(!newOut)return bOut>newIn||bIn>=newIn;
      return newIn<bOut&&newOut>bIn;
    });
    if(conflicts.length>0){
      const c=conflicts[0];
      if(!confirm('Room already booked:\n'+c.Person_Name+' ('+c.Status+')\n'+formatDate(c.Check_In)+' — '+(c.Check_Out?formatDate(c.Check_Out):'Open-ended')+'\n\nContinue anyway?'))return;
    }
  }

  // Property_Name: find from room's property (works even in "All properties" mode)
  const roomProp=room?properties.find(pr=>String(pr.id)===String(room.PropertyLookupId)):null;
  const propNameForSave=roomProp?roomProp.Title:(selectedProperty?selectedProperty.Title:'');
  // v15: Inherit room cleaning status so a clean (green) room stays green when a booking is created on it.
  const inheritedCS=(room&&room.Cleaning_Status)||'None';
  const fields={Person_Name:name,Company:company,Billing_Company:billingCompany||null,Check_In:checkIn?checkIn+'T15:00:00Z':null,Status:status,Door_Tag_Status:'Needs-print',Cleaning_Status:inheritedCS,Property_Name:propNameForSave,Floor:room?room.Floor:1,Notes:notes||null};
  fields.Include_Checkout_Fee=document.getElementById('fIncludeCheckoutFee').checked;
  fields.Continuation=document.getElementById('fContinuation').checked;
  if(checkOut)fields.Check_Out=checkOut+'T12:00:00Z';else fields.Check_Out=null;
  fields.RoomLookupId=parseInt(roomId);
  const btn=document.getElementById('bookingSaveBtn');btn.disabled=true;btn.textContent='Saving...';
  try{
    if(editingBookingId){delete fields.Door_Tag_Status;delete fields.Cleaning_Status;await updateListItem('Bookings',editingBookingId,fields);const l=allBookings.find(x=>x.id===editingBookingId);if(l){Object.assign(l,fields);l.Check_Out=fields.Check_Out}closeBookingModal();closeDetail();refreshLocal();loadData()}
    else{await createListItem('Bookings',fields);closeBookingModal();closeDetail();await loadData()}
  }catch(e){alert('Failed: '+e.message)}finally{btn.disabled=false;btn.textContent=editingBookingId?'Save changes':'Create booking'}
}

// --- DOOR TAG PRINT ---
// v15.3: Bruker per-eiendom HTML-mal (Properties.DoorTag_Template) — rediger via Templates-modalen.
function printDoorTag(bookingId){
  const b=allBookings.find(x=>x.id===bookingId);if(!b)return;
  const room=allRooms.find(r=>r.id===String(b.RoomLookupId));const roomTitle=room?room.Title:'?';
  const html=_renderDoorTagHtml(b);
  const w=window.open('','_blank','width=700,height=900');
  w.document.write('<!DOCTYPE html><html><head><title>Door Tag — Room '+roomTitle+'</title></head><body style="margin:0">'+html+'</body></html>');
  w.document.close();
  setTimeout(()=>w.print(),500);
  // Mark as printed
  updateListItem('Bookings',bookingId,{Door_Tag_Status:'Printed'}).then(()=>{b.Door_Tag_Status='Printed';renderFloors();updateStats()}).catch(console.error);
}

