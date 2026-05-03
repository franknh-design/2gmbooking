// ============================================================
// 2GM Booking v14.7.0 — render.js
// Etasjevisning, booking-detaljpanel, statistikk, hjelpere
// ============================================================

function doorTagBtn(b){
  if(!b)return'<button class="status-btn" disabled></button>';
  const s=b.Door_Tag_Status||'None';
  if(s==='Needs-print')return'<button class="status-btn needs-print" onclick="cycleDT(event,\''+b.id+'\')">✕</button>';
  if(s==='Printed')return'<button class="status-btn printed" onclick="cycleDT(event,\''+b.id+'\')">✓</button>';
  return'<button class="status-btn" onclick="cycleDT(event,\''+b.id+'\')"></button>';
}
function cleanBtn(b,room){
  // v14.5.10: Support cleaning status on empty rooms via Room.Cleaning_Status
  if(!b){
    if(!room)return'<button class="clean-btn" disabled></button>';
    const rs=room.Cleaning_Status||'None';
    if(rs==='Dirty')return'<button class="clean-btn dirty" onclick="cycleRoomCS(event,\''+room.id+'\')" title="Empty room — cleaning status"></button>';
    if(rs==='Clean')return'<button class="clean-btn clean" onclick="cycleRoomCS(event,\''+room.id+'\')" title="Empty room — cleaning status"></button>';
    return'<button class="clean-btn" onclick="cycleRoomCS(event,\''+room.id+'\')" title="Empty room — cleaning status"></button>';
  }
  const s=b.Cleaning_Status||'None';
  if(s==='Dirty')return'<button class="clean-btn dirty" onclick="cycleCS(event,\''+b.id+'\')"></button>';
  if(s==='Clean')return'<button class="clean-btn clean" onclick="cycleCS(event,\''+b.id+'\')"></button>';
  return'<button class="clean-btn" onclick="cycleCS(event,\''+b.id+'\')"></button>';
}
function batCell(l){if(l==null)return'<span class="muted">—</span>';if(l<30)return'<span class="pill danger">'+l+'%</span>';if(l<60)return'<span class="pill warning">'+l+'%</span>';return'<span>'+l+'%</span>'}
function datesCell(b){
  if(!b)return'<span class="empty-text">Empty</span>';
  const ci=formatDate(b.Check_In);const co=b.Check_Out?formatDate(b.Check_Out):'Open-ended';
  const today=new Date();today.setHours(0,0,0,0);const ind=new Date(b.Check_In);ind.setHours(0,0,0,0);
  const days=Math.round((ind-today)/864e5);let s='';
  if(b.Status==='Upcoming'||days>0){if(days>=0&&days<=4)s='color:var(--accent);font-weight:500;';else if(days>4&&days<=30)s='color:#EF9F27;font-weight:500;'}
  // v14.5.10: overdue badges
  let overdueBadge='';
  if(isOverdueCheckIn(b)){
    const d=daysOverdueCheckIn(b);
    overdueBadge=' <span style="background:rgba(209,67,67,.12);color:#A32D2D;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500" title="Should have been checked in '+d+' day'+(d===1?'':'s')+' ago">⚠ Check-in '+d+'d</span>';
  }else if(isOverdueCheckOut(b)){
    const d=daysOverdueCheckOut(b);
    overdueBadge=' <span style="background:rgba(209,67,67,.12);color:#A32D2D;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500" title="Should have been checked out '+d+' day'+(d===1?'':'s')+' ago">⚠ Check-out '+d+'d</span>';
  }
  // v14.5.18: needs-attention badge (orange) — adds on top of overdue badge if both apply
  const att=bookingNeedsAttention(b);
  if(att){
    const tipText=att.type==='invalid_status'
      ?'Status is '+(b.Status||'?')+' but Check-out was '+att.daysSinceCheckOut+' day'+(att.daysSinceCheckOut===1?'':'s')+' ago'
      :'Status Upcoming but Check-in was '+att.daysSinceCheckIn+' days ago';
    overdueBadge+=' <span style="background:rgba(239,159,39,.15);color:#854F0B;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:500" title="'+tipText+'">⚠ '+att.label+'</span>';
  }
  return'<span style="'+s+'">'+ci+'</span> — '+co+overdueBadge;
}

// Returns the full-tenant company for a property if active on the given date (default today), else null
function getActiveFullTenant(property,date){
  if(!property)return null;
  const company=(property.FullTenant_Company||'').trim();
  if(!company)return null;
  const rate=Number(property.FullTenant_RatePerRoom)||0;
  if(!rate)return null;
  const checkDate=date||new Date();
  const start=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):null;
  const end=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):null;
  if(start&&checkDate<start)return null;
  if(end&&checkDate>end)return null;
  return{company,rate,start,end};
}

// Checks if the property containing a given room is on full-tenant lease today
function getRoomFullTenant(room,date){
  if(!room)return null;
  const prop=properties.find(p=>String(p.id)===String(room.PropertyLookupId));
  return prop?getActiveFullTenant(prop,date):null;
}

// Returns active long-term contract for a single room on the given date, or null.
// Long-term = a fixed contract on a specific room (e.g. SalMar leases Leilighet 1A monthly).
// Distinct from full-tenant which applies to ALL rooms on a property uniformly.
function getActiveLongTermContract(room,date){
  if(!room)return null;
  const company=(room.LongTerm_Company||'').trim();
  if(!company)return null;
  const price=Number(room.LongTerm_Price)||0;
  if(!price)return null;
  const checkDate=date||new Date();
  const start=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):null;
  const end=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):null;
  if(start&&checkDate<start)return null;
  if(end&&checkDate>end)return null;
  const rateUnitRaw=(room.LongTerm_RateUnit||'Per day').toString().toLowerCase().trim();
  const isMonthly=rateUnitRaw.indexOf('month')>=0;
  return{company,price,start,end,isMonthly};
}

// Computes long-term contract amount for a single room over a period (handles pro-rata).
function computeLongTermForRoomPeriod(room,fromDate,toDate){
  const c=getActiveLongTermContract(room);
  if(!c)return null;
  // Recompute date overlap with agreement bounds (use start/end from contract)
  const agreementStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):new Date(1970,0,1);
  const agreementEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  const days=Math.floor((effTo-effFrom)/86400000)+1;
  let total,unitLabel,detailLabel;
  if(c.isMonthly){
    let monthFraction=0;
    let cursor=new Date(effFrom.getFullYear(),effFrom.getMonth(),1);
    while(cursor<=effTo){
      const monthStart=new Date(cursor.getFullYear(),cursor.getMonth(),1);
      const monthEnd=new Date(cursor.getFullYear(),cursor.getMonth()+1,0,23,59,59);
      const periodInMonthStart=new Date(Math.max(monthStart.getTime(),effFrom.getTime()));
      const periodInMonthEnd=new Date(Math.min(monthEnd.getTime(),effTo.getTime()));
      if(periodInMonthStart<=periodInMonthEnd){
        const daysInMonth=monthEnd.getDate();
        const periodDaysInMonth=Math.floor((periodInMonthEnd-periodInMonthStart)/86400000)+1;
        monthFraction+=periodDaysInMonth/daysInMonth;
      }
      cursor=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
    }
    total=Math.round(c.price*monthFraction*100)/100;
    detailLabel=c.price.toLocaleString('nb-NO')+' kr/mnd × '+monthFraction.toFixed(3)+' mnd';
  }else{
    total=Math.round(c.price*days*100)/100;
    detailLabel=c.price.toLocaleString('nb-NO')+' kr/dag × '+days+' dager';
  }
  return{room,company:c.company,price:c.price,isMonthly:c.isMonthly,days,total,detailLabel};
}

// Splits a long-term room's invoicing period into segments based on actual bookings
// in the room. Each segment is either a guest stay or an empty period (where the
// company still pays). Total of all segments matches the room's contract amount,
// with rounding adjustment on the last segment.
//
// Returns array of: {fromDate, toDate, days, name, isEmpty, total, dailyRate}
function segmentLongTermRoom(room,fromDate,toDate){
  const c=getActiveLongTermContract(room);
  if(!c)return null;
  // Determine the effective period (overlap of contract and invoicing period)
  const agreementStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):new Date(1970,0,1);
  const agreementEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  const totalDays=Math.floor((effTo-effFrom)/86400000)+1;
  // Compute the total contract amount for this period (with pro-rata)
  const contractCalc=computeLongTermForRoomPeriod(room,fromDate,toDate);
  if(!contractCalc)return null;
  const contractTotal=contractCalc.total;
  const dailyRate=contractTotal/totalDays;
  // Find all bookings on this room that overlap the period
  const roomBookings=allBookings.filter(b=>{
    if(String(b.RoomLookupId)!==String(room.id))return false;
    if(b.Status==='Cancelled')return false;
    if(!b.Check_In)return false;
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):effTo;co.setHours(0,0,0,0);
    if(co<effFrom||ci>effTo)return false;
    return true;
  }).map(b=>{
    const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
    const co=b.Check_Out?new Date(b.Check_Out):effTo;co.setHours(0,0,0,0);
    // Clip to effective period
    const sFrom=ci<effFrom?effFrom:ci;
    const sTo=co>effTo?effTo:co;
    return {bookingId:b.id,name:b.Person_Name||'(uten navn)',from:sFrom,to:sTo};
  }).sort((a,b)=>a.from-b.from);
  // Build segments: walk from effFrom forward, alternate between booking and empty
  const segments=[];
  let cursor=new Date(effFrom);
  for(let i=0;i<roomBookings.length;i++){
    const bk=roomBookings[i];
    // Empty segment before this booking?
    if(bk.from>cursor){
      const segTo=new Date(bk.from.getTime()-86400000);// day before booking starts
      segTo.setHours(0,0,0,0);
      if(segTo>=cursor){
        const days=Math.floor((segTo-cursor)/86400000)+1;
        segments.push({fromDate:new Date(cursor),toDate:segTo,days,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
      }
    }
    // The booking segment
    const bSegFrom=bk.from>cursor?bk.from:cursor;
    const bSegTo=bk.to;
    if(bSegTo>=bSegFrom){
      const days=Math.floor((bSegTo-bSegFrom)/86400000)+1;
      segments.push({fromDate:new Date(bSegFrom),toDate:new Date(bSegTo),days,name:bk.name,isEmpty:false,bookingId:bk.bookingId});
    }
    // Move cursor to day after booking ends
    const nextCursor=new Date(bSegTo.getTime()+86400000);
    nextCursor.setHours(0,0,0,0);
    if(nextCursor>cursor)cursor=nextCursor;
  }
  // Trailing empty segment after last booking
  if(cursor<=effTo){
    const days=Math.floor((effTo-cursor)/86400000)+1;
    segments.push({fromDate:new Date(cursor),toDate:new Date(effTo),days,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
  }
  // No bookings at all → one big empty segment
  if(segments.length===0){
    segments.push({fromDate:new Date(effFrom),toDate:new Date(effTo),days:totalDays,name:c.company+' (tomt)',isEmpty:true,bookingId:null});
  }
  // Compute totals per segment, with rounding adjustment on last segment
  let runningTotal=0;
  segments.forEach((s,idx)=>{
    if(idx===segments.length-1){
      // Last segment gets the rounding diff
      s.total=Math.round((contractTotal-runningTotal)*100)/100;
    }else{
      s.total=Math.round(s.days*dailyRate*100)/100;
      runningTotal+=s.total;
    }
    s.dailyRate=dailyRate;
  });
  return{
    room,company:c.company,price:c.price,isMonthly:c.isMonthly,
    contractTotal,totalDays,dailyRate,segments,
    contractDetailLabel:contractCalc.detailLabel
  };
}

// Compute full-tenant lease amount for a property within a date period.
// Handles pro-rata (partial overlap between period and agreement dates).
// Returns {days,rooms,rate,total,company,effectiveFrom,effectiveTo} or null if not applicable.
function computeFullTenantForPeriod(property,fromDate,toDate){
  if(!property)return null;
  const company=(property.FullTenant_Company||'').trim();
  if(!company)return null;
  const propertyRate=Number(property.FullTenant_RatePerRoom)||0;
  // Rate unit: 'Per day' (default, legacy) or 'Per month'
  const rateUnitRaw=(property.FullTenant_RateUnit||'Per day').toString().toLowerCase().trim();
  const isMonthly=rateUnitRaw.indexOf('month')>=0;
  const agreementStart=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):new Date(1970,0,1);
  const agreementEnd=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):new Date(2100,0,1);
  agreementStart.setHours(0,0,0,0);
  agreementEnd.setHours(23,59,59,999);
  // Effective overlap between period and agreement
  const effFrom=new Date(Math.max(fromDate.getTime(),agreementStart.getTime()));
  const effTo=new Date(Math.min(toDate.getTime(),agreementEnd.getTime()));
  effFrom.setHours(0,0,0,0);
  effTo.setHours(23,59,59,999);
  if(effFrom>effTo)return null;
  // Count days (inclusive)
  const days=Math.floor((effTo-effFrom)/86400000)+1;
  // Rooms on this property
  const propRooms=allRooms.filter(r=>String(r.PropertyLookupId)===String(property.id));
  const rooms=propRooms.length;
  if(rooms===0)return null;
  // Two pricing models:
  // A) Property has FullTenant_RatePerRoom set → uniform rate × rooms (Rigg 44 style)
  // B) Property rate is empty → sum each room's FullTenant_RoomPrice (Strandveien style, per-room)
  const useUniformRate=propertyRate>0;
  const sumRoomFullTenantPrices=propRooms.reduce((s,r)=>s+(Number(r.FullTenant_RoomPrice)||0),0);
  const usePerRoomRates=!useUniformRate&&sumRoomFullTenantPrices>0;
  if(!useUniformRate&&!usePerRoomRates)return null;
  let total,unitLabel,detailLabel,rate;
  // Pre-compute month fraction for both monthly modes
  let monthFraction=0;
  const breakdown=[];
  if(isMonthly){
    let cursor=new Date(effFrom.getFullYear(),effFrom.getMonth(),1);
    while(cursor<=effTo){
      const monthStart=new Date(cursor.getFullYear(),cursor.getMonth(),1);
      const monthEnd=new Date(cursor.getFullYear(),cursor.getMonth()+1,0,23,59,59);
      const periodInMonthStart=new Date(Math.max(monthStart.getTime(),effFrom.getTime()));
      const periodInMonthEnd=new Date(Math.min(monthEnd.getTime(),effTo.getTime()));
      if(periodInMonthStart<=periodInMonthEnd){
        const daysInMonth=monthEnd.getDate();
        const periodDaysInMonth=Math.floor((periodInMonthEnd-periodInMonthStart)/86400000)+1;
        monthFraction+=periodDaysInMonth/daysInMonth;
        breakdown.push(periodDaysInMonth+'/'+daysInMonth);
      }
      cursor=new Date(cursor.getFullYear(),cursor.getMonth()+1,1);
    }
  }
  if(useUniformRate){
    rate=propertyRate;
    if(isMonthly){
      total=Math.round(rate*rooms*monthFraction*100)/100;
      unitLabel='/mnd';
      detailLabel=rooms+' rom × '+rate.toLocaleString('nb-NO')+' kr/mnd × '+monthFraction.toFixed(3)+' mnd ('+breakdown.join(' + ')+')';
    }else{
      total=Math.round(rate*rooms*days*100)/100;
      unitLabel='/dag';
      detailLabel=rooms+' rom × '+rate.toLocaleString('nb-NO')+' kr/dag × '+days+' dager';
    }
  }else{
    // Per-room pricing model: sum each room's FullTenant_RoomPrice
    rate=sumRoomFullTenantPrices; // total per unit (day or month) for all rooms combined
    const roomsWithRate=propRooms.filter(r=>Number(r.FullTenant_RoomPrice)>0).length;
    if(isMonthly){
      total=Math.round(sumRoomFullTenantPrices*monthFraction*100)/100;
      unitLabel='/mnd (per rom)';
      detailLabel='Sum '+roomsWithRate+'/'+rooms+' rom-priser ('+sumRoomFullTenantPrices.toLocaleString('nb-NO')+' kr/mnd) × '+monthFraction.toFixed(3)+' mnd';
    }else{
      total=Math.round(sumRoomFullTenantPrices*days*100)/100;
      unitLabel='/dag (per rom)';
      detailLabel='Sum '+roomsWithRate+'/'+rooms+' rom-priser ('+sumRoomFullTenantPrices.toLocaleString('nb-NO')+' kr/dag) × '+days+' dager';
    }
  }
  return{days,rooms,rate,total,company,effectiveFrom:effFrom,effectiveTo:effTo,isMonthly,unitLabel,detailLabel,usePerRoomRates};
}

function renderRow(room,booking){
  const n=booking?booking.Person_Name:'';const c=booking?(booking.Company||''):'';
  // For empty rooms: find next upcoming booking, full-tenant, or long-term contract
  let emptyCell='<span class="empty-text">—</span>';
  if(!booking){
    const fullTenant=getRoomFullTenant(room);
    const longTerm=getActiveLongTermContract(room);
    const reserveLabel=fullTenant?fullTenant.company:(longTerm?longTerm.company:null);
    if(reserveLabel){
      emptyCell='<span style="color:#EF9F27;font-style:italic">🔒 Reservert '+escapeHtml(reserveLabel)+'</span>';
    }
    const upcoming=findNextUpcomingForRoom(room.id);
    if(upcoming){
      emptyCell=(reserveLabel?'<span style="color:#EF9F27;font-style:italic">🔒 '+escapeHtml(reserveLabel)+'</span>':'<span class="empty-text">—</span>')+' <span style="font-size:10px;color:#2C7A7B;font-style:italic" title="Upcoming booking">📅 '+escapeHtml(upcoming.Person_Name||'')+(upcoming.Check_In?' · '+formatDate(upcoming.Check_In):'')+'</span>';
    }
  }
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+doorTagBtn(booking)+'</td><td>'+cleanBtn(booking,room)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+'</td>'
    +'<td>'+(n?guestMarkedName(n):emptyCell)+(booking&&booking.Notes?'<span class="note-dot"></span>':'')+'</td>'
    +'<td class="muted">'+c+'</td>'
    +'<td style="text-align:right;font-variant-numeric:tabular-nums">'+batCell(room.Door_Battery_Level)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums">'+datesCell(booking)+'</td></tr>';
}

// Find the soonest Upcoming booking for a given room
function findNextUpcomingForRoom(roomId){
  const now=new Date();now.setHours(0,0,0,0);
  const ups=allBookings.filter(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Upcoming'&&b.Check_In);
  ups.sort((a,b)=>new Date(a.Check_In)-new Date(b.Check_In));
  return ups.find(b=>{const d=new Date(b.Check_In);d.setHours(0,0,0,0);return d>=now})||null;
}

function renderRowWithProperty(room,booking,propName){
  const n=booking?booking.Person_Name:'';const c=booking?(booking.Company||''):'';
  let emptyCell='<span class="empty-text">—</span>';
  if(!booking){
    const fullTenant=getRoomFullTenant(room);
    const longTerm=getActiveLongTermContract(room);
    const reserveLabel=fullTenant?fullTenant.company:(longTerm?longTerm.company:null);
    if(reserveLabel){
      emptyCell='<span style="color:#EF9F27;font-style:italic">🔒 Reservert '+escapeHtml(reserveLabel)+'</span>';
    }
    const upcoming=findNextUpcomingForRoom(room.id);
    if(upcoming){
      emptyCell=(reserveLabel?'<span style="color:#EF9F27;font-style:italic">🔒 '+escapeHtml(reserveLabel)+'</span>':'<span class="empty-text">—</span>')+' <span style="font-size:10px;color:#2C7A7B;font-style:italic">📅 '+escapeHtml(upcoming.Person_Name||'')+(upcoming.Check_In?' · '+formatDate(upcoming.Check_In):'')+'</span>';
    }
  }
  // v14.5.10 fix: 7 columns matching the header (T, C, Room, Name, Company, Bat., Dates)
  return'<tr onclick="showDetail(\''+room.id+'\')">'
    +'<td>'+doorTagBtn(booking)+'</td><td>'+cleanBtn(booking,room)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums;font-weight:500">'+room.Title+' <span class="muted" style="font-size:10px">'+escapeHtml(propName)+'</span></td>'
    +'<td>'+(n?guestMarkedName(n):emptyCell)+(booking&&booking.Notes?'<span class="note-dot"></span>':'')+'</td>'
    +'<td class="muted">'+c+'</td>'
    +'<td style="text-align:right;font-variant-numeric:tabular-nums">'+batCell(room.Door_Battery_Level)+'</td>'
    +'<td style="font-variant-numeric:tabular-nums">'+datesCell(booking)+'</td></tr>';
}

function renderFloors(){
  const sourceBk=(activeFilter==='dirty')?allBookings:bookings;
  const bMap={};
  // v14.5.19: When needsAttention filter is active, surface the problematic booking on each row,
  // even if it's behind another Active/Upcoming booking. Otherwise the row appears under the filter
  // but shows a non-problematic booking with no warning badge — confusing UX.
  if(activeFilter==='needsAttention'){
    // First pass: any problematic booking takes priority on its room
    allBookings.forEach(b=>{
      const rid=String(b.RoomLookupId||'');
      if(!rid)return;
      if(bookingNeedsAttention(b)!==null&&!bMap[rid])bMap[rid]=b;
    });
    // Second pass: rooms without problematic booking get their normal active/upcoming
    sourceBk.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&!bMap[rid])bMap[rid]=b});
  }else{
    sourceBk.forEach(b=>{const rid=String(b.RoomLookupId||'');if(rid&&(b.Status==='Active'||b.Status==='Upcoming')&&(!bMap[rid]||b.Status==='Active'))bMap[rid]=b});
  }
  const cols=7;
  const f1=getFilteredRoomsForFloor(1).sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  const f2=getFilteredRoomsForFloor(2).sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}));
  const allF1=rooms.filter(r=>r.Floor===1||String(r.Floor)==='1');
  const allF2=rooms.filter(r=>r.Floor===2||String(r.Floor)==='2');
  const noMatch='<tr><td colspan="'+cols+'" class="loading">No matching rooms</td></tr>';

  const isAllProps=selectedProperty===null;

  const renderFn=(r)=>{
    const b=bMap[r.id];
    if(activeFilter==='dirty'||isAllProps){
      const prop=properties.find(p=>String(p.id)===String(r.PropertyLookupId));
      return renderRowWithProperty(r,b,prop?prop.Title:'');
    }
    return renderRow(r,b);
  };

  document.getElementById('floor1Body').innerHTML=f1.length?f1.map(renderFn).join(''):noMatch;
  document.getElementById('floor2Body').innerHTML=f2.length?f2.map(renderFn).join(''):noMatch;

  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut','needsAttention'].includes(activeFilter);
  if(isStatFilter||isAllProps){
    document.getElementById('floor1Sub').textContent=f1.length+' rooms — all properties';
    document.getElementById('floor2Sub').textContent=f2.length+' rooms — all properties';
  }else{
    const f1range=allF1.length?'Rooms '+allF1.sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}))[0].Title+'–'+allF1[allF1.length-1].Title:'';
    const f2range=allF2.length?'Rooms '+allF2.sort((a,b)=>(a.Title||'').localeCompare(b.Title||'',undefined,{numeric:true}))[0].Title+'–'+allF2[allF2.length-1].Title:'';
    document.getElementById('floor1Sub').textContent=activeFilter?f1.length+' of '+allF1.length+' rooms':f1range;
    document.getElementById('floor2Sub').textContent=activeFilter?f2.length+' of '+allF2.length+' rooms':f2range;
  }
}

function updateStats(){
  // v14.5.10: All stat cards count across ALL assigned properties (regardless of selected property)
  const assignedPropIds=new Set(properties.map(p=>String(p.id)));
  const allAssignedRooms=allRooms.filter(r=>assignedPropIds.has(String(r.PropertyLookupId)));
  const allAssignedRoomIds=new Set(allAssignedRooms.map(r=>r.id));
  const tr=allAssignedRooms.length;
  // Active bookings (current) across all
  const today=new Date();today.setHours(0,0,0,0);
  const occupiedRoomIds=new Set();
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!allAssignedRoomIds.has(rid))return;
    // Active bookings: Status='Active' OR (Status='Upcoming' with Check_In <= today)
    if(b.Status==='Active'){occupiedRoomIds.add(rid);return}
    if(b.Status==='Upcoming'&&b.Check_In){
      const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
      if(ci.getTime()<=today.getTime())occupiedRoomIds.add(rid);
    }
  });
  document.getElementById('statCheckedIn').textContent=occupiedRoomIds.size+' / '+tr;
  document.getElementById('statEmpty').textContent=tr-occupiedRoomIds.size;

  // v14.5.10 PERF: pre-build bookings-by-room map ONCE
  const bookingsByRoom={};
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!allAssignedRoomIds.has(rid))return;
    if(!bookingsByRoom[rid])bookingsByRoom[rid]=[];
    bookingsByRoom[rid].push(b);
  });

  // Dirty rooms — both booked dirty AND empty rooms with Cleaning_Status='Dirty'
  const allDirtyRoomIds=new Set();
  allBookings.forEach(b=>{
    const rid=String(b.RoomLookupId||'');
    if(!allAssignedRoomIds.has(rid))return;
    if(b.Cleaning_Status==='Dirty'&&(b.Status==='Active'||b.Status==='Upcoming'))allDirtyRoomIds.add(rid);
    if(b.Status==='Active'&&b.Check_In){const w=calcWashDates(b.Check_In,b.Check_Out,b.id);if(w.some(x=>x.isToday))allDirtyRoomIds.add(rid)}
  });
  // v14.5.10: empty rooms with Cleaning_Status='Dirty'
  allAssignedRooms.forEach(r=>{
    if(!occupiedRoomIds.has(r.id)&&r.Cleaning_Status==='Dirty')allDirtyRoomIds.add(r.id);
  });
  document.getElementById('statDirty').textContent=allDirtyRoomIds.size;

  document.getElementById('statDoorTag').textContent=allBookings.filter(b=>{const rid=String(b.RoomLookupId||'');return allAssignedRoomIds.has(rid)&&b.Door_Tag_Status==='Needs-print'&&(b.Status==='Active'||b.Status==='Upcoming')}).length;
  document.getElementById('statBattery').textContent=allAssignedRooms.filter(r=>r.Door_Battery_Level!=null&&r.Door_Battery_Level<30).length;

  const overdueBookings=allBookings.filter(b=>{const rid=String(b.RoomLookupId||'');return allAssignedRoomIds.has(rid)});
  const overdueCheckInCount=overdueBookings.filter(b=>isOverdueCheckIn(b)).length;
  const overdueCheckOutCount=overdueBookings.filter(b=>isOverdueCheckOut(b)).length;
  document.getElementById('statOverdueCheckIn').textContent=overdueCheckInCount;
  document.getElementById('statOverdueCheckOut').textContent=overdueCheckOutCount;
  const ciBox=document.getElementById('statOverdueCheckInBox');
  if(ciBox)ciBox.style.cssText=overdueCheckInCount>0?'background:rgba(209,67,67,.10);border-color:#D14343':'';
  const coBox=document.getElementById('statOverdueCheckOutBox');
  if(coBox)coBox.style.cssText=overdueCheckOutCount>0?'background:rgba(209,67,67,.10);border-color:#D14343':'';

  // v14.5.18: Needs-attention count (invalid status + extreme overdue) — orange tint
  const needsAttentionCount=overdueBookings.filter(b=>bookingNeedsAttention(b)!==null).length;
  const naEl=document.getElementById('statNeedsAttention');
  if(naEl)naEl.textContent=needsAttentionCount;
  const naBox=document.getElementById('statNeedsAttentionBox');
  if(naBox)naBox.style.cssText=needsAttentionCount>0?'background:rgba(239,159,39,.12);border-color:#EF9F27':'';

  // v14.5.10 PERF: Optimized occupancy with pre-processed boundaries
  const now=new Date();const curMonth=now.getMonth();const curYear=now.getFullYear();
  const todayDate=now.getDate();
  let occupiedRoomDays=0;
  allAssignedRooms.forEach(room=>{
    const property=properties.find(p=>String(p.id)===String(room.PropertyLookupId));
    // Pre-compute Full-tenant boundaries once
    let ftStart=null,ftEnd=null,hasFT=false;
    if(property){
      const ftCompany=(property.FullTenant_Company||'').trim();
      const ftRate=Number(property.FullTenant_RatePerRoom)||0;
      if(ftCompany&&ftRate>0){
        hasFT=true;
        ftStart=property.FullTenant_StartDate?new Date(property.FullTenant_StartDate):null;
        ftEnd=property.FullTenant_EndDate?new Date(property.FullTenant_EndDate):null;
        if(ftStart)ftStart.setHours(0,0,0,0);
        if(ftEnd)ftEnd.setHours(0,0,0,0);
      }
    }
    // Long-term boundaries
    let ltStart=null,ltEnd=null,hasLT=false;
    const ltCompany=(room.LongTerm_Company||'').trim();
    const ltPrice=Number(room.LongTerm_Price)||0;
    if(ltCompany&&ltPrice>0){
      hasLT=true;
      ltStart=room.LongTerm_StartDate?new Date(room.LongTerm_StartDate):null;
      ltEnd=room.LongTerm_EndDate?new Date(room.LongTerm_EndDate):null;
      if(ltStart)ltStart.setHours(0,0,0,0);
      if(ltEnd)ltEnd.setHours(0,0,0,0);
    }
    // Pre-process bookings for this room
    const roomBookings=(bookingsByRoom[room.id]||[]).filter(b=>b.Status!=='Cancelled'&&b.Check_In).map(b=>{
      const ci=new Date(b.Check_In);ci.setHours(0,0,0,0);
      const co=b.Check_Out?new Date(b.Check_Out):now;
      if(co instanceof Date)co.setHours(0,0,0,0);
      return{ci,co};
    });
    // Iterate days
    for(let d=1;d<=todayDate;d++){
      const checkDate=new Date(curYear,curMonth,d);checkDate.setHours(0,0,0,0);
      if(hasFT&&(!ftStart||checkDate>=ftStart)&&(!ftEnd||checkDate<=ftEnd)){occupiedRoomDays++;continue}
      if(hasLT&&(!ltStart||checkDate>=ltStart)&&(!ltEnd||checkDate<=ltEnd)){occupiedRoomDays++;continue}
      let occ=false;
      for(let i=0;i<roomBookings.length;i++){
        if(checkDate>=roomBookings[i].ci&&checkDate<=roomBookings[i].co){occ=true;break}
      }
      if(occ)occupiedRoomDays++;
    }
  });
  const totalPossible=tr*todayDate;
  const occPct=totalPossible>0?Math.round(occupiedRoomDays/totalPossible*100):0;
  document.getElementById('statOccupancy').textContent=occPct+'%';
}

// --- DETAIL PANEL ---
function showDetail(roomId){
  const isStatFilter=activeFilter&&['dirty','checkedIn','empty','doorTag','battery','overdueCheckIn','overdueCheckOut','needsAttention'].includes(activeFilter);
  const room=isStatFilter?allRooms.find(r=>r.id===roomId):rooms.find(r=>r.id===roomId);
  if(!room)return;
  const sourceBk=isStatFilter?allBookings:bookings;
  const booking=sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Active')
    ||sourceBk.find(b=>String(b.RoomLookupId)===roomId&&b.Status==='Upcoming');
  selectedRoom=room;selectedBooking=booking;
  const p=document.getElementById('detailPanel');
  const prop=properties.find(pr=>String(pr.id)===String(room.PropertyLookupId));
  const propName=prop?prop.Title:'';

  if(!booking){
    // Check if room has an Upcoming booking (future)
    const upcoming=findNextUpcomingForRoom(room.id);
    // v14.5.10: cleaning status for empty rooms
    const roomCl=room.Cleaning_Status||'None';
    const roomClLabel={'None':'—','Dirty':'● Needs cleaning','Clean':'● Clean'}[roomCl];
    const roomClColor=roomCl==='Dirty'?'#A32D2D':(roomCl==='Clean'?'#0F6E56':'var(--text-tertiary)');
    let subHtml='Empty — '+propName+' · <span style="color:'+roomClColor+';font-weight:500">'+roomClLabel+'</span>';
    if(upcoming){
      const ci=upcoming.Check_In?formatDate(upcoming.Check_In):'?';
      const name=upcoming.Person_Name||'(unnamed)';
      const company=upcoming.Company?' · '+escapeHtml(upcoming.Company):'';
      subHtml='<div>Empty — '+propName+' · <span style="color:'+roomClColor+';font-weight:500">'+roomClLabel+'</span></div>'
        +'<div style="margin-top:8px;padding:8px 10px;background:rgba(44,122,123,.08);border-left:3px solid #2C7A7B;border-radius:4px;font-size:12px">'
        +'📅 <strong>Upcoming booking:</strong> '+escapeHtml(name)+company+' · Check-in <strong>'+ci+'</strong>'
        +'</div>';
    }
    let actions=(can('edit_bookings')?'<button class="primary" onclick="openNewBooking(\''+room.id+'\')">Create booking</button>':'');
    if(upcoming&&can('edit_bookings')){
      actions+='<button onclick="openEditBooking(\''+upcoming.id+'\')" style="background:#2C7A7B;color:#fff;border-color:#2C7A7B">See Upcoming</button>';
    }
    // v14.5.10: Door code + messaging buttons also for Upcoming bookings on empty rooms
    if(upcoming){
      actions+='<button onclick="window._currentDoorCodeBookingId=\''+upcoming.id+'\';showRoomDoorCode(\''+upcoming.id+'\')" style="background:rgba(239,159,39,.1);color:#a76800;border-color:#EF9F27" title="Vis nåværende dørkode for rommet">🔑 Vis kode</button>';
      actions+='<button onclick="window._currentDoorCodeBookingId=\''+upcoming.id+'\';generateRoomDoorCode(\''+upcoming.id+'\')" style="background:rgba(239,159,39,.1);color:#a76800;border-color:#EF9F27" title="Generer ny 6-sifret dørkode">🔑 Generer kode</button>';
      actions+='<button onclick="copyBookingSMS(\''+upcoming.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Copy SMS to clipboard">📱 Copy SMS</button>';
      actions+='<button onclick="openBookingSMS(\''+upcoming.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Open SMS app">📱 Send SMS</button>';
      actions+='<button onclick="copyBookingEmail(\''+upcoming.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Copy email to clipboard">📧 Copy email</button>';
      actions+='<button onclick="openBookingEmail(\''+upcoming.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Open email client">📧 Send email</button>';
    }
    if(upcoming&&can('cancel_bookings')){
      actions+='<button class="danger" onclick="cancelBookingConfirmed(\''+upcoming.id+'\')">Cancel Upcoming</button>';
    }
    actions+='<button onclick="closeDetail()">Close</button>';
    p.innerHTML='<div class="detail-grid"><div class="detail-main"><div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">'+subHtml+'</div></div><div class="detail-actions">'+actions+'</div></div>';
  }else{
    const dt={'None':'—','Needs-print':'✕ Needs print','Printed':'✓ Printed'}[booking.Door_Tag_Status]||'—';
    const cl={'None':'—','Dirty':'● Needs cleaning','Clean':'● Clean'}[booking.Cleaning_Status]||'—';
    const washHtml=getWashScheduleHtml(booking);
    let infoHtml='';
    if(can('view_bookings')){
      // Look up contact info from Persons list
      const person=allPersons.find(p=>(p.Title||'').toLowerCase()===(booking.Person_Name||'').toLowerCase()
        ||(p.Name||'').toLowerCase()===(booking.Person_Name||'').toLowerCase());
      const phone=person?(person.Mobile||person.Phone||person.Telefon||''):'';
      const email=person?(person.Email||''):'';
      const addr=person?(person.Address||''):'';
      // v14.5.10: Overdue banner at top of detail
      let overdueBanner='';
      if(isOverdueCheckIn(booking)){
        const d=daysOverdueCheckIn(booking);
        overdueBanner='<div style="background:rgba(209,67,67,.10);border-left:3px solid #D14343;padding:10px 14px;margin-bottom:12px;border-radius:6px;display:flex;align-items:center;gap:10px"><div style="flex:1;font-size:13px;color:#A32D2D"><strong>⚠ Overdue check-in:</strong> This booking should have been checked in '+d+' day'+(d===1?'':'s')+' ago.</div>'+(can('checkin_out')?'<button onclick="checkIn(\''+booking.id+'\')" style="padding:6px 14px;background:#1D9E75;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-family:inherit;font-weight:500">✓ Check in now</button>':'')+'</div>';
      }else if(isOverdueCheckOut(booking)){
        const d=daysOverdueCheckOut(booking);
        overdueBanner='<div style="background:rgba(209,67,67,.10);border-left:3px solid #D14343;padding:10px 14px;margin-bottom:12px;border-radius:6px;display:flex;align-items:center;gap:10px"><div style="flex:1;font-size:13px;color:#A32D2D"><strong>⚠ Overdue check-out:</strong> This booking should have been checked out '+d+' day'+(d===1?'':'s')+' ago.</div>'+(can('checkin_out')?'<button onclick="openCheckoutModal(\''+booking.id+'\')" style="padding:6px 14px;background:#1D9E75;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-family:inherit;font-weight:500">✓ Mark as Completed</button>':'')+'</div>';
      }
      // v14.5.18: Needs-attention banner (orange) — shown in addition to overdue banner if applicable
      const naAtt=bookingNeedsAttention(booking);
      if(naAtt){
        const explanation=naAtt.type==='invalid_status'
          ?'Status is <strong>'+(booking.Status||'?')+'</strong> but Check-out was <strong>'+naAtt.daysSinceCheckOut+' day'+(naAtt.daysSinceCheckOut===1?'':'s')+' ago</strong>. The booking should probably be marked as Completed.'
          :'Status is <strong>Upcoming</strong> but Check-in was <strong>'+naAtt.daysSinceCheckIn+' days ago</strong>. This booking may have been forgotten — verify whether the guest actually stayed.';
        overdueBanner+='<div style="background:rgba(239,159,39,.12);border-left:3px solid #EF9F27;padding:10px 14px;margin-bottom:12px;border-radius:6px"><div style="font-size:13px;color:#854F0B"><strong>⚠ '+naAtt.label+':</strong> '+explanation+'</div></div>';
      }
      infoHtml=overdueBanner+'<div class="detail-name">'+booking.Person_Name+'</div>'
        +'<div class="detail-sub">Room '+room.Title+' · '+(booking.Company||'')+' · '+propName+'</div>'
        +'<table class="detail-info">'
        +(phone?'<tr><td>Mobile</td><td><a href="tel:'+phone+'" style="color:var(--accent)">'+phone+'</a></td></tr>':'')
        +(email?'<tr><td>Email</td><td><a href="mailto:'+email+'" style="color:var(--accent)">'+email+'</a></td></tr>':'')
        +(addr?'<tr><td>Address</td><td style="white-space:pre-line">'+addr+'</td></tr>':'')
        +'<tr><td>Check-in</td><td>'+formatDate(booking.Check_In)+'</td></tr>'
        +'<tr><td>Check-out</td><td>'+(booking.Check_Out?formatDate(booking.Check_Out):'Open-ended')+'</td></tr>'
        +'<tr><td>Status</td><td>'+booking.Status+'</td></tr>'
        +'<tr><td>Door tag</td><td>'+dt+'</td></tr>'
        +'<tr><td>Cleaning</td><td>'+cl+'</td></tr>'
        +(booking.Notes?'<tr><td>Notes</td><td>'+booking.Notes+'</td></tr>':'')
        +'<tr><td>Battery</td><td>'+(typeof renderBatteryStatusHtml==='function'?renderBatteryStatusHtml(room):'(n/a)')+'</td></tr>'
        +((booking.Continuation===true||booking.Continuation==='true'||booking.Continuation===1)?'<tr><td>🔗 Continuation</td><td><span style="color:#7B61FF;font-weight:500">Yes — utvask skipped</span></td></tr>':'')
        +((booking.Billing_Company||'').trim()&&(booking.Billing_Company||'').trim()!==(booking.Company||'').trim()?'<tr><td>💳 Billing</td><td><span style="color:var(--accent);font-weight:500">'+escapeHtml(booking.Billing_Company)+'</span> <span style="color:var(--text-tertiary);font-size:11px">(rate &amp; invoice follow billing company)</span></td></tr>':'')
        +'</table>'
        +(can('view_prices')?(function(){
          const cost=calcBookingCost(booking,propName);
          // Always show the pricing block, even when no rate is set, so user sees the near-miss warning
          let extra='';
          if(cost.matchedName&&cost.matchedName.toLowerCase()!==(booking.Person_Name||'').toLowerCase()){
            extra='<div style="font-size:11px;color:var(--text-tertiary);margin-top:4px">Matched rate row: "'+cost.matchedName+'"</div>';
          }
          if(cost.nearMiss){
            extra+='<div style="margin-top:6px;padding:6px 8px;background:var(--bg-warning);border:1px solid #EF9F27;border-radius:4px;color:var(--text-warning);font-size:11px">⚠ '+cost.nearMiss+'</div>';
          }
          if(!cost.rate){
            return'<div style="margin-top:10px;padding:10px;background:var(--bg-secondary);border-radius:var(--radius-md);font-size:12px">'
              +'<strong>Pricing</strong> <span class="muted">(no rate set)</span>'
              +extra+'</div>';
          }
          return'<div style="margin-top:10px;padding:10px;background:var(--bg-secondary);border-radius:var(--radius-md);font-size:12px">'
            +'<strong>Pricing</strong> <span class="muted">('+cost.source+')</span><br>'
            +'Rate: <strong>'+cost.rate+' kr/night</strong> × '+cost.nights+' nights = <strong>'+cost.total.toLocaleString('nb-NO')+' kr</strong>'
            +extra+'</div>';
        })():'')
        +washHtml;
    }else{
      infoHtml='<div class="detail-name">Room '+room.Title+'</div><div class="detail-sub">'+cl+'</div>'+(typeof renderBatteryStatusHtml==='function'?renderBatteryStatusHtml(room):'')+washHtml;
    }
    let btns='';
    if(can('edit_bookings'))btns+='<button onclick="openEditBooking(\''+booking.id+'\')">Edit booking</button>';
    // "Add to Guests" button — only if not already in Persons list (fuzzy match)
    if(can('edit_bookings')&&booking.Person_Name){
      const inList=allPersons.some(p=>{
        const pn=(p.Name||p.Title||'').toLowerCase().trim();
        const bn=(booking.Person_Name||'').toLowerCase().trim();
        if(pn===bn)return true;
        const wa=pn.split(/[\s,]+/).filter(w=>w.length>1);
        const wb=bn.split(/[\s,]+/).filter(w=>w.length>1);
        if(wa.length<2||wb.length<2)return false;
        return wa.every(w=>bn.indexOf(w)>=0)||wb.every(w=>pn.indexOf(w)>=0);
      });
      if(!inList){
        btns+='<button onclick="addBookingToGuests(\''+booking.id+'\')" style="background:var(--bg-success);color:var(--text-success);border-color:var(--accent)">+ Add to Guests</button>';
      }
    }
    if(can('print_doortag'))btns+='<button onclick="printDoorTag(\''+booking.id+'\')">Print door tag</button>';
    // Door code / Tuya lock buttons (v14.7.0)
    // Gating: read-only display (Vis PIN / Vis kode) is open to anyone who can see
    // the booking detail; create/delete/list-all requires manage_lock since those
    // mutate the lock or expose all guests' PINs at once.
    {
      const hasTuyaId=room&&room.Tuya_Device_ID;
      const hasActivePin=!!(booking.Tuya_Password_ID);
      const canManageLock=can('manage_lock');
      if(hasTuyaId){
        if(!hasActivePin){
          if(canManageLock)btns+='<button id="btnTuyaCreate_'+booking.id+'" onclick="tuyaCreatePin(\''+booking.id+'\')" style="background:rgba(29,158,117,.12);color:#1D9E75;border-color:#1D9E75" title="Opprett og aktiver PIN direkte på låsen via Tuya API">🔑 Opprett PIN på lås</button>';
        }else{
          btns+='<button onclick="showTuyaPinDisplay(allRooms.find(r=>r.id===\''+room.id+'\'),allBookings.find(b=>b.id===\''+booking.id+'\').Tuya_Password_ID||\'?\',allBookings.find(b=>b.id===\''+booking.id+'\').Tuya_Password_ID,\''+booking.id+'\')" style="background:rgba(29,158,117,.12);color:#1D9E75;border-color:#1D9E75" title="Vis nåværende aktive PIN">🔑 Vis PIN</button>';
          // Nødutgang: manuell sletting hvis checkout-automatikken har feilet
          if(canManageLock)btns+='<button id="btnTuyaDelete_'+booking.id+'" onclick="tuyaDeletePin(\''+booking.id+'\')" style="background:rgba(209,67,67,.08);color:#A32D2D;border-color:#D14343;font-size:11px" title="Slett PIN manuelt (normalt skjer dette automatisk ved Check-out)">🗑️ Slett PIN (manuelt)</button>';
        }
        if(canManageLock)btns+='<button id="btnTuyaList_'+booking.id+'" onclick="tuyaListPins(\''+booking.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="List alle aktive PINs på denne låsen">📋 List PINs</button>';
      }else{
        // Fallback: manuell kode (ingen Tuya_Device_ID på rommet)
        btns+='<button onclick="showRoomDoorCode(\''+booking.id+'\')" style="background:rgba(239,159,39,.1);color:#a76800;border-color:#EF9F27" title="Vis PIN (manuell — Tuya_Device_ID ikke satt på rom)">🔑 Vis kode (manuell)</button>';
      }
    }
    // Messaging buttons
    btns+='<button onclick="copyBookingSMS(\''+booking.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Kopier SMS-tekst til utklippstavle">📱 Kopier SMS</button>';
    btns+='<button onclick="openBookingSMS(\''+booking.id+'\')" style="background:rgba(14,165,165,.1);color:#0EA5A5;border-color:#0EA5A5" title="Åpne SMS-app med ferdig tekst">📱 Send SMS</button>';
    btns+='<button onclick="copyBookingEmail(\''+booking.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Kopier e-post-tekst til utklippstavle">📧 Kopier e-post</button>';
    btns+='<button onclick="openBookingEmail(\''+booking.id+'\')" style="background:rgba(123,97,255,.1);color:#7B61FF;border-color:#7B61FF" title="Åpne e-postklient med ferdig tekst">📧 Send e-post</button>';
    if(booking.Status==='Upcoming'&&can('checkin_out'))btns+='<button class="primary" onclick="checkIn(\''+booking.id+'\')">Check in</button>';
    if(booking.Status==='Active'&&can('checkin_out'))btns+='<button class="primary" style="background:#EF9F27;border-color:#EF9F27" onclick="checkOut(\''+booking.id+'\')">Check out</button>';
    if(can('cleaning')){
      if(booking.Cleaning_Status==='Dirty')btns+='<button class="primary" onclick="markClean(\''+booking.id+'\')">Mark as cleaned ✓</button>';
      else btns+='<button onclick="markDirty(\''+booking.id+'\')">Mark as dirty</button>';
    }
    if(can('cancel_bookings'))btns+='<button class="danger" onclick="cancelBooking(\''+booking.id+'\')">Cancel booking</button>';
    btns+='<button onclick="closeDetail()">Close</button>';
    p.innerHTML='<div class="detail-grid"><div class="detail-main">'+infoHtml+'</div><div class="detail-actions">'+btns+'</div></div>';
  }
  p.classList.add('open');
  // Scroll to bring detail panel into view (it's now above the floor tables)
  setTimeout(()=>{p.scrollIntoView({behavior:'smooth',block:'nearest'})},50);
}
function closeDetail(){document.getElementById('detailPanel').classList.remove('open');selectedRoom=null;selectedBooking=null}
