// ============================================================
// 2GM Booking v14.7.0 — rates.js
// Prisberegning, dagspriser, checkout-gebyrer
// ============================================================

function getBookingForRoom(roomId){
  return bookings.find(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Active')
    ||bookings.find(b=>String(b.RoomLookupId)===String(roomId)&&b.Status==='Upcoming');
}

// --- PRICING ---
function _nameMatch(a,b){
  // Exact case-insensitive
  const la=(a||'').toLowerCase().trim();const lb=(b||'').toLowerCase().trim();
  if(!la||!lb)return false;
  if(la===lb)return true;
  // All words match (regardless of order) — handles "Marek Filas" vs "Filas, Marek"
  const wa=la.split(/[\s,]+/).filter(w=>w.length>1);
  const wb=lb.split(/[\s,]+/).filter(w=>w.length>1);
  if(wa.length<2||wb.length<2)return false;
  // a contains all words of b OR b contains all words of a
  return wa.every(w=>lb.indexOf(w)>=0)||wb.every(w=>la.indexOf(w)>=0);
}

function getDailyRate(personName,company,propertyTitle,roomId){
  // Priority: 1) Person+Property  2) Person (any)  3) Company+Property  4) Company (any)  5) Room rate  6) Property default
  const pn=(personName||'').toLowerCase();
  const co=(company||'').toLowerCase();
  const pt=(propertyTitle||'').toLowerCase();

  // 1. Person + specific property (fuzzy name, exact property)
  // Exclude rates with FeeType=Checkout (those are one-time fees, not nightly)
  const isNightly=r=>(r.FeeType||'').toLowerCase()!=='checkout';
  let rate=allRates.find(r=>isNightly(r)&&_nameMatch(r.Person_Name,personName)&&(r.Property||'').toLowerCase()===pt&&r.DailyRate);
  if(rate)return{rate:rate.DailyRate,source:'Person+Property',matchedName:rate.Person_Name};

  // 2. Person any property (fuzzy)
  rate=allRates.find(r=>isNightly(r)&&_nameMatch(r.Person_Name,personName)&&!(r.Property)&&r.DailyRate);
  if(rate)return{rate:rate.DailyRate,source:'Person',matchedName:rate.Person_Name};

  // 3. Company + specific property
  if(co){
    rate=allRates.find(r=>isNightly(r)&&(r.Company||'').toLowerCase()===co&&(r.Property||'').toLowerCase()===pt&&r.DailyRate);
    if(rate)return{rate:rate.DailyRate,source:'Company+Property'};
  }

  // 4. Company any property
  if(co){
    rate=allRates.find(r=>isNightly(r)&&(r.Company||'').toLowerCase()===co&&!(r.Property)&&r.DailyRate);
    if(rate)return{rate:rate.DailyRate,source:'Company'};
  }

  // 5. Room rate
  if(roomId){
    const room=allRooms.find(r=>r.id===String(roomId));
    if(room&&room.DailyRate)return{rate:room.DailyRate,source:'Room rate'};
  }

  // 6. Property default
  const prop=properties.find(p=>(p.Title||'').toLowerCase()===pt);
  if(prop&&prop.DailyRate)return{rate:prop.DailyRate,source:'Property default'};

  // No rate found — but check for near-misses and flag them
  const nearMiss=_findRateNearMiss(personName,company,propertyTitle);
  return{rate:0,source:'No rate set',nearMiss:nearMiss};
}

// Look up checkout fee (one-time cleaning fee at end of stay)
// Priority: 1) Company+Property  2) Company  3) Property  4) 0 (no fee)
function getCheckoutFee(company,propertyTitle){
  const co=(company||'').toLowerCase().trim();
  const pt=(propertyTitle||'').toLowerCase().trim();
  // Only consider rates explicitly marked as Checkout fee
  const checkoutRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='checkout'&&r.DailyRate);
  if(!checkoutRates.length)return 0;
  // 1. Company + specific property
  if(co){
    const r=checkoutRates.find(rt=>(rt.Company||'').toLowerCase()===co&&(rt.Property||'').toLowerCase()===pt);
    if(r)return Number(r.DailyRate)||0;
  }
  // 2. Company any property
  if(co){
    const r=checkoutRates.find(rt=>(rt.Company||'').toLowerCase()===co&&!(rt.Property));
    if(r)return Number(r.DailyRate)||0;
  }
  // 3. Property default
  const r=checkoutRates.find(rt=>(rt.Property||'').toLowerCase()===pt&&!(rt.Company));
  if(r)return Number(r.DailyRate)||0;
  return 0;
}

// Look up percent-based fee for a company (e.g. Jobzone 10% of month).
// Priority: Company+Property > Company. Returns percent as decimal (0.10 for 10%) or 0 if not configured.
function getPercentFeeRate(company,propertyTitle){
  const co=(company||'').toLowerCase().trim();
  if(!co)return 0;
  const pt=(propertyTitle||'').toLowerCase().trim();
  const percentRates=allRates.filter(r=>(r.FeeType||'').toLowerCase()==='percent'&&r.DailyRate);
  if(!percentRates.length)return 0;
  // 1. Company + specific property
  const r1=percentRates.find(rt=>(rt.Company||'').toLowerCase()===co&&(rt.Property||'').toLowerCase()===pt);
  if(r1)return (Number(r1.DailyRate)||0)/100;
  // 2. Company any property
  const r2=percentRates.find(rt=>(rt.Company||'').toLowerCase()===co&&!(rt.Property));
  if(r2)return (Number(r2.DailyRate)||0)/100;
  return 0;
}

// Does this company have a percent-based fee configured? (used to skip flat checkout fee)
function hasPercentFee(company,propertyTitle){
  return getPercentFeeRate(company,propertyTitle)>0;
}

// Detect rate config issues: rate exists for this name but property mismatch, or fuzzy company name
function _findRateNearMiss(personName,company,propertyTitle){
  const pn=(personName||'').toLowerCase().trim();
  const pt=(propertyTitle||'').toLowerCase().trim();
  const co=(company||'').toLowerCase().trim();
  // Does a rate exist with this name but a different property set?
  if(pn){
    const r=allRates.find(rt=>_nameMatch(rt.Person_Name,personName)&&rt.Property&&(rt.Property||'').toLowerCase()!==pt&&rt.DailyRate);
    if(r)return'Rate exists for "'+r.Person_Name+'" but only for property "'+r.Property+'" (this booking is at "'+propertyTitle+'")';
  }
  // Does a rate exist where Person_Name appears in the rate's Company field? (possible data entry mistake)
  if(pn){
    const r=allRates.find(rt=>(rt.Company||'').toLowerCase().includes(pn)&&rt.DailyRate);
    if(r)return'A rate with "'+personName+'" appears in the Company field of another rate row — possible data entry mistake';
  }
  return null;
}

function calcBookingNights(booking){
  if(!booking||!booking.Check_In)return 0;
  const ci=new Date(booking.Check_In);ci.setHours(0,0,0,0);
  const co=booking.Check_Out?new Date(booking.Check_Out):new Date();co.setHours(0,0,0,0);
  return Math.max(0,Math.round((co-ci)/864e5));
}

// v14.5.15: Resolve property title for a booking. Used by all rate-calc callsites
// so they work correctly in "All Properties" mode (where selectedProperty is null).
// Lookup order:
//   1. b.Property_Name (snapshot saved on booking) — BUT only if it matches a known property.
//      Old bookings sometimes have legacy values like "Private" that don't match — we treat those
//      as if Property_Name was missing and fall through to the room lookup.
//   2. Room → Property lookup via b.RoomLookupId
//   3. selectedProperty.Title (fallback for current view)
//   4. '' (last resort — calcBookingCost will return missing rate)
function getBookingPropertyTitle(b){
  if(!b)return selectedProperty?selectedProperty.Title:'';
  // 1. Try Property_Name, but verify it matches a known property
  const pname=b.Property_Name?String(b.Property_Name).trim():'';
  if(pname){
    const known=properties.find(p=>(p.Title||'').trim().toLowerCase()===pname.toLowerCase());
    if(known)return known.Title; // use canonical casing from Properties list
    // pname is set but not a known property (e.g. legacy "Private") — fall through to room lookup
  }
  // 2. Look up property via room
  const rid=String(b.RoomLookupId||'');
  if(rid){
    const room=allRooms.find(r=>r.id===rid);
    if(room&&room.PropertyLookupId){
      const prop=properties.find(p=>String(p.id)===String(room.PropertyLookupId));
      if(prop&&prop.Title)return prop.Title;
    }
  }
  // 3. Fallback to current view
  return selectedProperty?selectedProperty.Title:'';
}

function calcBookingCost(booking,propertyTitle){
  const nights=calcBookingNights(booking);
  // Rate follows billing company (if set), otherwise guest's own company
  const effectiveCompany=getEffectiveCompany(booking);
  const rateInfo=getDailyRate(booking.Person_Name,effectiveCompany,propertyTitle,booking.RoomLookupId);
  return{nights,rate:rateInfo.rate,total:nights*rateInfo.rate,source:rateInfo.source,matchedName:rateInfo.matchedName||null,nearMiss:rateInfo.nearMiss||null};
}

