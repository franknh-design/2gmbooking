// ============================================================
// 2GM Booking v14.7.0 — invoicing.js
// Fakturering, rapporter, PDF-eksport, priskontrakder
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
// FAKTURAGRUNNLAG / INVOICING (v14.5.10)
// ============================================================
let invoicingInitialized=false;

function toggleInvoicing(){
  ensureMainView();
  // Close other panels
  document.getElementById('incomingPanel').classList.remove('open');
  document.getElementById('archivePanel').classList.remove('open');
  const pp=document.getElementById('personsPanel');if(pp)pp.classList.remove('open');
  const cp=document.getElementById('companiesPanel');if(cp)cp.classList.remove('open');
  const pr=document.getElementById('pricingPanel');if(pr)pr.classList.remove('open');
  const ap=document.getElementById('adminPanel');if(ap)ap.classList.remove('open');
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
    const cost=calcBookingCost(b,getBookingPropertyTitle(b));
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
    const propTitleForB=getBookingPropertyTitle(b); // v14.5.12: per-booking
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuation&&!hasPercentFee(effectiveCo,propTitleForB)){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,propTitleForB);
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

  // LONG-TERM CONTRACTS (per-room): segmented by guests/gaps (v14.5.10)
  Object.keys(longTermByRoomId).forEach(rid=>{
    const room=allRooms.find(r=>r.id===rid);
    if(!room)return;
    const seg=segmentLongTermRoom(room,fromDate,toDate);
    if(!seg)return;
    seg.segments.forEach(s=>{
      const dateRange=formatDate(s.fromDate)+' → '+formatDate(s.toDate);
      items.push({
        booking:{id:s.bookingId||''},
        room:room.Title||'',
        name:s.isEmpty?s.name:(s.name+' ('+(room.Title||'')+')'),
        company:seg.company,
        guestCompany:'',
        hasBillingOverride:false,
        nights:s.days,
        rate:Math.round(s.dailyRate*100)/100,
        total:s.total,
        source:dateRange+' · '+s.days+' dager'+(s.isEmpty?' · tomt':''),
        nearMiss:null,
        lineType:s.isEmpty?'longterm_empty':'longterm',
        checkoutDate:null
      });
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

  // v14.5.18: Needs-attention banner — bookings with logically inconsistent state
  // (Status=Upcoming/Active but Check_Out passed, OR Upcoming with Check_In >30d ago)
  // These ARE included in totals — banner just warns the user to verify before sending invoice.
  const naItems=finalItems.filter(i=>i.booking&&bookingNeedsAttention(i.booking)!==null);
  // Get unique bookings (one booking can produce multiple line items: nights + utvask)
  const naBookingsMap={};
  naItems.forEach(i=>{if(i.booking&&!naBookingsMap[i.booking.id])naBookingsMap[i.booking.id]={booking:i.booking,room:i.room,name:i.name,issue:bookingNeedsAttention(i.booking)}});
  const naBookings=Object.values(naBookingsMap);
  let attentionBanner='';
  if(naBookings.length){
    const list=naBookings.map(x=>{
      const issueText=x.issue.type==='invalid_status'
        ?'Status='+(x.booking.Status||'?')+', Check-out passed '+x.issue.daysSinceCheckOut+' day'+(x.issue.daysSinceCheckOut===1?'':'s')+' ago'
        :'Never checked in ('+x.issue.daysSinceCheckIn+' days since Check-in)';
      return '<li style="margin:4px 0"><strong>'+escapeHtml(x.name||'?')+'</strong> · Room '+escapeHtml(x.room||'?')+' — '+escapeHtml(issueText)+'</li>';
    }).join('');
    attentionBanner='<div style="margin-bottom:12px;padding:10px 14px;background:rgba(239,159,39,.12);border-left:3px solid #EF9F27;border-radius:6px;font-size:12px;color:#854F0B">'
      +'<div style="font-weight:500;margin-bottom:6px">⚠ '+naBookings.length+' booking'+(naBookings.length!==1?'s':'')+' need attention — included in totals, but should be verified before sending invoice.</div>'
      +'<ul style="margin:4px 0 0 20px;padding:0">'+list+'</ul>'
      +'</div>';
  }

  if(!finalItems.length){
    body.innerHTML='<div style="text-align:center;padding:40px;color:var(--text-secondary)">No bookings in this period'+(cfSel.value!=='__ALL__'?' for '+escapeHtml(cfSel.value):'')+' on '+(selectedProperty?selectedProperty.Title:'selected property')+'.</div>';
    return;
  }

  let html=attentionBanner+warnings;

  // Grand totals — separate nights from checkout/percent fees and full-tenant leases
  const nightItems=finalItems.filter(i=>i.lineType==='nights');
  const feeItems=finalItems.filter(i=>i.lineType==='checkout'||i.lineType==='percent'||i.lineType==='fulltenant'||i.lineType==='longterm'||i.lineType==='longterm_empty');
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
        ?'<button onclick="event.stopPropagation();exportInvoicingCSV(\''+k.replace(/'/g,"\\'")+'\')" style="padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:var(--bg-success);color:var(--text-success);cursor:pointer;font-size:11px;font-family:inherit;margin-left:10px" title="Export XLSX for '+escapeHtml(k)+'">↓ XLSX</button>'
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
        const isLongTermEmpty=i.lineType==='longterm_empty';
        const ci=i.booking.Check_In?formatDate(i.booking.Check_In):'';
        const co=i.booking.Check_Out?formatDate(i.booking.Check_Out):'Open';
        let period;
        if(isFullTenant)period='🔒 Full-tenant lease';
        else if(isLongTerm||isLongTermEmpty)period='🔑 '+(i.source||'').split(' · ')[0];
        else if(isPercent)period='📊 Monthly percent fee';
        else if(isCheckout)period='🧹 Checkout '+formatDate(i.checkoutDate);
        else period=ci+' → '+co;
        const nightsCell=(isCheckout||isPercent)?'—':((isFullTenant||isLongTerm||isLongTermEmpty)?i.nights+' days':i.nights);
        let rateCell;
        if(isFullTenant)rateCell='<em style="color:var(--text-tertiary)">Full-tenant</em>';
        else if(isLongTerm||isLongTermEmpty)rateCell=i.rate?i.rate.toLocaleString('nb-NO',{maximumFractionDigits:2})+' kr/dag':'<em style="color:var(--text-tertiary)">Långtid</em>';
        else if(isPercent)rateCell='<em style="color:var(--text-tertiary)">%-basert</em>';
        else if(isCheckout)rateCell='<em style="color:var(--text-tertiary)">Utvask</em>';
        else rateCell=(i.rate?i.rate.toLocaleString('nb-NO')+' kr':'<span style="color:var(--text-danger)">— missing</span>');
        const totalCell=i.total?i.total.toLocaleString('nb-NO')+' kr':'—';
        let sourceCell;
        if(isFullTenant)sourceCell='<span style="color:#1D9E75">🔒 '+escapeHtml(i.source)+'</span>';
        else if(isLongTerm)sourceCell='<span style="color:#0EA5A5">'+escapeHtml(i.source)+'</span>';
        else if(isLongTermEmpty)sourceCell='<span style="color:#a76800;font-style:italic">'+escapeHtml(i.source)+'</span>';
        else if(isPercent)sourceCell='<span style="color:#EF9F27">📊 '+escapeHtml(i.source)+'</span>';
        else if(isCheckout)sourceCell='<span style="color:#7B61FF">🧹 Checkout fee</span>';
        else sourceCell=(i.nearMiss?'<span title="'+escapeHtml(i.nearMiss)+'" style="color:var(--text-warning)">⚠ '+escapeHtml(i.source)+'</span>':escapeHtml(i.source));
        let rowStyle;
        if(isFullTenant)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(29,158,117,.08)';
        else if(isLongTerm)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(14,165,165,.07)';
        else if(isLongTermEmpty)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(239,159,39,.05);font-style:italic';
        else if(isPercent)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:default;background:rgba(239,159,39,.06)';
        else if(isCheckout)rowStyle='border-top:.5px solid var(--border-tertiary);cursor:pointer;background:rgba(123,97,255,.04)';
        else rowStyle='border-top:.5px solid var(--border-tertiary);cursor:pointer';
        const hoverBg=isFullTenant?'rgba(29,158,117,.16)':(isLongTerm?'rgba(14,165,165,.14)':(isLongTermEmpty?'rgba(239,159,39,.10)':(isPercent?'rgba(239,159,39,.12)':(isCheckout?'rgba(123,97,255,.12)':'var(--bg-secondary)'))));
        const restBg=isFullTenant?'rgba(29,158,117,.08)':(isLongTerm?'rgba(14,165,165,.07)':(isLongTermEmpty?'rgba(239,159,39,.05)':(isPercent?'rgba(239,159,39,.06)':(isCheckout?'rgba(123,97,255,.04)':''))));
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
        else if(isLongTermEmpty)nameCell='<span style="color:#a76800;font-style:italic">'+escapeHtml(i.name)+'</span>';
        else if(isPercent)nameCell='<span style="color:var(--text-warning);font-weight:500">'+escapeHtml(i.name)+'</span>';
        else if(isCheckout)nameCell='<span style="color:var(--text-tertiary)">↳ '+guestMarkedName(i.name)+'</span>';
        else nameCell=guestMarkedName(i.name);
        // Full-tenant, long-term and percent rows are not clickable
        const clickAttr=(isPercent||isFullTenant||isLongTermEmpty)?'':(isLongTerm&&i.booking.id?'onclick="openEditBooking(\''+i.booking.id+'\')"':(isLongTerm?'':'onclick="openEditBooking(\''+i.booking.id+'\')"'));
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

// v14.5.11: Replaced CSV export with XLSX (SheetJS) — proper Excel formatting,
// new column order: Room, Guest, Company, [Billing], Check-in, Check-out, Nights, Rate, Total
// 'Rate source' column removed entirely. 'Billing company' kept but column hidden if all rows are empty.
// Total row is bold. Rate and Total columns get number formatting.
function exportInvoicingCSV(companyFilterName){
  // Function name kept for backward compat with onclick handlers — actually outputs XLSX now
  if(typeof XLSX==='undefined'){
    alert('XLSX-bibliotek (xlsx-js-style) er ikke lastet. Last siden på nytt (F5) og prøv igjen.');
    return;
  }
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
  // Each row: {room, guest, company, billing, checkIn, checkOut, nights, rate, total}
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
    const cost=calcBookingCost(b,getBookingPropertyTitle(b));
    const room=allRooms.find(r=>r.id===rid);
    const origCo=(b.Company||'').trim();
    const billingCo=effectiveCo!==origCo?effectiveCo:'';
    if(nights>0){
      rows.push({
        room:room?room.Title:'',
        guest:b.Person_Name||'',
        company:origCo,
        billing:billingCo,
        checkIn:formatDate(b.Check_In),
        checkOut:b.Check_Out?formatDate(b.Check_Out):'Open',
        nights:nights,
        rate:cost.rate||0,
        total:nights*(cost.rate||0)
      });
      if(effectiveCo){companyNightSum[effectiveCo]=(companyNightSum[effectiveCo]||0)+nights*(cost.rate||0)}
    }
    // Checkout fee line (skip if company has Percent fee, skip if Continuation)
    const isContinuationExp=(b.Continuation===true||b.Continuation==='true'||b.Continuation===1);
    const propTitleForB=getBookingPropertyTitle(b); // v14.5.12: per-booking
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuationExp&&!hasPercentFee(effectiveCo,propTitleForB)){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,propTitleForB);
        if(fee>0){
          rows.push({
            room:room?room.Title:'',
            guest:'Utvask: '+(b.Person_Name||''),
            company:origCo,
            billing:billingCo,
            checkIn:'Checkout '+formatDate(b.Check_Out),
            checkOut:'',
            nights:0,
            rate:fee,
            total:fee
          });
        }
      }
    }
  });
  // Full-tenant lease lines
  Object.keys(fullTenantByPropId).forEach(pid=>{
    const ft=fullTenantByPropId[pid];
    if(companyFilterName&&ft.company!==companyFilterName)return;
    const prop=properties.find(p=>String(p.id)===String(pid));
    rows.push({
      room:prop?prop.Title:'',
      guest:ft.company+' (full-tenant lease)',
      company:'',
      billing:ft.company,
      checkIn:ft.detailLabel||'',
      checkOut:'',
      nights:ft.days,
      rate:ft.rate,
      total:ft.total
    });
  });
  // Long-term per-room contracts
  Object.keys(longTermByRoomIdCsv).forEach(rid=>{
    const lt=longTermByRoomIdCsv[rid];
    if(companyFilterName&&lt.company!==companyFilterName)return;
    const room=allRooms.find(r=>r.id===rid);
    if(!room)return;
    const seg=segmentLongTermRoom(room,fromDate,toDate);
    if(!seg)return;
    seg.segments.forEach(s=>{
      rows.push({
        room:room.Title||'',
        guest:s.isEmpty?s.name:s.name,
        company:'',
        billing:seg.company,
        checkIn:formatDate(s.fromDate),
        checkOut:formatDate(s.toDate),
        nights:s.days,
        rate:Math.round(s.dailyRate*100)/100,
        total:s.total
      });
    });
  });
  // Percent-based fee lines
  Object.keys(companyNightSum).forEach(c=>{
    const pct=getPercentFeeRate(c,propTitleForPercent);
    if(pct>0){
      const feeAmount=Math.round(companyNightSum[c]*pct);
      rows.push({
        room:'',
        guest:c+' ('+(pct*100)+'% månedsgebyr)',
        company:'',
        billing:c,
        checkIn:periodStr,
        checkOut:'',
        nights:0,
        rate:feeAmount,
        total:feeAmount
      });
    }
  });
  // Sort: room, then billing/company, then guest
  rows.sort((a,b)=>(a.room||'').localeCompare(b.room||'','nb',{numeric:true})
    ||((a.billing||a.company)+'').localeCompare(((b.billing||b.company)+''),'nb')
    ||(a.guest+'').localeCompare((b.guest+''),'nb'));

  if(!rows.length){
    alert('Ingen data å eksportere for denne perioden.');
    return;
  }

  // Determine if Billing column should be shown (any non-empty value)
  const showBilling=rows.some(r=>r.billing&&r.billing.trim()!=='');

  // Build header + AOA (array of arrays) for SheetJS
  const headers=showBilling
    ?['Room','Guest','Company','Billing','Check-in','Check-out','Nights','Rate','Total']
    :['Room','Guest','Company','Check-in','Check-out','Nights','Rate','Total'];
  const aoa=[headers];
  rows.forEach(r=>{
    if(showBilling){
      aoa.push([r.room,r.guest,r.company,r.billing,r.checkIn,r.checkOut,r.nights,r.rate,r.total]);
    }else{
      aoa.push([r.room,r.guest,r.company,r.checkIn,r.checkOut,r.nights,r.rate,r.total]);
    }
  });
  // Total row
  const totalN=rows.reduce((s,r)=>s+(typeof r.nights==='number'?r.nights:0),0);
  const totalT=rows.reduce((s,r)=>s+(typeof r.total==='number'?r.total:0),0);
  if(showBilling){
    aoa.push(['','','','','','Total',totalN,'',totalT]);
  }else{
    aoa.push(['','','','','Total',totalN,'',totalT]);
  }

  // Build worksheet
  const ws=XLSX.utils.aoa_to_sheet(aoa);

  // Column widths
  const colWidths=showBilling
    ?[{wch:10},{wch:24},{wch:18},{wch:18},{wch:12},{wch:12},{wch:8},{wch:10},{wch:12}]
    :[{wch:10},{wch:24},{wch:18},{wch:12},{wch:12},{wch:8},{wch:10},{wch:12}];
  ws['!cols']=colWidths;

  // v14.5.13: Apply formatting using xlsx-js-style (writes styles into the file)
  const lastRow=aoa.length; // 1-based row count incl header
  const numColsRate=showBilling?7:6; // 0-indexed col for Rate
  const numColsTotal=showBilling?8:7; // 0-indexed col for Total
  const numColsNights=showBilling?6:5;
  const cellAddr=(r,c)=>XLSX.utils.encode_cell({r:r,c:c});
  const ncols=headers.length;

  // Header row (row 0) — bold, light gray background, bottom border
  for(let c=0;c<ncols;c++){
    const a=cellAddr(0,c);
    if(!ws[a])ws[a]={t:'s',v:''};
    ws[a].s={
      font:{bold:true,sz:11},
      fill:{patternType:'solid',fgColor:{rgb:'EEEEEE'}},
      alignment:{horizontal:c>=numColsNights?'right':'left',vertical:'center'},
      border:{bottom:{style:'thin',color:{rgb:'888888'}}}
    };
  }
  // Total row (last row, 0-indexed = lastRow-1) — bold + top border
  for(let c=0;c<ncols;c++){
    const a=cellAddr(lastRow-1,c);
    if(!ws[a])ws[a]={t:'s',v:''};
    ws[a].s={
      font:{bold:true,sz:11},
      border:{top:{style:'thin',color:{rgb:'000000'}}},
      alignment:{horizontal:c>=numColsNights?'right':'left'}
    };
  }
  // Number format for Rate, Total, Nights columns (data rows + total row)
  for(let r=1;r<lastRow;r++){
    [numColsRate,numColsTotal].forEach(c=>{
      const a=cellAddr(r,c);
      if(ws[a]&&typeof ws[a].v==='number'){
        ws[a].z='#,##0';
        ws[a].t='n';
        // Preserve any existing style (e.g. on total row) by merging
        const existingStyle=ws[a].s||{};
        ws[a].s={...existingStyle,numFmt:'#,##0',alignment:{...existingStyle.alignment,horizontal:'right'}};
      }
    });
    const an=cellAddr(r,numColsNights);
    if(ws[an]&&typeof ws[an].v==='number'){
      ws[an].z='0';
      ws[an].t='n';
      const existingStyle=ws[an].s||{};
      ws[an].s={...existingStyle,numFmt:'0',alignment:{...existingStyle.alignment,horizontal:'right'}};
    }
  }

  // Build workbook + filename
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Fakturagrunnlag');
  const propName=(selectedProperty?selectedProperty.Title:'Alle').replace(/\s+/g,'_');
  const companyPart=companyFilterName?'_'+companyFilterName.replace(/\s+/g,'_'):'';
  const filename='Fakturagrunnlag_'+propName+companyPart+'_'+periodStr+'.xlsx';
  XLSX.writeFile(wb,filename);
}

// ============================================================
// ADD GUEST FROM BOOKING (v14.5.10)
// ============================================================
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
    const cost=calcBookingCost(b,getBookingPropertyTitle(b));
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
    const propTitleForB=getBookingPropertyTitle(b); // v14.5.12: per-booking lookup
    if(b.Status==='Completed'&&b.Check_Out&&!isContinuation&&!hasPercentFee(effectiveCo,propTitleForB)){
      const checkoutDate=new Date(b.Check_Out);checkoutDate.setHours(0,0,0,0);
      const feeEnabled=(b.Include_Checkout_Fee===undefined||b.Include_Checkout_Fee===null||b.Include_Checkout_Fee===true||b.Include_Checkout_Fee==='true'||b.Include_Checkout_Fee===1);
      if(feeEnabled&&checkoutDate>=fromDate&&checkoutDate<=toDate){
        const fee=getCheckoutFee(effectiveCo,propTitleForB);
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
  // NOTE (v14.5.12): In "All Properties" mode propTitleForPercent is '', so
  // property-specific percent rules won't match. Acceptable limitation for now —
  // percent fees are rare and usually configured per-company, not per-property.
  // Run faktura per property for accurate percent-fee calculation.
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
  // Long-term per-room contracts (v14.5.10 segmented)
  Object.keys(longTermByRoomIdPdf).forEach(rid=>{
    const lt=longTermByRoomIdPdf[rid];
    if(companyFilterName&&lt.company!==companyFilterName)return;
    const key=lt.company;
    if(!groups[key])groups[key]={nights:[],fees:[],percent:null,fullTenant:null,longTerm:[]};
    if(!groups[key].longTerm)groups[key].longTerm=[];
    const room=allRooms.find(r=>r.id===rid);
    if(!room)return;
    const seg=segmentLongTermRoom(room,fromDate,toDate);
    if(!seg)return;
    seg.segments.forEach(s=>{
      groups[key].longTerm.push({
        roomTitle:room.Title||'',
        guestName:s.name,
        isEmpty:s.isEmpty,
        price:s.dailyRate,
        days:s.days,
        total:s.total,
        fromDate:s.fromDate,
        toDate:s.toDate,
        detailLabel:formatDate(s.fromDate)+' → '+formatDate(s.toDate)+' · '+s.days+' dager',
        isMonthly:lt.isMonthly
      });
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
      tableRows+='<tr class="ft-row"><td colspan="5"><strong>🔒 Full-tenant lease — '+escapeHtml(ft.property)+'</strong><br><small>'+escapeHtml(ft.detailLabel||'')+'</small></td><td class="num"><strong>'+fmtKr(ft.total)+'</strong></td></tr>';
    }

    // Long-term per-room contracts (v14.5.10 segmented): summary + collapsible detail rows
    if(g.longTerm&&g.longTerm.length){
      const ltTotal=g.longTerm.reduce((s,lt)=>s+lt.total,0);
      groupTotal+=ltTotal;
      // Count unique rooms
      const uniqueRooms=new Set(g.longTerm.map(s=>s.roomTitle)).size;
      const sectionId='lt-'+escapeHtml(key).replace(/[^a-zA-Z0-9]/g,'_');
      // Summary row — clickable for screen, always shows on print
      tableRows+='<tr class="lt-row lt-summary" onclick="document.querySelectorAll(\'.'+sectionId+'\').forEach(el=>el.classList.toggle(\'lt-hidden\'))">'
        +'<td colspan="5"><strong>🔑 Långtidsleie ('+uniqueRooms+' rom · '+g.longTerm.length+' segmenter)</strong> <span class="muted no-print">▼ klikk for detaljer</span></td>'
        +'<td class="num"><strong>'+fmtKr(ltTotal)+'</strong></td>'
        +'</tr>';
      // Detail rows — hidden by default on screen, always shown on print
      // Column order: Rom, Gjest, Periode, Netter, Sats, Sum
      g.longTerm.forEach(s=>{
        const styleExtra=s.isEmpty?';color:#a76800;font-style:italic':'';
        tableRows+='<tr class="lt-row lt-detail '+sectionId+' lt-hidden" style="background:'+(s.isEmpty?'rgba(239,159,39,.05)':'rgba(14,165,165,.05)')+'">'
          +'<td><small>'+escapeHtml(s.roomTitle)+'</small></td>'
          +'<td style="padding-left:24px'+styleExtra+'"><small>'+(s.isEmpty?'':'↳ ')+escapeHtml(s.guestName)+'</small></td>'
          +'<td><small>'+escapeHtml(s.detailLabel||'')+'</small></td>'
          +'<td class="num"><small>'+s.days+'</small></td>'
          +'<td class="num"><small>'+fmtKr(Math.round(s.price*100)/100)+'/dag</small></td>'
          +'<td class="num"><small>'+fmtKr(s.total)+'</small></td>'
          +'</tr>';
      });
    }

    // Night bookings — column order: Rom, Gjest, Periode, Netter, Sats, Sum
    g.nights.forEach(n=>{
      groupTotal+=n.total;
      const billingInfo=n.guestCompany&&n.guestCompany!==n.effectiveCo?'<br><small class="muted">Gjest jobber for: '+escapeHtml(n.guestCompany)+'</small>':'';
      tableRows+='<tr>'
        +'<td>'+escapeHtml(n.room)+'</td>'
        +'<td>'+escapeHtml(n.name)+billingInfo+'</td>'
        +'<td>'+formatDate(n.checkIn)+' → '+(n.checkOut?formatDate(n.checkOut):'Åpen')+'</td>'
        +'<td class="num">'+n.nightsCount+'</td>'
        +'<td class="num">'+fmtKr(n.rate)+'</td>'
        +'<td class="num">'+fmtKr(n.total)+'</td>'
        +'</tr>';
    });

    // Checkout fees — column order: Rom, Gjest, Periode, Netter, Sats, Sum
    g.fees.forEach(f=>{
      groupTotal+=f.fee;
      tableRows+='<tr class="fee-row">'
        +'<td>'+escapeHtml(f.room)+'</td>'
        +'<td>↳ Utvask: '+escapeHtml(f.name)+'</td>'
        +'<td>'+formatDate(f.checkoutDate)+'</td>'
        +'<td class="num">—</td>'
        +'<td class="num">—</td>'
        +'<td class="num">'+fmtKr(f.fee)+'</td>'
        +'</tr>';
    });

    // Percent fee
    if(g.percent){
      groupTotal+=g.percent.amount;
      tableRows+='<tr class="pct-row"><td colspan="5">📊 Månedsgebyr ('+(g.percent.rate*100)+'% av '+fmtKr(g.percent.base)+')</td><td class="num"><strong>'+fmtKr(g.percent.amount)+'</strong></td></tr>';
    }

    grandTotal+=groupTotal;

    bodyHtml+='<section class="company-section">'
      +'<h2>'+escapeHtml(key)+'</h2>'
      +'<table>'
      +'<thead><tr><th>Rom</th><th>Gjest</th><th>Periode</th><th class="num">Netter</th><th class="num">Sats</th><th class="num">Sum</th></tr></thead>'
      +'<tbody>'+tableRows+'</tbody>'
      +'<tfoot><tr><td colspan="5" class="num"><strong>Sum '+escapeHtml(key)+'</strong></td><td class="num"><strong>'+fmtKr(groupTotal)+'</strong></td></tr></tfoot>'
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
// PRICING TABS — Full-tenant + Long-term editors (v14.5.10)
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
// BACKUP & RESTORE (v14.5.10)
// ============================================================
const BACKUP_LISTS=['Properties','Rooms','Bookings','Persons','Cleaning_Log','Hours','Users','Rates','Companies'];

