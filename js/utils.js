// ============================================================
// 2GM Booking v14.7.0 — utils.js
// Dato-formatering, norske helligdager, escapeHtml
// ============================================================

function formatDate(d){if(!d)return'';const dt=new Date(d);return String(dt.getDate()).padStart(2,'0')+'.'+String(dt.getMonth()+1).padStart(2,'0')+'.'+dt.getFullYear()}
function toISODate(d){if(!d)return'';const dt=new Date(d);return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0')}

function getNextWeekday(date){const d=new Date(date);const day=d.getDay();if(day===0)d.setDate(d.getDate()+1);else if(day===6)d.setDate(d.getDate()+2);return d}

// ============================================================
// NORWEGIAN PUBLIC HOLIDAYS (v14.6.0) — calculated, no API needed
// Easter is calculated via Gauss' algorithm (valid for years 1583-4099).
// All movable holidays are derived from Easter Sunday.
// ============================================================

// Cache holiday-dates per year so we don't recompute on every call
const _holidayCache={};

// Gauss' Easter algorithm — returns Date object for Easter Sunday in given year
function _calculateEaster(year){
  const a=year%19;
  const b=Math.floor(year/100);
  const c=year%100;
  const d=Math.floor(b/4);
  const e=b%4;
  const f=Math.floor((b+8)/25);
  const g=Math.floor((b-f+1)/3);
  const h=(19*a+b-d-g+15)%30;
  const i=Math.floor(c/4);
  const k=c%4;
  const L=(32+2*e+2*i-h-k)%7;
  const m=Math.floor((a+11*h+22*L)/451);
  const month=Math.floor((h+L-7*m+114)/31);
  const day=((h+L-7*m+114)%31)+1;
  const date=new Date(year,month-1,day);date.setHours(0,0,0,0);
  return date;
}

function _addDays(date,days){
  const d=new Date(date);d.setDate(d.getDate()+days);d.setHours(0,0,0,0);return d;
}

// Returns object: { 'YYYY-MM-DD': 'Holiday name', ... } for all Norwegian public holidays in given year
function getNorwegianHolidays(year){
  if(_holidayCache[year])return _holidayCache[year];
  const map={};
  // Movable holidays first, then fixed — fixed overwrites if collision
  // (e.g. in 2027, 17. mai falls on 2. pinsedag — Grunnlovsdag should win as the more specific name)
  const easter=_calculateEaster(year);
  const movable=[
    [-3,'Skjærtorsdag'],
    [-2,'Langfredag'],
    [0,'1. påskedag'],
    [1,'2. påskedag'],
    [39,'Kristi himmelfartsdag'],
    [49,'1. pinsedag'],
    [50,'2. pinsedag']
  ];
  movable.forEach(([offset,name])=>{
    const dt=_addDays(easter,offset);
    map[toISODate(dt)]=name;
  });
  const fixedHolidays=[
    [0,1,'Nyttårsdag'],
    [4,1,'Arbeidernes dag'],
    [4,17,'Grunnlovsdag'],
    [11,25,'1. juledag'],
    [11,26,'2. juledag']
  ];
  fixedHolidays.forEach(([m,d,name])=>{
    const dt=new Date(year,m,d);dt.setHours(0,0,0,0);
    map[toISODate(dt)]=name;
  });
  _holidayCache[year]=map;
  return map;
}

// Returns holiday name string if date is a Norwegian public holiday, else null.
// Accepts Date object or ISO string.
function getHolidayName(dateOrIso){
  const d=dateOrIso instanceof Date?dateOrIso:new Date(dateOrIso);
  d.setHours(0,0,0,0);
  const year=d.getFullYear();
  const holidays=getNorwegianHolidays(year);
  return holidays[toISODate(d)]||null;
}

// True if date is Saturday, Sunday, or a Norwegian public holiday
function isNonWorkingDay(date){
  const d=date instanceof Date?date:new Date(date);
  const day=d.getDay();
  if(day===0||day===6)return true;
  return getHolidayName(d)!==null;
}

// v14.6.0: Returns next non-weekend, non-holiday day at or after given date.
// Replaces previous getNextWeekday for wash scheduling so we never schedule on holidays.
function getNextWorkingDay(date){
  const d=new Date(date);d.setHours(0,0,0,0);
  let safety=14; // max 14 days lookahead (Easter has 4 consecutive holidays max)
  while(isNonWorkingDay(d)&&safety-->0){
    d.setDate(d.getDate()+1);
  }
  return d;
}
function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]))}

