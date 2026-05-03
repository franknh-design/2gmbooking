// ============================================================
// 2GM Booking v14.7.0 — config.js
// Konfigurasjon, konstanter og global state
// ============================================================

// --- MSAL / AZURE AD ---
const _redirectUri=(function(){
  const o=window.location.origin;
  const p=window.location.pathname;
  const cleanPath=p.endsWith('/')?p:p.substring(0,p.lastIndexOf('/')+1);
  return o+cleanPath;
})();
const msalConfig={auth:{clientId:'f8e2259d-c440-41d3-94e3-3a2dce095817',authority:'https://login.microsoftonline.com/2b495272-f733-47a8-a771-bb744309fa17',redirectUri:_redirectUri},cache:{cacheLocation:'localStorage'}};
const msalInstance=new msal.PublicClientApplication(msalConfig);

// --- SHAREPOINT ---
const SITE_HOST='2gmeiendom.sharepoint.com';
const SITE_PATH='/sites/2GMBooking';
const LIST_IDS={Properties:'d842d574-f238-442a-be3d-77334727e89f',Rooms:'bfa962a0-5eb2-416c-abe8-adba06558c11',Bookings:'fe1dfe34-23df-4864-b0b1-b01bf60bfb75',Persons:'ebbe517d-83f8-4169-9423-70c63a3f8c07',Cleaning_Log:'6b1bd5f9-c54f-42ee-892f-d50c79481375',Hours:'9db53c54-70dd-483d-ad1d-565d0e4ac7ac',Users:'1b9b866f-0944-4f43-a80d-2a630e1e7c25',Rates:'a604493f-e879-48a0-bcab-cdeb9ae2195e',WashOverrides:'626a9546-60b2-4203-91fe-ca28a1a77e94'};

// --- TUYA PROXY ---
const TUYA_PROXY_BASE=(typeof window._tuyaProxyBase!=='undefined')?window._tuyaProxyBase:'https://DIN-FUNCTION-APP.azurewebsites.net/api';
const TUYA_FUNCTION_KEY=(typeof window._tuyaFunctionKey!=='undefined')?window._tuyaFunctionKey:'DIN_FUNCTION_KEY';

// --- TILLATELSER ---
const ALL_PERMS=[
  {key:'view_bookings',label:'View bookings'},
  {key:'edit_bookings',label:'Create/edit bookings'},
  {key:'checkin_out',label:'Check in/out'},
  {key:'cancel_bookings',label:'Cancel bookings'},
  {key:'cleaning',label:'Change cleaning status'},
  {key:'doortag',label:'Change door tag status'},
  {key:'print_doortag',label:'Print door tags'},
  {key:'view_hours',label:'View hours'},
  {key:'edit_hours',label:'Register hours'},
  {key:'edit_others_hours',label:'Register hours for others'},
  {key:'view_all_hours',label:'View all workers\' hours'},
  {key:'archive',label:'View archive'},
  {key:'import_export',label:'Import/Export'},
  {key:'view_prices',label:'View prices'},
  {key:'manage_rates',label:'Manage rates'},
  {key:'manage_companies',label:'Manage companies'},
  {key:'hours_reminder',label:'Daily hours reminder'},
  {key:'view_efficiency',label:'View cleaning efficiency analysis'},
  {key:'manage_lock',label:'Administrer låskoder (Tuya)'},
  {key:'admin',label:'User administration'}
];

// --- GLOBAL STATE ---
let accessToken=null,siteId=null;
let _tokenExpiresAt=0;
const TOKEN_REFRESH_MARGIN_MS=5*60*1000;
let currentUser={email:'',displayName:'',permissions:[]};
let properties=[],rooms=[],allRooms=[],bookings=[],allBookings=[],allUsers=[],allPersons=[],allRates=[],allCompanies=[],allWashOverrides=[];
let selectedProperty=null,selectedRoom=null,selectedBooking=null;
let editingBookingId=null,checkoutBookingId=null;
let activeFilter=null;
let currentView='main';
let _lastRefreshTime=Date.now();
let _knownBookingIds=new Set();
let _knownBookingModifiedMax='';
let _pollInterval=null;
let _sessionExpiredShown=false;
let msalReady=false;
