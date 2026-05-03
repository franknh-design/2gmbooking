// ============================================================
// 2GM Booking v14.7.0 — auth.js
// MSAL innlogging, token-håndtering
// ============================================================

function _clearStaleMsalInteraction(){
  try{
    // MSAL stores the in-progress flag with various key shapes depending on version.
    // Safest is to remove any localStorage/sessionStorage key containing 'interaction.status'.
    const keys=[];
    for(let i=0;i<localStorage.length;i++){
      const k=localStorage.key(i);
      if(k&&k.indexOf('interaction.status')>=0)keys.push(k);
    }
    keys.forEach(k=>{console.log('[Auth] Clearing stale MSAL key:',k);localStorage.removeItem(k)});
    const skeys=[];
    for(let i=0;i<sessionStorage.length;i++){
      const k=sessionStorage.key(i);
      if(k&&k.indexOf('interaction.status')>=0)skeys.push(k);
    }
    skeys.forEach(k=>{console.log('[Auth] Clearing stale MSAL session key:',k);sessionStorage.removeItem(k)});
  }catch(e){console.warn('[Auth] Could not clear stale interaction flag:',e.message)}
}

async function signIn(){
  if(!msalReady){alert('Vent litt...');return}
  // v14.5.16: proactively clear any stuck interaction flag from a previous failed login
  _clearStaleMsalInteraction();
  try{
    // v14.6.0: Switch from loginPopup to loginRedirect.
    // Reason: iOS Safari/Chrome (and some desktop Chrome configurations) block popup→opener
    // communication, causing 'monitor_window_timeout' errors or silent failures. Redirect flow
    // avoids popup entirely — full-page navigation to Microsoft and back.
    // The redirect-back is handled by handleRedirectPromise() in the init block (see bottom of app.js).
    await msalInstance.loginRedirect({scopes:['Sites.ReadWrite.All','Mail.Send']});
    // Code below this point will NOT execute on success — page will redirect to Microsoft.
    // If it does execute, loginRedirect was misconfigured.
  }catch(e){
    console.error('Login failed:',e);
    if(String(e.message||'').indexOf('interaction_in_progress')>=0){
      _clearStaleMsalInteraction();
      alert('Innloggingen ble avbrutt forrige gang. Trykk logg inn igjen.');
    }else{
      alert('Kunne ikke starte innlogging: '+e.message);
    }
  }
}

// Get token. interactive=true (default) means: fall back to popup if silent fails.
// interactive=false (used by background polling) means: never popup, just return null on failure.
async function getToken(interactive=true){
  // Reuse cached token if still valid
  if(accessToken&&_tokenExpiresAt&&Date.now()<(_tokenExpiresAt-TOKEN_REFRESH_MARGIN_MS)){
    return accessToken;
  }
  const a=msalInstance.getAllAccounts();
  if(!a.length){
    if(interactive&&!_sessionExpiredShown){
      _sessionExpiredShown=true;
      alert('Du er ikke logget inn. Last siden på nytt (F5) og logg inn på nytt.');
    }
    return null;
  }
  try{
    const r=await msalInstance.acquireTokenSilent({scopes:['Sites.ReadWrite.All','Mail.Send'],account:a[0]});
    if(!r||!r.accessToken)throw new Error('Silent returned empty token');
    accessToken=r.accessToken;
    _tokenExpiresAt=r.expiresOn?r.expiresOn.getTime():(Date.now()+50*60*1000);
    _sessionExpiredShown=false;
    return accessToken;
  }catch(e){
    console.warn('[Auth] Silent token failed:',e.message);
    if(!interactive){
      // Background call — fail silently, polling will skip this cycle
      return null;
    }
    // Interactive call — fall back to redirect (v14.6.0: was popup, but iOS Safari blocks popups)
    try{
      // acquireTokenRedirect navigates the page — code below this won't run on success
      await msalInstance.acquireTokenRedirect({scopes:['Sites.ReadWrite.All','Mail.Send']});
      // If we somehow get here, treat as failure
      throw new Error('Redirect did not navigate');
    }catch(redirectErr){
      console.error('[Auth] Redirect token failed:',redirectErr.message);
      accessToken=null;
      _tokenExpiresAt=0;
      if(!_sessionExpiredShown){
        _sessionExpiredShown=true;
        alert('Sesjonen har utløpt. Last siden på nytt (F5) og logg inn på nytt.');
      }
      throw new Error('Token unavailable');
    }
  }
}

// v14.6.0: Use logoutRedirect instead of logoutPopup for same reasons as login
function signOut(){msalInstance.logoutRedirect();document.getElementById('app').style.display='none';document.getElementById('loginScreen').style.display='block'}
function showApp(){document.getElementById('loginScreen').style.display='none';document.getElementById('app').style.display='block'}

// --- GRAPH API (v14.5.11 — all calls accept silent flag) ---
