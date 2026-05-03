// ============================================================
// 2GM Booking v14.7.0 — graph.js
// Microsoft Graph API-kall og SharePoint-operasjoner
// ============================================================

async function graphGet(ep,silent=false){
  const tok=await getToken(!silent);
  if(!tok){if(silent)return null;throw new Error('Token unavailable')}
  const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{headers:{Authorization:'Bearer '+tok,Accept:'application/json'}});
  if(!r.ok){if(silent&&r.status===401)return null;throw new Error('Graph error '+r.status+': '+await r.text())}
  return r.json();
}
async function graphPatch(ep,body){
  const tok=await getToken();
  const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'PATCH',headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json'},body:JSON.stringify(body)});
  if(!r.ok)throw new Error('Graph error '+r.status);
  return r.json();
}
async function graphPost(ep,body){
  const tok=await getToken();
  const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'POST',headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json'},body:JSON.stringify(body)});
  if(!r.ok){const t=await r.text();throw new Error('Graph error '+r.status+': '+t)}
  return r.json();
}
async function graphDelete(ep){
  const tok=await getToken();
  const r=await fetch('https://graph.microsoft.com/v1.0'+ep,{method:'DELETE',headers:{Authorization:'Bearer '+tok}});
  if(!r.ok)throw new Error('Graph error '+r.status);
  return true;
}

async function getSiteId(){if(siteId)return siteId;const r=await graphGet('/sites/'+SITE_HOST+':'+SITE_PATH);siteId=r.id;return siteId}
// Cache for dynamically resolved list IDs (lists not in LIST_IDS hardcoded map)
const _dynamicListIds={};
async function getListId(name){
  if(LIST_IDS[name])return LIST_IDS[name];
  if(_dynamicListIds[name])return _dynamicListIds[name];
  // Fall back to looking up by display name via Graph API
  const s=await getSiteId();
  try{
    const r=await graphGet('/sites/'+s+'/lists?$filter=displayName eq \''+name+'\'&$select=id,displayName');
    if(r.value&&r.value.length){
      _dynamicListIds[name]=r.value[0].id;
      return r.value[0].id;
    }
  }catch(e){console.error('Failed to lookup list '+name+':',e)}
  throw new Error('List not found: '+name);
}
async function getListItems(listName){const s=await getSiteId();const lid=await getListId(listName);let all=[];let url='/sites/'+s+'/lists/'+lid+'/items?$expand=fields&$top=500';while(url){const r=await graphGet(url);all=all.concat(r.value.map(i=>({id:i.id,...i.fields})));url=r['@odata.nextLink']?r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0',''):null}return all}

// Fetch a text file from a SharePoint document library.
// If pathInLibrary starts with a library name that's not the default, tries that library specifically.
// Example: 'Batteristatus/RoomBattery.csv' — first tries default lib with that path, then tries 'Batteristatus' as its own library.
async function fetchSiteFileText(pathInLibrary){
  const s=await getSiteId();
  await getToken();
  const errors=[];
  // Attempt 1: Default library (Shared Documents) with the full path
  try{
    const url='https://graph.microsoft.com/v1.0/sites/'+s+'/drive/root:/'+encodeURI(pathInLibrary);
    const r=await fetch(url,{headers:{Authorization:'Bearer '+accessToken}});
    if(r.ok){
      const item=await r.json();
      if(item['@microsoft.graph.downloadUrl']){
        const c=await fetch(item['@microsoft.graph.downloadUrl']);
        if(c.ok)return c.text();
      }
    }
    errors.push('Default library: '+r.status);
  }catch(e){errors.push('Default library: '+e.message)}
  // Attempt 2: Parse path as "LibraryName/file/path" — try as separate library
  const firstSlash=pathInLibrary.indexOf('/');
  if(firstSlash>0){
    const libName=pathInLibrary.substring(0,firstSlash);
    const remaining=pathInLibrary.substring(firstSlash+1);
    try{
      // Find the drive with matching name
      const drives=await graphGet('/sites/'+s+'/drives');
      const lib=drives.value.find(d=>d.name===libName||d.name.toLowerCase()===libName.toLowerCase());
      if(lib){
        const url='https://graph.microsoft.com/v1.0/drives/'+lib.id+'/root:/'+encodeURI(remaining);
        const r=await fetch(url,{headers:{Authorization:'Bearer '+accessToken}});
        if(r.ok){
          const item=await r.json();
          if(item['@microsoft.graph.downloadUrl']){
            const c=await fetch(item['@microsoft.graph.downloadUrl']);
            if(c.ok)return c.text();
          }
        }
        errors.push('Library "'+libName+'": '+r.status);
      }else{
        errors.push('Library "'+libName+'" not found among: '+drives.value.map(d=>d.name).join(', '));
      }
    }catch(e){errors.push('Library search: '+e.message)}
  }
  throw new Error('File not found. Tried:\n'+errors.join('\n'));
}
// Cache of known columns per list. Populated lazily on first save attempt.
const _knownColumnsByList={};
const _unknownFieldsByList={};

async function _discoverColumns(listName){
  if(_knownColumnsByList[listName])return _knownColumnsByList[listName];
  try{
    const s=await getSiteId();const lid=await getListId(listName);
    const res=await graphGet('/sites/'+s+'/lists/'+lid+'/columns?$select=name,displayName');
    const cols=new Set();
    (res.value||[]).forEach(c=>{if(c.name)cols.add(c.name)});
    // Also add common system fields that should always be allowed even if not in schema
    ['Title'].forEach(k=>cols.add(k));
    _knownColumnsByList[listName]=cols;
    console.log('[SharePoint] Discovered '+cols.size+' columns for '+listName+':',[...cols].sort().join(', '));
    return cols;
  }catch(e){
    console.warn('Could not discover columns for '+listName+':',e.message);
    _knownColumnsByList[listName]=new Set();
    return _knownColumnsByList[listName];
  }
}

async function _stripUnknownFieldsAsync(listName,fields){
  const cols=await _discoverColumns(listName);
  if(!cols||!cols.size)return fields; // discovery failed — let SharePoint reject as before
  const cleaned={};
  const skipped=[];
  Object.keys(fields).forEach(k=>{
    // Always allow Lookup-prefixed fields (e.g. RoomLookupId) — SharePoint resolves these
    if(k.endsWith('LookupId')||cols.has(k)){
      let v=fields[k];
      // PRAGMATIC: Yes/No fields cause 500 errors via Graph API.
      // Skip them entirely — SharePoint default value will be used.
      // TODO: figure out correct format. For now this gets bookings working.
      const isBool=(typeof v==='boolean'||v===0||v===1);
      const isYesNoField=(k==='Include_Checkout_Fee'||k==='Continuation');
      if(isBool&&isYesNoField){
        console.log('[SharePoint] Skipping Yes/No field "'+k+'" with value '+v+' (Graph API issue — using SharePoint default)');
        return; // skip this field
      }
      cleaned[k]=v;
    }else{
      skipped.push(k);
      if(!_unknownFieldsByList[listName])_unknownFieldsByList[listName]=new Set();
      _unknownFieldsByList[listName].add(k);
    }
  });
  if(skipped.length){
    console.warn('[SharePoint] Skipping unknown columns in '+listName+': '+skipped.join(', ')+'. Create these in SharePoint to enable.');
  }
  return cleaned;
}

async function createListItem(listName,fields){
  const cleaned=await _stripUnknownFieldsAsync(listName,fields);
  // Strip null/undefined values for create — SharePoint can throw 500 on unexpected null
  const final={};
  Object.keys(cleaned).forEach(k=>{if(cleaned[k]!==null&&cleaned[k]!==undefined)final[k]=cleaned[k]});
  const s=await getSiteId();const lid=await getListId(listName);
  console.log('[SharePoint] CREATE '+listName+' payload:',JSON.parse(JSON.stringify(final)));
  console.log('[SharePoint] CREATE '+listName+' field names:',Object.keys(final).join(', '));
  console.log('[SharePoint] Known columns:',[..._knownColumnsByList[listName]||[]].sort().join(', '));
  try{
    return await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:final});
  }catch(e){
    if(String(e.message||'').indexOf('500')<0&&String(e.message||'').indexOf('General exception')<0)throw e;
    // 500 with no useful info → systematic bisect
    console.warn('[BISECT] 500 received. Building payload up from minimal to find the broken field combination...');
    const keys=Object.keys(final);
    // Phase 1: try absolute minimal — just Title (or empty)
    const startMinimal={};
    if(final.Title)startMinimal.Title=final.Title;
    else startMinimal.Title='_BISECT_TEST_'+Date.now();
    let lastWorking=null;
    let lastWorkingItemId=null;
    try{
      console.log('[BISECT] Phase 1: minimal payload',startMinimal);
      const r=await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:startMinimal});
      console.log('[BISECT] ✓ Minimal succeeded with id='+r.id);
      lastWorking={...startMinimal};
      lastWorkingItemId=r.id;
    }catch(e2){
      console.warn('[BISECT] ✗ Even minimal payload failed:',e2.message);
      throw new Error('Save failed. Even a minimal payload (just Title) fails. This is a list-level problem in SharePoint, not a field problem. Original error: '+e.message);
    }
    // Phase 2: add fields one at a time
    let breakingField=null;
    let breakingValue=null;
    for(let i=0;i<keys.length;i++){
      const k=keys[i];
      if(k in lastWorking)continue;
      const testFields={...lastWorking,[k]:final[k]};
      try{
        console.log('[BISECT] Adding "'+k+'"='+JSON.stringify(final[k])+'...');
        // Delete previous test item before creating new one
        if(lastWorkingItemId){try{await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+lastWorkingItemId)}catch(e3){}}
        const r=await graphPost('/sites/'+s+'/lists/'+lid+'/items',{fields:testFields});
        lastWorking=testFields;
        lastWorkingItemId=r.id;
        console.log('[BISECT] ✓ OK with "'+k+'"');
      }catch(e2){
        console.warn('[BISECT] ✗ FAILED when adding "'+k+'"='+JSON.stringify(final[k])+':',e2.message);
        breakingField=k;
        breakingValue=final[k];
        break;
      }
    }
    // Cleanup last test item
    if(lastWorkingItemId){try{await graphDelete('/sites/'+s+'/lists/'+lid+'/items/'+lastWorkingItemId)}catch(e3){console.warn('[BISECT] Could not delete test item '+lastWorkingItemId+' — please remove manually')}}
    if(breakingField){
      throw new Error('Save failed. Adding field "'+breakingField+'" with value '+JSON.stringify(breakingValue)+' broke the request. Check SharePoint column type/required. Last working set: '+Object.keys(lastWorking).join(', '));
    }
    throw new Error('Save failed unexpectedly. Bisect added all fields without breaking but original payload still failed. Strange. Original error: '+e.message);
  }
}
async function updateListItem(listName,itemId,fields){
  const cleaned=await _stripUnknownFieldsAsync(listName,fields);
  const s=await getSiteId();const lid=await getListId(listName);
  return graphPatch('/sites/'+s+'/lists/'+lid+'/items/'+itemId+'/fields',cleaned);
}

// --- USER & PERMISSIONS ---
