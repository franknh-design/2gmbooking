# Tuya Smart Lock — Deploy & Configuration Checklist

One-time setup to make the Tuya buttons in v15 actually work end-to-end. Code is already shipped to `main`. Until all six sections below are done, the buttons either won't show or will fail with fetch errors against `DIN-FUNCTION-APP.azurewebsites.net`.

Estimated time: 30–60 minutes the first time.

---

## 1. SharePoint columns

Open `https://2gmeiendom.sharepoint.com/sites/2GMBooking` → Settings → Site contents.

### List `Rooms` — add columns:
| Column name (internal) | Type | Notes |
|---|---|---|
| `Tuya_Device_ID` | Single line of text (255) | Per-lock device ID. Empty rooms fall back to manual door code. |
| `Door_Code` | Single line of text (10) | Used by manual fallback flow + as `{room_door_code}` in SMS/email templates. May already exist. |
| `Door_Battery_Level` | Number (0 decimal places) | Already used by battery auto-refresh. May already exist. |
| `Door_Battery_Updated` | Date and Time | Already used by battery auto-refresh. May already exist. |

### List `Bookings` — add column:
| Column name (internal) | Type | Notes |
|---|---|---|
| `Tuya_Password_ID` | Single line of text (50) | Lock-side ID returned by `create_pin`, used later to call `delete_pin`. Stored as string because `_stripUnknownFieldsAsync` skips Yes/No fields and `Number` columns sometimes give SP grief; string is safest. |

**Verify column names match exactly** — SharePoint internal names are case-sensitive. Code references:
- `room.Tuya_Device_ID` (`tuya.js:49`, `bookings.js:51`)
- `booking.Tuya_Password_ID` (`tuya.js:91`, `bookings.js:54`)
- `room.Door_Code` (`tuya.js:58, 89`, `messaging.js:74`)

### Populate device IDs for known rooms
From `2gmbooking/tuya_lock.py` (working values, already tested):

| Room title (must match `Title` field) | `Tuya_Device_ID` |
|---|---|
| `701` | `bf61810d99855d670ahwkv` |
| `702` | `bf58f8f02b54f34865kd1a` |

Add additional rooms in the Tuya Smart Life app → Device → "..." → Device information → "Virtual ID" (or "Device ID").

---

## 2. Deploy Azure Function

Prerequisites:
- Azure subscription with rights to create resources
- Azure CLI (`az`) and Azure Functions Core Tools v4 (`func`) installed
  - `winget install Microsoft.AzureCLI`
  - `winget install Microsoft.Azure.FunctionsCoreTools`
- Logged in: `az login`

### Create resources (one-time)

Pick globally-unique names. Suggested:

```powershell
$RG = "rg-2gm-booking"
$LOCATION = "norwayeast"
$STORAGE = "st2gmbooking$([guid]::NewGuid().ToString('N').Substring(0,6))"  # must be globally unique, lowercase only
$APP = "fn-2gm-tuya"   # must be globally unique

az group create --name $RG --location $LOCATION
az storage account create --name $STORAGE --resource-group $RG --location $LOCATION --sku Standard_LRS
az functionapp create --name $APP --resource-group $RG --storage-account $STORAGE `
  --consumption-plan-location $LOCATION --runtime python --runtime-version 3.11 `
  --functions-version 4 --os-type Linux
```

### Set Tuya credentials as Function App Settings

```powershell
az functionapp config appsettings set --name $APP --resource-group $RG --settings `
  TUYA_CLIENT_ID="pwtyxkvk3wp7jqtreshk" `
  TUYA_CLIENT_SECRET="c68ac7114e1f467fb6d5ad092cbfa993" `
  TUYA_BASE_URL="https://openapi.tuyaeu.com"
```

(Values from `2gmbooking/tuya_lock.py` lines 30–32. If you've rotated the Tuya app credentials since then, use the new ones from the Tuya IoT Platform.)

### Allow CORS

```powershell
az functionapp cors add --name $APP --resource-group $RG --allowed-origins https://booking.2gm.no
az functionapp cors add --name $APP --resource-group $RG --allowed-origins https://franknh-design.github.io
```

### Publish the function code

From the cloned repo:

```powershell
cd "C:\Users\hauga\OneDrive - 2gm Eiendom AS\2Prosjekter\2gmbooking-git\azure_function\tuya_proxy"
func azure functionapp publish $APP --python
```

The publish command takes 2–5 minutes. When it finishes it prints the four endpoint URLs:
```
  create_pin: https://fn-2gm-tuya.azurewebsites.net/api/tuya/create_pin
  delete_pin: https://fn-2gm-tuya.azurewebsites.net/api/tuya/delete_pin
  list_pins:  https://fn-2gm-tuya.azurewebsites.net/api/tuya/list_pins
  health:     https://fn-2gm-tuya.azurewebsites.net/api/tuya/health
```

### Get the function key

```powershell
az functionapp keys list --name $APP --resource-group $RG --query "functionKeys.default" -o tsv
```

Copy the output — you'll paste it into `index.html` next.

### Smoke test

```powershell
$KEY = "<paste function key>"
curl "https://$APP.azurewebsites.net/api/tuya/health?code=$KEY"
```

Expected: `{"status":"ok","configured":true}`. If `configured:false`, app settings didn't propagate — restart the Function App and retry.

```powershell
curl "https://$APP.azurewebsites.net/api/tuya/list_pins?code=$KEY&device_id=bf61810d99855d670ahwkv"
```

Expected: JSON list of current PINs on lock 701. If you get `502` with `Token feilet`, the Tuya credentials are wrong.

---

## 3. Wire the function URL into `index.html`

Open `2gmbooking-git/index.html`, find the `<script src="https://alcdn.msauth.net/...msal-browser..."></script>` line (around line 9), and add **before** the `<script src="js/config.js">` line at the bottom:

```html
<script>
  window._tuyaProxyBase = 'https://fn-2gm-tuya.azurewebsites.net/api';
  window._tuyaFunctionKey = '<paste function key from step 2>';
</script>
```

`config.js:22-23` reads these globals if present and falls back to the placeholder strings if they're not — so injecting them via `window.*` keeps the secret out of the committed source. **Do commit this change** since GitHub Pages doesn't allow runtime config — but the function key alone is not enough to reach the locks (you also need the `code` query param, which the wrapper appends).

If you want stricter separation, create a separate file e.g. `config-prod.js` (gitignored), and load it from index.html before `js/config.js`. Then commit only the `<script src="config-prod.js">` line, not the keys themselves.

After editing:
```powershell
cd "C:\Users\hauga\OneDrive - 2gm Eiendom AS\2Prosjekter\2gmbooking-git"
git add index.html
git commit -m "Wire Tuya proxy URL + function key into production config"
git push origin main
```

---

## 4. Grant `manage_lock` permission to admins

In the SharePoint `Users` list, edit the rows for users who should be able to create/delete/list PINs. The `Permissions` field is a comma-separated string — add `manage_lock` to it.

Users with legacy `Role = SuperAdmin` or `Role = Admin` (and no explicit `Permissions` field) automatically get all permissions including `manage_lock` — no edit needed.

| User type | Action |
|---|---|
| New users with explicit `Permissions` | Append `,manage_lock` to the field |
| Legacy users with `Role = SuperAdmin` / `Admin` | Nothing — already covered |
| Cleaning staff / read-only users | Do **not** add `manage_lock` |

---

## 5. End-to-end smoke test

1. Open `https://booking.2gm.no` as a user with `manage_lock`.
2. Hard reload (`Ctrl+Shift+R`).
3. Find a booking on a room where `Tuya_Device_ID` is populated (use room `701` or `702` for first test).
4. Click the booking → in the detail panel you should see **🔑 Opprett PIN på lås**.
5. Click it. Confirm the PIN. Wait ~5 seconds.
6. Verify in the Tuya Smart Life app that a new PIN with the booking's name appears under the lock's "Temporary password" list.
7. Test the lock with the new PIN.
8. Check out the booking from the app — verify the PIN is auto-deleted from the Tuya app and the booking's `Tuya_Password_ID` field is cleared.

If any step fails, check `https://booking.2gm.no` browser console for the actual error, and the Azure Function's "Log stream" (Function App → Monitoring → Log stream) for server-side errors.

---

## 6. Rollback

If Tuya causes problems and you need to disable it without redeploying:

- **Per-room:** clear the `Tuya_Device_ID` field on the affected room. The UI falls back to the manual "Vis kode" button automatically.
- **System-wide:** revoke `manage_lock` from all users. Create/delete/list buttons disappear; only "Vis PIN"-display remains.
- **Total:** in `index.html`, change `window._tuyaProxyBase` to a non-existent URL (e.g. `'about:blank'`). Tuya fetch will fail fast and the rest of the app is unaffected.

---

## Where each thing lives in the code

| Concern | File / line |
|---|---|
| Browser → proxy base URL | `js/config.js:22-23` (reads `window._tuyaProxyBase` / `_tuyaFunctionKey`) |
| Tuya client (browser) | `js/tuya.js` |
| UI buttons + permission gating | `js/render.js:677-693` |
| Auto-delete on checkout | `js/bookings.js:50-58` |
| Manual fallback (no Tuya_Device_ID) | `js/tuya.js:330+` (`generateRoomDoorCode`, `showRoomDoorCode`, `showDoorCodeDisplay`) |
| Azure Function HTTP routes | `azure_function/tuya_proxy/function_app.py` |
| Tuya HMAC signing + AES PIN encryption | `azure_function/tuya_proxy/tuya_client.py` |
| Reference Python CLI (predates proxy) | `../2gmbooking/tuya_lock.py` (snapshot folder, not in git) |
