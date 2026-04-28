# Admin Guide — NC Connector for Outlook

This document describes installation, rollout and operation of **NC Connector for Outlook** (Outlook classic COM add-in).

## Contents
- [Installation (MSI)](#installation-msi)
- [Updates / upgrade behavior](#updates--upgrade-behavior)
- [Files & registry](#files--registry)
- [Settings (profile XML)](#settings-profile-xml)
- [Compose sharing lifecycle (3.0.4)](#compose-sharing-lifecycle-304)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [System address book required for user search and moderator selection](#system-address-book-required-for-user-search-and-moderator-selection)
- [Logging / support](#logging--support)
- [Troubleshooting](#troubleshooting)

## Installation (MSI)
1. Close Outlook.
2. Install the MSI (administrator rights required).

Silent install example:

```powershell
msiexec /i "NCConnectorForOutlook-3.0.4.msi" /qn /norestart
```

Afterwards, start Outlook. The **NC Connector** tab/group appears in the ribbon (Calendar/Appointment and Mail compose).

## Updates / upgrade behavior
- Updates/reinstalls are performed by installing an MSI package over the existing installation (same, older, or newer version).
- The MSI is configured as a **major upgrade** (stable `UpgradeCode`) so an existing installation is replaced automatically.
- User settings are kept because they are stored in the user profile.

## Files & registry

### Installation path
Default (x64):
- `C:\Program Files\NC4OL\`

Important files:
- `NcTalkOutlookAddIn.dll` (COM add-in)
- `NcTalkOutlookAddIn.dll.config` (binding redirects)
- `LICENSE.txt`

### Outlook add-in registration (HKLM)
The MSI registers the add-in, e.g.:
- `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

Common values:
- `LoadBehavior=3`
- `FriendlyName="NC Connector for Outlook"`
- `Description="Nextcloud Talk & Files integrated into Outlook"`

COM registration is handled via `HKLM\Software\Classes\CLSID\{...}` including `CodeBase` pointing to the installed DLL.

Installer marker key for the IFB URL reservation:
- `HKLM\Software\NC4OL\HttpUrl`

## Settings (profile XML)

### Location
Settings are stored per user and per Outlook profile:
- `%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml`
- Fallback profile file (when no profile name is available): `%LOCALAPPDATA%\NC4OL\settings_default.xml`

Password handling:
- `AppPasswordProtected` is stored encrypted via Windows DPAPI (`CurrentUser` scope).
- Plaintext `AppPassword` is not persisted in the new format.

Legacy migration on first start:
- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini`
- `%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\settings.ini`
- Legacy INI files are removed after successful migration.

### Important keys (excerpt)

```xml
<Settings SchemaVersion="1" Profile="Outlook">
  <ServerUrl>https://cloud.example.com</ServerUrl>
  <Username>max</Username>
  <AppPasswordProtected>BASE64_DPAPI_BLOB</AppPasswordProtected>
  <AuthMode>LoginFlow</AuthMode>
  <IfbEnabled>true</IfbEnabled>
  <IfbDays>30</IfbDays>
  <IfbPort>7777</IfbPort>
  <IfbCacheHours>24</IfbCacheHours>
  <DebugLoggingEnabled>false</DebugLoggingEnabled>
  <LogAnonymizationEnabled>true</LogAnonymizationEnabled>
  <FileLinkBasePath>NC Connector</FileLinkBasePath>
</Settings>
```

### Rollout / pre-seed
Because profile XML settings live in the user profile, common rollout approaches are:
- **Login Script / Intune / SCCM**: copy a prepared `settings_<OutlookProfile>.xml` to `%LOCALAPPDATA%\NC4OL\` (only if not present).
- **Group Policy Preferences**: distribute the file into the user profile.

Recommendation:
- Pre-seed only base URL and defaults.
- Let users fetch credentials via Login Flow v2 or enter them manually (recommended for DPAPI compatibility).

### Optional NC Connector backend policies

If the optional Nextcloud backend app `ncc_backend_4mc` is installed, Outlook also evaluates:
- `/apps/ncc_backend_4mc/api/v1/status`

Runtime behavior:
- checked when Talk wizard opens
- checked when Sharing wizard opens
- checked when Settings open or are saved
- valid active seat => backend policy values apply and `policy_editable=false` fields are locked in the UI
- missing backend / no seat / invalid seat => local Outlook settings remain active
- if the backend is unreachable, Outlook falls back to the locally saved add-in settings
- if the backend is reachable but the license/seat state is no longer usable, Outlook also falls back to the locally saved add-in settings
- invalid seat states remain visible in the UI so users can contact their administrator
- separate password delivery is only available when the backend endpoint exists and the current user has an active assigned seat
- backend custom templates stay inactive until the corresponding language override is set to `custom`
- the `custom` option is only shown when the backend endpoint exists and stays disabled unless the effective backend policy for that domain is actually `custom` and provides a template
- if `custom` is selected but the backend template is empty or unavailable, Outlook falls back to the local UI-default text block
- `policy.talk.event_description_type` may be `html` or `plain_text`; when `html` is active, Outlook sanitizes the Talk HTML template, applies an appointment-compat transform (legacy color/alignment fallbacks + Word-safe CSS stripping), and writes the block via HTML->RTF bridge for stable appointment rendering

Central policy can currently control:
- Talk defaults and lock state
- Sharing defaults and lock state
- share HTML/password templates
- Talk description language / custom invitation template

### Talk appointment-safe HTML subset (backend templates)

If backend policy/template delivery is enabled for Talk appointment descriptions (`event_description_type=html`), use this subset for robust Outlook rendering:

- Prefer table-based markup for layout (`table`, `tbody`, `tr`, `td`).
- Keep inline styles simple; avoid modern layout features (`flex`, `grid`) and advanced visual effects (`border-radius`, `overflow`, `object-fit`, `user-select`).
- Keep links explicit with full `https://` URLs.
- NC Connector adds legacy fallbacks automatically during appointment insert (`font color`, `bgcolor`, `align`, `valign`) and hardens anchor color rendering.

## Compose sharing lifecycle (3.0.4)

### Attachment automation and cleanup contract
- In compose attachment mode, created server artifacts are tracked immediately after share creation.
- Cleanup tracking is cleared only after a confirmed successful primary mail send.
- If a compose window is closed without a successful send, the add-in deletes the created share folder artifacts server-side (best effort, with send/close grace timer handling).
- Attachment automation evaluates new files both pre-add (`BeforeAttachmentAdd`) and post-add; if pre-add can resolve a local file path, NC flow can best-effort cancel host add before Outlook post-add handling.
- In Microsoft 365 / Exchange environments with server-side message-size limits, Outlook can block large attachments before add-in events fire; in those cases automation cannot intercept and users should use the `Insert Nextcloud share` button instead.
- In sharing wizard file-step, admins/users can add files and folders via Explorer drag & drop across the full step area (queue and action area), not only via explicit add buttons.

### Separate password follow-up mail
- If `Send password separately` is enabled, the main HTML block does not contain inline password text.
- This feature is only available when the NC Connector backend endpoint exists and the current user has an active assigned seat.
- Follow-up password mail dispatch runs only after the primary mail send is confirmed.
- Dispatch strategy:
  - first try automatic send
  - if automatic send fails, open a prefilled manual fallback draft.

## Internet Free/Busy Gateway (IFB)

### Purpose
IFB provides a local HTTP endpoint that Outlook can use to retrieve free/busy information from Nextcloud.

Endpoint:
- Configurable in `Settings -> IFB -> Local IFB port` (default `7777`).
- Default endpoint: `http://127.0.0.1:7777/nc-ifb/`

### URLACL (HTTP.SYS reservation)
The MSI reserves the URL namespace so the listener can run without administrative rights.

Check default reservation:

```powershell
netsh http show urlacl | Select-String -Pattern "7777/nc-ifb"
```

If you use a custom IFB port, add URLACL for that port (administrator shell):

```powershell
netsh http add urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/ user="S-1-1-0"
```

### Outlook registry (per user)
When enabling IFB, the add-in sets Outlook-specific registry values (HKCU) so Outlook uses the local free/busy URL.

Notes:
- IFB is optional (toggle in Settings).
- Without IFB, the add-in still works for Talk + Sharing.

## System address book required for user search and moderator selection

The following features require a reachable **Nextcloud system address book**:
- Moderator selection in the Talk wizard
- `Add users` default in add-in settings
- `Add guests` default in add-in settings

If the system address book is unavailable, these controls are disabled in the UI and show a warning with a setup link.

Nextcloud 31 activation:
- `sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"`

Nextcloud >= 32 activation:
- Nextcloud -> Admin Settings -> Groupware -> System Address Book (enable it)

Required in both versions:
- Nextcloud Admin Settings -> Sharing: enable username autocompletion / system address book access.

Repair hint (if Admin UI shows enabled but the system address book is still unavailable):
1. Reset + re-enable:
   - `sudo -E -u www-data php occ config:app:delete dav system_addressbook_exposed`
   - `sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"`
2. Rebuild the generated address book:
   - `sudo -E -u www-data php occ dav:sync-system-addressbook`
3. Verify endpoint:
   - `https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system/?export`

## Logging / support
Enable debug logging in Settings → **Debug**.
The debug tab also provides **Anonymize logs** (enabled by default).

Log files (daily rotation):
- `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`

Logs are categorized (e.g. `CORE`, `API`, `TALK`, `FILELINK`, `IFB`) and help with support cases.
When debug logging is enabled, runtime decision paths (including attachment pre-add gating and fallback reasons) are written to the same file; runtime exceptions are always written regardless of debug toggle state.
When anonymization is enabled, sensitive values are redacted before writing:
- configured Nextcloud URL/base host
- tokens and secrets in URLs/query strings/JSON payload fragments
- `Authorization` header credentials
- email addresses and user identifiers in common log field formats
- local user path segments (for example `C:\\Users\\<USER>\\...`)
Retention behavior:
- keep the latest 7 daily log files
- additionally remove log files older than 30 days (best effort cleanup)

## Troubleshooting

### Add-in not loading
1. In Outlook: `File → Options → Add-ins`
2. Check “COM Add-ins”: `NcTalkOutlook.AddIn`
3. Check “Disabled Items” (Outlook may disable add-ins after crashes)

Notes:
- If you run **32-bit Outlook** on **64-bit Windows**, Outlook reads COM add-in registration from the 32-bit registry view (`Wow6432Node`).
- Verify `LoadBehavior=3` exists in the correct registry view:
  - 64-bit: `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`
  - 32-bit: `HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

### IFB not responding
- Check which IFB port is configured in `Settings -> IFB` (default `7777`).
- Check if the configured port is bound:

```powershell
netstat -ano | Select-String ":<ifb-port>"
```

- Check URLACL (see above)
- Enable debug logging and inspect `IFB` entries

### Network / Nextcloud
- Server reachable, TLS ok?
- TLS behavior can now be switched in `Settings -> Advanced -> Transport security (TLS)` (`OS default` or forced TLS versions such as 1.2/1.3).
- NC Connector applies the selected policy at runtime via `ServicePointManager.SecurityProtocol`.
- Connection diagnostics operations (settings connection test and login flow) now enforce fresh HTTP/TLS handshakes, so TLS-mode switching is validated deterministically instead of reusing pooled keep-alive sockets.
- If secure-channel errors still occur, check certificate trust, DNS, proxy/TLS inspection, and machine TLS/Schannel policy before considering machine-wide registry/GPO overrides.
- App password valid?
- Talk installed/enabled?
- Password Policy app optional: if missing, passwords are generated locally (fallback)



