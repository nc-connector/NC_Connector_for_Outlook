# Admin Guide — NC Connector for Outlook

This document describes installation, rollout and operation of **NC Connector for Outlook** (Outlook classic COM add-in).

## Contents
- [Installation (MSI)](#installation-msi)
- [Updates / upgrade behavior](#updates--upgrade-behavior)
- [Files & registry](#files--registry)
- [Settings (profile XML)](#settings-profile-xml)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [System address book required for user search and moderator selection](#system-address-book-required-for-user-search-and-moderator-selection)
- [Logging / support](#logging--support)
- [Troubleshooting](#troubleshooting)

## Installation (MSI)
1. Close Outlook.
2. Install the MSI (administrator rights required).

Silent install example:

```powershell
msiexec /i "NCConnectorForOutlook-2.2.9.msi" /qn /norestart
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
  <IfbCacheHours>24</IfbCacheHours>
  <DebugLoggingEnabled>false</DebugLoggingEnabled>
  <FileLinkBasePath>90 Freigaben - extern</FileLinkBasePath>
</Settings>
```

### Rollout / pre-seed
Because profile XML settings live in the user profile, common rollout approaches are:
- **Login Script / Intune / SCCM**: copy a prepared `settings_<OutlookProfile>.xml` to `%LOCALAPPDATA%\NC4OL\` (only if not present).
- **Group Policy Preferences**: distribute the file into the user profile.

Recommendation:
- Pre-seed only base URL and defaults.
- Let users fetch credentials via Login Flow v2 or enter them manually (recommended for DPAPI compatibility).

## Internet Free/Busy Gateway (IFB)

### Purpose
IFB provides a local HTTP endpoint that Outlook can use to retrieve free/busy information from Nextcloud.

Default endpoint:
- `http://127.0.0.1:7777/nc-ifb/`

### URLACL (HTTP.SYS reservation)
The MSI reserves the URL namespace so the listener can run without administrative rights.

Check:

```powershell
netsh http show urlacl | Select-String -Pattern "7777/nc-ifb"
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

Log file:
- `%LOCALAPPDATA%\NC4OL\addin-runtime.log`

Logs are categorized (e.g. `CORE`, `API`, `TALK`, `FILELINK`, `IFB`) and help with support cases.

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
- Check if port 7777 is bound:

```powershell
netstat -ano | Select-String ":7777"
```

- Check URLACL (see above)
- Enable debug logging and inspect `IFB` entries

### Network / Nextcloud
- Server reachable, TLS ok?
- App password valid?
- Talk installed/enabled?
- Password Policy app optional: if missing, passwords are generated locally (fallback)


