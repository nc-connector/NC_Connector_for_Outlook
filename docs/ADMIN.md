# Admin Guide — NC Connector for Outlook

This document describes installation, rollout and operation of **NC Connector for Outlook** (Outlook classic COM add-in).

## Contents
- [Installation (MSI)](#installation-msi)
- [Updates / upgrade behavior](#updates--upgrade-behavior)
- [Files & registry](#files--registry)
- [Settings (`settings.ini`)](#settings-settingsini)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [Logging / support](#logging--support)
- [Troubleshooting](#troubleshooting)

## Installation (MSI)
1. Close Outlook.
2. Install the MSI (administrator rights required).

Silent install example:

```powershell
msiexec /i "NCConnectorForOutlook-2.2.7.msi" /qn /norestart
```

Afterwards, start Outlook. The **NC Connector** tab/group appears in the ribbon (Calendar/Appointment and Mail compose).

## Updates / upgrade behavior
- Updates are performed by installing a **newer MSI version**.
- The MSI is configured as a **major upgrade** (stable `UpgradeCode`) so an existing installation is replaced automatically.
- User settings are kept because they are stored in the user profile.

## Files & registry

### Installation path
Default (x64):
- `C:\Program Files\NextcloudTalkOutlookAddIn\`

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

## Settings (`settings.ini`)

### Location
Settings are stored per user:
- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini`

Legacy migration:
- If present, `%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\settings.ini` is migrated on first start.

### Important keys (excerpt)

```ini
ServerUrl=https://cloud.example.com
Username=max
AppPassword=xxxxxx
AuthMode=LoginFlow
IfbEnabled=true
IfbDays=30
IfbCacheHours=24
DebugLoggingEnabled=false
FileLinkBasePath=90 Freigaben - extern
SharingDefaultShareName=Share name
SharingDefaultPermCreate=false
SharingDefaultPermWrite=false
SharingDefaultPermDelete=false
SharingDefaultPasswordEnabled=true
SharingDefaultExpireDays=7
ShareBlockLang=default
EventDescriptionLang=default
TalkDefaultLobbyEnabled=true
TalkDefaultSearchVisible=true
TalkDefaultRoomType=EventConversation
TalkDefaultPasswordEnabled=true
TalkDefaultAddUsers=true
TalkDefaultAddGuests=false
```

Note: the app password is currently stored in clear text (planned improvement: Windows Credential Locker / DPAPI).

### Rollout / pre-seed
Because `settings.ini` lives in the user profile, common rollout approaches are:
- **Login Script / Intune / SCCM**: copy the file to `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini` (only if not present).
- **Group Policy Preferences**: distribute the file into the user profile.

Recommendation:
- Pre-seed only base URL and defaults.
- Let users fetch credentials via Login Flow v2 or enter them manually.

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

## Logging / support
Enable debug logging in Settings → **Debug**.

Log file:
- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`

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
