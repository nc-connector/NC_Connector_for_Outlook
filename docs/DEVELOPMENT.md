# Development Guide — NC Connector for Outlook

This document is a newcomer-friendly guide for building, debugging, and extending **NC Connector for Outlook** (Outlook classic COM add-in).

## Contents

- [Project purpose](#project-purpose)
- [Quick start](#quick-start)
- [Repository structure](#repository-structure)
- [Architecture](#architecture)
- [Network endpoints](#network-endpoints)
- [Localization (i18n)](#localization-i18n)
- [Logging](#logging)
- [Compatibility & version checks](#compatibility--version-checks)
- [Build & release](#build--release)
- [Local testing](#local-testing)
- [X-NCTALK-* property reference](#x-nctalk--property-reference)
- [Extension points](#extension-points)

## Project purpose

The add-in connects Outlook classic to a Nextcloud server and provides:

- **Nextcloud Talk** from calendar appointments (room creation, lobby, participant sync, moderator delegation)
- **Nextcloud sharing** from the mail compose window (upload + link share + HTML block insertion)
- **Internet Free/Busy (IFB)** via a local HTTP endpoint that proxies requests to Nextcloud

## Quick start

### Prerequisites

- Windows 10/11
- Outlook classic (x64 or x86)
- **.NET Framework 4.7.2** (target framework)
- MSBuild (e.g. Visual Studio Build Tools)
- **.NET SDK** (used by WiX v4 build via `dotnet`)

### Build MSI (recommended)

```powershell
cd "C:\\path\\to\\nc4ol-0.2.7"

# Optional: reference assemblies (only if needed)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages
$env:FrameworkPathOverride = "$PWD\\packages\\Microsoft.NETFramework.ReferenceAssemblies.net472\\build\\.NETFramework\\v4.7.2"

.\build.ps1 -Configuration Release
```

Output:

- `dist\\NCConnectorForOutlook-<version>.msi`

### Install & run locally

1. Install the MSI (administrator rights required):
   - `msiexec /i dist\\NCConnectorForOutlook-<version>.msi`
2. Start Outlook
3. Ribbon:
   - Calendar/appointment: **NC Connector → Insert Talk link**
   - Mail compose: **NC Connector → Insert Nextcloud share**
4. Open **NC Connector → Settings** and configure server URL + credentials.

## Repository structure

Top-level:

- `src/` — the COM add-in (WinForms UI + service layer)
- `installer/` — WiX v4 MSI project (files + registry + URLACL)
- `docs/` — admin/development documentation
- `assets/` — branding images used in README/screenshots
- `dist/` — build output (MSI)

Key code locations:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` — entry point, ribbon XML, Outlook event wiring, orchestration
- `src/NcTalkOutlookAddIn/Services/` — Nextcloud HTTP integrations (Talk, sharing, IFB, login flow)
- `src/NcTalkOutlookAddIn/UI/` — WinForms dialogs and wizards
- `src/NcTalkOutlookAddIn/Settings/` — persisted settings model + storage
- `src/NcTalkOutlookAddIn/Utilities/` — logging, theming, i18n, small shared helpers

## Architecture

### Main building blocks

- **COM add-in lifecycle**
  - `NextcloudTalkAddIn.OnConnection(...)` loads settings, enables logging (optional), initializes IFB, and wires Outlook events.
- **Service layer**
  - `Services/TalkService.cs` calls the Talk OCS API.
  - `Services/FileLinkService.cs` uploads via WebDAV and creates shares via OCS.
  - `Services/FreeBusyServer.cs` hosts the local IFB HTTP endpoint.
  - `Services/FreeBusyManager.cs` updates Outlook registry keys to point to the local IFB endpoint.
- **UI**
  - `UI/SettingsForm.cs` configures base URL, authentication, sharing defaults, IFB, and debug logging.
  - `UI/TalkLinkForm.cs` is the Talk wizard.
  - `UI/FileLinkWizardForm.cs` is the sharing wizard.
  - `UI/BrandedHeader.cs` is the shared header banner control.

### End-to-end flows

#### Talk link flow (appointments)

1. User clicks **Insert Talk link** in an appointment.
2. `UI/TalkLinkForm.cs` collects: title, password, lobby, listable flag, room type, participant sync options, optional delegation target.
3. `Services/TalkService.cs` creates the room via OCS.
4. `NextcloudTalkAddIn.ApplyRoomToAppointment(...)` updates the appointment:
   - `Location` (Talk URL)
   - a localized plain-text body block (incl. password and help URL)
   - persisted metadata as Outlook `UserProperties` (including `X-NCTALK-*` keys)
5. A runtime subscription is registered for the appointment (`AppointmentSubscription` inside `NextcloudTalkAddIn.cs`):
   - **Write** (save): updates lobby timer on time changes, updates room description, syncs participants, applies delegation
   - **Close** (discard without saving): deletes the room to avoid orphans (best-effort)
   - **BeforeDelete**: deletes the room (best-effort)

#### Sharing flow (mail compose)

1. User clicks **Insert Nextcloud share** while composing an email.
2. `UI/FileLinkWizardForm.cs` collects sharing settings and the file/folder selection.
3. `Services/FileLinkService.cs` performs WebDAV upload and creates the public share via OCS.
4. `Utilities/FileLinkHtmlBuilder.cs` generates the HTML block (header + link + password + permissions + expiration date).
5. `NextcloudTalkAddIn.InsertHtmlIntoMailItem(...)` inserts the HTML into the message body.

#### IFB flow

1. User enables IFB in Settings.
2. `Services/FreeBusyServer.cs` starts a local HTTP listener (default: `http://127.0.0.1:7777/nc-ifb/`).
3. `Services/FreeBusyManager.cs` updates Outlook registry values so Outlook requests free/busy data via the local endpoint.

## Network endpoints

The add-in uses Nextcloud **OCS** and **WebDAV** endpoints.

Talk (OCS, selection):

- Capabilities/version hint: `GET /ocs/v2.php/cloud/capabilities`
- Create room: `POST /ocs/v2.php/apps/spreed/api/v4/room`
- Delete room: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>`
- Lobby timer: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/webinar/lobby`
- Listable scope: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/listable`
- Description: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/description`
- Add participants: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants`
- Get participants: `GET /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants?includeStatus=true`
- Promote moderator: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/moderators`
- Self leave: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants/self`

Sharing:

- Create public share: `POST /ocs/v2.php/apps/files_sharing/api/v1/shares`
- Upload/folder creation: `remote.php/dav/...` (WebDAV)

IFB (DAV via proxy):

- Local listener: `http://127.0.0.1:7777/nc-ifb/...`
- The proxy talks to CalDAV and Addressbook endpoints under `remote.php/dav/...`

## Localization (i18n)

- Locale files:
  - `src/NcTalkOutlookAddIn/Resources/_locales/<lang>/messages.json`
- Runtime loader:
  - `src/NcTalkOutlookAddIn/Utilities/Strings.cs`

Notes:

- The default language is **German** (`de`).
- The UI language is derived from Windows UI culture. Some generated text blocks can be overridden via Settings (see “Language overrides”).
- Placeholders in `messages.json` use `$1`, `$2`, ... and are converted to `.NET` `string.Format` placeholders.

See `Translations.md` for the full language list and maintenance workflow.

## Logging

Debug logging is optional and is intended to make support cases reproducible.

- Enable: Settings → **Debug** → “Write debug log file”
- Log file: `%LOCALAPPDATA%\\NextcloudTalkOutlookAddInData\\addin-runtime.log`

Format:

- `[YYYY-MM-DD HH:mm:ss.fff] [CATEGORY] Message`

Example:

```
[2026-02-13 03:57:12.345] [TALK] BEGIN CreateRoom
[2026-02-13 03:57:12.910] [TALK] END CreateRoom (565 ms)
```

Implementation:

- `src/NcTalkOutlookAddIn/Utilities/DiagnosticsLogger.cs`
- `src/NcTalkOutlookAddIn/Utilities/LogCategories.cs`

Guidelines for new code:

- Log **start/end** of network operations (use `DiagnosticsLogger.BeginOperation(...)`).
- Log **decisions** (feature detection, version checks, fallbacks).
- Log **exceptions with context** (use `DiagnosticsLogger.LogException(...)`).
- Never swallow exceptions silently.

## Compatibility & version checks

### Outlook bitness (x86 on x64 Windows)

Outlook can be installed as a 32-bit application on 64-bit Windows. In that case it reads COM add-in registration from the 32-bit registry view (`Wow6432Node`).

The MSI registers add-in keys for **both** registry views:

- 64-bit: `HKLM\\Software\\Microsoft\\Office\\Outlook\\Addins\\NcTalkOutlook.AddIn`
- 32-bit: `HKLM\\Software\\Wow6432Node\\Microsoft\\Office\\Outlook\\Addins\\NcTalkOutlook.AddIn`

Installer definition:

- `installer/Product.wxs`

### Nextcloud feature detection

Some features depend on server capabilities.

- Event conversations require Nextcloud **>= 31**.
- Version checks and cached “server version hint” live in:
  - `src/NcTalkOutlookAddIn/Utilities/NextcloudVersionHelper.cs`

### UI theming (WinForms)

The add-in uses a dark theme where appropriate so dialogs match dark Outlook setups.

Implementation:

- `src/NcTalkOutlookAddIn/Utilities/UiThemeManager.cs`

Detection logic (best-effort):

1. Try Office/Outlook theme registry values (when available).
2. Fallback to Windows “app theme” (`AppsUseLightTheme`).
3. High contrast mode disables custom theming (system colors win).

## Build & release

### What `build.ps1` does

1. Builds the COM add-in (`NcTalkOutlookAddIn.sln`) via MSBuild
2. Reads the assembly version from `NcTalkOutlookAddIn.dll`
3. Builds the WiX v4 installer (`installer/NcConnectorOutlookInstaller.wixproj`)
4. Copies the MSI into `dist/`

### Versioning

- `src/NcTalkOutlookAddIn/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`

`build.ps1` derives the MSI `ProductVersion` from that (format `Major.Minor.Build`).

### Release checklist

1. Bump version in `AssemblyInfo.cs`
2. `.\build.ps1 -Configuration Release`
3. Install/upgrade MSI test (old version → new version)
4. Smoke test (Talk + sharing + IFB)
5. Optional: sign the MSI (if required in your environment)

## Local testing

Suggested smoke test sequence:

Note: there is currently no automated test suite in this repository. Use the smoke tests below to validate changes.

1. Enable debug logging in Settings.
2. Calendar: create a new appointment, insert a Talk link, save the appointment, then change start time and save again (lobby update).
3. Calendar: add attendees, save again (participant sync).
4. Mail: run the sharing wizard, upload 1–2 small files, insert the HTML block, and send to yourself.
5. IFB: enable IFB, then verify the local endpoint responds:
   - `Invoke-WebRequest http://127.0.0.1:7777/nc-ifb/ -UseBasicParsing`

## X-NCTALK-* property reference

The add-in persists appointment metadata as Outlook `UserProperties`. A subset of those properties uses `X-NCTALK-*` names to enable interoperability and stable re-sync behavior.

Unless stated otherwise:

- Properties are stored as **text** values in Outlook (`OlUserPropertyType.olText`).
- Boolean values are stored as `TRUE` / `FALSE` (uppercase).
- Timestamps are stored as **Unix epoch seconds** (UTC) in invariant culture.

Primary write location:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` → `ApplyRoomToAppointment(...)`

### Properties

| Property | Purpose | Type / format | Example | Written | Read / used | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| `X-NCTALK-TOKEN` | Talk room token | `string` | `a1b2c3d4` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Read is preferred over legacy token storage. |
| `X-NCTALK-URL` | Talk room URL | `string` | `https://cloud.example.com/call/a1b2c3d4` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Stored for interoperability. |
| `X-NCTALK-LOBBY` | Lobby enabled flag | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Used to decide whether lobby updates run on save. |
| `X-NCTALK-START` | Appointment start time (epoch seconds) | `int64` as string | `1739750400` | `ApplyRoomToAppointment(...)`, `AppointmentSubscription.OnWrite(...)` | (not read by add-in) | Updated when lobby is enabled and the start time changes. |
| `X-NCTALK-EVENT` | Room creation mode marker | `event` \| `standard` | `event` | `ApplyRoomToAppointment(...)` | `GetRoomType(...)` | Fallback exists via `NcTalkRoomType` user property. |
| `X-NCTALK-OBJECTID` | Time-window identifier | `"<start>#<end>"` | `1739750400#1739754000` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Stored for interoperability. |
| `X-NCTALK-ADD-USERS` | Participant sync: internal users | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Preferred over legacy `X-NCTALK-ADD-PARTICIPANTS`. |
| `X-NCTALK-ADD-GUESTS` | Participant sync: external emails | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Preferred over legacy `X-NCTALK-ADD-PARTICIPANTS`. |
| `X-NCTALK-ADD-PARTICIPANTS` | Participant sync: legacy combined toggle | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Used as fallback when the split toggles are missing (older items). |
| `X-NCTALK-DELEGATE` | Delegation target user ID | `string` | `alice` | `ApplyRoomToAppointment(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)`, `TryApplyDelegation(...)` | Stored in parallel to the `NcTalkDelegateId` user property for backward compatibility. |
| `X-NCTALK-DELEGATE-NAME` | Delegation target display name | `string` | `Alice Example` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Stored for interoperability. |
| `X-NCTALK-DELEGATED` | Delegation state marker | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)`, `TryApplyDelegation(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)` | Controls whether delegation is still pending. |
| `X-NCTALK-DELEGATE-READY` | Delegation “ready” marker | `TRUE` | `TRUE` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Reserved for interoperability; the add-in currently uses `X-NCTALK-DELEGATED` + delegate ID to detect pending delegation. |

## Extension points

### Add a new setting

1. Add property to `src/NcTalkOutlookAddIn/Settings/AddinSettings.cs`.
2. Persist it in `src/NcTalkOutlookAddIn/Settings/SettingsStorage.cs`.
3. Add UI in `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`.
4. Add translations (see `Translations.md`).

### Add a new Nextcloud API call

1. Add to the appropriate service:
   - Talk: `src/NcTalkOutlookAddIn/Services/TalkService.cs`
   - Sharing: `src/NcTalkOutlookAddIn/Services/FileLinkService.cs`
2. Add request/response model in `src/NcTalkOutlookAddIn/Models/` (if needed).
3. Add logging scopes and error handling.
4. Integrate in the UI/wizard and wire it up via `NextcloudTalkAddIn.cs`.

### Add a new localized string

1. Add a property to `src/NcTalkOutlookAddIn/Utilities/Strings.cs`.
2. Add the key to all locale files under `src/NcTalkOutlookAddIn/Resources/_locales/`.
3. Rebuild and verify the UI.
