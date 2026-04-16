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

## Release 3.0.2 delta summary

This release added parity-critical behavior that developers should preserve in future changes:

- Compose attachment automation now has deterministic two-mode handling (`always` vs. threshold prompt).
- Compose attachment automation now also evaluates pre-add (`BeforeAttachmentAdd`) and can best-effort cancel host add before Outlook post-add handling when a local path is resolvable.
- Compose threshold prompt is a strict two-action decision (`Share with NC Connector` / `Remove last selected attachments`) with batch removal semantics.
- Attachment-mode share output is specialized (`email_attachment` naming contract, read-only recipients, ZIP download link `/s/<token>/download`, no permissions row).
- Compose share cleanup is lifecycle-driven (armed after share creation, cleared only after confirmed send, delayed delete on unsent close race).
- Separate password follow-up dispatch is post-send only and includes automatic send plus manual fallback draft behavior.
- Talk defaults and Talk wizard use a centralized system-addressbook availability contract and deterministic lock state.
- Optional NC Connector backend policy mode is evaluated on wizard/settings entrypoints and can override/lock Talk + Sharing defaults, including central text/template payloads.
- Backend custom templates are only activated when the effective language override is `custom`; otherwise runtime stays on local UI-default text blocks.
- The `custom` option is only shown when the backend endpoint exists and stays disabled unless the effective backend policy for the respective domain is actually `custom` and provides a template.
- Separate password follow-up dispatch is seat-gated and only available with backend endpoint + active assigned seat.
- Backend attachment-threshold policy uses `attachments_min_size_mb` as both value and enable-state: a positive integer enables threshold mode, `null` disables it.
- Locked backend attachment-automation policy is also enforced in compose runtime through the live backend status, not only in Settings UI.
- If the backend is unreachable, the runtime falls back to local add-in settings immediately.
- If the backend is reachable but the license/seat state is no longer usable or the backend grace window has expired, the runtime also falls back to local add-in settings.
- Talk event-description templates may arrive as `html` or `plain_text`; when `html` is active, Outlook writes the Talk block into the open appointment editor HTML with stable markers so the rendered event description stays HTML while `Body` continues to provide the synchronized plain-text view.
- Share creation now follows the documented Nextcloud OCS contract more closely: create via `POST /shares` with `label`, then update mutable metadata like `note` via form-encoded `PUT /shares/{id}` arguments.
- Runtime settings and caches use `%LOCALAPPDATA%\\NC4OL` with profile-scoped XML settings migration.
- Transport security settings are runtime-live in `SettingsForm` (`OS default` / `TLS 1.2` / `TLS 1.3` / combined) and are intentionally validated on apply.
- Unsupported manual `TLS 1.3` runtime apply now fails closed (no implicit downgrade to TLS 1.2).
- Settings connectivity operations (connection test, login flow) enforce fresh HTTP/TLS handshakes (no pooled keep-alive reuse), so TLS mode toggles are tested deterministically.
- TLS transport diagnostics guidance text is mode-neutral and no longer assumes OS-default mode.

## Quick start

### Prerequisites

- Windows 10/11
- Outlook classic (x64 or x86)
- **.NET Framework 4.7.2** (target framework)
- MSBuild (e.g. Visual Studio Build Tools)
- **.NET SDK** (used by WiX v4 build via `dotnet`)

### Build MSI (recommended)

```powershell
cd "C:\\path\\to\\nc4ol-3.0.2"

# Optional: reference assemblies (only if needed)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages
$env:FrameworkPathOverride = "$PWD\\packages\\Microsoft.NETFramework.ReferenceAssemblies.net472\\build\\.NETFramework\\v4.7.2"

.\build.ps1 -Configuration Release
```

If WiX ICE validation is not available on the build host (for example `WIX0217` in restricted environments), use:

```powershell
.\build.ps1 -Configuration Release -SkipIceValidation
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
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs` — compose runtime subscription lifecycle
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs` — appointment runtime subscription lifecycle
- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` — appointment lifecycle orchestration for Talk room metadata/sync
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` — compose share cleanup + separate-password dispatch flow
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` — Talk template/body block rendering
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` — SMTP and attendee recipient resolution
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` — compose-subscription registry lifecycle
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
- **Shared utilities**
  - `Utilities/BrowserLauncher.cs` centralizes shell browser URL starts.
  - `Utilities/SizeFormatting.cs` centralizes MB display formatting.
  - `Utilities/ComInteropScope.cs` centralizes COM release/final-release patterns.

### End-to-end flows

#### Talk link flow (appointments)

1. User clicks **Insert Talk link** in an appointment.
2. `UI/TalkLinkForm.cs` collects: title, password, lobby, listable flag, room type, participant sync options, optional delegation target.
3. `Services/TalkService.cs` creates the room via OCS.
4. `Controllers/TalkAppointmentController.ApplyRoomToAppointment(...)` (invoked by `NextcloudTalkAddIn`) updates the appointment:
   - `Location` (Talk URL)
   - a localized plain-text body block (incl. password and help URL)
   - persisted metadata as Outlook `UserProperties` (including `X-NCTALK-*` keys)
5. A runtime subscription is registered for the appointment (`AppointmentSubscription` inside `NextcloudTalkAddIn.cs`):
   - **Write** (save): updates lobby timer on time changes, updates room description, syncs participants, applies delegation
   - If Outlook exposes the final changed start time only shortly after `Write`, a short deferred post-write verification retries the lobby update on the same opened appointment instead of broad calendar scanning.
   - **Close** (discard without saving): deletes the room to avoid orphans (best-effort)
   - **BeforeDelete**: deletes the room (best-effort)

#### Sharing flow (mail compose)

1. User clicks **Insert Nextcloud share** while composing an email.
2. `UI/FileLinkWizardForm.cs` collects sharing settings and the file/folder selection.
3. `Services/FileLinkService.cs` performs WebDAV upload, creates the public share via OCS (`label` on create), then updates mutable metadata like `note` via the documented OCS update arguments.
4. `Utilities/FileLinkHtmlBuilder.cs` generates the HTML block (header + link + password + permissions + expiration date).
5. `NextcloudTalkAddIn.InsertHtmlIntoMailItem(...)` inserts the HTML into the message body.

Compose runtime parity additions in `NextcloudTalkAddIn.cs` (`MailComposeSubscription`) with lifecycle logic delegated to `Controllers/ComposeShareLifecycleController`:

- Debounced attachment evaluation (`ComposeAttachmentEvalDebounceMs`) after compose attachment changes.
- Attachment automation modes:
  - always route attachments into NC sharing flow, or
  - threshold mode with a two-action prompt (`Share with NC Connector` / `Remove last selected attachments`).
- Pre-add attachment interception:
  - `BeforeAttachmentAdd` path resolves candidate file metadata early
  - can best-effort cancel host attachment add and launch NC sharing before Outlook post-add handling.
  - hard Outlook/Exchange size blocks can still happen before add-in callbacks and are not interceptable via official Outlook OOM events.
- Runtime host guard checks (live large-attachment setting) at:
  - pre-evaluation
  - pre-prompt-action handling
  - wizard finalize (enforced in `UI/FileLinkWizardForm.cs`).
- Attachment-mode wizard launch:
  - removes selected compose attachments
  - queues files as initial wizard selections
  - opens directly in file-step-equivalent mode.
- `UI/FileLinkWizardForm.cs` file-step queue accepts Explorer drag & drop for files/folders across queue and action-area controls.
- Compose share cleanup lifecycle:
  - arm immediately after share creation
  - clear only after successful send
  - delete server folder artifacts on unsent close (with send/close grace timer).
- Separate password-mail dispatch:
  - queue password-only HTML after share creation
  - capture recipients on send
  - dispatch only after successful primary send
  - auto-send first, then manual fallback draft on failure.

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
- Optional safety control (default on): “Anonymize logs”
- Daily log file format: `%LOCALAPPDATA%\\NC4OL\\addin-runtime.log_YYYYMMDD`
- Runtime exceptions are always written via `DiagnosticsLogger.LogException(...)`, even when debug logging is disabled.
- Retention: keep latest 7 daily log files and delete files older than 30 days (best effort cleanup).
- Anonymization redacts configured NC URL/base host, token/password-like values, authorization credentials, user identifiers, email addresses, and local user path fragments before log write.

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
  `LogException(...)` bypasses the optional debug switch and must remain the always-on error path.
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
| `X-NCTALK-START` | Appointment start time (epoch seconds) | `int64` as string | `1739750400` | `ApplyRoomToAppointment(...)`, `AppointmentSubscription.OnWrite(...)` | `TryReadRequiredIcalStartEpoch(...)`, `TryUpdateLobby(...)` | Authoritative lobby timer source on appointment save; updated when lobby is enabled and the start time changes. |
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


