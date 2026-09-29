# Development Guide — NC Connector for Outlook

This document is a newcomer-friendly guide for building, debugging, and extending **NC Connector for Outlook** (Outlook classic COM add-in).

Administrator rollout, configuration, operating checks, and incident runbooks are documented in [ADMIN.md](ADMIN.md).

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
- **Central backend email signatures** for matching Outlook sender accounts
- **Internet Free/Busy (IFB)** via a local HTTP endpoint that proxies requests to Nextcloud

## Quick start

### Prerequisites

- Windows 10/11 (64-bit)
- Outlook classic (x64 or x86)
- **.NET Framework 4.7.2** (target framework)
- MSBuild (e.g. Visual Studio Build Tools)
- **.NET SDK** (used by WiX v6 build via `dotnet`)
- **Nextcloud 32 or newer** (runtime server)

### Build MSI (recommended)

```powershell
cd "C:\path\to\nc4ol"

# Optional: reference assemblies (only if needed)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages -ExcludeVersion
$env:FrameworkPathOverride = "$PWD\\packages\\Microsoft.NETFramework.ReferenceAssemblies.net472\\build\\.NETFramework\\v4.7.2"

.\build.ps1 -Configuration Release
```

If WiX ICE validation is not available on the build host (for example `WIX0217` in restricted environments), use:

```powershell
.\build.ps1 -Configuration Release -SkipIceValidation
```

Output:

- `dist\NCConnectorForOutlook-<version>.msi`

### Install & run locally

1. Install the MSI (administrator rights required):
   - `msiexec /i dist\NCConnectorForOutlook-<version>.msi`
2. Start Outlook
3. Ribbon:
   - Calendar/appointment: **NC Connector → Insert Talk link**
   - Mail compose: **NC Connector → Insert Nextcloud share**
   - Inline reply/forward: **Message → NC Connector → Insert Nextcloud share**
4. Open **NC Connector → Settings** and configure server URL + credentials.

## Repository structure

Top-level:

- `src/` — the COM add-in (WinForms UI + service layer)
- `installer/` — WiX v6 MSI project (files + registry + URLACL)
- `docs/` — admin/development documentation
- `VENDOR.md` — bundled third-party sanitizer/runtime dependency notices and licenses
- `assets/` — branding images used in README/screenshots
- `dist/` — build output (MSI)

## Architecture

### Main building blocks

Key code locations:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` — entry point, ribbon XML, Outlook event wiring, orchestration
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs` — add-in bootstrap/teardown lifecycle (`OnConnection`, shutdown/disconnect)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs` — dedicated Outlook event hook/unhook wiring helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.CalendarSelection.cs` — selected-appointment rebinding for calendar-view deletion without a calendar scan
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs` — category-specific runtime logging helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs` — backend policy + Talk template/language resolver helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs` — deferred appointment-subscription ensure and Outlook event-restriction handling
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs` — compose subscription core state + lifecycle entry points (`Dispose`, identity, shared helpers)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs` — compose attachment events, timers, and prompt orchestration
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentPolicy.cs` — attachment policy snapshots, refresh, and send validation
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentMaterialization.cs` — attachment snapshots, local files, and removal from Outlook
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentQueue.cs` — queue handoff, before-add batches, and suppression completion
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Signature.cs` — backend email-signature policy application for the matching Outlook sender account
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Send.cs` — send gate, final recipient/account capture, and direct separate-password dispatch
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs` — `AfterWrite`, `Inspector.Close`, and inline `Unload` handling for newly inserted shares
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.ComposeLifecycle.cs` — composition-root bridge for queued share cleanup and direct password dispatch
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs` — appointment runtime subscription lifecycle
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.TalkAppointmentSync.cs` — STA capture and background dispatch for appointment changes
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.TalkRoomLifecycle.cs` — startup, retry recovery, and teardown for queued Talk room deletions
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs` — settings open/save/revert orchestration
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs` — FileLink ribbon launch + wizard orchestration
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs` — Talk ribbon flow orchestration (auth gate, wizard, room create/replace)
- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` with its `Sync` partial — appointment metadata, local snapshot capture, and remote room updates
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareCleanupTracker.cs` — in-memory pending-write state for newly inserted compose shares
- `src/NcTalkOutlookAddIn/Services/ComposeShareCleanupService.cs` — exact-origin deletion of unpersisted or insertion-failed server artifacts
- `src/NcTalkOutlookAddIn/Controllers/SeparatePasswordDeliveryController.cs` — password-mail body, recipient, sender, Secrets, and signature preparation and direct Outlook submission
- `src/NcTalkOutlookAddIn/Utilities/RecipientAddressList.cs` — shared recipient normalization, deduplication, and semicolon-separated lists
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` — Talk template/body block rendering
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` — SMTP and attendee recipient resolution
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` — compose-subscription registry lifecycle
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` — active mail, inline compose, Inspector identity, and dialog ownership
- `src/NcTalkOutlookAddIn/Controllers/MailBodyInsertionController.cs` — share-block insertion through the existing WordEditor and HTMLBody paths
- `src/NcTalkOutlookAddIn/Controllers/ManagedEmailSignatureController.cs` — the shared WordEditor signature-slot reconciler
- `src/NcTalkOutlookAddIn/Controllers/AppointmentHtmlBodyWriter.cs` — the appointment HTML-to-RTF bridge
- `src/NcTalkOutlookAddIn/Utilities/WordHtmlInsertionFile.cs` — shared HTML-document wrapping and temporary-file cleanup for Word insertion
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` — shared model for separate password follow-up dispatch queue entries
- `src/NcTalkOutlookAddIn/Services/` — Nextcloud HTTP integrations (Talk, sharing, IFB, login flow)
  - `Services/FileLinkQueueSnapshotBuilder.cs` builds the source-grouped queue tree used by the sharing wizard.
  - `Services/FileLinkDavClient.Browsing.cs` reads the user's DAV hierarchy and quota; `Services/FileLinkDavClient.Copy.cs` copies selected Nextcloud files into the share.
  - `Services/NcHttpClient.cs` is the shared request executor for auth headers, OCS headers, timeout/decompression, and optional fresh-connection mode.
  - All runtime HTTP calls (Talk, share/DAV, IFB, login flow, moderator avatar fetch) are routed through `NcHttpClient`.
  - `Services/EmailSignaturePolicyService.cs` resolves backend email-signature policy values against local settings and lock state.
  - `Services/UpdateCheckService.cs` checks `nc-connector.de` once per day for Outlook release metadata and stores the cached result in profile settings.
  - `Services/TalkAppointmentSyncCoordinator.cs` coalesces background Talk updates captured from Outlook events.
  - `Services/TalkRoomLifecycleCoordinator.cs` and `TalkRoomLifecycleStore.cs` persist and retry queued room deletions without scanning Outlook calendars.
  - `Services/IfbRegistryOwnershipManager.cs` and `IfbRegistryStateStore.cs` own IFB registry recovery; `FreeBusyServer.cs` validates the secret request path and limits concurrent requests.
  - `Services/ProtectedJsonStateStore.cs` provides the shared DPAPI-protected JSON and backup-recovery path used by the Talk deletion queue and IFB registry ownership state. Writes prefer atomic replacement and retain the established copy fallback; the typed stores retain feature-specific file names, entropy, validation, and diagnostics.
- `src/NcTalkOutlookAddIn/UI/` — WinForms dialogs and wizards
  - `UI/NextcloudFilePickerForm.cs` provides the **My Nextcloud** file and folder picker.
  - `UI/NextcloudPickerNavigation.cs` owns the picker history. The form advances it only after a successful folder load.
  - `UI/NextcloudPickerPreview.cs` loads and decodes one file preview. Selection, cancellation generation, and displayed-image ownership remain with the form; requests still run through `FileLinkService`.
  - `UI/ScaledForm.cs` is the shared DPI-scaling base for forms that use logical pixel layout helpers.
- `src/NcTalkOutlookAddIn/Settings/` — persisted settings model, storage, and managed setup policy
  - `Settings/ManagedSetupPolicy.cs` reads the managed Nextcloud URL from Windows policy registry keys.
  - `Settings/SettingsFileTransaction.cs` serializes profile writes through a named mutex and replaces a validated settings file while retaining its last valid backup.
- `src/NcTalkOutlookAddIn/Utilities/` — logging, theming, i18n, small shared helpers
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` — centralized HtmlSanitizer 9.0.892 policy for backend-provided share/talk HTML; active-content containers such as `template` are removed
- `src/NcTalkOutlookAddIn/Utilities/HtmlToPlainTextConverter.cs` — DOM-based HTML-to-plain-text rendering for plain-text email signatures
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` — centralized JSON payload normalization (`PrepareJsonPayload`), dictionary/string/int helpers, and OCS error extraction
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` — encapsulated pending-key tracking + throttled logging state for deferred appointment ensure
- `src/NcTalkOutlookAddIn/Utilities/NextcloudUriValidator.cs` — HTTPS base-URL validation and same-origin validation for server-supplied endpoints
- `src/NcTalkOutlookAddIn/Utilities/PictureConverter.cs` — shared Image -> IPictureDisp conversion helper for ribbon icons

#### Central email signature flow (mail compose)

The compose subscription evaluates the central signature after a compose surface opens, after sender or body-format changes, and once more in Outlook's cancellable send event.

Runtime rules:

- Backend signature insertion requires an active backend policy for the `email_signature` domain, an active assigned seat, non-empty `policy.email_signature.email_signature_template`, and `policy.email_signature.user_email`.
- Missing `policy.email_signature` support disables only central signatures and surfaces a backend update hint; Share/Talk policy domains remain independent.
- `email_signature_on_compose`, `email_signature_on_reply`, and `email_signature_on_forward` are backend defaults while their matching `policy_editable` value is `true`. A saved local value may therefore enable an editable backend default of `false`. A locked value (`policy_editable=false`) always wins, including a locked `false`.
- The effective Outlook sender identity must match `policy.email_signature.user_email`; other identities are left untouched. A `SentOnBehalfOfName`/From override for shared mailboxes or delegated Exchange identities takes precedence over `SendUsingAccount` and must resolve to the same SMTP address. If the sender identity cannot be resolved exactly, signature processing fails closed.
- New mail, reply, and forward use their corresponding effective setting. If compose insertion is active but reply or forward insertion is off, the matching sender clears an exact initial Outlook signature slot for that response and inserts no backend signature.
- Compose-kind resolution reads `PR_LAST_VERB_EXECUTED` first and then conversation metadata. If Outlook exposes only a generic inline `Response` and reply/forward values differ, background processing retries without mutation and the send gate blocks rather than guessing. If both values match, that common value applies.
- HTML, plain text, and RTF all use Outlook WordEditor. HTML and RTF import the sanitized template through a Word range; plain text uses `HtmlToPlainTextConverter` and a Word text range. `MailItem.HTMLBody` and `MailItem.Body` are not rewritten, and RTF remains RTF.
- Inspector and inline compose use the same reconciler. `Explorer.InlineResponse`, `Explorer.InlineResponseClose`, and `Inspectors.NewInspector` update the active surface, so popping an inline response into its own Inspector continues signature processing instead of retaining stale inline state.
- Backend signature HTML is sanitized through `HtmlTemplateSanitizer` with the same fail-closed policy used by sharing and Talk templates.
- The reconciler resolves the target in this order: NC Connector's `NcConnectorSignature` bookmark, Outlook's `_MailAutoSig` bookmark, then a safe structural insertion point. For a new mail that point is the end of the authored document; for reply/forward it is Outlook Word's protected position two characters before `_MailOriginal`, or a Word paragraph-border quote separator when `_MailOriginal` is absent. The actual `_MailOriginal` quote boundary remains its bookmark start and is tracked separately from the protected insertion target; using that boundary itself for fallback insertion would place the signature below Outlook's visible divider.
- An existing managed or `_MailAutoSig` slot is not trusted blindly: if its end is followed by meaningful authored content before the safe new-mail/quote target, replacement is staged at that safe target and the misplaced old slot is removed afterward. Reply/forward slot validation uses the actual quote boundary, not the protected insertion target, so Outlook's native signature table may end exactly at `_MailOriginal` and still be replaced in place. A slot entirely beyond the actual boundary is moved to the protected target; a slot that starts at or crosses the actual boundary is left untouched and reconciliation fails closed. This covers direct matching replies as well as identity changes after the user deleted the original text and signature and already-open drafts with a misplaced managed slot.
- The current cursor, the start of the message body, raw HTML prefixes, and localized reply-header text are never insertion or deletion fallbacks. If a reply/forward quote boundary cannot be located, the operation returns without changing authored or quoted content.
- The paragraph-border quote fallback excludes every paragraph inside a Word table, including nested permission cells in an inserted share. It checks `Range.Information(wdWithInTable)` before reading borders and fails closed if table membership cannot be read. A table border must never relocate the signature into the share during send-time reconciliation; the `_MailOriginal` bookmark path is unchanged.
- Table-based Outlook signatures are replaced only when `_MailAutoSig` is inside that table. Each successful insertion receives the `NcConnectorSignature` bookmark, including HTML, RTF, and plain text, so later updates and clears address only the managed range.
- Replacement is staged before the previous signature range is removed. When staging above an existing slot, a temporary Word bookmark tracks that old range while inserted content shifts its numeric offsets. If insertion, bookmark creation, tracking, or old-range deletion fails, the staged content is removed and the previous managed range is restored where possible.
- Cursor/selection preservation uses a temporary Word bookmark rather than stale absolute offsets. Safe fallback insertion adds only missing paragraph marks above the signature and one separator paragraph before quoted content.
- Sender and `BodyFormat` changes schedule another reconciliation. Changes observed while attachment processing is suppressing compose work, or while an inline item has no active surface, are deferred and resumed once a usable WordEditor returns.
- When policy becomes inactive or the sender no longer matches, NC Connector removes only `NcConnectorSignature`. It does not scan or rewrite arbitrary body content and does not remove a native signature from a non-matching identity.
- Signature processing only runs for unsent Outlook compose items. Opening a received or already sent message for reading must never modify its body.
- Before send, the pending debounce is stopped and the current sender, format, compose kind, policy, and managed slot are reconciled synchronously. With complete backend connection settings, sending is cancelled if no successful policy snapshot is available or if a required apply/clear operation cannot finish safely. The compose item stays open for correction and retry.

- Missing policy uses `email_signature_policy_unavailable`; a failed final apply/clear uses `email_signature_send_reconcile_failed`. Both cancel send and explain the next step without exposing internal source identifiers. The cache still accepts a matching last-success snapshot after a failed refresh; a fresh successful response, including refusal, replaces it. No network work is added to Outlook's send event.
- An incomplete backend setup does not create a signature requirement; cleanup is limited to best-effort removal of an exact `NcConnectorSignature` bookmark. An unsupported signature domain disables insertion as well, but with otherwise complete backend settings an existing managed range must still reconcile safely at send time. An `InlineResponseClose` that arrives just before send does not by itself block a previously reconciled, unchanged message.
- Separate password follow-up dispatch captures the successful policy and settings snapshot at share creation. On the primary mail's `Send` event it applies and reads back `SendUsingAccount`/`SentOnBehalfOfName`, submits only when the effective follow-up identity equals the captured primary sender, and adds the backend signature only when that identity also matches `policy.email_signature.user_email`. Plain source mail produces a plain follow-up; HTML/RTF source produces an HTML follow-up. A definite automatic-send failure displays the fully prepared message for manual delivery; an ambiguous submission is never repeated.
- Debug logging records the trigger, active surface, body format, compose kind, slot source, and reconciliation result without writing the signature template or sender address.

- **COM add-in lifecycle**
  - `NextcloudTalkAddIn.OnConnection(...)` loads settings, enables logging (optional), initializes IFB, and wires Outlook events.
- **Workflow controllers**
  - `SettingsWorkflowController`, `FileLinkLaunchController`, and `TalkRibbonController` own ribbon-triggered UI/runtime workflows.
  - `NextcloudTalkAddIn.cs` remains the COM/ribbon/event composition root and delegates feature flows to controllers.
  - `TalkRibbonController` prefetches backend policy + password policy before opening its wizard. `FileLinkLaunchController` also loads the required capability snapshot and passes it into the sharing wizard. Runtime policy data remains fresh on every entry.
  - FileLink, Talk, and settings prefetch completion is marshalled through `OutlookUiSynchronizationContext` before WinForms controls or Outlook COM objects are accessed. Outlook does not reliably provide a `SynchronizationContext` to COM callbacks; modal dialogs and Outlook interop must therefore return explicitly to the STA thread captured during add-in startup.
  - Lifecycle, policy/template resolution, and deferred subscription ensure are split into dedicated partial files to keep the root orchestration class maintainable.
- **Service layer**
  - `Services/TalkService.cs` calls the Talk OCS API.
  - `Services/FileLinkService.cs` orchestrates FileLink planning, remote root creation, transfer, and share creation.
  - `Services/NextcloudCapabilitiesService.cs` validates the global Nextcloud 32 requirement and caches the typed OCS capability snapshot shared by connection and feature flows.
  - `Services/FileLinkSelectionScanner.cs` scans local selections once and produces root-relative paths. `Services/FileLinkUploadPlanner.cs` assigns direct, chunked, or optional bulk transfer modes before remote mutation; `Services/FileLinkUploadPlan.cs` contains the resulting plan models.
  - `Services/FileLinkDavClient.cs` owns DAV collection lifecycle operations. Its `Probes` and `Requests` partials isolate exact resource checks, retry handling, failure mapping, and DAV URL construction.
  - `Services/FileLinkTransferService.cs` coordinates `FileLinkBulkUploader`, `FileLinkDirectUploader`, and `FileLinkChunkUploader`; `FileLinkSourceFile` validates the scanned source metadata around each transfer.
  - `Services/FileLinkShareClient.cs` creates the public share with one OCS request. Its `Recovery` partial resolves ambiguous create results through an exact-path lookup.
  - `Services/FileLinkUploadProgress.cs` aggregates phase and transfer progress and applies the update rate limit.
  - `Services/FreeBusyServer.cs` hosts the local IFB HTTP endpoint.
  - `Services/FreeBusyManager.cs` updates Outlook registry keys to point to the local IFB endpoint.
  - `Services/UpdateCheckService.cs` performs the homepage update check without blocking Outlook startup.
- **UI**
  - `UI/SettingsForm.cs` configures base URL, authentication, sharing defaults, IFB, and debug logging.
    The `General`, `Sharing`, `Talk`, `Signature`, `Ifb`, `Advanced`, `Debug`, and `About` partials own their tab layout and actions. Control construction, shared save validation, backend policy state, and cross-tab control updates remain in the main form; the existing `Language` partial keeps localization.
  - `UI/TalkLinkForm.cs` is the Talk wizard.
  - `UI/FileLinkWizardForm.cs` is the sharing wizard.
  - `UI/BrandedHeader.cs` is the shared header banner control and provides `AttachToParent(...)` for consistent form header setup.
  - `UI/ScaledForm.cs` centralizes `ScaleLogical(...)` so form-level DPI wrappers are not duplicated.
- **Shared utilities**
  - `Utilities/BrowserLauncher.cs` centralizes shell target starts for files and directories; `OpenUrl` rejects non-HTTPS targets.
  - `Utilities/SizeFormatting.cs` centralizes adaptive byte, transfer-rate, and MB display formatting.
  - `Utilities/ComInteropScope.cs` centralizes COM release/final-release patterns.
  - `Utilities/PasswordGenerationHelper.cs` centralizes password-policy min-length resolution, server-policy generation fallback, and shared minimum-length validation for Talk/FileLink forms.
  - `Utilities/FileLinkPath.cs` centralizes FileLink path normalization, combination, naming, sanitization, and depth calculation.
  - `Utilities/HtmlTemplateSanitizer.cs` applies a Thunderbird-aligned HTML policy for backend templates and fails closed if sanitization cannot be applied.

### Runtime configuration and policy processing

- `Settings/SettingsStorage.cs` selects a profile-specific XML file below `%LOCALAPPDATA%\NC4OL`, applies defaults for missing values, and protects the app password with Windows DPAPI in `CurrentUser` scope. `SettingsFileTransaction` writes and validates a temporary file under a cross-process profile mutex, then replaces the primary file and retains the previous valid file as `.bak`.
- A malformed password value clears only the password and blocks background settings writes. A malformed primary XML falls back to the valid backup. When no valid file remains, automatic writes stay blocked until an explicit Settings save succeeds. Runtime settings change only after that user-initiated save commits.
- `SettingsWorkflowController` persists the candidate configuration before applying runtime, TLS, or IFB changes. A write failure displays the localized save error and reports an unsuccessful save without changing runtime settings.
- `Settings/ManagedSetupPolicy.cs` reads `HKLM` before `HKCU` and, on 64-bit Windows, the 64-bit registry view before the 32-bit view. An unlocked URL fills an empty profile; a locked URL overrides the profile value.
- `Utilities/NextcloudUriValidator.cs` accepts only HTTPS Nextcloud base URLs without user information, query, or fragment. Login-flow and password-policy URLs supplied by a server must keep the configured scheme, host, and port.
- `Services/BackendPolicyService.cs` reads the optional backend status for Settings, Talk, FileLink, managed-signature, and saved-appointment deletion flows. Share and Talk can resolve to local values when the backend or seat is unavailable. The managed-signature send gate uses the stricter policy state described in the signature flow above.
- The TLS setting is applied through `ServicePointManager.SecurityProtocol`. Before that assignment, `TransportSecurityConfigurator` sets the .NET switches for system-default TLS and strong cryptography in code, so the choice does not depend only on `NcTalkOutlookAddIn.dll.config` inside Outlook's AppDomain. Connection tests and login-flow diagnostics request a fresh connection through `NcHttpClient`, so a changed TLS mode is tested with a new handshake instead of an existing pooled connection.
- `app.config` maps the versions of the bundled HtmlSanitizer dependencies. The runtime resolver handles only matching requests from the add-in and this dependency stack in the add-in directory; higher versions, different tokens or cultures, and unrelated requesters remain untouched. `tools/ci/Check-VendorAssemblyBindings.ps1` loads every transitive vendor reference in a fresh .NET Framework AppDomain with these redirects.

### License status presentation

#### Enterprise Rollout

`ManagedSetupPolicy` records registry-value presence separately from parsed values. Any present `NextcloudUrl`, `NextcloudUrlLocked`, or `ShowMainRibbonTab` activates rollout, including a lock value of `false`. `NextcloudUrl` and `ShowMainRibbonTab` resolve independently by hive/view precedence; the URL lock still belongs to the selected URL entry, not a lock-only entry in another hive. `AddinSettings` carries `IsEnterpriseRollout` and `ShowMainRibbonTab` as runtime-only fields. Clones preserve them, XML never writes them. Rollout instructions and migration impact are in [ADMIN.md](ADMIN.md#enterprise-rollout).

Only `ShowMainRibbonTab=false` hides the main tab and its Settings button and blocks the public Settings callback. `IsEnterpriseRollout` controls access checks, not settings visibility. With the tab visible, `SettingsForm` exposes all settings and saves non-authentication preferences normally. Share and Talk use the internal `OpenAuthenticationSettingsAsync` path for initial authentication or credential recovery. With the tab hidden, the form exposes only General and saves only credentials/authentication changes; the existing workflow still persists before applying runtime settings.

`OpenAuthenticationSettingsAsync` returns a successful-save result. It skips unauthenticated backend/address-book prefetch, starts `SettingsForm.BeginAuthentication`, and requires verified credentials before saving. Pending setup uses a connection invitation; typed authentication rejection uses reauthentication guidance, not a backend or Seat diagnosis. Credential edits require verification again. Transport failures retain their diagnostic path. Both launch controllers retain the original Outlook item across setup and resume only after persistence and runtime application succeed. `MailInteropController.IsItemOpenForRibbonAction` checks live Inspectors by COM identity or the active inline reply on the Outlook STA; it never reopens a closed item. Cancelled setup or a closed original item ends the action. Share recovery retries only its read-only prefetch once, not a wizard or mutation; attachment automation does not trigger interactive sign-in.

`PolicyUiHelper.GetEnterpriseRolloutNotice` combines confirmed backend availability with the existing personal Seat check, without a Community/Pro filter. Launch controllers check it before queue adoption or room creation. Attachment and signature processing disable their managed actions when access is unavailable without introducing an ordinary Outlook send restriction. Talk background updates and IFB requests check access before remote operations. Cleanup and previously accepted separate-password work retain their existing lifecycle.

`FetchEnterpriseRolloutPolicyStatus` uses the existing account-scoped policy snapshot through `StoreBackendPolicySnapshot`: failed refreshes retain a matching confirmed response, fresh successful refusals replace it, and failed refreshes do not advance the success timestamp. No persistent rollout flag, extra license expiry calculation or vendor change is introduced. Production-assembly policy tests cover registry precedence, ribbon callbacks, authentication-only UI, Seat parity, automation and cache transitions.

`AddinSettings.ResolvePolicyDefaults` is the common Share/Talk defaults resolver for Settings, both wizards, generated block languages, attachment automation and saved-room deletion. Resolution is field-specific: an active locked backend field wins, otherwise an explicit local value, otherwise an active editable backend default, otherwise the product default. The settings model records the presence of local choices, not whether they differ from defaults; clones copy that state independently. XML omits untouched fields and loading existing values records them as explicit choices. Settings records edits while suppressing programmatic overlays; a credentials-only save must not write displayed backend defaults into local preferences. Runtime resolved copies are never persisted.

`BackendPolicyService.ParseStatus` normalizes only an explicit legacy backend `share_expire_days=0` to one day, without mutating the incoming dictionary. Missing/null values and local expiration choices remain distinct. Attachment threshold `null` disables automation, legacy zero resolves to five MB and current positive thresholds remain within 1–10240 MB. The always-share switch takes precedence over threshold mode. `Invoke-OutlookPolicyMappingTests.ps1` tests raw persistence, real WinForms controls and the operative attachment resolver; stale successful policy snapshots retain their existing refresh behavior.

A response without a `status` object is a failed fetch (`invalid_payload`), not a confirmed missing seat. It uses the backend-unavailable notice and cannot replace a cached successful response. Optional license metadata remains optional for older backends.

`BackendPolicyService.ParseStatus` normalizes both plain and OCS-wrapped status responses. `BackendPolicyStatus` retains the optional `license_status`, `access_status`, `can_manage_license`, `grace_until_iso`, `license_activation.state`, `license_connection_error`, `license_last_sync_at_iso` and `license_offline_until_iso` metadata. Missing fields default to empty/false; only a JSON boolean `true` enables the license-management action.

`PolicyUiHelper` selects localized notices at display time for Settings, FileLink and Talk, and reuses the same causes in disabled-feature tooltips. It distinguishes license refusal from `seat_state=suspended_overlimit`; unknown seat states receive a generic seat message. Old backends without explanatory fields receive a generic access warning. Grace and connection notices use a yellow status style without altering feature access. Dates supplied by the backend are formatted in local time for display only, never evaluated as client-side entitlement deadlines.

An administrator without a seat receives the license diagnosis plus the local-use/no-seat notice in the banner. The disabled-feature tooltip names the missing personal seat instead. Grace wording describes availability for active assigned seats, not for every administrator. Offline-expiry notices include known synchronization dates even without a current connection-error flag.

Community and Pro use the same personal access checks. Global `overlicensed` does not disable active seats or their policy domains; `suspended_overlimit` does. `Invoke-OutlookPolicyMappingTests.ps1` checks 384 paired mode/access/seat/capacity/role/synchronization combinations.

The administration link is built from the validated configured HTTPS base URL, preserving its installation subpath, and targets `/index.php/settings/admin/ncc_backend_4mc`. It is offered only for license notices to users with `can_manage_license=true`. This is a UI hint, not authorization; the backend still enforces permissions. `WarningPanelUiHelper` sizes panels with or without an action link.

Access remains determined by the existing `is_valid`, `seat_assigned`, `seat_state` and policy-domain checks. No client-side activation, licensing-server request, background timer, COM-thread change or signature-cache change is introduced. `Invoke-OutlookPolicyMappingTests.ps1` covers parser compatibility, status/seat precedence, role-specific actions and unchanged access decisions.

### End-to-end flows

#### Talk link flow (appointments)

1. User clicks **Insert Talk link** in an appointment.
2. `UI/TalkLinkForm.cs` collects: title, password, lobby, listable flag, room type, participant sync options, optional delegation target.
3. `Controllers/TalkRibbonController.cs` prefetches backend policy status and password policy in parallel (`Task.WhenAll`) before opening the wizard.
   - A delegation target is rejected as self when it matches the canonical UID, configured login, or known primary email. Directory selections remain keyed by canonical UID.
4. Before the server request, the add-in captures the current subject, location, body and all `X-NCTALK-*` properties. `Services/TalkService.cs` then creates the new room via OCS while an existing room remains available.
5. `Controllers/TalkAppointmentController.ApplyRoomToAppointment(...)` (invoked by `NextcloudTalkAddIn`) updates the appointment:
   - `Location` (Talk URL)
   - a localized plain-text body block (incl. password and help URL)
   - persisted metadata as Outlook `UserProperties` (including `X-NCTALK-*` keys)
   - backend-provided custom Talk templates are sanitized before rendering (no raw HTML fallback)
   - talk appointment HTML is passed through an explicit compatibility transform (`HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(...)`) before insert
   - appointment HTML insert uses the HTML->RTF bridge (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`), not `AppointmentItem.HTMLBody` and not `HTMLEditor.body.innerHTML`
6. Only after all appointment writes succeed, the new runtime subscription is registered and an existing room is retired. If a write fails, the captured appointment state is restored and the newly created room is deleted; failed cleanup is placed in the unconditional durable deletion queue. If retiring the old room fails after a successful replacement, the new room stays attached and deletion of the old room is queued for retry.
7. The runtime subscription (`AppointmentSubscription` in `NextcloudTalkAddIn.AppointmentSubscription.cs`) then handles appointment events:
   - **Write** captures the required Outlook values on the STA thread. `TalkAppointmentSyncCoordinator` coalesces the immutable snapshots and performs lobby, description, participant, and delegation requests in the background.
   - If Outlook exposes the final changed start time only shortly after `Write`, a short deferred post-write capture reads that same opened appointment instead of scanning calendars.
   - **Close** of a newly created, unsaved appointment queues orphan-room cleanup.
   - **BeforeDelete** uses Outlook's appointment-specific deletion event. Organizer, token, delegation, and recurrence checks run on that appointment before `QueueSavedTalkRoomDeletion(...)` creates a deletion job with `PolicyRequired=true`; URL/location parsing is not a deletion source. The background worker resolves the effective `TalkDeleteRoomOnEventDelete` policy before deleting the room. The same event path covers deletion from an open appointment and from the calendar view.
   - `Explorer.SelectionChange` rebinds only selected Talk appointments. The current selection is processed once when each Explorer is hooked, so calendar-view deletion also works immediately after an Outlook restart without opening the appointment.
8. Startup initializes only the persistent deletion retry worker and hooks existing Explorer surfaces. It does not enumerate Outlook stores or calendar folders, scan calendar items, or retain folder-level `Items` subscriptions.
9. The DPAPI-protected deletion queue uses a primary file and backup. Nextcloud deletion runs in the background, and queued failures are retried after delay and after an Outlook restart. Cleanup of a newly created room from an unsaved, discarded appointment uses the same queue with `PolicyRequired=false`. When older state is loaded, only records already marked for deletion survive; tracking-only records are discarded.

#### Talk appointment-safe HTML subset (backend custom templates)

For stable rendering in Outlook appointment bodies (Word/RTF pipeline), backend Talk templates should stay within this subset:

- Table-first layout (`table`, `tbody`, `tr`, `td`) for structure.
- Inline styles are allowed, but NC Connector strips known Word-unreliable declarations for appointment rendering:
  - `display:flex|grid`, `flex*`, `grid*`, `border-radius*`, `overflow*`, `object-fit`, `user-select` (vendor-prefixed variants included).
- Color/alignment fallback is injected automatically during appointment compatibility transform:
  - `style=color` -> `<font color=...>`
  - `style=background-color` -> `bgcolor`
  - `style=text-align` -> `align`
  - `style=vertical-align` -> `valign`
- Anchor color hardening: link color is additionally wrapped as `<a><font color=...>...</font></a>` where needed.
- Unsupported/unsafe tags/attributes are still removed by the sanitizer (fail-closed policy).

#### Sharing flow (mail compose)

1. User clicks **Insert Nextcloud share** while composing an email.
2. `UI/FileLinkWizardForm.cs` collects sharing settings and groups local and **My Nextcloud** selections in one queue. Its folder tree is built from immutable selection snapshots. Interactive local scans run as cancellable background filesystem work; the completed snapshot is applied only after control returns to the wizard's captured WinForms context. The upload planner consumes the same snapshots, so files added later are not shared invisibly and a removed or changed queued file stops the upload as a changed source.
3. `Controllers/FileLinkLaunchController.cs` loads the required capability snapshot, backend policy status, and password policy in parallel (`Task.WhenAll`) before opening the wizard.
   - The wizard starts with the persisted FileLink defaults. A locked share-policy value overrides its local counterpart; an editable value leaves the saved Outlook setting unchanged.
4. `Services/FileLinkService.cs` orchestrates the upload and public-share flow through the dedicated planner, DAV, transfer, share, and progress components.
   - Queue admission uses `FileLinkSelection.IdentityComparer`: local Windows paths ignore case; Nextcloud paths preserve case. Selecting the same source path again is ignored, but distinct Nextcloud names such as `report.pdf` and `Report.pdf` both enter the queue. Destination name conflicts still use the existing rename resolution, independently of source identity.
   - `FileLinkQueueSnapshotBuilder` scans each local selection once when it enters the queue. Dialog and drag-and-drop selections run that recursive scan off the UI thread with cancellation; only the individual files materialized from Outlook attachments use the synchronous initial handoff. The root-relative snapshot preserves empty directories, rejects symbolic links and junctions, and captures file size and modification time. `FileLinkSelectionScanner` consumes that exact snapshot and `FileLinkUploadPlanner` assigns transfer modes without enumerating the selected folder again or touching the server.
   - `NextcloudFilePickerForm` browses the configured user's file space with depth-one DAV `PROPFIND` requests. Its address bar keeps a local back and forward history, navigates to parents or breadcrumb targets without repeating the account name, and refreshes the current listing without adding a history entry. The root uses the same Nextcloud symbol as Thunderbird. For a selected file, the picker first requests a 1024 × 1024 preview from Nextcloud's authenticated `/index.php/core/preview.png` route. The response is limited to 5 MiB, decoded off the UI thread, and accepted only within the decoded-image limits. If Nextcloud reports that no preview is available, supported raster images of at most 5 MiB fall back to a byte-limited DAV `GET` of the original; document originals are never downloaded for previews. Stale preview requests are cancelled. Confirming a folder records its complete descendant snapshot, including empty folders. Selected files are planned as server copies; the transfer service checks their current size and sends authenticated DAV `COPY` requests into the reserved share folder. Originals remain unchanged, and file content does not pass through Outlook for the transfer itself.
   - When the user leaves the first manual wizard step, a depth-zero DAV `PROPFIND` checks the target derived from the base path, the wizard's fixed date, and the sanitized share name. An occupied target keeps the wizard on that step. `FileLinkDavClient` still reserves the share root later with an atomic `MKCOL`, so a collision created after the preflight stops the upload safely. A `405` after an indeterminate first result counts as a successful reservation only when a depth-zero DAV `PROPFIND` confirms the exact path as a collection. A known `405` without an earlier indeterminate result remains a collision. Attachment automation skips the preflight and continues to try numbered names. Empty directories, parents needed by bulk or chunked transfers, and Direct parents shared by multiple files are created once, parent first, with at most three parallel requests per level. Single-file Direct path chains are created by `X-NC-WebDAV-Auto-Mkcol`.
   - `FileLinkTransferService` coordinates dedicated bulk, direct, and chunked uploaders. Non-bulk files up to 20 MiB use direct WebDAV `PUT` and the server-side `X-NC-WebDAV-Auto-Mkcol: 1` header. Larger files use Nextcloud chunked upload v2 under `/remote.php/dav/uploads/<user>/<upload-id>` and are assembled with `MOVE .file`. Direct and chunked files share the limit of three concurrent transfers.
   - When the typed capability snapshot exposes `dav.bulkupload = "1.0"`, at least 20 candidate files of at most 8 MiB can be packed into sequential multipart batches of at most 100 files and about 20 MiB. The planner selects bulk only when the batch plan saves at least 20 percent of all upload requests, counting base-path and share-root creation, planned directories, direct files, and every chunk-folder, chunk-`PUT`, and final `MOVE`. Before the first server change, sequential MD5 preparation reports its completed and total file count as a separate wizard phase.
   - After all transfers finish, `FileLinkShareClient` sends one OCS create-share `POST` with path, explicit permissions, password, expiration date, label, and note. It omits the legacy `publicUpload` parameter because Nextcloud would use it to replace the explicit permission mask. No metadata update request follows.
   - A missing response, a transient gateway/service response without an OCS result, or a successful response without usable share data makes the create result ambiguous. `FileLinkShareClient` records that path and performs an exact-path OCS lookup with child shares disabled before another create request. It reuses a matching public-link share, retries only after a confirmed empty result, and keeps blocking duplicate creation while the lookup result remains unknown.
   - Replay-safe `MKCOL`, direct `PUT`, chunk `PUT`, and bulk `POST` operations receive at most two retries for transport failures and selected transient HTTP responses. Each bulk retry rebuilds the same request body from the unchanged local plan. A final chunk `MOVE` is never sent twice blindly: after an indeterminate transport result, an exact depth-zero DAV probe accepts the transfer only when the target is a non-collection resource with the expected length.
   - The wizard shows scan, checksum, folder-preparation, and aggregate transfer progress. Intermediate checksum and transfer updates are limited to at most ten per second. Debug logs contain plan, retry, five-second aggregate progress, and completion records instead of per-file success noise.
5. `Utilities/FileLinkHtmlBuilder.cs` generates the HTML block (header + link + password + permissions + expiration date).
   - The trusted local HTML uses the transparent `header-transparent-164x48.png` resource over the single blue header-cell background. Word-safe table-cell padding and a 124 px label column preserve the compact Thunderbird layout in Outlook Classic. Field labels and expiration dates are encoded as visible no-break tokens, while permissions use nested presentation tables with fixed 14 px status cells so Outlook cannot stretch the icons.
   - backend-provided custom share templates are sanitized via `HtmlTemplateSanitizer` and fail closed on sanitizer errors.
   - Before placeholder replacement, custom templates remove the nearest block wrapper for optional values that are empty. Fixed labels such as `Password` therefore do not remain without a value in HTML or plain-text output.
   - `Models/AttachmentLinkTargetPolicy.cs` resolves `policy.share.attachment_link_target` (`zip_download` / `share_page`) against the nullable local setting. An invalid stored local value is treated as unset, so a valid editable backend value can seed it. ZIP is used when no valid local or usable backend value exists; a locked backend value wins.
   - `AttachmentMode` controls read-only permissions, rights-row suppression, and cleanup. The explicit attachment link target controls only the URL plus `{LINK_INTRO}` and `{LINK_LABEL}`. Manual shares always render the Nextcloud share page. Legacy templates without these placeholders keep their existing output.
   - ZIP URL derivation is fail-closed: the public absolute HTTP(S) URL must end in `/s/<token>` and match the OCS token. Invalid input throws before insertion; there is no original-URL fallback.
   - custom Share rendering prefers `policy.share.share_html_block_template_v2` and falls back to `policy.share.share_html_block_template`. This supports older backend releases while allowing current backends to keep the original response key placeholder-free for older clients.
   - current backends expose `policy.share.share_html_block_effective_language` for custom templates. Outlook uses it for generated link wording, field labels, permission names, and password hints; older backends without the field keep the previous UI-language fallback.
   - Custom-template placeholder values remain context-neutral because the same variable may occur in visible text or an attribute. No-break markup belongs in known visible template positions, not in generic substitution values.
   - plain-text compose keeps `MailItem.BodyFormat=olFormatPlain`; the share block is rendered as a framed text block with `#` separators and inserted through Outlook WordEditor. Inline replies/forwards keep two empty paragraphs above the block for the sender's own text. `MailItem.Body` is not rewritten.
6. `NextcloudTalkAddIn.TryInsertHtmlIntoMail(...)` / `TryInsertPlainTextIntoMail(...)` return the insertion result from `Controllers/MailBodyInsertionController.cs`. HTML compose uses WordEditor first so existing managed bookmarks stay intact; a direct `HTMLBody` write remains the compatibility fallback when the Inspector editor cannot be opened. If every insertion path fails, `FileLinkLaunchController` queues the newly created server artifacts for cleanup and reports the wizard as failed.

Compose runtime in `NextcloudTalkAddIn.cs` (`MailComposeSubscription`) delegates remote cleanup to `Services/ComposeShareCleanupService` and password delivery to `Controllers/SeparatePasswordDeliveryController`. Cleanup runs in the existing background task; password delivery stays on the Outlook STA in the primary mail's Send event:

- The FileLink ribbon entry is exposed in mail inspectors and in the Explorer inline reply/forward `Message` tab. Both entries call the same `FileLinkLaunchController` path.
- Inline replies/forwards insert the rendered share HTML through `Explorer.ActiveInlineResponseWordEditor`; the inline path does not rewrite `MailItem.HTMLBody` and keeps two empty paragraphs above the share block for the sender's own text.
- Debounced attachment evaluation (`ComposeAttachmentEvalDebounceMs`) after compose attachment changes.
- Attachment partials share one compose subscription and its existing event registrations and lifetime. Policy, materialization, and queue work do not introduce independent event handlers or task dispatch paths.
- Open compose windows invalidate their attachment-rule snapshot immediately after local settings are saved. After five minutes, attachment events refresh backend policy in the background while continuing with the last resolved rules. Send uses the same snapshot without a refresh warning; an unsuccessful refresh retains the previous rules and does not mark them fresh. Share and Talk wizards retain the policy loaded when they opened for their entire operation.
- Only a missing initial snapshot or a settings change leaves Send waiting for its initial rule lookup. The existing local-settings fallback remains when that first lookup cannot reach the backend. A send attempted before it finishes is cancelled with an informational retry notice; no STA-thread wait or automatic resend is added. An effective `AlwaysConnector` rule still blocks ordinary attachments with a required-routing warning. Neither notice claims an upload failed.
- Attachment automation modes:
  - always route attachments into NC sharing flow, or
  - threshold mode with a two-action prompt (`Share with NC Connector` / `Remove last selected attachments`).
- For a multi-file addition, the prompt pairs the last file's name with that file's size while the remove action still covers the complete added batch.
- Pre-add attachment interception:
  - `BeforeAttachmentAdd` path resolves candidate file metadata early
  - can best-effort cancel host attachment add and launch NC sharing before Outlook post-add handling.
  - an enforcing always-via-NC policy cancels the host add when the candidate cannot be materialized or evaluated.
  - hard Outlook/Exchange size blocks can still happen before add-in callbacks and are not interceptable via official Outlook OOM events.
- Runtime host guard checks (live large-attachment setting) at:
  - pre-evaluation
  - pre-prompt-action handling
  - wizard finalize (enforced in `UI/FileLinkWizardForm.cs` via `Services/OutlookAttachmentAutomationGuardService.cs`).
- Attachment-mode wizard launch:
  - after server prefetch, materializes the current compose attachments on the Outlook STA and queues them as initial wizard selections; capture, queue acceptance and host removal stay in one synchronous UI call, so attachment positions cannot become stale across the network wait
  - removes the Outlook attachments only after the complete initial queue has accepted them
  - does not restore adopted attachments when the user later cancels the wizard
  - opens directly in file-step-equivalent mode.
  - copies the effective attachment link target into `FileLinkRequest`; no per-share target switch is exposed.
- Outlook body resources with `PR_ATTACHMENT_HIDDEN=true`, such as signature images, are excluded from attachment batching, threshold totals, FileLink selection, host removal, and the required-routing send gate.
- `UI/FileLinkWizardForm.cs` file-step queue accepts Explorer drag & drop for files/folders across queue and action-area controls.
- Compose insertion and separate-password lifecycle:
  - `ComposeLifecycleOrigin` retains the exact server/account origin needed to delete the created share or issue a later Secrets request. Cleanup never falls back to the currently selected account.
  - if a newly created share cannot be inserted into the message, the controller attempts to delete its server folder with that captured origin.
  - after successful insertion, `MailComposeSubscription` tracks its `ComposeShareCleanupRecord` until Outlook raises `AfterWrite`. A completed write covers Save, AutoSave, and the write performed for Send/Outbox, so those paths release the cleanup record without deleting the share.
  - classic compose windows bind the concrete `InspectorEvents_10.Close` event. It fires only when that Inspector actually closes; if no successful write followed the insertion, the subscription queues deletion using the captured account and relative path. A cancelled close therefore keeps the share state armed.
  - inline compose keeps `Explorer.InlineResponseClose` as a surface-transition signal because Outlook also raises it for pop-out and navigation. `ItemEvents_10.Unload` remains the terminal item signal for an inline response and evaluates only previously captured cleanup state without reading the unloaded `MailItem`.
  - DAV cleanup uses the shared bounded FileLink retry path. A repeated delete remains idempotent because an already absent folder is accepted.
  - cleanup tracking is held in memory. Once Outlook has written the message, deleting that saved draft later or after an Outlook restart does not delete the share; the unused share must be removed manually.
  - `RegisterSeparatePasswordDispatch` initially keeps the password follow-up data only in the subscription's `_passwordDispatchQueue`. Save and AutoSave do not serialize that queue into the primary mail. Closing the compose window disposes the subscription, so a reopened draft, an Outlook restart before the first send attempt, or a new message created from an `.oft` template cannot restore the follow-up state.
  - `OnSend` is the direct dispatch boundary. The queue is consumed once, final Secrets/plain content is produced, the primary mail's resolved recipients and effective `SendUsingAccount`/`SentOnBehalfOfName` are applied, body and matching backend signature are finalized, recipients are resolved, and the follow-up is submitted without saving an intermediate Outlook draft.
  - no Drafts, Outbox, Sent-folder, MIME-marker, or restart-recovery path participates in password delivery. Outlook delayed/offline delivery still invokes `OnSend`, so the password follow-up is submitted immediately and does not wait for the primary item to leave the Outbox. Unexpected follow-up failures never set the primary send's `cancel` flag.
  - the send gate also cancels an enforcing attachment policy while an ordinary attachment still violates it.
- Separate password-mail dispatch:
  - queue password-only content after share creation and persist the origin settings needed by later Secrets requests
  - capture the successful share-time signature policy/settings snapshot before the queue entry is accepted
  - capture resolved recipients and sender identity on send
  - deduplicate resolved SMTP addresses across To, Cc, and Bcc before Secrets creation
  - when backend policy requests Nextcloud Secrets, create one one-time Secrets link per unique final recipient
  - Secrets links are encrypted locally with AES-GCM through Windows CNG; no new crypto dependency is bundled
  - if Secrets creation fails, fall back to the existing plain separate password mail and warn the user
  - preserve the source compose mode for HTML vs plain-text follow-up mails and submit directly from the primary `Send` callback without calling `Save`
  - on a definite automatic-send failure, open one fully prepared manual fallback; when Outlook's submission state is ambiguous, do not create or send a duplicate
  - if strict sender or recipient preparation fails before submission, open a fresh manual fallback with normalized To/Cc/Bcc strings and reconcile the managed signature after the Inspector is initialized.

#### IFB flow

1. User enables IFB in Settings.
2. `Services/FreeBusyManager.cs` creates a random request secret for each Outlook process. `Services/IfbRegistryStateStore.cs` keeps the DPAPI-protected ownership state with primary/backup recovery.
3. `Services/FreeBusyServer.cs` starts a local HTTP listener on the configured IFB port (`Settings -> IFB -> Local IFB port`, default: `7777`), accepts only `/nc-ifb/<request-secret>/freebusy/<address>.vfb`, and caps concurrent proxy requests at four.
4. `Services/IfbRegistryOwnershipManager.cs` records each original user value and registers `%NAME%@%SERVER%.vfb` below the secret endpoint. Outlook replaces both placeholders with the attendee's full SMTP address. Policy values are read for conflicts but are not written.
5. During upgrade, an unowned legacy `/nc-ifb/freebusy/...` value is adopted only in its pre-token form. An unowned tokenized path remains blocked. Without a saved external predecessor, disabling IFB removes the stale legacy value rather than restoring it.
6. On disable, an original value is restored only if the current value still equals the value written by NC Connector.
7. `IfbAddressBookCache` scopes cached identities by Outlook profile, normalized Nextcloud base URL including subpath, and canonical Nextcloud UID.

## Network endpoints

The add-in uses Nextcloud **OCS** and **WebDAV** endpoints.

`NextcloudUriValidator` normalizes the configured base URL before authenticated service construction. Runtime configuration rejects explicit HTTP, user information, query strings, and fragments. Absolute login-flow and password-policy endpoints returned by Nextcloud must use HTTPS and match the configured scheme, host, and port.

Authentication aliases and DAV identities are intentionally kept separate. Basic Auth uses the login entered by the user (which may be an email address), while `Services/NextcloudUserIdentityService.cs` resolves the canonical UID from `GET /ocs/v2.php/cloud/user?format=json`. User-scoped FileLink, CardDAV, and CalDAV paths use only `ocs.data.id`; a missing UID is treated as an error rather than silently substituting the login.

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

- Capabilities and required server version: `GET /ocs/v2.php/cloud/capabilities?format=json`
- Current canonical user ID: `GET /ocs/v2.php/cloud/user?format=json`
- Create public share: `POST /ocs/v2.php/apps/files_sharing/api/v1/shares`
- Upload/folder creation: `remote.php/dav/...` (WebDAV)
- Browse the user's files and quota: depth-one `PROPFIND /remote.php/dav/files/<user>/...`
- Preview a selected file: authenticated, byte-limited `GET /index.php/core/preview.png?file=...`; a missing generated preview falls back to `GET /remote.php/dav/files/<user>/...` only for supported raster images whose originals are at most 5 MiB
- Copy a selected Nextcloud file into the share: `COPY /remote.php/dav/files/<user>/...` with an absolute same-account `Destination`
- Optional small-file bulk upload: `POST /remote.php/dav/bulk` (`multipart/related`, only when `ocs.data.capabilities.dav.bulkupload` is exactly `"1.0"`)
- Large file upload: `MKCOL /remote.php/dav/uploads/<user>/<upload-id>`, chunk `PUT`s, then `MOVE /remote.php/dav/uploads/<user>/<upload-id>/.file` to the final file path

Secrets (optional separate password mode):

- Create encrypted secret: `POST /ocs/v2.php/apps/secrets/api/v1/secrets`
- Public one-time link: `/index.php/apps/secrets/share/<uuid>#<local-key>`
- The key stays in the URL fragment and is not sent to Nextcloud.

IFB (DAV via proxy):

- Reserved listener namespace: `http://127.0.0.1:<ifb-port>/nc-ifb/` (default `<ifb-port>=7777`)
- Accepted Outlook path: `/nc-ifb/<request-secret>/freebusy/<address>.vfb`; the secret is generated for each Outlook process and is not persisted. DPAPI-protected profile state stores registry ownership only
- Requests without the request secret return `404`
- The proxy talks to CalDAV and Addressbook endpoints under `remote.php/dav/...`

System address-book response contract:

- `Services/IfbAddressBookCache.cs` validates the complete CardDAV export before replacing any cached mappings or the last-success timestamp. Its cache remains scoped to the Outlook profile, Nextcloud base URL and configured login; the request path uses the resolved canonical UID.
- Only successful HTTP responses and HTTP 404 enter vCard validation. A complete, non-empty export with usable user UIDs is accepted regardless of content type, including on HTTP 404. Other HTTP errors remain errors even when the body resembles a valid export; the shared HTTP client is unchanged.
- Empty exports require both a successful HTTP status and a vCard content type. Reject HTML/JSON, truncated or mixed exports, nested cards and stray or mismatched component boundaries. Preserve legal line folding and UID-only contacts for user search and moderators.
- A failed refresh keeps the last good cache intact but marks the current scope for retry. Subsequent lookups must not hide the failure behind that old cache. Talk participant synchronization must resolve all recipients before applying changes and must not fall back to guest invitations on a failed address-book lookup.
- Keep response bodies and contact data out of diagnostics. `Invoke-OutlookTalkIfbLifecycleTests.ps1` covers response validation, scoped cache recovery and participant synchronization through the production paths.

Update check:

- Homepage endpoint: `GET https://nc-connector.de/wp-json/ncc/v1/update-check`
- Query values: `product=outlook`, installed version, channel, and a daily rotating client hash.
- Release and download targets from the response are retained only when they use HTTPS below `github.com/nc-connector/NC_Connector_for_Outlook/releases/`. Other values are discarded.
- `UpdateAvailable` is calculated locally from the installed and reported versions.
- The homepage only returns release metadata and counts one anonymous client per day.

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
- Daily log file format: `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`
- Runtime exceptions are always written via `DiagnosticsLogger.LogException(...)`, even when debug logging is disabled.
- Retention: keep latest 7 daily log files and delete files older than 30 days (best effort cleanup).
- Authorization values, URL credentials, structured token/password fields, Talk/share path tokens, and Secret fragments are redacted before every log write, even when optional anonymization is off.
- Anonymization additionally redacts the configured NC URL/base host, user identifiers, email addresses, and local user path fragments.

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
- FileLink hot paths log upload plans, retries, periodic aggregate progress, and completion summaries, not every successful request.
- Never swallow exceptions silently.

## Compatibility & version checks

### Outlook bitness (x86 on x64 Windows)

Outlook can be installed as a 32-bit application on 64-bit Windows. In that case it reads COM add-in registration from the 32-bit registry view (`Wow6432Node`).

The MSI registers add-in keys for **both** registry views:

- 64-bit: `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`
- 32-bit: `HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

Installer definition:

- `installer/Product.wxs`

### Nextcloud feature detection

All add-in functions require Nextcloud **32 or newer**. `NextcloudCapabilitiesService` reads the authenticated OCS capabilities endpoint, validates the structured version, and caches a typed snapshot for five minutes per server/user pair. Connection checks refresh the snapshot; feature flows reuse it and reject older or unversioned responses.

The same snapshot controls optional DAV bulk upload. Bulk is active only when `ocs.data.capabilities.dav.bulkupload` is exactly `"1.0"`, at least 20 candidates are no larger than 8 MiB, and sequential batches of at most 100 files and about 20 MiB reduce the complete upload request count by at least 20 percent. That comparison includes base-path and share-root creation, planned directories, direct files, and every chunk-folder, chunk-`PUT`, and final `MOVE`. Direct upload remains available when any condition is not met.

Implementation:

- `src/NcTalkOutlookAddIn/Services/NextcloudCapabilitiesService.cs`
- `src/NcTalkOutlookAddIn/Models/NextcloudCapabilitiesSnapshot.cs`
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
3. Builds the WiX v6 installer (`installer/NcConnectorOutlookInstaller.wixproj`)
4. Copies the MSI into `dist/`

### Versioning

- `src/NcTalkOutlookAddIn/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`

`build.ps1` derives the MSI `ProductVersion` from that (format `Major.Minor.Build`).

### Release checklist

1. Bump version in `AssemblyInfo.cs`
2. If vendored dependencies changed: update `VENDOR.md`
3. `.\build.ps1 -Configuration Release`
4. Install/upgrade MSI test (old version → new version)
5. Smoke test (Talk + sharing + IFB)
6. Optional: sign the MSI (if required in your environment)

## Local testing

Run the automated checks from `tools/ci/` through the jobs defined in
`.github/workflows/outlook-build-checks.yml`, then use the smoke tests below for Outlook COM behavior.

Suggested smoke test sequence:

1. Enable debug logging in Settings.
2. Calendar: create a new appointment, insert a Talk link, save the appointment, then change start time and save again (lobby update).
3. Calendar: restart Outlook, open the same appointment, change start time and save again (persistent metadata + lobby update).
4. Calendar: add attendees, save again (participant sync).
5. Calendar: enable saved-event room deletion, then delete saved Talk appointments from an open appointment and from the calendar view. Verify that each associated room is removed. Repeat once with a temporary Nextcloud failure, restart Outlook, and verify the queued retry.
6. Mail: run the sharing wizard, upload 1–2 small files, insert the HTML block, and send to yourself.
7. Mail: insert a share and discard the message before it is saved; verify that the exact server folder is removed. Repeat with Save or AutoSave and verify that the share remains.
8. Mail: with separate password delivery, create the share and click Send from the same compose window. Verify that the final follow-up leaves through the same effective Outlook account immediately and that no NC Connector password draft remains. Repeat with delayed/offline delivery and verify the documented direct boundary: the follow-up does not wait for the primary Outbox item. A closed and reopened primary draft or an `.oft` template with an existing share block is outside the supported flow and requires a new share before sending.
9. IFB: enable IFB, verify the URL reservation and TCP listener, then use Outlook's Scheduling Assistant with an address whose domain differs from the configured Nextcloud login and host. A direct request without the request secret is expected to return `404`.
10. Settings -> Advanced: click `Check now` and verify that latest version, last check, download link, and changelog summary update without blocking Outlook.

## X-NCTALK-* property reference

The add-in persists Talk appointment metadata as Outlook `UserProperties` using `X-NCTALK-*` names only. The old NC Connector-specific Outlook property names are no longer written or read.

Unless stated otherwise:

- Properties are stored as **text** values in Outlook (`OlUserPropertyType.olText`).
- Boolean values are stored as `TRUE` / `FALSE` (uppercase).
- Timestamps are stored as **Unix epoch seconds** (UTC) in invariant culture.

Primary write location:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` -> `ApplyRoomToAppointment(...)`
- local Outlook metadata refresh: `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` -> `PersistCoreIcalProperties(...)`

### Properties

| Property | Purpose | Type / format | Example | Written | Read / used | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| `X-NCTALK-TOKEN` | Talk room token | `string` | `a1b2c3d4` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Required for saved-event room deletion and runtime subscription; generic Talk URLs in `Location`/URL fields are ignored. |
| `X-NCTALK-URL` | Talk room URL | `string` | `https://cloud.example.com/call/a1b2c3d4` | `ApplyRoomToAppointment(...)` | `RegisterSubscription(...)` | Stored as local Outlook metadata; not used as a deletion source. |
| `X-NCTALK-LOBBY` | Lobby enabled flag | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Used to decide whether lobby updates run on save. |
| `X-NCTALK-START` | Appointment start time (epoch seconds) | `int64` as string | `1739750400` | `ApplyRoomToAppointment(...)`, `AppointmentSubscription.OnWrite(...)` | `GetIcalStartEpochOrNull(...)`, `TryReadAppointmentStartEpoch(...)` | Local metadata for subscription state; lobby updates use the current save/deferred start epoch directly. |
| `X-NCTALK-EVENT` | Room creation mode marker | `event` \| `standard` | `event` | `ApplyRoomToAppointment(...)` | `GetRoomType(...)` | No legacy Outlook property fallback. |
| `X-NCTALK-OBJECTID` | Time-window identifier | `"<start>#<end>"` | `1739750400#1739754000` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Stored as local Outlook metadata. |
| `X-NCTALK-ADD-USERS` | Participant sync: internal users | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Split participant sync flag for Nextcloud users. |
| `X-NCTALK-ADD-GUESTS` | Participant sync: external emails | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Split participant sync flag for guests. |
| `X-NCTALK-DELEGATE` | Delegation target user ID | `string` | `alice` | `ApplyRoomToAppointment(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)`, `TryApplyDelegation(...)` | No legacy Outlook property fallback. |
| `X-NCTALK-DELEGATE-NAME` | Delegation target display name | `string` | `Alice Example` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Stored as local Outlook metadata. |
| `X-NCTALK-DELEGATED` | Delegation state marker | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)`, `TryApplyDelegation(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)` | Controls whether delegation is still pending. |
| `X-NCTALK-DELEGATE-READY` | Delegation “ready” marker | `TRUE` | `TRUE` | `ApplyRoomToAppointment(...)` | (not read by add-in) | Local marker retained for the Outlook delegation contract; the add-in currently uses `X-NCTALK-DELEGATED` + delegate ID to detect pending delegation. |

## Extension points

### Add a new setting

1. Add property to `src/NcTalkOutlookAddIn/Settings/AddinSettings.cs`.
2. Persist it in `src/NcTalkOutlookAddIn/Settings/SettingsStorage.cs`.
3. Add UI in `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`.
4. Add translations (see `Translations.md`).

### Add a new Nextcloud API call

1. Add to the appropriate service:
   - Talk: `src/NcTalkOutlookAddIn/Services/TalkService.cs`
   - Sharing orchestration: `src/NcTalkOutlookAddIn/Services/FileLinkService.cs`
   - Sharing DAV collections: `src/NcTalkOutlookAddIn/Services/FileLinkDavClient.cs`
   - Sharing transfers: `src/NcTalkOutlookAddIn/Services/FileLinkTransferService.cs`
   - Sharing OCS share creation: `src/NcTalkOutlookAddIn/Services/FileLinkShareClient.cs`
   - Use `Services/NcHttpClient.cs` + `Utilities/NcJson.cs` for new OCS/JSON calls instead of introducing service-local request/parsing helpers.
2. Add request/response model in `src/NcTalkOutlookAddIn/Models/` (if needed).
3. Add logging scopes and error handling.
4. Integrate in the UI/wizard and wire it up via `NextcloudTalkAddIn.cs`.
5. For ribbon-triggered flows, prefer adding orchestration in the matching controller (`SettingsWorkflowController`, `FileLinkLaunchController`, `TalkRibbonController`) and keep `NextcloudTalkAddIn.cs` as a thin delegate layer.

### Add a new localized string

1. Add a property to `src/NcTalkOutlookAddIn/Utilities/Strings.cs`.
2. Add the key to all locale files under `src/NcTalkOutlookAddIn/Resources/_locales/`.
3. Rebuild and verify the UI.
