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

## Release 3.1.0 delta summary

This release expands Outlook compose support and central backend signatures:

- Backend-managed email signatures apply to matching Outlook sender identities in HTML/RTF and plain-text compose, including replies and forwards.
- Nextcloud share insertion is available from inline replies/forwards and uses WordEditor insertion so quoted content stays intact.
- Plain-text share blocks are inserted without rewriting `MailItem.Body`.
- Large files use Nextcloud chunked WebDAV upload v2 and the sharing wizard shows per-file upload speed.
- Separate password follow-up mails keep the original sender identity, receive the backend signature when policy and sender match, and still open a manual fallback draft if auto-send fails.
- Talk room deletion for saved appointments remains opt-in and Talk cleanup metadata stays local to Outlook.

## Quick start

### Prerequisites

- Windows 10/11
- Outlook classic (x64 or x86)
- **.NET Framework 4.7.2** (target framework)
- MSBuild (e.g. Visual Studio Build Tools)
- **.NET SDK** (used by WiX v6 build via `dotnet`)
- **Nextcloud 32 or newer** (runtime server)

### Build MSI (recommended)

```powershell
cd "C:\\path\\to\\nc4ol"

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

- `dist\\NCConnectorForOutlook-<version>.msi`

### Install & run locally

1. Install the MSI (administrator rights required):
   - `msiexec /i dist\\NCConnectorForOutlook-<version>.msi`
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

Key code locations:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` — entry point, ribbon XML, Outlook event wiring, orchestration
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs` — add-in bootstrap/teardown lifecycle (`OnConnection`, shutdown/disconnect)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs` — dedicated Outlook event hook/unhook wiring helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs` — category-specific runtime logging helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs` — backend policy + Talk template/language resolver helpers
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs` — deferred appointment-subscription ensure and Outlook event-restriction handling
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs` — compose subscription core state + lifecycle entry points (`Dispose`, identity, shared helpers)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs` — compose attachment interception/evaluation/share-launch flow
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Signature.cs` — backend email-signature policy application for the matching Outlook sender account
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs` — send/close cleanup lifecycle + separate-password dispatch handling
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs` — appointment runtime subscription lifecycle
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs` — settings open/save/revert orchestration
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs` — FileLink ribbon launch + wizard orchestration
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs` — Talk ribbon flow orchestration (auth gate, wizard, room create/replace)
- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` — appointment lifecycle orchestration for Talk room metadata/sync
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` — compose share cleanup + separate-password dispatch flow
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` — Talk template/body block rendering
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` — SMTP and attendee recipient resolution
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` — compose-subscription registry lifecycle
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` — shared mail/Inspector interop and the unified WordEditor signature-slot reconciler
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` — shared model for separate password follow-up dispatch queue entries
- `src/NcTalkOutlookAddIn/Services/` — Nextcloud HTTP integrations (Talk, sharing, IFB, login flow)
  - `Services/NcHttpClient.cs` is the shared request executor for auth headers, OCS headers, timeout/decompression, and optional fresh-connection mode.
  - All runtime HTTP calls (Talk, share/DAV, IFB, login flow, moderator avatar fetch) are routed through `NcHttpClient`.
  - `Services/EmailSignaturePolicyService.cs` resolves backend email-signature policy values against local settings and lock state.
  - `Services/UpdateCheckService.cs` checks `nc-connector.de` once per day for Outlook release metadata and stores the cached result in profile settings.
- `src/NcTalkOutlookAddIn/UI/` — WinForms dialogs and wizards
  - `UI/ScaledForm.cs` is the shared DPI-scaling base for forms that use logical pixel layout helpers.
- `src/NcTalkOutlookAddIn/Settings/` — persisted settings model, storage, and managed setup policy
  - `Settings/ManagedSetupPolicy.cs` reads the managed Nextcloud URL from Windows policy registry keys.
- `src/NcTalkOutlookAddIn/Utilities/` — logging, theming, i18n, small shared helpers
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` — centralized sanitizer for backend-provided share/talk HTML templates
- `src/NcTalkOutlookAddIn/Utilities/HtmlToPlainTextConverter.cs` — DOM-based HTML-to-plain-text rendering for plain-text email signatures
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` — centralized JSON payload normalization (`PrepareJsonPayload`), dictionary/string/int helpers, and OCS error extraction
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` — encapsulated pending-key tracking + throttled logging state for deferred appointment ensure
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
- Table-based Outlook signatures are replaced only when `_MailAutoSig` is inside that table. Each successful insertion receives the `NcConnectorSignature` bookmark, including HTML, RTF, and plain text, so later updates and clears address only the managed range.
- Replacement is staged before the previous signature range is removed. When staging above an existing slot, a temporary Word bookmark tracks that old range while inserted content shifts its numeric offsets. If insertion, bookmark creation, tracking, or old-range deletion fails, the staged content is removed and the previous managed range is restored where possible.
- Cursor/selection preservation uses a temporary Word bookmark rather than stale absolute offsets. Safe fallback insertion adds only missing paragraph marks above the signature and one separator paragraph before quoted content.
- Sender and `BodyFormat` changes schedule another reconciliation. Changes observed while attachment processing is suppressing compose work, or while an inline item has no active surface, are deferred and resumed once a usable WordEditor returns.
- When policy becomes inactive or the sender no longer matches, NC Connector removes only `NcConnectorSignature`. It does not scan or rewrite arbitrary body content and does not remove a native signature from a non-matching identity.
- Signature processing only runs for unsent Outlook compose items. Opening a received or already sent message for reading must never modify its body.
- Before send, the pending debounce is stopped and the current sender, format, compose kind, policy, and managed slot are reconciled synchronously. With complete backend connection settings, sending is cancelled if no successful policy snapshot is available or if a required apply/clear operation cannot finish safely. The compose item stays open for correction and retry.
- An incomplete backend setup does not create a signature requirement; cleanup is limited to best-effort removal of an exact `NcConnectorSignature` bookmark. An unsupported signature domain disables insertion as well, but with otherwise complete backend settings an existing managed range must still reconcile safely at send time. An `InlineResponseClose` that arrives just before send does not by itself block a previously reconciled, unchanged message.
- Separate password follow-up dispatch reuses the successful policy snapshot already verified for the primary send, sanitizes it once, and uses it for the complete queue. It applies and reads back `SendUsingAccount`/`SentOnBehalfOfName`, auto-sends only when the effective follow-up identity equals the sender captured from the successful primary mail, and adds the backend signature only when that effective identity also matches `policy.email_signature.user_email`. Plain source mail produces a plain follow-up; HTML/RTF source produces an HTML follow-up. The same snapshot is used for a manual fallback draft.
- Debug logging records the trigger, active surface, body format, compose kind, slot source, and reconciliation result without writing the signature template or sender address.

## Architecture

### Main building blocks

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
  - `UI/TalkLinkForm.cs` is the Talk wizard.
  - `UI/FileLinkWizardForm.cs` is the sharing wizard.
  - `UI/BrandedHeader.cs` is the shared header banner control and provides `AttachToParent(...)` for consistent form header setup.
  - `UI/ScaledForm.cs` centralizes `ScaleLogical(...)` so form-level DPI wrappers are not duplicated.
- **Shared utilities**
  - `Utilities/BrowserLauncher.cs` centralizes shell target starts (URLs, files, directories).
  - `Utilities/SizeFormatting.cs` centralizes adaptive byte, transfer-rate, and MB display formatting.
  - `Utilities/ComInteropScope.cs` centralizes COM release/final-release patterns.
  - `Utilities/PasswordGenerationHelper.cs` centralizes password-policy min-length resolution, server-policy generation fallback, and shared minimum-length validation for Talk/FileLink forms.
  - `Utilities/FileLinkPath.cs` centralizes FileLink path normalization, combination, naming, sanitization, and depth calculation.
  - `Utilities/HtmlTemplateSanitizer.cs` applies a Thunderbird-aligned HTML policy for backend templates and fails closed if sanitization cannot be applied.

### Runtime configuration and policy processing

- `Settings/SettingsStorage.cs` selects a profile-specific XML file below `%LOCALAPPDATA%\NC4OL`, applies defaults for missing values, and protects the app password with Windows DPAPI in `CurrentUser` scope. Its migration path copies legacy INI values into the profile files and removes legacy files only after every target write succeeds.
- `Settings/ManagedSetupPolicy.cs` reads `HKLM` before `HKCU` and, on 64-bit Windows, the 64-bit registry view before the 32-bit view. An unlocked URL fills an empty profile; a locked URL overrides the profile value.
- `Services/BackendPolicyService.cs` reads the optional backend status for Settings, Talk, FileLink, managed-signature, and saved-appointment deletion flows. Share and Talk can resolve to local values when the backend or seat is unavailable. The managed-signature send gate uses the stricter policy state described in the signature flow above.
- The TLS setting is applied through `ServicePointManager.SecurityProtocol`. Connection tests and login-flow diagnostics request a fresh connection through `NcHttpClient`, so a changed TLS mode is tested with a new handshake instead of an existing pooled connection. Other runtime HTTP calls continue to use the shared request executor.

### End-to-end flows

#### Talk link flow (appointments)

1. User clicks **Insert Talk link** in an appointment.
2. `UI/TalkLinkForm.cs` collects: title, password, lobby, listable flag, room type, participant sync options, optional delegation target.
3. `Controllers/TalkRibbonController.cs` prefetches backend policy status and password policy in parallel (`Task.WhenAll`) before opening the wizard.
4. `Services/TalkService.cs` creates the room via OCS.
5. `Controllers/TalkAppointmentController.ApplyRoomToAppointment(...)` (invoked by `NextcloudTalkAddIn`) updates the appointment:
   - `Location` (Talk URL)
   - a localized plain-text body block (incl. password and help URL)
   - persisted metadata as Outlook `UserProperties` (including `X-NCTALK-*` keys)
   - backend-provided custom Talk templates are sanitized before rendering (no raw HTML fallback)
   - talk appointment HTML is passed through an explicit compatibility transform (`HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(...)`) before insert
   - appointment HTML insert uses the HTML->RTF bridge (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`), not `AppointmentItem.HTMLBody` and not `HTMLEditor.body.innerHTML`
6. A runtime subscription is registered for the appointment (`AppointmentSubscription` in `NextcloudTalkAddIn.AppointmentSubscription.cs`):
   - **Write** (save): updates lobby timer on time changes, updates room description, syncs participants, applies delegation
   - If Outlook exposes the final changed start time only shortly after `Write`, a short deferred post-write verification retries the lobby update with the newly observed start time on the same opened appointment instead of broad calendar scanning.
   - **Close** (discard without saving): deletes the room to avoid orphans (best-effort)
   - **BeforeDelete**: queues room deletion in the background only when saved-event deletion is opted in (`TalkDeleteRoomOnEventDelete` or locked backend `talk_delete_room_on_event_delete`) and the appointment has `X-NCTALK-TOKEN`; URL/location parsing is not a deletion source

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
2. `UI/FileLinkWizardForm.cs` collects sharing settings and the file/folder selection.
3. `Controllers/FileLinkLaunchController.cs` loads the required capability snapshot, backend policy status, and password policy in parallel (`Task.WhenAll`) before opening the wizard.
4. `Services/FileLinkService.cs` orchestrates the upload and public-share flow through the dedicated planner, DAV, transfer, share, and progress components.
   - `FileLinkSelectionScanner` scans the local selection once. The root-relative scan result preserves empty directories, rejects symbolic links and junctions, and captures file size and modification time. `FileLinkUploadPlanner` then assigns transfer modes without touching the server.
   - `FileLinkDavClient` creates the share root with one atomic `MKCOL`. A collision stops a manual share; attachment automation tries numbered names without a preliminary `PROPFIND`. Empty directories, parents needed by bulk or chunked transfers, and Direct parents shared by multiple files are created once, parent first, with at most three parallel requests per level. Single-file Direct path chains are created by `X-NC-WebDAV-Auto-Mkcol`.
   - `FileLinkTransferService` coordinates dedicated bulk, direct, and chunked uploaders. Non-bulk files up to 20 MiB use direct WebDAV `PUT` and the server-side `X-NC-WebDAV-Auto-Mkcol: 1` header. Larger files use Nextcloud chunked upload v2 under `/remote.php/dav/uploads/<user>/<upload-id>` and are assembled with `MOVE .file`. Direct and chunked files share the limit of three concurrent transfers.
   - When the typed capability snapshot exposes `dav.bulkupload = "1.0"`, at least 20 candidate files of at most 8 MiB can be packed into sequential multipart batches of at most 100 files and about 20 MiB. The planner selects bulk only when the batch plan saves at least 20 percent of all upload requests, counting base-path and share-root creation, planned directories, direct files, and every chunk-folder, chunk-`PUT`, and final `MOVE`.
   - After all transfers finish, `FileLinkShareClient` sends one OCS create-share `POST` with path, explicit permissions, password, expiration date, label, and note. It omits the legacy `publicUpload` parameter because Nextcloud would use it to replace the explicit permission mask. No metadata update request follows.
   - A missing response, a transient gateway/service response without an OCS result, or a successful response without usable share data makes the create result ambiguous. `FileLinkShareClient` records that path and performs an exact-path OCS lookup with child shares disabled before another create request. It reuses a matching public-link share, retries only after a confirmed empty result, and keeps blocking duplicate creation while the lookup result remains unknown.
   - Replay-safe `MKCOL`, direct `PUT`, chunk `PUT`, and bulk `POST` operations receive at most two retries for transport failures and selected transient HTTP responses. Each bulk retry rebuilds the same request body from the unchanged local plan. A final chunk `MOVE` is never sent twice blindly: after an indeterminate transport result, an exact depth-zero DAV probe accepts the transfer only when the target is a non-collection resource with the expected length.
   - `FileLinkUploadProgress` limits phase updates to at most ten per second. Debug logs contain plan, retry, five-second aggregate progress, and completion records instead of per-file success noise.
5. `Utilities/FileLinkHtmlBuilder.cs` generates the HTML block (header + link + password + permissions + expiration date).
   - backend-provided custom share templates are sanitized via `HtmlTemplateSanitizer` and fail closed on sanitizer errors.
   - `Models/AttachmentLinkTargetPolicy.cs` resolves `policy.share.attachment_link_target` (`zip_download` / `share_page`) against the nullable local setting. An invalid stored local value is treated as unset, so a valid editable backend value can seed it. ZIP is used when no valid local or usable backend value exists; a locked backend value wins.
   - `AttachmentMode` controls read-only permissions, rights-row suppression, and cleanup. The explicit attachment link target controls only the URL plus `{LINK_INTRO}` and `{LINK_LABEL}`. Manual shares always render the Nextcloud share page. Legacy templates without these placeholders keep their existing output.
   - ZIP URL derivation is fail-closed: the public absolute HTTP(S) URL must end in `/s/<token>` and match the OCS token. Invalid input throws before insertion; there is no original-URL fallback.
   - custom Share rendering prefers `policy.share.share_html_block_template_v2` and falls back to `policy.share.share_html_block_template`. This supports older backend releases while allowing current backends to keep the original response key placeholder-free for older clients.
   - current backends expose `policy.share.share_html_block_effective_language` for custom templates. Outlook uses it for generated link wording, field labels, permission names, and password hints; older backends without the field keep the previous UI-language fallback.
   - plain-text compose keeps `MailItem.BodyFormat=olFormatPlain`; the share block is rendered as a framed text block with `#` separators and inserted through Outlook WordEditor. Inline replies/forwards keep two empty paragraphs above the block for the sender's own text. `MailItem.Body` is not rewritten.
6. `NextcloudTalkAddIn.InsertHtmlIntoMail(...)` / `InsertPlainTextIntoMail(...)` insert the rendered block into the message body (delegated to `Controllers/MailInteropController.cs`). HTML compose uses WordEditor first so existing managed bookmarks stay intact; a direct `HTMLBody` write remains the compatibility fallback when the Inspector editor cannot be opened.

Compose runtime parity additions in `NextcloudTalkAddIn.cs` (`MailComposeSubscription`) with lifecycle logic delegated to `Controllers/ComposeShareLifecycleController`:

- The FileLink ribbon entry is exposed in mail inspectors and in the Explorer inline reply/forward `Message` tab. Both entries call the same `FileLinkLaunchController` path.
- Inline replies/forwards insert the rendered share HTML through `Explorer.ActiveInlineResponseWordEditor`; the inline path does not rewrite `MailItem.HTMLBody` and keeps two empty paragraphs above the share block for the sender's own text.
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
  - wizard finalize (enforced in `UI/FileLinkWizardForm.cs` via `Services/OutlookAttachmentAutomationGuardService.cs`).
- Attachment-mode wizard launch:
  - removes selected compose attachments
  - queues files as initial wizard selections
  - opens directly in file-step-equivalent mode.
  - copies the effective attachment link target into `FileLinkRequest`; no per-share target switch is exposed.
- `UI/FileLinkWizardForm.cs` file-step queue accepts Explorer drag & drop for files/folders across queue and action-area controls.
- Compose share cleanup lifecycle:
  - arm immediately after share creation
  - clear only after successful send
  - delete server folder artifacts on unsent close (with send/close grace timer).
- Separate password-mail dispatch:
  - queue password-only HTML after share creation
  - capture recipients on send
  - capture the sender account on send and apply it to the follow-up mail before signature/body dispatch
  - when backend policy requests Nextcloud Secrets, split the final recipient list and create one one-time Secrets link per recipient
  - Secrets links are encrypted locally with AES-GCM through Windows CNG; no new crypto dependency is bundled
  - if Secrets creation fails, fall back to the existing plain separate password mail and warn the user
  - dispatch only after successful primary send and keep the source compose mode for HTML vs plain-text follow-up mails
  - auto-send first, then manual fallback draft on failure.

#### IFB flow

1. User enables IFB in Settings.
2. `Services/FreeBusyServer.cs` starts a local HTTP listener on the configured IFB port (`Settings -> IFB -> Local IFB port`, default: `7777`).
3. `Services/FreeBusyManager.cs` updates Outlook registry values so Outlook requests free/busy data via the local endpoint.

## Network endpoints

The add-in uses Nextcloud **OCS** and **WebDAV** endpoints.

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
- Optional small-file bulk upload: `POST /remote.php/dav/bulk` (`multipart/related`, only when `ocs.data.capabilities.dav.bulkupload` is exactly `"1.0"`)
- Large file upload: `MKCOL /remote.php/dav/uploads/<user>/<upload-id>`, chunk `PUT`s, then `MOVE /remote.php/dav/uploads/<user>/<upload-id>/.file` to the final file path

Secrets (optional separate password mode):

- Create encrypted secret: `POST /ocs/v2.php/apps/secrets/api/v1/secrets`
- Public one-time link: `/index.php/apps/secrets/share/<uuid>#<local-key>`
- The key stays in the URL fragment and is not sent to Nextcloud.

IFB (DAV via proxy):

- Local listener: `http://127.0.0.1:<ifb-port>/nc-ifb/...` (default `<ifb-port>=7777`)
- The proxy talks to CalDAV and Addressbook endpoints under `remote.php/dav/...`

Update check:

- Homepage endpoint: `GET https://nc-connector.de/wp-json/ncc/v1/update-check`
- Query values: `product=outlook`, installed version, channel, and a daily rotating client hash.
- Downloads still point directly to GitHub release assets. The homepage only returns release metadata and counts one anonymous client per day.

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
- FileLink hot paths log upload plans, retries, periodic aggregate progress, and completion summaries, not every successful request.
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
5. Mail: run the sharing wizard, upload 1–2 small files, insert the HTML block, and send to yourself.
6. IFB: enable IFB, then verify the local endpoint responds:
   - `Invoke-WebRequest http://127.0.0.1:<ifb-port>/nc-ifb/ -UseBasicParsing`
7. Settings -> Advanced: click `Check now` and verify that latest version, last check, download link, and changelog summary update without blocking Outlook.

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
