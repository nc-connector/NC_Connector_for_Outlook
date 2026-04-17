# Changelog

All notable changes to **NC Connector for Outlook** will be documented in this file.

This project follows the principles of **Keep a Changelog** and **Semantic Versioning**.

## [3.0.3] - 2026-04-17

### Added
- Daily runtime log rotation with 7-file retention and 30-day cleanup.
- Log anonymization option with redaction of sensitive runtime diagnostics.
- Backend HTML template hardening with fail-closed sanitizer policy.

### Changed
- Shared helper refactor for URL launch, size formatting, and COM release cleanup.
- Refactor cleanup pass to normalize shell/COM patterns and remove leftover redundancies.
- Runtime HTTP request handling centralized via `NcHttpClient` and `NcJson`.
- Cleanup refactor: removed duplicated helpers and redundant conditions.

### Fixed
- Password-notification icon lifecycle marshaled to the captured UI context.
- Removed ineffective password-selection JavaScript handlers for Outlook HTML rendering.

## [3.0.2] - 2026-04-15

### Added
- New runtime TLS controls in Outlook settings (`Use OS default TLS policy`, `Enable TLS 1.2`, `Enable TLS 1.3`) with immediate test/login-flow usage and persisted profile-specific configuration.

### Changed
- Release line/version references were aligned to `3.0.2` across assembly and installer defaults.
- TLS hint text in advanced settings/docs was tightened and no longer makes assumptions about machine-wide registry changes.
- TLS diagnostics guidance text was updated to be mode-neutral (no hardcoded OS-default assumption) and aligned across all supported locale files.

### Fixed
- Admin-controlled `?` hint glyph placement in advanced settings now avoids overlapping adjacent inputs in tight rows (language override section and neighboring fields).
- Connection diagnostics and login-flow connectivity checks now force fresh HTTP/TLS handshakes (no pooled keep-alive reuse), so runtime TLS mode switches are validated deterministically without requiring an Outlook restart.

## [3.0.1] - 2026-04-14

### Added
- The FileLink wizard now accepts Explorer drag & drop for files/folders across the entire file step (queue, surrounding pane, and action area), not just via explicit add buttons.
- TLS defaults are now hardened at add-in scope via `NcTalkOutlookAddIn.dll.config`, explicitly enabling strong crypto and OS-default TLS negotiation without changing machine-wide .NET registry settings.
- Settings connection diagnostics now classify transport failures (TLS handshake, certificate trust, DNS, proxy/connectivity, timeout) and surface actionable guidance instead of only generic secure-channel errors.

### Changed
- Release line/version references were aligned to `3.0.1` across assembly metadata, installer defaults, readmes, and admin/development docs.
- The new connection diagnostics strings were translated across all supported Outlook locales, so non-English users do not fall back to generic English placeholders.
- Compose attachment automation now evaluates attachments pre-add via `BeforeAttachmentAdd`; threshold/always flows can cancel the host add and open NC share mode before Outlook post-add handling.
- Compose automation subscription coverage was expanded with `ApplicationEvents_11.ItemLoad` so mail items loaded outside `NewInspector` (for example inline compose contexts) are also tracked for attachment automation.
- Repository text-format policy was standardized with new `.editorconfig` and `.gitattributes`; project text files were normalized to consistent CRLF line endings.
- With debug logging enabled, compose pre-add attachment automation now logs detailed candidate/decision/fallback diagnostics (including unresolved-path reasons) to improve support traceability.
- `NextcloudTalkAddIn` was split into dedicated `partial` units for runtime hooks and logging to reduce orchestration density in the main source file.
- COM cleanup now uses shared `ComInteropScope` helpers (including scoped wrappers), and high-risk recipient/explorer release paths were migrated to the centralized implementation.
- Oversized UI forms were modularized with `partial` feature files (`FileLinkWizardForm.DragDrop`, `SettingsForm.Language`, `TalkLinkForm.Moderator`) to improve maintainability and test focus.
- Talk body/template rendering and block sanitation were extracted into `Controllers/TalkDescriptionTemplateController`.
- Appointment attendee e-mail discovery and SMTP recipient resolution were extracted into `Controllers/OutlookRecipientResolverController`.
- Compose subscription lifecycle (get-or-create/remove/dispose-all) now runs through `Controllers/MailComposeSubscriptionRegistryController` instead of direct list/lock management in the add-in root.
- Talk appointment lifecycle logic (`ApplyRoomToAppointment`, runtime room trait resolution, room mutation sync, delegation and participant sync) was extracted into `Controllers/TalkAppointmentController`.
- Compose share cleanup and separate-password dispatch flow (including recipient normalization, auto-send, and manual fallback) was extracted into `Controllers/ComposeShareLifecycleController`.
- Legacy recipient helper forwarders in `NextcloudTalkAddIn` were removed; compose recipient normalization now calls `ComposeShareLifecycleController` helpers directly.
- Remaining recipient CSV/normalization passthrough wrappers were removed from the add-in root to avoid duplicate helper paths.
- `ComposeShareLifecycleController` COM release paths now use centralized `ComInteropScope.TryRelease(...)` for consistent exception-safe cleanup.
- `build.ps1` now supports `-SkipIceValidation` for environments where WiX ICE execution is unavailable (`WIX0217`), while keeping the default validated build path unchanged.
- Large nested runtime subscription classes were moved out of the root file into dedicated partial units:
  - `NextcloudTalkAddIn.MailComposeSubscription.cs`
  - `NextcloudTalkAddIn.AppointmentSubscription.cs`
- Minor UI redundancy cleanup (no behavior change) was applied in `FileLinkWizardForm`, `SettingsForm`, and `TalkLinkForm` (shared helpers for selection validation, resize wiring, and settings-option checkbox setup).

### Fixed
- Pre-add multi-file drag/drop is now debounced and batched into a single wizard launch instead of opening one wizard per file.
- Folder uploads in the FileLink wizard now correctly create required subfolders for mixed file+directory queues; reserved-name tracking no longer suppresses required DAV `MKCOL` calls.
- Upload status/progress UI now flushes buffered per-item progress updates immediately on failure/cancel/finalize paths, preventing stale bars or missing final state labels.
- Admin-controlled `?` hint glyphs now support explicit row anchors, preventing them from drifting into adjacent password and attachment threshold input fields in the FileLink wizard, sharing settings, and Talk password block.
- Pre-add attachment interception now also probes `Attachment.FileName` and `Attachment.DisplayName` for resolvable local paths when `Attachment.PathName` is unavailable, improving early capture reliability for drag/drop scenarios.
- Pre-add attachment candidate materialization now falls back to `Attachment.SaveAsFile(...)` and compares COM-reported size vs. measured file size, reducing false below-threshold decisions for unresolved path scenarios.
- `ApplicationEvents_11.ItemLoad` compose subscription is now limited to active inline-compose contexts to prevent duplicate compose subscriptions and duplicate threshold prompts when inspector-based subscriptions are already active.
- Add-in lifecycle teardown was de-duplicated by centralizing shutdown/disconnect cleanup into a shared idempotent path.

## [3.0.0] - 2026-03-30

### Added
- Optional NC Connector backend policy runtime for Talk and Sharing:
  - backend status endpoint is queried on Talk wizard open, Sharing wizard open, and Settings open/save
  - active valid seats enable backend policy values and `policy_editable` locks
  - paused/invalid seat states show an in-product warning and fall back to local add-in settings
  - central templates are supported for share HTML/password blocks and Talk description text
  - separate password follow-up delivery is explicitly gated behind backend endpoint + active assigned seat
  - backend custom text templates are only activated when the language override is set to `Custom`, otherwise local UI-default text remains active
  - backend attachment-threshold policy now treats `attachments_min_size_mb: null` as an explicit "disabled" state

### Changed
- Release line/version references were aligned to `3.0.0` across assembly metadata, installer defaults, readmes, and admin/development docs.
- Backend policy runtime is now fetched live on the relevant entry points and evaluated by compose attachment automation as well as in Settings/Wizard UI, so locked backend attachment rules stay authoritative without reusing stale cached policy data.
- Backend policy runtime now targets `/apps/ncc_backend_4mc/api/v1/status`; if the backend is unreachable, the license/seat state is no longer usable, or the backend grace window has expired, Outlook falls back to local add-in settings.
- Talk event descriptions now honor backend `event_description_type`; HTML templates are written into the open Outlook appointment editor with stable NC block markers while `Body` stays aligned for room-description sync.
- Locked settings and wizard controls now expose their admin/seat/backend explanation through active hint anchors instead of relying on WinForms tooltips on disabled controls.

### Fixed
- Outlook issue #4 (`HTTP 400` at final share creation) was hardened in the sharing service:
  - create-share now extracts and surfaces OCS server error details instead of only generic WebException text
  - share creation now follows the documented Nextcloud OCS contract more closely: `label` is sent on create, mutable metadata is updated separately through form-encoded OCS update arguments
  - the previous `HTTP 400` retry path that silently dropped optional `label` / `note` metadata was removed
- Runtime exceptions are now always written to `addin-runtime.log`, even when the optional debug log switch is off.
- Backend custom share HTML in attachment mode now removes only the `RIGHTS` row, preventing truncated or fragmentary HTML blocks.
- Compose mail insertion now prefers the `HTMLBody` path for backend HTML blocks, avoiding partial insertion from the old Word-editor fallback path.
- Talk appointment save handling now performs a short deferred post-write lobby verification so start-time changes from externally created appointments are still picked up after Outlook commits the final value.

## [2.3.0] - 2026-02-28

### Added
- Compose attachment automation settings:
  - `Always handle attachments via NC Connector`
  - `Offer upload above X MB`
- Two-action threshold prompt in compose:
  - `Share with NC Connector`
  - `Remove last selected attachments`
- Attachment-mode wizard launch context (direct file-step start, fixed `email_attachment` share naming with deterministic suffixes, `yyyyMMdd` date prefix).
- Separate password mail flow:
  - new default setting `Send password separately`
  - wizard toggle `Send password in separate email`
  - password-only follow-up HTML and post-send dispatch queue.
- Compose share cleanup lifecycle for unsent drafts (armed/cleared/delayed/delete runtime states).
- Built-in host large-attachment conflict guard with live setting checks in Settings UI and runtime flow gates.
- System-addressbook hardening for Talk defaults and Talk wizard:
  - centralized runtime availability contract (`available`, `error`, `count`, `forceRefresh`)
  - live checks on Talk click, settings open/save, and wizard open
  - deterministic lock state with context-specific tooltips and red warning blocks (Settings + Wizard).

### Changed
- Sharing HTML output in attachment mode:
  - ZIP download URL now uses `/s/<token>/download`
  - permissions row is hidden in attachment mode
  - inline password is hidden when separate password dispatch is enabled.
- Attachment-mode permissions are enforced as read-only regardless of sharing defaults.
- Logging depth for compose sharing flows now includes attachment-evaluation decisions, prompt actions, cleanup lifecycle transitions, and separate-password dispatch outcomes.
- Settings persistence is profile-aware and XML-based:
  - `%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml`
  - encrypted `AppPasswordProtected` (Windows DPAPI, CurrentUser scope)
  - automatic migration from legacy `settings.ini` paths.
- Runtime data path has been consolidated to `%LOCALAPPDATA%\NC4OL`:
  - debug log (`addin-runtime.log`)
  - IFB address book cache.
- Installer defaults have been renamed for consistent product naming:
  - install directory: `C:\Program Files\NC4OL`
  - IFB marker key: `HKLM\Software\NC4OL\HttpUrl`
  - legacy installer key `HKLM\Software\NextcloudTalkOutlookAddIn\HttpUrl` is removed on install.
- Delegation write-flow in appointment `OnWrite` was tightened:
  - pre-step execution order is deterministic (`room name` -> `lobby` -> `description` -> `participants`)
  - per-step pre-sync status is logged explicitly before moderator handover
  - pending delegation is no longer silently blocked by pre-step failures (handover still executes, failures stay visible in runtime logs).

### Fixed
- Debounced attachment evaluation to prevent duplicate triggers during rapid multi-selection.
- Runtime guard enforcement now blocks compose attachment automation not only in UI, but also before evaluation, before threshold action handling, and before attachment-mode finalize.
- Password policy capability parsing is now compatible with current Nextcloud payload variants (`minLength` etc.) and normalized generator URL formats.
- Password policy HTTP handling now enables gzip/deflate decompression and sanitizes JSON wrappers before parsing.
- Cross-client Talk edit interoperability was hardened for Thunderbird-created appointments:
  - missing local room traits are now bootstrapped from Talk server endpoints (`/object`, `/webinar/lobby`)
  - resolved room traits are persisted back into Outlook appointment properties
  - prevents false room-description update failures like `...: event` when editing TB-created event conversations.
- MSI maintenance behavior was aligned for support scenarios:
  - reinstall/update now supports same/older/newer package installation over an existing installation (`AllowDowngrades="yes"`).
- Room creation path no longer falls back from event conversation to standard conversation on create errors:
  - requested room type is now kept deterministic
  - if event conversation prerequisites are missing, creation fails fast with a clear service error.

### Documentation
- Updated `README.md` / `README.de.md` with 2.3.0 operational behavior:
  - profile-based settings XML + legacy migration cleanup
  - consolidated runtime path `%LOCALAPPDATA%\NC4OL`
  - compose cleanup and separate-password follow-up flow semantics.
- Clarified MSI reinstall/update semantics in README/admin docs (same/older/newer package install over existing installation).
- Expanded `docs/ADMIN.md` / `docs/ADMIN.de.md` with compose-sharing lifecycle details (cleanup contract and password follow-up dispatch behavior).
- Expanded `docs/DEVELOPMENT.md` / `docs/DEVELOPMENT.de.md` with explicit 2.3.0 implementation deltas and runtime contracts used for parity work.

## [2.2.7] - 2026-02-13

### Added
- More UI translations (see `Translations.md`).
- Tooltips across Settings and the wizards.
- Optional auto-add of event invitees to the Talk room (Nextcloud users via system address book, others via e-mail).
- Live Nextcloud password policy support for Talk and Sharing (minimum length + generator API, with secure fallback).

### Changed
- Unified, modernized UI across Talk wizard, Sharing wizard, and Settings.
- Legal/branding update: renamed to **NC Connector for Outlook**, with new app icon and header assets.

### Removed
- Buggy muzzle feature.

### Fixed
- Dark mode/theme handling to better follow the Outlook/Office theme.

### Documentation
- Expanded admin and developer documentation.


