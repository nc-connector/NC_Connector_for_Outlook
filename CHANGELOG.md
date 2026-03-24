# Changelog

All notable changes to **NC Connector for Outlook** will be documented in this file.

This project follows the principles of **Keep a Changelog** and **Semantic Versioning**.

## [2.2.9] - 2026-03-24

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
- Corrected all 2.2.9 branch version references in README/admin/development docs:
  - MSI examples now use `NCConnectorForOutlook-2.2.9.msi`
  - local path examples now use `nc4ol-2.2.9`.
- Clarified MSI upgrade/reinstall semantics in README/admin docs (same/older/newer package install over existing installation).

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
