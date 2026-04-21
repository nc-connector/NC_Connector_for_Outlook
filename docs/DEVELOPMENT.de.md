# Entwicklungsleitfaden — NC Connector for Outlook

Dieses Dokument ist ein einsteigerfreundlicher Leitfaden zum Bauen, Debuggen und Erweitern von **NC Connector for Outlook** (Outlook classic COM Add-in).

## Inhalt

- [Projektzweck](#projektzweck)
- [Schnellstart](#schnellstart)
- [Repository-Struktur](#repository-struktur)
- [Architektur](#architektur)
- [Netzwerk-Endpunkte](#netzwerk-endpunkte)
- [Lokalisierung (i18n)](#lokalisierung-i18n)
- [Logging](#logging)
- [Kompatibilitaet und Versionspruefungen](#kompatibilitaet-und-versionspruefungen)
- [Build und Release](#build-und-release)
- [Lokales Testen](#lokales-testen)
- [X-NCTALK-* Property-Referenz](#x-nctalk--property-referenz)
- [Erweiterungspunkte](#erweiterungspunkte)

## Projektzweck

Das Add-in verbindet Outlook classic mit einem Nextcloud-Server und bietet:

- **Nextcloud Talk** aus Kalendereintraegen (Room-Erstellung, Lobby, Teilnehmer-Sync, Moderator-Delegation)
- **Nextcloud Sharing** aus dem Mail-Compose-Fenster (Upload + Link-Share + HTML-Block-Insertion)
- **Internet Free/Busy (IFB)** ueber einen lokalen HTTP-Endpunkt, der Requests an Nextcloud weiterleitet

## Release-Delta 3.0.3

Dieses Release hat verhaltenskritische Paritaetsregeln eingefuehrt, die bei zukuenftigen Aenderungen erhalten bleiben muessen:

- Compose-Attachment-Automation hat jetzt eine deterministische Zwei-Modi-Verarbeitung (`always` vs. Threshold-Prompt).
- Compose-Attachment-Automation bewertet zusaetzlich pre-add (`BeforeAttachmentAdd`) und kann Host-Add best effort vor Outlook post-add abbrechen, wenn ein lokaler Pfad aufloesbar ist.
- Der Compose-Threshold-Prompt ist eine strikt zweiaktionale Entscheidung (`Share with NC Connector` / `Remove last selected attachments`) mit Batch-Remove-Semantik.
- Der Attachment-Mode-Share-Output ist spezialisiert (`email_attachment`-Namensvertrag, read-only Empfaenger, ZIP-Download-Link `/s/<token>/download`, keine Permissions-Zeile).
- Compose-Share-Cleanup ist lifecycle-getrieben (armed nach Share-Erstellung, cleared nur nach bestaetigtem Send, delayed delete bei unsent-close race).
- Separater Passwort-Follow-up-Dispatch erfolgt nur post-send und umfasst automatisches Senden plus manuellen Fallback-Entwurf.
- Talk-Defaults und Talk-Wizard nutzen einen zentralen System-Addressbook-Availability-Vertrag und deterministischen Lock-Status.
- Optionaler NC Connector Backend-Policy-Mode wird an Wizard-/Settings-Entry-Points ausgewertet und kann Talk + Sharing Defaults inkl. zentraler Text-/Template-Payloads ueberschreiben/locken.
- Backend-Custom-Templates werden nur aktiviert, wenn der effektive Sprach-Override `custom` ist; sonst bleibt die Runtime bei lokalen UI-Default-Textbloecken.
- Die Option `custom` wird nur angezeigt, wenn der Backend-Endpunkt existiert, und bleibt deaktiviert, solange die effektive Backend-Policy fuer die jeweilige Domain nicht wirklich `custom` ist und ein Template liefert.
- Separater Passwort-Follow-up-Dispatch ist seat-gated und nur mit Backend-Endpunkt + aktiv zugewiesenem Seat verfuegbar.
- Die Backend-Attachment-Threshold-Policy nutzt `attachments_min_size_mb` gleichzeitig als Wert und Enable-State: positive Integer aktivieren Threshold-Mode, `null` deaktiviert ihn.
- Gelocte Backend-Attachment-Automation-Policy wird auch in der Compose-Runtime ueber den Live-Backend-Status erzwungen, nicht nur in der Settings-UI.
- Wenn das Backend nicht erreichbar ist, faellt die Runtime sofort auf lokale Add-in-Settings zurueck.
- Wenn das Backend erreichbar ist, aber License/Seat-State nicht mehr nutzbar ist oder das Backend-Grace-Window abgelaufen ist, faellt die Runtime ebenfalls auf lokale Add-in-Settings zurueck.
- Talk-Event-Description-Templates koennen als `html` oder `plain_text` kommen; bei aktivem `html` schreibt Outlook den Talk-Block mit stabilen Markern in den offenen Appointment-Editor, sodass die gerenderte Event-Description HTML bleibt, waehrend `Body` weiterhin die synchronisierte Plain-Text-Ansicht liefert.
- Share-Erstellung folgt dem dokumentierten Nextcloud-OCS-Vertrag jetzt enger: Create via `POST /shares` mit `label`, danach Update veraenderbarer Metadaten wie `note` via form-encoded `PUT /shares/{id}`-Argumente.
- Runtime-Settings und Caches nutzen `%LOCALAPPDATA%\NC4OL` mit profile-spezifischer XML-Settings-Migration.
- Transport-Security-Settings sind runtime-live in `SettingsForm` (`OS default` / `TLS 1.2` / `TLS 1.3` / kombiniert) und werden beim Apply bewusst validiert.
- Nicht unterstuetztes manuelles `TLS 1.3`-Runtime-Apply faellt jetzt fail-closed (kein implizites Downgrade auf TLS 1.2).
- Settings-Connectivity-Operationen (Connection Test, Login Flow) erzwingen frische HTTP/TLS-Handshakes (kein gepooltes Keep-Alive-Reuse), sodass TLS-Mode-Toggles deterministisch getestet werden.
- TLS-Transport-Diagnostics-Text ist mode-neutral und geht nicht mehr von OS-default aus.

## Schnellstart

### Voraussetzungen

- Windows 10/11
- Outlook classic (x64 oder x86)
- **.NET Framework 4.7.2** (Target Framework)
- MSBuild (z. B. Visual Studio Build Tools)
- **.NET SDK** (wird vom WiX-v4-Build ueber `dotnet` verwendet)

### MSI bauen (empfohlen)

```powershell
cd "C:\path\to\nc4ol-3.0.4"

# Optional: Reference Assemblies (nur falls noetig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages
$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"

.\build.ps1 -Configuration Release
```

Wenn WiX-ICE-Validierung auf dem Build-Host nicht verfuegbar ist (z. B. `WIX0217` in restriktiven Umgebungen), nutze:

```powershell
.\build.ps1 -Configuration Release -SkipIceValidation
```

Output:

- `dist\NCConnectorForOutlook-<version>.msi`

### Lokal installieren und starten

1. MSI installieren (Administratorrechte erforderlich):
   - `msiexec /i dist\NCConnectorForOutlook-<version>.msi`
2. Outlook starten
3. Ribbon:
   - Kalender/Termin: **NC Connector → Talk-Link einfuegen**
   - Mail Compose: **NC Connector → Nextcloud-Freigabe einfuegen**
4. **NC Connector → Settings** oeffnen und Server-URL + Credentials konfigurieren.

## Repository-Struktur

Top-Level:

- `src/` — COM Add-in (WinForms UI + Service Layer)
- `installer/` — WiX-v4-MSI-Projekt (Files + Registry + URLACL)
- `docs/` — Admin-/Entwicklungsdokumentation
- `VENDOR.md` — Hinweise/Lizenzen fuer gebuendelte Third-Party-Sanitizer-/Runtime-Dependencies
- `assets/` — Branding-Bilder fuer README/Screenshots
- `dist/` — Build-Output (MSI)

Wichtige Code-Pfade:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` — Entry-Point, Ribbon XML, Outlook Event Wiring, Orchestrierung
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs` — Add-in Bootstrap/Teardown Lifecycle (`OnConnection`, shutdown/disconnect)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs` — dedizierte Outlook Event Hook/Unhook Wiring-Helper
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs` — kategorie-spezifische Runtime-Logging-Helper
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs` — Backend-Policy + Talk-Template/Language-Resolver-Helper
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs` — deferred appointment-subscription ensure und Outlook event-restriction handling
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs` — compose-subscription Core-Status + Lifecycle-Entry-Points (`Dispose`, identity, shared helpers)
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs` — compose attachment interception/evaluation/share-launch flow
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs` — send/close cleanup lifecycle + separate-password dispatch handling
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs` — appointment runtime subscription lifecycle
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs` — settings open/save/revert orchestration
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs` — FileLink ribbon launch + wizard orchestration
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs` — Talk ribbon flow orchestration (auth gate, wizard, room create/replace)
- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` — appointment lifecycle orchestration fuer Talk room metadata/sync
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` — compose share cleanup + separate-password dispatch flow
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` — Talk template/body block rendering
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` — SMTP- und attendee-recipient-resolution
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` — compose-subscription registry lifecycle
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` — gemeinsame mail/inspector interop helpers
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` — gemeinsames Modell fuer queue entries des separaten Passwort-Follow-up-Dispatch
- `src/NcTalkOutlookAddIn/Services/` — Nextcloud HTTP Integrationen (Talk, Sharing, IFB, Login Flow)
  - `Services/NcHttpClient.cs` ist der zentrale Request-Executor fuer Auth-Header, OCS-Header, Timeout/Decompression und optionalen Fresh-Connection-Mode.
  - Alle Runtime-HTTP-Calls (Talk, share/DAV, IFB, login flow, moderator avatar fetch) laufen ueber `NcHttpClient`.
- `src/NcTalkOutlookAddIn/UI/` — WinForms-Dialoge und Wizard
- `src/NcTalkOutlookAddIn/Settings/` — persistiertes Settings-Modell + Storage
- `src/NcTalkOutlookAddIn/Utilities/` — Logging, Theming, i18n, kleine gemeinsame Helper
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` — zentraler Sanitizer fuer backend-gelieferte Share/Talk-HTML-Templates
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` — zentrale JSON-Payload-Normalisierung (`PrepareJsonPayload`), dictionary/string/int helper und OCS-error extraction
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` — gekapselter pending-key tracking + throttled logging state fuer deferred appointment ensure
- `src/NcTalkOutlookAddIn/Utilities/PictureConverter.cs` — gemeinsamer Image -> IPictureDisp-Conversion-Helper fuer Ribbon-Icons

## Architektur

### Hauptbausteine

- **COM Add-in Lifecycle**
  - `NextcloudTalkAddIn.OnConnection(...)` laedt Settings, aktiviert Logging (optional), initialisiert IFB und verdrahtet Outlook-Events.
- **Workflow-Controller**
  - `SettingsWorkflowController`, `FileLinkLaunchController` und `TalkRibbonController` kapseln ribbon-getriggerte UI-/Runtime-Workflows.
  - `NextcloudTalkAddIn.cs` bleibt COM/Ribbon/Event-Composition-Root und delegiert Feature-Flows an Controller.
  - `TalkRibbonController` und `FileLinkLaunchController` prefetchen Backend-Policy + Passwort-Policy parallel vor Wizard-Open, holen aber trotzdem bei jedem Entry frische Runtime-Policy-Daten.
  - Lifecycle-, Policy/Template-Resolution- und deferred subscription ensure-Logik sind in dedizierte Partial-Dateien aufgeteilt, damit die Root-Orchestrierungsklasse wartbar bleibt.
- **Service Layer**
  - `Services/TalkService.cs` spricht die Talk-OCS-API an.
  - `Services/FileLinkService.cs` uploadet via WebDAV und erstellt Shares via OCS.
  - `Services/FreeBusyServer.cs` hostet den lokalen IFB-HTTP-Endpunkt.
  - `Services/FreeBusyManager.cs` schreibt Outlook-Registry-Werte so, dass Outlook den lokalen IFB-Endpunkt nutzt.
- **UI**
  - `UI/SettingsForm.cs` konfiguriert Base URL, Authentifizierung, Sharing-Defaults, IFB und Debug-Logging.
  - `UI/TalkLinkForm.cs` ist der Talk-Wizard.
  - `UI/FileLinkWizardForm.cs` ist der Sharing-Wizard.
  - `UI/BrandedHeader.cs` ist das gemeinsame Header-Banner-Control.
- **Gemeinsame Utilities**
  - `Utilities/BrowserLauncher.cs` zentralisiert Shell-Starts (URLs, Dateien, Verzeichnisse).
  - `Utilities/SizeFormatting.cs` zentralisiert MB-Formatierung.
  - `Utilities/ComInteropScope.cs` zentralisiert COM release/final-release patterns.
  - `Utilities/HtmlTemplateSanitizer.cs` erzwingt eine Thunderbird-alignte HTML-Policy fuer Backend-Templates und faellt fail-closed.

### End-to-end Flows

#### Talk-Link-Flow (Termine)

1. User klickt **Talk-Link einfuegen** in einem Termin.
2. `UI/TalkLinkForm.cs` sammelt: Titel, Passwort, Lobby, Listable-Flag, Room-Type, Teilnehmer-Sync-Optionen, optionales Delegationsziel.
3. `Controllers/TalkRibbonController.cs` prefetcht Backend-Policy-Status und Passwort-Policy parallel (`Task.WhenAll`) vor dem Oeffnen des Wizards.
4. `Services/TalkService.cs` erstellt den Room via OCS.
5. `Controllers/TalkAppointmentController.ApplyRoomToAppointment(...)` (aufgerufen von `NextcloudTalkAddIn`) aktualisiert den Termin:
   - `Location` (Talk URL)
   - lokalisierter Plain-Text-Body-Block (inkl. Passwort und Help-URL)
   - persistierte Metadaten als Outlook `UserProperties` (inkl. `X-NCTALK-*` Keys)
   - backend-gelieferte Custom-Talk-Templates werden vor Rendering sanitisiert (kein Raw-HTML-Fallback)
   - Talk-Appointment-HTML laeuft vor Insert durch einen expliziten Kompatibilitaets-Transform (`HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(...)`)
   - Appointment-HTML-Insert nutzt die HTML->RTF-Bridge (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`), nicht `AppointmentItem.HTMLBody` und nicht `HTMLEditor.body.innerHTML`
6. Eine Runtime-Subscription wird fuer den Termin registriert (`AppointmentSubscription` in `NextcloudTalkAddIn.AppointmentSubscription.cs`):
   - **Write** (save): aktualisiert Lobby-Timer bei Zeitveraenderungen, aktualisiert Room-Description, synct Teilnehmer, wendet Delegation an
   - wenn Outlook die finale geaenderte Startzeit erst kurz nach `Write` liefert, prueft eine kurze deferred post-write verification die Lobby-Aktualisierung erneut auf demselben offenen Termin statt breit im Kalender zu scannen.
   - **Close** (verwerfen ohne Speichern): loescht den Room, um Orphans zu vermeiden (best effort)
   - **BeforeDelete**: loescht den Room (best effort)

#### Appointment-sicheres HTML-Subset (Backend-Custom-Templates)

Fuer stabiles Rendering in Outlook-Termin-Bodies (Word/RTF-Pipeline) sollten Backend-Talk-Templates in diesem Subset bleiben:

- Table-first Layout (`table`, `tbody`, `tr`, `td`) fuer Struktur.
- Inline-Styles sind erlaubt, aber NC Connector entfernt bekannte Word-unzuverlaessige Deklarationen beim Appointment-Rendering:
  - `display:flex|grid`, `flex*`, `grid*`, `border-radius*`, `overflow*`, `object-fit`, `user-select` (inkl. vendor-prefixed Varianten).
- Color/Alignment-Fallbacks werden beim Appointment-Kompatibilitaets-Transform automatisch injiziert:
  - `style=color` -> `<font color=...>`
  - `style=background-color` -> `bgcolor`
  - `style=text-align` -> `align`
  - `style=vertical-align` -> `valign`
- Anchor-Color-Hardening: Link-Farbe wird zusaetzlich als `<a><font color=...>...</font></a>` verankert, wenn erforderlich.
- Nicht unterstuetzte/unsichere Tags/Attribute werden weiterhin vom Sanitizer entfernt (fail-closed).

#### Sharing-Flow (Mail Compose)

1. User klickt **Nextcloud-Freigabe einfuegen** waehrend Mail Compose.
2. `UI/FileLinkWizardForm.cs` sammelt Sharing-Settings und Datei/Ordner-Auswahl.
3. `Controllers/FileLinkLaunchController.cs` prefetcht Backend-Policy-Status und Passwort-Policy parallel (`Task.WhenAll`) vor dem Oeffnen des Wizards.
4. `Services/FileLinkService.cs` fuehrt WebDAV-Upload aus, erstellt den Public-Share via OCS (`label` bei Create) und aktualisiert danach veraenderliche Metadaten wie `note` ueber die dokumentierten OCS-Update-Argumente.
5. `Utilities/FileLinkHtmlBuilder.cs` erzeugt den HTML-Block (Header + Link + Passwort + Permissions + Ablaufdatum).
   - backend-gelieferte Custom-Share-Templates werden via `HtmlTemplateSanitizer` sanitisiert und fail-closed behandelt.
6. `NextcloudTalkAddIn.InsertHtmlIntoMail(...)` fuegt das HTML in den Message-Body ein (delegiert an `Controllers/MailInteropController.cs`).

Compose-Runtime-Paritaets-Erweiterungen in `NextcloudTalkAddIn.cs` (`MailComposeSubscription`) mit delegierter Lifecycle-Logik in `Controllers/ComposeShareLifecycleController`:

- Debounced Attachment-Evaluation (`ComposeAttachmentEvalDebounceMs`) nach Compose-Attachment-Aenderungen.
- Attachment-Automation-Modi:
  - Attachments immer in NC-Sharing-Flow routen, oder
  - Threshold-Mode mit Zwei-Aktionen-Prompt (`Share with NC Connector` / `Remove last selected attachments`).
- Pre-add Attachment-Interception:
  - `BeforeAttachmentAdd`-Pfad loest Candidate-Dateimetadaten frueh auf
  - kann Host-Attachment-Add best effort abbrechen und NC-Sharing vor Outlook post-add starten.
  - harte Outlook/Exchange-Groessenlimits koennen trotzdem vor Add-in-Callbacks greifen und sind ueber offizielle Outlook-OOM-Events nicht abfangbar.
- Runtime-Host-Guard-Checks (live large-attachment setting) an:
  - pre-evaluation
  - pre-prompt-action handling
  - wizard finalize (erzwungen in `UI/FileLinkWizardForm.cs` via `Services/OutlookAttachmentAutomationGuardService.cs`).
- Attachment-Mode-Wizard-Launch:
  - entfernt ausgewaehlte Compose-Attachments
  - queued Dateien als initiale Wizard-Selektionen
  - oeffnet direkt im file-step-aequivalenten Modus.
- `UI/FileLinkWizardForm.cs` file-step queue akzeptiert Explorer Drag & Drop fuer Dateien/Ordner ueber Queue und Action-Area-Controls.
- Compose-Share-Cleanup-Lifecycle:
  - arm sofort nach Share-Erstellung
  - clear nur nach erfolgreichem Send
  - delete Server-Folder-Artefakte bei unsent close (mit send/close grace timer).
- Separate Passwort-Mail-Dispatch:
  - queue password-only HTML nach Share-Erstellung
  - capture Empfaenger beim Send
  - dispatch nur nach erfolgreichem Primary-Send
  - zuerst auto-send, danach manueller Fallback-Entwurf bei Fehler.

#### IFB-Flow

1. User aktiviert IFB in den Settings.
2. `Services/FreeBusyServer.cs` startet einen lokalen HTTP-Listener auf dem konfigurierten IFB-Port (`Settings -> IFB -> Local IFB port`, Default: `7777`).
3. `Services/FreeBusyManager.cs` aktualisiert Outlook-Registry-Werte so, dass Outlook free/busy ueber den lokalen Endpunkt anfragt.

## Netzwerk-Endpunkte

Das Add-in nutzt Nextcloud-**OCS**- und **WebDAV**-Endpunkte.

Talk (OCS, Auswahl):

- Capabilities/Version-Hinweis: `GET /ocs/v2.php/cloud/capabilities`
- Room erstellen: `POST /ocs/v2.php/apps/spreed/api/v4/room`
- Room loeschen: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>`
- Lobby-Timer: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/webinar/lobby`
- Listable-Scope: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/listable`
- Description: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/description`
- Teilnehmer hinzufuegen: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants`
- Teilnehmer lesen: `GET /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants?includeStatus=true`
- Moderator promoten: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/moderators`
- Self leave: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants/self`

Sharing:

- Public Share erstellen: `POST /ocs/v2.php/apps/files_sharing/api/v1/shares`
- Upload/Ordner-Erstellung: `remote.php/dav/...` (WebDAV)

IFB (DAV via Proxy):

- Lokaler Listener: `http://127.0.0.1:<ifb-port>/nc-ifb/...` (Default `<ifb-port>=7777`)
- Der Proxy spricht mit CalDAV- und Addressbook-Endpunkten unter `remote.php/dav/...`

## Lokalisierung (i18n)

- Locale-Dateien:
  - `src/NcTalkOutlookAddIn/Resources/_locales/<lang>/messages.json`
- Runtime-Loader:
  - `src/NcTalkOutlookAddIn/Utilities/Strings.cs`

Hinweise:

- Default-Sprache ist **Deutsch** (`de`).
- Die UI-Sprache wird aus der Windows-UI-Culture abgeleitet. Einige generierte Textbloecke koennen in den Settings ueberschrieben werden (siehe "Language overrides").
- Platzhalter in `messages.json` verwenden `$1`, `$2`, ... und werden zu `.NET`-`string.Format`-Platzhaltern konvertiert.

Siehe `Translations.md` fuer vollstaendige Sprachliste und Pflege-Workflow.

## Logging

Debug-Logging ist optional und soll Support-Faelle reproduzierbar machen.

- Aktivieren: Settings -> **Debug** -> "Write debug log file"
- Optionale Sicherheitskontrolle (Default an): "Anonymize logs"
- Taegliches Log-Dateiformat: `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`
- Runtime-Exceptions werden immer via `DiagnosticsLogger.LogException(...)` geschrieben, auch wenn Debug-Logging deaktiviert ist.
- Retention: letzte 7 taegliche Logs behalten und Dateien aelter als 30 Tage loeschen (best effort cleanup).
- Anonymisierung maskiert konfigurierte NC URL/Base-Host, token/password-aehnliche Werte, Authorization-Credentials, User-Identifier, E-Mail-Adressen und lokale User-Pfadfragmente vor dem Log-Write.

Format:

- `[YYYY-MM-DD HH:mm:ss.fff] [CATEGORY] Message`

Beispiel:

```
[2026-02-13 03:57:12.345] [TALK] BEGIN CreateRoom
[2026-02-13 03:57:12.910] [TALK] END CreateRoom (565 ms)
```

Implementierung:

- `src/NcTalkOutlookAddIn/Utilities/DiagnosticsLogger.cs`
- `src/NcTalkOutlookAddIn/Utilities/LogCategories.cs`

Leitlinien fuer neuen Code:

- Logge **Start/Ende** von Netzwerktoperationen (nutze `DiagnosticsLogger.BeginOperation(...)`).
- Logge **Entscheidungen** (Feature Detection, Versionspruefungen, Fallbacks).
- Logge **Exceptions mit Kontext** (nutze `DiagnosticsLogger.LogException(...)`).
  `LogException(...)` umgeht den optionalen Debug-Schalter und muss der immer-aktive Error-Pfad bleiben.
- Exceptions niemals stillschweigend schlucken.

## Kompatibilitaet und Versionspruefungen

### Outlook-Bitness (x86 auf x64 Windows)

Outlook kann als 32-bit Anwendung auf 64-bit Windows installiert sein. In diesem Fall liest Outlook COM-Add-in-Registrierung aus der 32-bit Registry-View (`Wow6432Node`).

Die MSI registriert Add-in-Keys fuer **beide** Registry-Views:

- 64-bit: `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`
- 32-bit: `HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

Installer-Definition:

- `installer/Product.wxs`

### Nextcloud Feature Detection

Einige Features haengen von Server-Capabilities ab.

- Event-Conversations erfordern Nextcloud **>= 31**.
- Versionspruefungen und gecachter "server version hint" liegen in:
  - `src/NcTalkOutlookAddIn/Utilities/NextcloudVersionHelper.cs`

### UI-Theming (WinForms)

Das Add-in nutzt, wo passend, ein Dark Theme, damit Dialoge zu dunklen Outlook-Setups passen.

Implementierung:

- `src/NcTalkOutlookAddIn/Utilities/UiThemeManager.cs`

Detection-Logik (best effort):

1. Versuche Office/Outlook-Theme-Registry-Werte (wenn verfuegbar).
2. Fallback auf Windows "app theme" (`AppsUseLightTheme`).
3. High-Contrast-Modus deaktiviert Custom-Theming (Systemfarben gewinnen).

## Build und Release

### Was `build.ps1` macht

1. Baut das COM Add-in (`NcTalkOutlookAddIn.sln`) via MSBuild
2. Liest die Assembly-Version aus `NcTalkOutlookAddIn.dll`
3. Baut den WiX-v4-Installer (`installer/NcConnectorOutlookInstaller.wixproj`)
4. Kopiert die MSI nach `dist/`

### Versionierung

- `src/NcTalkOutlookAddIn/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`

`build.ps1` leitet daraus die MSI-`ProductVersion` ab (Format `Major.Minor.Build`).

### Release-Checkliste

1. Version in `AssemblyInfo.cs` erhoehen
2. Falls vendorte Dependencies geaendert wurden: `VENDOR.md` aktualisieren
3. `./build.ps1 -Configuration Release`
4. MSI Install/Upgrade testen (alte Version -> neue Version)
5. Smoke-Test (Talk + Sharing + IFB)
6. Optional: MSI signieren (falls in deiner Umgebung erforderlich)

## Lokales Testen

Empfohlene Smoke-Test-Sequenz:

Hinweis: Es gibt aktuell keine automatisierte Test-Suite in diesem Repository. Nutze die folgenden Smoke-Tests zur Validierung von Aenderungen.

1. Debug-Logging in den Settings aktivieren.
2. Kalender: neuen Termin erstellen, Talk-Link einfuegen, Termin speichern, danach Startzeit aendern und erneut speichern (Lobby-Update).
3. Kalender: Teilnehmer hinzufuegen, erneut speichern (Teilnehmer-Sync).
4. Mail: Sharing-Wizard ausfuehren, 1-2 kleine Dateien hochladen, HTML-Block einfuegen und an dich selbst senden.
5. IFB: IFB aktivieren, dann pruefen, ob der lokale Endpunkt antwortet:
   - `Invoke-WebRequest http://127.0.0.1:<ifb-port>/nc-ifb/ -UseBasicParsing`

## X-NCTALK-* Property-Referenz

Das Add-in persistiert Termin-Metadaten als Outlook `UserProperties`. Ein Teil dieser Properties nutzt `X-NCTALK-*`-Namen, um Interoperabilitaet und stabiles Re-Sync-Verhalten sicherzustellen.

Sofern nicht anders angegeben:

- Properties werden als **Textwerte** in Outlook gespeichert (`OlUserPropertyType.olText`).
- Boolesche Werte werden als `TRUE` / `FALSE` (grossgeschrieben) gespeichert.
- Zeitstempel werden als **Unix-Epoch-Sekunden** (UTC) in Invariant-Culture gespeichert.

Primaerer Write-Ort:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs` -> `ApplyRoomToAppointment(...)`

### Properties

| Property | Zweck | Typ / Format | Beispiel | Geschrieben | Gelesen / verwendet | Hinweise |
| --- | --- | --- | --- | --- | --- | --- |
| `X-NCTALK-TOKEN` | Talk-Room-Token | `string` | `a1b2c3d4` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Lesen bevorzugt gegenueber Legacy-Token-Storage. |
| `X-NCTALK-URL` | Talk-Room-URL | `string` | `https://cloud.example.com/call/a1b2c3d4` | `ApplyRoomToAppointment(...)` | (nicht vom Add-in gelesen) | Fuer Interoperabilitaet gespeichert. |
| `X-NCTALK-LOBBY` | Lobby-enabled Flag | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `EnsureSubscriptionForAppointment(...)` | Steuert, ob Lobby-Updates bei Save laufen. |
| `X-NCTALK-START` | Termin-Startzeit (Epoch-Sekunden) | `int64` als String | `1739750400` | `ApplyRoomToAppointment(...)`, `AppointmentSubscription.OnWrite(...)` | `TryReadRequiredIcalStartEpoch(...)`, `TryUpdateLobby(...)` | Autoritative Lobby-Timer-Quelle bei Save; wird aktualisiert, wenn Lobby aktiv ist und sich Startzeit aendert. |
| `X-NCTALK-EVENT` | Room-Creation-Mode-Marker | `event` \| `standard` | `event` | `ApplyRoomToAppointment(...)` | `GetRoomType(...)` | Fallback ueber `NcTalkRoomType`-UserProperty vorhanden. |
| `X-NCTALK-OBJECTID` | Time-window-Identifier | `"<start>#<end>"` | `1739750400#1739754000` | `ApplyRoomToAppointment(...)` | (nicht vom Add-in gelesen) | Fuer Interoperabilitaet gespeichert. |
| `X-NCTALK-ADD-USERS` | Teilnehmer-Sync: interne User | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Bevorzugt gegenueber Legacy `X-NCTALK-ADD-PARTICIPANTS`. |
| `X-NCTALK-ADD-GUESTS` | Teilnehmer-Sync: externe E-Mails | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Bevorzugt gegenueber Legacy `X-NCTALK-ADD-PARTICIPANTS`. |
| `X-NCTALK-ADD-PARTICIPANTS` | Teilnehmer-Sync: Legacy kombinierter Toggle | `TRUE` / `FALSE` | `TRUE` | `ApplyRoomToAppointment(...)` | `TrySyncRoomParticipants(...)` | Als Fallback genutzt, wenn gesplittete Toggles fehlen (aeltere Items). |
| `X-NCTALK-DELEGATE` | Delegationsziel-User-ID | `string` | `alice` | `ApplyRoomToAppointment(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)`, `TryApplyDelegation(...)` | Parallel zur `NcTalkDelegateId`-UserProperty fuer Backward-Compatibility gespeichert. |
| `X-NCTALK-DELEGATE-NAME` | Delegationsziel-Display-Name | `string` | `Alice Example` | `ApplyRoomToAppointment(...)` | (nicht vom Add-in gelesen) | Fuer Interoperabilitaet gespeichert. |
| `X-NCTALK-DELEGATED` | Delegations-Status-Marker | `TRUE` / `FALSE` | `FALSE` | `ApplyRoomToAppointment(...)`, `TryApplyDelegation(...)` | `IsDelegatedToOtherUser(...)`, `IsDelegationPending(...)` | Steuert, ob Delegation noch pending ist. |
| `X-NCTALK-DELEGATE-READY` | Delegation-"ready"-Marker | `TRUE` | `TRUE` | `ApplyRoomToAppointment(...)` | (nicht vom Add-in gelesen) | Fuer Interoperabilitaet reserviert; Add-in nutzt aktuell `X-NCTALK-DELEGATED` + Delegate-ID fuer Pending-Detection. |

## Erweiterungspunkte

### Neues Setting hinzufuegen

1. Property in `src/NcTalkOutlookAddIn/Settings/AddinSettings.cs` hinzufuegen.
2. In `src/NcTalkOutlookAddIn/Settings/SettingsStorage.cs` persistieren.
3. UI in `src/NcTalkOutlookAddIn/UI/SettingsForm.cs` hinzufuegen.
4. Uebersetzungen ergaenzen (siehe `Translations.md`).

### Neuen Nextcloud-API-Call hinzufuegen

1. Im passenden Service ergaenzen:
   - Talk: `src/NcTalkOutlookAddIn/Services/TalkService.cs`
   - Sharing: `src/NcTalkOutlookAddIn/Services/FileLinkService.cs`
   - Fuer neue OCS/JSON-Calls `Services/NcHttpClient.cs` + `Utilities/NcJson.cs` nutzen statt service-lokale Request/Parsing-Helper einzufuehren.
2. Request/Response-Modell in `src/NcTalkOutlookAddIn/Models/` ergaenzen (falls noetig).
3. Logging-Scopes und Error-Handling ergaenzen.
4. In UI/Wizard integrieren und ueber `NextcloudTalkAddIn.cs` verdrahten.
5. Fuer ribbon-getriggerte Flows Orchestrierung bevorzugt im passenden Controller ergaenzen (`SettingsWorkflowController`, `FileLinkLaunchController`, `TalkRibbonController`) und `NextcloudTalkAddIn.cs` als schlanke Delegate-Schicht halten.

### Neuen lokalisierten String hinzufuegen

1. Property in `src/NcTalkOutlookAddIn/Utilities/Strings.cs` hinzufuegen.
2. Key in allen Locale-Dateien unter `src/NcTalkOutlookAddIn/Resources/_locales/` ergaenzen.
3. Rebuild und UI verifizieren.