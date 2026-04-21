# DEVELOPMENT.de.md — NC Connector for Outlook

Dieses Dokument richtet sich an Entwickler und beschreibt Aufbau, Build und Release-Prozess des **NC Connector for Outlook** (Outlook classic COM Add-in).

## Inhalt

- [Projekt-Überblick](#projekt-überblick)
- [Voraussetzungen](#voraussetzungen)
- [Build (MSI)](#build-msi)
- [Lokales Testen](#lokales-testen)
- [Logging](#logging)
- [Code-Struktur](#code-struktur)
- [Versionierung & Release](#versionierung--release)

## Projekt-Überblick

Das Add-in integriert:

- **Nextcloud Talk** direkt aus dem Termin (Raum erstellen, Lobby, Moderator-Delegation, Teilnehmer-Automation)
- **Nextcloud Filelink** im E-Mail-Composer (Wizard, Upload, HTML-Block)
- **IFB (Internet Free/Busy)** als lokaler HTTP-Proxy zu Nextcloud

## Release 3.0.3 Delta-Ueberblick

In 3.0.3 gelten die paritaetskritischen Verhaltensregeln weiterhin und muessen bei Folgeaenderungen stabil bleiben:

- Compose-Anhangsautomatisierung mit deterministischen Modi (`always` vs. Schwellwert-Prompt).
- Compose-Anhangsautomatisierung prueft zusaetzlich pre-add (`BeforeAttachmentAdd`) und kann Host-Adds best effort vor der normalen Outlook-Post-Add-Verarbeitung abbrechen, wenn ein lokaler Pfad aufloesbar ist.
- Schwellwert-Prompt strikt mit zwei Aktionen (`Share with NC Connector` / `Remove last selected attachments`) und Batch-Entfernungssemantik.
- Spezialisierter Attachment-Share-Output (`email_attachment`-Namensvertrag, read-only Empfaengerrechte, ZIP-Link `/s/<token>/download`, keine Permissions-Zeile).
- Compose-Share-Cleanup mit Lifecycle-Vertrag (armed nach Share-Erstellung, cleared nur nach bestaetigtem Versand, delayed delete bei unsent-close-Race).
- Separater Passwort-Follow-up-Versand nur post-send inklusive Auto-Send und manuellem Fallback-Entwurf.
- Talk-Defaults und Talk-Wizard mit zentralem Systemadressbuch-Verfuegbarkeitsvertrag und deterministischem Lock-Zustand.
- Public-Link-Share-Erstellung folgt dem dokumentierten Nextcloud-OCS-Vertrag: `label` beim Create, veränderliche Metadaten wie `note` danach über den OCS-Update-Endpunkt.
- Runtime-Settings und Caches unter `%LOCALAPPDATA%\\NC4OL` mit profilbasierter XML-Migration.
- TLS-Transporteinstellungen sind in `SettingsForm` runtime-live schaltbar (`OS-Default` / `TLS 1.2` / `TLS 1.3` / kombiniert) und werden beim Anwenden bewusst validiert.
- Nicht unterstuetzte manuelle `TLS 1.3`-Anwendung faellt nicht mehr still auf TLS 1.2 zurueck, sondern bricht explizit mit Fehler ab.
- Verbindungspruefung und Login-Flow erzwingen frische HTTP/TLS-Handshakes (kein Keep-Alive-Reuse), damit TLS-Moduswechsel deterministisch getestet werden.
- Die TLS-Fehlerhilfe in der Verbindungsdiagnose ist modusneutral und behauptet nicht mehr pauschal `OS-Default`.

## Voraussetzungen

- Windows 10/11
- Outlook classic (typischerweise x64)
- **.NET Framework 4.7.2** (Target)
- MSBuild (z.B. Visual Studio Build Tools)
- **.NET SDK** (für WiX v4 Build via `dotnet`)

### Reference Assemblies (FrameworkPathOverride)

Auf manchen Build-Systemen fehlen die .NET Framework Reference Assemblies für 4.7.2 (insbesondere CI/Minimal-Installationen). In dem Fall kann man die NuGet-ReferenceAssemblies nutzen und `FrameworkPathOverride` setzen.

Beispiel:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.4"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

## Build (MSI)

Der empfohlene Build läuft immer über `build.ps1`:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.4"
$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
.\build.ps1 -Configuration Release
```

Wenn auf dem Build-Host die WiX-ICE-Validierung nicht verfuegbar ist (z. B. `WIX0217` in eingeschraenkten Umgebungen), verwende:

```powershell
.\build.ps1 -Configuration Release -SkipIceValidation
```

Output:

- `dist\NCConnectorForOutlook-<version>.msi`

Was das Script macht:

1) Build des COM Add-ins (`NcTalkOutlookAddIn.sln`) via MSBuild
2) Ermittelt die Assembly-Version aus `NcTalkOutlookAddIn.dll`
3) Build des WiX v4 Installers (`installer/NcConnectorOutlookInstaller.wixproj`)
4) Kopiert das MSI in `dist/`

## Lokales Testen

1) MSI installieren (als Admin):
   - `msiexec /i dist\NCConnectorForOutlook-<version>.msi`
2) Outlook starten
3) Kalendertermin öffnen:
   - Ribbon: **NC Connector → Talk-Link einfügen**
4) E-Mail erstellen:
   - Ribbon: **NC Connector → Nextcloud Freigabe hinzufügen**
5) Optional: IFB in Settings aktivieren, Port pruefen (`Einstellungen -> IFB`, Standard `7777`) und Endpunkt testen
   - `Invoke-WebRequest http://127.0.0.1:<ifb-port>/nc-ifb/ -UseBasicParsing`

## Logging

- Aktivierung: Settings → Tab **Debug**
- Option (Standard aktiv): `Logs anonymisieren`
- Datei (taeglich): `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`
- Runtime-Exceptions werden ueber `DiagnosticsLogger.LogException(...)` immer geschrieben, auch wenn Debug deaktiviert ist.
- Aufbewahrung: letzte 7 Tageslogs behalten, Logs aelter als 30 Tage (best effort) entfernen.
- Bei aktiver Anonymisierung werden NC-URL/Basis-Host, Token/Secrets, Authorization-Werte, E-Mails, Benutzerkennungen und lokale User-Pfadsegmente maskiert.

Kategorien (Beispiele):

- `CORE` (Start, Settings, Registry)
- `API` (HTTP Calls / Statuscodes)
- `TALK` (Room Lifecycle, Lobby, Delegation)
- `FILELINK` (Upload/Share)
- `IFB` (Requests, Cache, Outlook Registry)

## Code-Struktur

Root:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs`  
  Einstiegspunkt, Ribbon, Outlook-Events, Composition Root fuer die Workflows.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs`  
  Add-in-Bootstrap/Teardown (`OnConnection`, Shutdown/Disconnect).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs`
  Dedizierte Outlook-Event Hook-/Unhook-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs`
  Kategorienspezifische Runtime-Logging-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs`  
  Backend-Policy- und Talk-Template-/Sprach-Resolver.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs`  
  Deferred Appointment-Subscription-Ensure inkl. Outlook-Event-Restriction-Handling.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs`
  Runtime-Subscription-Core fuer Compose-Lifecycle-Zustand (`Dispose`, Identity, gemeinsame Helper).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs`
  Compose-Attachment-Interception/Evaluation/Share-Launch-Flow.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs`
  Send/Close-Cleanup-Lifecycle inkl. separatem Passwort-Dispatch.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs`
  Runtime-Subscription fuer Termin-Write/Close/Delete und Lifecycle-Cleanup.
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs`
  Orchestrierung fuer Settings-Open/Save/Revert.
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs`
  Orchestrierung fuer FileLink-Ribbon-Start und Wizard-Flow.
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs`
  Orchestrierung fuer Talk-Ribbon-Flow (Auth-Gate, Wizard, Room-Create/Replace).
- `TalkRibbonController` und `FileLinkLaunchController` prefetchen Backend-Policy + Passwort-Policy parallel (`Task.WhenAll`) vor dem Wizard-Open; Policy-Daten bleiben dabei pro Einstiegspunkt immer frisch.

Controller:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` (Talk-Termin-Lifecycle: Room-Metadaten, Lobby-/Description-/Delegation-/Teilnehmer-Sync)
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` (Compose-Share-Cleanup und separater Passwort-Versand inkl. Fallback)
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` (Talk-Template-/Block-Rendering)
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` (SMTP- und Attendee-Aufloesung)
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` (Compose-Subscription-Registry)
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` (gemeinsame Mail-/Inspector-Interop-Helper)
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` (gemeinsames Queue-Modell fuer separaten Passwort-Follow-up)

Services:

- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (DAV/Share Flow)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB; Port ueber Settings konfigurierbar, Standard `7777`)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)
- `src/NcTalkOutlookAddIn/Services/NcHttpClient.cs` (zentraler Request-Executor fuer Auth-Header, OCS-Header, Timeout/Decompression und optionalen Fresh-Connection-Mode)
  - Alle Runtime-HTTP-Aufrufe (Talk, Share/DAV, IFB, Login-Flow, Moderator-Avatar-Fetch) laufen zentral ueber `NcHttpClient`.

UI:

- `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`
- `src/NcTalkOutlookAddIn/UI/TalkLinkForm.cs`
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
- `src/NcTalkOutlookAddIn/UI/ComposeAttachmentPromptForm.cs` (2-Aktions-Prompt fuer Schwellwertmodus)
- `src/NcTalkOutlookAddIn/UI/BrandedHeader.cs` (Header-Banner inkl. `AttachToParent(...)` fuer konsistente Header-Initialisierung in Forms)
- `src/NcTalkOutlookAddIn/UI/ScaledForm.cs` (zentrale DPI-Skalierung via `ScaleLogical(...)`, damit Form-Wrapper nicht dupliziert werden)

Utilities:

- `src/NcTalkOutlookAddIn/Utilities/BrowserLauncher.cs` (zentraler Shell-Start fuer URLs, Dateien und Ordner)
- `src/NcTalkOutlookAddIn/Utilities/SizeFormatting.cs` (zentrale MB-Formatierung fuer UI-Texte)
- `src/NcTalkOutlookAddIn/Utilities/ComInteropScope.cs` (zentrale COM-Release-/FinalRelease-Helfer)
- `src/NcTalkOutlookAddIn/Utilities/PasswordGenerationHelper.cs` (zentralisiert Min-Length-Aufloesung, Server-Fallback-Generierung und gemeinsame Min-Length-Validierung fuer Talk/FileLink-Formulare)
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` (zentraler Sanitizer fuer Backend-HTML-Templates bei Share/Talk, fail-closed)
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` (zentrale JSON-Normalisierung inkl. `PrepareJsonPayload`, Dictionary-/String-/Int-Helfer und OCS-Fehlerextraktion)
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` (gekapselter Laufzeitzustand fuer Pending-Keys und Restriction-Log-Throttling)
- `src/NcTalkOutlookAddIn/Utilities/PictureConverter.cs` (gemeinsamer Image->IPictureDisp-Helfer fuer Ribbon-Icons)

Compose-Filelink-Paritaet (3.0.3):

- `MailComposeSubscription` in `NextcloudTalkAddIn.cs` steuert den Compose-Lifecycle fuer:
  - debouncte Anhangsauswertung (`ComposeAttachmentEvalDebounceMs`)
  - Pre-Add-Abfangpfad (`BeforeAttachmentAdd`) fuer fruehes Intercept
  - best-effort Abbruch des Host-Adds vor der normalen Outlook-Post-Add-Verarbeitung
  - harte Outlook-/Exchange-Groessenlimits koennen trotzdem vor Add-in-Callbacks greifen und sind ueber offizielle Outlook-OOM-Events nicht abfangbar
  - Always-via-NC und Schwellwertmodus
  - Batch-Entfernung (`Remove last selected attachments`)
  - Attachment-Mode-Wizardstart direkt im Datei-Schritt
  - Share-Cleanup bei unsent close inkl. Grace-Timer fuer Send/Close-Race
  - separates Passwort-Follow-up nach bestaetigtem erfolgreichem Hauptversand.
- `ComposeShareLifecycleController` kapselt die eigentliche Share-Cleanup-/Passwort-Dispatch-Logik; `MailComposeSubscription` haelt nur Queue- und Eventzustand.
- `TalkAppointmentController` kapselt Appointment-Schreib-/Sync-Pfade; `NextcloudTalkAddIn` delegiert diese Aufrufe statt die komplette Fachlogik im Root zu halten.
- Ribbon-getriggerte Flows werden im Controller-Slice gehalten (`SettingsWorkflowController`, `FileLinkLaunchController`, `TalkRibbonController`); `NextcloudTalkAddIn.cs` bleibt schlanke Delegate-/Composition-Root-Schicht.
  - Lifecycle-, Policy-/Template- und Deferred-Ensure-Logik sind in eigene Partial-Dateien ausgelagert, damit die Root-Klasse wartbar bleibt.
  - Custom-Talk-Templates aus dem Backend werden vor HTML-/Plain-Text-Rendering ueber `HtmlTemplateSanitizer` bereinigt (kein Raw-HTML-Fallback).
  - fuer Talk-Termine laeuft vor dem Insert ein expliziter Compat-Transform (`HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(...)`)
  - Appointment-HTML wird ueber HTML->RTF-Bridge geschrieben (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`), nicht ueber `AppointmentItem.HTMLBody` und nicht ueber `HTMLEditor.body.innerHTML`.
- `OutlookAttachmentAutomationGuardService` erzwingt den Host-Konflikt-Guard live:
  - vor Auswertung
  - vor Prompt-Aktionsverarbeitung
  - vor Wizard-Finalize im Attachment-Modus.
- `FileLinkHtmlBuilder` erzeugt im Attachment-Modus reduziertes HTML mit ZIP-Link `/s/<token>/download`.
- Custom-Share-Templates aus dem Backend werden im `FileLinkHtmlBuilder` vor der Einfuegung ueber `HtmlTemplateSanitizer` bereinigt (fail-closed).
- `FileLinkWizardForm` akzeptiert im Datei-Schritt Explorer-Drag-and-drop fuer Dateien/Ordner ueber Queue und Aktionsbereich.

### Appointment-sicheres HTML-Subset fuer Talk-Templates

Damit Backend-Talk-Templates in Outlook-Terminen stabil gerendert werden (Word/RTF-Pipeline), gilt:

- Layout bevorzugt tabellenbasiert aufbauen (`table`, `tbody`, `tr`, `td`).
- Inline-Styles sind erlaubt, aber Word-kritische CSS-Features werden im Appointment-Compat-Transform entfernt:
  - `display:flex|grid`, `flex*`, `grid*`, `border-radius*`, `overflow*`, `object-fit`, `user-select` (inkl. vendor-prefix Varianten).
- Farbausrichtung bekommt zusaetzliche Legacy-Fallbacks:
  - `style=color` -> `<font color=...>`
  - `style=background-color` -> `bgcolor`
  - `style=text-align` -> `align`
  - `style=vertical-align` -> `valign`
- Linkfarbe wird zusaetzlich abgesichert (`<a><font color=...>...</font></a>`), falls erforderlich.
- Unsichere/nicht erlaubte Tags/Attribute entfernt der Sanitizer weiterhin fail-closed.

Installer:

- `installer/NcConnectorOutlookInstaller.wixproj` (WiX v4 SDK Projekt)
- `installer/Product.wxs` (MSI Definition: Dateien + Registry + URLACL)
- `VENDOR.md` (Lizenzhinweise fuer gebuendelte Drittanbieter-Abhaengigkeiten)

## Versionierung & Release

### Version bump

- `src/NcTalkOutlookAddIn/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`

`build.ps1` leitet daraus die MSI `ProductVersion` ab (Format `Major.Minor.Build`).

### MSI Upgrade-Kompatibilität

Wichtig für Updates:

- UpgradeCode bleibt stabil (siehe `installer/Product.wxs`)
- COM GUID / ProgId bleiben stabil (siehe `NextcloudTalkAddIn.cs`)

### Release Checklist

1) Version bump
2) Bei geaenderten vendorten Abhaengigkeiten: `VENDOR.md` aktualisieren
3) `.\build.ps1 -Configuration Release`
4) MSI installieren/upgrade testen (alte Version → neue Version)
5) Talk + Filelink + IFB Smoke-Test
6) MSI ggf. signieren (falls in der Umgebung erforderlich)


