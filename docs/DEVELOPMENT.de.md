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

## Release 3.0.1 Delta-Ueberblick

In 3.0.1 gelten die paritaetskritischen Verhaltensregeln weiterhin und muessen bei Folgeaenderungen stabil bleiben:

- Compose-Anhangsautomatisierung mit deterministischen Modi (`always` vs. Schwellwert-Prompt).
- Compose-Anhangsautomatisierung prueft zusaetzlich pre-add (`BeforeAttachmentAdd`) und kann Host-Adds best effort vor der normalen Outlook-Post-Add-Verarbeitung abbrechen, wenn ein lokaler Pfad aufloesbar ist.
- Schwellwert-Prompt strikt mit zwei Aktionen (`Share with NC Connector` / `Remove last selected attachments`) und Batch-Entfernungssemantik.
- Spezialisierter Attachment-Share-Output (`email_attachment`-Namensvertrag, read-only Empfaengerrechte, ZIP-Link `/s/<token>/download`, keine Permissions-Zeile).
- Compose-Share-Cleanup mit Lifecycle-Vertrag (armed nach Share-Erstellung, cleared nur nach bestaetigtem Versand, delayed delete bei unsent-close-Race).
- Separater Passwort-Follow-up-Versand nur post-send inklusive Auto-Send und manuellem Fallback-Entwurf.
- Talk-Defaults und Talk-Wizard mit zentralem Systemadressbuch-Verfuegbarkeitsvertrag und deterministischem Lock-Zustand.
- Public-Link-Share-Erstellung folgt dem dokumentierten Nextcloud-OCS-Vertrag: `label` beim Create, veränderliche Metadaten wie `note` danach über den OCS-Update-Endpunkt.
- Runtime-Settings und Caches unter `%LOCALAPPDATA%\\NC4OL` mit profilbasierter XML-Migration.

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
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.1"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

## Build (MSI)

Der empfohlene Build läuft immer über `build.ps1`:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.1"
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
5) Optional: IFB in Settings aktivieren und Free/Busy prüfen

## Logging

- Aktivierung: Settings → Tab **Debug**
- Datei: `%LOCALAPPDATA%\NC4OL\addin-runtime.log`
- Runtime-Exceptions werden ueber `DiagnosticsLogger.LogException(...)` immer geschrieben, auch wenn Debug deaktiviert ist.

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
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs`
  Runtime-Subscription fuer Compose-Attachments/Share-Cleanup/Password-Dispatch.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs`
  Runtime-Subscription fuer Termin-Write/Close/Delete und Lifecycle-Cleanup.

Controller:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` (Talk-Termin-Lifecycle: Room-Metadaten, Lobby-/Description-/Delegation-/Teilnehmer-Sync)
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` (Compose-Share-Cleanup und separater Passwort-Versand inkl. Fallback)
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` (Talk-Template-/Block-Rendering)
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` (SMTP- und Attendee-Aufloesung)
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` (Compose-Subscription-Registry)

Services:

- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (DAV/Share Flow)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)

UI:

- `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`
- `src/NcTalkOutlookAddIn/UI/TalkLinkForm.cs`
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
- `src/NcTalkOutlookAddIn/UI/ComposeAttachmentPromptForm.cs` (2-Aktions-Prompt fuer Schwellwertmodus)
- `src/NcTalkOutlookAddIn/UI/BrandedHeader.cs` (Header-Banner)

Compose-Filelink-Paritaet (3.0.1):

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
- `OutlookAttachmentAutomationGuardService` erzwingt den Host-Konflikt-Guard live:
  - vor Auswertung
  - vor Prompt-Aktionsverarbeitung
  - vor Wizard-Finalize im Attachment-Modus.
- `FileLinkHtmlBuilder` erzeugt im Attachment-Modus reduziertes HTML mit ZIP-Link `/s/<token>/download`.
- `FileLinkWizardForm` akzeptiert im Datei-Schritt Explorer-Drag-and-drop fuer Dateien/Ordner ueber Queue und Aktionsbereich.

Installer:

- `installer/NcConnectorOutlookInstaller.wixproj` (WiX v4 SDK Projekt)
- `installer/Product.wxs` (MSI Definition: Dateien + Registry + URLACL)

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
2) `.\build.ps1 -Configuration Release`
3) MSI installieren/upgrade testen (alte Version → neue Version)
4) Talk + Filelink + IFB Smoke-Test
5) MSI ggf. signieren (falls in der Umgebung erforderlich)

