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
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-0.2.7"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

## Build (MSI)

Der empfohlene Build läuft immer über `build.ps1`:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-0.2.7"
$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
.\build.ps1 -Configuration Release
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
- Datei: `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`

Kategorien (Beispiele):

- `CORE` (Start, Settings, Registry)
- `API` (HTTP Calls / Statuscodes)
- `TALK` (Room Lifecycle, Lobby, Delegation)
- `FILELINK` (Upload/Share)
- `IFB` (Requests, Cache, Outlook Registry)

## Code-Struktur

Root:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs`  
  Einstiegspunkt, Ribbon, Outlook-Events, Orchestrierung der Workflows.

Services:

- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (DAV/Share Flow)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)

UI:

- `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`
- `src/NcTalkOutlookAddIn/UI/TalkLinkForm.cs`
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
- `src/NcTalkOutlookAddIn/UI/BrandedHeader.cs` (Header-Banner)

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
