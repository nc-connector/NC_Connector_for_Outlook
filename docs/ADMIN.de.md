# ADMIN.de.md — NC Connector for Outlook

Dieses Dokument beschreibt Installation, Rollout und Betrieb des **NC Connector for Outlook** (Outlook classic COM Add-in).

## Inhalt

- [Installation (MSI)](#installation-msi)
- [Updates / Upgrade-Verhalten](#updates--upgrade-verhalten)
- [Dateien & Registry](#dateien--registry)
- [Einstellungen (settings.ini)](#einstellungen-settingsini)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [Logging / Support](#logging--support)
- [Troubleshooting](#troubleshooting)

## Installation (MSI)

1) Outlook beenden.  
2) MSI installieren (Administrator-Rechte erforderlich).

Beispiel (silent):

```powershell
msiexec /i "NCConnectorForOutlook-2.2.7.msi" /qn /norestart
```

Danach Outlook starten. Im Ribbon erscheint der Tab **NC Connector** (Kalender/Termin + E-Mail).

## Updates / Upgrade-Verhalten

- Updates erfolgen durch Installation einer **neueren MSI-Version**.
- Die MSI ist als **Major Upgrade** konfiguriert (UpgradeCode bleibt stabil), damit eine vorhandene Installation automatisch ersetzt wird.
- Benutzer-Einstellungen bleiben erhalten, da sie im Benutzerprofil gespeichert sind.

## Dateien & Registry

### Installationspfad

Standard (x64):

- `C:\Program Files\NextcloudTalkOutlookAddIn\`

Wichtige Dateien:

- `NcTalkOutlookAddIn.dll` (COM Add-in)
- `NcTalkOutlookAddIn.dll.config` (Binding Redirects)
- `LICENSE.txt`

### Outlook Add-in Registrierung (HKLM)

Das Add-in wird per MSI unter anderem hier registriert:

- `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

Wichtige Werte:

- `LoadBehavior=3` (geladen)
- `FriendlyName="NC Connector for Outlook"`
- `Description="Nextcloud Talk & Files nahtlos in Outlook integriert"`

Die COM-Registrierung erfolgt über `HKLM\Software\Classes\CLSID\{...}` inklusive `CodeBase` auf das installierte DLL.

## Einstellungen (settings.ini)

### Speicherort

Einstellungen werden pro Benutzer gespeichert in:

- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini`

Legacy-Migration:

- Falls noch vorhanden, wird `%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\settings.ini` beim ersten Start übernommen.

### Wichtige Keys

Beispiele (Auszug):

```ini
ServerUrl=https://cloud.example.com
Username=max
AppPassword=xxxxxx
AuthMode=LoginFlow
IfbEnabled=true
IfbDays=30
IfbCacheHours=24
DebugLoggingEnabled=false
FileLinkBasePath=90 Freigaben - extern
TalkDefaultLobbyEnabled=true
TalkDefaultSearchVisible=true
TalkDefaultRoomType=EventConversation
TalkDefaultPasswordEnabled=true
TalkDefaultAddUsers=true
TalkDefaultAddGuests=false
```

Hinweis: Das App-Passwort wird aktuell in Klartext gespeichert (geplante Verbesserung: Windows Credential Locker / DPAPI).

### Rollout / Pre-Seed

Da `settings.ini` im Benutzerprofil liegt, gibt es mehrere typische Wege:

- **Login Script / Intune/SCCM**: Datei nach `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini` kopieren (nur falls noch nicht vorhanden).
- **Group Policy Preferences**: Datei in das Benutzerprofil verteilen.

Empfehlung:

- Nur Base-URL und Defaults pre-seeden.
- Credentials entweder über Login Flow v2 oder per Benutzer setzen lassen.

## Internet Free/Busy Gateway (IFB)

### Zweck

IFB stellt eine lokale HTTP-Quelle bereit, über die Outlook Free/Busy-Informationen aus Nextcloud beziehen kann.

Standard-Endpunkt:

- `http://127.0.0.1:7777/nc-ifb/`

### URLACL (HTTP.SYS Reservation)

Die MSI reserviert den URL-Namespace für alle authentifizierten Benutzer, damit der Listener ohne Admin-Rechte laufen kann.

Prüfen:

```powershell
netsh http show urlacl | Select-String -Pattern "7777/nc-ifb"
```

### Outlook Registry (per User)

Beim Aktivieren von IFB setzt das Add-in Outlook-spezifische Werte (HKCU), damit Outlook die Free/Busy-URL verwendet.

Hinweis:

- IFB ist optional (per Settings UI).
- Ohne IFB läuft das Add-in weiterhin für Talk + Filelink.

## Logging / Support

Debug-Logging ist im Settings-Tab **Debug** aktivierbar.

Log-Datei:

- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`

Die Logs sind kategorisiert (z.B. `CORE`, `API`, `TALK`, `FILELINK`, `IFB`) und helfen bei Supportfällen.

## Troubleshooting

### Add-in wird nicht geladen

1) In Outlook: `Datei → Optionen → Add-Ins`
2) Bereich „COM-Add-Ins“ prüfen: `NcTalkOutlook.AddIn`
3) „Deaktivierte Elemente“ prüfen (Outlook deaktiviert Add-ins bei Abstürzen)

### IFB reagiert nicht

- Prüfen, ob Port 7777 gebunden ist:

```powershell
netstat -ano | Select-String ":7777"
```

- URLACL prüfen (siehe oben)
- Debug-Log aktivieren und `IFB`-Einträge prüfen

### Netzwerk / Nextcloud

- Server erreichbar, TLS ok?
- App-Passwort gültig?
- Talk installiert?
- Password Policy App optional: bei fehlender App wird lokal generiert (Fallback)
