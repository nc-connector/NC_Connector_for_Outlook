# ADMIN.de.md — NC Connector for Outlook

Dieses Dokument beschreibt Installation, Rollout und Betrieb des **NC Connector for Outlook** (Outlook classic COM Add-in).

## Inhalt

- [Installation (MSI)](#installation-msi)
- [Updates / Upgrade-Verhalten](#updates--upgrade-verhalten)
- [Dateien & Registry](#dateien--registry)
- [Einstellungen (Profil-XML)](#einstellungen-profil-xml)
- [Compose-Freigabe-Lifecycle (3.0.3)](#compose-freigabe-lifecycle-303)
- [Talk-Termin-Templates (HTML) — Outlook-sicheres Subset](#talk-termin-templates-html--outlook-sicheres-subset)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [Systemadressbuch erforderlich fuer Benutzersuche und Moderator-Auswahl](#systemadressbuch-erforderlich-fuer-benutzersuche-und-moderator-auswahl)
- [Logging / Support](#logging--support)
- [Troubleshooting](#troubleshooting)

## Installation (MSI)

1) Outlook beenden.  
2) MSI installieren (Administrator-Rechte erforderlich).

Beispiel (silent):

```powershell
msiexec /i "NCConnectorForOutlook-3.0.3.msi" /qn /norestart
```

Danach Outlook starten. Im Ribbon erscheint der Tab **NC Connector** (Kalender/Termin + E-Mail).

## Updates / Upgrade-Verhalten

- Updates/Reinstallationen erfolgen durch Installation eines MSI-Pakets ueber die bestehende Installation (gleiche, aeltere oder neuere Version).
- Die MSI ist als **Major Upgrade** konfiguriert (UpgradeCode bleibt stabil), damit eine vorhandene Installation automatisch ersetzt wird.
- Benutzer-Einstellungen bleiben erhalten, da sie im Benutzerprofil gespeichert sind.

## Dateien & Registry

### Installationspfad

Standard (x64):

- `C:\Program Files\NC4OL\`

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

Installer-Markierung fuer die IFB-URL-Reservation:
- `HKLM\Software\NC4OL\HttpUrl`

## Einstellungen (Profil-XML)

### Speicherort

Einstellungen werden pro Benutzer und Outlook-Profil gespeichert:

- `%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml`
- Fallback-Datei (wenn kein Profilname verfuegbar ist): `%LOCALAPPDATA%\NC4OL\settings_default.xml`

Passwort-Speicherung:

- `AppPasswordProtected` wird per Windows DPAPI (`CurrentUser`) verschluesselt gespeichert.
- Ein Klartext-`AppPassword` wird im neuen Format nicht mehr persistiert.

Legacy-Migration:

- `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\settings.ini`
- `%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\settings.ini`
- Legacy-INI-Dateien werden nach erfolgreicher Migration entfernt.

### Wichtige Keys

Beispiel (Auszug):

```xml
<Settings SchemaVersion="1" Profile="Outlook">
  <ServerUrl>https://cloud.example.com</ServerUrl>
  <Username>max</Username>
  <AppPasswordProtected>BASE64_DPAPI_BLOB</AppPasswordProtected>
  <AuthMode>LoginFlow</AuthMode>
  <IfbEnabled>true</IfbEnabled>
  <IfbDays>30</IfbDays>
  <IfbPort>7777</IfbPort>
  <IfbCacheHours>24</IfbCacheHours>
  <DebugLoggingEnabled>false</DebugLoggingEnabled>
  <LogAnonymizationEnabled>true</LogAnonymizationEnabled>
  <FileLinkBasePath>NC Connector</FileLinkBasePath>
</Settings>
```

### Rollout / Pre-Seed

Da die Profil-XML im Benutzerprofil liegt, gibt es mehrere typische Wege:

- **Login Script / Intune/SCCM**: vorbereitete `settings_<OutlookProfile>.xml` nach `%LOCALAPPDATA%\NC4OL\` kopieren (nur falls noch nicht vorhanden).
- **Group Policy Preferences**: Datei in das Benutzerprofil verteilen.

Empfehlung:

- Nur Base-URL und Defaults pre-seeden.
- Credentials entweder ueber Login Flow v2 oder per Benutzer setzen lassen (empfohlen fuer DPAPI-Kompatibilitaet).

## Compose-Freigabe-Lifecycle (3.0.3)

### Attachment-Automatisierung und Cleanup-Vertrag
- Im Compose-Attachment-Modus werden serverseitige Artefakte direkt nach Share-Erstellung fuer Cleanup-Tracking registriert.
- Cleanup wird erst nach bestaetigtem erfolgreichem Hauptversand wieder entfernt.
- Wird das Compose-Fenster ohne erfolgreichen Versand geschlossen, loescht das Add-in die erzeugten Share-Ordner-Artefakte serverseitig (best effort, mit Grace-Timer fuer Send/Close-Race).
- Die Anhangsautomatisierung wertet neue Dateien sowohl pre-add (`BeforeAttachmentAdd`) als auch post-add aus; kann pre-add ein lokaler Dateipfad aufgeloest werden, kann der NC-Flow den Host-Add best effort vor der normalen Outlook-Post-Add-Verarbeitung abbrechen.
- In Microsoft-365-/Exchange-Umgebungen mit serverseitigen Nachrichtengroessenlimits kann Outlook grosse Anhaenge bereits vor den Add-in-Events blockieren; in diesen Faellen kann die Automatisierung technisch nicht greifen und der Benutzer soll stattdessen den Button `Nextcloud Freigabe hinzufuegen` verwenden.
- Im Datei-Schritt des Sharing-Wizards koennen Dateien und Ordner per Explorer-Drag-and-drop im gesamten Schrittbereich (Queue + Aktionsbereich) hinzugefuegt werden, nicht nur ueber die Add-Buttons.

### Separater Passwort-Follow-up-Versand
- Ist `Passwort separat senden` aktiv, enthaelt der Haupt-HTML-Block kein Inline-Passwort.
- Der Passwort-Follow-up-Versand startet erst nach bestaetigtem erfolgreichem Hauptversand.
- Versandstrategie:
  - zuerst automatischer Versand
  - bei Fehlern ein vorbefuellter manueller Fallback-Entwurf.

## Talk-Termin-Templates (HTML) — Outlook-sicheres Subset

Wenn fuer Talk-Terminbeschreibungen im Backend `event_description_type=html` genutzt wird, gilt fuer stabiles Outlook-Rendering:

- Templates werden zuerst sanitiziert (fail-closed, kein Raw-HTML-Fallback).
- Der Termin-Insert laeuft ueber HTML->RTF-Bridge (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`).
- Vor dem Insert laeuft ein expliziter Appointment-Compat-Transform:
  - bevorzugte Tabellenstruktur (`table`, `tbody`, `tr`, `td`)
  - Legacy-Fallbacks fuer Farben/Ausrichtung (`font color`, `bgcolor`, `align`, `valign`)
  - Linkfarbe wird bei Bedarf zusaetzlich als `<a><font color=...>...</font></a>` abgesichert
  - Word-kritische CSS-Features werden entfernt (`flex/grid`, `border-radius`, `overflow`, `object-fit`, `user-select`).

## Internet Free/Busy Gateway (IFB)

### Zweck

IFB stellt eine lokale HTTP-Quelle bereit, über die Outlook Free/Busy-Informationen aus Nextcloud beziehen kann.

Endpunkt:

- Konfigurierbar unter `Einstellungen -> IFB -> Lokaler IFB-Port` (Standard `7777`).
- Standard-Endpunkt: `http://127.0.0.1:7777/nc-ifb/`

### URLACL (HTTP.SYS Reservation)

Die MSI reserviert den URL-Namespace für alle authentifizierten Benutzer, damit der Listener ohne Admin-Rechte laufen kann.

Standard-Reservation pruefen:

```powershell
netsh http show urlacl | Select-String -Pattern "7777/nc-ifb"
```

Wenn ein eigener IFB-Port verwendet wird, URLACL fuer diesen Port hinterlegen (Admin-Shell):

```powershell
netsh http add urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/ user="S-1-1-0"
```

### Outlook Registry (per User)

Beim Aktivieren von IFB setzt das Add-in Outlook-spezifische Werte (HKCU), damit Outlook die Free/Busy-URL verwendet.

Hinweis:

- IFB ist optional (per Settings UI).
- Ohne IFB läuft das Add-in weiterhin für Talk + Filelink.

## Systemadressbuch erforderlich fuer Benutzersuche und Moderator-Auswahl

Die folgenden Funktionen brauchen ein erreichbares **Nextcloud-Systemadressbuch**:
- Moderator-Auswahl im Talk-Wizard
- Default `Benutzer hinzufuegen` in den Add-in-Einstellungen
- Default `Gaeste hinzufuegen` in den Add-in-Einstellungen

Wenn das Systemadressbuch nicht verfuegbar ist, werden diese Controls in der UI deaktiviert und mit Warnhinweis plus Setup-Link angezeigt.

Aktivierung in Nextcloud 31:
- `sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"`

Aktivierung in Nextcloud >= 32:
- Nextcloud -> Admin Settings -> Groupware -> System Address Book (aktivieren)

In beiden Versionen erforderlich:
- Nextcloud Admin Settings -> Sharing: Username-Autocomplete / Zugriff auf das Systemadressbuch aktivieren.

Reparaturhinweis (wenn im Admin-UI aktiv, aber faktisch nicht verfuegbar):
1. Reset + erneut aktivieren:
   - `sudo -E -u www-data php occ config:app:delete dav system_addressbook_exposed`
   - `sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"`
2. Systemadressbuch neu synchronisieren:
   - `sudo -E -u www-data php occ dav:sync-system-addressbook`
3. Endpoint pruefen:
   - `https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system/?export`

## Logging / Support

Debug-Logging ist im Settings-Tab **Debug** aktivierbar.
Im Debug-Tab gibt es zusaetzlich **Logs anonymisieren** (standardmaessig aktiviert).

Log-Dateien (taegliche Rotation):

- `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`

Die Logs sind kategorisiert (z.B. `CORE`, `API`, `TALK`, `FILELINK`, `IFB`) und helfen bei Supportfällen.
Wenn Debug-Logging aktiviert ist, werden auch Runtime-Entscheidungspfade (inkl. Attachment-Pre-Add-Gating und Fallback-Gruenden) in dieselbe Datei geschrieben; Runtime-Exceptions werden unabhaengig vom Debug-Schalter immer geloggt.
Wenn Anonymisierung aktiv ist, werden sensible Werte vor dem Schreiben maskiert:
- konfigurierte Nextcloud-URL/Basis-Host
- Token/Secrets in URLs, Query-Parametern und JSON-Fragmenten
- `Authorization`-Header-Werte
- E-Mail-Adressen und typische Benutzerkennungen in Logfeldern
- lokale Benutzerpfade (z.B. `C:\\Users\\<USER>\\...`)
Aufbewahrung:
- die letzten 7 Tageslogs bleiben erhalten
- zusaetzlich werden Logs aelter als 30 Tage (best effort) entfernt

## Troubleshooting

### Add-in wird nicht geladen

1) In Outlook: `Datei → Optionen → Add-Ins`
2) Bereich „COM-Add-Ins“ prüfen: `NcTalkOutlook.AddIn`
3) „Deaktivierte Elemente“ prüfen (Outlook deaktiviert Add-ins bei Abstürzen)

### IFB reagiert nicht

- Pruefen, welcher IFB-Port in `Einstellungen -> IFB` gesetzt ist (Standard `7777`).
- Pruefen, ob dieser Port gebunden ist:

```powershell
netstat -ano | Select-String ":<ifb-port>"
```

- URLACL prüfen (siehe oben)
- Debug-Log aktivieren und `IFB`-Einträge prüfen

### Netzwerk / Nextcloud

- Server erreichbar, TLS ok?
- TLS-Verhalten kann im Add-in unter `Einstellungen -> Erweitert -> Transportsicherheit (TLS)` umgeschaltet werden (`OS-Default` oder erzwungene TLS-Versionen wie 1.2/1.3).
- NC Connector setzt die Auswahl zur Laufzeit add-in-lokal ueber `ServicePointManager.SecurityProtocol`.
- Die Verbindungsdiagnose (Verbindungstest in den Einstellungen und Login-Flow) erzwingt nun frische HTTP/TLS-Handshakes, damit TLS-Moduswechsel deterministisch geprueft werden und nicht durch Keep-Alive-Reuse verfälscht sind.
- Wenn Secure-Channel-Fehler weiter auftreten, zuerst Zertifikatsvertrauen, DNS, Proxy/TLS-Inspection und die TLS-/Schannel-Richtlinien des Systems pruefen, bevor maschinenweite Registry-/GPO-Overrides erwogen werden.
- App-Passwort gültig?
- Talk installiert?
- Password Policy App optional: bei fehlender App wird lokal generiert (Fallback)



