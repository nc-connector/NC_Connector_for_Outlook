# ADMIN.de.md — NC Connector for Outlook

Dieses Dokument beschreibt Installation, Rollout und Betrieb des **NC Connector for Outlook** (Outlook classic COM Add-in).

## Inhalt

- [Installation (MSI)](#installation-msi)
- [Updates / Upgrade-Verhalten](#updates--upgrade-verhalten)
- [Dateien & Registry](#dateien--registry)
- [Einstellungen (Profil-XML)](#einstellungen-profil-xml)
- [Verwaltete Nextcloud-URL (Registry/GPO)](#verwaltete-nextcloud-url-registrygpo)
- [Compose-Freigabe-Lifecycle (3.1.0)](#compose-freigabe-lifecycle-310)
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
msiexec /i "NCConnectorForOutlook-<version>.msi" /qn /norestart
```

Danach Outlook starten. Im Ribbon erscheint der Tab **NC Connector** (Kalender/Termin + E-Mail).

## Updates / Upgrade-Verhalten

- Updates/Reinstallationen erfolgen durch Installation eines MSI-Pakets ueber die bestehende Installation (gleiche, aeltere oder neuere Version).
- Die MSI ist als **Major Upgrade** konfiguriert (UpgradeCode bleibt stabil), damit eine vorhandene Installation automatisch ersetzt wird.
- Benutzer-Einstellungen bleiben erhalten, da sie im Benutzerprofil gespeichert sind.
- Outlook fragt einmal pro Tag `nc-connector.de` nach Release-Metadaten. Downloads öffnen direkt die GitHub-Release-Dateien; die MSI läuft nicht über die Homepage.
- Das Start-Popup ist Opt-in (`Einstellungen -> Erweitert -> Über neue Versionen informieren`). Der Erweitert-Tab zeigt die gecachte aktuelle Version, den Download-Link und die Änderungsübersicht trotzdem an.
- Die Update-Anfrage sendet Produkt, installierte Version, Kanal und einen täglich wechselnden anonymen Client-Hash. Nextcloud-URLs, E-Mail-Adressen, Benutzernamen, Passwörter, Lizenzschlüssel oder Mandantendaten werden nicht gesendet.

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
  <UpdateNotifyEnabled>false</UpdateNotifyEnabled>
  <UpdateInstallId>LOCAL_RANDOM_ID</UpdateInstallId>
  <UpdateLastCheckedAtUtc>2026-06-12T08:00:00.0000000Z</UpdateLastCheckedAtUtc>
  <UpdateLatestVersion>3.1.1</UpdateLatestVersion>
  <UpdateReleaseUrl>https://github.com/nc-connector/NC_Connector_for_Outlook/releases/tag/v3.1.1</UpdateReleaseUrl>
  <UpdateDownloadUrl>https://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.1.1/NCConnectorForOutlook-3.1.1.msi</UpdateDownloadUrl>
  <FileLinkBasePath>NC Connector</FileLinkBasePath>
  <SharingAttachmentLinkTarget>zip_download</SharingAttachmentLinkTarget>
  <TalkDeleteRoomOnEventDelete>false</TalkDeleteRoomOnEventDelete>
  <EmailSignatureOnCompose>true</EmailSignatureOnCompose>
  <EmailSignatureOnReply>true</EmailSignatureOnReply>
  <EmailSignatureOnForward>true</EmailSignatureOnForward>
</Settings>
```

### Rollout / Pre-Seed

Da die Profil-XML im Benutzerprofil liegt, gibt es mehrere typische Wege:

- **Login Script / Intune/SCCM**: vorbereitete `settings_<OutlookProfile>.xml` nach `%LOCALAPPDATA%\NC4OL\` kopieren (nur falls noch nicht vorhanden).
- **Group Policy Preferences**: Datei in das Benutzerprofil verteilen.

Empfehlung:

- Nur Base-URL und Defaults pre-seeden.
- Credentials entweder ueber Login Flow v2 oder per Benutzer setzen lassen (empfohlen fuer DPAPI-Kompatibilitaet).

## Verwaltete Nextcloud-URL (Registry/GPO)

Outlook kann die Nextcloud-URL aus Windows-Policy-Registry-Keys lesen:

- `HKLM\Software\Policies\NC Connector`
- `HKCU\Software\Policies\NC Connector`

Werte:

- `NextcloudUrl` (`REG_SZ`): vollstaendige Nextcloud-URL, zum Beispiel `https://cloud.example.com`
- `NextcloudUrlLocked` (`REG_DWORD` oder String, optional): `1` / `true` sperrt das URL-Feld in den Einstellungen

Prioritaet:

- `HKLM` gewinnt vor `HKCU`
- 64-bit- und 32-bit-Registry-Views werden gelesen
- wenn der Wert nicht gesperrt ist, fuellt er nur leere Profile vor
- wenn der Wert gesperrt ist, nutzt Outlook immer die Registry-URL und deaktiviert das URL-Feld

Credentials bleiben benutzerspezifisch. Benutzer holen ein App-Passwort ueber den Login-Flow oder geben es manuell ein.

## Talk-Raum-Loeschschutz

Das Loeschen eines gespeicherten Outlook-Termins stellt die entfernte Talk-Raumloeschung nur an, wenn `TalkDeleteRoomOnEventDelete` lokal aktiviert oder per Backend-Policy `talk_delete_room_on_event_delete` gesperrt/aktiviert ist. Ausserdem muss der Termin NC-Connector-Metadaten (`X-NCTALK-TOKEN`) tragen. Generische Talk-Links in `Location` oder URL-Feldern werden ignoriert.

Der Cleanup fuer neu erzeugte Termine, die vor dem Speichern verworfen werden, bleibt davon unberuehrt und loescht den gerade erzeugten Raum weiterhin best effort.

## Zentrale E-Mail-Signatur

Wenn das optionale NC-Connector-Backend installiert ist, kann Outlook eine zentrale HTML-Signatur aus der Backend-Policy verwenden. Das ist bewusst an den Seat-Benutzer gekoppelt:

- der Backend-Endpunkt muss erreichbar sein
- dem aktuellen Benutzer muss ein aktiver Seat zugewiesen sein
- der Backend-Status muss `policy.email_signature` und `policy_editable.email_signature` enthalten
- der effektive Wert von `email_signature_on_compose` muss `true` sein
- `policy.email_signature.email_signature_template` muss HTML enthalten
- `policy.email_signature.user_email` muss die E-Mail-Adresse des Nextcloud-Benutzers enthalten
- die effektive Outlook-Absenderidentitaet muss genau zu dieser E-Mail-Adresse passen. Wenn Outlook einen `SentOnBehalfOfName`-/Von-Override fuer ein Shared Mailbox- oder delegiertes Exchange-Konto verwendet, muss genau dieser Override auf dieselbe SMTP-Adresse aufloesbar sein; andernfalls wird keine Backend-Signatur eingefuegt.

Die Backend-Werte fuer Compose, Antwort und Weiterleitung sind Vorgaben, solange der passende Wert `policy_editable.email_signature.<key>` auf `true` steht. Ein Benutzer darf deshalb eine lokal gespeicherte Option aktivieren, obwohl ihre editierbare Backend-Vorgabe `false` ist. Markiert das Backend einen Wert als nicht editierbar, gewinnt immer der Backend-Wert; ein gesperrtes `false` bleibt deaktiviert.

Wenn ein aelteres Backend Freigabe-/Talk-Policy liefert, aber noch keine `policy.email_signature`-Domain kennt, bleiben nur zentrale Signaturen deaktiviert. Die Settings zeigen dann einen Backend-Update-Hinweis; Freigabe und Talk bleiben von dieser Signatur-Domain getrennt.

Neue Mail, Antwort und Weiterleitung verwenden ihre jeweilige effektive Option. Outlook nutzt zuerst seine Operationsmetadaten, um Antwort und Weiterleitung zu unterscheiden. Bleibt eine Inline-Response unklar und unterscheiden sich beide Optionen, raet NC Connector nicht: Die Hintergrundverarbeitung laesst die Mail unveraendert und die finale Send-Pruefung fordert einen erneuten Versuch, statt mit unsicherem Signaturzustand zu senden. Sind beide Optionen gleich, gilt dieser gemeinsame Wert.

HTML, Plain Text und RTF verwenden Outlook WordEditor sowohl im normalen Inspector als auch in Inline-Antworten. Plain Text bleibt Plain Text, RTF bleibt RTF, und NC Connector schreibt weder `MailItem.Body` noch `MailItem.HTMLBody` neu. Beim Auskoppeln einer Inline-Antwort wechselt die Verarbeitung auf den neuen Inspector.

Das Add-in ersetzt nur Outlooks exakten Slot `_MailAutoSig` oder sein eigenes Bookmark `NcConnectorSignature`. Fehlen beide, wird die Signatur einer neuen Mail hinter allen vom Benutzer geschriebenen Text gesetzt. Bei Antwort/Weiterleitung landet sie hinter dem eigenen Antworttext und vor Outlooks Bookmark `_MailOriginal` oder einem ueber Word-Absatzrahmen erkannten Zitat-Trenner. Aktuelle Cursorposition, Body-Anfang und lokalisierte Header wie `From:` oder `Von:` sind keine Fallbacks. Fehlt eine sichere Zitatgrenze, bleiben eigener und zitierter Inhalt unveraendert.

Jede eingefuegte Backend-Signatur erhaelt in allen drei Formaten das Word-Bookmark `NcConnectorSignature`. Absenderwechsel, Formatwechsel und spaetere Updates bleiben dadurch auf die exakt verwaltete Range begrenzt. Der Wechsel von einem nicht passenden zum passenden Absender setzt die Signatur einer neuen Mail deshalb unter bereits geschriebenen Text statt darueber. Der Wechsel von der passenden Identitaet entfernt nur die verwaltete Bookmark-Range.

Das Backend-HTML laeuft durch den fail-closed Template-Sanitizer. Die formatierte Einfuegung wird vorbereitet und mit einem Bookmark versehen, bevor die vorige Signatur entfernt wird; ein Fehler loest keinen Body-weiten Ersatz aus. Die Word-Auswahl wird ueber ein temporaeres Bookmark wiederhergestellt, und Inspector sowie Inline-Response verwenden dieselben Abstandsregeln.

Outlook prueft Policy, Absender, Format, Compose-Typ und Bookmark vor Send ein letztes Mal synchron. Bei vollstaendigen Backend-Verbindungseinstellungen bricht es den Versand ab, wenn kein erfolgreicher Policy-Snapshot vorliegt oder ein erforderlicher Apply-/Clear-Vorgang nicht sicher abgeschlossen werden kann. Die Mail bleibt fuer Korrektur und einen neuen Sendeversuch offen. Eine unvollstaendige Einrichtung erzeugt keine Signaturpflicht und versucht nur, ein exaktes NC-Connector-Bookmark best effort zu entfernen. Eine nicht unterstuetzte Signatur-Domain deaktiviert ebenfalls die Einfuegung; bei ansonsten vollstaendiger Backend-Einrichtung verweigert die finale Pruefung den Versand aber weiterhin, wenn eine vorhandene verwaltete Range nicht sicher abgeglichen werden kann.

Ist die Policy inaktiv oder unvollstaendig oder passt der Absender nicht, entfernt NC Connector keine Outlook- oder Drittanbieter-Signatur dieser Identitaet. Ist die Compose-Policy fuer den passenden Absender aktiv, aber Antwort oder Weiterleitung deaktiviert, darf Outlook den exakt erkannten Outlook-Signaturplatz dieser Response leeren, ohne die Backend-Signatur einzufuegen.

## Compose-Freigabe-Lifecycle (3.1.0)

### Attachment-Automatisierung und Cleanup-Regeln
- Der Button `Nextcloud Freigabe hinzufuegen` ist auch in Outlook-Inline-Antworten/-Weiterleitungen im Tab Nachricht verfuegbar und nutzt denselben Wizard-Pfad wie Mail-Compose-Inspectoren.
- Inline-Antworten/-Weiterleitungen schreiben den Freigabeblock ueber Outlooks aktiven Inline-WordEditor. HTML/RTF-Mails behalten zwei leere Zeilen ueber dem Freigabeblock; Plain-Text-Mails bleiben Plain Text und verwenden den gerahmten `#`-Block.
- Im Compose-Attachment-Modus werden serverseitige Artefakte direkt nach Share-Erstellung fuer Cleanup-Tracking registriert.
- Unter `Einstellungen -> Freigaben -> Anhaenge` wird das Linkziel der Anhangsautomatisierung gewaehlt: `ZIP-Download` oder `Nextcloud-Freigabeseite`. Fehlt lokal und im Backend ein gueltiger Wert, ist `ZIP-Download` der Standard. Manuell erstellte Freigaben bleiben davon unberuehrt.
- Das Backend verwendet `policy.share.attachment_link_target` mit `zip_download` oder `share_page`; `policy_editable.share.attachment_link_target=false` sperrt die Einstellung. Ein editierbarer Backend-Wert dient nur als Vorgabe, solange noch kein lokaler Wert gespeichert wurde.
- Der Attachment-Modus bleibt bei beiden Linkzielen schreibgeschuetzt und behaelt seine bisherigen Cleanup-Regeln. Nur URL, `{LINK_INTRO}` und `{LINK_LABEL}` richten sich nach dem Linkziel.
- Im ZIP-Modus muss die oeffentliche absolute HTTP(S)-Freigabe-URL auf `/s/<token>` enden. Kann Outlook daraus nicht sicher `<Freigabe-URL>/download` bilden, wird die Einfuegung mit sichtbarer Fehlermeldung abgebrochen; die Original-URL wird nie mit ZIP-Text eingefuegt.
- Cleanup wird erst nach bestaetigtem erfolgreichem Hauptversand wieder entfernt.
- Wird das Compose-Fenster ohne erfolgreichen Versand geschlossen, loescht das Add-in die erzeugten Share-Ordner-Artefakte serverseitig (best effort, mit Grace-Timer fuer Send/Close-Race).
- Die Anhangsautomatisierung wertet neue Dateien sowohl pre-add (`BeforeAttachmentAdd`) als auch post-add aus; kann pre-add ein lokaler Dateipfad aufgeloest werden, kann der NC-Flow den Host-Add best effort vor der normalen Outlook-Post-Add-Verarbeitung abbrechen.
- In Microsoft-365-/Exchange-Umgebungen mit serverseitigen Nachrichtengroessenlimits kann Outlook grosse Anhaenge bereits vor den Add-in-Events blockieren; in diesen Faellen kann die Automatisierung technisch nicht greifen und der Benutzer soll stattdessen den Button `Nextcloud Freigabe hinzufuegen` verwenden.
- Im Datei-Schritt des Sharing-Wizards koennen Dateien und Ordner per Explorer-Drag-and-drop im gesamten Schrittbereich (Queue + Aktionsbereich) hinzugefuegt werden, nicht nur ueber die Add-Buttons.
- Datei-Uploads groesser als 20 MB nutzen Nextcloud Chunked Upload v2. Damit vermeiden wir lange einzelne WebDAV-`PUT`-Requests durch Proxies oder Webserver, die sehr grosse Request-Bodies ablehnen.
- Manuelle Freigaben verwenden immer die Nextcloud-Freigabeseite. Im Attachment-Modus entscheidet das konfigurierte Linkziel zwischen der Freigabeseite und `/s/<token>/download` als ZIP-Download.
- Custom-Share-Templates koennen `{LINK_INTRO}` und `{LINK_LABEL}` verwenden. Outlook fuellt beide Werte passend zum Modus; bestehende Templates ohne diese Variablen bleiben unveraendert nutzbar.
- Aktuelle Clients bevorzugen das versionierte Share-Template des Backends und fallen bei aelteren Backend-Versionen automatisch auf das bisherige Template-Feld zurueck. Eine Migration durch den Administrator ist nicht erforderlich.

### Separater Passwort-Follow-up-Versand
- Ist `Passwort separat senden` aktiv, enthaelt der Haupt-HTML-Block kein Inline-Passwort.
- Der Passwort-Follow-up-Versand startet erst nach bestaetigtem erfolgreichem Hauptversand.
- Der bereits beim Hauptversand erfolgreich gepruefte Policy-Snapshot wird einmal bereinigt und fuer alle Follow-up-Empfaenger sowie einen moeglichen manuellen Fallback-Entwurf wiederverwendet.
- Outlook setzt die Absenderidentitaet des Original-Compose und liest sie wieder aus. Automatischer Versand findet nur statt, wenn der effektive Follow-up-Absender exakt dem beim erfolgreichen Hauptversand erfassten Absender entspricht; andernfalls oeffnet Outlook den manuellen Fallback-Entwurf.
- Passt dieser effektive Absender zusaetzlich zu `policy.email_signature.user_email`, enthaelt der Follow-up die Backend-Signatur. Eine Plain-Text-Hauptmail erzeugt eine Plain-Text-Follow-up-Signatur; HTML/RTF erzeugt einen HTML-Follow-up. Fehlende Policy, Identitaetsabweichung oder Sanitizer-Fehler lassen die Signatur weg, statt ungeprueften Inhalt zu verwenden.
- Wenn die Backend-Policy Nextcloud Secrets auswaehlt, erzeugt Outlook pro finalem Empfaenger einen verschluesselten einmaligen Secrets-Link.
- Wenn die Secrets-Erstellung fehlschlaegt, faellt Outlook auf die bisherige separate Klartext-Passwortmail zurueck und zeigt einen Hinweis.
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
