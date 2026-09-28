# Betriebsanleitung — NC Connector für Outlook

Diese Anleitung richtet sich an Administratoren und Betriebsteams, die **NC Connector für Outlook** bereitstellen und betreiben. Sie beschreibt Voraussetzungen, Rollout, verwaltete Konfiguration, Servervorbereitung, Betriebsprüfungen, Logging und Störungsbehebung.

Quellcode-Aufbau, interne Verarbeitung, Protokollimplementierung, Builds und Entwicklertests sind in [DEVELOPMENT.de.md](DEVELOPMENT.de.md) dokumentiert.

## Inhalt

- [Zweck und Zuständigkeiten](#zweck-und-zuständigkeiten)
- [Voraussetzungen und Rollout-Planung](#voraussetzungen-und-rollout-planung)
- [Bereitstellung und Anwendungslebenszyklus](#bereitstellung-und-anwendungslebenszyklus)
- [Verwaltete Konfiguration und Benutzerdaten](#verwaltete-konfiguration-und-benutzerdaten)
- [Nextcloud-Server vorbereiten](#nextcloud-server-vorbereiten)
- [Optionales NC Connector Backend](#optionales-nc-connector-backend)
- [Funktionsbetrieb](#funktionsbetrieb)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [Sicherheit und Datenverarbeitung](#sicherheit-und-datenverarbeitung)
- [Monitoring und Support](#monitoring-und-support)
- [Runbooks zur Störungsbehebung](#runbooks-zur-störungsbehebung)

## Zweck und Zuständigkeiten

NC Connector ergänzt Outlook classic um diese Funktionen:

- Nextcloud-Datei- und Ordnerfreigaben aus neuen Mails, Antworten und Weiterleitungen
- Umleitung von Anhängen über Nextcloud
- Erstellen und Pflegen von Talk-Räumen aus Outlook-Terminen
- optionale zentral verwaltete E-Mail-Signaturen
- optionales Internet Free/Busy (IFB) über einen lokalen Nextcloud-Proxy

Die übliche Aufgabenteilung ist:

- **Windows-/Outlook-Administration:** MSI-Rollout, Add-in-Registrierung, verwaltete Registry-Werte, Client-Proxy und Zertifikatsvertrauen, IFB-URL-Reservierungen und Client-Logs
- **Nextcloud-Administration:** unterstützte Serverversion, benötigte Apps, öffentliches Routing, Files Sharing, Talk, Systemadressbuch, Speicherplatz und Reverse-Proxy-Limits
- **NC Connector Backend-Administration:** Seat-Zuweisung, zentral verwaltete Vorgaben und Sperren, Vorlagen, Signaturzuweisungen und separate Passwortzustellung

## Voraussetzungen und Rollout-Planung

### Client-Voraussetzungen

- 64-Bit-Windows 10 oder Windows 11
- Outlook classic 2019 oder neuer; das neue Outlook wird nicht unterstützt
- .NET Framework 4.7.2
- Administratorrechte für die MSI-Installation

Die MSI registriert sowohl die 64-Bit- als auch die 32-Bit-Ansicht von Outlook. 32-Bit-Outlook unter 64-Bit-Windows wird unterstützt.

### Nextcloud-Voraussetzungen

- Nextcloud 32 oder neuer für alle Add-in-Funktionen
- Files Sharing für Uploads und öffentliche Freigaben
- Talk für Meeting-Funktionen
- Nextcloud Secrets und NC Connector Backend für Passwortzustellung über einmalige Secret-Links
- das Nextcloud-Systemadressbuch für Benutzersuche, Teilnehmervorgaben und Moderatorauswahl

Das optionale NC Connector Backend wird für lokale Freigaben, Talk oder IFB nicht benötigt. Es ist für zentrale Policies, verwaltete Signaturen und separate Passwortzustellung erforderlich.

Die Nextcloud-App Password Policy ist optional. Ist sie verfügbar, liest NC Connector ihre Passwortvorgaben; andernfalls erzeugt es Passwörter mit seinem lokalen Generator.

### Netzwerk-Voraussetzungen

Clients benötigen HTTPS-Zugriff auf die konfigurierte öffentliche Nextcloud-Basis-URL einschließlich ihrer OCS- und WebDAV-Pfade. Ein öffentlicher Unterpfad wie `/nextcloud` bleibt Bestandteil der URL; `/index.php` darf nicht in die in NC Connector konfigurierte URL aufgenommen werden.

NC Connector lehnt explizite HTTP-URLs, in URLs eingebettete Zugangsdaten sowie Basis-URLs mit Query oder Fragment ab. Dynamische Login- und Passwort-Policy-Endpunkte müssen über HTTPS auf demselben Host, Port und Schema wie die konfigurierte Nextcloud-Basis-URL liegen.

Optionale ausgehende Ziele:

- `https://nc-connector.de/wp-json/ncc/v1/update-check` für tägliche Release-Metadaten
- GitHub-Release-Dateien, wenn ein Administrator oder Benutzer einen Download-Link öffnet

IFB lauscht ausschließlich auf `127.0.0.1` und benötigt keine eingehende Firewall-Regel für andere Computer.

### Checkliste vor dem Rollout

Vor einem breiten Rollout:

1. Vorgesehene Nextcloud-Basis-URL und einen eventuellen Unterpfad dokumentieren.
2. Nextcloud 32 oder neuer und Files Sharing prüfen.
3. Öffentliche Zertifikatskette, DNS, Proxy-Pfad und TLS-Inspection auf einem repräsentativen Arbeitsplatz prüfen.
4. Den [Pretty-URL-Test](#nextcloud-pretty-urls) abschließen.
5. Talk, Secrets, Systemadressbuch und NC Connector Backend nur für die vorgesehenen Funktionen aktivieren.
6. Vor dem Test zentral verwalteter Funktionen einem Pilotbenutzer einen Backend-Seat zuweisen.
7. Aktuelle und vorherige MSI für Rollout und Wiederherstellung bereithalten.
8. Einen Pilot-Abnahmetest für Installation, Verbindung, eine Freigabe, ein Talk-Meeting und jede verwendete optionale Funktion festlegen.

## Bereitstellung und Anwendungslebenszyklus

### Installation

1. Outlook schließen und warten, bis kein `OUTLOOK.EXE`-Prozess mehr läuft.
2. Die MSI mit Administratorrechten installieren.
3. Outlook starten.
4. **NC Connector -> Einstellungen** öffnen, die Nextcloud-Verbindung konfigurieren, den Verbindungstest ausführen und speichern.

Interaktive Installation:

```powershell
msiexec.exe /i "NCConnectorForOutlook-<version>.msi"
```

Unbeaufsichtigte Installation mit MSI-Log:

```powershell
msiexec.exe /i "NCConnectorForOutlook-<version>.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-install.log"
```

Erwartetes Ergebnis:

- **NC Connector** erscheint in den Ribbons von Outlook-Terminen und beim Verfassen von Mails.
- `C:\Program Files\NC4OL\` ist vorhanden.
- Das Add-in ist unter **Datei -> Optionen -> Add-Ins -> COM-Add-Ins** aufgeführt.

Wenn eine Prüfung fehlschlägt, mit [Add-in wird nicht geladen](#add-in-wird-nicht-geladen) fortfahren.

### Pilot-Abnahmetest

Diesen Test mit einem normalen Benutzerkonto durchführen:

1. In den **Einstellungen** die Nextcloud-Verbindung testen und speichern.
2. Eine neue Mail erstellen und eine kleine Nextcloud-Freigabe einfügen.
3. Einen neuen Termin erstellen und einen Talk-Link einfügen, wenn Talk eingesetzt wird.
4. Eine Antwort oder Weiterleitung testen, wenn Freigaben oder verwaltete Signaturen eingesetzt werden.
5. IFB am [lokalen Endpunkt](#internet-freebusy-gateway-ifb) testen, wenn es aktiviert ist.
6. Outlook schließen, erneut öffnen und den Verbindungstest wiederholen.

Add-in-Version, Outlook-Bitness, Nextcloud-Version und das Ergebnis jedes Schritts dokumentieren.

### Upgrade oder Rückkehr auf eine ältere Version

Die MSI ersetzt eine installierte neuere, gleiche oder ältere Version. Benutzerspezifische Einstellungen bleiben im Benutzerprofil erhalten.

Beim ersten Start mit aktiviertem IFB nach einem Upgrade migriert NC Connector nur seine alten Outlook-Werte im Format `/nc-ifb/freebusy/...` auf den aktuellen geschützten Pfad. Vorhandene externe Outlook-Free/Busy-Werte bleiben unverändert. Diese Migration betrifft benutzerspezifische Outlook-Werte; ein Löschen der HTTP-URL-Reservierung mit `netsh` ist dafür nicht erforderlich.

Das Add-in meldet Release-Metadaten, installiert aber keine Updates. **Einstellungen -> Erweitert -> Über neue Versionen informieren** steuert das Popup; die tägliche Metadatenabfrage läuft auch bei deaktiviertem Popup. Freigabe und Verteilung der MSI bleiben Aufgaben der Administration.

1. Outlook schließen.
2. `%LOCALAPPDATA%\NC4OL\settings_*.xml` sichern.
3. Die gewünschte MSI über die vorhandene Installation installieren.
4. Outlook starten und den Pilot-Abnahmetest für die eingesetzten Funktionen wiederholen.

Für die Rückkehr zur vorherigen Add-in-Version denselben Ablauf mit der vorherigen MSI wiederholen. Müssen zusätzlich Einstellungen zurückgespielt werden, Outlook vorher schließen und nur Dateien desselben Windows-Benutzers wiederherstellen. Ein geschütztes App-Passwort kann nach dem Kopieren zu einem anderen Windows-Konto oder Computer unlesbar sein; in diesem Fall neu authentifizieren.

### Deinstallation

1. Solange Outlook noch installiert ist, **NC Connector -> Einstellungen -> IFB** öffnen.
2. IFB deaktivieren und speichern. Dadurch wird der zuvor gespeicherte Outlook-Free/Busy-Pfad wiederhergestellt.
3. Outlook schließen und warten, bis kein `OUTLOOK.EXE`-Prozess mehr läuft.
4. **Windows-Einstellungen -> Apps -> Installierte Apps** verwenden oder:

```powershell
msiexec.exe /x "NCConnectorForOutlook-<version>.msi" /qn /norestart
```

Die MSI entfernt installierte Dateien, die Add-in-Registrierung und die Standard-IFB-URL-Reservierung. Benutzerspezifische Einstellungen, Caches und Logs unter `%LOCALAPPDATA%\NC4OL\` bleiben erhalten, damit eine Neuinstallation die Benutzerkonfiguration nicht löscht.

Dieses Profilverzeichnis erst löschen, wenn Einstellungen und Logs nicht mehr benötigt werden. Eine manuell für einen eigenen IFB-Port erstellte URL-Reservierung gehört nicht zur MSI; sie muss wie unter [Eigener IFB-Port](#eigener-ifb-port) beschrieben separat entfernt werden.

### Installationspfade und Registrierungsprüfung

Standard-Installationsverzeichnis:

```text
C:\Program Files\NC4OL\
```

Primäre Add-in-Registrierung:

```text
HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

32-Bit-Outlook unter 64-Bit-Windows liest:

```text
HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

`LoadBehavior` sollte `3` sein. Die MSI schreibt außerdem `HKLM\Software\NC4OL\HttpUrl` als Installationsmarker für die Standard-IFB-Reservierung.

## Verwaltete Konfiguration und Benutzerdaten

### Profildaten

Einstellungen werden pro Windows-Benutzer und Outlook-Profil gespeichert:

```text
%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml
```

Wenn Outlook keinen Profilnamen liefert, verwendet das Add-in:

```text
%LOCALAPPDATA%\NC4OL\settings_default.xml
```

Das App-Passwort wird als `AppPasswordProtected` mit dem Windows-Datenschutz für den aktuellen Benutzer gespeichert. Es ist kein portables Zugangsmittel.

Jeder erfolgreiche Speichervorgang schreibt zunächst eine validierte temporäre Datei im selben Verzeichnis und ersetzt danach die Hauptdatei. Die vorherige gültige Datei bleibt als `settings_<OutlookProfile>.xml.bak` erhalten. Ist nur das geschützte Passwort nicht lesbar, behält NC Connector die übrigen Einstellungen, leert das Passwort und blockiert automatische Settings-Schreibvorgänge, bis der Benutzer eine korrigierte Konfiguration ausdrücklich speichert.

Ausstehende Talk-Raum-Löschungen und IFB-Registry-Besitzstände werden getrennt pro Outlook-Profil unter `%LOCALAPPDATA%\NC4OL\` gespeichert. Diese Zustandsdateien verwenden Windows Data Protection und behalten neben der Primärdatei eine Sicherung. Geschützte Zustände nicht in ein anderes Windows-Konto kopieren.

Ältere `settings.ini`-Dateien unter diesen Verzeichnissen werden beim ersten Start migriert und erst nach erfolgreicher Migration entfernt:

```text
%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\
%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\
```

### Sicherung und Wiederherstellung

So wird ein Client-Profil gesichert:

1. Outlook schließen.
2. `settings_*.xml` und `settings_*.xml.bak` aus `%LOCALAPPDATA%\NC4OL\` kopieren.
3. Sollen ausstehende Talk-Raum-Löschungen die Sicherung überleben, zusätzlich `talk-room-lifecycle-*.dat*` und `ifb-registry-state-*.dat*` kopieren. Sie können nur für denselben Windows-Benutzer wiederhergestellt werden.
4. Windows-Benutzer und Outlook-Profilnamen dokumentieren.

So wird es wiederhergestellt:

1. Outlook schließen.
2. Die passende Haupt- und Sicherungsdatei für denselben Windows-Benutzer zurückspielen.
3. Outlook starten und den Verbindungstest ausführen.
4. Bei fehlgeschlagener Authentifizierung den Nextcloud-Login-Flow erneut verwenden.

Logs und der IFB-Adressbuch-Cache sind Betriebsdaten und für die Wiederherstellung der Konfiguration nicht erforderlich. Ohne die Talk-Lifecycle-Dateien gehen ausstehende Wiederholungen von Talk-Raum-Löschungen verloren.

### Rollout und Vorbelegung

Für die Nextcloud-URL die nachfolgende verwaltete Registry-Policy verwenden. Zentrale Funktionsvorgaben und Sperren über das optionale Backend verteilen.

Muss eine Profil-XML vorab verteilt werden:

- nur bereitstellen, wenn noch keine Profildatei vorhanden ist
- als Vorlage eine Datei verwenden, die mit derselben Add-in-Version erstellt wurde
- nur stabile, organisationsweit benötigte Vorgaben aufnehmen
- `AppPasswordProtected` vor der Verteilung entfernen
- jeden Benutzer über den Nextcloud-Login-Flow authentifizieren lassen
- das Ergebnis mit allen beim Rollout vorkommenden Outlook-Profilnamen prüfen

Ein geschütztes App-Passwort niemals zwischen Benutzern oder Computern kopieren.

### Verwaltete Nextcloud-URL

Unterstützte Policy-Pfade:

```text
HKLM\Software\Policies\NC Connector
HKCU\Software\Policies\NC Connector
```

Werte:

- `NextcloudUrl` (`REG_SZ`): vollständige öffentliche Nextcloud-URL
- `NextcloudUrlLocked` (`REG_DWORD` oder String, optional): `1` oder `true` sperrt das Feld

Beispiel für eine Computer-Policy:

```powershell
$policyPath = "HKLM:\Software\Policies\NC Connector"
New-Item -Path $policyPath -Force | Out-Null
New-ItemProperty -Path $policyPath -Name NextcloudUrl -PropertyType String -Value "https://cloud.example.com" -Force | Out-Null
New-ItemProperty -Path $policyPath -Name NextcloudUrlLocked -PropertyType DWord -Value 1 -Force | Out-Null
```

Priorität und Ergebnis:

- `HKLM` hat Vorrang vor `HKCU`.
- Die 64-Bit- und 32-Bit-Registry-Ansicht werden gelesen.
- Ein nicht gesperrter Wert füllt nur ein leeres Profil vor.
- Ein gesperrter Wert wird für jedes Profil verwendet und deaktiviert das URL-Feld.
- Zugangsdaten bleiben benutzerspezifisch.

Nach dem Verteilen der Policy Outlook neu starten und URL sowie Sperrstatus unter **NC Connector -> Einstellungen** prüfen.

## Nextcloud-Server vorbereiten

### Basisprüfungen

Für einen Pilotbenutzer:

1. **NC Connector -> Einstellungen** öffnen.
2. Die öffentliche Nextcloud-URL eintragen.
3. Über den Login-Flow oder mit einem App-Passwort authentifizieren.
4. Den Verbindungstest ausführen.

Der Test muss eine unterstützte Nextcloud-Version melden. Ein älterer Server oder eine Antwort ohne auswertbare Version wird abgelehnt.

Optionale Funktionen getrennt prüfen:

- eine öffentliche Freigabe für Files Sharing erstellen
- einen Talk-Raum für Talk erstellen
- nach Aktivierung des Systemadressbuchs nach einem Benutzer suchen
- backendverwaltete Einstellungen mit einem zugewiesenen Seat öffnen

### Nextcloud Pretty URLs

Pretty URLs sind eine serverweite Nextcloud-Routing-Voraussetzung. Sie betreffen Authentifizierung, Dateien, Apps, Talk und weitere Routen; sie sind keine reine Talk- oder Add-in-Einstellung.

NC Connector erstellt eine öffentliche Talk-URL in dieser Form:

```text
https://cloud.example.com/call/<TOKEN>
```

Funktioniert der Raum nur als `https://cloud.example.com/index.php/call/<TOKEN>`, routet der Webserver oder Reverse Proxy Pretty URLs nicht korrekt. Die öffentliche Route muss korrigiert werden; `/index.php` darf nicht zur in NC Connector konfigurierten URL hinzugefügt werden.

Bei einer Nextcloud-Installation unterhalb von `/nextcloud` lautet die erwartete URL:

```text
https://cloud.example.com/nextcloud/call/<TOKEN>
```

#### Kurztest

Im Webroot öffnen:

```text
https://cloud.example.com/index.php/login
https://cloud.example.com/login
```

Unterhalb von `/nextcloud` öffnen:

```text
https://cloud.example.com/nextcloud/index.php/login
https://cloud.example.com/nextcloud/login
```

Die URL ohne `/index.php` muss Nextcloud erreichen oder auf die Login-Seite weiterleiten. Ein Webserver-404 bedeutet, dass das Rewrite nicht aktiv ist.

#### Nginx

Die vollständige Nextcloud-Nginx-Konfiguration als Grundlage verwenden. Die folgenden Ausschnitte gehören in die passenden vorhandenen `server`- und PHP/FastCGI-Blöcke; keine doppelten Locations anlegen.

Im Webroot:

```nginx
location / {
    try_files $uri $uri/ /index.php$request_uri;
}
```

In der PHP/FastCGI-Location:

```nginx
fastcgi_param front_controller_active true;
```

Unterhalb von `/nextcloud` muss der Fallback diesen Pfad beibehalten:

```nginx
location /nextcloud {
    try_files $uri $uri/ /nextcloud/index.php$request_uri;
}
```

Prüfen und neu laden:

```bash
sudo nginx -t
sudo systemctl reload nginx
```

#### Apache

Apache muss `mod_rewrite` und `mod_env` laden. Der HTTP-Benutzer muss Nextclouds `.htaccess` schreiben können und der passende `<Directory>`-Block muss diese Regeln mit `AllowOverride All` erlauben.

Unter Debian oder Ubuntu:

```bash
sudo a2enmod rewrite env
sudo systemctl reload apache2
```

Für Nextcloud im Webroot in `config/config.php` setzen:

```php
'overwrite.cli.url' => 'https://cloud.example.com/',
'htaccess.RewriteBase' => '/',
```

Für Nextcloud unterhalb von `/nextcloud`:

```php
'overwrite.cli.url' => 'https://cloud.example.com/nextcloud',
'htaccess.RewriteBase' => '/nextcloud',
```

Hinter einem Reverse Proxy bezieht sich `htaccess.RewriteBase` auf den Backend-Apache-`DocumentRoot` nach der Proxy-Zuordnung. Entfernt der Proxy `/nextcloud` vor der Weiterleitung, `/` verwenden.

`.htaccess` mit dem tatsächlichen Installationspfad neu erzeugen:

```bash
cd /var/www/nextcloud
sudo -E -u www-data php occ maintenance:update:htaccess
sudo systemctl reload apache2
```

Erst nach Prüfung der Module, `AllowOverride`, Rewrite Base und neu erzeugten `.htaccess` kann der folgende Nextcloud-Ausweichweg getestet werden:

```php
'htaccess.IgnoreFrontController' => true,
```

`maintenance:update:htaccess` erneut ausführen und Apache neu laden.

Den Login-Test wiederholen und einen neu erstellten `/call/<TOKEN>`-Link von einem Client außerhalb des Servernetzes öffnen.

Offizielle Nextcloud-Referenzen:

- [Nginx-Konfiguration](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html)
- [Apache-Installation und Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls)
- [`maintenance:update:htaccess`](https://docs.nextcloud.com/server/32/admin_manual/occ_command.html#maintenance-commands)

### Systemadressbuch

Das Systemadressbuch wird benötigt für:

- Moderatorauswahl im Talk-Assistenten
- die Vorgabe **Benutzer hinzufügen**
- die Vorgabe **Gäste hinzufügen**
- die IFB-Adressauflösung

Unter **Nextcloud-Verwaltungseinstellungen -> Groupware -> Systemadressbuch** aktivieren. Zusätzlich unter **Verwaltungseinstellungen -> Teilen** die Benutzernamen-Autovervollständigung beziehungsweise den Systemadressbuchzugriff aktivieren.

Zeigt die Verwaltungsseite das Systemadressbuch als aktiv an, Clients können es aber weiterhin nicht verwenden:

```bash
sudo -E -u www-data php occ config:app:delete dav system_addressbook_exposed
sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"
sudo -E -u www-data php occ dav:sync-system-addressbook
```

Danach das erzeugte Adressbuch für einen Testbenutzer prüfen:

```text
https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system?export
```

Erwartetes Ergebnis: Benutzersuche und Moderatorfelder werden nach einer erneuten Outlook-Verbindung verfügbar. Ein vollständiger, gültiger Adressbuchexport funktioniert auch dann, wenn ein Reverse-Proxy fälschlich HTTP 404 oder einen falschen Inhaltstyp zurückgibt. Bei anderen HTTP-Fehlern wie 401 oder 403 muss der Zugriff auf den Endpunkt weiterhin korrigiert werden. Eine leere HTTP-404-Antwort ist kein Adressbuch.

Schlägt ein Abruf fehl oder ist die Antwort beschädigt, zeigt Outlook einen Adressbuchfehler an. Die betroffenen Felder bleiben bis zu einem erfolgreichen erneuten Abruf nicht verfügbar. Vorherige Kontakte bleiben im Cache erhalten, werden aber nicht als erfolgreicher neuer Abruf ausgegeben. Die Teilnehmerzuordnung stoppt, statt nicht aufgelöste interne Benutzer als Gäste einzuladen. Benutzer mit gültiger Nextcloud-UID ohne E-Mail-Adresse bleiben in Benutzersuche und Moderatorauswahl verfügbar; die Zuordnung eines E-Mail-Empfängers benötigt weiterhin eine E-Mail-Adresse.

Offizielle Nextcloud-Referenzen:

- [Systemadressbuch](https://docs.nextcloud.com/server/32/admin_manual/groupware/contacts.html#system-address-book)
- [`dav:sync-system-addressbook`](https://docs.nextcloud.com/server/32/admin_manual/occ_command.html#sync-system-address-book)

## Optionales NC Connector Backend

### Voraussetzungen und Betriebszustände

Backendverwaltete Funktionen benötigen:

- die installierte und aktivierte App `ncc_backend_4mc`
- Client-Zugriff auf `/apps/ncc_backend_4mc/api/v1/status`
- einen dem aktuellen Nextcloud-Benutzer zugewiesenen aktiven Seat
- eine von der installierten Backend-Version unterstützte Policy-Domain

Beobachtbares Verhalten je Zustand:

- **Keine Backend-Konfiguration:** Freigaben, Talk und IFB verwenden lokale Einstellungen. Zentrale Signaturen und separate Passwortzustellung sind nicht verfügbar.
- **Erreichbares Backend mit aktivem Seat:** Gespeicherte Outlook-Werte gelten für editierbare Felder. Gesperrte Backend-Werte überschreiben den lokalen Wert und können in Outlook nicht geändert werden.
- **Erreichbares Backend ohne nutzbaren Seat:** Freigaben und Talk verwenden lokale Einstellungen; Outlook zeigt den Seat- oder Lizenzstatus. Zentrale Signaturen und separate Passwortzustellung sind nicht verfügbar.
- **Backend vorübergehend nicht erreichbar:** Freigaben und Talk verwenden gespeicherte lokale Einstellungen. Eine passende Mail mit verpflichtender zentraler Signatur kann geöffnet und ungesendet bleiben, bis die Signatur-Policy wieder geprüft werden kann.
- **Backend ohne Signatur-Domain:** Freigabe- und Talk-Policies funktionieren weiter. Zentrale Signaturen bleiben deaktiviert und Outlook zeigt einen Update-Hinweis.

### Lizenzhinweise in Outlook

Einstellungen, Freigabe-Wizard und Talk-Dialog verwenden denselben Statushinweis. Eine aktive Lizenz erzeugt keine Lizenzwarnung. Während der Nachfrist zeigt ein gelber Hinweis, wie lange die Pro-Funktionen verfügbar bleiben; für Benutzer mit aktivem zugewiesenem Seat werden sie dadurch nicht deaktiviert.

Nach Ende der Nachfrist unterscheidet Outlook eine abgelaufene Lizenz von einer inaktiven oder ungültigen Lizenz, einem Aktivierungsproblem und einer überschrittenen Offline-Prüffrist. Ein tatsächlich pausierter Seat wird gesondert gemeldet. Benutzer ohne Seat sehen weiterhin den Hinweis auf die fehlende Zuweisung. Die grundlegenden Freigabe- und Talk-Funktionen bleiben mit lokalen Einstellungen nutzbar.

Vollständige Nextcloud-Administratoren sehen Lizenzhinweise auch ohne eigenen Seat und können **Lizenz im Backend verwalten** öffnen. Der Link führt zur NC-Connector-Verwaltung der konfigurierten Nextcloud. Andere Benutzer werden an ihren Nextcloud-Administrator verwiesen. Bei älteren Backends ohne diese zusätzlichen Angaben erscheint ein allgemeiner Hinweis statt einer vermuteten Ablaufursache oder eines Verwaltungslinks.

Ein Administrator ohne eigenen Seat sieht zusätzlich, dass Freigaben und Talk mit lokalen Einstellungen nutzbar bleiben. Der Tooltip an einer gesperrten Seat-Funktion nennt die fehlende Zuweisung, auch wenn der Banner zusätzlich eine Nachfrist oder ein Synchronisationsproblem meldet. Administratorrechte schalten keine persönlichen Funktionen frei.

Ein aktiv zugeteilter Community-Seat besitzt dieselben vorhandenen Funktionen wie ein aktiv zugeteilter Pro-Seat. Bei überschrittener Kapazität verlieren nur die überzähligen pausierten Seats ihre Seat-Funktionen; die übrigen aktiven Benutzer behalten Richtlinien und Funktionen.

Eine fehlgeschlagene Lizenzsynchronisierung wird für sich genommen nicht als ungültige Lizenz bezeichnet: Ursache können ein Verbindungsproblem, eine unbrauchbare Serverantwort oder nicht verfügbare lokale Aktivierungsdaten sein. Wenn das Backend die Angaben liefert, nennt der Hinweis die letzte erfolgreiche Synchronisierung und die Offline-Prüffrist. Scheitert dagegen der Abruf des Nextcloud-Backend-Status, erscheint ein eigener Verbindungshinweis. Tooltips deaktivierter Funktionen verwenden den passenden Grund; ein zusätzliches Lizenz-Popup bei jeder Aktion gibt es nicht.

Die Lizenzaktivierung bleibt Teil der normalen Backend-Synchronisierung. Outlook aktiviert keine Lizenzen und kontaktiert den Lizenzserver nicht direkt. Nach Korrektur einer Lizenz oder Seat-Zuweisung den betroffenen Dialog erneut öffnen oder die Verbindung in den Einstellungen aktualisieren, um den aktuellen Backend-Status zu laden.

### Policy-Rollout

Das Backend kann verwalten:

- Talk-Vorgaben und Raumlöschung bei gespeicherten Terminen
- Freigabe-Vorgaben, Passwortregeln und Linkziel für Anhänge
- Vorlagen für Freigaben, Passwortmails und Talk-Einladungen
- separate Passwortzustellung und optionale Secret-Links
- zentrale Signaturzuweisung sowie getrennte Schalter für neue Mail, Antwort und Weiterleitung

Werte editierbar lassen, wenn Benutzer ihre gespeicherte Outlook-Einstellung behalten oder ändern dürfen. Nur Einstellungen sperren, die Benutzer nicht ändern dürfen. Vor dem breiten Rollout sowohl einen Benutzer mit aktivem Seat als auch einen Benutzer ohne Seat testen.

### Vorlagen erstellen

Für eigene Freigabevorlagen:

- `{LINK_INTRO}` und `{LINK_LABEL}` verwenden, wenn der Text dem wirksamen Linkziel folgen soll
- manuelle Freigaben verwenden immer den Text für die Freigabeseite
- die Anhangsautomatisierung kann Text für ZIP-Download oder Freigabeseite verwenden
- Vorlagen ohne diese Variablen behalten ihren vorhandenen Text
- feste Beschriftung und Platzhalter eines optionalen Felds, etwa `{PASSWORD}`, im selben `tr`, `p`, `li` oder `div` platzieren; Outlook entfernt diesen vollständigen Block, wenn der Wert leer ist
- absolute `https://`-Links verwenden

Für HTML in Talk-Terminen:

- Tabellen (`table`, `tbody`, `tr`, `td`) für das Layout verwenden
- einfache Inline-Styles verwenden
- `flex`, `grid`, `border-radius`, `overflow`, `object-fit` und `user-select` vermeiden
- vollständige `https://`-Links verwenden

Nicht unterstütztes oder unsicheres HTML kann entfernt oder die Vorlage abgelehnt werden. Vorlagen in HTML-/RTF-Outlook-Terminen und mit jedem unterstützten Office-Design testen.

### Abnahmetest für verwaltete Signaturen

Eine zentrale Signatur wird nur angewendet, wenn die wirksame Outlook-**Von**-Adresse mit der im Backend zugewiesenen E-Mail-Adresse übereinstimmt. Eine Shared Mailbox oder delegierte **Von**-Adresse muss auf dieselbe SMTP-Adresse aufgelöst werden.

Vor dem Rollout diese Matrix testen:

1. Neue HTML-Mail mit passender Identität.
2. Neue Plaintext-Mail mit passender Identität.
3. Antwort und Weiterleitung mit von Anfang an ausgewählter passender Identität.
4. Antwort und Weiterleitung nach Wechsel von einer nicht passenden zur passenden Identität.
5. Eine nicht passende Identität.
6. Eine Shared oder delegierte Mailbox, falls eingesetzt.
7. Eine separate Passwort-Follow-up-Mail.

Erwartetes Ergebnis:

- bei einer neuen Mail steht die Signatur nach dem vom Benutzer geschriebenen Text
- bei Antworten und Weiterleitungen steht sie oberhalb der zitierten Nachricht
- beim Wechsel von der passenden Identität wird nur die verwaltete Signatur entfernt
- eine nicht passende Identität und ihre eigene Signatur bleiben unverändert
- getrennte Policies für neue Mail, Antwort und Weiterleitung werden befolgt
- ein gesperrter Backend-Wert kann in Outlook nicht geändert werden
- die Passwort-Follow-up-Mail erhält die Signatur nur bei ebenfalls passendem wirksamen Absender

Kann eine erforderliche abschließende Signaturprüfung nicht abgeschlossen werden, lässt Outlook die Mail geöffnet, statt sie mit ungeprüftem Signaturzustand zu senden.

Die Meldung unterscheidet eine noch nicht verfügbare Signaturrichtlinie von einer nicht sicher aktualisierbaren Signatur. Ist die Richtlinie nicht verfügbar, die Nextcloud-Verbindung prüfen und das Senden erneut versuchen. Eine zuvor bestätigte Richtlinie desselben Kontos bleibt nach einem fehlgeschlagenen Refresh nutzbar; eine neu empfangene Ablehnung wird wirksam.

## Funktionsbetrieb

### Freigaben und Uploads

Der Freigabe-Assistent akzeptiert lokale Dateien und Ordner sowie vorhandene Inhalte aus der eigenen Nextcloud des konfigurierten Benutzers. Die Warteschlange zeigt den vollständigen Zielpfad, gruppiert beide Quellen und nennt Einträge, Quellenanzahl, Gesamtgröße und Nextcloud-Speicherstatus. Der Picker **Meine Nextcloud** bietet Explorer-artige Schaltflächen für Zurück, Vorwärts, Hoch und Aktualisieren, klickbare Pfadsegmente, Filter, vertraute Dateisymbole, Dateidetails, einen eindeutigen Status bei Mehrfachauswahl, Speicherinformationen und bei Bedarf geladene Vorschauen für Bilder und vom Nextcloud-Server unterstützte Dokumente. Dokumentvorschauen, etwa für PDF- oder Office-Dateien, hängen von den auf dem Server aktivierten Vorschau-Anbietern ab. Diese Quelle funktioniert ohne NC Connector Backend.

Ausgewählte Nextcloud-Inhalte werden innerhalb desselben Kontos in den neuen Freigabeordner kopiert. Das Original bleibt unverändert und wird für die Übertragung nicht nach Outlook heruntergeladen. Für eine Vorschau fordert Outlook zuerst ein größenbegrenztes, von Nextcloud erzeugtes Bild an. Hat der Server für eine unterstützte Bilddatei keine erzeugte Vorschau, kann Outlook vorübergehend das Originalbild bis 5 MiB laden. Andere Originaldateien werden für Vorschauen nicht heruntergeladen. Lokale Inhalte verwenden weiterhin die vom Server und den ausgewählten Dateien unterstützte Uploadmethode. Implementierungsdetails stehen in [DEVELOPMENT.de.md](DEVELOPMENT.de.md#filelink-upload-architektur).

Betriebsgrenzen und Fehlerverhalten:

- symbolische Links und Junctions werden abgelehnt
- eine nach dem ersten Scan veränderte Quelldatei stoppt den Upload
- das konfigurierte Konto benötigt Lesezugriff auf ausgewählte Nextcloud-Inhalte und Schreibzugriff auf den Zielordner
- ein bereits vorhandener Stammordnername stoppt eine manuelle Freigabe; die Anhangsautomatisierung kann einen nummerierten Namen wählen
- HTTP `507` bedeutet zu wenig freien Nextcloud-Speicher
- Proxy-Timeouts und Request-Größenlimits können Uploads beeinträchtigen, obwohl Client und Nextcloud ansonsten funktionieren

Bei einer Störung mit einem großen Ordner die Anzahl ausgewählter Elemente, Gesamtgröße, Uhrzeit, angezeigte Phase, Nextcloud-Speicherstatus, Reverse-Proxy-Limits und `FILELINK`-Logeinträge erfassen.

### Anhangsautomatisierung

Unter **Einstellungen -> Freigabe -> Anhänge** können Administratoren oder Backend-Policy festlegen:

- Anhänge immer über NC Connector senden
- NC Connector oberhalb eines Größenschwellwerts anbieten
- `ZIP-Download` oder `Nextcloud-Freigabeseite` als Linkziel für Anhänge

`ZIP-Download` ist die Vorgabe, wenn weder ein lokaler noch ein Backend-Wert vorhanden ist. Das Linkziel gilt nur für die Anhangsautomatisierung; manuell erstellte Freigaben verlinken immer auf die Nextcloud-Freigabeseite.

Anhangsregeln, die älter als fünf Minuten sind, gelten während der Aktualisierung im Hintergrund weiter. Allein dieses Alter unterbricht den Versand nicht. Wurden noch keine Regeln geladen oder gerade Einstellungen geändert, bittet ein Informationshinweis darum, das Senden in einem Moment erneut zu versuchen. Die Mail bleibt offen und ihre Anhänge bleiben unverändert. Ein eigener Warnhinweis erklärt, wenn die wirksamen Einstellungen tatsächlich eine Freigabe über NC Connector vorschreiben; keiner der beiden Hinweise meldet einen fehlgeschlagenen Upload.

Beide Linkziele bleiben schreibgeschützte Freigaben. Kann aus der öffentlichen Freigabe keine gültige ZIP-Download-URL abgeleitet werden, stoppt das Einfügen mit einem Fehler. NC Connector beschriftet eine normale Freigabeseiten-URL nicht als ZIP-Download.

Outlook oder Exchange kann einen großen Anhang ablehnen, bevor ein Add-in-Ereignis läuft. In diesem Fall müssen Benutzer **Nextcloud-Freigabe einfügen** wählen und die Datei direkt im Freigabe-Assistenten hinzufügen.

### Ungesendete Mail und Freigabebereinigung

Nach dem Einfügen eines Freigabeblocks verfolgt NC Connector die neu erstellte Serverfreigabe, bis Outlook die Nachricht nach dieser Einfügung erfolgreich speichert. Wird das Verfassen-Fenster vor einer erfolgreichen Speicherung verworfen, entfernt NC Connector den exakten Serverordner mit dem beim Erstellen erfassten Kontokontext.

- Manuelles oder automatisches Speichern eines Entwurfs behält die Freigabe. Versand, „Später senden“ und Offline-Postausgang behalten sie ebenfalls.
- Ein abgebrochener Schließvorgang oder das Verschieben einer Inline-Antwort in ein eigenes Fenster entfernt die Freigabe nicht.
- Das spätere Löschen eines bereits von Outlook gespeicherten Entwurfs wird nicht verfolgt. Seine ungenutzte Freigabe muss manuell in Nextcloud entfernt werden.
- Kann der Freigabeblock nicht eingefügt werden, meldet der Wizard einen Fehler und versucht sofort, den neu erzeugten Serverordner mit dem erfassten Kontokontext zu entfernen.
- Eine erzwingende Anhangs-Policy blockiert den Versand, solange ein normaler Anhang in der Nachricht verbleibt, der über NC Connector hätte laufen müssen.
- Die separate Passwortzustellung wird beim Klick auf **Senden** direkt übergeben, wie im folgenden Abschnitt beschrieben.

### Separate Passwortzustellung

Separate Passwortzustellung benötigt NC Connector Backend und einen aktiven Seat.

- Die Hauptmail enthält kein Klartextpasswort.
- Die separate Passwortzustellung ist an die geöffnete Verfassen-Sitzung gebunden, in der die Freigabe erstellt wurde. Der Benutzer muss in genau dieser weiterhin geöffneten Nachricht auf **Senden** klicken. Speichern und AutoSave unterbrechen den Ablauf nicht, solange das Verfassen-Fenster geöffnet bleibt.
- Das Speichern der Hauptmail als Entwurf oder `.oft`-Vorlage mit anschließendem Schließen wird nicht unterstützt. Beim erneuten Öffnen des Entwurfs, einem Outlook-Neustart vor dem ersten Sendeversuch oder einer neuen Nachricht aus dieser Vorlage bleibt der sichtbare Freigabeblock erhalten, der Zustand für die Passwortzustellung jedoch nicht. Es wird keine Passwort-Follow-up-Mail erstellt; vor dem Versand muss in der endgültigen Nachricht eine neue Freigabe erstellt werden.
- Beim Klick auf **Senden** wird der Passwort-Follow-up vollständig aufgebaut und sofort über dasselbe wirksame Outlook-Konto übergeben. NC Connector speichert keinen zwischenzeitlichen Passwortentwurf und wertet weder Entwürfe, Postausgang noch Gesendet-Ordner aus.
- Diese direkte Grenze wartet nicht auf eine verzögerte oder Offline-Übermittlung der Hauptmail. Der Passwort-Follow-up kann daher bereits übergeben sein, während die Hauptmail noch im Postausgang liegt oder nach dem NC-Connector-Send-Callback von Outlook abgelehnt wird.
- Schlägt die Absenderprüfung oder automatische Übergabe eindeutig fehl, öffnet Outlook eine vollständig vorbereitete Nachricht zum manuellen Senden. Bei einem mehrdeutigen Outlook-Übermittlungsstatus wird kein automatisches Duplikat erzeugt.
- Im Secrets-Modus wird für jeden endgültigen Empfänger ein eigener einmaliger Secret-Link erstellt.
- Gleiche SMTP-Adressen in An, Cc und Bcc erhalten nur einen Secret-Follow-up.
- Schlägt die Secrets-Erstellung fehl, verwendet Outlook die Klartext-Passwort-Follow-up-Mail und zeigt eine Warnung.
- Eine passende Backend-Signatur wird nur eingefügt, wenn auch der Follow-up-Absender mit der zugewiesenen Signaturadresse übereinstimmt.

### Talk-Raum-Lebenszyklus

Das Löschen eines gespeicherten Outlook-Termins entfernt den zugehörigen entfernten Talk-Raum nur, wenn die Einstellung ausdrücklich aktiviert ist und der Termin NC Connector-Raummetadaten enthält. Die Einstellung ist standardmäßig deaktiviert. Ein in Ort oder Nachrichtentext kopierter Talk-Link reicht für eine entfernte Löschung nicht aus.

Das Löschen eines einzelnen Vorkommens oder einer Ausnahme einer Terminserie entfernt den gemeinsamen Raum nicht. Nur ein Nicht-Serientermin oder der Serienmaster kann die Raumlöschung vormerken.

NC Connector reagiert ausschließlich auf Outlook-Ereignisse des jeweiligen Talk-Termins. Das Öffnen oder Auswählen eines Talk-Termins bindet genau diesen Termin an das Löschereignis; das gilt auch für die nach einem Outlook-Neustart wiederhergestellte aktuelle Kalenderauswahl. Damit verwenden das Löschen aus dem geöffneten Termin und aus der Kalenderansicht dieselben Prüfungen und dieselbe Queue. Beim Outlook-Start werden weder Stores oder Kalenderordner aufgezählt noch Kalenderelemente durchsucht.

Vorgemerkte Raumlöschungen werden pro Outlook-Profil gespeichert und nach vorübergehenden Nextcloud-Fehlern oder einem Outlook-Neustart im Hintergrund wiederholt. Die Bereinigung eines neu erstellten Raums aus einem ungespeicherten und verworfenen Termin bleibt aktiv.

Die Moderatorübergabe lehnt den aktuellen Nextcloud-Benutzer ab, wenn die Wizard-Eingabe zur kanonischen Benutzer-ID, zum konfigurierten Login oder zur bekannten primären E-Mail-Adresse passt. Nach erfolgreicher Übergabe an eine andere Person verlässt der ursprüngliche Moderator den Raum weiterhin.

Vor der organisationsweiten Aktivierung der Raumlöschung für gespeicherte Termine:

1. Mit einem Pilotkonto einen Talk-Termin erstellen und speichern.
2. Outlook neu starten und den weiterhin ungeöffneten Termin direkt aus der Kalenderansicht löschen.
3. Prüfen, dass der entfernte Raum gelöscht wird.
4. Den Test mit dem Löschen aus dem geöffneten Termin wiederholen und Wiederherstellung oder Neuerstellung eines Raums für Benutzer dokumentieren.

## Internet Free/Busy Gateway (IFB)

### Zweck und Aktivierung

IFB lässt Outlook Nextcloud-Free/Busy-Daten über einen lokalen HTTP-Endpunkt abfragen.

1. Das Nextcloud-Systemadressbuch prüfen.
2. **NC Connector -> Einstellungen -> IFB** öffnen.
3. IFB aktivieren und Anzahl der Tage, Cache-Dauer und lokalen Port wählen. Vorgaben sind 30 Tage, 24 Cache-Stunden und Port `7777`.
4. Speichern und Outlook neu starten.

Reservierter Listener-Namespace:

```text
http://127.0.0.1:7777/nc-ifb/
```

Die MSI reserviert den Standard-URL-Namespace für authentifizierte Windows-Benutzer. NC Connector ergänzt die Outlook-Free/Busy-URL um ein zufälliges Pfadsegment; Anfragen ohne dieses Segment erhalten `404`. Der geheime Pfad wird intern verwaltet und in dieser Anleitung bewusst nicht angezeigt.

Beim Aktivieren von IFB werden nur benutzerspezifische Outlook-Free/Busy-Werte aktualisiert. Vorhandene Werte und Typen werden getrennt erfasst. Beim Deaktivieren stellt NC Connector einen Wert nur wieder her, solange er noch den von NC Connector geschriebenen Inhalt hat; spätere Änderungen durch Administratoren oder andere Anwendungen bleiben unangetastet. Werte unter `Software\Policies` werden nur auf Konflikte geprüft und nie geschrieben. Der Adressbuch-Cache ist nach Outlook-Profil, vollständiger Nextcloud-Basis-URL einschließlich Unterpfad und konfiguriertem Login getrennt. Der Adressbuchabruf selbst verwendet die kanonische Nextcloud-Benutzer-ID.

Der Listener läuft nur, solange Outlook läuft, IFB aktiviert ist und die gespeicherten Nextcloud-Zugangsdaten vollständig sind.

### Standard-Reservierung prüfen

```powershell
netsh http show urlacl | Select-String -Pattern "127.0.0.1:7777/nc-ifb"
Test-NetConnection 127.0.0.1 -Port 7777
```

Danach in Outlook einen Testtermin erstellen, eine Adresse aus dem Nextcloud-Systemadressbuch hinzufügen und den **Terminplanungs-Assistenten** öffnen. Erwartetes Ergebnis: Die Reservierung ist vorhanden, der TCP-Test ist erfolgreich und Outlook zeigt Free/Busy-Daten. Eine direkte Anfrage an den öffentlichen Pfad `/nc-ifb/freebusy/...` muss `404` liefern.

### Eigener IFB-Port

Gültige konfigurierte Ports reichen von `1024` bis `49151`. Die MSI erstellt nur für Port `7777` eine Reservierung. Für einen anderen Port eine PowerShell mit erhöhten Rechten öffnen und eine Reservierung für authentifizierte Benutzer hinzufügen:

```powershell
netsh http add urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/ sddl="D:(A;;GX;;;AU)"
```

Reservierung prüfen:

```powershell
netsh http show urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/
```

Wird der Port erneut geändert oder das Add-in entfernt, die manuell erstellte Reservierung löschen:

```powershell
netsh http delete urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/
```

Die Reservierung nicht an `Everyone` (`S-1-1-0`) vergeben.

## Sicherheit und Datenverarbeitung

- HTTPS für die Nextcloud-Basis-URL verwenden.
- Zertifikatsspeicher des Arbeitsplatzes, Proxy-Vertrauen und Windows-TLS-Policy aktuell halten.
- App-Passwörter sind für den aktuellen Windows-Benutzer geschützt; geschützte Zugangsdaten niemals verteilen.
- Backend-Vorlagen sind verwaltete Inhalte. Vor dem Rollout prüfen und Bearbeitungsrechte in Nextcloud begrenzen.
- Schlüssel für Secret-Links bleiben im URL-Fragment. Einen vollständigen Secret-Link vertraulich behandeln.
- IFB bindet nur an Loopback. Eine eigene URL-Reservierung vergibt Ausführungsrechte an authentifizierte lokale Benutzer, nicht an anonyme oder entfernte Benutzer.
- Die Anhangsbereinigung kann Serverdaten löschen, die für eine ungesendete Mail erstellt wurden. Beliebige öffentliche URLs werden nicht als Löschziele behandelt.
- Log-Anonymisierung ist standardmäßig aktiv. Jedes Log vor einer Weitergabe außerhalb der Organisation prüfen.

Die tägliche Update-Abfrage sendet Produkt, installierte Version, Kanal und einen wechselnden anonymen Client-Hash. Nextcloud-URL, E-Mail-Adresse, Benutzername, App-Passwort, Lizenzschlüssel oder Mandanteninhalte werden nicht gesendet. Downloads verlinken direkt auf GitHub-Release-Dateien.

## Monitoring und Support

### Regelmäßige Betriebsprüfungen

Nach einem Update von Add-in, Outlook, Proxy, Zertifikat oder Nextcloud:

1. Den Verbindungstest in den Einstellungen ausführen.
2. Eine kleine Freigabe erstellen und öffnen.
3. Einen Talk-Link erstellen und öffnen, wenn Talk aktiviert ist.
4. Die Benutzersuche testen, wenn das Systemadressbuch verwendet wird.
5. Jede gesperrte Backend-Einstellung mit einem Benutzer mit zugewiesenem Seat prüfen.
6. Die eingesetzten Signatur-Abnahmefälle testen.
7. Den lokalen IFB-Endpunkt testen, wenn IFB aktiviert ist.
8. **Einstellungen -> Erweitert -> Jetzt prüfen** öffnen und die angezeigten Release-Metadaten kontrollieren.

### Logs

Logging unter **NC Connector -> Einstellungen -> Debuggen** aktivieren. **Logs anonymisieren** aktiviert lassen, sofern der Support nicht ausdrücklich andere Daten anfordert.

Tägliche Dateien:

```text
%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD
```

Häufige Kategorien:

- `CORE`: Start, Einstellungen und Registrierung
- `API`: Nextcloud-Anfragen und Statuscodes
- `TALK`: Raum- und Terminoperationen
- `FILELINK`: Scan, Upload, Freigabe, Bereinigung und Passwort-Follow-up
- `IFB`: Listener, Cache und Free/Busy-Anfragen

Laufzeitfehler werden auch bei deaktiviertem Debug-Logging geschrieben. Bei aktivem Debug-Logging enthält die Datei zusätzlich Betriebsentscheidungen und periodischen Uploadfortschritt. Es bleiben die neuesten sieben Tagesdateien erhalten; Dateien älter als 30 Tage werden zusätzlich entfernt, soweit dies möglich ist.

### Supportpaket

Für eine reproduzierbare Störung:

1. Debug-Logging aktivieren.
2. Lokale Uhrzeit, Add-in-Version, Outlook-Version und Bitness, Windows-Version, Nextcloud-Version und relevante App-Versionen notieren.
3. Das Problem einmal reproduzieren.
4. Nur das betroffene Zeitfenster aus dem neuesten Log kopieren.
5. Sichtbare Fehlermeldung, angezeigten HTTP-Status und genauen Bedienschritt aufnehmen.
6. App-Passwörter, Autorisierungswerte, private Links, vollständige Nachrichtentexte, Empfängerlisten und Kundendaten vor der Weitergabe entfernen.

## Runbooks zur Störungsbehebung

### Add-in wird nicht geladen

1. **Outlook -> Datei -> Optionen -> Add-Ins** öffnen.
2. Unter **COM-Add-Ins** nach `NcTalkOutlook.AddIn` suchen.
3. **Deaktivierte Elemente** prüfen und das Add-in wieder aktivieren, falls Outlook es nach einem Absturz deaktiviert hat.
4. `LoadBehavior=3` im zur Outlook-Bitness passenden Registry-Pfad prüfen.
5. Prüfen, ob `C:\Program Files\NC4OL\NcTalkOutlookAddIn.dll` vorhanden ist.
6. Outlook schließen und die MSI reparieren oder neu installieren.

Wird das Add-in weiterhin nicht geladen, MSI-Log, Windows-Ereignisanzeige für Outlook/.NET und den passenden Registry-Pfad erfassen.

### Verbindungs- oder TLS-Test schlägt fehl

1. Die konfigurierte Basis-URL am betroffenen Arbeitsplatz öffnen.
2. DNS, Systemzeit, Zertifikatsvertrauen, Proxy-Authentifizierung und TLS-Inspection prüfen.
3. Prüfen, ob die URL den öffentlichen Unterpfad, aber nicht `/index.php` enthält.
4. Unter **Einstellungen -> Erweitert -> Transportsicherheit (TLS)** den von der Organisation freigegebenen Modus testen.
5. Den Verbindungstest erneut ausführen.
6. Das Ergebnis mit einem Arbeitsplatz außerhalb des betroffenen Proxy-Segments vergleichen.

Keine computerweiten TLS-Registry-Änderungen vornehmen, bevor Zertifikat, Proxy und Windows-Schannel-Policies geprüft wurden.

### Pretty URL oder Talk-Link liefert 404

Den [`/login`-Vergleich](#kurztest) ausführen. Funktioniert nur die URL mit `/index.php/login`, das Rewrite in Nginx, Apache oder Reverse Proxy korrigieren und erneut von außerhalb des Servernetzes testen.

### Upload bleibt bei null oder schlägt fehl

1. Notieren, ob der Assistent Scan, Ordnervorbereitung oder Upload anzeigt.
2. Vor der Bewertung der Netzwerkgeschwindigkeit den lokalen Scan abwarten; ein großer Quellbaum kann längere Zeit in dieser Phase verbringen.
3. Auf symbolische Links, Junctions, nicht lesbare oder während des Uploads veränderte Dateien prüfen.
4. Freien Nextcloud-Speicher prüfen; HTTP `507` bedeutet zu wenig Speicher.
5. Request-Body-Limits, Timeouts und WebDAV-Verarbeitung des Proxys prüfen.
6. Mit einer kleinen Datei, einer großen Datei und anschließend dem ursprünglichen Ordner reproduzieren.
7. `FILELINK`-Einträge für das betroffene Zeitfenster sammeln.

### Anhangsautomatisierung startet nicht

1. Konfigurierten Modus und Schwellwert unter **Anhänge** prüfen.
2. Prüfen, ob Outlook oder Exchange die Datei abgelehnt hat, bevor sie im Verfassen-Fenster erschien.
3. Auf ein anderes Outlook-Add-in prüfen, das große Anhänge verarbeitet.
4. **Nextcloud-Freigabe einfügen** als unterstützten Weg verwenden, wenn der Host den Anhang blockiert, bevor NC Connector ihn empfängt.

### Backend-Policy oder Seat wird nicht angewendet

1. Prüfen, ob `ncc_backend_4mc` installiert und aktiviert ist.
2. Prüfen, ob dem betroffenen Nextcloud-Benutzer ein aktiver Seat zugewiesen ist.
3. Client-Zugriff auf `/apps/ncc_backend_4mc/api/v1/status` prüfen.
4. Einstellungen öffnen und den angezeigten Backend- oder Seat-Status kontrollieren.
5. Prüfen, ob die Einstellung eine Vorgabe oder ein gesperrter Wert ist.
6. Mit einem anderen Pilotbenutzer mit zugewiesenem Seat vergleichen.

Freigaben und Talk können während eines Backend-Ausfalls lokale Einstellungen verwenden. Verwaltete Signaturen und separate Passwortzustellung benötigen einen gültigen Backend-Zustand.

### Verwaltete Signatur fehlt oder steht falsch

1. Backend-Seat und Signaturzuweisung prüfen.
2. Die wirksame Outlook-**Von**-SMTP-Adresse mit der im Backend zugewiesenen E-Mail-Adresse vergleichen.
3. Getrennte Schalter für neue Mail, Antwort und Weiterleitung prüfen.
4. Die passenden Fälle aus dem [Signatur-Abnahmetest](#abnahmetest-für-verwaltete-signaturen) wiederholen.
5. Einmal mit und einmal ohne Outlook-eigene Signatur testen.
6. `CORE`- und passende Compose-Logeinträge sammeln, ohne das Signatur-HTML weiterzugeben.

Eine blockierte abschließende Signaturprüfung nicht durch Kopieren unbekannten HTMLs in die Nachricht umgehen. Backend-Zugriff wiederherstellen oder Absender-/Policy-Zuweisung korrigieren.

### Benutzersuche oder Moderatorauswahl ist deaktiviert

1. Systemadressbuch und Sharing-/Autocomplete-Einstellungen prüfen.
2. `occ dav:sync-system-addressbook` ausführen.
3. Die erzeugte Adressbuch-URL für den betroffenen Benutzer testen.
4. Outlook neu starten und den Verbindungstest ausführen.

### Einstellungen können nicht geladen oder gespeichert werden

1. Outlook schließen und `settings_*.xml` sowie `settings_*.xml.bak` aus `%LOCALAPPDATA%\NC4OL\` in ein Supportverzeichnis kopieren.
2. Outlook starten und prüfen, ob die Werte aus der Sicherung wiederhergestellt wurden.
3. Ist nur das App-Passwort leer, erneut anmelden; die anderen lesbaren Einstellungen bleiben erhalten.
4. Sind beide Dateien ungültig, Einstellungen neu eintragen und ausdrücklich **Speichern** wählen. Hintergrundschreibvorgänge bleiben gesperrt, bis dieses Speichern erfolgreich war.
5. Schlägt das Speichern fehl, freien Speicherplatz, Zugriffsrechte, Endpoint-Security-Blockaden und die `CORE`-Logeinträge prüfen. Der Dialog bleibt geöffnet und die aktive Laufzeitkonfiguration wird nicht ersetzt.

Erwartetes Ergebnis: Ein erfolgreiches ausdrückliches Speichern erzeugt eine gültige Primärdatei und bewahrt die vorherige gültige Version als `.bak`.

### IFB antwortet nicht

1. Prüfen, ob IFB aktiviert ist und die Zugangsdaten vollständig sind.
2. Den konfigurierten Port prüfen.
3. Die passende URL-Reservierung prüfen.
4. Prüfen, ob ein anderer Prozess den Port belegt:

```powershell
netstat -ano | Select-String ":<ifb-port>"
```

5. `Test-NetConnection 127.0.0.1 -Port <ifb-port>` ausführen.
6. Einen Testtermin erstellen, eine bekannte Adresse aus dem Nextcloud-Systemadressbuch hinzufügen und den **Terminplanungs-Assistenten** öffnen.
7. `IFB`-Logeinträge für die Outlook-Anfrage und das CalDAV-Ergebnis prüfen. `404` bei einer direkten Anfrage ohne intern verwaltetes Pfadsegment ist zu erwarten.

Hat eine eigene Reservierung den falschen Principal, diese löschen und mit `D:(A;;GX;;;AU)` neu erstellen.
