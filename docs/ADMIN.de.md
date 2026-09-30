# Betriebsanleitung — NC Connector für Outlook

Diese Anleitung richtet sich an Administratoren und Betriebsteams, die **NC Connector für Outlook** bereitstellen und betreiben. Sie beschreibt Voraussetzungen, Rollout, verwaltete Konfiguration, Servervorbereitung, Betriebsprüfungen, Logging und Störungsbehebung.

Quellcode-Aufbau, interne Verarbeitung, Protokollimplementierung, Builds und Entwicklertests sind in [DEVELOPMENT.de.md](DEVELOPMENT.de.md) dokumentiert.

## Inhalt

- [Zweck und Zuständigkeiten](#zweck-und-zuständigkeiten)
- [Voraussetzungen und Rollout-Planung](#voraussetzungen-und-rollout-planung)
- [Bereitstellung und Anwendungslebenszyklus](#bereitstellung-und-anwendungslebenszyklus)
- [Verwaltete Konfiguration und Benutzerdaten](#verwaltete-konfiguration-und-benutzerdaten)
  - [Registry-Übersicht](#registry-übersicht)
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

Das optionale NC Connector Backend wird in einer nicht zentral verwalteten Installation für lokale Freigaben, Talk oder IFB nicht benötigt. Es ist für zentrale Policies, verwaltete Signaturen, separate Passwortzustellung und Enterprise Rollout erforderlich.

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
6. Den vorgesehenen Benutzern vor der Aktivierung von Enterprise Rollout einen aktiven Backend-Seat zuweisen.
7. Aktuelle und vorherige MSI für Rollout und Wiederherstellung bereithalten.
8. Die Einrichtung zunächst auf einem repräsentativen Arbeitsplatz durchführen, bevor die Konfiguration breit verteilt wird.

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

### Inbetriebnahme

1. Mit dem vorgesehenen Benutzerkonto anmelden, in den **Einstellungen** die Nextcloud-Verbindung testen und speichern.
2. Eine tatsächlich benötigte Funktion verwenden, beispielsweise eine Freigabe einfügen oder einen Talk-Termin erstellen. Für optionale Funktionen die jeweilige Einrichtung in dieser Anleitung beachten.
3. Bei einer Störung mit dem passenden Abschnitt unter [Runbooks zur Störungsbehebung](#runbooks-zur-störungsbehebung) fortfahren.

### Upgrade oder Rückkehr auf eine ältere Version

Die MSI ersetzt eine installierte neuere, gleiche oder ältere Version. Benutzerspezifische Einstellungen bleiben im Benutzerprofil erhalten.

Installation, direktes Upgrade und MSI-Reparatur entfernen alte NC-Connector-Free/Busy-Pfade aus den lokalen Windows-Benutzerprofilen. Das funktioniert auch, wenn `%LOCALAPPDATA%\NC4OL\` gelöscht oder umbenannt wurde. Beim nächsten Start mit aktiviertem IFB trägt NC Connector den aktuellen Pfad wieder ein. Andere Free/Busy-Anbieter und administrative Policies bleiben unverändert; manuelle Registry-Bereinigung oder `netsh`-Befehle sind nicht erforderlich.

Vor dem Setup Outlook in **allen Windows-Sitzungen** schließen. Läuft Outlook noch, bricht das Setup ab; Outlook wird nicht zwangsweise beendet. Die Rücknahme von Änderungen durch Windows Installer muss aktiviert sein.

Das Add-in meldet Release-Metadaten, installiert aber keine Updates. **Einstellungen -> Erweitert -> Über neue Versionen informieren** steuert das Popup; die tägliche Metadatenabfrage läuft auch bei deaktiviertem Popup. Freigabe und Verteilung der MSI bleiben Aufgaben der Administration.

1. Outlook schließen.
2. `%LOCALAPPDATA%\NC4OL\settings_*.xml` sichern.
3. Die gewünschte MSI über die vorhandene Installation installieren.
4. Outlook starten, die Verbindung und die vom Update betroffenen Funktionen prüfen.

Für die Rückkehr zur vorherigen Add-in-Version denselben Ablauf mit der vorherigen MSI wiederholen. Müssen zusätzlich Einstellungen zurückgespielt werden, Outlook vorher schließen und nur Dateien desselben Windows-Benutzers wiederherstellen. Ein geschütztes App-Passwort kann nach dem Kopieren zu einem anderen Windows-Konto oder Computer unlesbar sein; in diesem Fall neu authentifizieren.

### Deinstallation

1. Outlook in allen Windows-Sitzungen schließen und warten, bis kein `OUTLOOK.EXE`-Prozess mehr läuft.
2. **Windows-Einstellungen -> Apps -> Installierte Apps** verwenden oder:

```powershell
msiexec.exe /x "NCConnectorForOutlook-<version>.msi" /qn /norestart
```

Die MSI entfernt installierte Dateien, die Add-in-Registrierung, die Standard-IFB-URL-Reservierung und verbliebene eigene Outlook-Free/Busy-Pfade. Dafür muss IFB nicht vorher deaktiviert werden; auch der Datenordner wird für die Bereinigung nicht benötigt. Aktuelle Werte anderer Anbieter und administrative Policies bleiben unverändert. Benutzerspezifische Einstellungen, Caches und Logs unter `%LOCALAPPDATA%\NC4OL\` bleiben erhalten, damit eine Neuinstallation die Benutzerkonfiguration nicht löscht.

Soll vor der Deinstallation ein zuvor verwendeter externer Free/Busy-Anbieter wiederhergestellt werden, IFB in den Einstellungen deaktivieren, solange dessen gespeicherte Konfiguration noch vorhanden ist. Der Installer entfernt eigene Restwerte; die Adresse eines externen Anbieters kann er aus einem gelöschten Datenordner nicht rekonstruieren.

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

Die vorherige gültige Konfiguration bleibt als `settings_<OutlookProfile>.xml.bak` erhalten. Ist das geschützte Passwort nicht lesbar, bleiben die übrigen Einstellungen erhalten. Der Benutzer muss sich erneut anmelden und die korrigierte Konfiguration speichern; bis dahin wird die bestehende Datei nicht automatisch überschrieben.

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

Für Enterprise Rollout die nachfolgende verwaltete Registry-Policy verwenden. Vor der Aktivierung das Backend installieren und einrichten sowie jedem Benutzer einen aktiven Seat zuweisen. Community- und Pro-Seats ermöglichen denselben Zugriff.

Muss eine Profil-XML vorab verteilt werden:

- nur bereitstellen, wenn noch keine Profildatei vorhanden ist
- als Vorlage eine Datei verwenden, die mit derselben Add-in-Version erstellt wurde
- nur stabile, organisationsweit benötigte Vorgaben aufnehmen
- `AppPasswordProtected` vor der Verteilung entfernen
- jeden Benutzer über den Nextcloud-Login-Flow authentifizieren lassen
- die Datei unter dem Namen des vorgesehenen Outlook-Profils bereitstellen

Ein geschütztes App-Passwort niemals zwischen Benutzern oder Computern kopieren.

### Registry-Übersicht

Alle folgenden Werte liegen unter `HKLM\Software\Policies\NC Connector` oder `HKCU\Software\Policies\NC Connector`.

- **Vorrang:** HKLM 64-Bit → HKLM 32-Bit → HKCU 64-Bit → HKCU 32-Bit. Pro Wert gewinnt der erste vorhandene Eintrag, auch wenn er ungültig ist. Ausnahme: `NextcloudUrlLocked` gehört zur URL im selben Pfad und derselben Registry-Ansicht.
- **Managed-Modus:** Bereits einer dieser 15 Werte aktiviert [Enterprise Rollout](#enterprise-rollout), auch `0`, leer oder ungültig. Vor der Verteilung Backend einrichten und aktive Seats zuweisen. Community- und Pro-Seats sind gleichberechtigt.
- **Lokale Werte:** „Lokal; neu: …“ bedeutet: Gespeicherte Benutzereinstellung bleibt wirksam; nur ohne gespeicherten Wert gilt der genannte Produktstandard.
- **Gruppen-Regel:** Sobald ein TLS-, Logging- oder IFB-Wert gesetzt ist, wird die **gesamte jeweilige Gruppe** vorgegeben und gesperrt. Fehlende Werte innerhalb dieser Gruppe verwenden die angegebenen Produktstandards, **nicht** die lokalen Werte. Andere Gruppen bleiben davon unberührt.
- **Änderungen:** Registry-Vorgaben werden nach einem Outlook-Neustart wirksam. Nur `ShowMainRibbonTab=0` blendet Haupttab und Einstellungsbutton aus.

Für Ein/Aus-Werte `REG_DWORD` mit `0` / `1` verwenden. Alternativ werden Boolean-Strings wie `false` / `true` akzeptiert. Die Zahlenwerte für IFB benötigen dagegen tatsächlich `REG_DWORD`; `AuthMode` und `DefaultsSource` benötigen `REG_SZ`.

| Key | Typ / Werte | Wirkung und Abhängigkeiten | Wenn nicht gesetzt |
| --- | --- | --- | --- |
| `NextcloudUrl` | `REG_SZ`: öffentliche HTTPS-Nextcloud-Basis-URL, ggf. mit Unterpfad | Server vorgeben; überschreibt eine vorhandene lokale URL nur mit URL-Sperre | Lokale URL; neues Profil: leer |
| `NextcloudUrlLocked` | Ein/Aus | `1`: URL-Feld sperren; benötigt eine gültige `NextcloudUrl` im selben Pfad und derselben Registry-Ansicht | `0`: URL editierbar |
| `ShowMainRibbonTab` | Ein/Aus | `0`: Haupttab samt Einstellungen ausblenden; Share-/Talk-Aktionsbuttons bleiben erreichbar | `1`: sichtbar |
| `AuthMode` | `REG_SZ`: `LoginFlow` / `Manual` | Anmeldeart auswählen und sperren; Autostart nur mit passender Registry-URL, siehe [Anmeldung](#verwaltete-anmeldeart) | Lokal; neu: `LoginFlow`; kein verwalteter Autostart |
| `DefaultsSource` | `REG_SZ`: `local` / `backend` | Quelle editierbarer Standardwerte vorgeben; ausdrückliche Backend-Vorgabe hat Vorrang, siehe [Quellenauswahl](#quelle-der-standardwerte) | Backend-Vorgabe, sonst Benutzerauswahl, sonst `local` |
| `TransportTlsUseSystemDefault` | Ein/Aus; TLS-Gruppe | `1`: Windows wählt TLS; beide Versionsschalter sind dann wirkungslos | Lokal; neu: `0` |
| `TransportTlsEnable12` | Ein/Aus; TLS-Gruppe | TLS 1.2 erlauben, wenn Systemstandard aus ist | Lokal; neu: `1` |
| `TransportTlsEnable13` | Ein/Aus; TLS-Gruppe | TLS 1.3 erlauben, wenn Systemstandard aus ist; benötigt Laufzeit-/Windows-Unterstützung | Lokal; neu: `0` |
| `DebugLoggingEnabled` | Ein/Aus; Logging-Gruppe | Ausführliches Debug-Log einschalten; Fehler werden auch ohne Debug protokolliert | Lokal; neu: `0` |
| `LogAnonymizationEnabled` | Ein/Aus; Logging-Gruppe | Personenbezogene Logwerte anonymisieren; Passwort-/Token-Maskierung bleibt immer aktiv | Lokal; neu: `1` |
| `UpdateNotifyEnabled` | Ein/Aus | Update-Benachrichtigungen steuern und Schalter sperren; kein Abschalten des täglichen Abrufs oder von **Jetzt prüfen** | Lokal; neu: `0` |
| `IfbEnabled` | Ein/Aus; IFB-Gruppe | IFB einschalten; benötigt Anmeldung, gültigen aktiven Seat und passende URL-Reservierung | Lokal; neu: `0` |
| `IfbDays` | `REG_DWORD`: `10`, `30`, `60`, `90`; IFB-Gruppe | Zeitraum der Verfügbarkeitsdaten in Tagen | Lokal; neu: `30` |
| `IfbCacheHours` | `REG_DWORD`: `1`–`24`; IFB-Gruppe | Cache-Dauer in Stunden; gilt auch für den gemeinsamen Talk-Adressbuch-Cache | Lokal; neu: `24` |
| `IfbPort` | `REG_DWORD`: `1024`–`49151`; IFB-Gruppe | Lokaler IFB-Port; ein anderer Port benötigt eine eigene [URL-Reservierung](#eigener-ifb-port) | Lokal; neu: `7777` |

Die Gruppen-Regel bedeutet beispielsweise: Nur `TransportTlsEnable12=0` ergibt die ungültige TLS-Kombination `0 / 0 / 0`. Nur `IfbPort` zu setzen aktiviert IFB **nicht**; `IfbEnabled` bleibt dann `0`.

### Verwaltete Nextcloud-URL

Ohne URL-Sperre wird nur ein leeres URL-Feld vorbelegt. Mit Sperre ersetzt die Vorgabe die URL jedes Profils. Zugangsdaten bleiben benutzerspezifisch. Ein alleinstehendes `NextcloudUrlLocked` aktiviert zwar Enterprise Rollout, gibt aber keine URL vor und sperrt das Feld nicht.

Eine ungültige ausgewählte URL wird nicht verwendet; eine ungültige URL-Sperre gilt als aus. Ein ungültiger Ribbonwert lässt Haupttab und Einstellungen sichtbar. Ein nachrangiger Registry-Eintrag ersetzt keinen ungültigen höherrangigen Wert.

### Verwaltete Anmeldeart

Bei unvollständigen Zugangsdaten startet **Nextcloud-Freigabe einfügen** oder **Talk-Link einfügen** die Browseranmeldung direkt, wenn `AuthMode=LoginFlow` und eine gültige Registry-`NextcloudUrl` gesetzt sind. Die tatsächlich verwendete URL muss dieser Registry-URL entsprechen. Eine lokale URL allein, ein alleinstehender URL-Sperrwert oder `Manual` löst keinen Autostart aus.

Normales Öffnen der Einstellungen und vollständige Zugangsdaten lösen keinen Autostart aus. Pro Anmeldedialog erfolgt höchstens ein automatischer Versuch; nach Fehler oder Abbruch bleibt der Login-Button für einen bewussten neuen Versuch verfügbar.

Bei gültig vorgegebenem `AuthMode=LoginFlow` speichert die Erstanmeldung über Share oder Talk nach erfolgreichem Login und Verbindungstest automatisch und schließt den Dialog. Das gilt auch nach einem bewusst gestarteten erneuten Login. Erst nach erfolgreichem Speichern wird die ursprüngliche Aktion fortgesetzt, solange Mail oder Termin noch geöffnet sind. Normal geöffnete Einstellungen schließen sich nicht automatisch; `Manual` erfordert weiterhin **Speichern**.

Ungültiges `AuthMode` sperrt die Auswahl auf `LoginFlow` mit Konfigurationshinweis, aber ohne Autostart. Manuelles Anklicken des Login-Buttons bleibt möglich.

### Quelle der Standardwerte

`DefaultsSource` betrifft Freigaben, Talk, Anhangsautomatisierung, Sprachen erzeugter Texte und Signaturschalter. **local** bevorzugt gespeicherte lokale Entscheidungen, danach editierbare Backend-Startwerte. **backend** bevorzugt vorhandene Backend-Werte, danach lokale Werte. Einzelne erzwungene Policies gewinnen immer; editierbare Werte bleiben im Assistenten für die aktuelle Aktion änderbar. Die Signaturvorlage kommt weiterhin aus dem Backend.

Die Rangfolge bei gültigem aktivem Seat:

1. Ausdrückliche Backend-Quelle unter **Gruppeneinstellungen → Standardeinstellungen → Allgemein**; nur vollständige Nextcloud-Administratoren, nicht delegierbar. Mit **Im Add-on veränderbar** gewinnt eine gespeicherte Benutzerauswahl gegenüber diesem Vorschlag.
2. Bei **Keine Vorgabe** oder älterem Backend ohne dieses Feld: Registry-`DefaultsSource`, mit gesperrter Quellenauswahl.
3. Ohne Registry-Wert: Benutzerauswahl unter **Einstellungen → Erweitert → Quelle der Standardwerte**, sonst `local`.

Ohne gültigen aktiven Seat ist die Auswahl deaktiviert und die Quelle lokal; Enterprise Rollout verlangt trotzdem gültigen Zugriff. Bei wirksamer Quelle `backend` sind **Freigabe**, **Talk-Link** und **Signatur** in den Einstellungen ausgegraut. Eine ungültige Registry-Quelle verwendet gesperrtes `local` mit Konfigurationshinweis, sofern keine ausdrückliche Backend-Vorgabe gewinnt.

Nach Backend-Änderungen die Verbindung aktualisieren oder Einstellungen erneut öffnen. Zum Freigeben der Benutzerauswahl die Backend-Quelle auf **Keine Vorgabe** setzen und die Registry-Vorgabe entfernen. Ältere Clients ignorieren die neue Backend-Quellenvorgabe.

### Verwaltete Transportsicherheit (TLS)

Die drei Felder unter **Einstellungen → Erweitert → Transportsicherheit (TLS)** werden gemeinsam gesperrt. Ohne TLS-Registry-Gruppe gelten lokale Einstellungen, im frischen Profil Systemstandard aus / TLS 1.2 an / TLS 1.3 aus.

Ungültige Werte oder `0 / 0 / 0` blockieren Server-HTTP-Anfragen einschließlich Anmeldung und Update-Prüfungen mit Konfigurationshinweis. Es gibt kein stilles Einschalten von TLS 1.2 und keinen Fallback, wenn die Laufzeit eine gewählte TLS-Version ablehnt. Vorgabe korrigieren und Outlook neu starten; andere Zugangsdaten oder Seats beheben diesen Fehler nicht.

### Verwaltetes Logging

Beide Felder unter **Einstellungen → Debuggen** werden gemeinsam gesperrt. Ein ungültiger Wert verwendet nur für dieses Feld den Standard aus der Tabelle und zeigt einen Konfigurationshinweis; gültige Werte im anderen Feld bleiben wirksam. Der Hinweis wird auch ohne Debug protokolliert. Logging-Fehler blockieren keine Verbindung; siehe [Logs](#logs).

### Verwaltete Update-Benachrichtigungen

Der Schalter unter **Einstellungen → Erweitert** wird gesperrt. Ein ungültiger Wert verwendet `0` mit Konfigurationshinweis. Update-Abfragen bleiben unverändert; das Add-in installiert keine Updates automatisch.

### Verwaltetes Internet Free/Busy (IFB)

Aktivierung, Tage und Port unter **Einstellungen → IFB** sowie die Cache-Dauer unter **Erweitert** werden gemeinsam gesperrt. Details zu Anmeldung und Portbereitstellung stehen unter [IFB](#internet-freebusy-gateway-ifb). Ohne verwaltete IFB-Gruppe kann die Ersteinrichtung IFB einmalig vorwählen, sobald Zugangsdaten vorliegen und noch keine Benutzerentscheidung gespeichert ist; der Produktstandard ist aus.

Ein ungültiger Wert deaktiviert IFB mit Konfigurationshinweis, nicht Share oder Talk. Bei ungültiger Cache-Dauer verwendet der gemeinsame Adressbuch-Cache 24 Stunden; eine gültige Cache-Dauer bleibt auch bei einem anderen IFB-Fehler wirksam.

### Registry-Vorgaben zurücknehmen

Den betreffenden Wert aus allen zutreffenden Pfaden und Registry-Ansichten entfernen; bei TLS, Logging und IFB die **gesamte Gruppe** entfernen. Outlook neu starten. Nur den höherrangigen Eintrag zu löschen kann eine nachrangige Vorgabe freigeben. Gespeicherte lokale Anmeldeart, Quellenauswahl, TLS-, Logging-, Update- und IFB-Werte bleiben bei Übersteuerung erhalten. Die URL-Vorbelegung ist dagegen keine Sicherung einer früheren Serveradresse.

Andere vorhandene Keys halten Enterprise Rollout weiterhin aktiv. Zum vollständigen Beenden alle Werte der Tabelle entfernen.

### Enterprise Rollout

Enterprise Rollout gilt, sobald ein Wert der [Registry-Übersicht](#registry-übersicht) vorhanden ist. Fehlen alle fünfzehn Werte, bleibt das bisherige lokale Verhalten unverändert.

**Auswirkung beim Upgrade:** Auch eine bereits bestehende Registry-URL-Vorgabe aktiviert diesen Modus nach dem Update. Seats und Backend deshalb vor dem Rollout vorbereiten; eine ausschließlich per XML vorbelegte URL aktiviert ihn nicht.

- Nur `ShowMainRibbonTab=false` blendet den Haupttab im Explorer samt Einstellungsbutton aus. Fehlt der Wert oder ist er `true`, bleiben beide sichtbar und der vollständige Einstellungsdialog verfügbar, auch bei verwalteter URL oder URL-Sperre.
- Freigabe und Talk bleiben in ihren bisherigen Mail- und Terminbereichen erreichbar, einschließlich Inline-Antworten. Ihre Sichtbarkeit hängt nicht von `ShowMainRibbonTab` ab; es gibt keinen Ersatz- oder Statusbutton.
- Ein bestätigter gültiger, persönlich aktiver Seat erlaubt NC Connector. Ein Community-Seat ist einem Pro-Seat vollständig gleichgestellt. Globale Überbelegung allein sperrt einen aktiven Seat nicht.
- Ein bestätigt fehlendes Backend erzeugt den Installations-/Einrichtungshinweis. Ein fehlender, pausierter oder ungültiger Seat erzeugt den Seat-Hinweis zur zentral verwalteten Installation. Ein unbestätigter Verbindungs- oder Antwortfehler erzeugt einen Prüfhinweis, keine falsche Aussage über fehlendes Backend oder fehlenden Seat.
- Freigaben, Talk, Anhangsautomatisierung, verwaltete Signaturen und IFB benötigen den Rollout-Zugriff. Normale Outlook-Mails bleiben nutzbar; der Modus ist keine allgemeine Outlook-Versandsperre oder Schutz gegen Datenabfluss.
- Ein zuletzt bestätigter Zugriff kann bei einem vorübergehenden Abruffehler weiter gelten. Eine neue bestätigte Ablehnung ersetzt diesen Status. Bereits geöffnete Assistenten behalten ihren beim Öffnen geladenen Status.
- Bestehende Dateien, Freigaben oder Termine werden nicht wegen eines fehlenden Seats gelöscht. Das Verwerfen neu angelegter Freigaben und der Versand bereits vorbereiteter separater Passwort-Mails bleiben davon unberührt.

Erstanmeldung:

1. Nach dem Verteilen der Registry-Werte Outlook neu starten.
2. Ohne Zugangsdaten in einer Mail **Nextcloud-Freigabe einfügen** oder in einem Termin **Talk-Link einfügen** anklicken. Der vorhandene Einstellungsdialog öffnet sich direkt mit dem blauen Hinweis **Mit Nextcloud verbinden**, ohne vorgeschaltete Fehlermeldung. Das gilt auch für nicht verwaltete Installationen. Nur bei `ShowMainRibbonTab=false` sind die anderen Einstellungstabs nicht verfügbar; andernfalls ist der vollständige Dialog auch über **NC Connector -> Einstellungen** erreichbar.
3. Die Anmeldung abschließen. Bei gültig vorgegebenem `AuthMode=LoginFlow` werden die geprüften Zugangsdaten automatisch gespeichert und der Dialog geschlossen; sonst **Speichern** anklicken. Die [verwaltete Anmeldeart](#verwaltete-anmeldeart) erklärt den Browser-Autostart. Verwaltete URL und Sperre bleiben wirksam; vorhandene Einstellungen bleiben erhalten. Nach erfolgreichem Speichern wird die ursprüngliche Aktion fortgesetzt, sofern Mail oder Termin noch geöffnet sind (eine Inline-Antwort muss weiterhin aktiv sein). Abbrechen beendet die Aktion ohne weitere Meldung. Die Anmeldung allein lädt keine Dateien hoch und erstellt keinen Talk-Raum.
4. Bei einem Backend- oder Seat-Hinweis das Backend einrichten beziehungsweise dem angemeldeten Benutzer einen gültigen aktiven Seat zuweisen. Abgelehnte Zugangsdaten öffnen die Anmeldung erneut; Verbindungsfehler werden gesondert angezeigt. Ein ausgeblendeter Haupttab bleibt dabei ausgeblendet.

**Anzeige wiederherstellen:** Die wirksame Policy `ShowMainRibbonTab=false` entfernen oder auf `true` setzen und Outlook neu starten. Die verwaltete Zugriffsprüfung bleibt bestehen.

**Enterprise Rollout beenden:** Alle fünfzehn Auslöser an allen zutreffenden Policy-Pfaden und Registry-Ansichten entfernen und Outlook neu starten. Gespeicherte Zugangsdaten und Einstellungen bleiben erhalten.

Registry-Policies steuern den administrativen Rollout. Sie schützen nicht vor Benutzern, die diese Policy verändern oder das Add-in ersetzen können. Policy-Pfade und Berechtigungen zur Softwareverteilung entsprechend schützen.

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

Die lokalen Ausweichmöglichkeiten in diesem Abschnitt gelten für nicht zentral verwaltete Installationen. Bei [Enterprise Rollout](#enterprise-rollout) sind das Backend und ein aktiv zugewiesener Seat erforderlich; die Hinweise zur zentral verwalteten Installation haben Vorrang.

### Voraussetzungen und Betriebszustände

Backendverwaltete Funktionen benötigen:

- die installierte und aktivierte App `ncc_backend_4mc`
- Client-Zugriff auf `/apps/ncc_backend_4mc/api/v1/status`
- einen dem aktuellen Nextcloud-Benutzer zugewiesenen aktiven Seat
- eine von der installierten Backend-Version unterstützte Policy-Domain

Beobachtbares Verhalten je Zustand:

- **Keine Backend-Konfiguration:** Freigaben, Talk und IFB verwenden lokale Einstellungen. Zentrale Signaturen und separate Passwortzustellung sind nicht verfügbar.
- **Erreichbares Backend mit aktivem Seat:** Die [Quelle der Standardwerte](#quelle-der-standardwerte) bestimmt den Vorrang lokaler oder zentraler Vorgaben für editierbare Felder. Gesperrte Backend-Werte haben immer Vorrang und können in Outlook nicht geändert werden.
- **Erreichbares Backend ohne nutzbaren Seat:** Freigaben und Talk verwenden lokale Einstellungen; Outlook zeigt den Seat- oder Lizenzstatus. Zentrale Signaturen und separate Passwortzustellung sind nicht verfügbar.
- **Backend vorübergehend nicht erreichbar:** Ein zuvor bestätigter Status für dasselbe Konto kann Standardwertequelle und Policies beibehalten. Ohne nutzbaren bestätigten Status verwenden Freigaben und Talk bei nicht verwalteten Installationen lokale Einstellungen. Eine passende Mail mit verpflichtender zentraler Signatur kann geöffnet und ungesendet bleiben, bis die Signatur-Policy wieder geprüft werden kann.
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

Werte editierbar lassen, wenn Benutzer sie für eine einzelne Aktion anpassen dürfen. Die [Quelle der Standardwerte](#quelle-der-standardwerte) bestimmt lokale oder zentrale Startwerte; einzelne Einstellungen nur sperren, wenn Benutzer sie nicht ändern dürfen. Die Richtlinien gelten für Benutzer mit gültigem aktivem Seat; die vorgesehenen Zuweisungen vor dem Rollout einrichten.

Einstellungen, Freigabe- und Talk-Assistent, Anhangsautomatisierung und Signaturschalter verwenden die ausgewählte Quelle, ohne einzelne Policy-Sperren zu verändern. Auch ein ausdrücklich gespeichertes `false` oder ein Wert gleich dem Produktstandard bleibt eine Benutzerentscheidung. Das bloße Speichern der Zugangsdaten legt unberührte Optionen nicht fest. Administrative Übersteuerungen überschreiben die gespeicherte lokale Auswahl nie; sie gilt wieder, sobald lokale Standardwerte wirksam sind und das Feld entsperrt ist. Ohne nutzbaren Seat bleiben lokale Einstellungen bei nicht verwalteten Installationen verfügbar; Seat-Funktionen bleiben eingeschränkt.

Neue Backend-Ablaufvorgaben beginnen bei einem Tag. Eine Null-Tage-Vorgabe älterer Backends wird einheitlich als ein Tag interpretiert. Bestehende Freigaben und eine lokal gespeicherte Deaktivierung des Ablaufdatums bleiben unverändert. Anhangsschwellen liegen bei 1–10240 MB; ein ausdrückliches Backend-`null` deaktiviert die Schwelle, eine alte Backend-Null behält die etablierte Bedeutung von 5 MB.

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

Nicht unterstütztes oder unsicheres HTML kann entfernt oder die Vorlage abgelehnt werden. Vor der Verteilung die Darstellung der eigenen Vorlage in den tatsächlich verwendeten Outlook-Nachrichtenformaten und Office-Designs kontrollieren.

### Verwaltete Signaturen einrichten

1. Im Backend dem Benutzer einen aktiven Seat und die Signaturvorlage zuweisen.
2. Die zugewiesene E-Mail-Adresse mit der wirksamen Outlook-**Von**-Adresse abgleichen. Auch bei Shared Mailboxes und delegierten Absendern muss die tatsächliche Absender-SMTP-Adresse übereinstimmen; das angemeldete Nextcloud-Konto allein genügt nicht.
3. Die Signatur nach Bedarf für neue Mails, Antworten und Weiterleitungen aktivieren. Nur die Werte sperren, die der Benutzer nicht ändern darf.

Bei neuen Mails steht die Signatur nach dem selbst geschriebenen Text, bei Antworten und Weiterleitungen oberhalb der zitierten Nachricht. Wechselt der Benutzer zu einer anderen Absenderadresse, wird die verwaltete Signatur entfernt; die Signatur einer nicht zugewiesenen Identität bleibt unverändert. Auch separate Passwort-Mails erhalten die Vorlage nur bei passendem Absender.

Nach der Zuweisung eine Nachricht mit der vorgesehenen Absenderadresse öffnen und die Darstellung der Vorlage kontrollieren. Bei Abweichungen mit [Verwaltete Signatur fehlt oder steht falsch](#verwaltete-signatur-fehlt-oder-steht-falsch) fortfahren.

Kann eine erforderliche abschließende Signaturprüfung nicht abgeschlossen werden, lässt Outlook die Mail geöffnet, statt sie mit ungeprüftem Signaturzustand zu senden.

Die Meldung unterscheidet eine noch nicht verfügbare Signaturrichtlinie von einer nicht sicher aktualisierbaren Signatur. Ist die Richtlinie nicht verfügbar, die Nextcloud-Verbindung prüfen und das Senden erneut versuchen. Eine zuvor bestätigte Richtlinie desselben Kontos bleibt nach einem fehlgeschlagenen Refresh nutzbar; eine neu empfangene Ablehnung wird wirksam.

## Funktionsbetrieb

### Freigaben und Uploads

Die Sprache des Freigabe-HTML-Blocks wird unter **Einstellungen -> Freigabe** unterhalb der Freigabevorgaben gewählt. Vorhandene Auswahlen bleiben erhalten; eine vom Backend gesperrte Sprache bleibt schreibgeschützt.

Der Freigabe-Assistent akzeptiert lokale Dateien und Ordner sowie vorhandene Inhalte aus der eigenen Nextcloud des konfigurierten Benutzers. **Meine Nextcloud** zeigt Dateien, Ordner, Speicherinformationen und Vorschauen. Dokumentvorschauen, etwa für PDF- oder Office-Dateien, hängen von den auf dem Server aktivierten Vorschau-Anbietern ab. Diese Quelle funktioniert ohne NC Connector Backend, sofern Enterprise Rollout nicht aktiviert ist.

Ausgewählte Nextcloud-Inhalte werden innerhalb desselben Kontos in den neuen Freigabeordner kopiert. Das Original bleibt unverändert und wird für die Übertragung nicht nach Outlook heruntergeladen. Für eine Vorschau fordert Outlook zuerst ein größenbegrenztes, von Nextcloud erzeugtes Bild an. Hat der Server für eine unterstützte Bilddatei keine erzeugte Vorschau, kann Outlook vorübergehend das Originalbild bis 5 MiB laden. Andere Originaldateien werden für Vorschauen nicht heruntergeladen.

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

Outlook oder Exchange kann einen großen Anhang ablehnen, bevor NC Connector ihn übernehmen kann. In diesem Fall müssen Benutzer **Nextcloud-Freigabe einfügen** wählen und die Datei direkt im Freigabe-Assistenten hinzufügen.

### Ungesendete Mail und Freigabebereinigung

Wird eine Nachricht nach dem Einfügen einer neuen Freigabe verworfen, bevor Outlook sie gespeichert hat, entfernt NC Connector den dafür neu erstellten Serverordner. Das gilt nur für die neu angelegte Freigabe, nicht für ihre ursprünglichen Quelldateien.

- Manuelles oder automatisches Speichern eines Entwurfs behält die Freigabe. Versand, „Später senden“ und Offline-Postausgang behalten sie ebenfalls.
- Ein abgebrochener Schließvorgang oder das Verschieben einer Inline-Antwort in ein eigenes Fenster entfernt die Freigabe nicht.
- Das spätere Löschen eines bereits von Outlook gespeicherten Entwurfs wird nicht verfolgt. Seine ungenutzte Freigabe muss manuell in Nextcloud entfernt werden.
- Kann der Freigabeblock nicht eingefügt werden, meldet der Assistent einen Fehler und versucht, den neu erzeugten Serverordner zu entfernen.
- Eine erzwingende Anhangs-Policy blockiert den Versand, solange ein normaler Anhang in der Nachricht verbleibt, der über NC Connector hätte laufen müssen.
- Die separate Passwortzustellung wird beim Klick auf **Senden** direkt übergeben, wie im folgenden Abschnitt beschrieben.

### Separate Passwortzustellung

Separate Passwortzustellung benötigt NC Connector Backend und einen aktiven Seat.

- Die Hauptmail enthält kein Klartextpasswort.
- Die separate Passwortzustellung ist an die geöffnete Verfassen-Sitzung gebunden, in der die Freigabe erstellt wurde. Der Benutzer muss in genau dieser weiterhin geöffneten Nachricht auf **Senden** klicken. Speichern und AutoSave unterbrechen den Ablauf nicht, solange das Verfassen-Fenster geöffnet bleibt.
- Das Speichern der Hauptmail als Entwurf oder `.oft`-Vorlage mit anschließendem Schließen wird nicht unterstützt. Beim erneuten Öffnen des Entwurfs, einem Outlook-Neustart vor dem ersten Sendeversuch oder einer neuen Nachricht aus dieser Vorlage bleibt der sichtbare Freigabeblock erhalten, der Zustand für die Passwortzustellung jedoch nicht. Es wird keine Passwort-Follow-up-Mail erstellt; vor dem Versand muss in der endgültigen Nachricht eine neue Freigabe erstellt werden.
- Beim Klick auf **Senden** wird die Passwort-Mail sofort über dasselbe wirksame Outlook-Konto zum Versand übergeben.
- Die Passwort-Mail wartet nicht auf eine verzögerte oder Offline-Übermittlung der Hauptmail. Sie kann deshalb bereits versendet sein, während die Hauptmail noch im Postausgang liegt oder von Outlook abgelehnt wird.
- Schlägt die Absenderprüfung oder automatische Übergabe eindeutig fehl, öffnet Outlook eine vollständig vorbereitete Nachricht zum manuellen Senden. Bei einem mehrdeutigen Outlook-Übermittlungsstatus wird kein automatisches Duplikat erzeugt.
- Im Secrets-Modus wird für jeden endgültigen Empfänger ein eigener einmaliger Secret-Link erstellt.
- Gleiche SMTP-Adressen in An, Cc und Bcc erhalten nur einen Secret-Follow-up.
- Schlägt die Secrets-Erstellung fehl, verwendet Outlook die Klartext-Passwort-Follow-up-Mail und zeigt eine Warnung.
- Eine passende Backend-Signatur wird nur eingefügt, wenn auch der Follow-up-Absender mit der zugewiesenen Signaturadresse übereinstimmt.

### Talk-Raum-Lebenszyklus

Die Sprache des Talk-Beschreibungstextes wird unter **Einstellungen -> Talk-Link** unterhalb der Talk-Vorgaben gewählt, nicht mehr unter Erweitert. Vorhandene Auswahlen und Backend-Sperren bleiben unverändert.

Das Löschen eines gespeicherten Outlook-Termins entfernt den zugehörigen entfernten Talk-Raum nur, wenn die Einstellung ausdrücklich aktiviert ist und der Termin NC Connector-Raummetadaten enthält. Die Einstellung ist standardmäßig deaktiviert. Ein in Ort oder Nachrichtentext kopierter Talk-Link reicht für eine entfernte Löschung nicht aus.

Das Löschen eines einzelnen Vorkommens oder einer Ausnahme einer Terminserie entfernt den gemeinsamen Raum nicht. Nur ein Nicht-Serientermin oder der Serienmaster kann die Raumlöschung vormerken.

Die Raumlöschung wird beim Löschen eines ausgewählten Termins aus der Outlook-Kalenderansicht oder aus dem geöffneten Termin ausgelöst.

Vorgemerkte Raumlöschungen werden pro Outlook-Profil gespeichert und nach vorübergehenden Nextcloud-Fehlern oder einem Outlook-Neustart im Hintergrund wiederholt. Die Bereinigung eines neu erstellten Raums aus einem ungespeicherten und verworfenen Termin bleibt aktiv.

Die Moderatorrolle kann nur an eine andere Person übergeben werden. Nach erfolgreicher Übergabe verlässt der ursprüngliche Moderator den Raum.

Die Raumlöschung nur aktivieren, wenn das Löschen eines Outlook-Termins auch den zugehörigen Talk-Raum entfernen soll. Benutzer vorab über diese Folge informieren: Das Entfernen des Raums betrifft alle Teilnehmer, nicht nur den löschenden Benutzer.

## Internet Free/Busy Gateway (IFB)

### Zweck und Aktivierung

IFB lässt Outlook Nextcloud-Free/Busy-Daten über einen lokalen HTTP-Endpunkt abfragen. Es ist standardmäßig ausgeschaltet. Bei gesperrten Einstellungen die [verwaltete IFB-Policy](#verwaltetes-internet-freebusy-ifb) verwenden; andernfalls lokal konfigurieren:

1. Das Nextcloud-Systemadressbuch prüfen.
2. **NC Connector -> Einstellungen -> IFB** öffnen.
3. IFB aktivieren und Anzahl der Tage sowie lokalen Port wählen. Die gemeinsame Adressbuch-Cache-Dauer unter **Einstellungen -> Erweitert** setzen. Vorgaben sind 30 Tage, 24 Cache-Stunden und Port `7777`.
4. Speichern und Outlook neu starten.

Reservierter Listener-Namespace:

```text
http://127.0.0.1:7777/nc-ifb/
```

Die MSI reserviert den Standard-URL-Namespace für authentifizierte Windows-Benutzer. NC Connector ergänzt die Outlook-Free/Busy-URL um ein zufälliges Pfadsegment; Anfragen ohne dieses Segment erhalten `404`. Der geheime Pfad wird intern verwaltet und in dieser Anleitung bewusst nicht angezeigt.

Beim Aktivieren von IFB werden nur benutzerspezifische Outlook-Free/Busy-Werte aktualisiert. Beim Deaktivieren stellt NC Connector die vorherigen Werte wieder her, sofern sie seitdem nicht durch Administratoren oder andere Anwendungen geändert wurden. Werte unter `Software\Policies` werden nur auf Konflikte geprüft und nie geschrieben. Zwischengespeicherte Adressbuchdaten sind nach Outlook-Profil und Nextcloud-Konto getrennt.

Der Listener läuft nur, solange Outlook läuft, IFB effektiv aktiviert ist und die gespeicherten Nextcloud-Zugangsdaten vollständig sind. Ungültige verwaltete IFB-Einstellungen verhindern den Listener-Start; Enterprise Rollout benötigt für Anfragen außerdem bestätigten Backend-Zugriff und einen aktiv zugewiesenen Seat.

### Standard-Reservierung prüfen

```powershell
netsh http show urlacl | Select-String -Pattern "127.0.0.1:7777/nc-ifb"
Test-NetConnection 127.0.0.1 -Port 7777
```

Danach in Outlook einen Testtermin erstellen, eine Adresse aus dem Nextcloud-Systemadressbuch hinzufügen und den **Terminplanungs-Assistenten** öffnen. Erwartetes Ergebnis: Die Reservierung ist vorhanden, der TCP-Test ist erfolgreich und Outlook zeigt Free/Busy-Daten. Eine direkte Anfrage an den öffentlichen Pfad `/nc-ifb/freebusy/...` muss `404` liefern.

### Eigener IFB-Port

Gültige konfigurierte Ports reichen von `1024` bis `49151`. Die MSI erstellt nur für Port `7777` eine Reservierung. Das gilt gleichermaßen für lokale Einstellungen und verwaltete IFB-Policies; das Add-in erhöht seine Rechte nicht und erstellt keine eigene Reservierung. Für einen anderen Port muss ein Administrator eine PowerShell mit erhöhten Rechten öffnen und eine Reservierung für authentifizierte Benutzer hinzufügen:

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

Nach Änderungen die jeweils betroffenen Betriebsfunktionen kontrollieren:

- Nach Änderungen an Proxy, Zertifikaten, TLS oder Nextcloud den Verbindungstest in den Einstellungen ausführen.
- Nach einer Policy- oder Vorlagenänderung die wirksame Einstellung beziehungsweise Darstellung unter einem betroffenen Benutzerkonto kontrollieren.
- Nach einem Add-in- oder Outlook-Update die eingesetzten, vom Update betroffenen Funktionen auf einem repräsentativen Arbeitsplatz verwenden und bei Problemen die Logs auswerten.

### Logs

Logging unter **NC Connector -> Einstellungen -> Debuggen** aktivieren. **Logs anonymisieren** aktiviert lassen, sofern der Support nicht ausdrücklich andere Daten anfordert. Organisationsvorgaben können beide Felder sperren; siehe [Verwaltetes Logging](#verwaltetes-logging).

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
4. Unter **Einstellungen -> Erweitert -> Transportsicherheit (TLS)** den von der Organisation freigegebenen Modus testen. Ist die TLS-Gruppe gesperrt, stattdessen die [verwaltete TLS-Policy](#verwaltete-transportsicherheit-tls) prüfen, ungültige Werte am wirksamen Registry-Pfad korrigieren und Outlook neu starten. Ein TLS-Policy-Konfigurationsfehler muss vor Anmelde- oder Seat-Prüfungen behoben sein.
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
4. Einstellungen öffnen und den angezeigten Backend- oder Seat-Status kontrollieren. Wenn `ShowMainRibbonTab=false` die Einstellungen ausblendet, den Zugriffshinweis über Freigabe oder Talk prüfen.
5. Prüfen, ob die Einstellung eine Vorgabe oder ein gesperrter Wert ist.
6. Bei benutzerspezifischen Problemen mit einem funktionierenden Konto mit zugewiesenem Seat vergleichen.

In nicht zentral verwalteten Installationen können Freigaben und Talk während eines Backend-Ausfalls lokale Einstellungen verwenden. Verwaltete Signaturen und separate Passwortzustellung benötigen einen gültigen Backend-Zustand. Enterprise Rollout benötigt bestätigten Zugriff oder einen passenden zuletzt erfolgreich bestätigten Status wie oben beschrieben.

### Verwaltete Signatur fehlt oder steht falsch

1. Backend-Seat und Signaturzuweisung prüfen.
2. Die wirksame Outlook-**Von**-SMTP-Adresse mit der im Backend zugewiesenen E-Mail-Adresse vergleichen.
3. Getrennte Schalter für neue Mail, Antwort und Weiterleitung prüfen.
4. Für den betroffenen Nachrichtentyp Absender und Vorlage anhand von [Verwaltete Signaturen einrichten](#verwaltete-signaturen-einrichten) kontrollieren.
5. Bei doppelten Signaturen prüfen, ob zusätzlich eine Outlook-eigene Signatur eingerichtet ist.
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

1. Prüfen, ob IFB aktiviert ist und die Zugangsdaten vollständig sind. Bei gesperrten Einstellungen die [verwaltete IFB-Policy](#verwaltetes-internet-freebusy-ifb) prüfen, Konfigurationsfehler korrigieren und Outlook neu starten. Enterprise Rollout benötigt außerdem bestätigten Backend-Zugriff und einen aktiv zugewiesenen Seat.
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
