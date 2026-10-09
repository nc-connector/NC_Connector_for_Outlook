<a id="betriebsanleitung--nc-connector-für-outlook"></a>
<a id="operations-guide--nc-connector-for-outlook"></a>
<a id="zweck-und-zuständigkeiten"></a>
<a id="scope-and-responsibilities"></a>

# NC Connector für Outlook – Administration

Installation, zentrale Konfiguration und Fehlerbehandlung für Version 3.4.3.

<a id="contents"></a>

## Inhalt

- [Voraussetzungen](#voraussetzungen)
- [Installation und Anmeldung](#installation-und-anmeldung)
- [Enterprise Rollout](#enterprise-rollout)
- [Registry-Referenz](#registry-referenz)
- [Backend-Vorgaben und Signaturen](#backend-vorgaben-und-signaturen)
- [Hinweise für den Betrieb](#hinweise-für-den-betrieb)
- [Internet Free/Busy einrichten](#internet-freebusy-einrichten)
- [Update, Sicherung und Deinstallation](#update-sicherung-und-deinstallation)
- [Fehlerbehandlung](#fehlerbehandlung)
- [Logs und Support](#logs-und-support)

<a id="voraussetzungen-und-rollout-planung"></a>
<a id="requirements-and-rollout-planning"></a>
<a id="client-voraussetzungen"></a>
<a id="client-requirements"></a>
<a id="nextcloud-voraussetzungen"></a>
<a id="nextcloud-requirements"></a>
<a id="nextcloud-server-vorbereiten"></a>
<a id="nextcloud-server-preparation"></a>

## Voraussetzungen

| Bereich | Voraussetzung |
| --- | --- |
| Arbeitsplatz | Windows 10 oder 11, jeweils 64 Bit; .NET Framework 4.7.2 |
| Outlook | Outlook classic ab 2019, 32 oder 64 Bit. Das neue Outlook wird nicht unterstützt. |
| Installation | Administratorrechte; offene Arbeit speichern und Outlook vor dem Setup möglichst in allen Windows-Sitzungen schließen. |
| Nextcloud | Version 32 oder neuer, erreichbar über HTTPS |
| Freigaben | Nextcloud Files Sharing mit erlaubten öffentlichen Linkfreigaben, ausreichender Speicherplatz und Schreibrechte im Zielordner |
| Besprechungen | Nextcloud Talk |
| Benutzersuche und Verfügbarkeiten | Freigegebenes Nextcloud-Systemadressbuch |
| Zentrale Einstellungen und Signaturen, separate Passwort-Mail, Enterprise Rollout | NC Connector Backend und ein gültig zugeteilter NC Connector Seat für den jeweiligen Benutzer |
| Einmalige Passwort-Links | Zusätzlich die Nextcloud-App Secrets |

Ohne zentrale Verwaltung lassen sich Freigaben, Talk und Internet Free/Busy auch ohne Backend mit lokalen Einstellungen nutzen.

<a id="netzwerk-voraussetzungen"></a>
<a id="network-requirements"></a>

Als Serveradresse die öffentliche Nextcloud-URL verwenden, etwa `https://cloud.example.com` oder `https://cloud.example.com/nextcloud`. Kein `/index.php`, keine Zugangsdaten und keine Abfrageparameter anhängen. Zertifikate müssen am Arbeitsplatz als vertrauenswürdig gelten; Proxy und Firewall müssen auch OCS- und WebDAV-Zugriffe zulassen.

Für Update-Hinweise fragt das Add-in `https://nc-connector.de/wp-json/ncc/v1/update-check` ab. MSI-Downloads erfolgen über GitHub. IFB arbeitet ausschließlich auf dem lokalen Rechner und benötigt keine eingehende Firewall-Freigabe aus dem Netzwerk.

<a id="bereitstellung-und-anwendungslebenszyklus"></a>
<a id="deployment-and-application-lifecycle"></a>
<a id="installation"></a>
<a id="install"></a>
<a id="inbetriebnahme"></a>
<a id="initial-setup"></a>

## Installation und Anmeldung

1. Offene Arbeit speichern und Outlook möglichst in allen Windows-Sitzungen schließen.
2. Die MSI mit Administratorrechten installieren.
3. Outlook starten und **NC Connector → Einstellungen** öffnen.
4. Nextcloud-URL eintragen und anmelden. Beim Login Flow erfolgt die Anmeldung im Browser; bei manueller Anmeldung Benutzername und App-Passwort eintragen.
5. Verbindung testen und Einstellungen speichern.

Ist Outlook noch geöffnet, kann Windows Installer anbieten, es geordnet zu schließen. Eine unbeaufsichtigte Installation versucht dies automatisch. Läuft Outlook danach weiter, etwa in einer anderen Windows-Sitzung, stoppt das Setup vor dem Austausch der installierten Version. Outlook dann manuell schließen und die Installation wiederholen. Das Setup erzwingt kein Beenden und garantiert keinen automatischen Neustart von Outlook.

Für die unbeaufsichtigte Installation mit Protokoll:

```powershell
msiexec.exe /i "NCConnectorForOutlook-3.4.3.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-install.log"
```

<a id="basisprüfungen"></a>
<a id="base-service-checks"></a>
<a id="regelmäßige-betriebsprüfungen"></a>
<a id="routine-operating-checks"></a>

Anschließend sollte NC Connector unter **Datei → Optionen → Add-Ins → COM-Add-Ins** erscheinen. Vor der breiten Verteilung auf einem Arbeitsplatz eine Freigabe oder einen Talk-Link mit der vorgesehenen Konfiguration erstellen. Nach späteren Änderungen an Server, Proxy oder Richtlinien die Verbindung und die betroffene Funktion erneut prüfen.

<a id="verwaltete-konfiguration-und-benutzerdaten"></a>
<a id="managed-configuration-and-user-data"></a>

## Enterprise Rollout

Mit Registry-Vorgaben können Sie Serveradresse und Einstellungen verteilen, einzelne Bereiche sperren und das Einstellungsmenü ausblenden.

**Sobald einer der Werte aus der Registry-Referenz vorhanden ist, braucht jeder Benutzer das Backend und einen gültig zugeteilten NC Connector Seat.** Das gilt auch für ausgeschaltete, leere oder ungültige Werte. Beim Update auf 3.4.3 betrifft dies auch bereits vorhandene URL-Vorgaben aus der Registry. Backend und Seat-Zuweisungen deshalb vor dem Rollout vorbereiten.

<a id="checkliste-vor-dem-rollout"></a>
<a id="pre-deployment-checklist"></a>
<a id="rollout-und-vorbelegung"></a>
<a id="rollout-and-pre-seeding"></a>

### Rollout einrichten

1. Im Nextcloud Backend die vorgesehenen Benutzer mit Seats versorgen und die gewünschten Richtlinien festlegen.
2. Die benötigten Registry-Werte per Softwareverteilung oder Gruppenrichtlinie bereitstellen.
3. MSI installieren beziehungsweise aktualisieren und Outlook neu starten.
4. Die Erstanmeldung mit einem vorgesehenen Benutzerkonto prüfen.

Für eine weitgehend vorbereitete Anmeldung setzen Sie `NextcloudUrl`, `NextcloudUrlLocked=1` und `AuthMode=LoginFlow`. Soll der Haupttab samt Einstellungen verborgen bleiben, zusätzlich `ShowMainRibbonTab=0` setzen. Freigabe- und Talk-Schaltflächen bleiben in Mails und Terminen erreichbar.

<a id="verwaltete-anmeldeart"></a>
<a id="managed-sign-in-method"></a>

### Anmeldung bei ausgeblendeten Einstellungen

Fehlen Zugangsdaten, öffnet **Nextcloud-Freigabe einfügen** oder **Talk-Link einfügen** den Anmeldedialog. Bei `AuthMode=LoginFlow` und passender Registry-URL startet die Browseranmeldung automatisch.

Ist `AuthMode=LoginFlow` gültig vorgegeben, werden nach erfolgreicher Anmeldung und Verbindungsprüfung die Zugangsdaten automatisch gespeichert und der Dialog geschlossen – auch wenn der Benutzer die Browseranmeldung selbst gestartet hat. Ohne diese Vorgabe speichert er selbst. Bei `AuthMode=Manual` trägt er Benutzername und App-Passwort ein.

Danach wird die angefangene Aktion fortgesetzt, solange die Mail oder der Termin noch geöffnet ist. Eine fehlgeschlagene Anmeldung kann im geöffneten Dialog erneut versucht werden.

Ohne `ShowMainRibbonTab=0` bleiben Haupttab und Einstellungen sichtbar. Allein eine verwaltete URL blendet nichts aus.

<a id="registry-vorgaben-zurücknehmen"></a>
<a id="removing-registry-overrides"></a>

### Vorgaben ändern oder zurücknehmen

Registry-Änderungen werden nach einem Outlook-Neustart übernommen. Um die Einstellungen wieder einzublenden, `ShowMainRibbonTab` auf `1` setzen oder den Wert entfernen.

Um eine Einstellung wieder lokal freizugeben, ihre Vorgabe aus allen verwendeten Registry-Pfaden und Ansichten entfernen. Bei TLS, Logging und IFB jeweils die ganze Gruppe entfernen. Eine nachrangige Vorgabe kann sonst wieder wirksam werden.

Enterprise Rollout endet erst, wenn keiner der aufgeführten Werte mehr vorhanden ist. Gespeicherte Zugangsdaten und lokale Einstellungen bleiben erhalten. Eine früher verwendete Serveradresse wird durch das Entfernen der URL-Vorgabe nicht automatisch wiederhergestellt.

Enterprise Rollout beschränkt NC-Connector-Funktionen, nicht den normalen Outlook-Versand. Die ausdrückliche Vorgabe `SendPolicyFailureMode=failclosed` verhindert zusätzlich den Versand, solange die zentralen Versandvorgaben noch unbekannt sind. Ein bestätigt fehlender, pausierter oder ungültiger Seat deaktiviert geschützte Funktionen, blockiert aber keine normale Mail. Dies ersetzt keine Sicherheitsrichtlinie zur Verhinderung von Datenabfluss.

<a id="registry-übersicht"></a>
<a id="registry-reference"></a>

## Registry-Referenz

Die Werte werden unter einem dieser Pfade gesetzt:

```text
HKLM\Software\Policies\NC Connector
HKCU\Software\Policies\NC Connector
```

Vorrang pro Wert: **HKLM 64 Bit → HKLM 32 Bit → HKCU 64 Bit → HKCU 32 Bit**. Der erste vorhandene Wert gewinnt, auch wenn er ungültig ist. Die URL-Sperre muss bei der zugehörigen URL im selben Pfad und derselben Ansicht liegen.

Für Schalter `REG_DWORD` mit `0` = aus und `1` = an verwenden. Die Texte `false` und `true` werden ebenfalls akzeptiert. Auswahllisten benötigen `REG_SZ`, Zahlenwerte für IFB `REG_DWORD`.

<a id="verbindung-und-anmeldung"></a>
<a id="connection-and-sign-in"></a>
<a id="oberfläche-und-standardwerte"></a>
<a id="interface-and-defaults"></a>
<a id="verwaltete-nextcloud-url"></a>
<a id="managed-nextcloud-url"></a>

### Verbindung und Oberfläche

| Wert | Typ / Werte | Wirkung | Ohne Vorgabe |
| --- | --- | --- | --- |
| `NextcloudUrl` | `REG_SZ`, HTTPS-URL | Füllt eine noch leere Serveradresse aus. Ersetzt eine gespeicherte Adresse nur zusammen mit der URL-Sperre. | Gespeicherte Adresse; bei neuen Profilen leer |
| `NextcloudUrlLocked` | `REG_DWORD`, `0` / `1` | Sperrt die mit `NextcloudUrl` vorgegebene Adresse. Ohne gültige URL im selben Pfad keine URL-Sperre. | URL bearbeitbar |
| `AuthMode` | `REG_SZ`, `LoginFlow` / `Manual` | Wählt die Anmeldeart und sperrt die Auswahl. | Gespeicherte Auswahl; bei neuen Profilen `LoginFlow` |
| `ShowMainRibbonTab` | `REG_DWORD`, `0` / `1` | `0` blendet Haupttab und Einstellungsbutton aus. | Haupttab sichtbar |
| `DefaultsSource` | `REG_SZ`, `local` / `backend` | Legt die Quelle der Standardwerte fest und sperrt die Auswahl. Eine ausdrückliche Backend-Vorgabe hat Vorrang. | Backend-Vorgabe, sonst Benutzerauswahl, sonst `local` |

Der automatische Login startet nur über Freigabe oder Talk bei unvollständigen Zugangsdaten. Die verwendete URL muss der Registry-URL entsprechen. Das normale Öffnen der Einstellungen startet keine Anmeldung und schließt den Dialog auch nicht automatisch.

<a id="versand-bei-dienstausfällen"></a>
<a id="sending-during-service-outages"></a>

### Versand bei Dienstausfällen

| Wert | Typ / Werte | Wirkung | Ohne Vorgabe |
| --- | --- | --- | --- |
| `SendPolicyFailureMode` | `REG_SZ`, `failopen` / `failclosed` | Bestimmt, ob der Versand bei nicht verfügbarer Nextcloud oder nicht verfügbarem Backend trotz zutreffender Signatur- oder Anhangsregel fortgesetzt werden darf. Benutzer können diese Vorgabe nicht ändern. | `failopen` |

Ein ungültiger Wert ergibt `failopen` und einen Konfigurationshinweis in Einstellungen und Log. Die Existenz aktiviert Enterprise Rollout, auch bei `failopen` oder ungültigem Inhalt. Registry-Änderungen benötigen einen Outlook-Neustart.

| Zustand der aktuellen Mail | `failopen` | `failclosed` |
| --- | --- | --- |
| Bestätigt keine zutreffende Signatur- oder Anhangsregel | Normal senden, ohne Policy-Warnung. | Normal senden, ohne Policy-Warnung. |
| Policy-Zustand unbekannt, etwa beim ersten Start mit nicht verfügbarem Server | Normal senden; keine Signatur- oder Anhangswarnung erfinden. | Erst beim Klick auf Senden blockieren, bis die zentralen Vorgaben geprüft werden können. Entwürfe bleiben speicherbar. |
| Zutreffende Regel bekannt, aber Nextcloud oder Backend nicht verfügbar | Mit nicht modalem Hinweis senden, wenn die Regel nicht erfüllt werden kann; ungeteilte Dateien bleiben normale Anhänge. Eine korrekt nutzbare zwischengespeicherte Signatur wird ohne unnötige Warnung eingefügt. | Versand blockieren, bis Verfügbarkeit bestätigt und die Regel erfüllt ist. Eine zwischengespeicherte Policy allein ersetzt die Prüfung nicht. |
| Bekannte zutreffende Regel nicht erfüllt und Dienstprüfung noch ausstehend | Blockieren, bis die Regel erfüllt oder ein Ausfall bestätigt ist. Eine korrekt eingefügte zwischengespeicherte Signatur erlaubt den Versand ohne Warnung. | Blockieren, bis die Prüfung erfolgreich und die Regel erfüllt ist. |
| Dienste verfügbar, aber vorgeschriebene Signatur nicht einfügbar oder Anhänge noch nicht geteilt | Versand mit Erklärung und Handlungshinweis blockieren. | Gleiches Verhalten. |

Regeln werden für die einzelne Mail ausgewertet, nicht für die gesamte Outlook-Sitzung. Eine Signatur gilt nur für den passenden Absender und aktivierten Nachrichtentyp. Eine Anhangsregel greift nur bei wirksamem **Immer über NC Connector** oder überschrittenem zentral verbindlichem Schwellwert. Ein lokales optionales Uploadangebot ist kein verbindlicher Schwellwert.

Beim Outlook-Start erscheint kein Versandrichtlinien-Dialog. Relevante Hinweise erscheinen beim Verfassen, Antworten, Weiterleiten, Hinzufügen betroffener Anhänge oder bei Freigabe-/Talk-Aktionen; blockierende Fehler erst beim Senden. Eine fehlgeschlagene Freigabe- oder Talk-Aktion erzeugt keine Vorgaben für spätere normale Mails. Abgelehnte Zugangsdaten, fehlende Berechtigungen, voller Speicher und lokale Verarbeitungsfehler sind keine Ausfallausnahmen. HTTP `429` ist eine vorübergehende Anfragenbegrenzung; eine vom Server genannte Wartezeit beachten. Eine blockierte Mail wird nie automatisch erneut gesendet.

<a id="verwaltete-transportsicherheit-tls"></a>
<a id="managed-transport-security-tls"></a>

### TLS

Ist mindestens ein TLS-Wert gesetzt, werden alle drei TLS-Einstellungen gesperrt. Fehlende Gruppenwerte erhalten die Standards aus der Tabelle. Ohne TLS-Vorgaben bleiben die gespeicherten lokalen Werte wirksam; ein neues Profil verwendet dieselben Standards.

| Wert | Typ / Werte | Wirkung | Standard |
| --- | --- | --- | --- |
| `TransportTlsUseSystemDefault` | `REG_DWORD`, `0` / `1` | Bei `1` bestimmt Windows die TLS-Version; die beiden Versionsschalter werden ignoriert. | `0` |
| `TransportTlsEnable12` | `REG_DWORD`, `0` / `1` | Erlaubt TLS 1.2 bei ausgeschaltetem Systemstandard. | `1` |
| `TransportTlsEnable13` | `REG_DWORD`, `0` / `1` | Erlaubt TLS 1.3, sofern Windows und Laufzeit es unterstützen. | `0` |

Nicht alle drei Werte auf `0` setzen: Dann sind keine Serververbindungen möglich. Ein einzelnes `TransportTlsEnable12=0` führt wegen der Gruppenstandards ebenfalls dazu.

<a id="logging-und-update-benachrichtigungen"></a>
<a id="logging-and-update-notifications"></a>
<a id="verwaltetes-logging"></a>
<a id="managed-logging"></a>
<a id="verwaltete-update-benachrichtigungen"></a>
<a id="managed-update-notifications"></a>

### Logging und Update-Hinweise

Die beiden Logging-Werte gehören zusammen. Ist einer gesetzt, werden beide Felder gesperrt und der fehlende Wert erhält seinen Standard. Ohne Logging-Vorgaben gelten die lokalen Einstellungen, im neuen Profil Debug aus und Anonymisierung an.

| Wert | Typ / Werte | Wirkung | Standard |
| --- | --- | --- | --- |
| `DebugLoggingEnabled` | `REG_DWORD`, `0` / `1` | Schaltet ausführliche Protokolle ein. Fehler werden auch ohne Debug protokolliert. | `0` |
| `LogAnonymizationEnabled` | `REG_DWORD`, `0` / `1` | Anonymisiert personenbezogene Angaben im Log. Passwörter und Tokens werden unabhängig davon maskiert. | `1` |
| `UpdateNotifyEnabled` | `REG_DWORD`, `0` / `1` | Steuert und sperrt die Update-Benachrichtigung; unabhängig von der Logging-Gruppe. | Lokale Auswahl; im neuen Profil `0` |

`UpdateNotifyEnabled=0` unterbindet nur den Hinweis, nicht die tägliche Versionsabfrage oder **Jetzt prüfen**. Das Add-in installiert keine Updates selbst.

<a id="ifb"></a>
<a id="verwaltetes-internet-freebusy-ifb"></a>
<a id="managed-internet-freebusy-ifb"></a>

### Internet Free/Busy

Ist mindestens ein IFB-Wert gesetzt, werden alle vier Werte einschließlich der Adressbuch-Cache-Dauer unter **Erweitert** vorgegeben und gesperrt. Fehlende Gruppenwerte erhalten die Standards aus der Tabelle. Ohne Vorgaben gelten die lokalen Einstellungen.

| Wert | Typ / Werte | Wirkung | Standard |
| --- | --- | --- | --- |
| `IfbEnabled` | `REG_DWORD`, `0` / `1` | Aktiviert die Anzeige von Nextcloud-Verfügbarkeiten in Outlook. | `0` |
| `IfbDays` | `REG_DWORD`, `10`, `30`, `60`, `90` | Abgefragter Zeitraum in Tagen | `30` |
| `IfbCacheHours` | `REG_DWORD`, `1`–`24` | Cache-Dauer des Systemadressbuchs in Stunden; gilt auch für Talk. | `24` |
| `IfbPort` | `REG_DWORD`, `1024`–`49151` | Lokaler Port; für andere Ports als `7777` ist eine eigene URL-Reservierung nötig. | `7777` |

Nur einen Port vorzugeben aktiviert IFB nicht. Beim ersten Einrichten ohne Registry-Vorgabe kann der Anmeldedialog IFB vorwählen; die gewünschte Einstellung vor dem Speichern kontrollieren.

<a id="optionales-nc-connector-backend"></a>
<a id="optional-nc-connector-backend"></a>
<a id="policy-rollout"></a>

## Backend-Vorgaben und Signaturen

Das Backend verwaltet Freigabe- und Talk-Einstellungen, Vorlagen, Signaturen und separate Passwortzustellung. Richtlinien gelten für Benutzer mit gültig zugeteiltem Seat. Administratorrechte ersetzen diese Zuweisung nicht.

Lassen Sie **Im Add-on veränderbar** eingeschaltet, wenn Benutzer eine Vorgabe anpassen dürfen. Schalten Sie es aus, wenn der Wert verbindlich gelten soll. Bei Ablaufzeiten mindestens einen Tag wählen. Eine Anhangsschwelle wird über ihren Schalter deaktiviert, nicht durch Eingabe von null; eingeschaltet sind 1–10240 MB möglich.

<a id="default-values-source"></a>

### Quelle der Standardwerte

Unter **Gruppeneinstellungen → Standardeinstellungen → Allgemein** bestimmt ein Nextcloud-Administrator, woher Outlook die Startwerte für neue Aktionen übernimmt. Diese Einstellung kann nicht an Gruppenadministratoren delegiert werden.

Sie gilt gemeinsam für Freigaben, Talk, Anhangsautomatisierung, die Sprache erzeugter Texte und die Signaturschalter. Die Signaturvorlage selbst kommt immer aus dem Backend.

| Backend-Einstellung | Ergebnis in Outlook |
| --- | --- |
| **Lokal** oder **Backend**, nicht im Add-on veränderbar | Die Backend-Auswahl gilt; auch eine abweichende Registry-Vorgabe wird überstimmt. |
| **Lokal** oder **Backend**, im Add-on veränderbar | Benutzer können unter **Erweitert → Quelle der Standardwerte** wählen. Bis dahin gilt die Backend-Auswahl. |
| **Keine Vorgabe**; ebenso bei älterem Backend ohne diese Einstellung | `DefaultsSource` aus der Registry gilt. Fehlt auch dieser Wert, gilt die Benutzerauswahl, sonst **Lokal**. |

Bei **Lokal** werden gespeicherte lokale Werte verwendet. Fehlt eine lokale Auswahl, folgen Backend- und anschließend Produktstandards. Bei **Backend** ist die Reihenfolge umgekehrt: Backend, lokale Auswahl, Produktstandard.

Einzeln gesperrte Backend-Werte gelten unabhängig von dieser Auswahl. Editierbare Felder können im Assistenten für die aktuelle Aktion weiterhin angepasst werden.

Mit Quelle **Backend** sind die Einstellungstabs **Freigabe**, **Talk-Link** und **Signatur** ausgegraut. Die bisher gespeicherten lokalen Werte bleiben erhalten. Ohne gültigen Seat ist die Quellenauswahl nicht verfügbar. Ältere Add-in-Versionen übernehmen die neue Quellenvorgabe nicht.

Nach einer Backend-Änderung die Einstellungen oder den betroffenen Assistenten erneut öffnen. Um die Quellenauswahl freizugeben, im Backend **Im Add-on veränderbar** aktivieren oder **Keine Vorgabe** wählen und zusätzlich die Registry-Vorgabe entfernen.

<a id="verwaltete-signaturen-einrichten"></a>
<a id="configure-managed-signatures"></a>

### Signaturen

1. Im Backend dem Benutzer eine Signaturvorlage zuweisen.
2. Prüfen, ob die zugewiesene E-Mail-Adresse zur tatsächlichen **Von**-Adresse in Outlook passt. Das gilt auch für freigegebene Postfächer und delegierte Absender.
3. Die Signatur für neue Mails, Antworten und Weiterleitungen nach Bedarf aktivieren.
4. Eine Mail mit der vorgesehenen Absenderadresse öffnen und die Darstellung kontrollieren.

Falls doppelte Signaturen erscheinen, zusätzlich eingerichtete Outlook-Signaturen prüfen. Die Sprache des Freigabeblocks wird unter **Freigabe**, die Sprache des Talk-Beschreibungstextes unter **Talk-Link** eingestellt.

<a id="vorlagen-erstellen"></a>
<a id="template-authoring"></a>

### Eigene Vorlagen

Outlook stellt einfaches HTML mit Tabellen und Inline-Styles am zuverlässigsten dar. Verwenden Sie absolute HTTPS-Links; vermeiden Sie Weblayouts mit Flexbox oder Grid. Eigene Vorlagen vor der Verteilung in Outlook kontrollieren.

Bei Freigabevorlagen passen sich `{LINK_INTRO}` und `{LINK_LABEL}` an das gewählte Linkziel an. Beschriftung und Platzhalter eines optionalen Feldes, etwa `{PASSWORD}`, im selben HTML-Absatz oder derselben Tabellenzeile unterbringen. So bleibt bei leerem Wert keine einzelne Beschriftung stehen.

<a id="funktionsbetrieb"></a>
<a id="feature-operation"></a>

## Hinweise für den Betrieb

<a id="freigaben-und-uploads"></a>
<a id="sharing-and-uploads"></a>

### Dateien und Anhänge

**Meine Nextcloud** kopiert ausgewählte Dateien und Ordner in einen neuen Freigabeordner. Die Originale bleiben erhalten. Dafür werden Lesezugriff auf die Quelle, Schreibzugriff auf das Ziel und genügend Speicher benötigt. Symbolische Links und Junctions werden nicht unterstützt.

<a id="anhangsautomatisierung"></a>
<a id="attachment-automation"></a>

Die Anhangsautomatisierung wird unter **Freigabe → Anhänge** eingerichtet. Sie kann Anhänge immer über NC Connector senden oder die Nutzung ab einem Schwellwert anbieten. Das Linkziel ist wahlweise ZIP-Download oder Freigabeseite; ohne Vorgabe gilt ZIP-Download. Manuelle Freigaben verlinken auf die Freigabeseite.

Outlook oder Exchange kann einen Anhang abweisen, bevor NC Connector ihn erhält. In diesem Fall die Datei direkt über **Nextcloud-Freigabe einfügen** auswählen. Ein zentral gesperrter Größen-Schwellwert macht das Teilen bei Überschreitung verbindlich; ein lokales Uploadangebot nicht. Bei verfügbaren Diensten verhindert eine verbindliche Anhangsregel den Versand, solange eine betroffene Datei noch normal an der Mail hängt. Bei Dienstausfällen gilt [SendPolicyFailureMode](#versand-bei-dienstausfällen).

Originalanhänge bleiben erhalten, bis ihre Dateien erfolgreich geteilt wurden und der Freigabelink in die Mail eingefügt wurde. Abbruch des Assistenten sowie fehlgeschlagener Upload oder Einfügung erhalten die Anhänge. Werden Dateien aus der Assistentenauswahl entfernt, werden nur die tatsächlich geteilten Originale gelöscht; spätere Ergänzungen und andere gleichnamige Dateien bleiben erhalten.

<a id="ungesendete-mail-und-freigabebereinigung"></a>
<a id="unsent-mail-and-share-cleanup"></a>

Beim Verwerfen einer noch nicht gespeicherten Mail versucht NC Connector, ihren neu angelegten Freigabeordner zu entfernen. Bereits gespeicherte Entwürfe behalten ihre Freigaben. Wird ein solcher Entwurf später gelöscht, müssen nicht mehr benötigte Freigaben in Nextcloud manuell entfernt werden.

<a id="separate-password-delivery"></a>

### Separate Passwortzustellung

Die Passwortzustellung wird beim Klick auf **Senden** angestoßen und verwendet dasselbe Outlook-Konto wie die Hauptmail. Beachten Sie diese Grenzen:

- **Mail bis zum Versand geöffnet lassen.** Nach Schließen und erneutem Öffnen eines Entwurfs oder einem Outlook-Neustart muss die Freigabe neu erstellt werden. Ein noch sichtbarer Freigabeblock stellt den Passwortversand nicht wieder her. Das gilt auch für Vorlagen mit einem vorhandenen Freigabeblock. Speichern und AutoSave bei weiterhin geöffneter Mail sind möglich.
- **Bei verzögertem Versand kann das Passwort zuerst ankommen.** Die Passwort-Mail wartet nicht auf die Hauptmail im Postausgang. Sie kann auch schon versendet sein, wenn Outlook die Hauptmail anschließend ablehnt.
- Scheitert der Versand eindeutig, öffnet sich eine vorbereitete Passwort-Mail zum manuellen Senden. Bei unklarem Ergebnis wird nicht automatisch erneut versendet, um Duplikate zu vermeiden.
- Mit Secrets erhält jeder Empfänger einen eigenen einmaligen Link. **Scheitert die Erstellung, wird das Passwort mit Warnung als Klartext in der separaten Mail versendet.**

Diese Einschränkungen bei der Einführung der Funktion an die Benutzer weitergeben.

<a id="talk-raum-lebenszyklus"></a>
<a id="talk-room-lifecycle"></a>

### Talk-Räume und Termine

Das Löschen eines gespeicherten Outlook-Termins löscht dessen Talk-Raum nur, wenn dies unter **Talk-Link** ausdrücklich aktiviert wurde. Standardmäßig ist es ausgeschaltet. Die Löschung betrifft den Raum für alle Teilnehmer.

Bei Terminserien bleibt der gemeinsame Raum beim Löschen einzelner Vorkommen bestehen. Ein nur in den Termin kopierter Talk-Link löst keine Raumlöschung aus. Bei vorübergehenden Verbindungsfehlern werden vorgemerkte Raumlöschungen später wiederholt, auch nach einem Outlook-Neustart.

<a id="internet-freebusy-gateway-ifb"></a>
<a id="zweck-und-aktivierung"></a>
<a id="purpose-and-activation"></a>

## Internet Free/Busy einrichten

IFB zeigt Nextcloud-Verfügbarkeiten im Outlook-Terminplanungs-Assistenten. Es benötigt das Systemadressbuch, eine Anmeldung und ein laufendes Outlook.

1. Das [Systemadressbuch](#benutzersuche-oder-verfügbarkeiten-fehlen) für den Benutzer bereitstellen.
2. Unter **Einstellungen → IFB** aktivieren und Zeitraum sowie Port festlegen; die Cache-Dauer steht unter **Erweitert**. Alternativ die Registry-Gruppe verwenden.
3. Speichern und Outlook neu starten.
4. In einem Termin eine bekannte Nextcloud-E-Mail-Adresse hinzufügen und die Verfügbarkeit im Terminplanungs-Assistenten prüfen.

<a id="standard-reservierung-prüfen"></a>
<a id="verify-the-default-reservation"></a>

Die MSI richtet den Standardport `7777` ein. Bei laufendem Outlook mit aktiviertem IFB und abgeschlossener Anmeldung lässt sich die lokale Verbindung so prüfen:

```powershell
netsh http show urlacl url=http://127.0.0.1:7777/nc-ifb/
Test-NetConnection 127.0.0.1 -Port 7777
```

IFB ist nur auf dem eigenen Rechner erreichbar. Die für Outlook hinterlegte URL enthält einen automatisch verwalteten geheimen Pfad. Ein direkter Browseraufruf ohne diesen Pfad liefert deshalb `404` und ist kein geeigneter Funktionstest.

<a id="eigener-ifb-port"></a>
<a id="custom-ifb-port"></a>

### Anderen Port verwenden

Für einen anderen Port zwischen `1024` und `49151` muss ein Administrator die URL-Reservierung in einer erhöhten PowerShell anlegen. Beispiel für Port `8888`:

```powershell
netsh http add urlacl url=http://127.0.0.1:8888/nc-ifb/ sddl="D:(A;;GX;;;AU)"
```

Danach denselben Port im Add-in oder per `IfbPort` setzen und Outlook neu starten. Die Berechtigung gilt für authentifizierte Windows-Benutzer; nicht durch `Everyone` ersetzen.

Wird der eigene Port nicht mehr genutzt, die selbst angelegte Reservierung entfernen:

```powershell
netsh http delete urlacl url=http://127.0.0.1:8888/nc-ifb/
```

## Update, Sicherung und Deinstallation

<a id="upgrade-oder-rückkehr-auf-eine-ältere-version"></a>
<a id="upgrade-or-return-to-an-older-version"></a>

### Update oder Rückkehr zur vorherigen Version

1. Offene Arbeit speichern und Outlook möglichst in allen Windows-Sitzungen schließen. Ist es noch geöffnet, kann Windows Installer das geordnete Schließen anbieten; eine unbeaufsichtigte Installation versucht dies automatisch.
2. Die [Profildaten sichern](#profildaten-sichern-und-wiederherstellen).
3. Die gewünschte MSI über die vorhandene Installation installieren. Die vorherige MSI für einen Rückweg aufbewahren.
4. Outlook starten und Anmeldung sowie die verwendeten Funktionen prüfen.

Läuft Outlook danach weiter, etwa in einer anderen Windows-Sitzung, stoppt das Setup vor dem Entfernen der alten Version und der IFB-Bereinigung. Outlook manuell schließen und das Setup erneut starten. Es erzwingt kein Beenden und garantiert keinen automatischen Neustart von Outlook.

Benutzereinstellungen bleiben erhalten. Für die Rückkehr zu einer älteren Version deren MSI auf dieselbe Weise installieren; bei Bedarf die zugehörige Sicherung zurückspielen. Windows-Installer-Rollback darf durch die Softwareverteilung nicht deaktiviert sein.

<a id="profildaten"></a>
<a id="profile-data"></a>
<a id="sicherung-und-wiederherstellung"></a>
<a id="backup-and-restore"></a>

### Profildaten sichern und wiederherstellen

Die Daten liegen unter `%LOCALAPPDATA%\NC4OL\`:

| Dateien | Inhalt |
| --- | --- |
| `settings_<OutlookProfile>.xml` und `.xml.bak` | Einstellungen und vorherige gültige Fassung; ohne ermittelbaren Profilnamen heißt die Datei `settings_default.xml`. |
| `talk-room-lifecycle-*.dat*` | Noch ausstehende Talk-Raum-Löschungen |
| `ifb-registry-state-*.dat*` | Für IFB gesicherte vorherige Registry-Werte |
| `addin-runtime.log_YYYYMMDD` | Betriebsprotokolle |

Für eine Sicherung Outlook schließen und die Einstellungs- sowie Zustandsdateien kopieren. Windows-Benutzer und Outlook-Profil zuordnen. Zur Wiederherstellung Outlook ebenfalls schließen und die passenden Dateien zurückspielen.

App-Passwort und Zustandsdateien sind für den Windows-Benutzer geschützt. Sie sind nicht als Vorlage für andere Konten oder Computer geeignet. Ist das Passwort nach einer Wiederherstellung nicht lesbar, erneut anmelden. Logs und Adressbuch-Cache müssen nicht wiederhergestellt werden.

Falls Sie statt Registry-Werten eine XML-Vorlage verteilen, nur noch nicht vorhandene Profildateien anlegen und `AppPasswordProtected` vorher entfernen. Jeden Benutzer selbst anmelden lassen. Eine XML-Vorlage allein aktiviert Enterprise Rollout nicht.

<a id="deinstallation"></a>
<a id="uninstall"></a>

### Deinstallation und Neuinstallation

Outlook in allen Sitzungen schließen und das Add-in über **Windows-Einstellungen → Apps** deinstallieren. Die MSI entfernt Programmdateien, Add-in-Registrierung, die Standard-IFB-Reservierung und eigene Free/Busy-Restwerte. Auch Installation, Upgrade und Reparatur bereinigen solche alten IFB-Werte, selbst wenn der Datenordner zuvor gelöscht wurde.

Für die unbeaufsichtigte Deinstallation mit Protokoll:

```powershell
msiexec.exe /x "NCConnectorForOutlook-3.4.3.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-uninstall.log"
```

Benutzereinstellungen, Caches und Logs unter `%LOCALAPPDATA%\NC4OL\` bleiben erhalten. Deshalb kann das Add-in nach einer Neuinstallation noch angemeldet sein. Für einen vollständigen Neustart der Einrichtung den Ordner bei geschlossenem Outlook nach einer Sicherung umbenennen. Das setzt die Konfiguration aller darin gespeicherten Outlook-Profile zurück; ausstehende Talk-Raum-Löschungen werden dann nicht mehr ausgeführt.

Soll ein früherer Free/Busy-Anbieter wiederhergestellt werden, IFB vor der Deinstallation deaktivieren, solange die gespeicherten vorherigen Werte noch vorhanden sind. Aus einem gelöschten Datenordner lassen sich diese nicht rekonstruieren. Aktuelle Fremdwerte und administrative Policies werden nicht entfernt. Eigene Port-Reservierungen müssen Sie [separat zurücknehmen](#anderen-port-verwenden).

<a id="runbooks-zur-störungsbehebung"></a>
<a id="troubleshooting-runbooks"></a>

## Fehlerbehandlung

### Setup meldet weiterhin geöffnetes Outlook

1. Ungespeicherte Arbeit in Outlook sichern.
2. Outlook in allen Windows-Sitzungen schließen und prüfen, ob noch ein `OUTLOOK.EXE`-Prozess läuft.
3. Eine verbliebene Instanz geordnet schließen und das Setup erneut starten. Outlook nicht zwangsweise beenden, solange ungespeicherte Inhalte vorhanden sein könnten.

Windows Installer kann das Schließen anbieten und versucht es bei unbeaufsichtigtem Setup automatisch. Bleibt Outlook trotzdem aktiv, blockiert NC Connector den Austausch der alten Version; ein automatischer Neustart von Outlook ist nicht zugesichert.

<a id="add-in-wird-nicht-geladen"></a>
<a id="add-in-does-not-load"></a>
<a id="installationspfade-und-registrierungsprüfung"></a>
<a id="installed-paths-and-registration-checks"></a>

### Add-in oder Haupttab fehlt

1. Prüfen, ob Outlook classic verwendet wird.
2. Fehlt nur der Haupttab, `ShowMainRibbonTab` prüfen. Bei `0` ist das beabsichtigt; Freigabe und Talk müssen weiterhin in Mails und Terminen erreichbar sein.
3. Unter **Datei → Optionen → Add-Ins** sowohl **COM-Add-Ins** als auch **Deaktivierte Elemente** prüfen. Das Add-in heißt `NcTalkOutlook.AddIn`.
4. Prüfen, ob `C:\Program Files\NC4OL\NcTalkOutlookAddIn.dll` vorhanden ist und `LoadBehavior` im passenden Pfad den Wert `3` hat:

```text
64-Bit-Outlook: HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
32-Bit-Outlook: HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

Fehlen Dateien oder Registrierung, Outlook schließen und die MSI reparieren. Bei weiterem Fehler Installationslog und Windows-Ereignisanzeige für Outlook/.NET auswerten.

<a id="verbindungs--oder-tls-test-schlägt-fehl"></a>
<a id="connection-or-tls-test-fails"></a>

### Anmeldung oder Verbindung schlägt fehl

1. Die konfigurierte Nextcloud-URL am Arbeitsplatz im Browser öffnen. URL, DNS, Systemzeit, Zertifikat und Proxy prüfen.
2. Bei einer TLS-Konfigurationsmeldung zuerst die Registry-Vorgaben korrigieren, anschließend Outlook neu starten. Andere Zugangsdaten lösen einen TLS-Fehler nicht.
3. Bei HTTP `401` erneut anmelden; bei `403` die Berechtigungen und vorgeschaltete Zugriffssperren prüfen.
4. Den Verbindungstest erneut ausführen. Falls nur ein Arbeitsplatz betroffen ist, dessen Proxy-, Zertifikats- und Endpoint-Security-Konfiguration mit einem funktionierenden Arbeitsplatz vergleichen.

Beim Klick auf Freigabe oder Talk wird die Nextcloud-Verbindung vor dem Assistenten erneut geprüft. Ein nicht erreichbarer Server zeigt in beiden Aktionen denselben Hinweis mit dem Angebot, die Einstellungen zu öffnen; eine frühere erfolgreiche Verbindung überspringt diese Prüfung nicht. Nach dem Speichern geprüfter Zugangsdaten wird vor der Fortsetzung nochmals geprüft. Einrichtungsabbruch oder Schließen der ursprünglichen Mail/des Termins beendet die Aktion. Die Anhangsautomatisierung öffnet keine Anmeldung automatisch; fehlgeschlagenes Teilen erhält die Originalanhänge. Für den normalen Versand gilt weiterhin der konfigurierte `SendPolicyFailureMode`.

Keine Zertifikatsprüfung abschalten und nicht versuchsweise computerweite TLS-Einstellungen ändern.

### Eine Registry-Vorgabe wird nicht wie erwartet angewendet

Zuerst Pfad, Registry-Ansicht, Datentyp und Schreibweise anhand der Tabellen prüfen. Danach Outlook neu starten. Ein höherrangiger ungültiger Eintrag wird nicht durch einen gültigen Eintrag weiter unten ersetzt.

| Betroffene Vorgabe | Verhalten bei ungültigem Wert / Abhilfe |
| --- | --- |
| URL, URL-Sperre, Haupttab | Ungültige URL wird nicht verwendet; ungültige Sperre sperrt die URL nicht; ungültiger Ribbonwert lässt den Haupttab sichtbar. Wert im wirksamen Pfad korrigieren. |
| Anmeldeart | Auswahl bleibt auf Login Flow gesperrt, der automatische Start entfällt. `LoginFlow` oder `Manual` als `REG_SZ` setzen. |
| Standardwertquelle | Ohne vorrangige Backend-Vorgabe wird gesperrtes `local` verwendet. `local` oder `backend` als `REG_SZ` setzen. |
| Versand bei Ausfällen | Ein ungültiges `SendPolicyFailureMode` ergibt `failopen` mit Konfigurationshinweis. `failopen` oder `failclosed` als `REG_SZ` setzen. |
| TLS | Ungültige Werte oder drei ausgeschaltete Schalter verhindern Serveranfragen, auch Anmeldung und Update-Prüfung. Werte beziehungsweise gewählten TLS-Modus korrigieren. |
| Logging | Das betroffene Feld verwendet seinen Standard; die Verbindung wird nicht blockiert. |
| Update-Hinweise | Hinweise bleiben ausgeschaltet. Die Versionsabfrage läuft weiter. |
| IFB | IFB wird deaktiviert. Bei ungültiger Cache-Dauer gelten 24 Stunden; andere gültige Cache-Vorgaben bleiben erhalten. |

<a id="backend-policy-oder-seat-wird-nicht-angewendet"></a>
<a id="backend-policy-or-seat-is-not-applied"></a>
<a id="voraussetzungen-und-betriebszustände"></a>
<a id="prerequisites-and-operating-states"></a>
<a id="lizenzhinweise-in-outlook"></a>
<a id="license-notices-in-outlook"></a>

### Backend-Hinweis, fehlender Seat oder unerwartete Einstellungen

| Anzeige / Problem | Was prüfen? |
| --- | --- |
| Backend fehlt | Nextcloud-App `ncc_backend_4mc` installieren beziehungsweise aktivieren. Zugriff auf `/apps/ncc_backend_4mc/api/v1/status` prüfen. |
| Seat fehlt, ist pausiert oder ungültig | Im Backend die Zuweisung des betroffenen Benutzers und den Lizenzstatus prüfen. Bei pausierter Zuweisung Seat-Anzahl oder Lizenzumfang anpassen. |
| Prüfung nicht möglich | Verbindung zwischen Outlook und Nextcloud prüfen; das ist für sich genommen keine Lizenzablehnung. |
| Lizenzsynchronisierung fehlgeschlagen | Im Backend die Verbindung zum Lizenzserver und den letzten erfolgreichen Abgleich prüfen. |
| Nachfrist oder Aktivierungsproblem | Lizenzübersicht im Backend öffnen und den dort genannten Handlungsbedarf beheben. |
| Wert oder Einstellungstab ist gesperrt | Registry-Vorgaben, Quelle der Standardwerte und **Im Add-on veränderbar** im Backend prüfen. |
| Alter oder unerwarteter Startwert | Quelle der Standardwerte und lokale Auswahl kontrollieren; nach Änderungen den Dialog erneut öffnen. |

Bei überschrittenem Lizenzumfang zuerst die konkret pausierten Zuweisungen prüfen. Weiterhin aktive Seats müssen nicht neu zugewiesen werden.

Ohne Enterprise Rollout bleiben Freigaben und Talk bei fehlendem gültigem Seat mit lokalen Einstellungen nutzbar; zentrale Signaturen und separate Passwortzustellung nicht. Im Enterprise Rollout müssen Backend und Seat-Zuweisung für die Nutzung von NC Connector wieder verfügbar sein. Ein kurzzeitiger Verbindungsfehler kann mit dem zuletzt bestätigten Status überbrückt werden.

<a id="benutzersuche-oder-verfügbarkeiten-fehlen"></a>
<a id="system-address-book"></a>
<a id="benutzersuche-oder-moderatorauswahl-ist-deaktiviert"></a>
<a id="user-search-or-moderator-selection-is-disabled"></a>

### Systemadressbuch

Fehlen Benutzer in der Suche, bleibt die Moderatorauswahl gesperrt oder fehlen Nextcloud-Verfügbarkeiten, zuerst den Adressbuchzugriff prüfen:

1. In Nextcloud unter **Verwaltungseinstellungen → Groupware → Systemadressbuch** die Bereitstellung einschalten. Auch die Freigabe- und Autovervollständigungsregeln unter **Teilen** für den betroffenen Benutzer prüfen.
2. Das Systemadressbuch auf dem Nextcloud-Server neu aufbauen. Im Nextcloud-Verzeichnis ausführen; HTTP-Benutzer und Aufruf bei Containerinstallationen anpassen:

```bash
sudo -E -u www-data php occ dav:sync-system-addressbook
```

3. Mit dem betroffenen Konto den Export prüfen:

```text
https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system?export
```

4. In Outlook die Verbindung erneut prüfen und die Suche wiederholen. Bei IFB zusätzlich die lokale Verbindung prüfen, siehe unten.

Erwartet wird ein vCard-Adressbuch, keine HTML-Fehlerseite. Bei `401` die Anmeldung, bei `403` die Zugriffsrechte prüfen. Benutzer ohne E-Mail-Adresse können gesucht werden, lassen sich aber nicht über eine E-Mail-Adresse zuordnen.

Zeigt Nextcloud das Systemadressbuch als aktiviert an, der Export bleibt aber nicht erreichbar, die gespeicherte Servereinstellung kontrollieren:

```bash
sudo -E -u www-data php occ config:app:get dav system_addressbook_exposed
```

Soll das Systemadressbuch für Clients bereitstehen, muss der Wert `yes` sein. Eine abweichende Einstellung gezielt korrigieren und das Adressbuch erneut aufbauen:

```bash
sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"
sudo -E -u www-data php occ dav:sync-system-addressbook
```

<a id="nextcloud-pretty-urls"></a>
<a id="kurztest"></a>
<a id="quick-test"></a>
<a id="pretty-url-oder-talk-link-liefert-404"></a>
<a id="pretty-url-or-talk-link-returns-404"></a>

### Talk-Link liefert im Browser 404

`https://<cloud>/login` mit `https://<cloud>/index.php/login` vergleichen. Funktioniert nur die zweite Adresse, die Rewrite-Regeln des Webservers oder Reverse Proxys nach der offiziellen Nextcloud-Konfiguration korrigieren. Bei Installationen unter `/nextcloud` den Unterpfad in beiden URLs beibehalten. `/index.php` nicht als Ausweichlösung in die Add-in-Serveradresse aufnehmen.

Danach den Login und einen neu erzeugten Talk-Link von einem Arbeitsplatz aus erneut öffnen. Serverkonfiguration vor Änderungen sichern und vor dem Neuladen mit dem jeweiligen Webserver prüfen.

<a id="nginx"></a>

Für **Nginx** die [offizielle Nextcloud-Konfiguration](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html) verwenden, statt einzelne Rewrite-Regeln aus fremden Konfigurationen zu übernehmen.

<a id="apache"></a>

Bei **Apache** Rewrite-Module, `AllowOverride` und Rewrite Base prüfen; nach Änderungen Nextclouds `.htaccess` neu erzeugen. Die Schritte stehen in der [Nextcloud-Anleitung für Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls).

<a id="upload-bleibt-bei-null-oder-schlägt-fehl"></a>
<a id="upload-remains-at-zero-or-fails"></a>
<a id="anhangsautomatisierung-startet-nicht"></a>
<a id="attachment-automation-does-not-start"></a>

### Upload oder Anhangsautomatisierung schlägt fehl

1. Prüfen, ob die Anzeige noch beim lokalen Dateiscan oder bereits beim Upload steht. Große Ordner brauchen vor dem Upload Zeit zum Einlesen.
2. Lesbarkeit der Quelldateien, Schreibrechte und freien Nextcloud-Speicher prüfen. HTTP `507` weist auf fehlenden Speicher hin.
3. Bei Fehlern nur mit großen Dateien die Größenlimits und Timeouts des Reverse Proxys prüfen.
4. Bei ausbleibender Automatisierung die Einstellungen unter **Freigabe → Anhänge** kontrollieren. Hat Outlook den Anhang bereits abgelehnt, die Datei direkt im Freigabe-Assistenten auswählen.
5. Mit einer kleinen Datei eingrenzen und das zugehörige Zeitfenster im `FILELINK`-Log auswerten.

Wird der Versand wegen ungeprüfter zentraler Vorgaben blockiert, `SendPolicyFailureMode` prüfen, die Dienste wieder verfügbar machen, die Mail geöffnet lassen und nach erfolgreicher Prüfung erneut senden. Der Hinweis bedeutet nicht, dass ein Upload fehlgeschlagen ist. Originalanhänge bleiben nach abgebrochenem oder fehlgeschlagenem Teilen erhalten; vor einem neuen Versuch kontrollieren.

<a id="verwaltete-signatur-fehlt-oder-steht-falsch"></a>
<a id="managed-signature-is-missing-or-misplaced"></a>

### Signatur fehlt oder verhindert das Senden

Seat, Signaturzuweisung, tatsächliche **Von**-Adresse und die Schalter für neue Mail, Antwort oder Weiterleitung prüfen. Bei doppelter Signatur zusätzlich Outlooks eigene Signaturkonfiguration kontrollieren.

Bei einem Ausfall richtet sich der Versand nach [SendPolicyFailureMode](#versand-bei-dienstausfällen). Eine Warnung über die fehlende Signatur erscheint nur bei einer für diese Mail bekannt zutreffenden Signatur, nicht allein wegen einer fehlgeschlagenen Policy-Abfrage. Bei `failclosed` blockiert auch ein unbekannter Policy-Zustand das Senden bis zur Prüfung. Kann eine vorgeschriebene Signatur trotz verfügbarer Dienste nicht eingefügt werden, bleibt der Versand in beiden Modi gesperrt. Die Mail geöffnet lassen und die Fehlermeldung mit passendem Logzeitraum an den Support geben.

<a id="einstellungen-können-nicht-geladen-oder-gespeichert-werden"></a>
<a id="settings-cannot-be-loaded-or-saved"></a>

### Einstellungen lassen sich nicht laden oder speichern

1. Outlook schließen und die betroffenen `settings_*.xml`- und `.xml.bak`-Dateien sichern.
2. Outlook neu starten. Das Add-in versucht bei beschädigter Hauptdatei die vorherige gültige Fassung zu laden.
3. Fehlt nur das App-Passwort, erneut anmelden. Sind beide Dateien beschädigt, Einstellungen neu eintragen und ausdrücklich speichern.
4. Schlägt das Speichern weiter fehl, freien Speicherplatz, Zugriffsrechte und Endpoint-Security prüfen; anschließend die `CORE`-Einträge auswerten.

<a id="ifb-antwortet-nicht"></a>
<a id="ifb-does-not-respond"></a>

### IFB startet nicht oder zeigt keine Daten

1. Aktivierung, Anmeldung und Registry-Gruppe prüfen. Bei zentral verwalteter Installation zusätzlich den Backend-Status prüfen.
2. URL-Reservierung und Port wie unter [IFB einrichten](#internet-freebusy-einrichten) prüfen. Für einen eigenen Port reicht die Portangabe allein nicht aus.
3. Bei Portkonflikt den belegenden Prozess ermitteln, beispielsweise für den Standardport:

```powershell
netstat -ano | Select-String ":7777"
```

4. Startet IFB, liefert aber keine Verfügbarkeit, Systemadressbuch und verwendete E-Mail-Adresse kontrollieren. Die Anfrage und das CalDAV-Ergebnis stehen im `IFB`-Log.
5. Nach Neuinstallation bei der Meldung **An unowned NC Connector IFB value already exists** Outlook in allen Sitzungen schließen und die aktuelle MSI reparieren. Die Reparatur entfernt alte eigene IFB-Einträge; fremde Anbieterwerte nicht manuell löschen.

<a id="monitoring-und-support"></a>
<a id="monitoring-and-support"></a>
<a id="logs"></a>

## Logs und Support

Unter **Einstellungen → Debuggen** das ausführliche Logging einschalten und **Logs anonymisieren** aktiviert lassen. Bei ausgeblendetem Haupttab kann die Administration die beiden Logging-Registry-Werte verwenden.

Die Dateien liegen unter `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`. Fehler werden auch ohne Debug-Logging erfasst. Die Aufbewahrung ist auf sieben Tagesdateien begrenzt; Dateien älter als 30 Tage werden zusätzlich entfernt.

<a id="supportpaket"></a>
<a id="support-package"></a>

Für eine Supportanfrage:

1. Uhrzeit, Add-in-Version, Outlook-Version und Bitness sowie Nextcloud- und relevante App-Versionen notieren.
2. Das Problem mit eingeschaltetem Debug-Logging einmal nachvollziehen.
3. Bedienschritt, Fehlermeldung und passendes Logzeitfenster zusammenstellen. Bei Installationsproblemen das MSI-Log ergänzen.
4. Vor der Weitergabe auf Zugangsdaten, private Links, Empfänger- und Kundendaten prüfen – auch bei eingeschalteter Anonymisierung.

`CORE` betrifft Einstellungen und Start, `API` die Serverkommunikation, `FILELINK` Freigaben, `TALK` Besprechungen und `IFB` die Verfügbarkeitsanzeige. Debug-Logging nach der Diagnose wieder auf den vorgesehenen Betriebswert zurückstellen.

<a id="sicherheit-und-datenverarbeitung"></a>
<a id="security-and-data-handling"></a>

Zugangsdaten und vollständige Freigabe- oder Secret-Links vertraulich behandeln. Backend-Vorlagen und Registry-Vorgaben nur durch berechtigte Administratoren ändern lassen. Die tägliche Versionsabfrage übermittelt Produkt, Version, Kanal und einen wechselnden anonymen Client-Hash, keine Nextcloud-Zugangsdaten oder Nachrichteninhalte.

Weiterführend:

- [Supportformular](https://nc-connector.de/support/)
- [Nextcloud-Systemadressbuch](https://docs.nextcloud.com/server/32/admin_manual/groupware/contacts.html#system-address-book)
- [Nextcloud-Konfiguration für Nginx](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html)
- [Nextcloud-Konfiguration für Apache und Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls)
