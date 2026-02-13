<div align="center" style="background:#0082C9; padding:1px 0;"><img src="assets/header-solid-blue-1920x480.png" alt="Addon" height="80"></div>

[English](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.md) | [Deutsch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.de.md)
[Admin Guide](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/ADMIN.md) | [Development Guide](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/DEVELOPMENT.md) | [Translations](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md)

# NC Connector for Outlook

NC Connector for Outlook verbindet Outlook nahtlos mit Ihrer Nextcloud. Das Add-in automatisiert Talk-Raeume fuer Termine, stellt einen lokalen Free/Busy-Proxy bereit und liefert einen leistungsfaehigen Filelink-Assistenten fuer E-Mails. Ziel ist ein professioneller Workflow vom Kalender bis zur Dateiablage -- ohne Medienbruch und mit klarer Administrierbarkeit.

Dies ist ein Community-Projekt und kein offizielles Produkt der Nextcloud GmbH.

## Highlights

- **Ein Klick zu Nextcloud Talk** 
Termin öffnen, Nextcloud Talk wählen, Raum konfigurieren, Moderator definieren. Optional können eingeladene Teilnehmer direkt in den Raum übernommen werden (getrennt nach internen Nextcloud-Benutzern und externen E-Mail-Gästen). Der Wizard schreibt Titel/Ort/Beschreibung inklusive Hilfe-Link automatisch in den Termin.
- **Sharing deluxe** 
Compose-Button Nextcloud Freigabe hinzufügen startet den Freigabe-Assistenten mit Upload-Queue, Passwortgenerator, Ablaufdatum und Notizfeld. Die fertige Freigabe landet als formatiertes HTML direkt in der E-Mail.
- **Enterprise-Sicherheit** 
Lobby bis Startzeit, Moderator-Delegation, automatisches Aufräumen nicht gespeicherter Termine, Pflicht-Passwörter und Ablauffristen schützen sensible Meetings und Dateien.
- **Internet Free/Busy Gateway (IFB)**  
Lokaler HTTP-Listener beantwortet Outlook-Free/Busy-Anfragen direkt aus Nextcloud. Registry-Werte fuer Suchpfad und Read-URL werden gesetzt. Bei HTTP 404 faellt das Add-in auf Scheduling-POST zurueck, sodass Verfuegbarkeiten bereitstehen.
- **Debug-Logging auf Knopfdruck**  
Im Debug-Tab aktivierbar. Schreibt strukturierte Logs (Authentifizierung, Termin- und Filelink-Flows, IFB) nach `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`. Der Speicherort wird im UI angezeigt.

## Changelog

Siehe [`CHANGELOG.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/CHANGELOG.md).

## Funktionsüberblick

### Nextcloud Talk direkt aus dem Termin
- Talk-Popup mit Lobby, Passwort, Listbarkeit, Raumtyp und Moderatorensuche.
- Automatische Einträge von Titel, Ort, Beschreibung (inkl. Hilfe-Link und Passwort) in das Terminfenster.
- Room-Tracking, Lobby-Updates, Delegations-Workflow und Cleanup, falls der Termin verworfen oder verschoben wird.
- Kalender-Aenderungen (Drag-and-drop oder Dialog-Edit) halten Lobby/Startzeit des Talk-Raums synchron.
- Optionales Teilnehmer-Sync nach dem Speichern des Termins:
  - **Benutzer:** interne Nextcloud-Benutzer werden direkt dem Raum hinzugefügt.
  - **Gäste:** externe E-Mail-Adressen werden als Gäste eingeladen (ggf. zusätzliche Einladung per E-Mail durch Nextcloud).

### Nextcloud Sharing im Compose-Fenster
- Vier Schritte (Freigabe, Ablaufdatum, Dateien, Notiz) mit passwortgeschütztem Upload-Ordner.
- Upload-Queue mit Duplikatprüfung, Fortschrittsanzeige und optionaler Freigabe.
- Automatische HTML-Bausteine mit Link, Passwort, Ablaufdatum und optionaler Notiz.

### Administration & Compliance
- Login Flow V2 (App-Passwort wird automatisch angelegt) und zentrale Optionen (Basis-URL, Debug-Modus, Freigabe-Pfade, Defaultwerte fuer Freigabe/Talk).
- Vollständige Internationalisierung (siehe [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md)) und strukturierte Debug-Logs für Support-Fälle.

## Sprache & Übersetzungen

- Die UI-Sprache folgt der Outlook/Office-Bedienoberfläche (Office UI language). Wenn Outlook auf **Systemeinstellungen verwenden** steht, entspricht das in der Regel der Windows-Anzeigesprache.
- Unterstützte Sprachen sind in [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md) dokumentiert. Fallback ist `de`, danach `en`.

### Sprach-Overrides (Textbausteine)

In den Einstellungen unter **Erweitert** können Sie die Sprache für eingefügte Textbausteine unabhängig von der UI-Sprache festlegen:

- **Freigabe-HTML-Block** (E-Mail): Sprache des formatierten HTML-Blocks beim Teilen.
- **Talk-Beschreibungstext** (Termin): Sprache des eingefügten Textblocks (z.B. Passwortzeile / Hilfe-Link).

Option `Default (UI)` nutzt die aktuelle UI-Sprache (inkl. Fallbacks).

## Systemanforderungen

- Windows 10 oder Windows 11 (64 Bit)  
- Microsoft Outlook classic >=2019  
- .NET Framework 4.7.2 Runtime  
- Nextcloud Server mit Talk- und Filesharing-App

## Installation und Updates

1. Outlook schliessen.  
2. Aktuelle MSI (z.B. `NCConnectorForOutlook-2.2.7.msi`) ausfuehren und den UAC-Prompt bestaetigen (Administratorrechte sind erforderlich). Das Setup richtet URLACL sowie alle benoetigten Registry-Schluessel fuer IFB ein.  
3. Outlook starten und im Ribbon **NC Connector** auf **Einstellungen** klicken.  
4. Login-Modus waehlen, Verbindungstest ausfuehren, Einstellungen speichern. Bei erfolgreichem Test bleibt IFB automatisch aktiv.  
5. Filelink-Basisverzeichnis pruefen und Debug-Logging bei Bedarf aktivieren.

Updates erfolgen durch Installation einer hoeheren MSI-Version. Persoenliche Einstellungen (`settings.ini`) bleiben erhalten. Die Deinstallation entfernt das Add-in, stoppt den IFB-Listener und setzt die Registry-Werte zurueck.

## Troubleshooting

- **Debug-Log**: Tab *Debug* aktivieren. Log-Datei: `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`.  
- **Add-in nicht sichtbar**: Installation muss mit Adminrechten erfolgen. Pruefe `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn` und ggf. Repair in einer Admin-Konsole: `msiexec /i "NCConnectorForOutlook-2.2.7.msi" ADDLOCAL=ALL`.  
- **IFB testen**: `powershell -Command "Invoke-WebRequest http://127.0.0.1:7777/nc-ifb/freebusy/<mail>.vfb -UseBasicParsing"`. Bei Abweichungen Registry unter `HKCU\Software\Microsoft\Office\<Version>\Outlook\Options\Calendar` pruefen.  
- **TLS/Proxy pruefen**: `powershell -Command "Test-NetConnection <Ihre-Domain> -Port 443"`. Bei SSL-Warnungen Zertifikate/Proxy kontrollieren.  
- **Filelink-Fehler**: Debug-Log liefert HTTP-Statuscodes und Exception-Meldungen. Pflichtfelder im Wizard sind validiert.

## Screenshots

<details>
<summary><strong>Settings</strong></summary>

| <a href="Screenshots/settings.jpg"><img src="Screenshots/settings.jpg" alt="Settings Dialog" width="230"></a> |
| --- |

</details>

<details>
<summary><strong>Talk-Link Workflow</strong></summary>

| <a href="Screenshots/1_talk.jpg"><img src="Screenshots/1_talk.jpg" alt="Talk Schritt 1" width="230"></a> | <a href="Screenshots/2_talk.jpg"><img src="Screenshots/2_talk.jpg" alt="Talk Schritt 2" width="230"></a> |
| --- | --- |

</details>

<details open>
<summary><strong>Filelink Wizard</strong></summary>

| <a href="Screenshots/1_filelink.jpg"><img src="Screenshots/1_filelink.jpg" alt="Filelink Schritt 1" width="230"></a> | <a href="Screenshots/2_filelink.jpg"><img src="Screenshots/2_filelink.jpg" alt="Filelink Schritt 2" width="230"></a> |
| --- | --- |
| <a href="Screenshots/3_filelink.jpg"><img src="Screenshots/3_filelink.jpg" alt="Filelink Schritt 3" width="230"></a> | <a href="Screenshots/4_filelink.jpg"><img src="Screenshots/4_filelink.jpg" alt="Filelink Schritt 4" width="230"></a> |
| <a href="Screenshots/5_filelink.jpg"><img src="Screenshots/5_filelink.jpg" alt="Filelink Schritt 5" width="230"></a> | |

</details>

<details>
<summary><strong>Internet Free/Busy</strong></summary>

| <a href="Screenshots/ifb.jpg"><img src="Screenshots/ifb.jpg" alt="IFB Einstellungen" width="230"></a> |
| --- |

</details>





