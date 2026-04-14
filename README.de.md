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
Compose-Button Nextcloud Freigabe hinzufügen startet den Freigabe-Assistenten mit Upload-Queue, Passwortgenerator, Ablaufdatum, Notizfeld, Anhangs-Automatisierung und optionalem separatem Passwort-Follow-up. Die fertige Freigabe landet als formatiertes HTML direkt in der E-Mail.
- **Enterprise-Sicherheit** 
Lobby bis Startzeit, Moderator-Delegation, automatisches Aufräumen nicht gespeicherter Termine, Pflicht-Passwörter und Ablauffristen schützen sensible Meetings und Dateien.
- **Zentrale Backend-Policies (optional)**
Ist das optionale NC-Connector-Backend installiert, koennen Talk- und Sharing-Defaults zentral gesteuert werden. Beim Oeffnen von Wizard und Settings prueft das Add-in den Backend-Status, uebernimmt bei gueltigem Seat die Policy-Werte und sperrt admin-kontrollierte Optionen sichtbar im UI.
- **Internet Free/Busy Gateway (IFB)**  
Lokaler HTTP-Listener beantwortet Outlook-Free/Busy-Anfragen direkt aus Nextcloud. Registry-Werte fuer Suchpfad und Read-URL werden gesetzt. Bei HTTP 404 faellt das Add-in auf Scheduling-POST zurueck, sodass Verfuegbarkeiten bereitstehen.
- **Debug-Logging auf Knopfdruck**  
Im Debug-Tab aktivierbar. Schreibt strukturierte Logs (Authentifizierung, Termin- und Filelink-Flows, IFB) nach `%LOCALAPPDATA%\NC4OL\addin-runtime.log`. Laufzeit-Exceptions werden dort auch dann geschrieben, wenn der Debug-Schalter aus ist. Der Speicherort wird im UI angezeigt.

## Changelog

Siehe [`CHANGELOG.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/CHANGELOG.md).

## Funktionsüberblick

### Nextcloud Talk direkt aus dem Termin
- Talk-Popup mit Lobby, Passwort, Listbarkeit, Raumtyp und Moderatorensuche.
- Automatische Einträge von Titel, Ort, Beschreibung (inkl. Hilfe-Link und Passwort) in das Terminfenster.
- Room-Tracking, Lobby-Updates, Delegations-Workflow und Cleanup, falls der Termin verworfen oder verschoben wird.
- Kalender-Aenderungen (Drag-and-drop oder Dialog-Edit) halten Lobby/Startzeit des Talk-Raums synchron.
- Wenn Sie einen Termin mit Moderator-Delegation speichern, aktualisiert NC Connector zuerst Raumname, Lobbyzeit, Beschreibung und Teilnehmer und uebergibt danach die Moderation.
- Live-Verfuegbarkeitschecks fuer das Systemadressbuch (bei Talk-Klick, Settings Open/Save und Wizard-Open) mit deterministischem Lock-Verhalten:
  - `Benutzer hinzufuegen`, `Gaeste hinzufuegen` und Moderator-Controls werden bei Nichtverfuegbarkeit deaktiviert.
  - In den Settings erscheint ein roter Warnblock mit Setup-Link.
  - Im Talk-Wizard erscheint in der Moderator-Sektion ein roter Inline-Warnblock.
- Optionales Teilnehmer-Sync nach dem Speichern des Termins:
  - **Benutzer:** interne Nextcloud-Benutzer werden direkt dem Raum hinzugefügt.
  - **Gäste:** externe E-Mail-Adressen werden als Gäste eingeladen (ggf. zusätzliche Einladung per E-Mail durch Nextcloud).

### Nextcloud Sharing im Compose-Fenster
- Vier Schritte (Freigabe, Ablaufdatum, Dateien, Notiz) mit passwortgeschütztem Upload-Ordner.
- Upload-Queue mit Duplikatprüfung, Fortschrittsanzeige und optionaler Freigabe.
- Automatische HTML-Bausteine mit Link, Passwort, Ablaufdatum und optionaler Notiz.
- Optionale Anhangs-Automatisierung:
  - neue Anhaenge immer ueber NC Connector verarbeiten, oder
  - Prompt ab konfigurierbarer Gesamtgroesse anzeigen.
- Die Anhangs-Automatisierung nutzt zusaetzlich einen Pre-Add-Check (`BeforeAttachmentAdd`) und kann Host-Adds best effort bereits vor der normalen Outlook-Post-Add-Verarbeitung abbrechen, wenn ein aufloesbarer lokaler Dateipfad vorliegt.
- Der Schwellwert-Prompt hat genau zwei Aktionen:
  - `Share with NC Connector`
  - `Remove last selected attachments` (Batch-semantisch).
- Attachment-Modus:
  - fixer Share-Basisname `email_attachment` mit deterministischen Suffixen (`_1`, `_2`, ...)
  - Empfaengerberechtigung immer read-only
  - HTML-Ausgabe mit ZIP-Download-URL `/s/<token>/download` ohne Permissions-Zeile.
- Die Datei-Queue akzeptiert Explorer-Drag-and-drop fuer Dateien und Ordner ueber den kompletten Datei-Schritt (Queue + Aktionsbereich), nicht nur ueber die Add-Buttons.
- Optionaler separater Passwort-Mail-Flow:
  - nur verfuegbar mit optionalem NC-Connector-Backend und aktivem, dem aktuellen Benutzer zugewiesenem Seat
  - Inline-Passwort im Hauptblock ausblenden
  - Passwort nach erfolgreichem Hauptversand als Follow-up-Mail senden
  - bei Auto-Send-Fehlern vorbefuellten manuellen Fallback-Entwurf oeffnen.

### Administration & Compliance
- Login Flow V2 (App-Passwort wird automatisch angelegt) und zentrale Optionen (Basis-URL, Debug-Modus, Freigabe-Pfade, Defaultwerte fuer Freigabe/Talk).
- Optionaler NC-Connector-Backend-Status/Policy-Vertrag:
  - Pruefung beim Oeffnen von Talk-Wizard, Sharing-Wizard und Settings sowie beim Speichern der Settings
  - gueltiger aktiver Seat aktiviert Backend-Policy-Werte und Admin-Locks
- fehlendes Backend / kein Seat / ungueltiger Seat / abgelaufene Grace-Zeit faellt auf lokale Outlook-Settings zurueck
  - ungueltige Seat-Zustaende werden sichtbar im UI angezeigt, damit sich Benutzer an den Administrator wenden koennen
- Backend-Templates fuer Freigabe- und Talk-Texte werden nur aktiv, wenn in den Sprach-Overrides `Benutzerdefiniert` gewaehlt ist
- `Benutzerdefiniert` wird nur angezeigt, wenn der NC-Connector-Backend-Endpunkt existiert, und bleibt deaktiviert, solange die effektive Backend-Policy fuer diesen Bereich nicht wirklich `custom` ist oder keine Vorlage liefert
- ist `Benutzerdefiniert` aktiv, aber die Backend-Vorlage leer oder nicht verfuegbar, faellt Outlook auf den lokalen UI-Default-Text zurueck
- Vollständige Internationalisierung (siehe [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md)) und strukturierte Debug-Logs für Support-Fälle.

## Sprache & Übersetzungen

- Die UI-Sprache folgt der Outlook/Office-Bedienoberfläche (Office UI language). Wenn Outlook auf **Systemeinstellungen verwenden** steht, entspricht das in der Regel der Windows-Anzeigesprache.
- Unterstützte Sprachen sind in [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md) dokumentiert. Fallback ist `de`, danach `en`.

### Sprach-Overrides (Textbausteine)

In den Einstellungen unter **Erweitert** können Sie die Sprache für eingefügte Textbausteine unabhängig von der UI-Sprache festlegen:

- **Freigabe-HTML-Block** (E-Mail): Sprache des formatierten HTML-Blocks beim Teilen.
- **Talk-Beschreibungstext** (Termin): Sprache des eingefügten Textblocks (z.B. Passwortzeile / Hilfe-Link).

Option `Default (UI)` nutzt die aktuelle UI-Sprache (inkl. Fallbacks).

Option `Benutzerdefiniert` wird nur angezeigt, wenn der NC-Connector-Backend-Endpunkt existiert. Auswaehlbar wird sie erst dann, wenn die effektive Backend-Policy fuer den jeweiligen Bereich wirklich `custom` ist und eine Backend-Vorlage vorliegt. Andernfalls bleibt die Option deaktiviert und Outlook verwendet weiter den lokalen UI-Default-Block.

## Systemanforderungen

- Windows 10 oder Windows 11 (64 Bit)  
- Microsoft Outlook classic >=2019  
- .NET Framework 4.7.2 Runtime  
- Nextcloud Server mit Talk- und Filesharing-App

## Installation und Updates

1. Outlook schliessen.  
2. Aktuelle MSI (z.B. `NCConnectorForOutlook-3.0.1.msi`) ausfuehren und den UAC-Prompt bestaetigen (Administratorrechte sind erforderlich). Das Setup richtet URLACL sowie alle benoetigten Registry-Schluessel fuer IFB ein.  
3. Outlook starten und im Ribbon **NC Connector** auf **Einstellungen** klicken.  
4. Login-Modus waehlen, Verbindungstest ausfuehren, Einstellungen speichern. Bei erfolgreichem Test bleibt IFB automatisch aktiv.  
5. Filelink-Basisverzeichnis pruefen und Debug-Logging bei Bedarf aktivieren.

Updates erfolgen durch Installation eines MSI-Pakets ueber die bestehende Installation (gleiche, aeltere oder neuere Version). Persoenliche Einstellungen bleiben erhalten und werden in profilbasierte XML-Dateien (`settings_<OutlookProfile>.xml`) unter `%LOCALAPPDATA%\NC4OL` migriert. Die Deinstallation entfernt das Add-in, stoppt den IFB-Listener und setzt die Registry-Werte zurueck.

### Release-Notizen 3.0.1 (Betrieb)

- Runtime-Artefakte sind unter `%LOCALAPPDATA%\NC4OL` gebuendelt:
  - Settings-Dateien (`settings_<OutlookProfile>.xml`)
  - IFB-/Systemadressbuch-Cache
  - Debug-Log (`addin-runtime.log`)
- Legacy-INI-Settings aus aelteren Builds werden beim ersten Start migriert und nach erfolgreicher Migration entfernt.
- Im Compose-Attachment-Modus wird das serverseitige Cleanup direkt nach Share-Erstellung armed und nur nach bestaetigtem erfolgreichem Mailversand wieder cleared.
- Wenn separates Passwort-Senden aktiv ist, blendet die Hauptmail das Inline-Passwort aus und der Passwort-Follow-up-Versand wird erst nach bestaetigtem erfolgreichem Hauptversand gestartet. Dieses Feature ist nur mit Backend-Endpunkt + aktiv zugewiesenem Seat verfuegbar.

## Troubleshooting

- **Debug-Log**: Tab *Debug* für ausführliche Traces aktivieren. Log-Datei: `%LOCALAPPDATA%\NC4OL\addin-runtime.log`. Bei aktiviertem Debug-Logging sind auch Attachment-Pre-Add-Entscheidungen/Fallback-Gruende enthalten. Laufzeit-Exceptions werden dort auch bei deaktiviertem Debug-Logging geschrieben.  
- **Add-in nicht sichtbar**: Installation muss mit Adminrechten erfolgen. Pruefe `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn` und ggf. Repair in einer Admin-Konsole: `msiexec /i "NCConnectorForOutlook-3.0.1.msi" ADDLOCAL=ALL`.  
- **IFB testen**: `powershell -Command "Invoke-WebRequest http://127.0.0.1:7777/nc-ifb/freebusy/<mail>.vfb -UseBasicParsing"`. Bei Abweichungen Registry unter `HKCU\Software\Microsoft\Office\<Version>\Outlook\Options\Calendar` pruefen.  
- **TLS/Proxy pruefen**: `powershell -Command "Test-NetConnection <Ihre-Domain> -Port 443"`. Bei SSL-Warnungen Zertifikate/Proxy kontrollieren. NC Connector aktiviert starke Kryptografie und die TLS-Standards des Betriebssystems add-in-lokal ueber `NcTalkOutlookAddIn.dll.config`; wenn Secure-Channel-Fehler trotzdem auftreten, pruefen Sie Zertifikatsvertrauen, TLS-pruefende Proxys, DNS und die TLS-/Schannel-Richtlinien des Systems. Maschinenweite Registry-/GPO-Overrides bleiben eine optionale Admin-Massnahme, sind aber kein Installer-Eingriff.  
- **Anhangsautomatisierung greift nicht bei grossen Dateien**: In Microsoft-365-/Exchange-Umgebungen kann Outlook Anhaenge vor den Add-in-Events blockieren (z. B. bei serverseitigen Groessenlimits). In diesen Faellen den Button **`Nextcloud Freigabe hinzufuegen`** verwenden.  
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







