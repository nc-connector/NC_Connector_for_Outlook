<div align="center" style="background:#0082C9; padding:1px 0;"><img src="assets/header-solid-blue-1920x480.png" alt="Addon" height="80"></div>

[English](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.md) | [Deutsch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.de.md)
[Admin-Handbuch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/ADMIN.de.md) | [Entwicklerdoku](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/DEVELOPMENT.de.md) | [Übersetzungen](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md) | [VENDOR](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/VENDOR.md)

# NC Connector for Outlook

NC Connector bringt Nextcloud-Freigaben, Talk-Meetings, zentrale Signaturen und Free/Busy-Daten in Outlook classic. Für Organisationen, die Outlook behalten und Nextcloud als eigene Infrastruktur nutzen wollen.


## Was das Add-in macht

- Nextcloud-Freigaben direkt aus neuen Mails, Antworten und Weiterleitungen erstellen
- große Dateien per Nextcloud Chunked WebDAV Upload v2 hochladen und als Link senden
- Passwort, Ablaufdatum, Berechtigungen und separate Passwortzustellung steuern
- Passwörter wahlweise als Klartext-Mail oder als Nextcloud Secret-Link senden
- Talk-Räume direkt aus Outlook-Terminen erstellen und aktualisieren
- zentrale E-Mail-Signaturen aus dem optionalen Backend anwenden
- Outlook Free/Busy über einen lokalen Nextcloud-Proxy bereitstellen
- Debug-Logs für Supportfälle schreiben, optional anonymisiert

## Optionales Backend

Ohne Backend funktionieren Freigaben, Talk und IFB lokal in Outlook. Mit NC Connector Backend kommen zentrale Steuerung und Team-Funktionen hinzu:

- Seat-Zuteilung und Richtlinien
- Vorgaben für Freigaben, Talk und Signaturen
- eigene HTML-Vorlagen für Freigaben, Passwortmails und Talk-Einladungen
- separate Passwortzustellung und optional Nextcloud Secret-Links
- Sperren einzelner Optionen durch Administratoren

## Freigaben

Der Freigabe-Assistent lädt Dateien und Ordner nach Nextcloud hoch und fügt den fertigen Freigabeblock in die Mail ein. HTML/RTF bekommt einen formatierten Block, Plaintext einen klaren Textblock.

Weitere Punkte:

- verfügbar in Compose-Fenstern, Antworten, Weiterleitungen und Inline-Antworten
- optionales Ablaufdatum und eigene Berechtigungen pro Freigabe
- Anhangsautomatisierung für große Anhänge oder immer über NC Connector
- separate Passwortmails werden erst nach erfolgreichem Versand der Hauptmail verschickt
- bei Auto-Send-Fehlern öffnet sich eine vorbereitete manuelle Passwortmail

## Talk

Aus einem Outlook-Termin kann direkt ein Nextcloud Talk-Raum erstellt werden. Der Dialog unterstützt Lobby, Passwort, Raumtyp und Moderation.

NC Connector kann Terminänderungen mit dem Raum abgleichen und eingeladene Teilnehmer übernehmen. Das Löschen gespeicherter Talk-Termine entfernt Räume nur nach ausdrücklicher Aktivierung.

## Signaturen

Mit Backend kann Outlook zentral verwaltete E-Mail-Signaturen einfügen oder lokale Signaturen entfernen, wenn die Policy das vorgibt. NC Connector greift dabei nur die Signatur der passenden Absenderadresse an. Signaturen anderer Konten bleiben unberührt.

## Installation

1. Outlook schließen.
2. Aktuelle MSI aus den [GitHub Releases](https://github.com/nc-connector/NC_Connector_for_Outlook/releases) installieren.
3. Outlook starten und **NC Connector -> Einstellungen** öffnen.
4. Nextcloud-URL eintragen.
5. Login mit Nextcloud oder manuelles App-Passwort nutzen.
6. Verbindung testen und speichern.

Updates werden durch Installation der neuen MSI über die bestehende Version eingespielt. Persönliche Einstellungen bleiben erhalten.

## Voraussetzungen

- Windows 10 oder Windows 11
- Outlook classic 2019 oder neuer
- .NET Framework 4.7.2
- Nextcloud mit Files Sharing
- für Talk-Funktionen: Nextcloud Talk
- für Secret-Link-Passwortzustellung: Nextcloud Secrets und NC Connector Backend

## Sprache

Die UI-Sprache folgt der Outlook/Office-Sprache. Unterstützte Sprachen sind in [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md) dokumentiert. Fallback ist Deutsch, danach Englisch.

Textbausteine für Freigaben und Talk können in den Einstellungen unabhängig von der UI-Sprache gesetzt werden. Backend-Vorlagen werden nur genutzt, wenn das Backend vorhanden ist und die Policy sie freigibt.

## Fehleranalyse

Debug-Logs lassen sich in den Einstellungen aktivieren. Die Dateien liegen unter:

`%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`

Die Anonymisierung ist standardmäßig aktiv und maskiert Server-URL, Zugangsdaten, E-Mail-Adressen und lokale Benutzerpfade.

Für typische Setup-, IFB- und Backend-Policy-Probleme siehe das [Admin-Handbuch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/ADMIN.de.md).

## Weitere Dokumentation

- [Changelog](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/CHANGELOG.md)
- [Admin-Handbuch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/ADMIN.de.md)
- [Entwicklerdoku](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/DEVELOPMENT.de.md)
- [Drittanbieter-Lizenzen](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/VENDOR.md)

## Screenshots

<details>
<summary><strong>Einstellungen</strong></summary>

| <a href="Screenshots/settings.jpg"><img src="Screenshots/settings.jpg" alt="Einstellungen" width="230"></a> |
| --- |

</details>

<details>
<summary><strong>Talk-Link</strong></summary>

| <a href="Screenshots/1_talk.jpg"><img src="Screenshots/1_talk.jpg" alt="Talk Schritt 1" width="230"></a> | <a href="Screenshots/2_talk.jpg"><img src="Screenshots/2_talk.jpg" alt="Talk Schritt 2" width="230"></a> |
| --- | --- |

</details>

<details open>
<summary><strong>Freigabe-Assistent</strong></summary>

| <a href="Screenshots/1_filelink.jpg"><img src="Screenshots/1_filelink.jpg" alt="Freigabe Schritt 1" width="230"></a> | <a href="Screenshots/2_filelink.jpg"><img src="Screenshots/2_filelink.jpg" alt="Freigabe Schritt 2" width="230"></a> |
| --- | --- |
| <a href="Screenshots/3_filelink.jpg"><img src="Screenshots/3_filelink.jpg" alt="Freigabe Schritt 3" width="230"></a> | <a href="Screenshots/4_filelink.jpg"><img src="Screenshots/4_filelink.jpg" alt="Freigabe Schritt 4" width="230"></a> |
| <a href="Screenshots/5_filelink.jpg"><img src="Screenshots/5_filelink.jpg" alt="Freigabe Schritt 5" width="230"></a> | |

</details>

<details>
<summary><strong>Internet Free/Busy</strong></summary>

| <a href="Screenshots/ifb.jpg"><img src="Screenshots/ifb.jpg" alt="IFB Einstellungen" width="230"></a> |
| --- |

</details>
