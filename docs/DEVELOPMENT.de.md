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
- **Zentrale Backend-E-Mail-Signaturen** fuer passende Outlook-Absenderkonten
- **IFB (Internet Free/Busy)** als lokaler HTTP-Proxy zu Nextcloud

## Release 3.1.0 Delta-Ueberblick

Diese Release-Linie erweitert Compose-Unterstuetzung und zentrale Backend-Signaturen:

- Backend-gesteuerte E-Mail-Signaturen gelten fuer passende Outlook-Absenderidentitaeten in HTML/RTF und Plain Text, auch bei Antworten und Weiterleitungen.
- Nextcloud-Freigaben koennen aus Inline-Antworten/-Weiterleitungen eingefuegt werden und laufen ueber WordEditor, damit zitierte Inhalte erhalten bleiben.
- Plain-Text-Freigabebloecke werden eingefuegt, ohne `MailItem.Body` direkt umzuschreiben.
- Grosse Dateien nutzen Nextcloud Chunked WebDAV Upload v2; der Freigabe-Wizard zeigt Uploadgeschwindigkeit pro Datei.
- Separate Passwort-Follow-up-Mails behalten die Absenderidentitaet des Original-Compose, bekommen bei Policy-/Absender-Match die Backend-Signatur und oeffnen bei Auto-Send-Fehler weiterhin einen manuellen Fallback-Entwurf.
- Talk-Raumloeschung fuer gespeicherte Termine bleibt Opt-in; Talk-Cleanup-Metadaten bleiben lokal in Outlook.

## Voraussetzungen

- Windows 10/11
- Outlook classic (typischerweise x64)
- **.NET Framework 4.7.2** (Target)
- MSBuild (z.B. Visual Studio Build Tools)
- **.NET SDK** (für den WiX-v6-Build via `dotnet`)

### Reference Assemblies (FrameworkPathOverride)

Auf manchen Build-Systemen fehlen die .NET Framework Reference Assemblies für 4.7.2 (insbesondere CI/Minimal-Installationen). In dem Fall kann man die NuGet-ReferenceAssemblies nutzen und `FrameworkPathOverride` setzen.

Beispiel:

```powershell
cd "C:\Pfad\zum\nc4ol"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages -ExcludeVersion

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

## Build (MSI)

Der empfohlene Build läuft immer über `build.ps1`:

```powershell
cd "C:\Pfad\zum\nc4ol"
$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
.\build.ps1 -Configuration Release
```

Wenn auf dem Build-Host die WiX-ICE-Validierung nicht verfuegbar ist (z. B. `WIX0217` in eingeschraenkten Umgebungen), verwende:

```powershell
.\build.ps1 -Configuration Release -SkipIceValidation
```

Output:

- `dist\NCConnectorForOutlook-<version>.msi`

Was das Script macht:

1) Build des COM Add-ins (`NcTalkOutlookAddIn.sln`) via MSBuild
2) Ermittelt die Assembly-Version aus `NcTalkOutlookAddIn.dll`
3) Build des WiX-v6-Installers (`installer/NcConnectorOutlookInstaller.wixproj`)
4) Kopiert das MSI in `dist/`

## Lokales Testen

Die automatisierten Prüfungen unter `tools/ci/` laufen über die Jobs in
`.github/workflows/outlook-build-checks.yml`. Anschließend decken die folgenden Smoke-Tests das Outlook-COM-Verhalten ab:

1) MSI installieren (als Admin):
   - `msiexec /i dist\NCConnectorForOutlook-<version>.msi`
2) Outlook starten
3) Kalendertermin öffnen:
   - Ribbon: **NC Connector → Talk-Link einfügen**
4) E-Mail erstellen:
   - Ribbon: **NC Connector → Nextcloud Freigabe hinzufügen**
   - Inline-Antwort/-Weiterleitung: **Nachricht → NC Connector → Nextcloud Freigabe hinzufügen**
5) Optional: IFB in Settings aktivieren, Port pruefen (`Einstellungen -> IFB`, Standard `7777`) und Endpunkt testen
   - `Invoke-WebRequest http://127.0.0.1:<ifb-port>/nc-ifb/ -UseBasicParsing`
6) Einstellungen -> Erweitert: `Jetzt prüfen` ausführen und kontrollieren, dass aktuelle Version, letzte Prüfung, Download-Link und Änderungsübersicht angezeigt werden.

## Logging

- Aktivierung: Settings → Tab **Debug**
- Option (Standard aktiv): `Logs anonymisieren`
- Datei (taeglich): `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`
- Runtime-Exceptions werden ueber `DiagnosticsLogger.LogException(...)` immer geschrieben, auch wenn Debug deaktiviert ist.
- Aufbewahrung: letzte 7 Tageslogs behalten, Logs aelter als 30 Tage (best effort) entfernen.
- Bei aktiver Anonymisierung werden NC-URL/Basis-Host, Token/Secrets, Authorization-Werte, E-Mails, Benutzerkennungen und lokale User-Pfadsegmente maskiert.

Kategorien (Beispiele):

- `CORE` (Start, Settings, Registry)
- `API` (HTTP Calls / Statuscodes)
- `TALK` (Room Lifecycle, Lobby, Delegation)
- `FILELINK` (Upload/Share)
- `IFB` (Requests, Cache, Outlook Registry)

## Code-Struktur

Root:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs`  
  Einstiegspunkt, Ribbon, Outlook-Events, Composition Root fuer die Workflows.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs`  
  Add-in-Bootstrap/Teardown (`OnConnection`, Shutdown/Disconnect).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs`
  Dedizierte Outlook-Event Hook-/Unhook-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs`
  Kategorienspezifische Runtime-Logging-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs`  
  Backend-Policy- und Talk-Template-/Sprach-Resolver.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs`  
  Deferred Appointment-Subscription-Ensure inkl. Outlook-Event-Restriction-Handling.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs`
  Runtime-Subscription-Core fuer Compose-Lifecycle-Zustand (`Dispose`, Identity, gemeinsame Helper).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs`
  Compose-Attachment-Interception/Evaluation/Share-Launch-Flow.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Signature.cs`
  Backend-E-Mail-Signatur-Policy fuer das passende Outlook-Absenderkonto.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs`
  Send/Close-Cleanup-Lifecycle inkl. separatem Passwort-Dispatch.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs`
  Runtime-Subscription fuer Termin-Write/Close/Delete und Lifecycle-Cleanup.
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs`
  Orchestrierung fuer Settings-Open/Save/Revert.
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs`
  Orchestrierung fuer FileLink-Ribbon-Start und Wizard-Flow.
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs`
  Orchestrierung fuer Talk-Ribbon-Flow (Auth-Gate, Wizard, Room-Create/Replace).
- `TalkRibbonController` und `FileLinkLaunchController` prefetchen Backend-Policy + Passwort-Policy parallel (`Task.WhenAll`) vor dem Wizard-Open; Policy-Daten bleiben dabei pro Einstiegspunkt immer frisch.
- Nach dem FileLink-, Talk- oder Settings-Prefetch wechselt `OutlookUiSynchronizationContext` vor jedem WinForms- oder Outlook-COM-Zugriff zurueck auf den beim Add-in-Start erfassten Outlook-STA-Thread. Outlook stellt COM-Callbacks nicht verlaesslich einen `SynchronizationContext` bereit; modale Dialoge und Outlook-Interop werden deshalb explizit zurueckgeschaltet.

Controller:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` (Talk-Termin-Lifecycle: Room-Metadaten, Lobby-/Description-/Delegation-/Teilnehmer-Sync)
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` (Compose-Share-Cleanup und separater Passwort-Versand inkl. Fallback)
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` (Talk-Template-/Block-Rendering)
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` (SMTP- und Attendee-Aufloesung)
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` (Compose-Subscription-Registry)
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` (gemeinsame Mail-/Inspector-Interop-Helper und einheitlicher WordEditor-Signatur-Slot-Reconciler)
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` (gemeinsames Queue-Modell fuer separaten Passwort-Follow-up)
- `src/NcTalkOutlookAddIn/Settings/ManagedSetupPolicy.cs` (verwaltete Nextcloud-URL aus Registry/GPO)

Services:

- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (DAV/Share Flow)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB; Port ueber Settings konfigurierbar, Standard `7777`)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)
- `src/NcTalkOutlookAddIn/Services/NcHttpClient.cs` (zentraler Request-Executor fuer Auth-Header, OCS-Header, Timeout/Decompression und optionalen Fresh-Connection-Mode)
  - Alle Runtime-HTTP-Aufrufe (Talk, Share/DAV, IFB, Login-Flow, Moderator-Avatar-Fetch) laufen zentral ueber `NcHttpClient`.
- `src/NcTalkOutlookAddIn/Services/EmailSignaturePolicyService.cs` (loest Backend-E-Mail-Signatur-Policy gegen lokale Settings und Lock-State auf)
- `src/NcTalkOutlookAddIn/Services/UpdateCheckService.cs` (fragt einmal pro Tag `nc-connector.de` nach Outlook-Release-Metadaten und speichert das Ergebnis in den Profil-Settings)

Update-Check:

- Endpoint: `GET https://nc-connector.de/wp-json/ncc/v1/update-check`
- Gesendet werden Produkt, installierte Version, Kanal und ein taeglich wechselnder anonymer Client-Hash.
- Downloads zeigen direkt auf GitHub-Release-Dateien; die Homepage zaehlt nur taegliche Installationen und liefert Release-Metadaten.

UI:

- `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`
- `src/NcTalkOutlookAddIn/UI/TalkLinkForm.cs`
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
- `src/NcTalkOutlookAddIn/UI/ComposeAttachmentPromptForm.cs` (2-Aktions-Prompt fuer Schwellwertmodus)
- `src/NcTalkOutlookAddIn/UI/BrandedHeader.cs` (Header-Banner inkl. `AttachToParent(...)` fuer konsistente Header-Initialisierung in Forms)
- `src/NcTalkOutlookAddIn/UI/ScaledForm.cs` (zentrale DPI-Skalierung via `ScaleLogical(...)`, damit Form-Wrapper nicht dupliziert werden)

Utilities:

- `src/NcTalkOutlookAddIn/Utilities/BrowserLauncher.cs` (zentraler Shell-Start fuer URLs, Dateien und Ordner)
- `src/NcTalkOutlookAddIn/Utilities/SizeFormatting.cs` (zentrale MB-Formatierung fuer UI-Texte)
- `src/NcTalkOutlookAddIn/Utilities/ComInteropScope.cs` (zentrale COM-Release-/FinalRelease-Helfer)
- `src/NcTalkOutlookAddIn/Utilities/PasswordGenerationHelper.cs` (zentralisiert Min-Length-Aufloesung, Server-Fallback-Generierung und gemeinsame Min-Length-Validierung fuer Talk/FileLink-Formulare)
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` (zentraler Sanitizer fuer Backend-HTML-Templates bei Share/Talk, fail-closed)
- `src/NcTalkOutlookAddIn/Utilities/HtmlToPlainTextConverter.cs` (DOM-basierte HTML-zu-Plain-Text-Ausgabe fuer Plain-Text-E-Mail-Signaturen)
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` (zentrale JSON-Normalisierung inkl. `PrepareJsonPayload`, Dictionary-/String-/Int-Helfer und OCS-Fehlerextraktion)
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` (gekapselter Laufzeitzustand fuer Pending-Keys und Restriction-Log-Throttling)
- `src/NcTalkOutlookAddIn/Utilities/PictureConverter.cs` (gemeinsamer Image->IPictureDisp-Helfer fuer Ribbon-Icons)

### Zentrale E-Mail-Signatur im Compose-Fenster

Die Compose-Subscription wertet die zentrale Signatur nach dem Oeffnen einer Compose-Oberflaeche, nach Absender- oder BodyFormat-Aenderungen und ein letztes Mal in Outlooks abbrechbarem Send-Event aus.

Runtime-Regeln:

- Backend-Signatur-Einfuegung benoetigt eine aktive Backend-Policy fuer die Domain `email_signature`, einen aktiven zugewiesenen Seat, ein nicht leeres `policy.email_signature.email_signature_template` und `policy.email_signature.user_email`.
- Fehlende `policy.email_signature`-Unterstuetzung deaktiviert nur zentrale Signaturen und zeigt einen Backend-Update-Hinweis; Freigabe-/Talk-Policy-Domains bleiben unabhaengig.
- `email_signature_on_compose`, `email_signature_on_reply` und `email_signature_on_forward` sind Backend-Vorgaben, solange der passende `policy_editable`-Wert `true` ist. Ein gespeicherter lokaler Wert darf deshalb eine editierbare Backend-Vorgabe `false` aktivieren. Ein gesperrter Wert (`policy_editable=false`) gewinnt immer, auch ein gesperrtes `false`.
- Die effektive Outlook-Absenderidentitaet muss zu `policy.email_signature.user_email` passen; andere Identitaeten bleiben unberuehrt. Ein `SentOnBehalfOfName`-/Von-Override fuer Shared Mailbox- oder delegierte Exchange-Identitaeten hat Vorrang vor `SendUsingAccount` und muss auf dieselbe SMTP-Adresse aufloesbar sein. Wenn die Absenderidentitaet nicht eindeutig aufgeloest werden kann, bleibt die Signaturverarbeitung fail-closed.
- Neue Mail, Antwort und Weiterleitung verwenden ihre jeweilige effektive Einstellung. Ist Compose-Einfuegung aktiv, aber Antwort oder Weiterleitung deaktiviert, entfernt der passende Absender dort einen exakt erkannten initialen Outlook-Signaturplatz und fuegt keine Backend-Signatur ein.
- Die Compose-Typ-Aufloesung liest zuerst `PR_LAST_VERB_EXECUTED` und danach Conversation-Metadaten. Liefert Outlook nur eine generische Inline-`Response` und unterscheiden sich Antwort-/Weiterleitungswerte, wiederholt die Hintergrundverarbeitung die Pruefung ohne Mutation und das Send-Gate blockiert, statt zu raten. Sind beide Werte gleich, gilt dieser gemeinsame Wert.
- HTML, Plain Text und RTF laufen einheitlich ueber Outlook WordEditor. HTML und RTF importieren das bereinigte Template ueber eine Word-Range; Plain Text verwendet `HtmlToPlainTextConverter` und eine Word-Text-Range. `MailItem.HTMLBody` und `MailItem.Body` werden nicht neu geschrieben; RTF bleibt RTF.
- Inspector und Inline-Compose nutzen denselben Reconciler. `Explorer.InlineResponse`, `Explorer.InlineResponseClose` und `Inspectors.NewInspector` aktualisieren die aktive Oberflaeche, damit eine ausgekoppelte Inline-Antwort im Inspector weiterverarbeitet wird und keinen veralteten Inline-Zustand behaelt.
- Backend-Signatur-HTML laeuft durch `HtmlTemplateSanitizer` mit derselben fail-closed Policy wie Freigabe- und Talk-Templates.
- Der Reconciler sucht das Ziel in dieser Reihenfolge: NC Connectors Bookmark `NcConnectorSignature`, Outlooks Bookmark `_MailAutoSig`, danach einen sicheren strukturellen Einfuegepunkt. Bei einer neuen Mail ist das das Ende des selbst verfassten Dokuments; bei Antwort/Weiterleitung Outlook Words geschuetzte Position zwei Zeichen vor `_MailOriginal` oder, wenn `_MailOriginal` fehlt, ein ueber Word-Absatzrahmen erkannter Zitat-Trenner. Die tatsaechliche `_MailOriginal`-Zitatgrenze bleibt der Bookmark-Start und wird getrennt vom geschuetzten Einfuegeziel gefuehrt; eine Fallback-Einfuegung direkt an dieser Grenze wuerde die Signatur unter Outlooks sichtbaren Trenner setzen.
- Auch ein vorhandener verwalteter oder `_MailAutoSig`-Slot wird nicht blind vertraut: Folgt zwischen seinem Ende und dem sicheren Ziel fuer neue Mail bzw. Zitatgrenze bedeutungsvoller eigener Text, wird der Ersatz am sicheren Ziel vorbereitet und der falsch platzierte alte Slot erst danach entfernt. Die Pruefung eines Antwort-/Weiterleitungs-Slots erfolgt gegen die tatsaechliche Zitatgrenze und nicht gegen das geschuetzte Einfuegeziel; Outlooks native Signaturtabelle darf deshalb exakt an `_MailOriginal` enden und wird dann an Ort und Stelle ersetzt. Ein Slot vollstaendig hinter der tatsaechlichen Grenze wird zum geschuetzten Ziel verschoben; beginnt oder kreuzt ein Slot die tatsaechliche Grenze, bleibt er unveraendert und der Abgleich endet fail-closed. Das deckt direkt passende Antworten ebenso ab wie Identitaetswechsel nach dem Loeschen des urspruenglichen Texts und der Signatur sowie bereits offene Entwuerfe mit falsch platzierter verwalteter Signatur.
- Aktuelle Cursorposition, Nachrichtentext-Anfang, rohe HTML-Prefixe und lokalisierte Antwort-Header sind nie Einfuege- oder Loesch-Fallbacks. Kann bei Antwort/Weiterleitung keine Zitatgrenze gefunden werden, endet der Vorgang ohne Aenderung an selbst verfasstem oder zitiertem Inhalt.
- Tabellenbasierte Outlook-Signaturen werden nur ersetzt, wenn `_MailAutoSig` in dieser Tabelle liegt. Jede erfolgreiche Einfuegung bekommt das Bookmark `NcConnectorSignature`, auch HTML, RTF und Plain Text; spaetere Updates und Clears adressieren nur diese verwaltete Range.
- Die neue Signatur wird vorbereitet, bevor die bisherige Range entfernt wird. Wird oberhalb eines vorhandenen Slots vorbereitet, verfolgt ein temporaeres Word-Bookmark dessen alte Range, waehrend eingefuegter Inhalt die numerischen Positionen verschiebt. Schlagen Einfuegung, Bookmark-Erstellung, Tracking oder das Entfernen der alten Range fehl, wird der vorbereitete Inhalt entfernt und die vorherige verwaltete Range nach Moeglichkeit wiederhergestellt.
- Cursor und Auswahl werden ueber ein temporaeres Word-Bookmark statt ueber veraltete absolute Offsets wiederhergestellt. Der sichere Fallback ergaenzt nur fehlende Absatzmarken oberhalb der Signatur und einen Trennabsatz vor zitiertem Inhalt.
- Absender- und `BodyFormat`-Aenderungen planen einen neuen Abgleich. Aenderungen waehrend unterdrueckter Attachment-Verarbeitung oder ohne aktive Oberflaeche werden zurueckgestellt und fortgesetzt, sobald wieder ein verwendbarer WordEditor vorhanden ist.
- Wird die Policy inaktiv oder passt der Absender nicht mehr, entfernt NC Connector nur `NcConnectorSignature`. Beliebiger Body-Inhalt wird nicht durchsucht oder neu geschrieben; eine native Signatur einer nicht passenden Identitaet bleibt unberuehrt.
- Signaturverarbeitung laeuft nur fuer ungesendete Outlook-Compose-Items. Das Oeffnen einer empfangenen oder bereits gesendeten Nachricht zum Lesen darf den Body nie veraendern.
- Vor dem Senden stoppt Outlook den ausstehenden Debounce und gleicht aktuellen Absender, Format, Compose-Typ, Policy und verwalteten Slot synchron ab. Bei vollstaendigen Backend-Verbindungseinstellungen wird der Versand abgebrochen, wenn kein erfolgreicher Policy-Snapshot vorliegt oder ein erforderlicher finaler Apply-/Clear-Vorgang nicht sicher abgeschlossen werden kann. Das Compose-Item bleibt fuer Korrektur und erneuten Versand offen.
- Eine unvollstaendige Backend-Einrichtung erzeugt keine Signaturpflicht; der Cleanup beschraenkt sich auf das best-effort Entfernen eines exakt gefundenen `NcConnectorSignature`-Bookmarks. Eine nicht unterstuetzte Signatur-Domain deaktiviert ebenfalls die Einfuegung, bei ansonsten vollstaendiger Backend-Einrichtung muss eine vorhandene verwaltete Range vor Send aber weiterhin sicher abgeglichen werden. Ein kurz vor Send eintreffendes `InlineResponseClose` blockiert eine bereits abgeglichene, unveraenderte Mail nicht allein wegen des fehlenden Inline-Editors.
- Der separate Passwort-Follow-up-Dispatch verwendet den bereits beim Hauptversand erfolgreich geprueften Policy-Snapshot, bereinigt ihn einmal und nutzt ihn fuer die gesamte Queue. Er setzt und prueft `SendUsingAccount`/`SentOnBehalfOfName`, sendet nur automatisch, wenn die effektive Follow-up-Identitaet dem beim erfolgreichen Hauptversand erfassten Absender entspricht, und fuegt die Backend-Signatur nur ein, wenn diese effektive Identitaet zusaetzlich zu `policy.email_signature.user_email` passt. Eine Plain-Text-Quelle erzeugt einen Plain-Text-Follow-up; HTML/RTF erzeugt einen HTML-Follow-up. Derselbe Snapshot gilt fuer einen manuellen Fallback-Entwurf.
- Das Debug-Log erfasst Trigger, aktive Oberflaeche, Body-Format, Compose-Typ, Slot-Quelle und Abgleichsergebnis, schreibt aber weder Signatur-Template noch Absenderadresse.

Compose-Filelink-Paritaet (3.1.0):

- Der FileLink-Ribbon-Einstieg ist im Mail-Inspector und im Explorer-Tab `Nachricht` fuer Inline-Antworten/-Weiterleitungen sichtbar. Beide Einstiege laufen ueber denselben `FileLinkLaunchController`.
- Inline-Antworten/-Weiterleitungen fuegen das gerenderte Freigabe-HTML ueber `Explorer.ActiveInlineResponseWordEditor` ein; der Inline-Pfad schreibt nicht direkt in `MailItem.HTMLBody` und behaelt zwei leere Absaetze ueber dem Freigabeblock fuer eigenen Text.
- Normale HTML-Compose-Fenster verwenden zuerst den Inspector-WordEditor, damit verwaltete Bookmarks erhalten bleiben. Nur wenn dieser Editor nicht geoeffnet werden kann, bleibt die direkte `MailItem.HTMLBody`-Route als Kompatibilitaetsfallback aktiv.
- `MailComposeSubscription` in `NextcloudTalkAddIn.cs` steuert den Compose-Lifecycle fuer:
  - debouncte Anhangsauswertung (`ComposeAttachmentEvalDebounceMs`)
  - Pre-Add-Abfangpfad (`BeforeAttachmentAdd`) fuer fruehes Intercept
  - best-effort Abbruch des Host-Adds vor der normalen Outlook-Post-Add-Verarbeitung
  - harte Outlook-/Exchange-Groessenlimits koennen trotzdem vor Add-in-Callbacks greifen und sind ueber offizielle Outlook-OOM-Events nicht abfangbar
  - Always-via-NC und Schwellwertmodus
  - Batch-Entfernung (`Remove last selected attachments`)
  - Attachment-Mode-Wizardstart direkt im Datei-Schritt
  - Share-Cleanup bei unsent close inkl. Grace-Timer fuer Send/Close-Race
  - separates Passwort-Follow-up nach bestaetigtem erfolgreichem Hauptversand; Empfaenger und Absenderkonto werden beim Senden aus dem Original-Compose uebernommen
  - bei Backend-Policy `Nextcloud Secret Link` wird die finale Empfaengerliste aufgeteilt und pro Empfaenger ein eigener einmaliger Secrets-Link erstellt
  - Secrets-Links werden lokal per AES-GCM ueber Windows CNG verschluesselt; es wird keine neue Crypto-Abhaengigkeit gebuendelt
  - wenn Secrets-Erstellung fehlschlaegt, faellt der Versand auf die bisherige separate Klartext-Passwortmail zurueck und zeigt einen Hinweis.
- `ComposeShareLifecycleController` kapselt die eigentliche Share-Cleanup-/Passwort-Dispatch-Logik; `MailComposeSubscription` haelt nur Queue- und Eventzustand.
- `TalkAppointmentController` kapselt Appointment-Schreib-/Sync-Pfade; `NextcloudTalkAddIn` delegiert diese Aufrufe statt die komplette Fachlogik im Root zu halten.
- Nach Appointment-Write werden die lokalen Outlook-`X-NCTALK-*`-Metadaten aktualisiert; serverseitige CalDAV-VEVENTs werden dafuer nicht gepatcht.
- Gespeicherte Talk-Termine stellen die entfernte Raumloeschung nur mit Opt-in (`TalkDeleteRoomOnEventDelete` bzw. Backend-Policy `talk_delete_room_on_event_delete`) und vorhandenen `X-NCTALK-TOKEN`-Metadaten im Hintergrund an; generische Talk-URLs in `Location`/URL-Feldern werden nicht als Loeschquelle ausgewertet.
- Der Cleanup fuer verworfene, noch nicht gespeicherte neue Termine bleibt davon getrennt aktiv.
- Ribbon-getriggerte Flows werden im Controller-Slice gehalten (`SettingsWorkflowController`, `FileLinkLaunchController`, `TalkRibbonController`); `NextcloudTalkAddIn.cs` bleibt schlanke Delegate-/Composition-Root-Schicht.
  - Lifecycle-, Policy-/Template- und Deferred-Ensure-Logik sind in eigene Partial-Dateien ausgelagert, damit die Root-Klasse wartbar bleibt.
  - Custom-Talk-Templates aus dem Backend werden vor HTML-/Plain-Text-Rendering ueber `HtmlTemplateSanitizer` bereinigt (kein Raw-HTML-Fallback).
  - fuer Talk-Termine laeuft vor dem Insert ein expliziter Compat-Transform (`HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(...)`)
  - Appointment-HTML wird ueber HTML->RTF-Bridge geschrieben (`MailItem.HTMLBody` -> `AppointmentItem.RTFBody`), nicht ueber `AppointmentItem.HTMLBody` und nicht ueber `HTMLEditor.body.innerHTML`.
- `OutlookAttachmentAutomationGuardService` erzwingt den Host-Konflikt-Guard live:
  - vor Auswertung
  - vor Prompt-Aktionsverarbeitung
  - vor Wizard-Finalize im Attachment-Modus.
- `Models/AttachmentLinkTargetPolicy.cs` loest `policy.share.attachment_link_target` (`zip_download` / `share_page`) gegen den nullable lokalen Wert auf. Ein ungueltiger gespeicherter lokaler Wert gilt als nicht gesetzt, sodass ein gueltiger editierbarer Backend-Wert ihn vorgeben kann. ZIP gilt nur ohne gueltigen lokalen oder nutzbaren Backend-Wert; ein gesperrter Backend-Wert gewinnt.
- `AttachmentMode` steuert Read-only-Berechtigungen, das Ausblenden der Rechtezeile und Cleanup. Das explizite Linkziel steuert nur URL, `{LINK_INTRO}` und `{LINK_LABEL}`; manuelle Freigaben bleiben immer auf der Nextcloud-Freigabeseite. Im Wizard gibt es keinen Schalter pro Freigabe.
- Die ZIP-URL-Ableitung ist fail-closed: Die absolute oeffentliche HTTP(S)-URL muss auf `/s/<token>` enden und zum OCS-Token passen. Ungueltige Eingaben brechen vor dem Einfuegen ab; es gibt keinen Fallback auf die Original-URL.
- Custom-Share-Templates aus dem Backend werden im `FileLinkHtmlBuilder` vor der Einfuegung ueber `HtmlTemplateSanitizer` bereinigt (fail-closed).
- `{LINK_INTRO}` und `{LINK_LABEL}` werden anhand des effektiven Linkziels aufgeloest. Bestehende Templates ohne diese Platzhalter behalten ihre bisherige Ausgabe.
- Fuer Custom-Share-Templates bevorzugt Outlook `policy.share.share_html_block_template_v2` und faellt auf `policy.share.share_html_block_template` zurueck. Damit funktionieren aeltere Backend-Versionen weiter, waehrend aktuelle Backends den bisherigen Antwortschluessel fuer aeltere Clients platzhalterfrei halten koennen.
- Aktuelle Backends liefern fuer Custom-Templates `policy.share.share_html_block_effective_language`. Outlook verwendet diese Sprache fuer erzeugte Linktexte, Feldbezeichnungen, Berechtigungsnamen und Passworthinweise; bei aelteren Backends ohne dieses Feld bleibt der bisherige Fallback auf die UI-Sprache erhalten.
- Plain-Text-Compose bleibt `MailItem.BodyFormat=olFormatPlain`; der Freigabeblock wird als Textblock mit `#`-Rahmen gerendert und ueber Outlook WordEditor eingefuegt. Inline-Antworten/-Weiterleitungen behalten zwei leere Absaetze ueber dem Block fuer eigenen Text. `MailItem.Body` wird nicht neu geschrieben.
- `FileLinkWizardForm` akzeptiert im Datei-Schritt Explorer-Drag-and-drop fuer Dateien/Ordner ueber Queue und Aktionsbereich.
- `FileLinkService` nutzt fuer Dateien bis 20 MB einen direkten WebDAV-`PUT`. Groessere Dateien laufen ueber Nextcloud Chunked Upload v2 unter `/remote.php/dav/uploads/<user>/<upload-id>` und werden danach per `MOVE .file` an den finalen DAV-Pfad zusammengesetzt.

### Appointment-sicheres HTML-Subset fuer Talk-Templates

Damit Backend-Talk-Templates in Outlook-Terminen stabil gerendert werden (Word/RTF-Pipeline), gilt:

- Layout bevorzugt tabellenbasiert aufbauen (`table`, `tbody`, `tr`, `td`).
- Inline-Styles sind erlaubt, aber Word-kritische CSS-Features werden im Appointment-Compat-Transform entfernt:
  - `display:flex|grid`, `flex*`, `grid*`, `border-radius*`, `overflow*`, `object-fit`, `user-select` (inkl. vendor-prefix Varianten).
- Farbausrichtung bekommt zusaetzliche Legacy-Fallbacks:
  - `style=color` -> `<font color=...>`
  - `style=background-color` -> `bgcolor`
  - `style=text-align` -> `align`
  - `style=vertical-align` -> `valign`
- Linkfarbe wird zusaetzlich abgesichert (`<a><font color=...>...</font></a>`), falls erforderlich.
- Unsichere/nicht erlaubte Tags/Attribute entfernt der Sanitizer weiterhin fail-closed.

Installer:

- `installer/NcConnectorOutlookInstaller.wixproj` (WiX-v6-SDK-Projekt)
- `installer/Product.wxs` (MSI Definition: Dateien + Registry + URLACL)
- `VENDOR.md` (Lizenzhinweise fuer gebuendelte Drittanbieter-Abhaengigkeiten)

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
2) Bei geaenderten vendorten Abhaengigkeiten: `VENDOR.md` aktualisieren
3) `.\build.ps1 -Configuration Release`
4) MSI installieren/upgrade testen (alte Version → neue Version)
5) Talk + Filelink + IFB Smoke-Test
6) MSI ggf. signieren (falls in der Umgebung erforderlich)
