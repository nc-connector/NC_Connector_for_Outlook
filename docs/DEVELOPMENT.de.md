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

## Release 3.0.4 Delta-Ueberblick

Diese Release-Linie ist ein gezieltes Sicherheitsupdate fuer den Talk-Termin-Cleanup:

- Das Loeschen gespeicherter Termine stellt die entfernte Raumloeschung nur noch dann im Hintergrund an, wenn `TalkDeleteRoomOnEventDelete` oder die gesperrte Backend-Policy `talk_delete_room_on_event_delete` es explizit aktiviert.
- Die Runtime-Subscription-Erkennung benoetigt `X-NCTALK-TOKEN`; generische Talk-URLs in `Location` oder URL-Feldern werden nicht als Loeschquelle verwendet.
- Der Cleanup fuer neu erzeugte Termine, die vor dem Speichern verworfen werden, bleibt aktiv und ist vom Opt-in fuer gespeicherte Termine getrennt.
- Talk-Metadaten werden lokal in Outlook als `X-NCTALK-*` UserProperties/MAPI-Felder persistiert. Serverseitige Kalender-`.ics`-Objekte werden nicht gepatcht.

## Voraussetzungen

- Windows 10/11
- Outlook classic (typischerweise x64)
- **.NET Framework 4.7.2** (Target)
- MSBuild (z.B. Visual Studio Build Tools)
- **.NET SDK** (für WiX v4 Build via `dotnet`)

### Reference Assemblies (FrameworkPathOverride)

Auf manchen Build-Systemen fehlen die .NET Framework Reference Assemblies für 4.7.2 (insbesondere CI/Minimal-Installationen). In dem Fall kann man die NuGet-ReferenceAssemblies nutzen und `FrameworkPathOverride` setzen.

Beispiel:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.4"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

## Build (MSI)

Der empfohlene Build läuft immer über `build.ps1`:

```powershell
cd "C:\Users\Bastian\VS-Code\NC-E-T_new\nc4ol-3.0.4"
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
3) Build des WiX v4 Installers (`installer/NcConnectorOutlookInstaller.wixproj`)
4) Kopiert das MSI in `dist/`

## Lokales Testen

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

Controller:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` (Talk-Termin-Lifecycle: Room-Metadaten, Lobby-/Description-/Delegation-/Teilnehmer-Sync)
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs` (Compose-Share-Cleanup und separater Passwort-Versand inkl. Fallback)
- `src/NcTalkOutlookAddIn/Controllers/EmailSignaturePlainTextController.cs` (Plain-Text-Signatur-Einfuegung/-Cleanup ueber Outlook-WordEditor-Signatur-Bookmarks)
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` (Talk-Template-/Block-Rendering)
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` (SMTP- und Attendee-Aufloesung)
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` (Compose-Subscription-Registry)
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` (gemeinsame Mail-/Inspector-Interop-Helper)
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` (gemeinsames Queue-Modell fuer separaten Passwort-Follow-up)

Services:

- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (DAV/Share Flow)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB; Port ueber Settings konfigurierbar, Standard `7777`)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)
- `src/NcTalkOutlookAddIn/Services/NcHttpClient.cs` (zentraler Request-Executor fuer Auth-Header, OCS-Header, Timeout/Decompression und optionalen Fresh-Connection-Mode)
  - Alle Runtime-HTTP-Aufrufe (Talk, Share/DAV, IFB, Login-Flow, Moderator-Avatar-Fetch) laufen zentral ueber `NcHttpClient`.
- `src/NcTalkOutlookAddIn/Services/EmailSignaturePolicyService.cs` (loest Backend-E-Mail-Signatur-Policy gegen lokale Settings und Lock-State auf)

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

Die Compose-Subscription prueft die Backend-Policy fuer die zentrale E-Mail-Signatur nach dem Oeffnen eines Compose-Fensters und wenn Outlook senderbezogene Eigenschaften aendert.

Runtime-Regeln:

- Backend-Signatur-Einfuegung benoetigt eine aktive Backend-Policy fuer die Domain `email_signature`, einen aktiven zugewiesenen Seat, ein nicht leeres `policy.email_signature.email_signature_template` und `policy.email_signature.user_email`.
- Fehlende `policy.email_signature`-Unterstuetzung deaktiviert nur zentrale Signaturen und zeigt einen Backend-Update-Hinweis; Freigabe-/Talk-Policy-Domains bleiben unabhaengig.
- Die effektive Outlook-Absenderidentitaet muss zu `policy.email_signature.user_email` passen; andere Identitaeten bleiben unberuehrt. Ein `SentOnBehalfOfName`-/Von-Override fuer Shared Mailbox- oder delegierte Exchange-Identitaeten hat Vorrang vor `SendUsingAccount` und muss auf dieselbe SMTP-Adresse aufloesbar sein. Wenn die Absenderidentitaet nicht eindeutig aufgeloest werden kann, bleibt die Signaturverarbeitung fail-closed.
- Die lokalen Einstellungen `EmailSignatureOnCompose`, `EmailSignatureOnReply` und `EmailSignatureOnForward` koennen die Einfuegung fuer den jeweiligen Compose-Typ deaktivieren, solange das Backend den Wert nicht sperrt.
- Wenn die Einfuegung beim Verfassen aktiv ist, aber Antworten oder Weiterleitungen deaktiviert sind, entfernt der passende Absender dort trotzdem den initialen lokalen Signaturplatz und fuegt keine Backend-Signatur ein.
- Fuer HTML/RTF-Compose besitzt das passende Absenderkonto auch bei Antworten und Weiterleitungen den initialen Signaturplatz. Beim Compose-Start erkannte Outlook-native oder Drittanbieter-Signaturen werden nur entfernt, wenn die Grenze zur zitierten Nachricht strukturell erkennbar ist; andernfalls bleiben zitierte Nachricht und Trenner erhalten.
- Fuer Plain-Text-Compose bleibt Outlooks Body-Format erhalten. Das bereinigte Backend-HTML wird zu Plain Text gerendert und ueber Outlooks WordEditor-Signaturplatz (`_MailAutoSig`) oder das verwaltete NC-Connector-Word-Bookmark eingefuegt. Plain-Text-Verarbeitung parst keine Antwort-Header und schreibt nicht direkt in `MailItem.Body`.
- Wenn die Compose-Signatur-Policy inaktiv ist oder der Absender nicht passt, entfernt NC Connector nur den eigenen markierten HTML-Signaturblock oder das eigene verwaltete Plain-Text-Word-Bookmark aus dem aktuellen Compose-Item. Outlook-native oder Drittanbieter-Signaturen werden nicht entfernt.
- Backend-Signatur-HTML laeuft durch `HtmlTemplateSanitizer` mit derselben fail-closed Policy wie Freigabe- und Talk-Templates.
- HTML/RTF-Signaturen werden als markierte HTML-Bloecke geschrieben, damit spaetere Policy-/Sender-Aenderungen gezielt nur NC-Connector-eigene Inhalte aktualisieren oder entfernen. Plain-Text-Signaturen werden fuer die offene Compose-Session ueber das verwaltete Word-Bookmark verfolgt.
- Signaturverarbeitung laeuft nur fuer ungesendete Outlook-Compose-Items. Das Oeffnen einer empfangenen oder bereits gesendeten Nachricht zum Lesen darf den Body nie veraendern.
- Separate Passwort-Follow-up-Mails entstehen ohne Compose-Inspector. Vor dem Auto-Versand nutzen sie die `email_signature`-Policy wie eine neue Mail und haengen die Backend-Signatur nur an, wenn der beim urspruenglichen Compose ermittelte Absender zu `policy.email_signature.user_email` passt.
- Inspector-Compose-Fenster nutzen fuer HTML/RTF `MailItem.HTMLBody` und fuer Plain Text WordEditor. Der HTML/RTF-Pfad merkt sich vor dem Body-Rewrite die aktive Word-Auswahl und stellt danach die Auswahl-Schrift wieder her, damit Outlooks aktuelle Schreibschrift aktiv bleibt.
- Inline-Antworten/-Weiterleitungen werden ueber Outlooks `Explorer.InlineResponse`-Event verfolgt und ueber `Explorer.ActiveInlineResponseWordEditor` geschrieben. Inline-Word-Importe verwenden ein UTF-8-BOM-HTML-Dokument, damit Nicht-ASCII-Zeichen in Signaturen erhalten bleiben.
- Die Antwort-/Weiterleitungs-Erkennung nutzt zuerst `PR_LAST_VERB_EXECUTED`. Wenn Outlook den Wert noch nicht gesetzt hat, nutzen ausgekoppelte Antworten/Weiterleitungen `PR_CONVERSATION_INDEX`; Inline-Antworten/-Weiterleitungen gelten ueber `Explorer.InlineResponse` als Antwort.
- Inline-HTML/RTF-Ersetzung nutzt Outlooks versteckte Word-Bookmarks `_MailAutoSig` und `_MailOriginal`. Wenn `_MailOriginal` fehlt, wird der Trenner zur zitierten Nachricht ueber Word-Absatzrahmen erkannt. Tabellenbasierte Outlook- oder Drittanbieter-Signaturen werden mit `Word.Table.Delete()` entfernt, wenn das Signatur-Bookmark in einer Tabelle liegt. Reine Textmarker wie `From:` oder `Von:` werden nie als Schnittposition benutzt.
- Die HTML/RTF-Signatur ersetzt den Compose-Signaturplatz vor der zitierten Nachrichtengrenze, behaelt zwei leere Absaetze ueber der Signatur fuer eigenen Text und eine leere Zeile zwischen Signatur und Antwort-/Weiterleitungs-Trenner.
- Beim reinen Entfernen der Signatur in Antworten/Weiterleitungen bleibt oberhalb des Zitats ein leerer Schreibbereich erhalten; der Word-Cursor wird dorthin zurueckgesetzt.

Compose-Filelink-Paritaet (3.0.4):

- Der FileLink-Ribbon-Einstieg ist im Mail-Inspector und im Explorer-Tab `Nachricht` fuer Inline-Antworten/-Weiterleitungen sichtbar. Beide Einstiege laufen ueber denselben `FileLinkLaunchController`.
- Inline-Antworten/-Weiterleitungen fuegen das gerenderte Freigabe-HTML ueber `Explorer.ActiveInlineResponseWordEditor` ein; der Inline-Pfad schreibt nicht direkt in `MailItem.HTMLBody` und behaelt zwei leere Absaetze ueber dem Freigabeblock fuer eigenen Text.
- `MailComposeSubscription` in `NextcloudTalkAddIn.cs` steuert den Compose-Lifecycle fuer:
  - debouncte Anhangsauswertung (`ComposeAttachmentEvalDebounceMs`)
  - Pre-Add-Abfangpfad (`BeforeAttachmentAdd`) fuer fruehes Intercept
  - best-effort Abbruch des Host-Adds vor der normalen Outlook-Post-Add-Verarbeitung
  - harte Outlook-/Exchange-Groessenlimits koennen trotzdem vor Add-in-Callbacks greifen und sind ueber offizielle Outlook-OOM-Events nicht abfangbar
  - Always-via-NC und Schwellwertmodus
  - Batch-Entfernung (`Remove last selected attachments`)
  - Attachment-Mode-Wizardstart direkt im Datei-Schritt
  - Share-Cleanup bei unsent close inkl. Grace-Timer fuer Send/Close-Race
  - separates Passwort-Follow-up nach bestaetigtem erfolgreichem Hauptversand; Empfaenger und Absenderkonto werden beim Senden aus dem Original-Compose uebernommen.
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
- `FileLinkHtmlBuilder` erzeugt im Attachment-Modus reduziertes HTML mit ZIP-Link `/s/<token>/download`.
- Custom-Share-Templates aus dem Backend werden im `FileLinkHtmlBuilder` vor der Einfuegung ueber `HtmlTemplateSanitizer` bereinigt (fail-closed).
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

- `installer/NcConnectorOutlookInstaller.wixproj` (WiX v4 SDK Projekt)
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

