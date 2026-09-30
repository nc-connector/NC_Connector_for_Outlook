# Entwicklungsleitfaden — NC Connector für Outlook

Dieses Dokument richtet sich an Entwickler und beschreibt Aufbau, Build und Release-Prozess des **NC Connector for Outlook** (Outlook classic COM Add-in).

Rollout, Konfiguration, Betriebsprüfungen und Störungs-Runbooks für Administratoren stehen in [ADMIN.de.md](ADMIN.de.md).

## Inhalt

- [Projektzweck](#projektzweck)
- [Schnellstart](#schnellstart)
- [Repository-Struktur](#repository-struktur)
- [Architektur](#architektur)
- [Netzwerk-Endpunkte](#netzwerk-endpunkte)
- [Lokalisierung (i18n)](#lokalisierung-i18n)
- [Logging](#logging)
- [Kompatibilität und Versionsprüfungen](#kompatibilität-und-versionsprüfungen)
- [Build und Release](#build-und-release)
- [Lokales Testen](#lokales-testen)
- [Referenz der X-NCTALK-*-Eigenschaften](#referenz-der-x-nctalk--eigenschaften)
- [Erweiterungspunkte](#erweiterungspunkte)

## Projektzweck

Das Add-in integriert:

- **Nextcloud Talk** direkt aus dem Termin (Raum erstellen, Lobby, Moderator-Delegation, Teilnehmer-Automation)
- **Nextcloud-Freigaben** im E-Mail-Composer (Upload, Linkfreigabe und Einfügen des HTML- oder Textblocks)
- **Zentrale Backend-E-Mail-Signaturen** fuer passende Outlook-Absenderkonten
- **IFB (Internet Free/Busy)** als lokaler HTTP-Proxy zu Nextcloud

## Schnellstart

### Voraussetzungen

- Windows 10 oder Windows 11 (64-Bit)
- Outlook classic (x64 oder x86)
- **.NET Framework 4.7.2** (Target)
- MSBuild (z.B. Visual Studio Build Tools)
- **.NET SDK** (für den WiX-v6-Build via `dotnet`)
- **Nextcloud 32 oder neuer** (Laufzeit-Server)

**Referenz-Assemblies (`FrameworkPathOverride`)**

Auf manchen Build-Systemen fehlen die .NET Framework Reference Assemblies für 4.7.2 (insbesondere CI/Minimal-Installationen). In dem Fall kann man die NuGet-ReferenceAssemblies nutzen und `FrameworkPathOverride` setzen.

Beispiel:

```powershell
cd "C:\Pfad\zum\nc4ol"

# Optional: Reference Assemblies lokal holen (nur wenn nötig)
nuget install Microsoft.NETFramework.ReferenceAssemblies.net472 -OutputDirectory packages -ExcludeVersion

$env:FrameworkPathOverride = "$PWD\packages\Microsoft.NETFramework.ReferenceAssemblies.net472\build\.NETFramework\v4.7.2"
```

### MSI bauen (empfohlen)

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

### Lokal installieren und starten

1. MSI mit Administratorrechten installieren:
   - `msiexec /i dist\NCConnectorForOutlook-<version>.msi`
2. Outlook starten.
3. Die Funktionen über das Ribbon öffnen:
   - Termin: **NC Connector → Talk-Link einfügen**
   - E-Mail: **NC Connector → Nextcloud Freigabe hinzufügen**
   - Inline-Antwort oder -Weiterleitung: **Nachricht → NC Connector → Nextcloud Freigabe hinzufügen**
4. Unter **NC Connector → Einstellungen** die Server-URL und Zugangsdaten konfigurieren.

## Repository-Struktur

- `src/` — COM-Add-in mit WinForms-Oberfläche und Services
- `installer/` — WiX-v6-MSI-Projekt
- `docs/` — Betriebs- und Entwicklungsdokumentation
- `VENDOR.md` — Hinweise und Lizenzen für gebündelte Drittanbieterkomponenten
- `assets/` — Branding-Ressourcen
- `dist/` — erzeugte MSI-Dateien

## Architektur

### Zentrale Bausteine

Root:

- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs`  
  Einstiegspunkt, Ribbon, Outlook-Events, Composition Root fuer die Workflows.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Lifecycle.cs`  
  Add-in-Bootstrap/Teardown (`OnConnection`, Shutdown/Disconnect).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Hooks.cs`
  Dedizierte Outlook-Event Hook-/Unhook-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.CalendarSelection.cs`
  Rebind ausgewählter Termine für das Löschen aus der Kalenderansicht ohne Kalenderscan.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.Logging.cs`
  Kategorienspezifische Runtime-Logging-Helper.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.PolicyTemplates.cs`  
  Backend-Policy- und Talk-Template-/Sprach-Resolver.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.SubscriptionEnsure.cs`  
  Deferred Appointment-Subscription-Ensure inkl. Outlook-Event-Restriction-Handling.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs`
  Runtime-Subscription-Core fuer Compose-Lifecycle-Zustand (`Dispose`, Identity, gemeinsame Helper).
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs`
  Anhangsereignisse, Timer und Prompt-Orchestrierung.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentPolicy.cs`
  Policy-Snapshots, Aktualisierung und Versandprüfung für Anhänge.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentMaterialization.cs`
  Anhangs-Snapshots, lokale Dateien und Entfernung aus Outlook.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.AttachmentQueue.cs`
  Queue-Übergabe, Before-add-Batches und Abschluss der Ereignisunterdrückung.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Signature.cs`
  Backend-E-Mail-Signatur-Policy fuer das passende Outlook-Absenderkonto.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.Send.cs`
  Send-Gate und dauerhaftes Vorbereiten separater Passwortmails.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs`
  `AfterWrite`-, `Inspector.Close`- und Inline-`Unload`-Behandlung für neu eingefügte Freigaben.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.ComposeLifecycle.cs`
  Composition-Root-Brücke für asynchron angestoßene Freigabebereinigung und bestätigten Passwortversand.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.AppointmentSubscription.cs`
  Runtime-Subscription fuer Termin-Write/Close/Delete und Lifecycle-Cleanup.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.TalkAppointmentSync.cs`
  STA-Erfassung und Hintergrundverarbeitung von Terminänderungen.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.TalkRoomLifecycle.cs`
  Start, Retry-Wiederherstellung und Teardown für vorgemerkte Talk-Raum-Löschungen.
- `src/NcTalkOutlookAddIn/Controllers/SettingsWorkflowController.cs`
  Orchestrierung fuer Settings-Open/Save/Revert.
- `src/NcTalkOutlookAddIn/Controllers/FileLinkLaunchController.cs`
  Orchestrierung fuer FileLink-Ribbon-Start und Wizard-Flow.
- `src/NcTalkOutlookAddIn/Controllers/TalkRibbonController.cs`
  Orchestrierung fuer Talk-Ribbon-Flow (Auth-Gate, Wizard, Room-Create/Replace).
- `TalkRibbonController` lädt Backend-Policy und Passwort-Policy vor dem Wizard. `FileLinkLaunchController` lädt zusätzlich den erforderlichen Capability-Snapshot und übergibt ihn an den Freigabe-Wizard. Policy-Daten bleiben pro Einstiegspunkt frisch.
- Nach dem FileLink-, Talk- oder Settings-Prefetch wechselt `OutlookUiSynchronizationContext` vor jedem WinForms- oder Outlook-COM-Zugriff zurueck auf den beim Add-in-Start erfassten Outlook-STA-Thread. Outlook stellt COM-Callbacks nicht verlaesslich einen `SynchronizationContext` bereit; modale Dialoge und Outlook-Interop werden deshalb explizit zurueckgeschaltet.

Controller:

- `src/NcTalkOutlookAddIn/Controllers/TalkAppointmentController.cs` mit `Sync`-Partial (Terminmetadaten, lokaler Snapshot und entfernte Raumaktualisierung)
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareCleanupTracker.cs` (In-Memory-Status neu eingefügter, noch nicht geschriebener Compose-Freigaben)
- `src/NcTalkOutlookAddIn/Controllers/SeparatePasswordDeliveryController.cs` (Body-, Empfänger-, Absender-, Secrets- und Signaturaufbereitung sowie direkte Outlook-Übergabe der Passwortmail)
- `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs` (Talk-Template-/Block-Rendering)
- `src/NcTalkOutlookAddIn/Controllers/OutlookRecipientResolverController.cs` (SMTP- und Attendee-Aufloesung)
- `src/NcTalkOutlookAddIn/Controllers/MailComposeSubscriptionRegistryController.cs` (Compose-Subscription-Registry)
- `src/NcTalkOutlookAddIn/Controllers/MailInteropController.cs` (aktive Mail, Inline-Compose, Inspector-Identität und Dialogbesitzer)
- `src/NcTalkOutlookAddIn/Controllers/MailBodyInsertionController.cs` (Freigabeblöcke über die bestehenden WordEditor- und HTMLBody-Pfade)
- `src/NcTalkOutlookAddIn/Controllers/ManagedEmailSignatureController.cs` (gemeinsamer WordEditor-Signatur-Slot-Reconciler)
- `src/NcTalkOutlookAddIn/Controllers/AppointmentHtmlBodyWriter.cs` (HTML-zu-RTF-Brücke für Termine)
- `src/NcTalkOutlookAddIn/Models/SeparatePasswordDispatchEntry.cs` (gemeinsames Queue-Modell fuer separaten Passwort-Follow-up)
- `src/NcTalkOutlookAddIn/Settings/ManagedSetupPolicy.cs` (verwaltete Nextcloud-URL, Anmeldeart, Ribbon-Sichtbarkeit, Standardwertequelle sowie Transport-TLS-, Logging-, Update-Benachrichtigungs- und IFB-Policies aus Registry/GPO)
- `src/NcTalkOutlookAddIn/Settings/SettingsFileTransaction.cs` (serialisiert Profil-Schreibvorgänge über einen benannten Mutex, ersetzt nur validierte Dateien und behält die letzte gültige Sicherung)

Services:

- `src/NcTalkOutlookAddIn/Services/ComposeShareCleanupService.cs` (Löschung nicht persistierter oder nicht eingefügter Serverartefakte über den exakten Ursprung)
- `src/NcTalkOutlookAddIn/Services/TalkService.cs` (Talk API Calls)
- `src/NcTalkOutlookAddIn/Services/FileLinkService.cs` (Orchestrierung von Uploadplan, Freigabe-Stammordner, Transfer und Share-Erstellung)
- `src/NcTalkOutlookAddIn/Services/FileLinkQueueSnapshotBuilder.cs` (quellengruppierter Ordnerbaum für die Freigabe-Warteschlange)
- `src/NcTalkOutlookAddIn/Services/FileLinkDavClient.Browsing.cs` (DAV-Hierarchie und Speicherwerte des Benutzers) sowie `FileLinkDavClient.Copy.cs` (Kopieren ausgewählter Nextcloud-Dateien in die Freigabe)
- `src/NcTalkOutlookAddIn/Services/NextcloudCapabilitiesService.cs` (globale Nextcloud-32-Prüfung und typisierter OCS-Capabilities-Snapshot mit fünf Minuten Cache)
- `src/NcTalkOutlookAddIn/Services/FileLinkSelectionScanner.cs` (einmaliger lokaler Scan und Pfade relativ zum Freigabe-Stammordner)
- `src/NcTalkOutlookAddIn/Services/FileLinkUploadPlanner.cs` (Auswahl von Direct, Chunked oder optionalem Bulk vor der ersten serverseitigen Änderung)
- `src/NcTalkOutlookAddIn/Services/FileLinkUploadPlan.cs` (Modelle des fertigen Uploadplans)
- `src/NcTalkOutlookAddIn/Services/FileLinkDavClient.cs` mit `Probes`- und `Requests`-Partials (DAV-Verzeichnis-Lebenszyklus, exakte Ressourcenprüfungen, Löschung, Wiederholungen, Fehlerabbildung und URL-Aufbau)
- `src/NcTalkOutlookAddIn/Services/FileLinkTransferService.cs` mit `FileLinkBulkUploader`, `FileLinkDirectUploader`, `FileLinkChunkUploader` und `FileLinkSourceFile` (Transferkoordination, Protokollpfade und Prüfung der lokalen Quelldatei)
- `src/NcTalkOutlookAddIn/Services/FileLinkShareClient.cs` mit `Recovery`-Partial (öffentliche Freigabe mit einem OCS-Request und Prüfung unklarer Erstellergebnisse)
- `src/NcTalkOutlookAddIn/Services/FileLinkUploadProgress.cs` (aggregierter Phasen- und Transferfortschritt mit Begrenzung der Aktualisierungsrate)
- `src/NcTalkOutlookAddIn/Services/FreeBusyServer.cs` + `FreeBusyManager.cs` (IFB; Port ueber Settings konfigurierbar, Standard `7777`)
- `src/NcTalkOutlookAddIn/Services/PasswordPolicyService.cs` (Nextcloud Password Policy + Fallback)
- `src/NcTalkOutlookAddIn/Services/NcHttpClient.cs` (zentraler Request-Executor fuer Auth-Header, OCS-Header, Timeout/Decompression und optionalen Fresh-Connection-Mode)
  - Alle Runtime-HTTP-Aufrufe (Talk, Share/DAV, IFB, Login-Flow, Moderator-Avatar-Fetch) laufen zentral ueber `NcHttpClient`.
- `src/NcTalkOutlookAddIn/Services/EmailSignaturePolicyService.cs` (loest Backend-E-Mail-Signatur-Policy gegen lokale Settings und Lock-State auf)
- `src/NcTalkOutlookAddIn/Services/UpdateCheckService.cs` (fragt einmal pro Tag `nc-connector.de` nach Outlook-Release-Metadaten und speichert das Ergebnis in den Profil-Settings)
- `src/NcTalkOutlookAddIn/Services/TalkAppointmentSyncCoordinator.cs` (fasst entfernte Talk-Aktualisierungen aus Outlook-Ereignissen zusammen)
- `src/NcTalkOutlookAddIn/Services/TalkRoomLifecycleCoordinator.cs` und `TalkRoomLifecycleStore.cs` (Persistenz und Wiederholung vorgemerkter Raumlöschungen ohne Outlook-Kalenderscan)
- `src/NcTalkOutlookAddIn/Services/IfbRegistryOwnershipManager.cs` und `IfbRegistryStateStore.cs` (IFB-Registry-Wiederherstellung); `FreeBusyServer.cs` prüft den geheimen Anfragepfad und begrenzt parallele Anfragen.
- `src/NcTalkOutlookAddIn/Services/ProtectedJsonStateStore.cs` stellt den gemeinsamen Pfad für DPAPI-geschütztes JSON und Backup-Wiederherstellung der Talk-Löschwarteschlange und des IFB-Registry-Besitzstatus bereit. Schreibvorgänge bevorzugen den atomaren Austausch und behalten den bestehenden Copy-Fallback; die typisierten Stores behalten ihre fachlichen Dateinamen, Entropie, Validierung und Diagnosemeldungen.

Update-Check:

- Endpoint: `GET https://nc-connector.de/wp-json/ncc/v1/update-check`
- Gesendet werden Produkt, installierte Version, Kanal und ein taeglich wechselnder anonymer Client-Hash.
- Downloads zeigen direkt auf GitHub-Release-Dateien; die Homepage zaehlt nur taegliche Installationen und liefert Release-Metadaten.

UI:

- `src/NcTalkOutlookAddIn/UI/SettingsForm.cs`
  Die Partials `General`, `Sharing`, `Talk`, `Signature`, `Ifb`, `Advanced`, `Debug` und `About` enthalten Layout und Aktionen ihres Tabs. Control-Erzeugung, gemeinsame Speichervalidierung, Backend-Policy-Zustand und tabübergreifende Control-Aktualisierung bleiben in der Hauptform; das bestehende `Language`-Partial enthält weiterhin die Lokalisierung.
  Die Sprachauswahl des Freigabeblocks gehört zu `Sharing`, die Sprache des Talk-Beschreibungstextes zu `Talk`. Beide verwenden die bisherigen Auswahlhandler, Policy-Sperren und gespeicherten Werte; `Advanced` enthält und positioniert keine Sprachauswahl mehr.
- `src/NcTalkOutlookAddIn/UI/TalkLinkForm.cs`
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
- `src/NcTalkOutlookAddIn/UI/NextcloudFilePickerForm.cs` (Datei- und Ordnerauswahl für **Meine Nextcloud**)
- `src/NcTalkOutlookAddIn/UI/NextcloudPickerNavigation.cs` hält den Verlauf. Die Form übernimmt eine neue Verlaufsposition erst nach erfolgreichem Laden des Ordners.
- `src/NcTalkOutlookAddIn/UI/NextcloudPickerPreview.cs` lädt und dekodiert eine Dateivorschau. Auswahl, Abbruchgeneration und Besitz des angezeigten Bildes bleiben bei der Form; Anfragen laufen weiterhin über `FileLinkService`.
- `src/NcTalkOutlookAddIn/UI/ComposeAttachmentPromptForm.cs` (2-Aktions-Prompt fuer Schwellwertmodus)
- `src/NcTalkOutlookAddIn/UI/BrandedHeader.cs` (Header-Banner inkl. `AttachToParent(...)` fuer konsistente Header-Initialisierung in Forms)
- `src/NcTalkOutlookAddIn/UI/ScaledForm.cs` (zentrale DPI-Skalierung via `ScaleLogical(...)`, damit Form-Wrapper nicht dupliziert werden)

Utilities:

- `src/NcTalkOutlookAddIn/Utilities/RecipientAddressList.cs` (gemeinsame Normalisierung, Deduplizierung und Semikolonlisten für Empfänger)
- `src/NcTalkOutlookAddIn/Utilities/WordHtmlInsertionFile.cs` (gemeinsame HTML-Dokumenthülle und Bereinigung temporärer Dateien für Word-Einfügungen)
- `src/NcTalkOutlookAddIn/Utilities/BrowserLauncher.cs` (zentraler Shell-Start für Dateien und Ordner; `OpenUrl` lehnt Nicht-HTTPS-Ziele ab)
- `src/NcTalkOutlookAddIn/Utilities/SizeFormatting.cs` (zentrale MB-Formatierung fuer UI-Texte)
- `src/NcTalkOutlookAddIn/Utilities/ComInteropScope.cs` (zentrale COM-Release-/FinalRelease-Helfer)
- `src/NcTalkOutlookAddIn/Utilities/PasswordGenerationHelper.cs` (zentralisiert Min-Length-Aufloesung, Server-Fallback-Generierung und gemeinsame Min-Length-Validierung fuer Talk/FileLink-Formulare)
- `src/NcTalkOutlookAddIn/Utilities/FileLinkPath.cs` (zentrale Normalisierung, Kombination, Benennung, Bereinigung und Tiefenberechnung für FileLink-Pfade)
- `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs` (zentrale HtmlSanitizer-9.0.892-Policy für Backend-HTML bei Freigaben und Talk; aktive Inhaltscontainer wie `template` werden entfernt)
- `src/NcTalkOutlookAddIn/Utilities/HtmlToPlainTextConverter.cs` (DOM-basierte HTML-zu-Plain-Text-Ausgabe fuer Plain-Text-E-Mail-Signaturen)
- `src/NcTalkOutlookAddIn/Utilities/NcJson.cs` (zentrale JSON-Normalisierung inkl. `PrepareJsonPayload`, Dictionary-/String-/Int-Helfer und OCS-Fehlerextraktion)
- `src/NcTalkOutlookAddIn/Utilities/DeferredAppointmentEnsureState.cs` (gekapselter Laufzeitzustand fuer Pending-Keys und Restriction-Log-Throttling)
- `src/NcTalkOutlookAddIn/Utilities/NextcloudUriValidator.cs` (HTTPS-Basis-URL-Prüfung und Same-Origin-Prüfung für serverseitig gelieferte Endpunkte)
- `src/NcTalkOutlookAddIn/Utilities/PictureConverter.cs` (gemeinsamer Image->IPictureDisp-Helfer fuer Ribbon-Icons)

#### Zentrale E-Mail-Signatur im Compose-Fenster

Die Compose-Subscription wertet die zentrale Signatur nach dem Oeffnen einer Compose-Oberflaeche, nach Absender- oder BodyFormat-Aenderungen und ein letztes Mal in Outlooks abbrechbarem Send-Event aus.

Runtime-Regeln:

- Backend-Signatur-Einfuegung benoetigt eine aktive Backend-Policy fuer die Domain `email_signature`, einen aktiven zugewiesenen Seat, ein nicht leeres `policy.email_signature.email_signature_template` und `policy.email_signature.user_email`.
- Fehlende `policy.email_signature`-Unterstuetzung deaktiviert nur zentrale Signaturen und zeigt einen Backend-Update-Hinweis; Freigabe-/Talk-Policy-Domains bleiben unabhaengig.
- `email_signature_on_compose`, `email_signature_on_reply` und `email_signature_on_forward` folgen bei passendem `policy_editable=true` der effektiven Standardwertequelle: Die lokale Quelle bevorzugt eine gespeicherte lokale Auswahl, die Backend-Quelle den Backend-Startwert. Ein gesperrter Wert (`policy_editable=false`) gewinnt immer, auch ein gesperrtes `false`. Die Signaturvorlage bleibt bei beiden Quellen ausschließlich im Backend.
- Die effektive Outlook-Absenderidentitaet muss zu `policy.email_signature.user_email` passen; andere Identitaeten bleiben unberuehrt. Ein `SentOnBehalfOfName`-/Von-Override fuer Shared Mailbox- oder delegierte Exchange-Identitaeten hat Vorrang vor `SendUsingAccount` und muss auf dieselbe SMTP-Adresse aufloesbar sein. Wenn die Absenderidentitaet nicht eindeutig aufgeloest werden kann, bleibt die Signaturverarbeitung fail-closed.
- Neue Mail, Antwort und Weiterleitung verwenden ihre jeweilige effektive Einstellung. Ist Compose-Einfuegung aktiv, aber Antwort oder Weiterleitung deaktiviert, entfernt der passende Absender dort einen exakt erkannten initialen Outlook-Signaturplatz und fuegt keine Backend-Signatur ein.
- Die Compose-Typ-Aufloesung liest zuerst `PR_LAST_VERB_EXECUTED` und danach Conversation-Metadaten. Liefert Outlook nur eine generische Inline-`Response` und unterscheiden sich Antwort-/Weiterleitungswerte, wiederholt die Hintergrundverarbeitung die Pruefung ohne Mutation und das Send-Gate blockiert, statt zu raten. Sind beide Werte gleich, gilt dieser gemeinsame Wert.
- HTML, Plain Text und RTF laufen einheitlich ueber Outlook WordEditor. HTML und RTF importieren das bereinigte Template ueber eine Word-Range; Plain Text verwendet `HtmlToPlainTextConverter` und eine Word-Text-Range. `MailItem.HTMLBody` und `MailItem.Body` werden nicht neu geschrieben; RTF bleibt RTF.
- Inspector und Inline-Compose nutzen denselben Reconciler. `Explorer.InlineResponse`, `Explorer.InlineResponseClose` und `Inspectors.NewInspector` aktualisieren die aktive Oberflaeche, damit eine ausgekoppelte Inline-Antwort im Inspector weiterverarbeitet wird und keinen veralteten Inline-Zustand behaelt.
- Backend-Signatur-HTML laeuft durch `HtmlTemplateSanitizer` mit derselben fail-closed Policy wie Freigabe- und Talk-Templates.
- Der Reconciler sucht das Ziel in dieser Reihenfolge: NC Connectors Bookmark `NcConnectorSignature`, Outlooks Bookmark `_MailAutoSig`, danach einen sicheren strukturellen Einfuegepunkt. Bei einer neuen Mail ist das das Ende des selbst verfassten Dokuments; bei Antwort/Weiterleitung Outlook Words geschuetzte Position zwei Zeichen vor `_MailOriginal` oder, wenn `_MailOriginal` fehlt, ein ueber Word-Absatzrahmen erkannter Zitat-Trenner. Die tatsaechliche `_MailOriginal`-Zitatgrenze bleibt der Bookmark-Start und wird getrennt vom geschuetzten Einfuegeziel gefuehrt; eine Fallback-Einfuegung direkt an dieser Grenze wuerde die Signatur unter Outlooks sichtbaren Trenner setzen.
- Auch ein vorhandener verwalteter oder `_MailAutoSig`-Slot wird nicht blind vertraut: Folgt zwischen seinem Ende und dem sicheren Ziel fuer neue Mail bzw. Zitatgrenze bedeutungsvoller eigener Text, wird der Ersatz am sicheren Ziel vorbereitet und der falsch platzierte alte Slot erst danach entfernt. Die Pruefung eines Antwort-/Weiterleitungs-Slots erfolgt gegen die tatsaechliche Zitatgrenze und nicht gegen das geschuetzte Einfuegeziel; Outlooks native Signaturtabelle darf deshalb exakt an `_MailOriginal` enden und wird dann an Ort und Stelle ersetzt. Ein Slot vollstaendig hinter der tatsaechlichen Grenze wird zum geschuetzten Ziel verschoben; beginnt oder kreuzt ein Slot die tatsaechliche Grenze, bleibt er unveraendert und der Abgleich endet fail-closed. Das deckt direkt passende Antworten ebenso ab wie Identitaetswechsel nach dem Loeschen des urspruenglichen Texts und der Signatur sowie bereits offene Entwuerfe mit falsch platzierter verwalteter Signatur.
- Aktuelle Cursorposition, Nachrichtentext-Anfang, rohe HTML-Prefixe und lokalisierte Antwort-Header sind nie Einfuege- oder Loesch-Fallbacks. Kann bei Antwort/Weiterleitung keine Zitatgrenze gefunden werden, endet der Vorgang ohne Aenderung an selbst verfasstem oder zitiertem Inhalt.
- Der Absatzrahmen-Fallback fuer den Zitat-Trenner schliesst jeden Absatz innerhalb einer Word-Tabelle aus, einschliesslich verschachtelter Berechtigungszellen einer eingefuegten Freigabe. Vor der Rahmenpruefung liest er `Range.Information(wdWithInTable)`; ist die Tabellenzuordnung nicht lesbar, endet der Abgleich fail-closed. Ein Tabellenrahmen darf die Signatur beim Sendeabgleich niemals in die Freigabe verschieben; der `_MailOriginal`-Bookmark-Pfad bleibt unveraendert.
- Tabellenbasierte Outlook-Signaturen werden nur ersetzt, wenn `_MailAutoSig` in dieser Tabelle liegt. Jede erfolgreiche Einfuegung bekommt das Bookmark `NcConnectorSignature`, auch HTML, RTF und Plain Text; spaetere Updates und Clears adressieren nur diese verwaltete Range.
- Die neue Signatur wird vorbereitet, bevor die bisherige Range entfernt wird. Wird oberhalb eines vorhandenen Slots vorbereitet, verfolgt ein temporaeres Word-Bookmark dessen alte Range, waehrend eingefuegter Inhalt die numerischen Positionen verschiebt. Schlagen Einfuegung, Bookmark-Erstellung, Tracking oder das Entfernen der alten Range fehl, wird der vorbereitete Inhalt entfernt und die vorherige verwaltete Range nach Moeglichkeit wiederhergestellt.
- Cursor und Auswahl werden ueber ein temporaeres Word-Bookmark statt ueber veraltete absolute Offsets wiederhergestellt. Der sichere Fallback ergaenzt nur fehlende Absatzmarken oberhalb der Signatur und einen Trennabsatz vor zitiertem Inhalt.
- Absender- und `BodyFormat`-Aenderungen planen einen neuen Abgleich. Aenderungen waehrend unterdrueckter Attachment-Verarbeitung oder ohne aktive Oberflaeche werden zurueckgestellt und fortgesetzt, sobald wieder ein verwendbarer WordEditor vorhanden ist.
- Wird die Policy inaktiv oder passt der Absender nicht mehr, entfernt NC Connector nur `NcConnectorSignature`. Beliebiger Body-Inhalt wird nicht durchsucht oder neu geschrieben; eine native Signatur einer nicht passenden Identitaet bleibt unberuehrt.
- Signaturverarbeitung laeuft nur fuer ungesendete Outlook-Compose-Items. Das Oeffnen einer empfangenen oder bereits gesendeten Nachricht zum Lesen darf den Body nie veraendern.
- Vor dem Senden stoppt Outlook den ausstehenden Debounce und gleicht aktuellen Absender, Format, Compose-Typ, Policy und verwalteten Slot synchron ab. Bei vollstaendigen Backend-Verbindungseinstellungen wird der Versand abgebrochen, wenn kein erfolgreicher Policy-Snapshot vorliegt oder ein erforderlicher finaler Apply-/Clear-Vorgang nicht sicher abgeschlossen werden kann. Das Compose-Item bleibt fuer Korrektur und erneuten Versand offen.

- Eine fehlende Policy verwendet `email_signature_policy_unavailable`, ein fehlgeschlagener finaler Apply-/Clear-Vorgang `email_signature_send_reconcile_failed`. Beide brechen den Versand mit einer verständlichen Handlungsanweisung ab, ohne interne Fehlerbezeichner anzuzeigen. Bei fehlgeschlagenem Refresh bleibt ein passender erfolgreicher Snapshot nutzbar; eine frische erfolgreiche Antwort, auch eine Ablehnung, ersetzt ihn. Im Outlook-Send-Event kommt kein Netzwerkaufruf hinzu.
- Eine unvollstaendige Backend-Einrichtung erzeugt keine Signaturpflicht; der Cleanup beschraenkt sich auf das best-effort Entfernen eines exakt gefundenen `NcConnectorSignature`-Bookmarks. Eine nicht unterstuetzte Signatur-Domain deaktiviert ebenfalls die Einfuegung, bei ansonsten vollstaendiger Backend-Einrichtung muss eine vorhandene verwaltete Range vor Send aber weiterhin sicher abgeglichen werden. Ein kurz vor Send eintreffendes `InlineResponseClose` blockiert eine bereits abgeglichene, unveraenderte Mail nicht allein wegen des fehlenden Inline-Editors.
- Der separate Passwort-Follow-up-Dispatch erfasst den erfolgreichen Policy- und Settings-Snapshot bereits beim Erstellen der Freigabe. Im `Send`-Ereignis der Hauptmail setzt und prueft er `SendUsingAccount`/`SentOnBehalfOfName`, übergibt nur automatisch, wenn die effektive Follow-up-Identitaet dem erfassten Absender der Hauptmail entspricht, und fuegt die Backend-Signatur nur ein, wenn diese Identitaet zusätzlich zu `policy.email_signature.user_email` passt. Eine Plain-Text-Quelle erzeugt einen Plain-Text-Follow-up; HTML/RTF erzeugt einen HTML-Follow-up. Bei eindeutig fehlgeschlagener automatischer Übergabe zeigt Outlook die vollständig vorbereitete Nachricht an; ein mehrdeutiger Übermittlungsstatus wird nie wiederholt.
- Das Debug-Log erfasst Trigger, aktive Oberflaeche, Body-Format, Compose-Typ, Slot-Quelle und Abgleichsergebnis, schreibt aber weder Signatur-Template noch Absenderadresse.

### Laufzeitkonfiguration und Policy-Verarbeitung

`AuthMode` ist eine optionale `REG_SZ`-Policy mit `LoginFlow` oder `Manual`. Sie wird unabhängig nach der bestehenden Hive-/View-Reihenfolge aufgelöst; der erste vorhandene Wert gewinnt, auch bei ungültigem Inhalt oder Typ. Ihre Existenz aktiviert Enterprise Rollout und sperrt beide Authentifizierungs-Radiobuttons. Ungültige ausgewählte Eingaben verwenden den gesperrten Fallback `LoginFlow`, zeigen `Strings.ManagedAuthModeInvalid` und erlauben niemals den Login-Autostart. Gültige Sperren verwenden `Strings.PolicyAdminControlledTooltip`. Der effektive Modus bleibt vom ursprünglichen lokalen Modus getrennt: Vollständiges und reines Anmeldespeichern erhalten die lokale Auswahl, Klone übernehmen beide Zustände und nach Entfernen der Registry-Vorgabe gilt wieder die lokale Auswahl. Die administrative Anmeldeart wird nicht als neue Benutzerpräferenz persistiert.

Der Login-Autostart erweitert den bestehenden Anmeldeeinstieg für Freigabe/Talk, nicht das normale Öffnen der Einstellungen. Er verlangt unvollständige Zugangsdaten, die ausdrückliche gültige Registry-Vorgabe `AuthMode=LoginFlow` und eine gültige Registry-`NextcloudUrl`, die der tatsächlich verwendeten effektiven URL entspricht. Eine ausschließlich lokal gespeicherte URL, ein alleinstehender URL-Sperrwert, eine abweichende effektive URL, `Manual` oder ungültige Policy-Daten lösen ihn nicht aus. Der vorhandene Login Flow startet pro Anmeldeaufruf höchstens einmal automatisch; nach Abbruch oder Fehler bleibt der Dialog für einen ausdrücklich angeforderten neuen Versuch über den Login-Button verfügbar, ohne automatische Schleife. Vollständige Zugangsdaten lösen keine neue automatische Erstanmeldung aus. Erfolg verwendet den bestehenden Weg zum geprüften Speichern und Fortsetzen am ursprünglichen Outlook-Element; es entsteht weder eine zweite Anmeldeimplementierung noch eine doppelte Freigabe-/Talk-Ausführung. TLS-Validierung und Community-/Pro-Seat-Zugriffsprüfung bleiben unverändert.

Ein geprüfter Login schließt eine anmeldungspflichtige Share-/Talk-Einrichtung bei gültig verwaltetem `AuthMode=LoginFlow` automatisch ab, auch nach bewusstem Wiederholen. Nach Wiederherstellen von temporärem TLS und Aufheben des Busy-Zustands wartet `StartLoginFlowAsync` auf `SaveSettingsWithErrorHandlingAsync`, denselben Task wie der Speichern-Button. Vorhandene Serverstatus-Aktualisierung und Validierung bleiben in `SaveSettingsAsync`; Persistenz und Aktionsfortsetzung bleiben beim `SettingsWorkflowController`. Fehlgeschlagene Anmeldung speichert nicht; fehlgeschlagene Validierung lässt den Dialog offen. Ein späterer Persistenzfehler wird vom Controller gemeldet und setzt die Aktion nicht fort. Normale Einstellungen, unverwaltete Anmeldung, Manual und ungültiges AuthMode speichern oder schließen nicht automatisch.

`DefaultsSource` ist eine optionale `REG_SZ`-Policy (`local` oder `backend`) mit der bestehenden unabhängigen Hive-/View-Priorität. Ihre Existenz aktiviert Enterprise Rollout; der erste vorhandene Eintrag verdeckt nachrangige Einträge auch bei ungültigem Inhalt. Ein ungültiger ausgewählter Eintrag ergibt `local` und behält Verwaltungssperre sowie Konfigurationshinweis. Registry-Änderungen benötigen einen Outlook-Neustart. Gespeicherte Benutzerauswahl und deren Existenz bleiben getrennt von Laufzeitübersteuerungen; XML und Klone erhalten die lokale Auswahl, ohne die effektive administrative Quelle zu persistieren.

Die Backend-Statusantwort kann auf oberster Ebene `defaults_source` (`inherit`, `local`, `backend`) und `defaults_source_editable` (JSON-Boolean) enthalten. Fehlende Quellenmetadaten oder `inherit` ignorieren sowohl die Quellenübersteuerung als auch ihr Editierbarkeitsflag: Es gilt die Registry-Quelle, falls vorhanden, sonst die gespeicherte Benutzerauswahl, sonst `local`. Eine ausdrückliche Backend-Quelle ersetzt die Registry-Quelle einschließlich ihrer Sperre. Bei `defaults_source_editable=true` gewinnt eine ausdrücklich gespeicherte Benutzerauswahl gegenüber diesem Backend-Startwert; andernfalls ist die Backend-Quelle fest vorgegeben. Ältere Backends behalten damit das geerbte Verhalten, ältere Clients ignorieren die zusätzlichen Felder. Quellenauswahl und Backend-Standardwerte benötigen die bestehende Prüfung auf bestätigt gültigen persönlichen Zugriff und aktiv zugeteilten Seat, für Community und Pro identisch. Ein fehlender oder pausierter Seat schaltet keine Backend-Standardwerte frei; vorhandene Snapshot-/Cache- und Enterprise-Rollout-Zugriffsregeln bleiben unverändert.

`SettingsForm` platziert die Quellenauswahl unter Erweitert. Bei effektiver Backend-Quelle werden Freigaben, Talk und Signatur ausgegraut und können nicht ausgewählt werden; ihre lokalen Werte bleiben erhalten. Der Wechsel einer editierbaren Quelle zurück zu lokal aktiviert die Tabs wieder unter den bestehenden Funktions-/Feldbeschränkungen. Der Speicherpfad erfasst nur Benutzeränderungen, nicht die über eine Laufzeitüberlagerung angezeigten Backend-Startwerte. Freigabe- und Talk-Assistent erlauben weiterhin Änderungen für die einzelne Aktion, soweit ihre bestehenden einzelnen Policy-Regeln dies zulassen. Die Quellenauswahl bevorzugt Standardwerte und ist keine zusätzliche Feld-Policy-Sperre.

`UpdateNotifyEnabled` verwendet dieselbe ausschließlich zur Laufzeit wirksame Registry-Überlagerung wie TLS und Logging. Die Existenz aktiviert Enterprise Rollout; der erste vorhandene Wert nach Hive-/View-Priorität gewinnt. Ein ungültiger ausgewählter Wert ergibt false und hält ein Gültigkeitsflag für den Tooltip am gesperrten Schalter bereit. `SettingsStorage` persistiert nur `LocalUpdateNotifyEnabled`; der Einstellungsdialog kopiert keinen verwalteten effektiven Wert hinein. Klone behalten Policy und lokalen Zustand; Entfernen der Policy stellt die lokale Auswahl wieder her. `UpdateCheckService.ShouldNotify` liest den bestehenden effektiven Getter und behält die tägliche/versionsbezogene Duplikatunterdrückung; automatische und manuelle Update-Abrufe bleiben unverändert. Produktive Tests prüfen Benachrichtigungsentscheidung, echte Felder, Speichern, Registry-Priorität, XML sowie Community-/Pro-Gleichstellung.

- `Settings/SettingsStorage.cs` wählt eine profilspezifische XML-Datei unter `%LOCALAPPDATA%\NC4OL`, setzt Standardwerte für fehlende Einträge und schützt das App-Passwort mit Windows-DPAPI im `CurrentUser`-Kontext. `SettingsFileTransaction` schreibt und validiert eine temporäre Datei unter einem profilbezogenen prozessübergreifenden Mutex, ersetzt danach die Primärdatei und behält die vorherige gültige Datei als `.bak`.
- Ein fehlerhafter Passwortwert leert nur das Passwort und sperrt Settings-Schreibvorgänge im Hintergrund. Bei ungültiger Primär-XML wird die gültige Sicherung geladen. Ist keine gültige Datei vorhanden, bleiben automatische Schreibvorgänge gesperrt, bis das ausdrückliche Speichern im Einstellungsdialog erfolgreich war. Erst danach ändert sich die aktive Laufzeitkonfiguration.
- `SettingsWorkflowController` persistiert die Kandidatenkonfiguration vor Änderungen an Laufzeit-, TLS- oder IFB-Zustand. Ein Schreibfehler zeigt den lokalisierten Speicherfehler und meldet einen erfolglosen Speichervorgang, ohne die Laufzeiteinstellungen zu ändern.
- `Settings/ManagedSetupPolicy.cs` liest `Software\Policies\NC Connector` in dieser Reihenfolge: HKLM 64-Bit, HKLM 32-Bit, HKCU 64-Bit, HKCU 32-Bit; unter 32-Bit-Windows werden nur die 32-Bit-Ansichten gelesen. Eine nicht gesperrte URL füllt ein leeres Profil; eine gesperrte URL überschreibt den Profilwert. `TransportTlsUseSystemDefault`, `TransportTlsEnable12` und `TransportTlsEnable13` werden unabhängig aufgelöst und übernehmen jeweils den ersten vorhandenen Eintrag. Existenz und Gültigkeit bleiben getrennt: Ein ungültiger höherrangiger TLS-Eintrag verdeckt nachrangige Einträge und erzeugt einen Konfigurationsfehler statt eines Fallback-Werts.
- Jeder vorhandene TLS-Registry-Wert, einschließlich `false` oder eines ungültigen Werts, aktiviert Enterprise Rollout und sperrt die gesamte TLS-Gruppe. Fehlende Geschwisterwerte verwenden die Produktstandards `false / true / false`, niemals die lokalen Profileinstellungen. Der Systemstandardmodus ignoriert die beiden ausdrücklichen Versionsauswahlen. Ein ungültiger effektiver TLS-Wert oder die verwaltete Kombination `false / false / false` endet fail-closed; der Kompatibilitätsfallback für eine leere nicht verwaltete Auswahl darf bei einer verwalteten Policy TLS 1.2 nicht stillschweigend wieder aktivieren.
- `AddinSettings` hält die ursprünglichen lokalen TLS-Werte getrennt von einer ausschließlich zur Laufzeit wirksamen verwalteten Überlagerung. Die TLS-Property-Getter liefern bei aktiver Policy die effektiven verwalteten Werte; lokale Werte bleiben zum Persistieren erhalten und gelten wieder, sobald alle TLS-Policy-Werte entfernt sind. Klone übernehmen beide Zustände. `SettingsStorage` schreibt nur die ursprünglichen lokalen TLS-Werte; das Speichern anderer Einstellungen macht eine administrative Vorgabe deshalb nicht zur dauerhaften lokalen Auswahl. Die Einstellungs-UI darf gesperrte effektive Werte nicht in lokale TLS-Werte zurückschreiben; alle drei Felder bleiben auch nach Busy-State- oder Backend-Status-Aktualisierungen gesperrt.
- `Utilities/NextcloudUriValidator.cs` akzeptiert nur HTTPS-Nextcloud-Basis-URLs ohne Benutzerinformationen, Query oder Fragment. Login-Flow- und Passwort-Policy-URLs aus Serverantworten müssen Schema, Host und Port der konfigurierten Basis-URL beibehalten.
- `Services/BackendPolicyService.cs` liest den optionalen Backend-Status für Einstellungen, Talk, FileLink, verwaltete Signaturen und die Raumlöschung gespeicherter Termine. Freigabe und Talk können bei fehlendem Backend oder Seat lokale Werte verwenden. Das Send-Gate für verwaltete Signaturen nutzt den strengeren Policy-Zustand aus dem Signaturablauf.
- `TransportSecurityConfigurator` validiert TLS-Einstellungen zentral und wendet sie über `ServicePointManager.SecurityProtocol` an. Zuvor setzt er die .NET-Schalter für System-Default-TLS und starke Kryptografie programmatisch; die Auswahl hängt im von Outlook bereitgestellten AppDomain deshalb nicht allein von `NcTalkOutlookAddIn.dll.config` ab. `NcHttpClient` und der Update-Prüfpfad erzeugen Anfragen über `TransportSecurityConfigurator.CreateRequest`: Unter der gemeinsamen Add-in-Sperre lehnt diese Methode ungültiges verwaltetes TLS vor der Anfrageerzeugung ab oder setzt das verwaltete Protokoll unmittelbar vor dem Konstruktor erneut. Der .NET-Framework-Konstruktor übernimmt `ServicePointManager.SecurityProtocol`; `HttpWebRequest.SslProtocols` ist im Ziel .NET Framework 4.7.2 intern und keine öffentliche Einstellung pro Anfrage. Siehe den [Microsoft-Referenzquelltext](https://github.com/microsoft/referencesource/blob/main/System/net/System/Net/HttpWebRequest.cs#L4700-L4703).
- TLS-Übernahme, Anfrageerzeugung und temporäres Wiederherstellen verwenden dieselbe Sperre. `Restore` ersetzt kein aktives verwaltetes Protokoll. Das koordiniert NC Connectors eigene Einstellungsvorschau, Dialogabbrüche und Anfragen; fremder Code, der den AppDomain-weiten Standard ohne diese Sperre verändert, wird dadurch nicht synchronisiert. Die Policy ist eine Add-in-Rollout-Vorgabe, keine Isolation gegenüber anderen Add-ins und keine computerweite TLS-Härtung. Ungültiges verwaltetes TLS zeigt beim Start sowie im Einstellungs-/Login-Ablauf vor Anmelde- oder Seat-Prüfungen einen Konfigurationsfehler. Eine von der Laufzeit abgelehnte TLS-1.3-Auswahl behält das bestehende Verhalten ohne Fallback. Verbindungstest und Login-Flow-Diagnose fordern über `NcHttpClient` eine neue Verbindung an, damit ein geänderter TLS-Modus mit einem neuen Handshake statt über eine vorhandene gepoolte Verbindung geprüft wird. Registry-Änderungen erfordern einen Outlook-Neustart.
- `app.config` bildet die Versionen der gebündelten HtmlSanitizer-Abhängigkeiten ab. Der Laufzeit-Resolver bedient nur passende Anfragen des Add-ins und dieses Abhängigkeitsstapels aus dem Add-in-Verzeichnis; höhere Versionen, andere Tokens, Kulturen oder fremde Requester bleiben unberührt. `tools/ci/Check-VendorAssemblyBindings.ps1` lädt jede transitive Vendor-Referenz in einer frischen .NET-Framework-AppDomain mit diesen Redirects.

### Darstellung des Lizenzstatus

#### Enterprise Rollout

`ManagedSetupPolicy` erfasst die Existenz der Registry-Werte getrennt von ihren interpretierten Werten. Jeder vorhandene Wert `NextcloudUrl`, `NextcloudUrlLocked`, `AuthMode`, `ShowMainRibbonTab`, `DefaultsSource`, `TransportTlsUseSystemDefault`, `TransportTlsEnable12`, `TransportTlsEnable13`, `DebugLoggingEnabled`, `LogAnonymizationEnabled`, `UpdateNotifyEnabled`, `IfbEnabled`, `IfbDays`, `IfbCacheHours` oder `IfbPort` aktiviert den Rollout, auch `false`, leere oder ungültige Werte. Fehlen alle fünfzehn Werte, gilt das bisherige lokale Verhalten. `NextcloudUrl`, `AuthMode`, `ShowMainRibbonTab`, `DefaultsSource` und jeder TLS-, Logging-, Update-Benachrichtigungs- und IFB-Wert werden unabhängig nach Hive-/View-Priorität aufgelöst; die URL-Sperre gehört weiterhin zum ausgewählten URL-Eintrag, nicht zu einem alleinstehenden Sperrwert in einem anderen Hive. `AddinSettings` führt `IsEnterpriseRollout`, `ShowMainRibbonTab` und die verwalteten Anmeldeart-/Standardwertequellen-/TLS-/Logging-/Update-Benachrichtigungs-/IFB-Überlagerungen ausschließlich zur Laufzeit. Klone übernehmen diesen Zustand; XML schreibt diese administrativen Vorgaben nie. Rollout, Neustartanforderungen und das Entfernen aller fünfzehn Auslöser stehen in [ADMIN.de.md](ADMIN.de.md#enterprise-rollout).

Nur `ShowMainRibbonTab=false` blendet den Haupttab samt Einstellungsbutton aus und sperrt den öffentlichen Settings-Callback. `IsEnterpriseRollout` steuert die Zugriffsprüfung, nicht die Sichtbarkeit der Einstellungen. Bei sichtbarem Tab zeigt `SettingsForm` alle Einstellungen und speichert auch Einstellungen außerhalb der Anmeldung regulär; eine TLS-Policy sperrt dabei nur die vollständige TLS-Gruppe. Freigabe und Talk verwenden für Erstanmeldung oder erneute Anmeldung den internen Pfad `OpenAuthenticationSettingsAsync`. Bei ausgeblendetem Tab zeigt der Dialog nur Allgemein und speichert nur Änderungen der Zugangsdaten/Anmeldeart; der bestehende Workflow persistiert weiterhin vor der Laufzeitübernahme. Die verwaltete TLS-Prüfung gilt auch für diesen reinen Anmeldepfad und hängt weder von Zugangsdaten noch von einer verfügbaren Backend-/Seat-Antwort ab. Ein gültiger aktiver Community-Seat bleibt einem Pro-Seat gleichgestellt.

`OpenAuthenticationSettingsAsync` gibt zurück, ob erfolgreich gespeichert wurde. Der Pfad überspringt Backend-/Adressbuchabrufe vor der Anmeldung, startet `SettingsForm.BeginAuthentication` und verlangt geprüfte Zugangsdaten vor dem Speichern. Eine ausstehende Einrichtung zeigt eine Verbindungsaufforderung; typisierte Anmeldeablehnungen zeigen eine Aufforderung zur erneuten Anmeldung statt einer Backend- oder Seat-Diagnose. Geänderte Zugangsdaten müssen erneut geprüft werden. Transportfehler behalten ihren Diagnosepfad. Beide Startcontroller behalten das ursprüngliche Outlook-Element während der Anmeldung und setzen erst nach erfolgreicher Speicherung und Laufzeitübernahme fort. `MailInteropController.IsItemOpenForRibbonAction` prüft offene Inspectors anhand ihrer COM-Identität oder die aktive Inline-Antwort auf dem Outlook-STA; geschlossene Elemente werden nicht erneut geöffnet. Abbrechen oder ein geschlossenes ursprüngliches Element beendet die Aktion. Freigaben wiederholen bei erneuter Anmeldung höchstens einmal ihre rein lesenden Vorabprüfungen, keinen Assistenten und keine Mutation; die Anhangsautomatisierung startet keine interaktive Anmeldung.

`PolicyUiHelper.GetEnterpriseRolloutNotice` verbindet bestätigte Backend-Verfügbarkeit mit der vorhandenen persönlichen Seat-Prüfung, ohne Community-/Pro-Filter. Die Startcontroller prüfen vor Queue-Übernahme oder Raumerstellung. Anhangs- und Signaturverarbeitung deaktivieren ihre verwalteten Aktionen bei fehlendem Zugriff, ohne eine allgemeine Outlook-Versandsperre einzuführen. Talk-Hintergrundaktualisierung und IFB-Anfragen prüfen vor entfernten Operationen. Bereinigung und bereits angenommene separate Passwortzustellung behalten ihren bisherigen Lebenszyklus.

`FetchEnterpriseRolloutPolicyStatus` nutzt über `StoreBackendPolicySnapshot` den bestehenden kontogebundenen Policy-Snapshot: Fehlgeschlagene Aktualisierungen behalten eine passende bestätigte Antwort, frische erfolgreiche Ablehnungen ersetzen sie, und Fehler verschieben den Erfolgszeitpunkt nicht. Kein persistierter Rollout-Schalter, keine zusätzliche Lizenzfristberechnung und keine Vendoränderung. Tests gegen die produktive Assembly prüfen Registry-Priorität, Ribbon-Callbacks, reine Anmeldeoberfläche, Seat-Gleichstellung, Automatisierung und Cacheübergänge.

`AddinSettings.ResolvePolicyDefaults` ist die gemeinsame Share-/Talk-Auflösung für Einstellungen, beide Assistenten, Sprachen erzeugter Textblöcke, Anhangsautomatisierung und das Löschen gespeicherter Räume. Eine aktive gesperrte Backend-Vorgabe gewinnt immer. Bei lokaler Quelle steht ein ausdrücklicher lokaler Wert vor einem aktiven editierbaren Backend-Startwert; bei Backend-Quelle steht dieser Startwert vor dem lokalen Wert. Fehlende Werte verwenden weiterhin den Produktstandard. `EmailSignaturePolicyService` verwendet dieselbe Quellenpräferenz für die drei Einfügeschalter, nicht für die ausschließlich zentrale Vorlage. Das Einstellungsmodell merkt sich, ob eine lokale Auswahl existiert, nicht ob sie vom Standard abweicht; Klone kopieren diesen Zustand unabhängig. XML lässt unberührte Felder aus, vorhandene gespeicherte Werte werden als ausdrückliche Auswahl geladen. Die Einstellungen erfassen Änderungen, unterdrücken dabei aber programmatische Overlays; ein reines Speichern der Zugangsdaten darf angezeigte Backend-Startwerte nicht lokal festschreiben. Aufgelöste Laufzeitkopien werden nie persistiert.

`BackendPolicyService.ParseStatus` normalisiert nur ein ausdrückliches altes Backend-`share_expire_days=0` auf einen Tag, ohne das Eingabedictionary zu verändern. Fehlende/Null-Werte und lokale Ablaufentscheidungen bleiben getrennt. Anhangsschwelle `null` deaktiviert die Automatisierung, alte Null ergibt fünf MB, aktuelle positive Schwellen bleiben im Bereich 1–10240 MB. Der Immer-Schalter hat Vorrang vor dem Schwellenmodus. `Invoke-OutlookPolicyMappingTests.ps1` prüft die Persistenz, echte WinForms-Felder und den operativen Anhangsresolver; die Aktualisierung alter erfolgreicher Policy-Snapshots bleibt unverändert.

Eine Antwort ohne `status`-Objekt ist ein fehlgeschlagener Abruf (`invalid_payload`), kein bestätigter fehlender Seat. Sie verwendet den Hinweis auf den nicht abrufbaren Backendstatus und ersetzt keinen erfolgreich gespeicherten Status. Optionale Lizenzangaben bleiben für ältere Backends optional.

`BackendPolicyService.ParseStatus` normalisiert direkte und OCS-umschlossene Statusantworten. `BackendPolicyStatus` übernimmt die optionalen Angaben `license_status`, `access_status`, `can_manage_license`, `grace_until_iso`, `license_activation.state`, `license_connection_error`, `license_last_sync_at_iso` und `license_offline_until_iso`. Fehlende Felder bleiben leer/false; nur ein JSON-Boolean `true` schaltet den Lizenzverwaltungslink frei.

`PolicyUiHelper` wählt die lokalisierten Hinweise beim Anzeigen in Einstellungen, FileLink und Talk und verwendet dieselben Ursachen in Tooltips deaktivierter Funktionen. Lizenzablehnungen werden von `seat_state=suspended_overlimit` unterschieden; unbekannte Seat-Zustände erhalten einen allgemeinen Seat-Hinweis. Ältere Backends ohne erklärende Felder erhalten einen allgemeinen Zugriffshinweis. Nachfrist- und Verbindungshinweise verwenden einen gelben Statusstil, ohne die Funktionsfreigabe zu ändern. Vom Backend gelieferte Zeitpunkte werden nur zur Anzeige in Ortszeit formatiert und nicht als clientseitige Berechtigungsfristen ausgewertet.

Ein Administrator ohne Seat erhält im Banner die Lizenzdiagnose plus den Hinweis auf lokale Nutzung und fehlenden Seat. Der Tooltip an der gesperrten Funktion nennt dagegen den fehlenden persönlichen Seat. Nachfristtexte beschreiben die Verfügbarkeit für aktiv zugeteilte Seats, nicht für jeden Administrator. Hinweise zur abgelaufenen Offlinefrist enthalten bekannte Synchronisationszeitpunkte auch ohne aktuellen Verbindungsfehler.

Community und Pro verwenden dieselbe persönliche Zugriffsprüfung. Globales `overlicensed` sperrt aktive Seats und ihre Policy-Domains nicht; `suspended_overlimit` dagegen schon. `Invoke-OutlookPolicyMappingTests.ps1` prüft 384 paarweise Kombinationen aus Betriebsart, Zugriff, Seat, Kapazität, Rolle und Synchronisierung.

Der Verwaltungslink wird aus der validierten konfigurierten HTTPS-Basis-URL unter Beibehaltung des Installationsunterpfads gebildet und führt zu `/index.php/settings/admin/ncc_backend_4mc`. Er erscheint nur bei Lizenzhinweisen für Benutzer mit `can_manage_license=true`. Dies ist ein UI-Hinweis, keine Autorisierung; das Backend prüft die Berechtigungen weiterhin selbst. `WarningPanelUiHelper` berechnet die Panelhöhe mit oder ohne Aktionslink.

Die Freigabe bleibt von den bisherigen Prüfungen auf `is_valid`, `seat_assigned`, `seat_state` und Policy-Domains abhängig. Es gibt keine clientseitige Aktivierung, Lizenzserveranfrage, neuen Hintergrundtimer, COM-Thread-Änderung oder Änderung des Signaturcaches. `Invoke-OutlookPolicyMappingTests.ps1` prüft Parser-Kompatibilität, Status-/Seat-Priorität, rollenabhängige Aktionen und unveränderte Freigabeentscheidungen.

### Ende-zu-Ende-Abläufe

#### Talk-Link-Ablauf (Termine)

1. Der Benutzer klickt in einem Termin auf **Talk-Link einfügen**.
2. `TalkLinkForm` erfasst Titel, Passwort, Lobby, Sichtbarkeit, Raumtyp, Teilnehmersynchronisierung und optionales Ziel der Moderatorübergabe.
3. `TalkRibbonController` lädt Backend- und Passwort-Policy parallel, bevor der Wizard geöffnet wird. Ein Delegationsziel wird als eigener Benutzer abgelehnt, wenn es zur kanonischen UID, zum konfigurierten Login oder zur bekannten primären E-Mail-Adresse passt.
4. Vor der Serveranfrage sichert das Add-in Betreff, Ort, Body und alle `X-NCTALK-*`-Eigenschaften des Termins. `TalkService` erstellt anschließend den neuen Raum, während ein vorhandener Raum weiterhin verfügbar bleibt.
5. `TalkAppointmentController.ApplyRoomToAppointment(...)` schreibt URL, lokalisierten Body-Block und `X-NCTALK-*`-Metadaten in den Termin.
6. Erst wenn alle Terminänderungen erfolgreich waren, wird die neue Appointment-Subscription registriert und ein vorhandener Raum entfernt. Scheitert eine Terminänderung, stellt das Add-in den gesicherten Zustand wieder her und löscht ausschließlich den neu erstellten Raum; eine fehlgeschlagene Bereinigung wird in die unbedingte persistente Löschqueue eingestellt. Scheitert nach erfolgreicher Ersetzung nur die Löschung des alten Raums, bleibt der neue Raum verknüpft und die alte Raumlöschung wird zur Wiederholung vorgemerkt.
7. Beim Speichern liest die Appointment-Subscription die benötigten Outlook-Werte auf dem STA-Thread. `TalkAppointmentSyncCoordinator` fasst unveränderliche Snapshots zusammen und führt Lobby-, Beschreibungs-, Teilnehmer- und Delegationsaufrufe im Hintergrund aus.
8. `BeforeDelete` verwendet Outlooks terminspezifisches Löschereignis. Organizer-, Token-, Delegations- und Serienprüfung laufen an diesem Termin, bevor `QueueSavedTalkRoomDeletion(...)` einen Löschauftrag mit `PolicyRequired=true` erstellt; URL- oder Ortsauswertung ist keine Löschquelle. Der Hintergrund-Worker wertet die wirksame `TalkDeleteRoomOnEventDelete`-Policy vor der Raumlöschung aus. Derselbe Ereignispfad gilt für die Löschung aus dem geöffneten Termin und aus der Kalenderansicht. `Explorer.SelectionChange` bindet nur ausgewählte Talk-Termine; beim Hook jedes Explorers wird dessen aktuelle Auswahl einmal verarbeitet, damit die Kalenderlöschung direkt nach einem Outlook-Neustart ohne vorheriges Öffnen funktioniert.
9. Beim Start werden nur der persistente Lösch-Retry-Worker und bestehende Explorer-Oberflächen initialisiert. Stores oder Kalenderordner werden nicht aufgezählt, Kalenderelemente nicht durchsucht und keine ordnerbezogenen `Items`-Subscriptions gehalten.
10. Die DPAPI-geschützte Löschqueue besitzt Primärdatei und Sicherung. Die Nextcloud-Löschung läuft im Hintergrund; fehlgeschlagene Löschungen werden verzögert und nach einem Outlook-Neustart wiederholt. Die Bereinigung eines neu erstellten Raums aus einem ungespeicherten und verworfenen Termin verwendet dieselbe Queue mit `PolicyRequired=false`. Beim Laden älterer Zustände bleiben nur bereits zur Löschung vorgemerkte Einträge erhalten; reine Tracking-Einträge werden verworfen.

#### HTML-Subset für Talk-Termine (Backend-Vorlagen)

Backend-Talk-Templates verwenden für die Outlook-Word-/RTF-Pipeline bevorzugt Tabellen (`table`, `tbody`, `tr`, `td`). Der Kompatibilitätsschritt entfernt `display:flex|grid`, `flex*`, `grid*`, `border-radius*`, `overflow*`, `object-fit` und `user-select`, ergänzt Legacy-Attribute für Farbe und Ausrichtung und entfernt nicht erlaubte Tags oder Attribute.

#### Freigabeablauf (E-Mail verfassen)

- Der FileLink-Ribbon-Einstieg ist im Mail-Inspector und im Explorer-Tab `Nachricht` fuer Inline-Antworten/-Weiterleitungen sichtbar. Beide Einstiege laufen ueber denselben `FileLinkLaunchController`.
- Der Wizard startet mit den gespeicherten lokalen FileLink-Vorgaben. Ein gesperrter Share-Policy-Wert überschreibt den zugehörigen lokalen Wert; ein editierbarer Wert lässt die gespeicherte Outlook-Einstellung unverändert.
- Der Datei-Schritt gruppiert lokale und **Meine Nextcloud**-Auswahlen in derselben Warteschlange. Der sichtbare Ordnerbaum entsteht aus unveränderlichen Auswahl-Snapshots. Interaktive lokale Scans laufen als abbrechbare Dateisystemarbeit im Hintergrund; erst der fertige Snapshot wird wieder im erfassten WinForms-Kontext in die Queue übernommen. Der Uploadplan verwendet genau dieselben Snapshots; später hinzugekommene Dateien werden nicht unsichtbar mit hochgeladen, und eine entfernte oder veränderte Datei aus der Queue stoppt den Upload als geänderte Quelle.
- Inline-Antworten/-Weiterleitungen fuegen das gerenderte Freigabe-HTML ueber `Explorer.ActiveInlineResponseWordEditor` ein; der Inline-Pfad schreibt nicht direkt in `MailItem.HTMLBody` und behaelt zwei leere Absaetze ueber dem Freigabeblock fuer eigenen Text.
- Normale HTML-Compose-Fenster verwenden zuerst den Inspector-WordEditor, damit verwaltete Bookmarks erhalten bleiben. Nur wenn dieser Editor nicht geoeffnet werden kann, bleibt die direkte `MailItem.HTMLBody`-Route als Kompatibilitaetsfallback aktiv.
- Offene Verfassenfenster verwerfen ihren Snapshot der Anhangsregeln unmittelbar nach dem Speichern lokaler Einstellungen. Nach fünf Minuten aktualisieren Anhangsereignisse die Backend-Policy im Hintergrund und verwenden währenddessen die zuletzt ermittelten Regeln weiter. Auch der Versand verwendet diesen Stand ohne Aktualisierungshinweis; eine fehlgeschlagene Abfrage behält die bisherigen Regeln bei und setzt ihr Alter nicht zurück. Freigabe- und Talk-Assistenten behalten die beim Öffnen geladene Policy für ihren gesamten Vorgang.
- Nur ein fehlender erster Snapshot oder eine Einstellungsänderung lässt den Versand auf die erste Regelabfrage warten. Erreicht diese erste Abfrage das Backend nicht, bleibt der bisherige Rückgriff auf lokale Einstellungen bestehen. Ein vor ihrem Abschluss gestarteter Sendeversuch wird mit einem Informationshinweis zum erneuten Senden abgebrochen; weder ein Warten auf dem STA-Thread noch ein automatischer Neuversand kommt hinzu. Eine wirksame `AlwaysConnector`-Regel blockiert gewöhnliche Anhänge weiterhin mit dem Hinweis auf verpflichtende Freigaben. Keiner dieser Hinweise behauptet einen fehlgeschlagenen Upload.
- Die Anhangs-Partials teilen sich eine Compose-Subscription mit ihren bestehenden Ereignisregistrierungen und ihrer Lebensdauer. Policy, Dateierfassung und Queue-Übergabe führen keine unabhängigen Ereignishandler oder Task-Aufrufpfade ein.
- `MailComposeSubscription` debounct Anhangsänderungen und verarbeitet Always-via-NC sowie den Schwellwertmodus. `BeforeAttachmentAdd` versucht die Dateidaten früh zu erfassen; bei einer erzwingenden Policy wird ein nicht materialisierbarer oder nicht prüfbarer Host-Anhang abgebrochen. Bereits hinzugefügte Outlook-Anhänge werden erst entfernt, nachdem die vollständige Startauswahl in der Queue liegt. Ein späterer Abbruch des Wizards stellt übernommene Anhänge nicht wieder her. Harte Outlook-/Exchange-Grenzen können weiterhin vor einem Add-in-Ereignis greifen.
- Bei einer Mehrfachauswahl zeigt der Schwellwertdialog den Namen und die Größe derselben zuletzt hinzugefügten Datei. Die Entfernen-Aktion umfasst weiterhin den vollständigen zuletzt hinzugefügten Batch.
- Die Erfassung der aktuellen Outlook-Anhänge erfolgt erst nach dem Server-Prefetch auf dem Outlook-STA. Dateierfassung, Queue-Annahme und Entfernung der Originale bleiben in einem synchronen UI-Aufruf; Anhangspositionen werden nicht über die Wartezeit der Netzwerkanfragen hinweg gespeichert.
- Outlook-Body-Ressourcen mit `PR_ATTACHMENT_HIDDEN=true`, beispielsweise Signaturbilder, werden weder in Anhangs-Batches und Schwellwertsummen noch in FileLink-Auswahl, Host-Entfernung oder Send-Gate einbezogen.
- `NextcloudTalkAddIn.TryInsertHtmlIntoMail(...)` und `TryInsertPlainTextIntoMail(...)` geben den Einfügestatus von `MailBodyInsertionController` zurück. Scheitern alle Einfügepfade, stellt `FileLinkLaunchController` die neu erzeugten Serverartefakte zur Bereinigung ein und meldet den Wizard als fehlgeschlagen.
- `ComposeLifecycleOrigin` hält den exakten Server-/Kontokontext für das Löschen der erstellten Freigabe oder eine spätere Secrets-Anfrage. Die Bereinigung fällt nie auf das aktuell ausgewählte Konto zurück.
- Kann eine neu erstellte Freigabe nicht in die Mail eingefügt werden, versucht der Controller, ihren Serverordner mit diesem erfassten Kontext zu löschen.
- Nach erfolgreicher Einfügung verfolgt `MailComposeSubscription` den `ComposeShareCleanupRecord`, bis Outlook `AfterWrite` auslöst. Ein abgeschlossener Schreibvorgang umfasst Speichern, automatisches Speichern und den Schreibvorgang für Versand/Postausgang; diese Pfade geben den Bereinigungseintrag frei, ohne die Freigabe zu löschen.
- Klassische Compose-Fenster binden das konkrete `InspectorEvents_10.Close`-Ereignis. Es wird erst ausgelöst, wenn dieser Inspector tatsächlich schließt. Folgte auf die Einfügung kein erfolgreicher Schreibvorgang, stößt die Subscription die Löschung mit dem erfassten Konto und relativen Pfad an. Ein abgebrochener Schließvorgang lässt den Bereinigungsstatus daher aktiv.
- Inline-Compose behandelt `Explorer.InlineResponseClose` nur als Oberflächenwechsel, weil Outlook das Ereignis auch bei Pop-out und Navigation auslöst. `ItemEvents_10.Unload` bleibt dort das abschließende Item-Signal und wertet ausschließlich den vorher erfassten Bereinigungsstatus aus, ohne auf das entladene `MailItem` zuzugreifen.
- DAV-Bereinigungen nutzen den gemeinsamen begrenzten FileLink-Retry-Pfad. Ein wiederholtes Löschen bleibt idempotent, weil ein bereits fehlender Ordner akzeptiert wird.
- Die Bereinigungsverfolgung liegt im Arbeitsspeicher. Hat Outlook die Nachricht geschrieben, löscht das spätere Löschen dieses gespeicherten Entwurfs – auch nach einem Outlook-Neustart – die Freigabe nicht; die ungenutzte Freigabe muss manuell entfernt werden.
- `RegisterSeparatePasswordDispatch` hält die Daten für den Passwort-Follow-up zunächst nur in `_passwordDispatchQueue` der Subscription. Speichern und AutoSave schreiben diese Queue nicht in die Hauptmail. Beim Schließen des Verfassen-Fensters wird die Subscription verworfen; ein erneut geöffneter Entwurf, ein Outlook-Neustart vor dem ersten Sendeversuch oder eine neue Nachricht aus einer `.oft`-Vorlage kann den Follow-up-Zustand deshalb nicht wiederherstellen.
- `OnSend` ist die direkte Übergabegrenze. Die Queue wird einmal verbraucht, der endgültige Secrets-/Klartextinhalt erzeugt, Empfänger und wirksames `SendUsingAccount`/`SentOnBehalfOfName` der Hauptmail übernommen, Body und passende Backend-Signatur fertiggestellt, Empfänger aufgelöst und der Follow-up ohne zwischengespeicherten Outlook-Entwurf übergeben.
- Für die Passwortzustellung werden weder Entwürfe, Postausgang, Gesendet-Ordner, MIME-Marker noch eine Neustart-Wiederherstellung verwendet. Auch Outlooks verzögerte oder Offline-Übermittlung ruft `OnSend` auf; der Passwort-Follow-up wird deshalb sofort übergeben und wartet nicht auf die Hauptmail im Postausgang. Unerwartete Follow-up-Fehler setzen das `cancel`-Flag der Hauptmail nie.
- Das Send-Gate bricht den Versand bei einer erzwingenden Anhangs-Policy ab, solange noch ein gewöhnlicher regelwidriger Anhang vorhanden ist.
- `ComposeShareCleanupService` führt die Serverbereinigung im bestehenden Hintergrundtask aus. `SeparatePasswordDeliveryController` übernimmt Aufbereitung und direkte Übergabe der Passwortmail auf dem Outlook-STA im Send-Ereignis der Hauptmail. SMTP-Adressen werden gemeinsam über An, Cc und Bcc dedupliziert. Im Secrets-Modus entsteht ein Einmal-Link pro eindeutiger Adresse. Bei eindeutigem Auto-Send-Fehler wird genau ein vollständig vorbereiteter manueller Fallback geöffnet; bei mehrdeutigem Outlook-Status wird kein Duplikat erzeugt. Scheitert die strikte Absender- oder Empfängeraufbereitung vor der Übergabe, verwendet der Controller denselben Fallback mit normalisierten An-/Cc-/Bcc-Feldern und gleicht die verwaltete Signatur nach Initialisierung des Inspectors ab.
- Secrets-Links werden lokal per AES-GCM über Windows CNG verschlüsselt. Schlägt die Secrets-Erstellung fehl, wird die Klartext-Passwortmail verwendet und ein Hinweis angezeigt.
- `OutlookAttachmentAutomationGuardService` erzwingt den Host-Konflikt-Guard live:
  - vor Auswertung
  - vor Prompt-Aktionsverarbeitung
  - vor Wizard-Finalize im Attachment-Modus.
- `Models/AttachmentLinkTargetPolicy.cs` löst `policy.share.attachment_link_target` (`zip_download` / `share_page`) anhand der effektiven Standardwertequelle gegen den nullable lokalen Wert auf. Ein ungültiger gespeicherter lokaler Wert gilt als nicht gesetzt. ZIP gilt nur ohne gültigen lokalen oder nutzbaren Backend-Wert; ein gesperrter Backend-Wert gewinnt immer.
- `AttachmentMode` steuert Read-only-Berechtigungen, das Ausblenden der Rechtezeile und Cleanup. Das explizite Linkziel steuert nur URL, `{LINK_INTRO}` und `{LINK_LABEL}`; manuelle Freigaben bleiben immer auf der Nextcloud-Freigabeseite. Im Wizard gibt es keinen Schalter pro Freigabe.
- Die ZIP-URL-Ableitung ist fail-closed: Die absolute oeffentliche HTTP(S)-URL muss auf `/s/<token>` enden und zum OCS-Token passen. Ungueltige Eingaben brechen vor dem Einfuegen ab; es gibt keinen Fallback auf die Original-URL.
- Custom-Share-Templates aus dem Backend werden im `FileLinkHtmlBuilder` vor der Einfuegung ueber `HtmlTemplateSanitizer` bereinigt (fail-closed).
- Vor der Platzhalterersetzung entfernt der Renderer bei leeren optionalen Werten den naechsten umgebenden Block. Feste Beschriftungen wie `Passwort` bleiben deshalb weder in HTML- noch in Plain-Text-Ausgaben ohne Wert stehen.
- `{LINK_INTRO}` und `{LINK_LABEL}` werden anhand des effektiven Linkziels aufgeloest. Bestehende Templates ohne diese Platzhalter behalten ihre bisherige Ausgabe.
- Fuer Custom-Share-Templates bevorzugt Outlook `policy.share.share_html_block_template_v2` und faellt auf `policy.share.share_html_block_template` zurueck. Damit funktionieren aeltere Backend-Versionen weiter, waehrend aktuelle Backends den bisherigen Antwortschluessel fuer aeltere Clients platzhalterfrei halten koennen.
- Aktuelle Backends liefern fuer Custom-Templates `policy.share.share_html_block_effective_language`. Outlook verwendet diese Sprache fuer erzeugte Linktexte, Feldbezeichnungen, Berechtigungsnamen und Passworthinweise; bei aelteren Backends ohne dieses Feld bleibt der bisherige Fallback auf die UI-Sprache erhalten.
- Plain-Text-Compose bleibt `MailItem.BodyFormat=olFormatPlain`; der Freigabeblock wird als Textblock mit `#`-Rahmen gerendert und ueber Outlook WordEditor eingefuegt. Inline-Antworten/-Weiterleitungen behalten zwei leere Absaetze ueber dem Block fuer eigenen Text. `MailItem.Body` wird nicht neu geschrieben.
- `FileLinkWizardForm` akzeptiert im Datei-Schritt Explorer-Drag-and-drop fuer Dateien/Ordner ueber Queue und Aktionsbereich.
- `FileLinkTransferService` nutzt fuer Dateien bis 20 MiB einen direkten WebDAV-`PUT`. Groessere Dateien laufen ueber Nextcloud Chunked Upload v2 unter `/remote.php/dav/uploads/<user>/<upload-id>` und werden danach per `MOVE .file` an den finalen DAV-Pfad zusammengesetzt.

##### Upload-Architektur

- Alle Funktionen setzen Nextcloud 32 oder neuer voraus. `NextcloudCapabilitiesService` validiert die strukturierte Version der authentifizierten OCS-Capabilities-Antwort und speichert den typisierten Snapshot fünf Minuten pro Server/Benutzer zwischen. Verbindungsprüfungen aktualisieren ihn; Funktionseinstiege lehnen ältere Server oder Antworten ohne auswertbare Version ab.
- `FileLinkService` orchestriert die fachlich getrennten Komponenten für Planung, DAV-Verzeichnisse, Transfer, Share-Erstellung und Fortschritt.
- Die Queue-Aufnahme verwendet `FileLinkSelection.IdentityComparer`: Lokale Windows-Pfade ignorieren Groß-/Kleinschreibung, Nextcloud-Pfade unterscheiden sie. Dieselbe Quelle wird nicht doppelt aufgenommen; unterschiedliche Nextcloud-Namen wie `report.pdf` und `Report.pdf` bleiben beide erhalten. Namenskonflikte am Ziel verwenden unabhängig davon weiterhin die bestehende Umbenennung.
- `FileLinkQueueSnapshotBuilder` scannt jede lokale Auswahl genau einmal beim Einfügen in die Queue. Auswahlen aus Dialogen und Drag-and-drop führen den rekursiven Scan mit Abbruchmöglichkeit außerhalb des UI-Threads aus; nur die bereits von Outlook materialisierten einzelnen Anhangsdateien verwenden die synchrone Startübergabe. Der relativ zum Freigabe-Stammordner aufgebaute Snapshot bewahrt leere Verzeichnisse, lehnt symbolische Links und Junctions ab und speichert Dateigröße sowie Änderungszeit. `FileLinkSelectionScanner` verwendet exakt diesen Snapshot; `FileLinkUploadPlanner` weist anschließend die Transferarten zu, ohne den ausgewählten Ordner erneut zu enumerieren oder den Server zu verändern.
- `NextcloudFilePickerForm` liest den Dateibereich des konfigurierten Benutzers mit DAV-`PROPFIND` der Tiefe eins. Die Adressleiste führt einen lokalen Zurück- und Vorwärtsverlauf, navigiert ohne wiederholten Kontonamen zu Elternordnern oder angeklickten Pfadsegmenten und aktualisiert den aktuellen Ordner ohne neuen Verlaufseintrag. Für den Stamm wird dasselbe Nextcloud-Symbol wie in Thunderbird verwendet. Für eine ausgewählte Datei fordert der Picker zuerst über Nextclouds authentifizierte Route `/index.php/core/preview.png` eine Vorschau mit 1024 × 1024 Pixeln an. Die Antwort ist auf 5 MiB begrenzt, wird außerhalb des UI-Threads dekodiert und nur innerhalb der Grenzwerte für dekodierte Bilder akzeptiert. Meldet Nextcloud, dass keine Vorschau verfügbar ist, wird nur bei unterstützten Rasterbildern mit höchstens 5 MiB Größe das Original über einen bytebegrenzten DAV-`GET` geladen; Dokument-Originale werden für Vorschauen nie heruntergeladen. Überholte Vorschauanfragen werden abgebrochen. Beim Bestätigen eines Ordners wird der vollständige Nachfahren-Snapshot einschließlich leerer Ordner festgehalten. Ausgewählte Dateien plant der Scanner als serverseitige Kopien. Der Transfer prüft ihre aktuelle Größe und kopiert sie mit authentifiziertem DAV-`COPY` in den reservierten Freigabeordner. Das Original bleibt unverändert; für die eigentliche Übertragung läuft kein Dateiinhalt durch Outlook.
- Beim Weiter aus dem ersten Schritt prüft der manuelle Wizard den aus Basispfad, festgehaltenem Wizard-Datum und bereinigtem Freigabenamen abgeleiteten Zielpfad mit einem DAV-`PROPFIND` der Tiefe null. Ein vorhandenes Ziel hält den Wizard im ersten Schritt. `FileLinkDavClient` reserviert den Freigabe-Stammordner beim späteren Upload atomar mit `MKCOL`, damit eine Kollision zwischen Vorprüfung und Upload sicher abbricht. Ein `405` nach einem unklaren ersten Ergebnis gilt nur dann als erfolgreiche Reservierung, wenn ein DAV-`PROPFIND` der Tiefe null den exakten Pfad als Collection bestätigt. Ein bekannter `405` ohne vorherige Unklarheit bleibt eine Kollision. Die Anhangsautomatisierung überspringt die Vorprüfung und probiert weiterhin nummerierte Freigabenamen. Leere Verzeichnisse, für Bulk oder Chunked benötigte Elternpfade und von mehreren Direct-Dateien gemeinsam genutzte Eltern werden einmal, Eltern vor Kindern, mit maximal drei parallelen Requests pro Ebene angelegt. Direct-Pfadketten mit nur einer Datei legt `X-NC-WebDAV-Auto-Mkcol` an.
- `FileLinkTransferService` koordiniert getrennte Bulk-, Direct- und Chunked-Uploader. Dateien außerhalb von Bulk bis 20 MiB werden mit dem serverseitig ausgewerteten Header `X-NC-WebDAV-Auto-Mkcol: 1` über direkte `PUT`-Requests hochgeladen. Dateien über 20 MiB verwenden Chunked Upload v2. Direct- und Chunked-Dateien teilen sich das Limit von maximal drei parallelen Transfers.
- Nur wenn `ocs.data.capabilities.dav.bulkupload` exakt `"1.0"` meldet, kommen mindestens 20 Kandidaten mit höchstens 8 MiB pro Datei für DAV-Bulk infrage. Sequentielle Multipart-Batches sind auf 100 Dateien und ungefähr 20 MiB begrenzt. Der Planner wählt Bulk nur, wenn mindestens 20 Prozent aller Upload-Requests entfallen. Die Berechnung umfasst Basispfad, Freigabe-Stammordner, geplante Verzeichnisse, direkte Dateien sowie jeden Chunk-Ordner, Chunk-`PUT` und abschließenden `MOVE`. Vor der ersten Serveränderung meldet die sequenzielle MD5-Berechnung ihren aktuellen und gesamten Dateizähler als eigene Wizard-Phase.
- Nach Abschluss aller Transfers sendet `FileLinkShareClient` genau einen OCS-Create-Share-`POST` mit Pfad, expliziten Berechtigungen, Passwort, Ablaufdatum, Label und Notiz. Der veraltete Parameter `publicUpload` entfällt, weil Nextcloud damit die explizite Berechtigungsmaske ersetzen würde. Ein nachträglicher Metadaten-`PUT` findet nicht statt.
- Bei ausbleibender Antwort, einer temporären Gateway-/Service-Antwort ohne OCS-Ergebnis oder einer erfolgreichen Antwort ohne verwertbare Freigabedaten ist das Ergebnis des Erstellaufrufs unklar. `FileLinkShareClient` merkt sich den Pfad und führt vor einem weiteren Erstellaufruf eine OCS-Abfrage für exakt diesen Pfad ohne untergeordnete Freigaben aus. Eine passende öffentliche Freigabe wird übernommen; ein neuer Versuch erfolgt nur nach einem bestätigten leeren Ergebnis. Solange die Abfrage selbst unklar bleibt, wird kein zweiter Erstellaufruf gesendet.
- Wiederholbare `MKCOL`-, direkte `PUT`-, Chunk-`PUT`- und Bulk-`POST`-Operationen erhalten bei Transportfehlern und ausgewählten temporären HTTP-Antworten maximal zwei Wiederholungen. Jeder Bulk-Versuch baut denselben Request-Body aus dem unveränderten lokalen Plan neu auf. Ein abschließendes Chunk-`MOVE` wird nie blind ein zweites Mal gesendet: Nach einem unklaren Transportergebnis wird das exakte Ziel mit einem DAV-Depth-0-Request geprüft und nur als erfolgreich gewertet, wenn es keine Collection ist und die erwartete Länge besitzt.
- Der Wizard zeigt Scan, Prüfsummenberechnung, Ordnervorbereitung sowie aggregierte Dateien, Bytes und Transferrate. Zwischenstände der Prüfsummen- und Transferphasen sind auf maximal zehn Meldungen pro Sekunde begrenzt; das Debug-Log schreibt Plan, Wiederholungen, aggregierten Fortschritt im Fünf-Sekunden-Takt und Abschluss.

#### IFB-Ablauf

`ManagedSetupPolicy` löst `IfbEnabled`, `IfbDays`, `IfbCacheHours` und `IfbPort` unabhängig nach der bestehenden Priorität HKLM 64-Bit, HKLM 32-Bit, HKCU 64-Bit, HKCU 32-Bit auf. Existenz und Gültigkeit bleiben getrennt: Ein ausgewählter ungültiger Wert verdeckt nachrangige Einträge. Jeder vorhandene IFB-Wert aktiviert Enterprise Rollout und sperrt alle vier IFB-Einstellungen, auch bei false oder ungültigem Inhalt. Fehlende Geschwisterwerte verwenden die Produktstandards `false / 30 / 24 / 7777`, nicht lokale Werte. `IfbEnabled` verwendet den bestehenden Boolean-Parser; die drei numerischen Policies verlangen Registry-DWORD-Ganzzahlen. Zulässige Tage, Cache-Stunden, Ports und den Rollout beschreibt [ADMIN.de.md](ADMIN.de.md#verwaltetes-internet-freebusy-ifb).

`AddinSettings.HasManagedIfb` und `IsManagedIfbValid` beschreiben die ausschließlich zur Laufzeit wirksame Überlagerung. Bestehende IFB-Getter liefern effektive Werte; separate ursprüngliche lokale Werte bleiben die einzigen IFB-Auswahlen, die `SettingsStorage` serialisiert. Klone behalten lokale Werte, Policy und Gültigkeit. Nach Entfernen aller IFB-Policy-Einträge gelten wieder die lokalen Einstellungen. `SettingsForm` zeigt die Überlagerung an, ohne sie beim vollständigen oder reinen Anmeldespeichern zurückzuschreiben; Initialisierung, Busy-State-Wechsel und Backend-Status-Aktualisierungen erhalten alle vier Sperren. Gültige Sperren verwenden `Strings.PolicyAdminControlledTooltip`; ungültige Policies verwenden `Strings.ManagedIfbPolicyInvalid` (`managed_ifb_policy_invalid`) in jeder unterstützten Sprache.

Ein ausgewählter ungültiger IFB-Wert macht die Gruppe ungültig, beendet IFB über den bestehenden Manager-Pfad, stellt eigene Outlook-Registry-Einträge wieder her und verhindert Listener-Start sowie entfernte IFB-Aufrufe. Der lokalisierte Fehler bleibt auf IFB begrenzt; Anmeldung, Verbindungstests, Freigaben, Talk und Updates werden dadurch nicht gesperrt. `IfbCacheHours` versorgt auch den gemeinsamen Talk-Adressbuch-Cache: Ein ungültiger Cache-Wert ergibt 24 Stunden; ein gültiger Wert bleibt wirksam, wenn ein anderes IFB-Feld ungültig ist. Gültiges deaktiviertes IFB verwendet den bisherigen Stop-/Restore-Pfad. Gültiges aktiviertes IFB behält den von Zugangsdaten abhängigen Start und die bestehende Enterprise-Rollout-Zugriffsprüfung vor jeder entfernten Anfrage, mit identischen Community-/Pro-Seat-Voraussetzungen und ohne neuen Zugriffsmodus. Registry-Policy-Änderungen benötigen einen Outlook-Neustart; ein eigener Port benötigt eine administrativ erstellte URLACL, ohne Rechteerhöhung oder automatische Reservierung durch das Add-in.

1. Die effektive lokale oder gültige verwaltete Konfiguration aktiviert IFB.
2. `FreeBusyManager` erzeugt für jeden Outlook-Prozess ein zufälliges Request-Secret. `IfbRegistryStateStore` hält den DPAPI-geschützten Besitzstand mit Primär-/Backup-Wiederherstellung.
3. `FreeBusyServer` startet den Loopback-Listener, akzeptiert nur `/nc-ifb/<request-secret>/freebusy/<address>.vfb` und begrenzt gleichzeitige Proxy-Anfragen auf vier.
4. `IfbRegistryOwnershipManager` merkt sich jeden ursprünglichen Benutzerwert und registriert `%NAME%@%SERVER%.vfb` unterhalb des geheimen Endpunkts. Outlook ersetzt beide Platzhalter durch die vollständige SMTP-Adresse des Teilnehmers. Policy-Werte werden nur auf Konflikte geprüft.
5. Die MSI entfernt eigene alte Pfade bei Installation, direktem Upgrade, Reparatur und vollständiger Deinstallation. Zur Laufzeit können alte Pfade ohne Token weiterhin übernommen werden; nicht zugeordnete tokenisierte Pfade bleiben gesperrt, damit kein anderes laufendes Profil übernommen wird.
6. Bei erhaltenem Zuordnungsjournal kann ein nach MSI-Wartung fehlender Registry-Wert sofort neu eingetragen werden. Beim Deaktivieren wird der Vorgänger wiederhergestellt, wenn der aktuelle Wert fehlt oder noch unserem geschriebenen Wert entspricht. Ein abweichender vorhandener Wert wird nie überschrieben.
7. `IfbAddressBookCache` trennt Einträge nach Outlook-Profil, normalisierter Nextcloud-Basis-URL einschließlich Unterpfad und kanonischer Nextcloud-UID.

`installer/IfbCleanup` ist ein eigenständiges .NET-Framework-Programm, eingebettet in die Binary-Tabelle der MSI. `build.ps1` baut es vor WiX. Es verwendet `IfbRegistryEndpoint` gemeinsam mit dem Add-in und entfernt ausschließlich die beiden bekannten Outlook-Suchwerte, wenn sie vollständig dem eigenen Loopback-URL-Format entsprechen. Policies und fremde Werte bleiben ausgeschlossen. Der Helfer ermittelt lokale Benutzerprofile, verwendet geladene HKU-Hives oder bindet bestehende NTUSER.DAT-Dateien kurzzeitig ein und läuft erhöht ohne Impersonation. Outlook muss sitzungsübergreifend geschlossen sein. Benutzereinstellungen und DPAPI-Zuordnungsdateien werden weder gelesen noch gelöscht oder entschlüsselt.

Die Rollback-Aktion steht vor der verzögerten Bereinigung; entfernte Rohwerte und Registry-Typen werden zuvor in einem nur für Administratoren/SYSTEM zugänglichen HKLM-Transaktionsjournal gesichert. Rollback stellt ausschließlich fehlende Werte in noch bestehenden Schlüsseln wieder her; Commit entfernt das Journal. Ein weiterer Wartungslauf arbeitet ein verbliebenes Journal vor der Bereinigung ab. Das ausgehende Upgrade-Paket überspringt die Bereinigung bei `UPGRADINGPRODUCTCODE`; das neue Paket führt sie aus. Windows-Installer-Rollback muss aktiviert bleiben. Setup-Meldungen werden aus `installer/IfbCleanup/SetupMessages.xml` in allen unterstützten Sprachen eingebettet.

`Invoke-OutlookIfbInstallerTests.ps1` prüft Bereinigung und Rollback mit isolierten HKCU-Testschlüsseln, nie mit echten Benutzer-Hives. `Check-MsiPackage.ps1` prüft Aktionstypen und Reihenfolge im gebauten Paket und wertet Bedingungen für Installation, Upgrade, Reparatur und Deinstallation aus, ohne diese auszuführen. `Invoke-OutlookTalkIfbLifecycleTests.ps1` deckt Wiederanlauf mit erhaltenem Journal und Vorgängerwiederherstellung ab. Echte MSI-Installation und Offline-Hive-Einbindung benötigen weiterhin einen Installationstest-Rechner.

## Netzwerk-Endpunkte

Das Add-in verwendet Nextcloud-**OCS**- und **WebDAV**-Endpunkte.

`NextcloudUriValidator` normalisiert die konfigurierte Basis-URL vor der Erstellung authentifizierter Services. Die Laufzeitkonfiguration lehnt explizites HTTP, Benutzerinformationen, Query und Fragment ab. Absolute Login-Flow- und Passwort-Policy-Endpunkte aus Nextcloud-Antworten müssen HTTPS verwenden und zu Schema, Host und Port der konfigurierten Basis-URL passen.

Authentifizierungsalias und DAV-Identität bleiben getrennt: Basic Auth verwendet den eingegebenen Login, während `Services/NextcloudUserIdentityService.cs` die kanonische UID über `GET /ocs/v2.php/cloud/user?format=json` ermittelt. Benutzerbezogene FileLink-, CardDAV- und CalDAV-Pfade verwenden ausschließlich `ocs.data.id`; eine fehlende UID ist ein Fehler.

Talk (Auswahl):

- Capabilities/Versionshinweis: `GET /ocs/v2.php/cloud/capabilities`
- Raum erstellen: `POST /ocs/v2.php/apps/spreed/api/v4/room`
- Raum löschen: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>`
- Lobby-Timer: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/webinar/lobby`
- Sichtbarkeit: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/listable`
- Beschreibung: `PUT /ocs/v2.php/apps/spreed/api/v4/room/<token>/description`
- Teilnehmer hinzufügen: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants`
- Teilnehmer lesen: `GET /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants?includeStatus=true`
- Moderator hochstufen: `POST /ocs/v2.php/apps/spreed/api/v4/room/<token>/moderators`
- Raum selbst verlassen: `DELETE /ocs/v2.php/apps/spreed/api/v4/room/<token>/participants/self`

Freigaben:

- Capabilities und erforderliche Serverversion: `GET /ocs/v2.php/cloud/capabilities?format=json`
- Aktuelle kanonische Benutzer-ID: `GET /ocs/v2.php/cloud/user?format=json`
- Öffentliche Freigabe erstellen: `POST /ocs/v2.php/apps/files_sharing/api/v1/shares`
- Upload/Ordneranlage: `remote.php/dav/...` (WebDAV)
- Dateien und Speicherwerte des Benutzers lesen: `PROPFIND /remote.php/dav/files/<user>/...` mit Tiefe eins
- Vorschau einer ausgewählten Datei: authentifiziertes, bytebegrenztes `GET /index.php/core/preview.png?file=...`; fehlt die generierte Vorschau, folgt `GET /remote.php/dav/files/<user>/...` nur für unterstützte Rasterbilder, deren Original höchstens 5 MiB groß ist
- Ausgewählte Nextcloud-Datei in die Freigabe kopieren: `COPY /remote.php/dav/files/<user>/...` mit absolutem `Destination` im selben Konto
- Optionaler Bulk-Upload kleiner Dateien: `POST /remote.php/dav/bulk` (`multipart/related`, nur bei exakt `ocs.data.capabilities.dav.bulkupload = "1.0"`)
- Upload großer Dateien: `MKCOL /remote.php/dav/uploads/<user>/<upload-id>`, Chunk-`PUT`s, danach `MOVE /remote.php/dav/uploads/<user>/<upload-id>/.file` zum Zielpfad

Secrets (optionale separate Passwortzustellung):

- Verschlüsseltes Secret erstellen: `POST /ocs/v2.php/apps/secrets/api/v1/secrets`
- Öffentlicher Einmal-Link: `/index.php/apps/secrets/share/<uuid>#<local-key>`
- Der Schlüssel bleibt im URL-Fragment und wird nicht an Nextcloud übertragen.

IFB (DAV über lokalen Proxy):

- Reservierter Listener-Namespace: `http://127.0.0.1:<ifb-port>/nc-ifb/` (Standardport `7777`)
- Akzeptierter Outlook-Pfad: `/nc-ifb/<request-secret>/freebusy/<address>.vfb`; das Secret wird für jeden Outlook-Prozess erzeugt und nicht persistiert. Der DPAPI-geschützte Profilzustand enthält nur den Registry-Besitzstand
- Anfragen ohne Request-Secret erhalten `404`
- Der Proxy greift auf CalDAV- und Adressbuch-Endpunkte unter `remote.php/dav/...` zu.

Antwortvertrag des Systemadressbuchs:

- `Services/IfbAddressBookCache.cs` prüft den vollständigen CardDAV-Export, bevor Kontaktzuordnungen oder der letzte Erfolgszeitpunkt ersetzt werden. Der Cache bleibt an Outlook-Profil, Nextcloud-Basis-URL und konfigurierten Login gebunden; der Anfragepfad verwendet die aufgelöste kanonische UID.
- Nur erfolgreiche HTTP-Antworten und HTTP 404 gelangen in die vCard-Prüfung. Ein vollständiger, nicht leerer Export mit nutzbaren Benutzer-UIDs wird unabhängig vom Inhaltstyp akzeptiert, auch bei HTTP 404. Andere HTTP-Fehler bleiben auch mit scheinbar gültigem Export Fehler; der gemeinsame HTTP-Client bleibt unverändert.
- Leere Exporte benötigen sowohl einen erfolgreichen HTTP-Status als auch einen vCard-Inhaltstyp. HTML/JSON, abgeschnittene oder gemischte Exporte, verschachtelte Karten sowie überzählige oder falsch benannte Komponentengrenzen ablehnen. Legale Zeilenfaltung und Kontakte mit UID ohne E-Mail für Benutzersuche und Moderatoren erhalten.
- Ein fehlgeschlagener Abruf erhält den letzten gültigen Cache, markiert den aktuellen Kontokontext jedoch für einen erneuten Versuch. Nachfolgende Zugriffe dürfen den Fehler nicht mit diesem alten Cache verdecken. Die Talk-Teilnehmerzuordnung muss sämtliche Empfänger vor Änderungen auflösen und darf bei Adressbuchfehlern nicht auf Gasteinladungen ausweichen.
- Keine Antwortinhalte oder Kontaktdaten protokollieren. `Invoke-OutlookTalkIfbLifecycleTests.ps1` prüft Antwortvalidierung, kontogebundene Cache-Wiederherstellung und Teilnehmerzuordnung über die produktiven Pfade.

Updateprüfung:

- Homepage-Endpunkt: `GET https://nc-connector.de/wp-json/ncc/v1/update-check`
- Query-Werte: `product=outlook`, installierte Version, Kanal und täglich wechselnder Client-Hash
- Release- und Download-Ziele aus der Antwort bleiben nur erhalten, wenn sie HTTPS unter `github.com/nc-connector/NC_Connector_for_Outlook/releases/` verwenden. Andere Werte werden verworfen.
- `UpdateAvailable` wird lokal aus installierter und gemeldeter Version berechnet.
- Die Homepage liefert nur Release-Metadaten und zählt einen anonymen Client pro Tag.

## Lokalisierung (i18n)

- Übersetzungen liegen unter `src/NcTalkOutlookAddIn/Resources/_locales/<sprache>/messages.json`.
- `src/NcTalkOutlookAddIn/Utilities/Strings.cs` lädt die aktive Sprache und formatiert Platzhalter.
- Die Standardsprache ist Deutsch (`de`). Die Outlook-/Office-Oberflächensprache hat Vorrang; nicht unterstützte Sprachen fallen auf Deutsch und danach Englisch zurück.
- Sprachabhängige Links müssen auf die passende deutsche oder englische Anleitung zeigen.

Die vollständige Sprachliste und der Pflegeablauf stehen in `Translations.md`.

## Logging

`ManagedSetupPolicy` löst `DebugLoggingEnabled` und `LogAnonymizationEnabled` unabhängig nach der vorhandenen Registry-Priorität auf. Bereits ein vorhandener Wert aktiviert Enterprise Rollout und sperrt die gesamte Logging-Gruppe. Fehlende Geschwisterwerte verwenden false/true. Fehlerhafte ausgewählte Werte verwenden nur ihren jeweiligen Standard, behalten die Verwaltungssperre und erzeugen einen Konfigurationshinweis sowie einen immer geschriebenen Diagnoseeintrag; sie führen keine Netzwerksperre ein. Nachrangige ungültige Werte werden ignoriert, wenn ein höherrangiges Feld ausgewählt wurde.

`AddinSettings` liefert über die vorhandenen Getter effektive Logging-Werte und behält die ursprünglichen lokalen Werte für XML-Persistenz. Klone übernehmen die Überlagerung; nach deren Entfernung gelten wieder die lokalen Werte. `SettingsForm` zeigt die effektiven Werte an, schreibt aber keine gesperrte Vorgabe als lokale Auswahl zurück. Busy-State-Änderungen erhalten die Sperren. Der vorhandene Pfad `ConfigureDiagnosticsLogger` übernimmt die effektiven Werte beim Start sowie nach Speichern/Zurücknehmen; verpflichtende Geheimnis-Maskierung und immer aktive Fehlerprotokollierung bleiben unverändert. Produktive Policy-Tests prüfen Registry-Auflösung, datensparsame Defaults, XML, UI und den Laufzeitlogger.

- Aktivierung: **Einstellungen -> Debuggen -> Debug-Logdatei schreiben**
- Option (Standard aktiv): **Logs anonymisieren**
- Tägliche Datei: `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`
- Runtime-Exceptions laufen über `DiagnosticsLogger.LogException(...)` und werden auch bei deaktiviertem Debug-Logging geschrieben.
- Aufbewahrung: die letzten sieben Tageslogs behalten und Dateien, die älter als 30 Tage sind, nach Möglichkeit entfernen.
- Authorization-Werte, URL-Zugangsdaten, strukturierte Token-/Passwortfelder, Talk-/Share-Pfadtoken und Secret-Fragmente werden vor jedem Schreiben maskiert, auch wenn die optionale Anonymisierung aus ist.
- Die Anonymisierung maskiert zusätzlich Nextcloud-URL/Basis-Host, Benutzerkennungen, E-Mail-Adressen und lokale Benutzerpfade.

Kategorien:

- `CORE`: Start, Einstellungen und Registry
- `API`: HTTP-Aufrufe und Statuscodes
- `TALK`: Raum-Lebenszyklus, Lobby und Delegation
- `FILELINK`: Upload, Freigabe, Bereinigung und Passwort-Follow-up
- `IFB`: Anfragen, Cache und Outlook-Registry

FileLink-Uploadpfade protokollieren Uploadplan, Wiederholungen, periodischen Gesamtfortschritt und Abschlusszusammenfassung, aber nicht jeden erfolgreichen Datei-Request.

## Kompatibilität und Versionsprüfungen

### Outlook-Bitness

Outlook kann als 32-Bit-Anwendung auf einem 64-Bit-Windows installiert sein. Das MSI registriert das COM-Add-in deshalb in beiden Registry-Ansichten:

- 64-Bit: `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`
- 32-Bit: `HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn`

Die Definition liegt in `installer/Product.wxs`.

### Nextcloud-Funktionserkennung

Alle Add-in-Funktionen setzen Nextcloud 32 oder neuer voraus. `NextcloudCapabilitiesService` liest die authentifizierte OCS-Capabilities-Antwort, prüft die strukturierte Versionsnummer und speichert den typisierten Snapshot fünf Minuten pro Server-/Benutzerkombination zwischen. Antworten ohne auswertbare Version und ältere Server werden abgelehnt.

Die DAV-Bulk-Funktion wird nur bei exakt gemeldetem `ocs.data.capabilities.dav.bulkupload = "1.0"` und einem Uploadplan mit ausreichender Request-Einsparung verwendet. Ohne diese Voraussetzungen bleiben Direct- und Chunked-Upload verfügbar.

### WinForms-Theme

`Utilities/UiThemeManager.cs` liest nach Möglichkeit das Office-/Outlook-Theme und fällt danach auf das Windows-App-Theme zurück. Im Hochkontrastmodus gelten die Systemfarben.

## Build und Release

### Was `build.ps1` macht

1. COM-Add-in (`NcTalkOutlookAddIn.sln`) mit MSBuild bauen
2. Assembly-Version aus `NcTalkOutlookAddIn.dll` lesen
3. WiX-v6-Installer (`installer/NcConnectorOutlookInstaller.wixproj`) bauen
4. MSI nach `dist/` kopieren

### Versionierung

- `src/NcTalkOutlookAddIn/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`

`build.ps1` leitet daraus die MSI `ProductVersion` ab (Format `Major.Minor.Build`).

**MSI-Upgrade-Kompatibilität**

Wichtig für Updates:

- UpgradeCode bleibt stabil (siehe `installer/Product.wxs`)
- COM GUID / ProgId bleiben stabil (siehe `NextcloudTalkAddIn.cs`)

### Release-Checkliste

1) Version bump
2) Bei geaenderten vendorten Abhaengigkeiten: `VENDOR.md` aktualisieren
3) `.\build.ps1 -Configuration Release`
4) MSI installieren/upgrade testen (alte Version → neue Version)
5) Talk + Filelink + IFB Smoke-Test
6) MSI ggf. signieren (falls in der Umgebung erforderlich)

## Lokales Testen

Die automatisierten Prüfungen unter `tools/ci/` laufen über die Jobs in `.github/workflows/outlook-build-checks.yml`. Outlook-COM-Verhalten wird anschließend in Outlook geprüft.

Vorgeschlagener Ablauf:

1. Debug-Logging in den Einstellungen aktivieren.
2. Kalender: Neuen Termin erstellen, einen Talk-Link einfügen, speichern, Startzeit ändern und erneut speichern.
3. Kalender: Outlook neu starten, denselben Termin öffnen, die Startzeit ändern und erneut speichern.
4. Kalender: Teilnehmer hinzufügen und erneut speichern.
5. Kalender: Die Raumlöschung gespeicherter Termine aktivieren und gespeicherte Talk-Termine einmal aus dem geöffneten Termin und einmal aus der Kalenderansicht löschen. Prüfen, dass der jeweilige Raum entfernt wird. Den Test einmal mit einem vorübergehenden Nextcloud-Fehler wiederholen, Outlook neu starten und den vorgemerkten Retry prüfen.
6. Mail: Freigabe-Wizard ausführen, ein oder zwei kleine Dateien hochladen, den Freigabeblock einfügen und an das eigene Konto senden.
7. Mail: Eine Freigabe einfügen und die Nachricht vor dem Speichern verwerfen; prüfen, dass der exakte Serverordner entfernt wird. Mit Speichern oder AutoSave wiederholen und prüfen, dass die Freigabe bestehen bleibt.
8. Mail: Bei separater Passwortzustellung die Freigabe erstellen und im selben Verfassen-Fenster auf **Senden** klicken. Prüfen, dass der fertige Follow-up sofort über dasselbe wirksame Outlook-Konto abgeht und kein NC-Connector-Passwortentwurf verbleibt. Mit verzögerter/Offline-Übermittlung wiederholen und die dokumentierte direkte Grenze prüfen: Der Follow-up wartet nicht auf die Hauptmail im Postausgang. Ein geschlossener und erneut geöffneter Hauptentwurf oder eine `.oft`-Vorlage mit vorhandenem Freigabeblock liegt außerhalb des unterstützten Ablaufs und benötigt vor dem Versand eine neue Freigabe.
9. IFB: IFB aktivieren, URL-Reservierung und TCP-Listener prüfen und danach im Outlook-Terminplanungs-Assistenten eine Adresse verwenden, deren Domain vom konfigurierten Nextcloud-Login und -Host abweicht. Eine direkte Anfrage ohne Request-Secret muss `404` liefern.
10. **Einstellungen -> Erweitert -> Jetzt prüfen** ausführen und kontrollieren, dass aktuelle Version, letzte Prüfung, Download-Link und Änderungsübersicht ohne blockierte Outlook-Oberfläche aktualisiert werden.

## Referenz der X-NCTALK-*-Eigenschaften

Das Add-in speichert Talk-Terminmetadaten ausschließlich als Outlook-`UserProperties` mit `X-NCTALK-*`-Namen. Frühere NC-Connector-spezifische Outlook-Eigenschaften werden nicht mehr gelesen oder geschrieben.

Sofern nicht anders angegeben:

- Die Werte sind Text (`OlUserPropertyType.olText`).
- Boolean-Werte werden als `TRUE` oder `FALSE` gespeichert.
- Zeitstempel sind Unix-Epoch-Sekunden in UTC und invariantem Zahlenformat.

Die primären Schreibpfade liegen in `TalkAppointmentController.ApplyRoomToAppointment(...)` und `TalkAppointmentController.PersistCoreIcalProperties(...)`.

### Eigenschaften

| Eigenschaft | Zweck | Format | Verwendung |
| --- | --- | --- | --- |
| `X-NCTALK-TOKEN` | Talk-Raumtoken | Text | Voraussetzung für Subscription, Retry und optionale Raumlöschung; beliebige Talk-URLs sind keine Löschquelle. |
| `X-NCTALK-URL` | Talk-Raum-URL | Absolute URL | Lokale Outlook-Metadaten, nicht als Löschquelle verwendet. |
| `X-NCTALK-LOBBY` | Lobby aktiv | `TRUE` / `FALSE` | Steuert Lobby-Aktualisierungen beim Speichern. |
| `X-NCTALK-START` | Terminstart | Unix-Sekunden | Laufzeit- und Lobby-Zustand. |
| `X-NCTALK-EVENT` | Raummodus | `event` / `standard` | Unterscheidet Ereignis- und Gruppenraum. |
| `X-NCTALK-OBJECTID` | Zeitfenster | `<start>#<end>` | Lokale Terminmetadaten. |
| `X-NCTALK-ADD-USERS` | Interne Teilnehmer synchronisieren | `TRUE` / `FALSE` | Getrennter Schalter für Nextcloud-Benutzer. |
| `X-NCTALK-ADD-GUESTS` | Externe Teilnehmer synchronisieren | `TRUE` / `FALSE` | Getrennter Schalter für Gäste. |
| `X-NCTALK-DELEGATE` | Ziel der Moderatorübergabe | Benutzer-ID | Erkennung und Wiederholung ausstehender Übergaben. |
| `X-NCTALK-DELEGATE-NAME` | Anzeigename des Ziels | Text | Lokale Anzeigeinformation. |
| `X-NCTALK-DELEGATED` | Status der Übergabe | `TRUE` / `FALSE` | Markiert eine noch ausstehende oder abgeschlossene Übergabe. |
| `X-NCTALK-DELEGATE-READY` | Bereitschaftsmarker | `TRUE` | Bestandteil des Outlook-Delegationsablaufs. |

## Erweiterungspunkte

### Neue Einstellung ergänzen

1. Eigenschaft in `Settings/AddinSettings.cs` anlegen.
2. Lesen und Schreiben in `Settings/SettingsStorage.cs` ergänzen.
3. Oberfläche in `UI/SettingsForm.cs` erweitern.
4. Den Schlüssel in allen Dateien unter `Resources/_locales/` ergänzen.
5. Persistenz-, Locale- und UI-Tests anpassen.

### Neuen Nextcloud-API-Aufruf ergänzen

1. Den zuständigen Service erweitern:
   - Talk: `Services/TalkService.cs`
   - FileLink-Orchestrierung: `Services/FileLinkService.cs`
   - DAV: `Services/FileLinkDavClient.cs`
   - Transfers: `Services/FileLinkTransferService.cs`
   - OCS-Freigabe: `Services/FileLinkShareClient.cs`
2. Für OCS/JSON den gemeinsamen `NcHttpClient` und `NcJson` verwenden.
3. Bei Bedarf ein Request-/Response-Modell unter `Models/` anlegen.
4. Operationslogging, Fehlerabbildung und Tests ergänzen.
5. Ribbon-Abläufe über den zuständigen Controller verdrahten; `NextcloudTalkAddIn.cs` bleibt Composition Root.

### Neue Übersetzung ergänzen

1. Zugriffseigenschaft in `Utilities/Strings.cs` anlegen.
2. Den Schlüssel in jeder `Resources/_locales/<sprache>/messages.json` ergänzen.
3. Locale- und Nutzungsprüfungen ausführen und die Oberfläche kontrollieren.
