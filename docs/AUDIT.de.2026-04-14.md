# Audit-Bericht (2026-04-14) — NC Connector for Outlook 3.0.1

## Scope
- Quellcode-Audit fuer `src/NcTalkOutlookAddIn` (Core, Controller, UI, Utilities, Services)
- Doku-Konsistenz-Audit fuer `README*.md`, `docs/ADMIN*.md`, `docs/DEVELOPMENT*.md`, `CHANGELOG.md`
- Build-Check in der aktuellen lokalen Shell-Umgebung

## Ergebnis (Kurzfassung)
- Doppelter/weitergeleiteter Hilfscode im Compose-Recipient-Flow wurde weiter reduziert.
- COM-Release-Pfade im Compose-Share-Lifecycle wurden auf die zentrale Utility vereinheitlicht.
- Komplexe Attachment-Pre-Add-Logik ist dokumentiert und mit gezielten Kommentarbloeken nachvollziehbar.
- Doku-Dateien sind in den kritischen Verhaltensthemen konsistent (best-effort Pre-Add, Outlook/Exchange-Hardlimit, Fallback ueber Share-Button).
- Kein offensichtlicher toter Codepfad im geprueften Scope gefunden, der ohne Funktionsverlust entfernbar waere.

## Umgesetzte Bereinigungen in diesem Audit-Durchlauf
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.cs`
  - verbleibende Recipient-Normalisierungs-Forwarder entfernt:
    - `BuildNormalizedRecipientCsv(...)`
    - `CountRecipientsInCsv(...)`
    - `NormalizeRecipientAddress(...)`
  - no-op Interface-Callbacks (`OnAddInsUpdate`, `OnStartupComplete`) klar als absichtliche IDTExtensibility2-Pflichtpfade kommentiert.
- `src/NcTalkOutlookAddIn/NextcloudTalkAddIn.MailComposeSubscription.cs`
  - Recipient-Normalisierung ruft direkt `ComposeShareLifecycleController` auf (keine Root-Weiterleitung mehr).
  - Namespace-Import fuer Controller explizit gesetzt (`using NcTalkOutlookAddIn.Controllers;`).
- `src/NcTalkOutlookAddIn/Controllers/ComposeShareLifecycleController.cs`
  - wiederholte COM-Release-Logik auf `ComInteropScope.TryRelease(...)` umgestellt.
  - Ablauf im Fehlerfall fuer Passwort-Follow-up (Auto-Send -> manueller Fallback) mit Kommentar klar dokumentiert.
- `src/NcTalkOutlookAddIn/UI/FileLinkWizardForm.cs`
  - missverstaendliche lokale Variable `obsoleteContext` in `previousContext` umbenannt.

## Doku-Konsistenzpruefung (Bestand gegen Verhalten)
- Konsistent dokumentiert:
  - `README.md`, `README.de.md`, `docs/ADMIN.md`, `docs/ADMIN.de.md`:
    - Pre-Add ist best effort, nicht garantiert
    - Outlook/Exchange kann grosse Anhaenge vor Add-in-Events blockieren
    - in diesem Fall soll der Button `Insert Nextcloud share` / `Nextcloud Freigabe hinzufuegen` genutzt werden
- Build-Doku konsistent:
  - `build.ps1 -SkipIceValidation` ist in `docs/DEVELOPMENT.md` und `docs/DEVELOPMENT.de.md` beschrieben.

## Technische Befunde (priorisiert)

### Hoch
1. Outlook-Hostlimit vor Add-in-Events bleibt harte Produktgrenze
- Wenn Outlook/Exchange einen Anhang vor `BeforeAttachmentAdd` blockiert, kann die Automatisierung technisch nicht eingreifen.
- Status: korrekt dokumentiert, aber nicht rein clientseitig loesbar.

### Mittel
2. Sehr grosse UI-Dateien bleiben Wartungsrisiko
- `UI/FileLinkWizardForm.cs` ~3254 Zeilen
- `UI/SettingsForm.cs` ~2148 Zeilen
- `UI/TalkLinkForm.cs` ~1072 Zeilen
- Empfehlung: weiter in Presenter/State-Objekte schneiden.

3. Service-Layer ohne formale Contract-Tests
- Kritische HTTP-Flows (`TalkService`, `FileLinkService`) sind weiterhin smoke-test-lastig.
- Empfehlung: minimaler Satz API-Contract-Tests mit mockbaren Responses.

### Niedrig
4. Lokaler Build in dieser Shell nicht reproduzierbar ohne Office/Extensibility-Referenzen
- `dotnet build NcTalkOutlookAddIn.sln -c Release` lief in dieser Umgebung nicht durch, weil `Extensibility`, `office`, `stdole`, `Microsoft.Office.Interop.Outlook` lokal nicht aufloesbar waren.
- Das ist ein Umgebungs-/Tooling-Thema, kein Nachweis fuer einen funktionalen Regressionseffekt der hier gemachten Cleanup-Edits.

## Was explizit nicht entfernt wurde (bewusst)
- IDTExtensibility2-Callbacks `OnAddInsUpdate` und `OnStartupComplete`:
  - no-op, aber erforderlich fuer den COM-Lifecyclevertrag.
- Legacy-Migrationspfade in `SettingsStorage`:
  - nicht tot, sondern noetig fuer Upgrades von Altinstallationen.

## Empfohlene naechste Schritte
1. `ComposeAttachmentAutomationController` als naechste Entkopplungsstufe aus `MailComposeSubscription` extrahieren.
2. COM-Scope-Konvention (`ComInteropScope`) weiter in verbleibende manuelle Release-Pfade migrieren.
3. Contract-Tests fuer Talk/Share-HTTP-Flows aufbauen (Create/Update/Delete + DAV Edge-Cases).
4. Build-Agent dokumentieren (erforderliche Office-Interop/Extensibility-Komponenten), um lokale Build-Differenzen zu vermeiden.
