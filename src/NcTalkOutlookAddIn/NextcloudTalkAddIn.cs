/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    /**
     * Einstiegs- und Ribbon-Implementierung fuer das Nextcloud Talk Outlook Add-in.
     * Die Klasse registriert sich als klassisches COM Add-in ueber IDTExtensibility2
     * und stellt gleichzeitig das Ribbon-XML fuer Terminfenster bereit.
     */
    [ComVisible(true)]
    [Guid("A8CC9257-A153-4A01-AB35-D66CB3D44AAA")]
    [ProgId("NcTalkOutlook.AddIn")]
    public sealed class NextcloudTalkAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private const string SenderAddressProperty = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E";
        private const string AlternateSenderProperty = "http://schemas.microsoft.com/mapi/proptag/0x0065001E";
        private const string MuzzleAccountProperty = "NcMuzzleAccountKey";

        private Outlook.Application _outlookApplication;
        private SettingsStorage _settingsStorage;
        private AddinSettings _currentSettings;
        private readonly Dictionary<string, AppointmentSubscription> _activeSubscriptions = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _knownRoomTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByToken = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private FreeBusyManager _freeBusyManager;
        private bool _itemSendHooked;
        private bool _itemLoadHooked;
        private readonly List<OutboxSubscription> _outboxSubscriptions = new List<OutboxSubscription>();
        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;
        private readonly List<ExplorerSubscription> _explorerSubscriptions = new List<ExplorerSubscription>();
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByEntryId = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private bool _initialScanPerformed;
        private readonly Queue<MuzzlePendingInfo> _pendingMuzzleItems = new Queue<MuzzlePendingInfo>();
        private readonly object _pendingMuzzleLock = new object();

        private const string PropertyToken = "NcTalkRoomToken";
        private const string PropertyRoomType = "NcTalkRoomType";
        private const string PropertyLobby = "NcTalkLobbyEnabled";
        private const string PropertySearchVisible = "NcTalkSearchVisible";
        private const string PropertyPasswordSet = "NcTalkPasswordSet";
        private const string PropertyStartEpoch = "NcTalkStartEpoch";
        private const string PropertyDataVersion = "NcTalkDataVersion";
        private const string BodySectionHeader = "Nextcloud Talk";
        private const string BodyJoinLine = "Jetzt an der Besprechung teilnehmen :";
        private const string BodyPasswordPrefix = "Passwort:";
        private const string BodyHelpQuestion = "Benoetigen Sie Hilfe?";
        private const string BodyHelpLink = "https://docs.nextcloud.com/server/latest/user_manual/de/talk/join_a_call_or_chat_as_guest.html";
        private const int PropertyVersionValue = 1;

        private static void LogCore(string message)
        {
            DiagnosticsLogger.Log("Core", message);
        }

        private static void LogSettings(string message)
        {
            DiagnosticsLogger.Log("Settings", message);
        }

        private static void LogTalk(string message)
        {
            DiagnosticsLogger.Log("Talk", message);
        }

        /**
         * Outlook ruft diese Methode beim Laden des Add-ins auf.
         * Hier speichern wir die Application-Instanz fuer spaetere Aktionen.
         */
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            _settingsStorage = new SettingsStorage();
            _currentSettings = _settingsStorage.Load();
            DiagnosticsLogger.SetEnabled(_currentSettings != null && _currentSettings.DebugLoggingEnabled);
            LogCore("Add-in verbunden (Outlook-Version=" + (_outlookApplication != null ? _outlookApplication.Version : "unbekannt") + ").");
            if (_currentSettings != null)
            {
                LogSettings("Einstellungen geladen (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", Debug=" + _currentSettings.DebugLoggingEnabled + ").");
            }
            _freeBusyManager = new FreeBusyManager(_settingsStorage.DataDirectory);
            _freeBusyManager.Initialize(_outlookApplication);

            EnsureItemSendHook();
            EnsureOutboxHook();
            EnsureItemLoadHook();
            EnsureInspectorHook();
            EnsureExplorerHook();
            ApplyIfbSettings();
        }

        /**
         * Outlook signalisiert, dass das Add-in entladen wird.
         * Cleanup-Hooks folgen spaeter, sobald wir Ressourcen halten.
         */
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            UnhookItemSend();
            UnhookOutbox();
            UnhookItemLoad();
            UnhookInspector();
            UnhookExplorer();
            if (_freeBusyManager != null && _currentSettings != null && _currentSettings.IfbEnabled)
            {
                try
                {
                    var clone = _currentSettings.Clone();
                    clone.IfbEnabled = false;
                    _freeBusyManager.ApplySettings(clone);
                }
                catch
                {
                }
            }
            if (_freeBusyManager != null)
            {
                _freeBusyManager.Dispose();
            }
            _freeBusyManager = null;
            _outlookApplication = null;
            LogCore("Add-in getrennt (removeMode=" + removeMode + ").");
        }

        /**
         * Wird nach dem Laden aller Add-ins aufgerufen; derzeit nicht verwendet.
         */
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /**
         * ErmÃ¶glicht es, vor dem Laden weitere Initialisierungen vorzunehmen.
         */
        public void OnStartupComplete(ref Array custom)
        {
        }

        /**
         * Wird aufgerufen, wenn Outlook beendet wird; reserviert fuer spaetere Cleanup-Schritte.
         */
        public void OnBeginShutdown(ref Array custom)
        {
            UnhookItemSend();
            UnhookOutbox();
            UnhookItemLoad();
            UnhookInspector();
            UnhookExplorer();
            if (_freeBusyManager != null && _currentSettings != null && _currentSettings.IfbEnabled)
            {
                try
                {
                    var clone = _currentSettings.Clone();
                    clone.IfbEnabled = false;
                    _freeBusyManager.ApplySettings(clone);
                }
                catch
                {
                }
            }
            if (_freeBusyManager != null)
            {
                _freeBusyManager.Dispose();
            }
            _freeBusyManager = null;
        }

        public string GetCustomUI(string ribbonID)
        {
            if (string.Equals(ribbonID, "Microsoft.Outlook.Appointment", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabAppointment'>
        <group id='NcTalkGroup' label='{0}'>
          <button id='NcTalkCreateButton'
                  label='{1}'
                  size='large'
                  getImage='OnGetButtonImage'
                  onAction='OnTalkButtonPressed'
                  screentip='{2}'
                  supertip='{3}' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>",
                    EscapeXml(Strings.RibbonAppointmentGroupLabel),
                    EscapeXml(Strings.RibbonTalkButtonLabel),
                    EscapeXml(Strings.RibbonTalkButtonScreenTip),
                    EscapeXml(Strings.RibbonTalkButtonSuperTip));
            }

            if (string.Equals(ribbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='NcTalkExplorerTab' label='{0}' insertAfterMso='TabMail'>
        <group id='NcTalkExplorerGroup' label='{1}'>
          <button id='NcTalkSettingsExplorerButton'
                  label='{2}'
                  size='large'
                  getImage='OnGetButtonImage'
                  onAction='OnSettingsButtonPressed'
                  screentip='{3}'
                  supertip='{4}' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>",
                    EscapeXml(Strings.RibbonExplorerTabLabel),
                    EscapeXml(Strings.RibbonExplorerGroupLabel),
                    EscapeXml(Strings.RibbonSettingsButtonLabel),
                    EscapeXml(Strings.RibbonSettingsScreenTip),
                    EscapeXml(Strings.RibbonSettingsSuperTip));
            }

            if (string.Equals(ribbonID, "Microsoft.Outlook.Mail.Compose", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabNewMailMessage'>
        <group id='NcTalkMailGroup' label='{0}'>
          <button id='NcTalkFileLinkButton'
                  label='{1}'
                  size='large'
                  getImage='OnGetButtonImage'
                  onAction='OnFileLinkButtonPressed'
                  screentip='{2}'
                  supertip='{3}' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>",
                    EscapeXml(Strings.RibbonMailGroupLabel),
                    EscapeXml(Strings.RibbonFileLinkButtonLabel),
                    EscapeXml(Strings.RibbonFileLinkButtonScreenTip),
                    EscapeXml(Strings.RibbonFileLinkButtonSuperTip));
            }

            return null;
        }

        /**
         * Outlook uebergibt das Ribbon-Handle unmittelbar nach dem Laden.
         * Wir merken uns die Instanz fuer spaetere Refresh-Operationen.
         */
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            // Ribbon-Handle aktuell nicht benötigt; Methode bleibt als Interface-Stub bestehen.
        }

        public void OnTalkButtonPressed(IRibbonControl control)
        {
            EnsureSettingsLoaded();

            if (!SettingsAreComplete())
            {
                LogTalk("Talk-Link abgebrochen: Einstellungen unvollstaendig.");
                MessageBox.Show(
                    Strings.ErrorMissingCredentials,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                OnSettingsButtonPressed(control);
                return;
            }

            if (!EnsureAuthenticationValid(control))
            {
                LogTalk("Talk-Link abgebrochen: Authentifizierung fehlgeschlagen.");
                return;
            }

            Outlook.AppointmentItem appointment = GetActiveAppointment();
            if (appointment == null)
            {
                LogTalk("Talk-Link abgebrochen: Kein aktiver Termin gefunden.");
                MessageBox.Show(
                    Strings.ErrorNoAppointment,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var subject = appointment.Subject ?? string.Empty;
            var start = appointment.Start == DateTime.MinValue ? DateTime.Now : appointment.Start;
            var end = appointment.End == DateTime.MinValue ? start.AddHours(1) : appointment.End;
            LogTalk("Talk-Link gestartet (Betreff='" + subject + "', Start=" + start.ToString("o") + ", Ende=" + end.ToString("o") + ").");

            using (var dialog = new TalkLinkForm(_currentSettings ?? new AddinSettings(), subject, start, end))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    LogTalk("Talk-Link Dialog abgebrochen.");
                    return;
                }

                var request = new TalkRoomRequest
                {
                    Title = dialog.TalkTitle,
                    Password = dialog.TalkPassword,
                    LobbyEnabled = dialog.LobbyUntilStart,
                    SearchVisible = dialog.SearchVisible,
                    RoomType = dialog.SelectedRoomType,
                    AppointmentStart = appointment.Start,
                    AppointmentEnd = appointment.End,
                    Description = BuildInitialRoomDescription(dialog.TalkPassword)
                };
                LogTalk("Raumanfrage vorbereitet (Titel='" + request.Title + "', Typ=" + request.RoomType + ", Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ", PasswortGesetzt=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = GetUserPropertyText(appointment, PropertyToken);
                if (!string.IsNullOrWhiteSpace(existingToken))
                {
                    LogTalk("Vorhandener Raum (Token=" + existingToken + ") gefunden – Ersetzanfrage.");
                    var overwrite = MessageBox.Show(
                        Strings.ConfirmReplaceRoom,
                        Strings.DialogTitle,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (overwrite != DialogResult.Yes)
                    {
                        LogTalk("Ersetzen abgelehnt – Vorgang beendet.");
                        return;
                    }

                    var existingType = GetRoomType(appointment);
                    bool existingIsEvent = existingType.HasValue && existingType.Value == TalkRoomType.EventConversation;
                    LogTalk("Versuche vorhandenen Raum zu loeschen (Event=" + existingIsEvent + ").");

                    if (!TryDeleteRoom(existingToken, existingIsEvent))
                    {
                        LogTalk("Loeschen des vorhandenen Raums fehlgeschlagen.");
                        return;
                    }
                }

                TalkRoomCreationResult result;
                try
                {
                    using (new WaitCursorScope())
                    {
                        LogTalk("Sende CreateRoom-Request an Nextcloud.");
                        var service = CreateTalkService();
                        result = service.CreateRoom(request);
                    }
                    LogTalk("Raum erfolgreich erstellt (Token=" + result.RoomToken + ", URL=" + result.RoomUrl + ", Event=" + result.CreatedAsEventConversation + ").");
                }
                catch (TalkServiceException ex)
                {
                    LogTalk("Talk-Raum konnte nicht erstellt werden: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoom, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    LogTalk("Unerwarteter Fehler beim Erstellen des Talk-Raums: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                ApplyRoomToAppointment(appointment, request, result);
                LogTalk("Raumdaten in Termin gespeichert (EntryID=" + (appointment.EntryID ?? "n/a") + ").");

                MessageBox.Show(
                    string.Format(Strings.InfoRoomCreated, request.Title),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        public void OnSettingsButtonPressed(IRibbonControl control)
        {
            EnsureSettingsLoaded();
            LogSettings("Einstellungsdialog geoeffnet.");

            using (var form = new SettingsForm(_currentSettings, _outlookApplication))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _currentSettings = form.Result ?? new AddinSettings();
                    LogSettings("Einstellungen uebernommen (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", Debug=" + _currentSettings.DebugLoggingEnabled + ").");
                    ApplyIfbSettings();
                    if (_settingsStorage != null)
                    {
                        _settingsStorage.Save(_currentSettings);
                    }

                    EnsureItemSendHook();
                }
                else
                {
                    LogSettings("Einstellungsdialog ohne Aenderungen geschlossen.");
                }
            }
        }

        /**
         * Liefert das eingebettete Talk-Icon als COM-kompatibles PictureDisp.
         */
        public stdole.IPictureDisp OnGetButtonImage(IRibbonControl control)
        {
            string resourceName = "NcTalkOutlookAddIn.Resources.talk-96.png";
            if (control != null)
            {
                if (string.Equals(control.Id, "NcTalkSettingsExplorerButton", StringComparison.OrdinalIgnoreCase))
                {
                    resourceName = "NcTalkOutlookAddIn.Resources.logo-nextcloud.png";
                }
                else if (string.Equals(control.Id, "NcTalkFileLinkButton", StringComparison.OrdinalIgnoreCase))
                {
                    resourceName = "NcTalkOutlookAddIn.Resources.nextcloud-filelink.png";
                }
            }

            using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (resourceStream == null)
                {
                    return null;
                }

                using (var bitmap = new Bitmap(resourceStream))
                {
                    return PictureConverter.ToPictureDisp(bitmap);
                }
            }
        }

        public void OnFileLinkButtonPressed(IRibbonControl control)
        {
            EnsureSettingsLoaded();
            if (!SettingsAreComplete())
            {
                MessageBox.Show(
                    Strings.ErrorMissingCredentials,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                OnSettingsButtonPressed(control);
                return;
            }

            Outlook.MailItem mail = GetActiveMailItem();
            if (mail == null)
            {
                MessageBox.Show(
                    Strings.ErrorNoMailItem,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string basePath = string.IsNullOrWhiteSpace(_currentSettings.FileLinkBasePath)
                ? "90 Freigaben - extern"
                : _currentSettings.FileLinkBasePath;

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);

            using (var wizard = new FileLinkWizardForm(configuration, basePath))
            {
                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    LogCore("Filelink erstellt (Ordner=\"" + wizard.Result.FolderName + "\").");
                    string html = FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot);
                    InsertHtmlIntoMail(mail, html);
                }
            }
        }

        private Outlook.AppointmentItem GetActiveAppointment()
        {
            if (_outlookApplication == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = _outlookApplication.ActiveInspector();
            }
            catch
            {
                inspector = null;
            }

            if (inspector != null)
            {
                return inspector.CurrentItem as Outlook.AppointmentItem;
            }

            return null;
        }

        private bool SettingsAreComplete()
        {
            return _currentSettings != null
                   && !string.IsNullOrWhiteSpace(_currentSettings.ServerUrl)
                   && !string.IsNullOrWhiteSpace(_currentSettings.Username)
                   && !string.IsNullOrWhiteSpace(_currentSettings.AppPassword);
        }

        private void ApplyIfbSettings()
        {
            if (_currentSettings != null)
            {
                DiagnosticsLogger.SetEnabled(_currentSettings.DebugLoggingEnabled);
            }

            if (_freeBusyManager == null || _currentSettings == null)
            {
                return;
            }

            try
            {
                LogCore("IFB wird angewendet (Enabled=" + _currentSettings.IfbEnabled + ", Days=" + _currentSettings.IfbDays + ", CacheHours=" + _currentSettings.IfbCacheHours + ").");
                _freeBusyManager.ApplySettings(_currentSettings);
            }
            catch (Exception ex)
            {
                LogCore("IFB konnte nicht gestartet werden: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningIfbStartFailed, ex.Message));
            }
        }

        /**
         * Fuehrt einen schnellen Verbindungstest durch und bietet bei Fehlern den Sprung in die Einstellungen an.
         */
        private bool EnsureAuthenticationValid(IRibbonControl control)
        {
            try
            {
                using (new WaitCursorScope())
                {
                    var service = CreateTalkService();
                    string response;
                    LogTalk("Starte Verifizierungs-Request fuer Anmeldedaten.");
                    if (service.VerifyConnection(out response))
                    {
                        UpdateStoredServerVersion(response);
                        LogTalk("Anmeldedaten verifiziert (Antwort=" + (string.IsNullOrEmpty(response) ? "OK" : response) + ").");
                        return true;
                    }

                    string message = string.IsNullOrEmpty(response)
                        ? "Die Anmeldedaten konnten nicht verifiziert werden."
                        : "Die Anmeldedaten konnten nicht verifiziert werden: " + response;
                    LogTalk("Anmeldedaten ungueltig: " + message);
                    return PromptOpenSettings(message, control);
                }
            }
            catch (TalkServiceException ex)
            {
                string message;
                if ((int)ex.StatusCode == 0)
                {
                    message = Strings.ErrorServerUnavailable;
                }
                else if (ex.IsAuthenticationError)
                {
                    message = string.Format(Strings.ErrorAuthenticationRejected, ex.Message);
                }
                else
                {
                    message = string.Format(Strings.ErrorConnectionFailed, ex.Message);
                }

                LogTalk("Fehler bei der Verbindungspruefung: " + message);
                return PromptOpenSettings(message, control);
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler bei der Verbindungspruefung: " + ex.Message);
                return PromptOpenSettings(string.Format(Strings.ErrorUnknownAuthentication, ex.Message), control);
            }
        }

        private bool PromptOpenSettings(string message, IRibbonControl control)
        {
            var result = MessageBox.Show(
                string.Format(Strings.PromptOpenSettings, message),
                Strings.DialogTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                OnSettingsButtonPressed(control);
            }

            return false;
        }

        private TalkService CreateTalkService()
        {
            return new TalkService(new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword));
        }

        private void ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            if (appointment == null || result == null)
            {
                return;
            }
            LogTalk("Schreibe Talk-Daten in Termin (Token=" + result.RoomToken + ", Typ=" + request.RoomType + ", Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ", PasswortGesetzt=" + (!string.IsNullOrEmpty(request.Password)) + ").");

            if (!string.IsNullOrWhiteSpace(request.Title))
            {
                appointment.Subject = request.Title.Trim();
            }

            appointment.Location = result.RoomUrl;
            appointment.Body = UpdateBodyWithTalkBlock(appointment.Body, result.RoomUrl, request.Password);

            SetUserProperty(appointment, PropertyToken, Outlook.OlUserPropertyType.olText, result.RoomToken);
            SetUserProperty(appointment, PropertyRoomType, Outlook.OlUserPropertyType.olText, request.RoomType.ToString());
            SetUserProperty(appointment, PropertyLobby, Outlook.OlUserPropertyType.olYesNo, request.LobbyEnabled);
            SetUserProperty(appointment, PropertySearchVisible, Outlook.OlUserPropertyType.olYesNo, request.SearchVisible);
            SetUserProperty(appointment, PropertyPasswordSet, Outlook.OlUserPropertyType.olYesNo, !string.IsNullOrEmpty(request.Password));
            LogTalk("Benutzerfelder aktualisiert (Token gesetzt, Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ").");

            long? startEpoch = ConvertToUnixTimestamp(appointment.Start);
            if (startEpoch.HasValue)
            {
                SetUserProperty(appointment, PropertyStartEpoch, Outlook.OlUserPropertyType.olText, startEpoch.Value.ToString(CultureInfo.InvariantCulture));
            }

            SetUserProperty(appointment, PropertyDataVersion, Outlook.OlUserPropertyType.olNumber, PropertyVersionValue);

            RegisterSubscription(appointment, result);
            TryUpdateRoomDescription(appointment, result.RoomToken, result.CreatedAsEventConversation);
        }

        private static string UpdateBodyWithTalkBlock(string existingBody, string roomUrl, string password)
        {
            var body = existingBody ?? string.Empty;
            body = RemoveExistingTalkBlock(body);

            var builder = new StringBuilder();
            builder.AppendLine(BodySectionHeader);
            builder.AppendLine();
            builder.AppendLine(BodyJoinLine);
            builder.AppendLine(roomUrl ?? string.Empty);
            builder.AppendLine();

            bool hasPassword = !string.IsNullOrWhiteSpace(password);
            if (hasPassword)
            {
                builder.AppendLine(BodyPasswordPrefix + " " + password);
                builder.AppendLine();
            }

            builder.AppendLine(BodyHelpQuestion);
            builder.AppendLine();
            builder.AppendLine(BodyHelpLink);

            if (body.Length == 0)
            {
                return builder.ToString();
            }

            if (!body.EndsWith("\r\n", StringComparison.Ordinal))
            {
                body += "\r\n";
            }

            return body + "\r\n" + builder;
        }

        private static string RemoveExistingTalkBlock(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return body;
            }

            int headerIndex = body.IndexOf(BodySectionHeader, StringComparison.Ordinal);
            if (headerIndex < 0)
            {
                return body;
            }

            int helpIndex = body.IndexOf(BodyHelpLink, headerIndex, StringComparison.Ordinal);
            if (helpIndex < 0)
            {
                return body;
            }

            int blockStart = headerIndex;
            while (blockStart > 0 && (body[blockStart - 1] == '\r' || body[blockStart - 1] == '\n'))
            {
                blockStart--;
            }

            int blockEnd = helpIndex + BodyHelpLink.Length;
            while (blockEnd < body.Length && (body[blockEnd] == '\r' || body[blockEnd] == '\n' || body[blockEnd] == ' ' || body[blockEnd] == '\t'))
            {
                blockEnd++;
            }

            var prefix = body.Substring(0, blockStart).TrimEnd('\r', '\n');
            var suffix = body.Substring(blockEnd).TrimStart('\r', '\n');

            if (prefix.Length == 0)
            {
                return suffix;
            }

            if (suffix.Length == 0)
            {
                return prefix;
            }

            return prefix + "\r\n\r\n" + suffix;
        }

        private static string BuildInitialRoomDescription(string password)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                return string.Empty;
            }

            return BodyPasswordPrefix + " " + password.Trim();
        }

        private static string BuildDescriptionPayload(Outlook.AppointmentItem appointment)
        {
            if (appointment == null)
            {
                return null;
            }

            string body = appointment.Body ?? string.Empty;
            return body.Trim();
        }

        private static void SetUserProperty(Outlook.AppointmentItem appointment, string name, Outlook.OlUserPropertyType type, object value)
        {
            if (appointment == null)
            {
                return;
            }

            var properties = appointment.UserProperties;
            var property = properties[name] ?? properties.Add(name, type, Type.Missing, Type.Missing);
            property.Value = value;
        }

        private static string GetEntryId(Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment != null ? appointment.EntryID : null;
            }
            catch
            {
                return null;
            }
        }

        private void UpdateStoredServerVersion(string response)
        {
            if (_currentSettings == null || string.IsNullOrWhiteSpace(response))
            {
                return;
            }

            Version parsed;
            if (!NextcloudVersionHelper.TryParse(response, out parsed))
            {
                return;
            }

            string versionText = parsed.ToString();
            if (string.Equals(_currentSettings.LastKnownServerVersion, versionText, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _currentSettings.LastKnownServerVersion = versionText;
            if (_settingsStorage != null)
            {
                try
                {
                    _settingsStorage.Save(_currentSettings);
                }
                catch
                {
                    // Wir ignorieren Fehler beim asynchronen Speichern der Version.
                }
            }
        }

        private static string GetUserPropertyText(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return null;
            }

            var property = appointment.UserProperties[name];
            return property != null ? property.Value as string : null;
        }

        private static bool GetUserPropertyBool(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return false;
            }

            var property = appointment.UserProperties[name];
            if (property == null)
            {
                return false;
            }

            object value = property.Value;
            if (value == null)
            {
                return false;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            if (value is int)
            {
                return (int)value != 0;
            }

            string text = value as string;
            if (!string.IsNullOrEmpty(text))
            {
                bool boolParsed;
                if (bool.TryParse(text, out boolParsed))
                {
                    return boolParsed;
                }

                int numericParsed;
                if (int.TryParse(text, out numericParsed))
                {
                    return numericParsed != 0;
                }
            }

            return false;
        }

        private void RefreshEntryBinding(AppointmentSubscription subscription)
        {
            if (subscription == null)
            {
                return;
            }

            string oldEntryId = subscription.EntryId;
            string newEntryId = GetEntryId(subscription.Appointment);

            if (string.Equals(oldEntryId, newEntryId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!string.IsNullOrEmpty(oldEntryId))
            {
                AppointmentSubscription current;
                if (_subscriptionByEntryId.TryGetValue(oldEntryId, out current) && current == subscription)
                {
                    _subscriptionByEntryId.Remove(oldEntryId);
                }
            }

            subscription.UpdateEntryId(newEntryId);

            if (!string.IsNullOrEmpty(newEntryId))
            {
                AppointmentSubscription existing;
                if (_subscriptionByEntryId.TryGetValue(newEntryId, out existing) && existing != subscription)
                {
                    existing.Dispose();
                }

                _subscriptionByEntryId[newEntryId] = subscription;
            }
        }

        private static TalkRoomType? GetRoomType(Outlook.AppointmentItem appointment)
        {
            string value = GetUserPropertyText(appointment, PropertyRoomType);
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            TalkRoomType parsed;
            if (Enum.TryParse<TalkRoomType>(value, true, out parsed))
            {
                return parsed;
            }

            return null;
        }

        private bool IsOrganizer(Outlook.AppointmentItem appointment)
        {
            if (appointment == null)
            {
                return false;
            }

            switch (appointment.MeetingStatus)
            {
                case Outlook.OlMeetingStatus.olNonMeeting:
                case Outlook.OlMeetingStatus.olMeeting:
                case Outlook.OlMeetingStatus.olMeetingCanceled:
                    return true;
                default:
                    return false;
            }
        }

        private static void RemoveUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return;
            }

            var property = appointment.UserProperties[name];
            if (property != null)
            {
                property.Delete();
            }
        }

        private void ClearTalkProperties(Outlook.AppointmentItem appointment)
        {
            RemoveUserProperty(appointment, PropertyToken);
            RemoveUserProperty(appointment, PropertyRoomType);
            RemoveUserProperty(appointment, PropertyLobby);
            RemoveUserProperty(appointment, PropertySearchVisible);
            RemoveUserProperty(appointment, PropertyPasswordSet);
            RemoveUserProperty(appointment, PropertyStartEpoch);
            RemoveUserProperty(appointment, PropertyDataVersion);
        }

        private void RegisterSubscription(Outlook.AppointmentItem appointment, TalkRoomCreationResult result)
        {
            if (result == null)
            {
                return;
            }

            RegisterSubscription(appointment, result.RoomToken, result.LobbyEnabled, result.CreatedAsEventConversation);
        }

        private void RegisterSubscription(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }
            LogTalk("Registriere Termin-Subscription (Token=" + roomToken + ", Lobby=" + lobbyEnabled + ", Event=" + isEventConversation + ").");

            string entryId = GetEntryId(appointment);
            if (!string.IsNullOrEmpty(entryId))
            {
                AppointmentSubscription existingByEntry;
                if (_subscriptionByEntryId.TryGetValue(entryId, out existingByEntry))
                {
                    if (existingByEntry.IsFor(appointment))
                    {
                        return;
                    }

                    existingByEntry.Dispose();
                }
            }

            AppointmentSubscription existingByToken;
            if (_subscriptionByToken.TryGetValue(roomToken, out existingByToken))
            {
                if (existingByToken.IsFor(appointment))
                {
                    return;
                }

                existingByToken.Dispose();
            }

            var key = Guid.NewGuid().ToString("N");
            var subscription = new AppointmentSubscription(this, appointment, key, roomToken, lobbyEnabled, isEventConversation, entryId);
            _activeSubscriptions[key] = subscription;
            _subscriptionByToken[roomToken] = subscription;
            _knownRoomTokens.Add(roomToken);

            if (!string.IsNullOrEmpty(entryId))
            {
                _subscriptionByEntryId[entryId] = subscription;
            }
        }

        private void UnregisterSubscription(string key, string roomToken, string entryId)
        {
            LogTalk("Entferne Termin-Subscription (Token=" + (roomToken ?? "n/a") + ", EntryId=" + (entryId ?? "n/a") + ").");
            if (!string.IsNullOrEmpty(key))
            {
                _activeSubscriptions.Remove(key);
            }

            if (!string.IsNullOrEmpty(roomToken))
            {
                _knownRoomTokens.Remove(roomToken);
                _subscriptionByToken.Remove(roomToken);
            }

            if (!string.IsNullOrEmpty(entryId))
            {
                _subscriptionByEntryId.Remove(entryId);
            }
        }
        private bool TryDeleteRoom(string roomToken, bool isEventConversation)
        {
            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return true;
            }

            try
            {
                LogTalk("Loesche Raum (Token=" + roomToken + ", Event=" + isEventConversation + ").");
                var service = CreateTalkService();
                service.DeleteRoom(roomToken, isEventConversation);
                LogTalk("Raum erfolgreich geloescht (Token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Raum konnte nicht geloescht werden: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningRoomDeleteFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler beim Loeschen des Raums: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningRoomDeleteFailed, ex.Message));
            }

            return false;
        }

        private bool TryUpdateLobby(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            try
            {
                var service = CreateTalkService();
                DateTime start = appointment.Start;
                DateTime end = appointment.End;

                if (start == DateTime.MinValue)
                {
                    start = DateTime.Now;
                }

                if (end == DateTime.MinValue)
                {
                    end = start.AddHours(1);
                }

                LogTalk("Aktualisiere Lobby (Token=" + roomToken + ", Start=" + start.ToString("o") + ", Ende=" + end.ToString("o") + ", Event=" + isEventConversation + ").");
                service.UpdateLobby(roomToken, start, end, isEventConversation);
                LogTalk("Lobby erfolgreich aktualisiert (Token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Lobby konnte nicht aktualisiert werden: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler beim Aktualisieren der Lobby: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }

            return false;
        }

        private bool TryUpdateRoomDescription(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string description = BuildDescriptionPayload(appointment);
            if (description == null)
            {
                description = string.Empty;
            }

            try
            {
                var service = CreateTalkService();
                LogTalk("Aktualisiere Raumbeschreibung (Token=" + roomToken + ", Event=" + isEventConversation + ", TextLaenge=" + description.Length + ").");
                service.UpdateDescription(roomToken, description, isEventConversation);
                LogTalk("Raumbeschreibung aktualisiert (Token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Raumbeschreibung konnte nicht aktualisiert werden: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningDescriptionUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler beim Aktualisieren der Raumbeschreibung: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningDescriptionUpdateFailed, ex.Message));
            }

            return false;
        }

        private static long? ConvertToUnixTimestamp(DateTime? value)
        {
            if (!value.HasValue || value.Value == DateTime.MinValue)
            {
                return null;
            }

            var date = value.Value;
            if (date.Kind == DateTimeKind.Unspecified)
            {
                date = DateTime.SpecifyKind(date, DateTimeKind.Local);
            }

            var offset = new DateTimeOffset(date);
            return offset.ToUnixTimeSeconds();
        }

        private static string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("'", "&apos;")
                .Replace("\"", "&quot;");
        }

        private static void ShowWarning(string message)
        {
            MessageBox.Show(
                message,
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private sealed class WaitCursorScope : IDisposable
        {
            private readonly Cursor _previous;

            public WaitCursorScope()
            {
                _previous = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
            }

            public void Dispose()
            {
                Cursor.Current = _previous;
            }
        }

        private sealed class AppointmentSubscription : IDisposable
        {
            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.AppointmentItem _appointment;
            private readonly string _key;
            private readonly string _roomToken;
            private readonly bool _lobbyEnabled;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly bool _isEventConversation;
            private long? _lastLobbyTimer;
            private bool _roomDeleted;
            private bool _disposed;
            private string _entryId;

            internal AppointmentSubscription(
                NextcloudTalkAddIn owner,
                Outlook.AppointmentItem appointment,
                string key,
                string roomToken,
                bool lobbyEnabled,
                bool isEventConversation,
                string entryId)
            {
                _owner = owner;
                _appointment = appointment;
                _key = key;
                _roomToken = roomToken;
                _lobbyEnabled = lobbyEnabled;
                _isEventConversation = isEventConversation;
                _lastLobbyTimer = ConvertToUnixTimestamp(appointment != null ? (DateTime?)appointment.Start : null);
                _entryId = entryId;
                _events = appointment as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.BeforeDelete += OnBeforeDelete;
                    _events.Write += OnWrite;
                    _events.Close += OnClose;
                }

                LogTalk("Subscription registriert (Token=" + _roomToken + ", Lobby=" + _lobbyEnabled + ", Event=" + _isEventConversation + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }

            private void OnWrite(ref bool cancel)
            {
                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("OnWrite ignoriert (kein Token).");
                    return;
                }

                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnWrite ignoriert (nicht Organisator, Token=" + _roomToken + ").");
                    return;
                }

                LogTalk("OnWrite fuer Termin (Token=" + _roomToken + ").");

                if (_lobbyEnabled)
                {
                    long? current = ConvertToUnixTimestamp(_appointment != null ? (DateTime?)_appointment.Start : null);
                    if (current.HasValue && current != _lastLobbyTimer)
                    {
                        LogTalk("Versuche Lobby-Update waehrend OnWrite (Token=" + _roomToken + ").");
                        if (_owner.TryUpdateLobby(_appointment, _roomToken, _isEventConversation))
                        {
                            _lastLobbyTimer = current;
                            if (current.HasValue)
                            {
                                SetUserProperty(_appointment, PropertyStartEpoch, Outlook.OlUserPropertyType.olText, current.Value.ToString(CultureInfo.InvariantCulture));
                            }

                            LogTalk("Lobby-Update erfolgreich (Token=" + _roomToken + ").");
                        }
                        else
                        {
                            LogTalk("Lobby-Update fehlgeschlagen (Token=" + _roomToken + ").");
                        }
                    }
                }

                LogTalk("Aktualisiere Raumbeschreibung waehrend OnWrite (Token=" + _roomToken + ").");
                _owner.TryUpdateRoomDescription(_appointment, _roomToken, _isEventConversation);
                _owner.RefreshEntryBinding(this);
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("BeforeDelete ignoriert (nicht Organisator, Token=" + _roomToken + ").");
                    return;
                }

                LogTalk("BeforeDelete -> EnsureRoomDeleted (Token=" + _roomToken + ").");
                EnsureRoomDeleted();
            }

            private void OnClose(ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnClose ohne Organisator (Token=" + _roomToken + ").");
                    Dispose();
                    return;
                }

                if (!_roomDeleted && _appointment != null && !_appointment.Saved)
                {
                    LogTalk("OnClose ohne Speichern – Raum wird geloescht (Token=" + _roomToken + ").");
                    EnsureRoomDeleted();
                    _owner.ClearTalkProperties(_appointment);
                    Dispose();
                    return;
                }

                LogTalk("OnClose abgeschlossen (Token=" + _roomToken + ", Deleted=" + _roomDeleted + ").");
            }

            private void EnsureRoomDeleted()
            {
                if (_roomDeleted || !_owner.IsOrganizer(_appointment))
                {
                    if (_roomDeleted)
                    {
                        LogTalk("EnsureRoomDeleted: Raum bereits geloescht (Token=" + _roomToken + ").");
                    }

                    return;
                }

                if (_owner.TryDeleteRoom(_roomToken, _isEventConversation))
                {
                    _owner.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    LogTalk("EnsureRoomDeleted erfolgreich (Token=" + _roomToken + ").");
                    Dispose();
                }
                else
                {
                    LogTalk("EnsureRoomDeleted fehlgeschlagen (Token=" + _roomToken + ").");
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    LogTalk("Subscription.Dispose erneut aufgerufen (Token=" + _roomToken + ").");
                    return;
                }

                if (_events != null)
                {
                    _events.BeforeDelete -= OnBeforeDelete;
                    _events.Write -= OnWrite;
                    _events.Close -= OnClose;
                }

                _owner.UnregisterSubscription(_key, _roomToken, _entryId);
                _disposed = true;
                LogTalk("Subscription.Dispose abgeschlossen (Token=" + _roomToken + ").");
            }

            internal Outlook.AppointmentItem Appointment
            {
                get { return _appointment; }
            }

            internal string EntryId
            {
                get { return _entryId; }
            }

            internal bool IsFor(Outlook.AppointmentItem appointment)
            {
                return appointment != null && appointment == _appointment;
            }

            internal bool MatchesToken(string token)
            {
                return string.Equals(_roomToken, token, StringComparison.OrdinalIgnoreCase);
            }

            internal void UpdateEntryId(string entryId)
            {
                _entryId = entryId;
                LogTalk("Subscription EntryId aktualisiert (Token=" + _roomToken + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }
        }

        private sealed class ExplorerSubscription : IDisposable
        {
            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.Explorer _explorer;
            private readonly Outlook.ExplorerEvents_10_Event _events;
            private bool _disposed;

            internal ExplorerSubscription(NextcloudTalkAddIn owner, Outlook.Explorer explorer)
            {
                _owner = owner;
                _explorer = explorer;
                _events = explorer as Outlook.ExplorerEvents_10_Event;
                if (_events != null)
                {
                    _events.SelectionChange += OnSelectionChange;
                    _events.Close += OnClose;
                }
            }

            internal bool IsFor(Outlook.Explorer explorer)
            {
                return explorer != null && explorer == _explorer;
            }

            private void OnSelectionChange()
            {
                if (_disposed || _explorer == null)
                {
                    return;
                }

                Outlook.Selection selection = null;
                try
                {
                    selection = _explorer.Selection;
                    if (selection == null)
                    {
                        return;
                    }

                    int count = selection.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        try
                        {
                            var appointment = selection[index] as Outlook.AppointmentItem;
                            if (appointment != null)
                            {
                                _owner.EnsureSubscriptionForAppointment(appointment);
                            }
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    if (selection != null && Marshal.IsComObject(selection))
                    {
                        try
                        {
                            Marshal.ReleaseComObject(selection);
                        }
                        catch
                        {
                        }
                    }
                }
            }

            private void OnClose()
            {
                Dispose();
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                if (_events != null)
                {
                    try
                    {
                        _events.SelectionChange -= OnSelectionChange;
                        _events.Close -= OnClose;
                    }
                    catch
                    {
                    }
                }

                if (_explorer != null && Marshal.IsComObject(_explorer))
                {
                    try
                    {
                        Marshal.ReleaseComObject(_explorer);
                    }
                    catch
                    {
                    }
                }

                _owner.RemoveExplorerSubscription(this);
                _disposed = true;
            }
        }

        private void EnsureSettingsLoaded()
        {
            if (_currentSettings == null)
            {
                if (_settingsStorage != null)
                {
                    _currentSettings = _settingsStorage.Load();
                }

                if (_currentSettings == null)
                {
                    _currentSettings = new AddinSettings();
                }

                EnsureItemSendHook();
                EnsureOutboxHook();
                EnsureItemLoadHook();
                EnsureInspectorHook();
                EnsureExplorerHook();
                EnsureExistingAppointmentsSubscribed();
            }
        }

        private static class PictureConverter
        {
            /**
             * Interne Hilfsklasse, die Image -> IPictureDisp Konvertierung kapselt.
             */
            private sealed class AxHostPictureConverter : AxHost
            {
                public AxHostPictureConverter() : base(string.Empty)
                {
                }

                /**
                 * Konvertiert ein Image unter Verwendung der WinForms-Infrastruktur.
                 */
                public static stdole.IPictureDisp ImageToPictureDisp(Image image)
                {
                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
                }
            }

            /**
             * Stellt die Konvertierung fuer aufrufende Stellen bereit.
             */
            public static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return AxHostPictureConverter.ImageToPictureDisp(image);
            }
        }

        /**
         * Registriert einen Items-Listener auf dem Postausgang, damit blockierte Besprechungs-Mails
         * unmittelbar entfernt werden und Outlook den Versand trotzdem als abgeschlossen betrachtet.
         */
        private void EnsureOutboxHook()
        {
            if (_outlookApplication == null)
            {
                return;
            }

            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;

            try
            {
                session = _outlookApplication.Session;
                if (session == null)
                {
                    DiagnosticsLogger.Log("Muzzle", "EnsureOutboxHook: session unavailable.");
                    return;
                }

                var blockedStoreIds = GetBlockedStoreIds(session);
                CleanupOutboxSubscriptions(blockedStoreIds);

                if (blockedStoreIds.Count == 0)
                {
                    DiagnosticsLogger.Log("Muzzle", "EnsureOutboxHook: kein Konto fuer den Maulkorb aktiviert – keine Outbox-Listener noetig.");
                    return;
                }

                stores = session.Stores;
                if (stores == null || stores.Count == 0)
                {
                    DiagnosticsLogger.Log("Muzzle", "EnsureOutboxHook: no stores available.");
                    return;
                }

                foreach (Outlook.Store store in stores)
                {
                    if (store == null)
                    {
                        continue;
                    }

                    string storeId = SafeGetString(() => store.StoreID) ?? string.Empty;
                    if (!blockedStoreIds.Contains(storeId))
                    {
                        continue;
                    }

                    if (IsOutboxSubscriptionActive(storeId))
                    {
                        continue;
                    }

                    AttachOutboxForStore(store, storeId);
                }

                if (_outboxSubscriptions.Count == 0)
                {
                    DiagnosticsLogger.Log("Muzzle", "EnsureOutboxHook: keine Outbox-Subscriptions konnten erstellt werden.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "EnsureOutboxHook failed: " + ex.Message);
            }
            finally
            {
                ReleaseComReference(stores);
                ReleaseComReference(session);
            }
        }

        private HashSet<string> GetBlockedStoreIds(Outlook.NameSpace session)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (session == null || _currentSettings == null)
            {
                return result;
            }

            Outlook.Accounts accounts = null;
            try
            {
                accounts = session.Accounts;
                if (accounts == null || accounts.Count == 0)
                {
                    return result;
                }

                bool globalEnabled = _currentSettings.OutlookMuzzleEnabled;
                bool hasPerAccount = _currentSettings.OutlookMuzzleAccounts != null && _currentSettings.OutlookMuzzleAccounts.Count > 0;

                foreach (Outlook.Account account in accounts)
                {
                    if (account == null)
                    {
                        continue;
                    }

                    Outlook.Store accountStore = null;
                    string storeId = string.Empty;
                    try
                    {
                        accountStore = account.DeliveryStore;
                        if (accountStore != null)
                        {
                            storeId = SafeGetString(() => accountStore.StoreID);
                        }
                    }
                    catch
                    {
                        storeId = string.Empty;
                    }
                    finally
                    {
                        ReleaseComReference(accountStore);
                    }

                    if (string.IsNullOrEmpty(storeId))
                    {
                        continue;
                    }

                    string normalized = OutlookAccountHelper.NormalizeAccountIdentifier(account);
                    bool enabled = _currentSettings.IsMuzzleEnabledForAccount(normalized);

                    if (!hasPerAccount && globalEnabled)
                    {
                        result.Add(storeId);
                        continue;
                    }

                    if (enabled)
                    {
                        result.Add(storeId);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "GetBlockedStoreIds failed: " + ex.Message);
            }
            finally
            {
                ReleaseComReference(accounts);
            }

            return result;
        }

        private void CleanupOutboxSubscriptions(HashSet<string> activeStoreIds)
        {
            if (_outboxSubscriptions.Count == 0)
            {
                return;
            }

            var stillActive = new List<OutboxSubscription>(_outboxSubscriptions.Count);
            foreach (var subscription in _outboxSubscriptions)
            {
                if (subscription == null)
                {
                    continue;
                }

                if (activeStoreIds.Contains(subscription.StoreId))
                {
                    stillActive.Add(subscription);
                }
                else
                {
                    subscription.Dispose(OnOutboxItemAdd);
                }
            }

            _outboxSubscriptions.Clear();
            _outboxSubscriptions.AddRange(stillActive);
        }

        private void AttachOutboxForStore(Outlook.Store store, string storeId)
        {
            Outlook.MAPIFolder outbox = null;
            Outlook.Items items = null;
            string storeName = SafeGetString(() => store.DisplayName);
            string storeLabel = string.IsNullOrEmpty(storeName) ? (string.IsNullOrEmpty(storeId) ? "<unknown store>" : storeId) : storeName;

            try
            {
                outbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
            }
            catch
            {
                outbox = null;
            }

            if (outbox == null)
            {
                DiagnosticsLogger.Log("Muzzle", "Outbox hook skipped (Store=" + storeLabel + "): kein Outbox-Ordner.");
                return;
            }

            try
            {
                items = outbox.Items;
            }
            catch
            {
                items = null;
            }

            if (items == null)
            {
                DiagnosticsLogger.Log("Muzzle", "Outbox hook skipped (Store=" + storeLabel + "): Items-Auflistung fehlt.");
                ReleaseComReference(outbox);
                return;
            }

            try
            {
                items.ItemAdd += OnOutboxItemAdd;
                _outboxSubscriptions.Add(new OutboxSubscription(storeId, storeLabel, items));
                DiagnosticsLogger.Log("Muzzle", "Outbox hook attached (Store=" + storeLabel + ").");
                items = null; // Ownership transfer to subscription
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "Outbox hook failed (Store=" + storeLabel + "): " + ex.Message);
                try
                {
                    items.ItemAdd -= OnOutboxItemAdd;
                }
                catch
                {
                }
                ReleaseComReference(items);
            }
            finally
            {
                ReleaseComReference(outbox);
            }
        }

        private bool IsOutboxSubscriptionActive(string storeId)
        {
            foreach (var subscription in _outboxSubscriptions)
            {
                if (subscription.IsForStore(storeId))
                {
                    return true;
                }
            }

            return false;
        }

        /**
         * Loest den Items-Listener des Postausgangs und gibt COM-Ressourcen frei.
         */
        private void UnhookOutbox()
        {
            if (_outboxSubscriptions.Count == 0)
            {
                return;
            }

            ReleaseOutboxItems();
        }

        /**
         * Hilfsmethode zum Freigeben des gehaltenen Items-Objektes.
         */
        private void ReleaseOutboxItems()
        {
            if (_outboxSubscriptions.Count == 0)
            {
                return;
            }

            foreach (var subscription in _outboxSubscriptions)
            {
                subscription.Dispose(OnOutboxItemAdd);
            }

            _outboxSubscriptions.Clear();
        }

        private void EnsureInspectorHook()
        {
            if (_outlookApplication == null || _inspectors != null)
            {
                return;
            }

            try
            {
                _inspectors = _outlookApplication.Inspectors;
                if (_inspectors != null)
                {
                    _inspectors.NewInspector += OnNewInspector;
                }
            }
            catch
            {
                _inspectors = null;
            }
        }

        private void UnhookInspector()
        {
            if (_inspectors == null)
            {
                return;
            }

            try
            {
                _inspectors.NewInspector -= OnNewInspector;
            }
            catch
            {
            }
            finally
            {
                Marshal.FinalReleaseComObject(_inspectors);
                _inspectors = null;
            }
        }

        private void EnsureExplorerHook()
        {
            if (_outlookApplication == null)
            {
                return;
            }

            if (_explorers == null)
            {
                try
                {
                    _explorers = _outlookApplication.Explorers;
                    if (_explorers != null)
                    {
                        _explorers.NewExplorer += OnNewExplorer;
                    }
                }
                catch
                {
                    _explorers = null;
                }
            }

            try
            {
                HookExplorer(_outlookApplication.ActiveExplorer());
            }
            catch
            {
            }
        }

        private void UnhookExplorer()
        {
            foreach (var subscription in _explorerSubscriptions.ToArray())
            {
                subscription.Dispose();
            }

            _explorerSubscriptions.Clear();

            if (_explorers != null)
            {
                try
                {
                    _explorers.NewExplorer -= OnNewExplorer;
                }
                catch
                {
                }
                finally
                {
                    Marshal.FinalReleaseComObject(_explorers);
                    _explorers = null;
                }
            }
        }

        private void EnsureExistingAppointmentsSubscribed()
        {
            if (_outlookApplication == null || _initialScanPerformed)
            {
                return;
            }

            Outlook.MAPIFolder calendar = null;
            Outlook.Items items = null;

            try
            {
                var session = _outlookApplication.Session;
                if (session == null)
                {
                    return;
                }

                try
                {
                    calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                }
                catch
                {
                    calendar = null;
                }

                if (calendar == null)
                {
                    return;
                }

                try
                {
                    items = calendar.Items;
                    if (items == null)
                    {
                        return;
                    }

                    items.IncludeRecurrences = true;
                    object raw = items.GetFirst();
                    while (raw != null)
                    {
                        object current = raw;
                        try
                        {
                            var appointment = current as Outlook.AppointmentItem;
                            if (appointment != null)
                            {
                                EnsureSubscriptionForAppointment(appointment);
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            if (current != null && Marshal.IsComObject(current))
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(current);
                                }
                                catch
                                {
                                }
                            }
                        }

                        raw = items.GetNext();
                    }
                }
                catch
                {
                }
            }
            finally
            {
                if (items != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(items);
                    }
                    catch
                    {
                    }
                }

                if (calendar != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(calendar);
                    }
                    catch
                    {
                    }
                }

                _initialScanPerformed = true;
            }
        }

        private void EnsureItemLoadHook()
        {
            if (_outlookApplication == null || _itemLoadHooked)
            {
                return;
            }

            try
            {
                _outlookApplication.ItemLoad += OnApplicationItemLoad;
                _itemLoadHooked = true;
            }
            catch
            {
                _itemLoadHooked = false;
            }
        }

        private void UnhookItemLoad()
        {
            if (_outlookApplication == null || !_itemLoadHooked)
            {
                return;
            }

            try
            {
                _outlookApplication.ItemLoad -= OnApplicationItemLoad;
            }
            catch
            {
            }
            finally
            {
                _itemLoadHooked = false;
            }
        }

        private void EnsureItemSendHook()
        {
            if (_outlookApplication == null || _itemSendHooked)
            {
                return;
            }

            try
            {
                _outlookApplication.ItemSend += OnItemSend;
                _itemSendHooked = true;
            }
            catch
            {
                _itemSendHooked = false;
            }
        }

        private void UnhookItemSend()
        {
            if (_outlookApplication == null || !_itemSendHooked)
            {
                return;
            }

            try
            {
                _outlookApplication.ItemSend -= OnItemSend;
            }
            catch
            {
            }
            finally
            {
                _itemSendHooked = false;
            }
        }

        private void OnItemSend(object item, ref bool cancel)
        {
            if (!IsMuzzleFeatureActive())
            {
                return;
            }

            string messageClass = TryGetMessageClass(item);
            if (!ShouldTrackMuzzleForSend(item, messageClass))
            {
                DiagnosticsLogger.Log("Muzzle", "OnItemSend ignored (MessageClass=" + (messageClass ?? "<unknown>") + ").");
                return;
            }

            DiagnosticsLogger.Log("Muzzle", "OnItemSend triggered.");
            EnsureOutboxHook();

            string accountKey = GetAccountKeyForItem(item);
            if (string.IsNullOrEmpty(accountKey))
            {
                DiagnosticsLogger.Log("Muzzle", "Could not resolve account for ItemSend.");
                return;
            }

            if (_currentSettings == null || !_currentSettings.IsMuzzleEnabledForAccount(accountKey))
            {
                DiagnosticsLogger.Log("Muzzle", "Account '" + accountKey + "' nicht fuer den Maulkorb aktiviert - ItemSend wird nicht blockiert.");
                return;
            }

            string stampedAccount = StampMuzzleAccount(item, accountKey);
            if (string.IsNullOrEmpty(stampedAccount))
            {
                DiagnosticsLogger.Log("Muzzle", "Could not stamp muzzle account (resolution failed).");
                return;
            }

            string entryId = CaptureEntryIdForMuzzle(item);
            DiagnosticsLogger.Log("Muzzle", "Stamped muzzle account '" + stampedAccount + "' onto pending item (EntryID=" + (string.IsNullOrEmpty(entryId) ? "<empty>" : entryId) + ").");
            EnqueuePendingMuzzleItem(stampedAccount, entryId);
        }

        /**
         * Wird aus dem Outbox-Listener aufgerufen, um blockierte Items unmittelbar zu entfernen.
         */
        private void OnOutboxItemAdd(object item)
        {
            if (item == null)
            {
                return;
            }

            try
            {
                DiagnosticsLogger.Log("Muzzle", "OnOutboxItemAdd (" + item.GetType().Name + ").");

                if (!IsMuzzleFeatureActive())
                {
                    DiagnosticsLogger.Log("Muzzle", "Muzzle inactive -> skipping.");
                    return;
                }

                var messageClass = TryGetMessageClass(item);
                if (!IsMuzzleTarget(messageClass))
                {
                    DiagnosticsLogger.Log("Muzzle", "MessageClass " + (messageClass ?? "<empty>") + " not targeted.");
                    return;
                }

                MuzzlePendingInfo pending = DequeuePendingMuzzleItem();
                string queuedAccount = pending != null ? pending.AccountKey : string.Empty;
                string queuedEntryId = pending != null ? pending.EntryId : string.Empty;
                if (!string.IsNullOrEmpty(queuedAccount))
                {
                    DiagnosticsLogger.Log("Muzzle", "Using queued account '" + queuedAccount + "' (EntryID=" + (string.IsNullOrEmpty(queuedEntryId) ? "<empty>" : queuedEntryId) + ").");
                }
                else
                {
                    DiagnosticsLogger.Log("Muzzle", "No queued account available for Outbox item.");
                }

                if (!ShouldMuzzleItem(item, queuedAccount))
                {
                    DiagnosticsLogger.Log("Muzzle", "Account not configured for blocking.");
                    return;
                }

                if (TryDeleteMuzzleItem(item, queuedEntryId))
                {
                    DiagnosticsLogger.Log("Muzzle", "Item deleted from Outbox.");
                }
                else
                {
                    DiagnosticsLogger.Log("Muzzle", "Failed to delete Outbox item – invite may be sent.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "OnOutboxItemAdd error: " + ex.Message);
            }
            finally
            {
                ReleaseComReference(item);
            }
        }

        private void OnApplicationItemLoad(object item)
        {
            if (item == null)
            {
                return;
            }

            var appointment = item as Outlook.AppointmentItem;
            if (appointment == null)
            {
                return;
            }

            EnsureSubscriptionForAppointment(appointment);
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {
            if (inspector == null)
            {
                return;
            }

            try
            {
                var appointment = inspector.CurrentItem as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    EnsureSubscriptionForAppointment(appointment);
                }
            }
            catch
            {
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return;
            }

            try
            {
                HookExplorer(explorer);
            }
            catch
            {
            }
        }

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment)
        {
            if (appointment == null)
            {
                return;
            }

            try
            {
                string roomToken = GetUserPropertyText(appointment, PropertyToken);
                if (string.IsNullOrWhiteSpace(roomToken))
                {
                    return;
                }

                string entryId = GetEntryId(appointment);
                if (!string.IsNullOrEmpty(entryId))
                {
                    AppointmentSubscription existingByEntry;
                    if (_subscriptionByEntryId.TryGetValue(entryId, out existingByEntry))
                    {
                        if (existingByEntry.IsFor(appointment) && existingByEntry.MatchesToken(roomToken))
                        {
                            return;
                        }

                        existingByEntry.Dispose();
                    }
                }

                AppointmentSubscription existingByToken;
                if (_subscriptionByToken.TryGetValue(roomToken, out existingByToken))
                {
                    if (existingByToken.IsFor(appointment))
                    {
                        return;
                    }

                    existingByToken.Dispose();
                }

                var roomType = GetRoomType(appointment);
                bool isEventConversation = roomType.HasValue && roomType.Value == TalkRoomType.EventConversation;
                bool lobbyEnabled = GetUserPropertyBool(appointment, PropertyLobby);

                RegisterSubscription(appointment, roomToken, lobbyEnabled, isEventConversation);
            }
            catch
            {
            }
        }

        private void HookExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return;
            }

            foreach (var existing in _explorerSubscriptions)
            {
                if (existing.IsFor(explorer))
                {
                    return;
                }
            }

            var subscription = new ExplorerSubscription(this, explorer);
            _explorerSubscriptions.Add(subscription);
        }

        private void RemoveExplorerSubscription(ExplorerSubscription subscription)
        {
            if (subscription == null)
            {
                return;
            }

            _explorerSubscriptions.Remove(subscription);
        }

        /**
         * Ermittelt die MessageClass eines Outlook-Items.
         */
        private static string TryGetMessageClass(object item)
        {
            try
            {
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    return mail.MessageClass;
                }

                Outlook.MeetingItem meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    return meeting.MessageClass;
                }

                Outlook.AppointmentItem appointment = item as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    return appointment.MessageClass;
                }

                Outlook._MailItem mailInterface = item as Outlook._MailItem;
                if (mailInterface != null)
                {
                    return mailInterface.MessageClass;
                }

                Outlook._MeetingItem meetingInterface = item as Outlook._MeetingItem;
                if (meetingInterface != null)
                {
                    return meetingInterface.MessageClass;
                }

                Outlook._AppointmentItem appointmentInterface = item as Outlook._AppointmentItem;
                if (appointmentInterface != null)
                {
                    return appointmentInterface.MessageClass;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "TryGetSendUsingAccount exception: " + ex.Message);
            }

            return null;
        }

        /**
         * Liest die EntryID eines Items (falls vorhanden) aus.
         */
        private static string TryGetEntryId(object item)
        {
            try
            {
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    return mail.EntryID ?? string.Empty;
                }

                Outlook.MeetingItem meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    return meeting.EntryID ?? string.Empty;
                }

                Outlook.AppointmentItem appointment = item as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    return appointment.EntryID ?? string.Empty;
                }

                Outlook._MailItem mailInterface = item as Outlook._MailItem;
                if (mailInterface != null)
                {
                    return mailInterface.EntryID ?? string.Empty;
                }

                Outlook._MeetingItem meetingInterface = item as Outlook._MeetingItem;
                if (meetingInterface != null)
                {
                    return meetingInterface.EntryID ?? string.Empty;
                }

                Outlook._AppointmentItem appointmentInterface = item as Outlook._AppointmentItem;
                if (appointmentInterface != null)
                {
                    return appointmentInterface.EntryID ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "TryGetEntryId failed: " + ex.Message);
            }

            return string.Empty;
        }

        /**
         * Versucht eine stabile EntryID fuer den Maulkorb zu erfassen. Falls sie noch fehlt, wird das Item zunaechst gespeichert.
         */
        private string CaptureEntryIdForMuzzle(object item)
        {
            string entryId = TryGetEntryId(item);
            if (!string.IsNullOrEmpty(entryId))
            {
                return entryId;
            }

            if (TryPersistItemForEntryId(item))
            {
                entryId = TryGetEntryId(item);
                if (!string.IsNullOrEmpty(entryId))
                {
                    DiagnosticsLogger.Log("Muzzle", "EntryID nach Save ermittelt: " + entryId);
                    return entryId;
                }
            }

            DiagnosticsLogger.Log("Muzzle", "EntryID bleibt leer (Item konnte nicht gespeichert werden).");
            return string.Empty;
            }

        private static bool TryPersistItemForEntryId(object item)
        {
            try
            {
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    mail.Save();
                    DiagnosticsLogger.Log("Muzzle", "MailItem fuer EntryID gespeichert.");
                    return true;
                }

                Outlook.MeetingItem meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    meeting.Save();
                    DiagnosticsLogger.Log("Muzzle", "MeetingItem fuer EntryID gespeichert.");
                    return true;
                }

                Outlook.AppointmentItem appointment = item as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    appointment.Save();
                    DiagnosticsLogger.Log("Muzzle", "AppointmentItem fuer EntryID gespeichert.");
                    return true;
                }

                Outlook._MailItem mailInterface = item as Outlook._MailItem;
                if (mailInterface != null)
                {
                    mailInterface.Save();
                    DiagnosticsLogger.Log("Muzzle", "_MailItem fuer EntryID gespeichert.");
                    return true;
                }

                Outlook._MeetingItem meetingInterface = item as Outlook._MeetingItem;
                if (meetingInterface != null)
                {
                    meetingInterface.Save();
                    DiagnosticsLogger.Log("Muzzle", "_MeetingItem fuer EntryID gespeichert.");
                    return true;
                }

                Outlook._AppointmentItem appointmentInterface = item as Outlook._AppointmentItem;
                if (appointmentInterface != null)
                {
                    appointmentInterface.Save();
                    DiagnosticsLogger.Log("Muzzle", "_AppointmentItem fuer EntryID gespeichert.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "EntryID-Save fehlgeschlagen: " + ex.Message);
            }

            return false;
        }

        private Outlook.MailItem GetActiveMailItem()
        {
            if (_outlookApplication == null)
            {
                return null;
            }

            try
            {
                Outlook.Inspector inspector = _outlookApplication.ActiveInspector();
                if (inspector == null)
                {
                    return null;
                }

                return inspector.CurrentItem as Outlook.MailItem;
            }
            catch
            {
                return null;
            }
        }

        private static void InsertHtmlIntoMail(Outlook.MailItem mail, string htmlFragment)
        {
            if (mail == null || string.IsNullOrEmpty(htmlFragment))
            {
                return;
            }

            const string topSpacing = "<p style=\"margin:0;\">&nbsp;</p><p style=\"margin:0;\">&nbsp;</p>";
            const string bottomSpacing = "<p style=\"margin:0;\">&nbsp;</p>";
            string fragmentWithSpacing = topSpacing + htmlFragment + bottomSpacing;

            string body = mail.HTMLBody ?? string.Empty;
            if (string.IsNullOrWhiteSpace(body))
            {
                mail.HTMLBody = "<html><body>" + fragmentWithSpacing + "</body></html>";
                return;
            }

            int bodyTagIndex = body.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (bodyTagIndex >= 0)
            {
                int insertIndex = body.IndexOf('>', bodyTagIndex);
                if (insertIndex >= 0)
                {
                    insertIndex += 1;
                    mail.HTMLBody = body.Insert(insertIndex, fragmentWithSpacing);
                    return;
                }
            }

            mail.HTMLBody = fragmentWithSpacing + body;
        }

        private bool IsMuzzleFeatureActive()
        {
            return _currentSettings != null && _currentSettings.HasAnyMuzzleAccountEnabled();
        }

        private static bool ShouldTrackMuzzleForSend(object item, string messageClass)
        {
            if (item is Outlook.AppointmentItem || item is Outlook._AppointmentItem)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(messageClass))
            {
                return false;
            }

            string normalized = messageClass.Trim().ToUpperInvariant();
            if (normalized == "IPM.APPOINTMENT")
            {
                return true;
            }

            return IsMuzzleTarget(normalized);
        }

        private bool ShouldMuzzleItem(object item, string queuedAccountKey = null)
        {
            if (!IsMuzzleFeatureActive())
            {
                return false;
            }

            string accountKey = !string.IsNullOrEmpty(queuedAccountKey) ? queuedAccountKey : TryGetStampedAccount(item);
            if (string.IsNullOrEmpty(accountKey))
            {
                accountKey = GetAccountKeyForItem(item);
            }

            bool enabled = _currentSettings.IsMuzzleEnabledForAccount(accountKey);
            DiagnosticsLogger.Log("Muzzle", "Outbox item inspected (Account='" + (accountKey ?? "<empty>") + "', Enabled=" + enabled + ").");
            return enabled;
        }

        private string GetAccountKeyForItem(object item)
        {
            Outlook.Account account = null;
            try
            {
                account = TryGetSendUsingAccount(item);
                if (account != null)
                {
                    string normalized = OutlookAccountHelper.NormalizeAccountIdentifier(account);
                    DiagnosticsLogger.Log("Muzzle", "SendUsingAccount resolved: " + normalized);
                    return normalized;
                }

                string sender = TryGetSenderAddress(item);
                if (!string.IsNullOrWhiteSpace(sender))
                {
                    string normalized = OutlookAccountHelper.NormalizeAccountIdentifier(sender);
                    DiagnosticsLogger.Log("Muzzle", "Sender address fallback: " + normalized);
                    return normalized;
                }

                DiagnosticsLogger.Log("Muzzle", "Account resolution failed – empty key.");
                return string.Empty;
            }
            finally
            {
                ReleaseComReference(account);
            }
        }

        private string StampMuzzleAccount(object item, string accountKey)
        {
            if (string.IsNullOrEmpty(accountKey))
            {
                return string.Empty;
            }

            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = TryGetUserProperties(item);
                if (properties == null)
                {
                    DiagnosticsLogger.Log("Muzzle", "StampMuzzleAccount: UserProperties unavailable.");
                    return accountKey;
                }

                property = properties.Find(MuzzleAccountProperty);
                if (property == null)
                {
                    property = properties.Add(MuzzleAccountProperty, Outlook.OlUserPropertyType.olText);
                }
                property.Value = accountKey;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "StampMuzzleAccount failed: " + ex.Message);
            }
            finally
            {
                ReleaseComReference(property);
                ReleaseComReference(properties);
            }

            return accountKey;
        }

        private void EnqueuePendingMuzzleItem(string accountKey, string entryId)
        {
            if (string.IsNullOrWhiteSpace(accountKey) && string.IsNullOrWhiteSpace(entryId))
            {
                DiagnosticsLogger.Log("Muzzle", "Pending queue not updated (empty account + EntryID).");
                return;
            }

            lock (_pendingMuzzleLock)
            {
                _pendingMuzzleItems.Enqueue(new MuzzlePendingInfo(accountKey, entryId));
                const int maxPending = 10;
                while (_pendingMuzzleItems.Count > maxPending)
                {
                    MuzzlePendingInfo removed = _pendingMuzzleItems.Dequeue();
                    string removedAccount = removed != null ? removed.AccountKey : string.Empty;
                    string removedEntry = removed != null ? removed.EntryId : string.Empty;
                    DiagnosticsLogger.Log("Muzzle", "Pending queue trimmed, discarded '" + (string.IsNullOrEmpty(removedAccount) ? "<empty>" : removedAccount) + "' (EntryID=" + (string.IsNullOrEmpty(removedEntry) ? "<empty>" : removedEntry) + ").");
                }
            }
        }

        private MuzzlePendingInfo DequeuePendingMuzzleItem()
        {
            lock (_pendingMuzzleLock)
            {
                while (_pendingMuzzleItems.Count > 0)
                {
                    MuzzlePendingInfo candidate = _pendingMuzzleItems.Dequeue();
                    if (candidate != null && (!string.IsNullOrWhiteSpace(candidate.AccountKey) || !string.IsNullOrWhiteSpace(candidate.EntryId)))
                    {
                        return candidate;
                    }

                    DiagnosticsLogger.Log("Muzzle", "Discarded empty pending entry (Account/EntryID fehlten).");
                }
            }

            return null;
        }

        private string TryGetStampedAccount(object item)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = TryGetUserProperties(item);
                if (properties == null)
                {
                    return string.Empty;
                }

                property = properties.Find(MuzzleAccountProperty);
                if (property == null)
                {
                    return string.Empty;
                }

                var value = property.Value as string;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    DiagnosticsLogger.Log("Muzzle", "Stamped account restored: " + value.Trim());
                    return value.Trim();
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "TryGetStampedAccount failed: " + ex.Message);
                return string.Empty;
            }
            finally
            {
                ReleaseComReference(property);
                ReleaseComReference(properties);
            }
        }

        private static Outlook.UserProperties TryGetUserProperties(object item)
        {
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    return mail.UserProperties;
                }

                var meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    return meeting.UserProperties;
                }

                var mailInterface = item as Outlook._MailItem;
                if (mailInterface != null)
                {
                    return mailInterface.UserProperties;
                }

                var meetingInterface = item as Outlook._MeetingItem;
                if (meetingInterface != null)
                {
                    return meetingInterface.UserProperties;
                }
            }
            catch
            {
            }

            return null;
        }

        private static Outlook.Account TryGetSendUsingAccount(object item)
        {
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    return mail.SendUsingAccount;
                }

                var mailInterface = item as Outlook._MailItem;
                if (mailInterface != null)
                {
                    return mailInterface.SendUsingAccount;
                }

                var meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    return meeting.SendUsingAccount;
                }

                var meetingInterface = item as Outlook._MeetingItem;
                if (meetingInterface != null)
                {
                    return meetingInterface.SendUsingAccount;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "TryGetSendUsingAccount exception: " + ex.Message);
            }

            return null;
        }

        private static string TryGetSenderAddress(object item)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                var mail = item as Outlook.MailItem;
                var meeting = item as Outlook.MeetingItem;
                var mailInterface = item as Outlook._MailItem;
                var meetingInterface = item as Outlook._MeetingItem;
                if (mail != null)
                {
                    accessor = mail.PropertyAccessor;
                }
                else if (meeting != null)
                {
                    accessor = meeting.PropertyAccessor;
                }
                else if (mailInterface != null)
                {
                    accessor = mailInterface.PropertyAccessor;
                }
                else if (meetingInterface != null)
                {
                    accessor = meetingInterface.PropertyAccessor;
                }

                if (accessor == null)
                {
                    return string.Empty;
                }

                string address = SafeGetAddressFromAccessor(accessor, SenderAddressProperty);
                if (string.IsNullOrWhiteSpace(address))
                {
                    address = SafeGetAddressFromAccessor(accessor, "http://schemas.microsoft.com/mapi/proptag/0x0065001E");
                }

                if (string.IsNullOrWhiteSpace(address) && mail != null)
                {
                    address = SafeGetString(() => mail.SenderEmailAddress);
                }
                else if (string.IsNullOrWhiteSpace(address) && mailInterface != null)
                {
                    address = SafeGetString(() => mailInterface.SenderEmailAddress);
                }

                if (string.IsNullOrWhiteSpace(address) && meeting != null)
                {
                    address = SafeGetString(() => meeting.SenderEmailAddress);
                }
                else if (string.IsNullOrWhiteSpace(address) && meetingInterface != null)
                {
                    address = SafeGetString(() => meetingInterface.SenderEmailAddress);
                }

                if (string.IsNullOrWhiteSpace(address))
                {
                    return string.Empty;
                }

                return address.Trim();
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComReference(accessor);
            }
        }

        /**
         * Prueft, ob ein Item vom Maulkorb betroffen ist.
         */
        private static bool IsMuzzleTarget(string messageClass)
        {
            if (string.IsNullOrWhiteSpace(messageClass))
            {
                return false;
            }

            string normalized = messageClass.Trim().ToUpperInvariant();
            return normalized == "IPM.SCHEDULE.MEETING.REQUEST" ||
                   normalized == "IPM.SCHEDULE.MEETING.UPDATE" ||
                   normalized == "IPM.SCHEDULE.MEETING.CANCELED" ||
                   normalized == "IPM.SCHEDULE.MEETING.CANCELLED";
        }

        /**
         * Entfernt das gekennzeichnete Item endgueltig aus dem Postausgang (direkt oder via EntryID-Fallback).
         */
        private bool TryDeleteMuzzleItem(object item, string fallbackEntryId = null)
        {
            if (item == null)
            {
                return false;
            }

            if (TryDeleteDirect(item))
            {
                return true;
            }

            string entryId = TryGetEntryId(item);
            if (string.IsNullOrEmpty(entryId) && !string.IsNullOrEmpty(fallbackEntryId))
            {
                entryId = fallbackEntryId;
                DiagnosticsLogger.Log("Muzzle", "DeleteItem: nutze gespeicherte EntryID " + entryId + ".");
            }
            if (string.IsNullOrEmpty(entryId) || _outlookApplication == null)
            {
                DiagnosticsLogger.Log("Muzzle", "DeleteItem: kein EntryID-Fallback moeglich.");
                return false;
            }

            Outlook.NameSpace session = null;
            object fetched = null;
            try
            {
                session = _outlookApplication.Session;
                if (session == null)
                {
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: Outlook-Session nicht verfuegbar.");
                    return false;
                }

                fetched = session.GetItemFromID(entryId);
                if (fetched == null)
                {
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: EntryID " + entryId + " nicht gefunden.");
                    return false;
                }

                if (TryDeleteDirect(fetched))
                {
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem ueber EntryID erfolgreich.");
                    return true;
                }

                DiagnosticsLogger.Log("Muzzle", "DeleteItem ueber EntryID ohne Wirkung.");
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "DeleteItem via EntryID fehlgeschlagen: " + ex.Message);
                return false;
            }
            finally
            {
                ReleaseComReference(fetched);
                ReleaseComReference(session);
            }
        }

        private static bool TryDeleteDirect(object item)
        {
            try
            {
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    mail.Delete();
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: MailItem deleted.");
                    return true;
                }

                Outlook.MeetingItem meeting = item as Outlook.MeetingItem;
                if (meeting != null)
                {
                    meeting.Delete();
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: MeetingItem deleted.");
                    return true;
                }

                Outlook._MailItem mailInterfaceOnly = item as Outlook._MailItem;
                if (mailInterfaceOnly != null)
                {
                    mailInterfaceOnly.Delete();
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: _MailItem deleted.");
                    return true;
                }

                Outlook._MeetingItem meetingInterfaceOnly = item as Outlook._MeetingItem;
                if (meetingInterfaceOnly != null)
                {
                    meetingInterfaceOnly.Delete();
                    DiagnosticsLogger.Log("Muzzle", "DeleteItem: _MeetingItem deleted.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Muzzle", "DeleteItem direct failed: " + ex.Message);
            }

            return false;
        }
        private static void ReleaseComReference(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
            }
        }

        private static string SafeGetAddressFromAccessor(Outlook.PropertyAccessor accessor, string uri)
        {
            if (accessor == null || string.IsNullOrEmpty(uri))
            {
                return string.Empty;
            }

            try
            {
                var result = accessor.GetProperty(uri);
                return result as string ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetString(Func<string> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return string.Empty;
            }
        }

        /**
         * Puffer fuer OnItemSend -> OnOutboxItemAdd, haelt Account + EntryID.
         */
        private sealed class MuzzlePendingInfo
        {
            public string AccountKey { get; private set; }
            public string EntryId { get; private set; }

            public MuzzlePendingInfo(string accountKey, string entryId)
            {
                AccountKey = accountKey ?? string.Empty;
                EntryId = entryId ?? string.Empty;
            }
        }

        /**
         * Haelt eine Outbox-Subscription pro Store, damit ItemAdd pro Postausgang aktiv bleibt.
         */
        private sealed class OutboxSubscription
        {
            public string StoreId { get; private set; }
            public string StoreLabel { get; private set; }
            public Outlook.Items Items { get; private set; }

            public OutboxSubscription(string storeId, string storeLabel, Outlook.Items items)
            {
                StoreId = storeId ?? string.Empty;
                StoreLabel = storeLabel ?? string.Empty;
                Items = items;
            }

            public bool IsForStore(string storeId)
            {
                if (string.IsNullOrEmpty(StoreId) && string.IsNullOrEmpty(storeId))
                {
                    return true;
                }

                return string.Equals(StoreId, storeId ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            }

            public void Dispose(Outlook.ItemsEvents_ItemAddEventHandler handler)
            {
                if (Items == null)
                {
                    return;
                }

                try
                {
                    Items.ItemAdd -= handler;
                }
                catch
                {
                }

                try
                {
                    Marshal.FinalReleaseComObject(Items);
                }
                catch
                {
                }
            }
        }
}

}












