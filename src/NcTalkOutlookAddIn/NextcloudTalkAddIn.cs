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
     * Entry point and ribbon implementation for NC Connector for Outlook.
     * Registers as a classic COM add-in via IDTExtensibility2 and provides the
     * ribbon XML for appointment windows.
     */
    [ComVisible(true)]
    [Guid("A8CC9257-A153-4A01-AB35-D66CB3D44AAA")]
    [ProgId("NcTalkOutlook.AddIn")]
    public sealed class NextcloudTalkAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private Outlook.Application _outlookApplication;
        private SettingsStorage _settingsStorage;
        private AddinSettings _currentSettings;
        private readonly Dictionary<string, AppointmentSubscription> _activeSubscriptions = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _knownRoomTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByToken = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private FreeBusyManager _freeBusyManager;
        private bool _itemLoadHooked;
        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;
        private readonly List<ExplorerSubscription> _explorerSubscriptions = new List<ExplorerSubscription>();
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByEntryId = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private bool _initialScanPerformed;

        private const string PropertyToken = "NcTalkRoomToken";
        private const string PropertyRoomType = "NcTalkRoomType";
        private const string PropertyLobby = "NcTalkLobbyEnabled";
        private const string PropertySearchVisible = "NcTalkSearchVisible";
        private const string PropertyPasswordSet = "NcTalkPasswordSet";
        private const string PropertyStartEpoch = "NcTalkStartEpoch";
        private const string PropertyDataVersion = "NcTalkDataVersion";
        private const string PropertyAddUsers = "NcTalkAddUsers";
        private const string PropertyAddGuests = "NcTalkAddGuests";
        private const string PropertyDelegateId = "NcTalkDelegateId";
        private const string PropertyDelegated = "NcTalkDelegated";
        private const string IcalToken = "X-NCTALK-TOKEN";
        private const string IcalUrl = "X-NCTALK-URL";
        private const string IcalLobby = "X-NCTALK-LOBBY";
        private const string IcalStart = "X-NCTALK-START";
        private const string IcalEvent = "X-NCTALK-EVENT";
        private const string IcalObjectId = "X-NCTALK-OBJECTID";
        private const string IcalAddUsers = "X-NCTALK-ADD-USERS";
        private const string IcalAddGuests = "X-NCTALK-ADD-GUESTS";
        private const string IcalAddParticipants = "X-NCTALK-ADD-PARTICIPANTS";
        private const string IcalDelegate = "X-NCTALK-DELEGATE";
        private const string IcalDelegateName = "X-NCTALK-DELEGATE-NAME";
        private const string IcalDelegated = "X-NCTALK-DELEGATED";
        private const string IcalDelegateReady = "X-NCTALK-DELEGATE-READY";
        private const string BodySectionHeader = "Nextcloud Talk";
        private const string TalkHelpUrlMarker = "join_a_call_or_chat_as_guest.html";
        private const int PropertyVersionValue = 1;

        private static void LogCore(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Core, message);
        }

        private static void LogSettings(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Core, message);
        }

        private static void LogTalk(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Talk, message);
        }

        /**
         * Outlook calls this method when the add-in is loaded.
         * Stores the Application instance for later actions.
         */
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            _settingsStorage = new SettingsStorage();
            _currentSettings = _settingsStorage.Load();
            DiagnosticsLogger.SetEnabled(_currentSettings != null && _currentSettings.DebugLoggingEnabled);
            TryApplyOfficeUiLanguage();
            LogCore("Add-in verbunden (Outlook-Version=" + (_outlookApplication != null ? _outlookApplication.Version : "unbekannt") + ").");
            if (_currentSettings != null)
            {
                LogSettings("Einstellungen geladen (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", Debug=" + _currentSettings.DebugLoggingEnabled + ").");
            }
            _freeBusyManager = new FreeBusyManager(_settingsStorage.DataDirectory);
            _freeBusyManager.Initialize(_outlookApplication);
            EnsureItemLoadHook();
            EnsureInspectorHook();
            EnsureExplorerHook();
            ApplyIfbSettings();
        }

        private void TryApplyOfficeUiLanguage()
        {
            try
            {
                if (_outlookApplication == null)
                {
                    return;
                }

                LanguageSettings languageSettings = _outlookApplication.LanguageSettings;
                if (languageSettings == null)
                {
                    return;
                }

                int lcid = languageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
                CultureInfo culture = CultureInfo.GetCultureInfo(lcid);
                Strings.SetPreferredUiLanguage(culture.Name);
                LogCore("Office UI language erkannt: " + culture.Name + " (LCID=" + lcid + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to detect Office UI language.", ex);
            }
        }

        /**
         * Outlook signals that the add-in is unloading.
         * Cleanup hooks follow once resources are held.
         */
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
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
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to disable IFB during add-in disconnect.", ex);
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
         * Called after all add-ins have been loaded; currently unused.
         */
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /**
         * Called after Outlook startup is complete; currently unused.
         */
        public void OnStartupComplete(ref Array custom)
        {
        }

        /**
         * Called when Outlook shuts down; reserved for future cleanup steps.
         */
        public void OnBeginShutdown(ref Array custom)
        {
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
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to disable IFB during Outlook shutdown.", ex);
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
         * Outlook passes the ribbon handle right after loading.
         * Stores the instance for later refresh operations.
         */
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            // Ribbon handle not needed right now; keep as an interface stub.
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

            var configuration = new TalkServiceConfiguration(_currentSettings.ServerUrl, _currentSettings.Username, _currentSettings.AppPassword);
            PasswordPolicyInfo passwordPolicy = null;
            try
            {
                passwordPolicy = new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogTalk("Passwort-Policy konnte nicht geladen werden: " + ex.Message);
                passwordPolicy = null;
            }

            List<NextcloudUser> userDirectory = null;
            try
            {
                var cache = new IfbAddressBookCache(_settingsStorage != null ? _settingsStorage.DataDirectory : null);
                userDirectory = cache.GetUsers(configuration, _currentSettings.IfbCacheHours);
            }
            catch (Exception ex)
            {
                LogTalk("System-Adressbuch konnte nicht geladen werden: " + ex.Message);
                userDirectory = new List<NextcloudUser>();
            }

            using (var dialog = new TalkLinkForm(_currentSettings ?? new AddinSettings(), configuration, passwordPolicy, userDirectory, subject, start, end))
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
                    AppointmentStart = start,
                    AppointmentEnd = end,
                    Description = BuildInitialRoomDescription(dialog.TalkPassword, _currentSettings != null ? _currentSettings.EventDescriptionLang : "default"),
                    AddUsers = dialog.AddUsers,
                    AddGuests = dialog.AddGuests,
                    DelegateModeratorId = dialog.DelegateModeratorId,
                    DelegateModeratorName = dialog.DelegateModeratorName
                };
                LogTalk("Raumanfrage vorbereitet (Titel='" + request.Title + "', Typ=" + request.RoomType + ", Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ", PasswortGesetzt=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
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
                }
                else
                {
                    LogSettings("Einstellungsdialog ohne Aenderungen geschlossen.");
                }
            }
        }

        /**
         * Returns the embedded app icon as a COM-compatible PictureDisp.
         */
        public stdole.IPictureDisp OnGetButtonImage(IRibbonControl control)
        {
            string resourceName = "NcTalkOutlookAddIn.Resources.app.png";

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

            PasswordPolicyInfo passwordPolicy = null;
            try
            {
                passwordPolicy = new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogCore("Passwort-Policy (Freigabe) konnte nicht geladen werden: " + ex.Message);
                passwordPolicy = null;
            }

            using (var wizard = new FileLinkWizardForm(_currentSettings ?? new AddinSettings(), configuration, passwordPolicy, basePath))
            {
                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    LogCore("Filelink erstellt (Ordner=\"" + wizard.Result.FolderName + "\").");
                    string html = FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, _currentSettings != null ? _currentSettings.ShareBlockLang : "default");
                    InsertHtmlIntoMail(mail, html);
                }
            }
        }

        private Outlook.MailItem GetActiveMailItem()
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook ActiveInspector.", ex);
                inspector = null;
            }

            if (inspector != null)
            {
                try
                {
                    return inspector.CurrentItem as Outlook.MailItem;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read CurrentItem from ActiveInspector.", ex);
                }
            }

            Outlook.Explorer explorer = null;
            try
            {
                explorer = _outlookApplication.ActiveExplorer();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook ActiveExplorer.", ex);
                explorer = null;
            }

            if (explorer != null)
            {
                object inlineResponse = null;
                try
                {
                    inlineResponse = explorer.ActiveInlineResponse;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read ActiveInlineResponse from Explorer.", ex);
                    inlineResponse = null;
                }

                var mailItem = inlineResponse as Outlook.MailItem;
                if (mailItem != null)
                {
                    return mailItem;
                }
            }

            return null;
        }

        private void InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {
            if (mail == null || string.IsNullOrWhiteSpace(html))
            {
                return;
            }

            IDataObject previousClipboard = null;
            bool restoreClipboard = false;

            try
            {
                previousClipboard = Clipboard.GetDataObject();
                restoreClipboard = previousClipboard != null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read clipboard data object.", ex);
                previousClipboard = null;
                restoreClipboard = false;
            }

            try
            {
                Clipboard.SetText(html, TextDataFormat.Html);
                if (TryPasteClipboardIntoMailInspector(mail))
                {
                    LogCore("HTML-Block in E-Mail eingefuegt (WordEditor).");
                    return;
                }
            }
            catch (Exception ex)
            {
                LogCore("HTML-Insert ueber WordEditor fehlgeschlagen: " + ex.Message);
            }
            finally
            {
                if (restoreClipboard)
                {
                    try
                    {
                        Clipboard.SetDataObject(previousClipboard);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore clipboard after HTML insertion.", ex);
                    }
                }
            }

            try
            {
                string existing = mail.HTMLBody ?? string.Empty;
                string insertHtml = "<br><br>" + html;
                int bodyTagIndex = existing.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyTagIndex >= 0)
                {
                    int bodyTagEnd = existing.IndexOf(">", bodyTagIndex);
                    if (bodyTagEnd >= 0)
                    {
                        mail.HTMLBody = existing.Insert(bodyTagEnd + 1, insertHtml);
                    }
                    else
                    {
                        mail.HTMLBody = insertHtml + existing;
                    }
                }
                else
                {
                    mail.HTMLBody = insertHtml + existing;
                }

                LogCore("HTML-Block in E-Mail eingefuegt (HTMLBody Fallback).");
            }
            catch (Exception ex)
            {
                LogCore("HTML-Insert ueber HTMLBody fehlgeschlagen: " + ex.Message);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private bool TryPasteClipboardIntoMailInspector(Outlook.MailItem mail)
        {
            Outlook.Inspector inspector = null;
            object wordEditor = null;
            object application = null;
            object selection = null;

            try
            {
                inspector = mail.GetInspector;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to access MailItem.GetInspector for HTML paste.", ex);
                inspector = null;
            }

            if (inspector == null)
            {
                return false;
            }

            try
            {
                wordEditor = inspector.WordEditor;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to access Inspector.WordEditor for HTML paste.", ex);
                wordEditor = null;
            }

            if (wordEditor == null)
            {
                return false;
            }

            try
            {
                application = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                if (application == null)
                {
                    return false;
                }

                selection = application.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, application, null);
                if (selection == null)
                {
                    return false;
                }

                // Insert near the top so we are reliably above the signature block.
                try
                {
                    // wdStory = 6
                    selection.GetType().InvokeMember("HomeKey", BindingFlags.InvokeMethod, null, selection, new object[] { 6, 0 });
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to move cursor in Word editor before pasting HTML (best-effort).", ex);
                }

                selection.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, selection, null);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to paste HTML into Word editor.", ex);
                return false;
            }
            finally
            {
                if (selection != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(selection);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Word selection COM object.", ex);
                    }
                }

                if (application != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(application);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Word application COM object.", ex);
                    }
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook ActiveInspector for appointment.", ex);
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
         * Runs a quick connection test and offers to open the settings on failures.
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
                        ? Strings.ErrorCredentialsNotVerified
                        : string.Format(CultureInfo.CurrentCulture, Strings.ErrorCredentialsNotVerifiedFormat, response);
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
            LogTalk("Schreibe Talk-Daten in Termin (Token=" + result.RoomToken + ", Typ=" + request.RoomType + ", Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ", PasswortGesetzt=" + (!string.IsNullOrEmpty(request.Password)) + ", AddUsers=" + request.AddUsers + ", AddGuests=" + request.AddGuests + ", Delegate=" + (string.IsNullOrEmpty(request.DelegateModeratorId) ? "n/a" : request.DelegateModeratorId) + ").");

            if (!string.IsNullOrWhiteSpace(request.Title))
            {
                appointment.Subject = request.Title.Trim();
            }

            appointment.Location = result.RoomUrl;
            appointment.Body = UpdateBodyWithTalkBlock(appointment.Body, result.RoomUrl, request.Password, _currentSettings != null ? _currentSettings.EventDescriptionLang : "default");

            SetUserProperty(appointment, PropertyToken, Outlook.OlUserPropertyType.olText, result.RoomToken);
            var storedRoomType = result.CreatedAsEventConversation ? TalkRoomType.EventConversation : TalkRoomType.StandardRoom;
            SetUserProperty(appointment, PropertyRoomType, Outlook.OlUserPropertyType.olText, storedRoomType.ToString());
            SetUserProperty(appointment, PropertyLobby, Outlook.OlUserPropertyType.olYesNo, request.LobbyEnabled);
            SetUserProperty(appointment, PropertySearchVisible, Outlook.OlUserPropertyType.olYesNo, request.SearchVisible);
            SetUserProperty(appointment, PropertyPasswordSet, Outlook.OlUserPropertyType.olYesNo, !string.IsNullOrEmpty(request.Password));
            LogTalk("Benutzerfelder aktualisiert (Token gesetzt, Lobby=" + request.LobbyEnabled + ", Suche=" + request.SearchVisible + ").");

            long? startEpoch = TimeUtilities.ToUnixTimeSeconds(appointment.Start);
            long? endEpoch = TimeUtilities.ToUnixTimeSeconds(appointment.End);
            if (startEpoch.HasValue)
            {
                SetUserProperty(appointment, PropertyStartEpoch, Outlook.OlUserPropertyType.olText, startEpoch.Value.ToString(CultureInfo.InvariantCulture));
            }

            SetUserProperty(appointment, PropertyDataVersion, Outlook.OlUserPropertyType.olNumber, PropertyVersionValue);
            SetUserProperty(appointment, PropertyAddUsers, Outlook.OlUserPropertyType.olYesNo, request.AddUsers);
            SetUserProperty(appointment, PropertyAddGuests, Outlook.OlUserPropertyType.olYesNo, request.AddGuests);

            if (!string.IsNullOrWhiteSpace(request.DelegateModeratorId))
            {
                string delegateId = request.DelegateModeratorId.Trim();
                string delegateName = !string.IsNullOrWhiteSpace(request.DelegateModeratorName) ? request.DelegateModeratorName.Trim() : delegateId;

                SetUserProperty(appointment, PropertyDelegateId, Outlook.OlUserPropertyType.olText, delegateId);
                SetUserProperty(appointment, PropertyDelegated, Outlook.OlUserPropertyType.olYesNo, false);

                SetUserProperty(appointment, IcalDelegate, Outlook.OlUserPropertyType.olText, delegateId);
                SetUserProperty(appointment, IcalDelegateName, Outlook.OlUserPropertyType.olText, delegateName);
                SetUserProperty(appointment, IcalDelegated, Outlook.OlUserPropertyType.olText, "FALSE");
                SetUserProperty(appointment, IcalDelegateReady, Outlook.OlUserPropertyType.olText, "TRUE");
            }
            else
            {
                RemoveUserProperty(appointment, PropertyDelegateId);
                RemoveUserProperty(appointment, PropertyDelegated);

                RemoveUserProperty(appointment, IcalDelegate);
                RemoveUserProperty(appointment, IcalDelegateName);
                RemoveUserProperty(appointment, IcalDelegated);
                RemoveUserProperty(appointment, IcalDelegateReady);
            }

            SetUserProperty(appointment, IcalToken, Outlook.OlUserPropertyType.olText, result.RoomToken);
            SetUserProperty(appointment, IcalUrl, Outlook.OlUserPropertyType.olText, result.RoomUrl);
            SetUserProperty(appointment, IcalLobby, Outlook.OlUserPropertyType.olText, request.LobbyEnabled ? "TRUE" : "FALSE");
            if (startEpoch.HasValue)
            {
                SetUserProperty(appointment, IcalStart, Outlook.OlUserPropertyType.olText, startEpoch.Value.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                RemoveUserProperty(appointment, IcalStart);
            }

            SetUserProperty(appointment, IcalEvent, Outlook.OlUserPropertyType.olText, result.CreatedAsEventConversation ? "event" : "standard");
            string objectId = (startEpoch.HasValue && endEpoch.HasValue)
                ? (startEpoch.Value.ToString(CultureInfo.InvariantCulture) + "#" + endEpoch.Value.ToString(CultureInfo.InvariantCulture))
                : null;
            if (!string.IsNullOrWhiteSpace(objectId))
            {
                SetUserProperty(appointment, IcalObjectId, Outlook.OlUserPropertyType.olText, objectId);
            }
            else
            {
                RemoveUserProperty(appointment, IcalObjectId);
            }

            SetUserProperty(appointment, IcalAddUsers, Outlook.OlUserPropertyType.olText, request.AddUsers ? "TRUE" : "FALSE");
            SetUserProperty(appointment, IcalAddGuests, Outlook.OlUserPropertyType.olText, request.AddGuests ? "TRUE" : "FALSE");
            SetUserProperty(appointment, IcalAddParticipants, Outlook.OlUserPropertyType.olText, (request.AddUsers || request.AddGuests) ? "TRUE" : "FALSE");

            RegisterSubscription(appointment, result);
            TryUpdateRoomDescription(appointment, result.RoomToken, result.CreatedAsEventConversation);
        }

        private static string UpdateBodyWithTalkBlock(string existingBody, string roomUrl, string password, string languageOverride)
        {
            string body = RemoveExistingTalkBlock(existingBody ?? string.Empty);

            string joinLabel = Strings.GetInLanguage(languageOverride, "ui_description_join_label", "Join the meeting now:");
            string passwordLineFormat = Strings.GetInLanguage(languageOverride, "ui_description_password_line", "Password: {0}");
            string helpLabel = Strings.GetInLanguage(languageOverride, "ui_description_help_label", "Need help?");
            string helpUrl = Strings.GetInLanguage(
                languageOverride,
                "ui_description_help_url",
                "https://docs.nextcloud.com/server/latest/user_manual/en/talk/join_a_call_or_chat_as_guest.html");

            var lines = new List<string>
            {
                BodySectionHeader,
                string.Empty,
                joinLabel,
                roomUrl ?? string.Empty,
                string.Empty
            };

            if (!string.IsNullOrWhiteSpace(password))
            {
                lines.Add(string.Format(CultureInfo.InvariantCulture, passwordLineFormat, password.Trim()));
                lines.Add(string.Empty);
            }

            lines.Add(helpLabel);
            lines.Add(string.Empty);
            lines.Add(helpUrl);

            string block = string.Join("\r\n", lines).TrimEnd('\r', '\n');

            if (string.IsNullOrWhiteSpace(body))
            {
                return block;
            }

            return body.TrimEnd('\r', '\n') + "\r\n\r\n" + block;
        }

        private static string RemoveExistingTalkBlock(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return body;
            }

            var lines = new List<string>();
            using (var reader = new StringReader(body))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            bool removed = false;
            int index = 0;

            while (index < lines.Count)
            {
                if (!string.Equals((lines[index] ?? string.Empty).Trim(), BodySectionHeader, StringComparison.OrdinalIgnoreCase))
                {
                    index++;
                    continue;
                }

                int blockEnd = FindTalkBlockEnd(lines, index);
                if (blockEnd < 0)
                {
                    index++;
                    continue;
                }

                int blockStart = index;
                while (blockStart > 0 && string.IsNullOrWhiteSpace(lines[blockStart - 1]))
                {
                    blockStart--;
                }

                int removeEnd = blockEnd;
                while (removeEnd + 1 < lines.Count && string.IsNullOrWhiteSpace(lines[removeEnd + 1]))
                {
                    removeEnd++;
                }

                lines.RemoveRange(blockStart, removeEnd - blockStart + 1);
                removed = true;
                index = blockStart;
            }

            if (!removed)
            {
                return body;
            }

            return string.Join("\r\n", lines).Trim('\r', '\n');
        }

        private static int FindTalkBlockEnd(List<string> lines, int headerIndex)
        {
            int maxIndex = Math.Min(lines.Count - 1, headerIndex + 40);
            for (int i = headerIndex; i <= maxIndex; i++)
            {
                string line = lines[i] ?? string.Empty;
                if (line.IndexOf(TalkHelpUrlMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string BuildInitialRoomDescription(string password, string languageOverride)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                return string.Empty;
            }

            string passwordLineFormat = Strings.GetInLanguage(languageOverride, "ui_description_password_line", "Password: {0}");
            return string.Format(CultureInfo.InvariantCulture, passwordLineFormat, password.Trim());
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read AppointmentItem.EntryID.", ex);
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
                catch (Exception ex)
                {
                    // Ignore errors when saving the version asynchronously.
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to persist server version hint.", ex);
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

        private static bool HasUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            if (appointment == null)
            {
                return false;
            }

            try
            {
                return appointment.UserProperties[name] != null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to check user property '" + name + "'.", ex);
                return false;
            }
        }

        private static string GetUserPropertyTextPrefer(Outlook.AppointmentItem appointment, string primaryName, string legacyName)
        {
            string primaryValue = GetUserPropertyText(appointment, primaryName);
            if (!string.IsNullOrWhiteSpace(primaryValue))
            {
                return primaryValue;
            }

            return GetUserPropertyText(appointment, legacyName);
        }

        private static bool GetUserPropertyBoolPrefer(Outlook.AppointmentItem appointment, string primaryName, string legacyName)
        {
            if (HasUserProperty(appointment, primaryName))
            {
                return GetUserPropertyBool(appointment, primaryName);
            }

            return GetUserPropertyBool(appointment, legacyName);
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
            string eventRaw = GetUserPropertyText(appointment, IcalEvent);
            if (!string.IsNullOrWhiteSpace(eventRaw))
            {
                string normalized = eventRaw.Trim();
                if (string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.EventConversation;
                }

                if (string.Equals(normalized, "standard", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.StandardRoom;
                }
            }

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
            RemoveUserProperty(appointment, PropertyAddUsers);
            RemoveUserProperty(appointment, PropertyAddGuests);
            RemoveUserProperty(appointment, PropertyDelegateId);
            RemoveUserProperty(appointment, PropertyDelegated);
            RemoveUserProperty(appointment, IcalToken);
            RemoveUserProperty(appointment, IcalUrl);
            RemoveUserProperty(appointment, IcalLobby);
            RemoveUserProperty(appointment, IcalStart);
            RemoveUserProperty(appointment, IcalEvent);
            RemoveUserProperty(appointment, IcalObjectId);
            RemoveUserProperty(appointment, IcalAddUsers);
            RemoveUserProperty(appointment, IcalAddGuests);
            RemoveUserProperty(appointment, IcalAddParticipants);
            RemoveUserProperty(appointment, IcalDelegate);
            RemoveUserProperty(appointment, IcalDelegateName);
            RemoveUserProperty(appointment, IcalDelegated);
            RemoveUserProperty(appointment, IcalDelegateReady);
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

        private bool IsDelegatedToOtherUser(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyTextPrefer(appointment, IcalDelegate, PropertyDelegateId) ?? string.Empty;
            bool delegated = GetUserPropertyBoolPrefer(appointment, IcalDelegated, PropertyDelegated);

            if (!delegated || string.IsNullOrWhiteSpace(delegateId))
            {
                delegateId = string.Empty;
                return false;
            }

            string currentUser = _currentSettings != null ? (_currentSettings.Username ?? string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(currentUser))
            {
                return true;
            }

            return !string.Equals(delegateId.Trim(), currentUser.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private bool IsDelegationPending(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyTextPrefer(appointment, IcalDelegate, PropertyDelegateId) ?? string.Empty;
            bool delegated = GetUserPropertyBoolPrefer(appointment, IcalDelegated, PropertyDelegated);

            if (delegated)
            {
                delegateId = string.Empty;
                return false;
            }

            return !string.IsNullOrWhiteSpace(delegateId);
        }

        private void TrySyncRoomParticipants(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken) || _currentSettings == null)
            {
                return;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Teilnehmer-Sync uebersprungen (Delegation=" + delegateId + ", Token=" + roomToken + ").");
                return;
            }

            bool addParticipantsLegacy = GetUserPropertyBool(appointment, IcalAddParticipants);
            bool hasAddUsers = HasUserProperty(appointment, IcalAddUsers) || HasUserProperty(appointment, PropertyAddUsers);
            bool hasAddGuests = HasUserProperty(appointment, IcalAddGuests) || HasUserProperty(appointment, PropertyAddGuests);

            bool addUsers = hasAddUsers ? GetUserPropertyBoolPrefer(appointment, IcalAddUsers, PropertyAddUsers) : addParticipantsLegacy;
            bool addGuests = hasAddGuests ? GetUserPropertyBoolPrefer(appointment, IcalAddGuests, PropertyAddGuests) : addParticipantsLegacy;
            if (!addUsers && !addGuests)
            {
                return;
            }

            var configuration = new TalkServiceConfiguration(_currentSettings.ServerUrl, _currentSettings.Username, _currentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                return;
            }

            List<string> attendeeEmails = GetAppointmentAttendeeEmails(appointment);
            if (attendeeEmails.Count == 0)
            {
                return;
            }

            var cache = new IfbAddressBookCache(_settingsStorage != null ? _settingsStorage.DataDirectory : null);
            string selfEmail;
            cache.TryGetPrimaryEmailForUid(configuration, _currentSettings.IfbCacheHours, _currentSettings.Username, out selfEmail);

            int userAdds = 0;
            int guestAdds = 0;
            int skipped = 0;

            try
            {
                var service = CreateTalkService();

                foreach (string email in attendeeEmails)
                {
                    if (string.IsNullOrWhiteSpace(email))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(selfEmail) &&
                        string.Equals(email, selfEmail, StringComparison.OrdinalIgnoreCase))
                    {
                        skipped++;
                        continue;
                    }

                    string uid;
                    if (cache.TryGetUid(configuration, _currentSettings.IfbCacheHours, email, out uid) &&
                        !string.IsNullOrWhiteSpace(uid))
                    {
                        if (!addUsers)
                        {
                            skipped++;
                            continue;
                        }

                        if (service.AddUserParticipant(roomToken, uid))
                        {
                            userAdds++;
                        }

                        continue;
                    }

                    if (!addGuests)
                    {
                        skipped++;
                        continue;
                    }

                    if (service.AddGuestParticipant(roomToken, email))
                    {
                        guestAdds++;
                    }
                }
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Teilnehmer-Sync fehlgeschlagen: " + ex.Message);
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler beim Teilnehmer-Sync: " + ex.Message);
            }

            LogTalk("Teilnehmer-Sync abgeschlossen (Users=" + userAdds + ", Guests=" + guestAdds + ", Skipped=" + skipped + ", Token=" + roomToken + ").");
        }

        private void TryApplyDelegation(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken) || _currentSettings == null)
            {
                return;
            }

            string delegateId;
            if (!IsDelegationPending(appointment, out delegateId))
            {
                return;
            }

            string currentUser = _currentSettings.Username ?? string.Empty;
            if (!string.IsNullOrEmpty(currentUser) &&
                string.Equals(delegateId, currentUser, StringComparison.OrdinalIgnoreCase))
            {
                LogTalk("Delegation ignoriert (Delegate == aktueller Benutzer).");
                RemoveUserProperty(appointment, PropertyDelegateId);
                RemoveUserProperty(appointment, PropertyDelegated);
                RemoveUserProperty(appointment, IcalDelegate);
                RemoveUserProperty(appointment, IcalDelegateName);
                RemoveUserProperty(appointment, IcalDelegated);
                RemoveUserProperty(appointment, IcalDelegateReady);
                return;
            }

            try
            {
                var service = CreateTalkService();
                LogTalk("Delegation starten (Token=" + roomToken + ", Delegate=" + delegateId + ").");

                service.AddUserParticipant(roomToken, delegateId);

                var participants = service.GetParticipants(roomToken);
                int attendeeId = 0;
                foreach (var participant in participants)
                {
                    if (participant == null)
                    {
                        continue;
                    }

                    if (string.Equals(participant.ActorType, "users", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(participant.ActorId, delegateId, StringComparison.OrdinalIgnoreCase))
                    {
                        attendeeId = participant.AttendeeId;
                        break;
                    }
                }

                if (attendeeId <= 0)
                {
                    LogTalk("Delegation fehlgeschlagen: attendeeId nicht gefunden (Delegate=" + delegateId + ").");
                    return;
                }

                string promoteError;
                if (!service.PromoteModerator(roomToken, attendeeId, out promoteError))
                {
                    LogTalk("Delegation fehlgeschlagen: Moderator konnte nicht gesetzt werden (" + promoteError + ").");
                    string warning = string.IsNullOrWhiteSpace(promoteError)
                        ? Strings.WarningModeratorTransferFailed
                        : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, promoteError);
                    ShowWarning(warning);
                    return;
                }

                bool left = service.LeaveRoom(roomToken);
                LogTalk("Delegation abgeschlossen (Token=" + roomToken + ", Delegate=" + delegateId + ", LeftSelf=" + left + ").");

                SetUserProperty(appointment, PropertyDelegated, Outlook.OlUserPropertyType.olYesNo, true);
                SetUserProperty(appointment, IcalDelegated, Outlook.OlUserPropertyType.olText, "TRUE");
                RemoveUserProperty(appointment, IcalDelegateReady);
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Delegation fehlgeschlagen: " + ex.Message);
                string message = ex.Message;
                ShowWarning(string.IsNullOrWhiteSpace(message)
                    ? Strings.WarningModeratorTransferFailed
                    : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, message));
            }
            catch (Exception ex)
            {
                LogTalk("Unerwarteter Fehler bei Delegation: " + ex.Message);
                string message = ex.Message;
                ShowWarning(string.IsNullOrWhiteSpace(message)
                    ? Strings.WarningModeratorTransferFailed
                    : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, message));
            }
        }

        private static List<string> GetAppointmentAttendeeEmails(Outlook.AppointmentItem appointment)
        {
            var emails = new List<string>();
            if (appointment == null)
            {
                return emails;
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment recipients.", ex);
                recipients = null;
            }

            if (recipients == null)
            {
                return emails;
            }

            try
            {
                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        if (recipient == null)
                        {
                            continue;
                        }

                        int type = 0;
                        try
                        {
                            type = recipient.Type;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read recipient.Type.", ex);
                            type = 0;
                        }

                        // 1=Required, 2=Optional, 3=Resource
                        if (type == 3)
                        {
                            continue;
                        }

                        string email = TryGetRecipientSmtpAddress(recipient);
                        if (string.IsNullOrWhiteSpace(email))
                        {
                            continue;
                        }

                        email = email.Trim().ToLowerInvariant();
                        if (!emails.Contains(email))
                        {
                            emails.Add(email);
                        }
                    }
                    finally
                    {
                        if (recipient != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(recipient);
                            }
                            catch (Exception ex)
                            {
                                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to release Recipient COM object.", ex);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to enumerate appointment recipients.", ex);
            }
            finally
            {
                try
                {
                    Marshal.ReleaseComObject(recipients);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to release Recipients COM object.", ex);
                }
            }

            return emails;
        }

        private static string TryGetRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            if (recipient == null)
            {
                return null;
            }

            string address = null;
            try
            {
                address = recipient.Address;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.Address.", ex);
                address = null;
            }

            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            Outlook.AddressEntry entry = null;
            try
            {
                entry = recipient.AddressEntry;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.AddressEntry.", ex);
                entry = null;
            }

            if (entry != null)
            {
                try
                {
                    Outlook.ExchangeUser exUser = null;
                    try
                    {
                        exUser = entry.GetExchangeUser();
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve Exchange user from address entry.", ex);
                        exUser = null;
                    }

                    if (exUser != null)
                    {
                        try
                        {
                            address = exUser.PrimarySmtpAddress;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read ExchangeUser.PrimarySmtpAddress.", ex);
                            address = null;
                        }

                        try
                        {
                            Marshal.ReleaseComObject(exUser);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to release ExchangeUser COM object.", ex);
                        }
                    }
                }
                finally
                {
                    try
                    {
                        Marshal.ReleaseComObject(entry);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to release AddressEntry COM object.", ex);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            try
            {
                Outlook.PropertyAccessor accessor = recipient.PropertyAccessor;
                if (accessor != null)
                {
                    try
                    {
                        const string SmtpSchema = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        address = accessor.GetProperty(SmtpSchema) as string;
                    }
                    finally
                    {
                            try
                            {
                                Marshal.ReleaseComObject(accessor);
                            }
                            catch (Exception ex)
                            {
                                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to release PropertyAccessor COM object.", ex);
                            }
                        }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve SMTP address via PropertyAccessor.", ex);
            }

            return !string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0 ? address : null;
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

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Lobby-Update uebersprungen (Delegation=" + delegateId + ", Token=" + roomToken + ").");
                return true;
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

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Beschreibung-Update uebersprungen (Delegation=" + delegateId + ", Token=" + roomToken + ").");
                return true;
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
                _lastLobbyTimer = TimeUtilities.ToUnixTimeSeconds(appointment != null ? (DateTime?)appointment.Start : null);
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

                string delegateId;
                if (_owner.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("OnWrite uebersprungen (Delegation=" + delegateId + ", Token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                LogTalk("OnWrite fuer Termin (Token=" + _roomToken + ").");

                if (_lobbyEnabled)
                {
                    long? current = TimeUtilities.ToUnixTimeSeconds(_appointment != null ? (DateTime?)_appointment.Start : null);
                    if (current.HasValue && current != _lastLobbyTimer)
                    {
                        LogTalk("Versuche Lobby-Update waehrend OnWrite (Token=" + _roomToken + ").");
                        if (_owner.TryUpdateLobby(_appointment, _roomToken, _isEventConversation))
                        {
                            _lastLobbyTimer = current;
                            if (current.HasValue)
                            {
                                SetUserProperty(_appointment, PropertyStartEpoch, Outlook.OlUserPropertyType.olText, current.Value.ToString(CultureInfo.InvariantCulture));
                                SetUserProperty(_appointment, IcalStart, Outlook.OlUserPropertyType.olText, current.Value.ToString(CultureInfo.InvariantCulture));
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
                _owner.TrySyncRoomParticipants(_appointment, _roomToken, _isEventConversation);
                _owner.TryApplyDelegation(_appointment, _roomToken, _isEventConversation);
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

                string delegateId;
                if (_owner.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("EnsureRoomDeleted uebersprungen (Delegation=" + delegateId + ", Token=" + _roomToken + ").");
                    _owner.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    Dispose();
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
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process explorer selection item.", ex);
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to handle explorer selection change.", ex);
                }
                finally
                {
                    if (selection != null && Marshal.IsComObject(selection))
                    {
                        try
                        {
                            Marshal.ReleaseComObject(selection);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Selection COM object.", ex);
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
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to detach explorer event handlers.", ex);
                    }
                }

                if (_explorer != null && Marshal.IsComObject(_explorer))
                {
                    try
                    {
                        Marshal.ReleaseComObject(_explorer);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Explorer COM object.", ex);
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
                EnsureItemLoadHook();
                EnsureInspectorHook();
                EnsureExplorerHook();
                EnsureExistingAppointmentsSubscribed();
            }
        }

        private static class PictureConverter
        {
            /**
             * Internal helper class that wraps Image -> IPictureDisp conversion.
             */
            private sealed class AxHostPictureConverter : AxHost
            {
                public AxHostPictureConverter() : base(string.Empty)
                {
                }

                /**
                 * Converts an Image using WinForms infrastructure.
                 */
                public static stdole.IPictureDisp ImageToPictureDisp(Image image)
                {
                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
                }
            }

            /**
             * Exposes the conversion for callers.
             */
            public static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return AxHostPictureConverter.ImageToPictureDisp(image);
            }
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Inspectors.NewInspector.", ex);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Inspectors.NewInspector.", ex);
            }
            finally
            {
                try
                {
                    Marshal.FinalReleaseComObject(_inspectors);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Inspectors COM object.", ex);
                }
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
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Explorers.NewExplorer.", ex);
                    _explorers = null;
                }
            }

            try
            {
                HookExplorer(_outlookApplication.ActiveExplorer());
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook current ActiveExplorer.", ex);
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
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Explorers.NewExplorer.", ex);
                }
                finally
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(_explorers);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Explorers COM object.", ex);
                    }
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
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to access default calendar folder.", ex);
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
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to inspect calendar item during initial scan.", ex);
                        }
                        finally
                        {
                            if (current != null && Marshal.IsComObject(current))
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(current);
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release calendar item COM object.", ex);
                                }
                            }
                        }

                        raw = items.GetNext();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed while scanning calendar items for existing subscriptions.", ex);
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
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Items COM object.", ex);
                    }
                }

                if (calendar != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(calendar);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Calendar folder COM object.", ex);
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
            catch (Exception ex)
            {
                _itemLoadHooked = false;
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Outlook Application.ItemLoad.", ex);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Outlook Application.ItemLoad.", ex);
            }
            finally
            {
                _itemLoadHooked = false;
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewInspector event.", ex);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewExplorer event.", ex);
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
                string roomToken = GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
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
                bool lobbyEnabled = GetUserPropertyBoolPrefer(appointment, IcalLobby, PropertyLobby);

                RegisterSubscription(appointment, roomToken, lobbyEnabled, isEventConversation);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure subscription for appointment.", ex);
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
    }
}
