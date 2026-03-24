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
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
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
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByToken = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private FreeBusyManager _freeBusyManager;
        private bool _itemLoadHooked;
        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;
        private readonly List<ExplorerSubscription> _explorerSubscriptions = new List<ExplorerSubscription>();
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByEntryId = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly List<MailComposeSubscription> _mailComposeSubscriptions = new List<MailComposeSubscription>();
        private readonly object _mailComposeSyncRoot = new object();
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private readonly HashSet<string> _pendingAppointmentEnsureKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _pendingAppointmentEnsureSyncRoot = new object();
        private DateTime _lastDeferredAppointmentEnsureRestrictionLogUtc = DateTime.MinValue;
        private int _deferredAppointmentEnsureRestrictionSuppressedCount;
        private SynchronizationContext _uiSynchronizationContext;
        private bool _initialScanPerformed;
        private const int ComposeAttachmentEvalDebounceMs = 250;
        private const int ComposeShareCleanupSendGraceMs = 15000;

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
        private static readonly Regex TalkUrlTokenRegex = new Regex(@"https?://[^\s""'<>]+/call/(?<token>[A-Za-z0-9_-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

        private void LogDeferredAppointmentEnsureRestriction(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            lock (_pendingAppointmentEnsureSyncRoot)
            {
                DateTime nowUtc = DateTime.UtcNow;
                bool shouldEmit = _lastDeferredAppointmentEnsureRestrictionLogUtc == DateTime.MinValue ||
                                  (nowUtc - _lastDeferredAppointmentEnsureRestrictionLogUtc).TotalSeconds >= 5;
                if (shouldEmit)
                {
                    if (_deferredAppointmentEnsureRestrictionSuppressedCount > 0)
                    {
                        message = message + " (suppressed " + _deferredAppointmentEnsureRestrictionSuppressedCount.ToString(CultureInfo.InvariantCulture) + " similar entries)";
                    }

                    DiagnosticsLogger.Log(LogCategories.Core, message);
                    _lastDeferredAppointmentEnsureRestrictionLogUtc = nowUtc;
                    _deferredAppointmentEnsureRestrictionSuppressedCount = 0;
                }
                else
                {
                    _deferredAppointmentEnsureRestrictionSuppressedCount++;
                }
            }
        }

        private static void LogFileLink(string message)
        {
            DiagnosticsLogger.Log(LogCategories.FileLink, message);
        }

        /**
         * Outlook calls this method when the add-in is loaded.
         * Stores the Application instance for later actions.
         */
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            _uiSynchronizationContext = SynchronizationContext.Current;
            string outlookProfileName = ResolveCurrentOutlookProfileName();
            _settingsStorage = new SettingsStorage(outlookProfileName);
            _currentSettings = _settingsStorage.Load();
            DiagnosticsLogger.SetEnabled(_currentSettings != null && _currentSettings.DebugLoggingEnabled);
            TryApplyOfficeUiLanguage();
            LogCore("Add-in connected (Outlook version=" + (_outlookApplication != null ? _outlookApplication.Version : "unknown") + ").");
            if (!string.IsNullOrWhiteSpace(outlookProfileName))
            {
                LogCore("Using Outlook profile settings: " + outlookProfileName + ".");
            }
            if (_currentSettings != null)
            {
                LogSettings("Settings loaded (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", Debug=" + _currentSettings.DebugLoggingEnabled + ").");
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
                LogCore("Office UI language detected: " + culture.Name + " (LCID=" + lcid + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to detect Office UI language.", ex);
            }
        }

        private string ResolveCurrentOutlookProfileName()
        {
            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            object session = null;
            try
            {
                session = _outlookApplication.Session;
                if (session == null)
                {
                    return string.Empty;
                }

                object rawProfileName = session.GetType().InvokeMember(
                    "CurrentProfileName",
                    BindingFlags.GetProperty,
                    null,
                    session,
                    null,
                    CultureInfo.InvariantCulture);

                string profileName = rawProfileName as string;
                return string.IsNullOrWhiteSpace(profileName) ? string.Empty : profileName.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve current Outlook profile name.", ex);
                return string.Empty;
            }
            finally
            {
                if (session != null && Marshal.IsComObject(session))
                {
                    try
                    {
                        Marshal.ReleaseComObject(session);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to release Outlook session COM object after profile resolution.", ex);
                    }
                }
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
            UnhookMailComposeSubscriptions();
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
            _uiSynchronizationContext = null;
            lock (_pendingAppointmentEnsureSyncRoot)
            {
                _pendingAppointmentEnsureKeys.Clear();
            }
            LogCore("Add-in disconnected (removeMode=" + removeMode + ").");
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
            UnhookMailComposeSubscriptions();
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
            _uiSynchronizationContext = null;
            lock (_pendingAppointmentEnsureSyncRoot)
            {
                _pendingAppointmentEnsureKeys.Clear();
            }
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
                LogTalk("Talk link cancelled: settings are incomplete.");
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
                LogTalk("Talk link cancelled: authentication failed.");
                return;
            }

            Outlook.AppointmentItem appointment = GetActiveAppointment();
            if (appointment == null)
            {
                LogTalk("Talk link cancelled: no active appointment found.");
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
            LogTalk("Talk link started (subject='" + subject + "', start=" + start.ToString("o") + ", end=" + end.ToString("o") + ").");

            var configuration = new TalkServiceConfiguration(_currentSettings.ServerUrl, _currentSettings.Username, _currentSettings.AppPassword);
            PasswordPolicyInfo passwordPolicy = null;
            try
            {
                passwordPolicy = new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogTalk("Password policy could not be loaded: " + ex.Message);
                passwordPolicy = null;
            }

            var addressbookCache = new IfbAddressBookCache(_settingsStorage != null ? _settingsStorage.DataDirectory : null);
            LogTalk("System address book status check requested (context=talk_click, forceRefresh=True).");
            var talkClickAddressbookStatus = addressbookCache.GetSystemAddressbookStatus(
                configuration,
                _currentSettings.IfbCacheHours,
                true);
            LogTalk(
                "System address book status result (context=talk_click, available=" + talkClickAddressbookStatus.Available +
                ", count=" + talkClickAddressbookStatus.Count +
                ", hasError=" + (!string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error)) + ").");
            if (!talkClickAddressbookStatus.Available && !string.IsNullOrWhiteSpace(talkClickAddressbookStatus.Error))
            {
                LogTalk("System address book unavailable on talk click: " + talkClickAddressbookStatus.Error);
            }

            var talkWizardAddressbookStatus = talkClickAddressbookStatus;
            LogTalk(
                "System address book status reused (context=talk_wizard_open, source=talk_click, available=" + talkWizardAddressbookStatus.Available +
                ", count=" + talkWizardAddressbookStatus.Count +
                ", hasError=" + (!string.IsNullOrWhiteSpace(talkWizardAddressbookStatus.Error)) + ").");
            if (!talkWizardAddressbookStatus.Available && !string.IsNullOrWhiteSpace(talkWizardAddressbookStatus.Error))
            {
                LogTalk("System address book unavailable on talk wizard open: " + talkWizardAddressbookStatus.Error);
            }

            List<NextcloudUser> userDirectory;
            try
            {
                userDirectory = talkWizardAddressbookStatus.Available
                    ? addressbookCache.GetUsers(configuration, _currentSettings.IfbCacheHours, false)
                    : new List<NextcloudUser>();
            }
            catch (Exception ex)
            {
                LogTalk("System address book could not be loaded: " + ex.Message);
                userDirectory = new List<NextcloudUser>();
            }

            using (var dialog = new TalkLinkForm(
                _currentSettings ?? new AddinSettings(),
                configuration,
                passwordPolicy,
                userDirectory,
                talkWizardAddressbookStatus,
                subject,
                start,
                end))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    LogTalk("Talk link dialog cancelled.");
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
                LogTalk("Room request prepared (title='" + request.Title + "', type=" + request.RoomType + ", lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ", passwordSet=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
                if (!string.IsNullOrWhiteSpace(existingToken))
                {
                    LogTalk("Existing room found (token=" + existingToken + "), replacement requested.");
                    var overwrite = MessageBox.Show(
                        Strings.ConfirmReplaceRoom,
                        Strings.DialogTitle,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (overwrite != DialogResult.Yes)
                    {
                        LogTalk("Replacement declined, operation ended.");
                        return;
                    }

                    var existingType = GetRoomType(appointment);
                    bool existingIsEvent = existingType.HasValue && existingType.Value == TalkRoomType.EventConversation;
                    LogTalk("Attempting to delete existing room (event=" + existingIsEvent + ").");

                    if (!TryDeleteRoom(existingToken, existingIsEvent))
                    {
                        LogTalk("Deleting existing room failed.");
                        return;
                    }
                }

                TalkRoomCreationResult result;
                try
                {
                    using (new WaitCursorScope())
                    {
                        LogTalk("Sending CreateRoom request to Nextcloud.");
                        var service = CreateTalkService();
                        result = service.CreateRoom(request);
                    }
                    LogTalk("Room created successfully (token=" + result.RoomToken + ", URL=" + result.RoomUrl + ", event=" + result.CreatedAsEventConversation + ").");
                }
                catch (TalkServiceException ex)
                {
                    LogTalk("Talk room could not be created: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoom, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    LogTalk("Unexpected error while creating talk room: " + ex.Message);
                    MessageBox.Show(
                        string.Format(Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                ApplyRoomToAppointment(appointment, request, result);
                LogTalk("Room data stored in appointment (EntryID=" + (appointment.EntryID ?? "n/a") + ").");

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
            LogSettings("Settings dialog opened.");

            using (var form = new SettingsForm(_currentSettings, _outlookApplication))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _currentSettings = form.Result ?? new AddinSettings();
                    LogSettings("Settings applied (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", Debug=" + _currentSettings.DebugLoggingEnabled + ").");
                    ApplyIfbSettings();
                    if (_settingsStorage != null)
                    {
                        _settingsStorage.Save(_currentSettings);
                    }
                }
                else
                {
                    LogSettings("Settings dialog closed without changes.");
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

            EnsureMailComposeSubscription(mail, ResolveActiveInspectorIdentityKey());
            RunFileLinkWizardForMail(mail, null);
        }

        private bool RunFileLinkWizardForMail(Outlook.MailItem mail, FileLinkWizardLaunchOptions launchOptions)
        {
            if (mail == null || _currentSettings == null)
            {
                return false;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);

            string basePath = string.IsNullOrWhiteSpace(_currentSettings.FileLinkBasePath)
                ? "90 Freigaben - extern"
                : _currentSettings.FileLinkBasePath;

            PasswordPolicyInfo passwordPolicy = null;
            try
            {
                passwordPolicy = new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogFileLink("Sharing password policy could not be loaded: " + ex.Message);
                passwordPolicy = null;
            }

            using (var wizard = new FileLinkWizardForm(_currentSettings ?? new AddinSettings(), configuration, passwordPolicy, basePath, launchOptions))
            {
                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    string languageOverride = _currentSettings != null ? _currentSettings.ShareBlockLang : "default";
                    LogFileLink("Share created (folder=\"" + wizard.Result.FolderName + "\").");

                    MailComposeSubscription composeSubscription = EnsureMailComposeSubscription(mail, ResolveActiveInspectorIdentityKey());
                    if (composeSubscription != null)
                    {
                        composeSubscription.ArmShareCleanup(wizard.Result);
                    }

                    string html = FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, languageOverride);

                    if (composeSubscription != null
                        && wizard.RequestSnapshot != null
                        && AddinSettings.SeparatePasswordFeatureEnabled
                        && wizard.RequestSnapshot.PasswordSeparateEnabled
                        && !string.IsNullOrWhiteSpace(wizard.Result.Password))
                    {
                        string passwordOnlyHtml = FileLinkHtmlBuilder.BuildPasswordOnly(wizard.Result, languageOverride);
                        composeSubscription.RegisterSeparatePasswordDispatch(
                            wizard.Result,
                            wizard.RequestSnapshot,
                            passwordOnlyHtml);
                    }

                    InsertHtmlIntoMail(mail, html);
                    return true;
                }
            }

            return false;
        }

        private sealed class SeparatePasswordDispatchEntry
        {
            internal string ShareLabel { get; set; }

            internal string ShareUrl { get; set; }

            internal string Password { get; set; }

            internal string Html { get; set; }

            internal string To { get; set; }

            internal string Cc { get; set; }

            internal string Bcc { get; set; }
        }

        private MailComposeSubscription EnsureMailComposeSubscription(Outlook.MailItem mail, string inspectorIdentityOverride = null)
        {
            if (mail == null)
            {
                return null;
            }

            string mailIdentityKey = ResolveMailComIdentityKey(mail);
            string inspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityOverride)
                ? ResolveMailInspectorIdentityKey(mail)
                : inspectorIdentityOverride.Trim();

            lock (_mailComposeSyncRoot)
            {
                foreach (var existing in _mailComposeSubscriptions)
                {
                    if (existing.IsFor(mail, mailIdentityKey, inspectorIdentityKey))
                    {
                        return existing;
                    }
                }

                var created = new MailComposeSubscription(this, mail, mailIdentityKey, inspectorIdentityKey);
                _mailComposeSubscriptions.Add(created);
                return created;
            }
        }

        private static string ResolveMailComIdentityKey(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return string.Empty;
            }

            IntPtr unk = IntPtr.Zero;
            try
            {
                unk = Marshal.GetIUnknownForObject(mail);
                if (unk == IntPtr.Zero)
                {
                    return string.Empty;
                }

                return unchecked((ulong)unk.ToInt64()).ToString("X16", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve MailItem COM identity key.", ex);
                return string.Empty;
            }
            finally
            {
                if (unk != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.Release(unk);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release MailItem COM identity pointer.", ex);
                    }
                }
            }
        }

        private static string ResolveMailInspectorIdentityKey(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                return ResolveComIdentityKey(inspector, "Inspector");
            }
            catch (COMException ex)
            {
                uint errorCode = unchecked((uint)ex.ErrorCode);
                if ((errorCode & 0xFFFFu) == 0x0108u)
                {
                    LogFileLink(
                        "MailItem.GetInspector unavailable while resolving compose inspector identity (hresult=0x"
                        + errorCode.ToString("X8", CultureInfo.InvariantCulture)
                        + ").");
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.GetInspector for compose identity.", ex);
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.GetInspector for compose identity.", ex);
                return string.Empty;
            }
            finally
            {
                if (inspector != null && Marshal.IsComObject(inspector))
                {
                    try
                    {
                        Marshal.ReleaseComObject(inspector);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release compose Inspector COM object.", ex);
                    }
                }
            }
        }

        private sealed class NativeWindowOwner : IWin32Window
        {
            private readonly IntPtr _handle;

            internal NativeWindowOwner(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle
            {
                get { return _handle; }
            }
        }

        private IWin32Window TryCreateMailInspectorDialogOwner(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                if (inspector == null)
                {
                    return null;
                }

                int hwnd = ReadInspectorWindowHandle(inspector);

                return hwnd > 0 ? new NativeWindowOwner(new IntPtr(hwnd)) : null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve compose prompt owner inspector.", ex);
                return null;
            }
            finally
            {
                if (inspector != null && Marshal.IsComObject(inspector))
                {
                    try
                    {
                        Marshal.ReleaseComObject(inspector);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release compose prompt owner Inspector COM object.", ex);
                    }
                }
            }
        }

        private static int ReadInspectorWindowHandle(Outlook.Inspector inspector)
        {
            if (inspector == null)
            {
                return 0;
            }

            foreach (string propertyName in new[] { "HWND", "Hwnd" })
            {
                try
                {
                    PropertyInfo property = inspector.GetType().GetProperty(propertyName);
                    if (property == null)
                    {
                        continue;
                    }

                    object value = property.GetValue(inspector, null);
                    if (value == null)
                    {
                        continue;
                    }

                    int hwnd;
                    if (int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out hwnd) && hwnd > 0)
                    {
                        return hwnd;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read inspector window handle property '" + propertyName + "'.",
                        ex);
                }
            }

            return 0;
        }

        private static string ResolveComIdentityKey(object comObject, string objectName)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return string.Empty;
            }

            IntPtr unk = IntPtr.Zero;
            try
            {
                unk = Marshal.GetIUnknownForObject(comObject);
                if (unk == IntPtr.Zero)
                {
                    return string.Empty;
                }

                return unchecked((ulong)unk.ToInt64()).ToString("X16", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to resolve COM identity key for " + (objectName ?? "object") + ".",
                    ex);
                return string.Empty;
            }
            finally
            {
                if (unk != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.Release(unk);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to release COM identity pointer for " + (objectName ?? "object") + ".",
                            ex);
                    }
                }
            }
        }

        private void RemoveMailComposeSubscription(MailComposeSubscription subscription)
        {
            if (subscription == null)
            {
                return;
            }

            lock (_mailComposeSyncRoot)
            {
                _mailComposeSubscriptions.Remove(subscription);
            }
        }

        private void UnhookMailComposeSubscriptions()
        {
            MailComposeSubscription[] current;
            lock (_mailComposeSyncRoot)
            {
                current = _mailComposeSubscriptions.ToArray();
                _mailComposeSubscriptions.Clear();
            }

            foreach (var subscription in current)
            {
                try
                {
                    subscription.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose mail compose subscription.", ex);
                }
            }
        }

        private bool TryGetAttachmentAutomationGuardState(string stage, string composeKey, out OutlookAttachmentAutomationGuardService.GuardState state)
        {
            state = null;
            try
            {
                state = _attachmentGuardService.ReadLiveState();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read live attachment automation guard state.", ex);
                return false;
            }

            if (state == null || !state.LockActive)
            {
                return false;
            }

            LogFileLink(
                "Compose attachment automation blocked by host setting (stage="
                + (stage ?? string.Empty)
                + ", composeKey="
                + (composeKey ?? string.Empty)
                + ", thresholdMb="
                + state.ThresholdMb.ToString(CultureInfo.InvariantCulture)
                + ", source="
                + (state.Source ?? string.Empty)
                + ").");
            return true;
        }

        private bool TryDeleteComposeShareFolder(string relativeFolder, string reason, string shareId, string shareLabel)
        {
            if (string.IsNullOrWhiteSpace(relativeFolder))
            {
                return true;
            }

            EnsureSettingsLoaded();
            if (_currentSettings == null || !SettingsAreComplete())
            {
                LogFileLink(
                    "Compose share cleanup skipped (settings incomplete): relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty));
                return false;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);
            var service = new FileLinkService(configuration);
            try
            {
                service.DeleteShareFolder(relativeFolder, CancellationToken.None);
                LogFileLink(
                    "Compose share cleanup delete success (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (shareId ?? string.Empty)
                    + ", shareLabel="
                    + (shareLabel ?? string.Empty)
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Compose share cleanup delete failure (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (shareId ?? string.Empty)
                    + ", shareLabel="
                    + (shareLabel ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
        }

        private void DispatchSeparatePasswordMailQueue(string composeKey, List<SeparatePasswordDispatchEntry> queue)
        {
            if (queue == null || queue.Count == 0 || _outlookApplication == null)
            {
                return;
            }

            int attemptedDispatches = 0;
            int successfulDispatches = 0;
            int autoSendFailures = 0;
            int fallbackOpenedCount = 0;
            int fallbackOpenFailures = 0;
            string lastFailureMessage = string.Empty;
            var sentRecipients = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var dispatch in queue)
            {
                if (dispatch == null || string.IsNullOrWhiteSpace(dispatch.Password) || string.IsNullOrWhiteSpace(dispatch.Html))
                {
                    continue;
                }

                attemptedDispatches++;
                Outlook.MailItem passwordMail = null;
                try
                {
                    passwordMail = _outlookApplication.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    if (passwordMail == null)
                    {
                        throw new InvalidOperationException("Password mail draft could not be created.");
                    }

                    passwordMail.Subject = BuildSeparatePasswordMailSubject(dispatch);
                    passwordMail.HTMLBody = dispatch.Html;
                    List<string> resolvedRecipients = ApplySeparatePasswordRecipientsForSend(passwordMail, dispatch, composeKey);
                    int resolvedRecipientCount = resolvedRecipients.Count;

                    LogFileLink(
                        "Separate password mail send start (composeKey="
                        + (composeKey ?? string.Empty)
                        + ", to="
                        + BuildNormalizedRecipientCsv(dispatch.To)
                        + ", cc="
                        + BuildNormalizedRecipientCsv(dispatch.Cc)
                        + ", bcc="
                        + BuildNormalizedRecipientCsv(dispatch.Bcc)
                        + ", resolvedRecipients="
                        + resolvedRecipientCount.ToString(CultureInfo.InvariantCulture)
                        + ").");

                    ((Outlook._MailItem)passwordMail).Send();
                    successfulDispatches++;
                    AddRecipientAddresses(sentRecipients, resolvedRecipients);
                    LogFileLink("Separate password mail send done (composeKey=" + (composeKey ?? string.Empty) + ").");
                }
                catch (Exception ex)
                {
                    autoSendFailures++;
                    lastFailureMessage = ex.Message ?? string.Empty;
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Separate password mail auto-send failed (composeKey=" + (composeKey ?? string.Empty) + ").",
                        ex);
                    bool fallbackOpened = TryOpenSeparatePasswordFallback(dispatch, composeKey);
                    if (fallbackOpened)
                    {
                        fallbackOpenedCount++;
                    }
                    else
                    {
                        fallbackOpenFailures++;
                    }
                }
                finally
                {
                    if (passwordMail != null && Marshal.IsComObject(passwordMail))
                    {
                        try
                        {
                            Marshal.ReleaseComObject(passwordMail);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release password MailItem COM object.", ex);
                        }
                    }
                }
            }

            int recipientCount = sentRecipients.Count;
            if (attemptedDispatches > 0 && successfulDispatches == attemptedDispatches && autoSendFailures == 0 && recipientCount > 0)
            {
                LogFileLink(
                    "Separate password mail sent (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", attempted="
                    + attemptedDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", successful="
                    + successfulDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", recipients="
                    + recipientCount.ToString(CultureInfo.InvariantCulture)
                    + ").");
                ShowPasswordMailSuccessNotification(recipientCount);
            }
            else
            {
                LogFileLink(
                    "Separate password mail partially sent (manual fallback required) (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", attempted="
                    + attemptedDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", successful="
                    + successfulDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", recipients="
                    + recipientCount.ToString(CultureInfo.InvariantCulture)
                    + ", fallbackOpened="
                    + fallbackOpenedCount.ToString(CultureInfo.InvariantCulture)
                    + ", fallbackOpenFailures="
                    + fallbackOpenFailures.ToString(CultureInfo.InvariantCulture)
                    + ", autoSendFailures="
                    + autoSendFailures.ToString(CultureInfo.InvariantCulture)
                    + ").");

                if (autoSendFailures > 0 && fallbackOpenedCount == 0 && !string.IsNullOrWhiteSpace(lastFailureMessage))
                {
                    ShowPasswordMailFailureDialog(lastFailureMessage);
                }
            }
        }

        private bool TryOpenSeparatePasswordFallback(SeparatePasswordDispatchEntry dispatch, string composeKey)
        {
            if (dispatch == null || _outlookApplication == null)
            {
                return false;
            }

            Outlook.MailItem fallback = null;
            try
            {
                fallback = _outlookApplication.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (fallback == null)
                {
                    return false;
                }

                string toRecipients = BuildNormalizedRecipientCsv(dispatch.To);
                string ccRecipients = BuildNormalizedRecipientCsv(dispatch.Cc);
                string bccRecipients = BuildNormalizedRecipientCsv(dispatch.Bcc);
                if (CountRecipientsInCsv(toRecipients) + CountRecipientsInCsv(ccRecipients) + CountRecipientsInCsv(bccRecipients) <= 0)
                {
                    throw new InvalidOperationException("Separate password fallback draft has no valid recipients.");
                }

                fallback.To = toRecipients;
                fallback.CC = ccRecipients;
                fallback.BCC = bccRecipients;
                fallback.Subject = BuildSeparatePasswordMailSubject(dispatch);
                fallback.HTMLBody = dispatch.Html ?? string.Empty;
                fallback.Display(false);
                LogFileLink(
                    "Separate password mail manual fallback opened (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", to="
                    + toRecipients
                    + ", cc="
                    + ccRecipients
                    + ", bcc="
                    + bccRecipients
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password mail manual fallback failed (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                return false;
            }
            finally
            {
                if (fallback != null && Marshal.IsComObject(fallback))
                {
                    try
                    {
                        Marshal.ReleaseComObject(fallback);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release password fallback MailItem COM object.", ex);
                    }
                }
            }
        }

        private List<string> ApplySeparatePasswordRecipientsForSend(Outlook.MailItem mail, SeparatePasswordDispatchEntry dispatch, string composeKey)
        {
            if (mail == null)
            {
                throw new InvalidOperationException("Password mail is not available.");
            }

            List<string> toRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.To : string.Empty);
            List<string> ccRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.Cc : string.Empty);
            List<string> bccRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.Bcc : string.Empty);
            int totalRecipients = toRecipients.Count + ccRecipients.Count + bccRecipients.Count;
            if (totalRecipients <= 0)
            {
                throw new InvalidOperationException("Separate password mail has no valid recipients.");
            }

            var resolvedRecipients = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                if (recipients == null)
                {
                    throw new InvalidOperationException("Password mail recipients collection is not available.");
                }

                AddResolvedRecipients(recipients, toRecipients, Outlook.OlMailRecipientType.olTo, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, ccRecipients, Outlook.OlMailRecipientType.olCC, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, bccRecipients, Outlook.OlMailRecipientType.olBCC, composeKey, resolvedRecipients);

                bool resolvedAll = recipients.ResolveAll();
                if (!resolvedAll)
                {
                    throw new InvalidOperationException("Separate password mail recipients could not be resolved.");
                }

                if (resolvedRecipients.Count <= 0)
                {
                    throw new InvalidOperationException("Separate password mail has no resolvable recipients.");
                }

                return resolvedRecipients;
            }
            finally
            {
                if (recipients != null && Marshal.IsComObject(recipients))
                {
                    try
                    {
                        Marshal.ReleaseComObject(recipients);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release password Recipients COM object.", ex);
                    }
                }
            }
        }

        private void AddResolvedRecipients(
            Outlook.Recipients recipients,
            List<string> addresses,
            Outlook.OlMailRecipientType type,
            string composeKey,
            List<string> resolvedRecipients)
        {
            if (recipients == null || addresses == null || addresses.Count == 0)
            {
                return;
            }

            for (int i = 0; i < addresses.Count; i++)
            {
                string address = addresses[i] ?? string.Empty;
                Outlook.Recipient recipient = null;
                try
                {
                    recipient = recipients.Add(address);
                    if (recipient == null)
                    {
                        throw new InvalidOperationException("Recipient could not be added.");
                    }

                    recipient.Type = (int)type;
                    bool resolved = recipient.Resolve();
                    if (!resolved)
                    {
                        throw new InvalidOperationException(
                            "Recipient could not be resolved (composeKey="
                            + (composeKey ?? string.Empty)
                            + ", address="
                            + address
                            + ", type="
                            + type.ToString()
                            + ").");
                    }

                    AddUniqueRecipient(resolvedRecipients, address);
                }
                finally
                {
                    if (recipient != null && Marshal.IsComObject(recipient))
                    {
                        try
                        {
                            Marshal.ReleaseComObject(recipient);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release password Recipient COM object.", ex);
                        }
                    }
                }
            }
        }

        private void ShowPasswordMailFailureDialog(string detailMessage)
        {
            if (string.IsNullOrWhiteSpace(detailMessage))
            {
                return;
            }

            try
            {
                MessageBox.Show(
                    detailMessage.Trim(),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to show separate password failure dialog.", ex);
            }
        }

        private string BuildSeparatePasswordMailSubject(SeparatePasswordDispatchEntry dispatch)
        {
            string baseSubject = Strings.SharingPasswordMailSubject;
            string shareLabel = dispatch != null ? (dispatch.ShareLabel ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(shareLabel))
            {
                return baseSubject;
            }

            return string.Format(CultureInfo.CurrentCulture, Strings.SharingPasswordMailSubjectWithLabel, shareLabel);
        }

        private static void AddRecipientAddresses(HashSet<string> recipients, List<string> addresses)
        {
            if (recipients == null || addresses == null)
            {
                return;
            }

            for (int i = 0; i < addresses.Count; i++)
            {
                string normalized = NormalizeRecipientAddress(addresses[i]);
                if (!string.IsNullOrWhiteSpace(normalized))
                {
                    recipients.Add(normalized);
                }
            }
        }

        private static void AddUniqueRecipient(List<string> recipients, string address)
        {
            if (recipients == null)
            {
                return;
            }

            string normalized = NormalizeRecipientAddress(address);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            for (int i = 0; i < recipients.Count; i++)
            {
                if (string.Equals(recipients[i], normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            recipients.Add(normalized);
        }

        private static List<string> ExtractRecipientAddresses(string csv)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(csv))
            {
                return list;
            }

            string[] parts = csv.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                AddUniqueRecipient(list, parts[i]);
            }

            return list;
        }

        private static string BuildNormalizedRecipientCsv(string csv)
        {
            List<string> recipients = ExtractRecipientAddresses(csv);
            return recipients.Count == 0 ? string.Empty : string.Join("; ", recipients.ToArray());
        }

        private static int CountRecipientsInCsv(string csv)
        {
            return ExtractRecipientAddresses(csv).Count;
        }

        private static string NormalizeRecipientAddress(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            string value = raw.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            int lt = value.LastIndexOf('<');
            int gt = value.LastIndexOf('>');
            if (lt >= 0 && gt > lt)
            {
                value = value.Substring(lt + 1, gt - lt - 1).Trim();
            }

            value = value.Trim().Trim('\'', '"');
            return value.Trim();
        }

        private void ShowPasswordMailSuccessNotification(int recipientCount)
        {
            if (recipientCount <= 0)
            {
                return;
            }

            try
            {
                var notifyIcon = new NotifyIcon();
                notifyIcon.Icon = BrandingAssets.GetAppIcon(32);
                notifyIcon.Visible = true;
                notifyIcon.BalloonTipTitle = Strings.SharingPasswordMailNotificationTitle;
                notifyIcon.BalloonTipText = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.SharingPasswordMailNotificationSuccess,
                    recipientCount.ToString(CultureInfo.CurrentCulture));
                notifyIcon.ShowBalloonTip(5000);
                Task.Delay(7000).ContinueWith(_ =>
                {
                    try
                    {
                        notifyIcon.Visible = false;
                        notifyIcon.Dispose();
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose password notification icon.", ex);
                    }
                });
                LogFileLink("Separate password notification shown (recipients=" + recipientCount.ToString(CultureInfo.InvariantCulture) + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Separate password notification failed.", ex);
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

        private string ResolveActiveInspectorIdentityKey()
        {
            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = _outlookApplication.ActiveInspector();
                return ResolveComIdentityKey(inspector, "ActiveInspector");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve active inspector identity key.", ex);
                return string.Empty;
            }
            finally
            {
                if (inspector != null && Marshal.IsComObject(inspector))
                {
                    try
                    {
                        Marshal.ReleaseComObject(inspector);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release active Inspector COM object.", ex);
                    }
                }
            }
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
                    LogCore("Inserted HTML block into mail (WordEditor).");
                    return;
                }
            }
            catch (Exception ex)
            {
                LogCore("Failed to insert HTML via WordEditor: " + ex.Message);
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

                LogCore("Inserted HTML block into mail (HTMLBody fallback).");
            }
            catch (Exception ex)
            {
                LogCore("Failed to insert HTML via HTMLBody fallback: " + ex.Message);
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
                LogCore("Applying IFB (Enabled=" + _currentSettings.IfbEnabled + ", Days=" + _currentSettings.IfbDays + ", CacheHours=" + _currentSettings.IfbCacheHours + ").");
                _freeBusyManager.ApplySettings(_currentSettings);
            }
            catch (Exception ex)
            {
                LogCore("Failed to start IFB: " + ex.Message);
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
                    LogTalk("Starting credential verification request.");
                    if (service.VerifyConnection(out response))
                    {
                        UpdateStoredServerVersion(response);
                        LogTalk("Credentials verified (response=" + (string.IsNullOrEmpty(response) ? "OK" : response) + ").");
                        return true;
                    }

                    string message = string.IsNullOrEmpty(response)
                        ? Strings.ErrorCredentialsNotVerified
                        : string.Format(CultureInfo.CurrentCulture, Strings.ErrorCredentialsNotVerifiedFormat, response);
                    LogTalk("Invalid credentials: " + message);
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

                LogTalk("Connection check failed: " + message);
                return PromptOpenSettings(message, control);
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error during connection check: " + ex.Message);
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
            LogTalk("Writing talk data to appointment (token=" + result.RoomToken + ", type=" + request.RoomType + ", lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ", passwordSet=" + (!string.IsNullOrEmpty(request.Password)) + ", addUsers=" + request.AddUsers + ", addGuests=" + request.AddGuests + ", delegate=" + (string.IsNullOrEmpty(request.DelegateModeratorId) ? "n/a" : request.DelegateModeratorId) + ").");

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
            LogTalk("User fields updated (token set, lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ").");

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

        private bool TryStampIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            startEpoch = 0;
            if (appointment == null)
            {
                LogTalk("Failed to stamp X-NCTALK-START: appointment is null (token=" + (roomToken ?? "n/a") + ").");
                return false;
            }

            DateTime start;
            try
            {
                start = appointment.Start;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Appointment.Start while stamping X-NCTALK-START (token=" + (roomToken ?? "n/a") + ").", ex);
                return false;
            }

            long? epoch = TimeUtilities.ToUnixTimeSeconds(start);
            if (!epoch.HasValue || epoch.Value <= 0)
            {
                RemoveUserProperty(appointment, IcalStart);
                RemoveUserProperty(appointment, PropertyStartEpoch);
                LogTalk("Failed to stamp X-NCTALK-START: Appointment.Start is missing/invalid (token=" + (roomToken ?? "n/a") + ", start=" + start.ToString("o", CultureInfo.InvariantCulture) + ").");
                return false;
            }

            startEpoch = epoch.Value;
            string value = startEpoch.ToString(CultureInfo.InvariantCulture);
            SetUserProperty(appointment, IcalStart, Outlook.OlUserPropertyType.olText, value);
            SetUserProperty(appointment, PropertyStartEpoch, Outlook.OlUserPropertyType.olText, value);
            LogTalk("X-NCTALK-START stamped (token=" + (roomToken ?? "n/a") + ", startEpoch=" + value + ").");
            return true;
        }

        private bool TryReadRequiredIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            startEpoch = 0;
            string rawValue = GetUserPropertyText(appointment, IcalStart);
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                LogTalk("Lobby update blocked: X-NCTALK-START is missing (token=" + (roomToken ?? "n/a") + ").");
                return false;
            }

            long parsed;
            if (!long.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
            {
                LogTalk("Lobby update blocked: X-NCTALK-START is invalid (token=" + (roomToken ?? "n/a") + ", value='" + rawValue + "').");
                return false;
            }

            startEpoch = parsed;
            return true;
        }

        private static long? GetIcalStartEpochOrNull(Outlook.AppointmentItem appointment)
        {
            string rawValue = GetUserPropertyText(appointment, IcalStart);
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            long parsed;
            if (!long.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
            {
                return null;
            }

            return parsed;
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

        private static string TryGetEntryIdForDeferredKey(Outlook.AppointmentItem appointment)
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

        private static string TryGetGlobalAppointmentIdForDeferredKey(Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment != null ? appointment.GlobalAppointmentID : null;
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

        private static string ExtractTokenFromTalkUrlText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            Match match = TalkUrlTokenRegex.Match(text);
            if (!match.Success)
            {
                return null;
            }

            string token = match.Groups["token"] != null ? match.Groups["token"].Value : null;
            return string.IsNullOrWhiteSpace(token) ? null : token.Trim();
        }

        private string ResolveRoomTokenForAppointment(Outlook.AppointmentItem appointment)
        {
            string roomToken = GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
            if (!string.IsNullOrWhiteSpace(roomToken))
            {
                return roomToken.Trim();
            }

            string location = null;
            try
            {
                location = appointment != null ? appointment.Location : null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read appointment location while resolving Talk token.", ex);
            }

            string extracted = ExtractTokenFromTalkUrlText(location);
            if (string.IsNullOrWhiteSpace(extracted))
            {
                return null;
            }

            try
            {
                SetUserProperty(appointment, IcalToken, Outlook.OlUserPropertyType.olText, extracted);
                SetUserProperty(appointment, PropertyToken, Outlook.OlUserPropertyType.olText, extracted);
                LogTalk("Room token bootstrapped from location (token=" + extracted + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to persist bootstrapped room token.", ex);
            }

            return extracted;
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

        private void ResolveRuntimeRoomTraits(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool fallbackLobbyEnabled,
            bool fallbackIsEventConversation,
            out bool lobbyKnown,
            out bool lobbyEnabled,
            out bool isEventConversation)
        {
            bool? resolvedLobby = null;
            bool? resolvedEventConversation = null;

            try
            {
                bool hasLobbyFlag = HasUserProperty(appointment, IcalLobby) || HasUserProperty(appointment, PropertyLobby);
                if (hasLobbyFlag)
                {
                    resolvedLobby = GetUserPropertyBoolPrefer(appointment, IcalLobby, PropertyLobby);
                }

                var roomType = GetRoomType(appointment);
                if (roomType.HasValue)
                {
                    resolvedEventConversation = roomType.Value == TalkRoomType.EventConversation;
                }

                if (!resolvedEventConversation.HasValue && TryIsEventConversationFromObjectId(appointment))
                {
                    resolvedEventConversation = true;
                    PersistEventConversationTraits(appointment, roomToken);
                    LogTalk("Conversation type inferred from X-NCTALK-OBJECTID (token=" + roomToken + ").");
                }

                TryHydrateMissingRoomTraitsFromServer(appointment, roomToken, ref resolvedLobby, ref resolvedEventConversation);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve runtime room traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }

            lobbyKnown = resolvedLobby.HasValue;
            lobbyEnabled = resolvedLobby.HasValue ? resolvedLobby.Value : fallbackLobbyEnabled;
            isEventConversation = resolvedEventConversation.HasValue ? resolvedEventConversation.Value : fallbackIsEventConversation;
        }

        private static bool TryIsEventConversationFromObjectId(Outlook.AppointmentItem appointment)
        {
            try
            {
                string objectId = GetUserPropertyText(appointment, IcalObjectId);
                return !string.IsNullOrWhiteSpace(objectId);
            }
            catch
            {
                return false;
            }
        }

        private void PersistEventConversationTraits(Outlook.AppointmentItem appointment, string roomToken)
        {
            if (appointment == null)
            {
                return;
            }

            try
            {
                SetUserProperty(appointment, IcalEvent, Outlook.OlUserPropertyType.olText, "event");
                SetUserProperty(appointment, PropertyRoomType, Outlook.OlUserPropertyType.olText, TalkRoomType.EventConversation.ToString());
                LogTalk("Conversation type persisted as event (token=" + (roomToken ?? "n/a") + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist event conversation traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }
        }

        private void PersistLobbyTraits(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled)
        {
            if (appointment == null)
            {
                return;
            }

            try
            {
                SetUserProperty(appointment, IcalLobby, Outlook.OlUserPropertyType.olText, lobbyEnabled ? "TRUE" : "FALSE");
                SetUserProperty(appointment, PropertyLobby, Outlook.OlUserPropertyType.olYesNo, lobbyEnabled);
                LogTalk("Lobby flag persisted from runtime update (token=" + (roomToken ?? "n/a") + ", lobby=" + lobbyEnabled + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist lobby traits from runtime update (token=" + (roomToken ?? "n/a") + ").", ex);
            }
        }

        private static bool IsEventConversationDescriptionError(TalkServiceException ex)
        {
            if (ex == null || string.IsNullOrWhiteSpace(ex.Message))
            {
                return false;
            }

            string normalized = ex.Message.Trim();
            return string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase) ||
                   normalized.IndexOf("event conversation", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsMissingOrForbiddenRoomMutationError(TalkServiceException ex)
        {
            if (ex == null)
            {
                return false;
            }

            return ex.StatusCode == HttpStatusCode.NotFound ||
                   ex.StatusCode == HttpStatusCode.Forbidden;
        }

        private static string GetNormalizedRoomName(Outlook.AppointmentItem appointment)
        {
            if (appointment == null)
            {
                return string.Empty;
            }

            try
            {
                string subject = appointment.Subject;
                return string.IsNullOrWhiteSpace(subject) ? string.Empty : subject.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment subject for room name sync.", ex);
                return string.Empty;
            }
        }

        private void TryHydrateMissingRoomTraitsFromServer(
            Outlook.AppointmentItem appointment,
            string roomToken,
            ref bool? lobbyEnabled,
            ref bool? isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }

            bool needLobby = !lobbyEnabled.HasValue;
            bool needEvent = !isEventConversation.HasValue;
            if (!needLobby && !needEvent)
            {
                return;
            }

            try
            {
                bool? resolvedLobby;
                bool? resolvedEvent;
                var service = CreateTalkService();
                if (!service.TryReadRoomTraits(roomToken, out resolvedLobby, out resolvedEvent))
                {
                    return;
                }

                if (needLobby && resolvedLobby.HasValue)
                {
                    lobbyEnabled = resolvedLobby.Value;
                    try
                    {
                        SetUserProperty(appointment, IcalLobby, Outlook.OlUserPropertyType.olText, resolvedLobby.Value ? "TRUE" : "FALSE");
                        SetUserProperty(appointment, PropertyLobby, Outlook.OlUserPropertyType.olYesNo, resolvedLobby.Value);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist lobby flag after server bootstrap.", ex);
                    }

                    LogTalk("Lobby flag bootstrapped from server (token=" + roomToken + ", lobby=" + resolvedLobby.Value + ").");
                }

                if (needEvent && resolvedEvent.HasValue)
                {
                    isEventConversation = resolvedEvent.Value;
                    try
                    {
                        SetUserProperty(appointment, IcalEvent, Outlook.OlUserPropertyType.olText, resolvedEvent.Value ? "event" : "standard");
                        SetUserProperty(
                            appointment,
                            PropertyRoomType,
                            Outlook.OlUserPropertyType.olText,
                            resolvedEvent.Value ? TalkRoomType.EventConversation.ToString() : TalkRoomType.StandardRoom.ToString());
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist conversation type after server bootstrap.", ex);
                    }

                    LogTalk("Conversation type bootstrapped from server (token=" + roomToken + ", event=" + resolvedEvent.Value + ").");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to bootstrap room traits from server (token=" + roomToken + ").", ex);
            }
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
            LogTalk("Registering appointment subscription (token=" + roomToken + ", lobby=" + lobbyEnabled + ", event=" + isEventConversation + ").");

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

            if (!string.IsNullOrEmpty(entryId))
            {
                _subscriptionByEntryId[entryId] = subscription;
            }
        }

        private void UnregisterSubscription(string key, string roomToken, string entryId)
        {
            LogTalk("Removing appointment subscription (token=" + (roomToken ?? "n/a") + ", EntryId=" + (entryId ?? "n/a") + ").");
            if (!string.IsNullOrEmpty(key))
            {
                _activeSubscriptions.Remove(key);
            }

            if (!string.IsNullOrEmpty(roomToken))
            {
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

        // Returns true when participant sync completed without runtime/service failures.
        // This status is used for deterministic pre-delegation logging in OnWrite.
        private bool TrySyncRoomParticipants(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken) || _currentSettings == null)
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Participant sync skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            bool addParticipantsLegacy = GetUserPropertyBool(appointment, IcalAddParticipants);
            bool hasAddUsers = HasUserProperty(appointment, IcalAddUsers) || HasUserProperty(appointment, PropertyAddUsers);
            bool hasAddGuests = HasUserProperty(appointment, IcalAddGuests) || HasUserProperty(appointment, PropertyAddGuests);

            bool addUsers = hasAddUsers ? GetUserPropertyBoolPrefer(appointment, IcalAddUsers, PropertyAddUsers) : addParticipantsLegacy;
            bool addGuests = hasAddGuests ? GetUserPropertyBoolPrefer(appointment, IcalAddGuests, PropertyAddGuests) : addParticipantsLegacy;
            if (!addUsers && !addGuests)
            {
                return true;
            }

            var configuration = new TalkServiceConfiguration(_currentSettings.ServerUrl, _currentSettings.Username, _currentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                LogTalk("Participant sync failed: talk service configuration incomplete.");
                return false;
            }

            List<string> attendeeEmails = GetAppointmentAttendeeEmails(appointment);
            if (attendeeEmails.Count == 0)
            {
                return true;
            }

            var cache = new IfbAddressBookCache(_settingsStorage != null ? _settingsStorage.DataDirectory : null);
            string selfEmail;
            cache.TryGetPrimaryEmailForUid(configuration, _currentSettings.IfbCacheHours, _currentSettings.Username, out selfEmail);

            int userAdds = 0;
            int guestAdds = 0;
            int skipped = 0;
            bool hadFailures = false;

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
                        else
                        {
                            hadFailures = true;
                            LogTalk("Participant sync failed while adding Nextcloud user (uid=" + uid + ", token=" + roomToken + ").");
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
                    else
                    {
                        hadFailures = true;
                        LogTalk("Participant sync failed while adding guest (email=" + email + ", token=" + roomToken + ").");
                    }
                }
            }
            catch (TalkServiceException ex)
            {
                hadFailures = true;
                LogTalk("Participant sync failed: " + ex.Message);
            }
            catch (Exception ex)
            {
                hadFailures = true;
                LogTalk("Unexpected error during participant sync: " + ex.Message);
            }

            LogTalk("Participant sync completed (users=" + userAdds + ", guests=" + guestAdds + ", skipped=" + skipped + ", failed=" + hadFailures + ", token=" + roomToken + ").");
            return !hadFailures;
        }

        private void TryApplyDelegation(Outlook.AppointmentItem appointment, string roomToken)
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
                LogTalk("Delegation ignored (delegate == current user).");
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
                LogTalk("Starting delegation (token=" + roomToken + ", delegate=" + delegateId + ").");

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
                    LogTalk("Delegation failed: attendeeId not found (delegate=" + delegateId + ").");
                    return;
                }

                string promoteError;
                if (!service.PromoteModerator(roomToken, attendeeId, out promoteError))
                {
                    LogTalk("Delegation failed: moderator could not be assigned (" + promoteError + ").");
                    string warning = string.IsNullOrWhiteSpace(promoteError)
                        ? Strings.WarningModeratorTransferFailed
                        : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, promoteError);
                    ShowWarning(warning);
                    return;
                }

                bool left = service.LeaveRoom(roomToken);
                LogTalk("Delegation completed (token=" + roomToken + ", delegate=" + delegateId + ", leftSelf=" + left + ").");

                SetUserProperty(appointment, PropertyDelegated, Outlook.OlUserPropertyType.olYesNo, true);
                SetUserProperty(appointment, IcalDelegated, Outlook.OlUserPropertyType.olText, "TRUE");
                RemoveUserProperty(appointment, IcalDelegateReady);
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Delegation failed: " + ex.Message);
                string message = ex.Message;
                ShowWarning(string.IsNullOrWhiteSpace(message)
                    ? Strings.WarningModeratorTransferFailed
                    : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, message));
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error during delegation: " + ex.Message);
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
                LogTalk("Deleting room (token=" + roomToken + ", event=" + isEventConversation + ").");
                var service = CreateTalkService();
                service.DeleteRoom(roomToken, isEventConversation);
                LogTalk("Room deleted successfully (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                LogTalk("Room could not be deleted: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningRoomDeleteFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error while deleting room: " + ex.Message);
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
                LogTalk("Lobby update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            try
            {
                var service = CreateTalkService();
                long startEpoch;
                if (!TryReadRequiredIcalStartEpoch(appointment, roomToken, out startEpoch))
                {
                    return false;
                }

                DateTime startUtc;
                try
                {
                    startUtc = DateTimeOffset.FromUnixTimeSeconds(startEpoch).UtcDateTime;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Lobby update blocked: X-NCTALK-START out of range (token=" + roomToken + ", startEpoch=" + startEpoch.ToString(CultureInfo.InvariantCulture) + ").", ex);
                    return false;
                }

                DateTime end;
                try
                {
                    end = appointment.End;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Appointment.End during lobby update (token=" + roomToken + ").", ex);
                    return false;
                }

                if (end == DateTime.MinValue)
                {
                    end = startUtc;
                }

                LogTalk("Updating lobby (token=" + roomToken + ", startEpoch=" + startEpoch.ToString(CultureInfo.InvariantCulture) + ", startUtc=" + startUtc.ToString("o") + ", end=" + end.ToString("o") + ", event=" + isEventConversation + ").");
                service.UpdateLobby(roomToken, startUtc, end, isEventConversation);
                LogTalk("Lobby updated successfully (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    LogTalk("Lobby update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Lobby could not be updated: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error while updating lobby: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }

            return false;
        }

        private bool TryUpdateRoomName(Outlook.AppointmentItem appointment, string roomToken)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                LogTalk("Room name update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            string roomName = GetNormalizedRoomName(appointment);
            if (string.IsNullOrWhiteSpace(roomName))
            {
                LogTalk("Room name update skipped (empty subject, token=" + roomToken + ").");
                return true;
            }

            try
            {
                var service = CreateTalkService();
                LogTalk("Updating room name (token=" + roomToken + ", length=" + roomName.Length + ").");
                service.UpdateRoomName(roomToken, roomName);
                LogTalk("Room name updated (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    LogTalk("Room name update skipped for event conversation (token=" + roomToken + ").");
                    PersistEventConversationTraits(appointment, roomToken);
                    return true;
                }

                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    LogTalk("Room name update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Room name could not be updated: " + ex.Message);
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error while updating room name: " + ex.Message);
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
                LogTalk("Description update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
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
                LogTalk("Updating room description (token=" + roomToken + ", event=" + isEventConversation + ", textLength=" + description.Length + ").");
                service.UpdateDescription(roomToken, description, isEventConversation);
                LogTalk("Room description updated (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    LogTalk("Room description update skipped after event-conversation response (token=" + roomToken + ").");
                    PersistEventConversationTraits(appointment, roomToken);
                    return true;
                }

                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    LogTalk("Room description update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                LogTalk("Room description could not be updated: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningDescriptionUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error while updating room description: " + ex.Message);
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

        private sealed class MailComposeSubscription : IDisposable
        {
            private sealed class AttachmentBatchEntry
            {
                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentBatchInfo
            {
                internal int Count { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentSnapshot
            {
                internal int Index { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentAutomationSettings
            {
                internal bool AlwaysConnector { get; set; }

                internal bool OfferAboveEnabled { get; set; }

                internal int ThresholdMb { get; set; }

                internal long ThresholdBytes { get; set; }
            }

            private sealed class ComposeShareCleanupEntry
            {
                internal string RelativeFolder { get; set; }

                internal string ShareId { get; set; }

                internal string ShareLabel { get; set; }

                internal DateTime CreatedUtc { get; set; }
            }

            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.MailItem _mail;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly string _mailIdentityKey;
            private readonly string _inspectorIdentityKey;
            private readonly string _composeKey;
            private readonly System.Windows.Forms.Timer _attachmentEvalTimer = new System.Windows.Forms.Timer();
            private readonly System.Windows.Forms.Timer _cleanupGraceTimer = new System.Windows.Forms.Timer();
            private readonly List<AttachmentBatchEntry> _pendingAddedBatch = new List<AttachmentBatchEntry>();
            private readonly List<ComposeShareCleanupEntry> _cleanupEntries = new List<ComposeShareCleanupEntry>();
            private readonly List<SeparatePasswordDispatchEntry> _passwordDispatchQueue = new List<SeparatePasswordDispatchEntry>();
            private bool _attachmentSuppressed;
            private bool _attachmentPromptOpen;
            private bool _sendPending;
            private DateTime _sendPendingAtUtc;
            private bool _awaitingGraceCloseResolution;
            private bool _disposed;

            internal MailComposeSubscription(NextcloudTalkAddIn owner, Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                _owner = owner;
                _mail = mail;
                _mailIdentityKey = string.IsNullOrWhiteSpace(mailIdentityKey)
                    ? ResolveMailComIdentityKey(mail)
                    : mailIdentityKey.Trim();
                _inspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityKey)
                    ? ResolveMailInspectorIdentityKey(mail)
                    : inspectorIdentityKey.Trim();
                _composeKey = BuildComposeKey(mail, _mailIdentityKey, _inspectorIdentityKey);

                _attachmentEvalTimer.Interval = ComposeAttachmentEvalDebounceMs;
                _attachmentEvalTimer.Tick += OnAttachmentEvalTimerTick;

                _cleanupGraceTimer.Interval = ComposeShareCleanupSendGraceMs;
                _cleanupGraceTimer.Tick += OnCleanupGraceTimerTick;

                _events = mail as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.AttachmentAdd += OnAttachmentAdd;
                    _events.PropertyChange += OnPropertyChange;
                    _events.Send += OnSend;
                    _events.Close += OnClose;
                }

                LogFileLink(
                    "Compose subscription registered (composeKey="
                    + _composeKey
                    + ", mailIdentity="
                    + (_mailIdentityKey ?? string.Empty)
                    + ", inspectorIdentity="
                    + (_inspectorIdentityKey ?? string.Empty)
                    + ").");
            }

            internal bool IsFor(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                if (mail == null)
                {
                    return false;
                }

                if (ReferenceEquals(mail, _mail) || mail == _mail)
                {
                    return true;
                }

                if (string.IsNullOrWhiteSpace(_mailIdentityKey))
                {
                    if (string.IsNullOrWhiteSpace(_inspectorIdentityKey))
                    {
                        return false;
                    }
                }

                if (!string.IsNullOrWhiteSpace(_mailIdentityKey)
                    && string.Equals(_mailIdentityKey, mailIdentityKey ?? string.Empty, StringComparison.Ordinal))
                {
                    return true;
                }

                return !string.IsNullOrWhiteSpace(_inspectorIdentityKey)
                    && string.Equals(_inspectorIdentityKey, inspectorIdentityKey ?? string.Empty, StringComparison.Ordinal);
            }

            internal void ArmShareCleanup(FileLinkResult result)
            {
                if (result == null)
                {
                    return;
                }

                string relativeFolder = string.IsNullOrWhiteSpace(result.RelativePath)
                    ? string.Empty
                    : result.RelativePath.Trim();
                if (string.IsNullOrWhiteSpace(relativeFolder))
                {
                    LogFileLink("Compose share cleanup arm skipped (composeKey=" + _composeKey + ", reason=missing_relative_folder).");
                    return;
                }

                _cleanupGraceTimer.Stop();
                _awaitingGraceCloseResolution = false;
                _sendPending = false;

                _cleanupEntries.Add(new ComposeShareCleanupEntry
                {
                    RelativeFolder = relativeFolder,
                    ShareId = result.ShareToken ?? string.Empty,
                    ShareLabel = result.FolderName ?? string.Empty,
                    CreatedUtc = DateTime.UtcNow
                });

                LogFileLink(
                    "Compose share cleanup armed (composeKey="
                    + _composeKey
                    + ", relativeFolder="
                    + relativeFolder
                    + ", shareId="
                    + (result.ShareToken ?? string.Empty)
                    + ", shareLabel="
                    + (result.FolderName ?? string.Empty)
                    + ", armedCount="
                    + _cleanupEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            internal void RegisterSeparatePasswordDispatch(FileLinkResult result, FileLinkRequest request, string passwordOnlyHtml)
            {
                if (result == null || request == null || string.IsNullOrWhiteSpace(passwordOnlyHtml))
                {
                    return;
                }

                string password = result.Password ?? string.Empty;
                if (string.IsNullOrWhiteSpace(password))
                {
                    return;
                }

                var entry = new SeparatePasswordDispatchEntry
                {
                    ShareLabel = result.FolderName ?? string.Empty,
                    ShareUrl = result.ShareUrl ?? string.Empty,
                    Password = password.Trim(),
                    Html = passwordOnlyHtml,
                    To = BuildNormalizedRecipientCsv(ReadMailRecipientList("To")),
                    Cc = BuildNormalizedRecipientCsv(ReadMailRecipientList("CC")),
                    Bcc = BuildNormalizedRecipientCsv(ReadMailRecipientList("BCC"))
                };

                _passwordDispatchQueue.Add(entry);
                LogFileLink(
                    "Separate password dispatch registered (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", hasShareUrl="
                    + (!string.IsNullOrWhiteSpace(entry.ShareUrl)).ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private static string BuildComposeKey(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                if (mail != null)
                {
                    try
                    {
                        string entryId = mail.EntryID;
                        if (!string.IsNullOrWhiteSpace(entryId))
                        {
                            return entryId.Trim();
                        }
                    }
                    catch (COMException ex)
                    {
                        uint errorCode = unchecked((uint)ex.ErrorCode);
                        if ((errorCode & 0xFFFFu) == 0x0108u)
                        {
                            LogFileLink(
                                "MailItem.EntryID unavailable while building compose key (hresult=0x"
                                + errorCode.ToString("X8", CultureInfo.InvariantCulture)
                                + ").");
                        }
                        else
                        {
                            DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.EntryID for compose key.", ex);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.EntryID for compose key.", ex);
                    }
                }

                if (!string.IsNullOrWhiteSpace(mailIdentityKey))
                {
                    return mailIdentityKey.Trim();
                }

                if (!string.IsNullOrWhiteSpace(inspectorIdentityKey))
                {
                    return inspectorIdentityKey.Trim();
                }

                return Guid.NewGuid().ToString("N");
            }

            private void OnAttachmentAdd(Outlook.Attachment attachment)
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                _pendingAddedBatch.Add(new AttachmentBatchEntry
                {
                    Name = ReadAttachmentName(attachment),
                    SizeBytes = ReadAttachmentSizeBytes(attachment)
                });

                LogFileLink(
                    "Compose attachment added (composeKey="
                    + _composeKey
                    + ", pendingBatchCount="
                    + _pendingAddedBatch.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                ScheduleAttachmentEvaluation();
            }

            private void OnPropertyChange(string name)
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                string propertyName = string.IsNullOrWhiteSpace(name) ? string.Empty : name.Trim();
                if (propertyName.IndexOf("Attach", StringComparison.OrdinalIgnoreCase) < 0
                    && !string.Equals(propertyName, "HasAttachment", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                LogFileLink(
                    "Compose property changed (composeKey="
                    + _composeKey
                    + ", property="
                    + propertyName
                    + ").");
                ScheduleAttachmentEvaluation();
            }

            private void ScheduleAttachmentEvaluation()
            {
                if (_disposed)
                {
                    return;
                }

                try
                {
                    _attachmentEvalTimer.Stop();
                    _attachmentEvalTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to schedule compose attachment evaluation (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void OnAttachmentEvalTimerTick(object sender, EventArgs e)
            {
                _attachmentEvalTimer.Stop();

                try
                {
                    EvaluateAttachmentAutomation();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Compose attachment evaluation failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void EvaluateAttachmentAutomation()
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("evaluate", _composeKey, out guardState))
                {
                    _pendingAddedBatch.Clear();
                    return;
                }

                AttachmentAutomationSettings settings = ReadAttachmentAutomationSettings();
                if (!settings.AlwaysConnector && !settings.OfferAboveEnabled)
                {
                    _pendingAddedBatch.Clear();
                    LogFileLink("Compose attachment evaluation skipped (composeKey=" + _composeKey + ", reason=automation_disabled).");
                    return;
                }

                List<AttachmentSnapshot> attachments = SnapshotAttachments();
                if (attachments.Count == 0)
                {
                    _pendingAddedBatch.Clear();
                    LogFileLink("Compose attachment evaluation skipped (composeKey=" + _composeKey + ", reason=no_attachments).");
                    return;
                }

                long totalBytes = SumAttachmentBytes(attachments);
                AttachmentBatchInfo lastAdded = BuildLastAddedBatchInfo(attachments);

                LogFileLink(
                    "Compose attachment evaluation (composeKey="
                    + _composeKey
                    + ", attachmentCount="
                    + attachments.Count.ToString(CultureInfo.InvariantCulture)
                    + ", totalBytes="
                    + totalBytes.ToString(CultureInfo.InvariantCulture)
                    + ", lastAddedCount="
                    + lastAdded.Count.ToString(CultureInfo.InvariantCulture)
                    + ", alwaysConnector="
                    + settings.AlwaysConnector.ToString(CultureInfo.InvariantCulture)
                    + ", offerAboveEnabled="
                    + settings.OfferAboveEnabled.ToString(CultureInfo.InvariantCulture)
                    + ", thresholdBytes="
                    + settings.ThresholdBytes.ToString(CultureInfo.InvariantCulture)
                    + ").");

                if (settings.AlwaysConnector)
                {
                    StartComposeAttachmentShareFlow("always", totalBytes, settings.ThresholdMb, lastAdded);
                    return;
                }

                if (!settings.OfferAboveEnabled || totalBytes <= settings.ThresholdBytes)
                {
                    return;
                }

                string reasonText = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.AttachmentPromptReason,
                    FormatSizeMb(totalBytes),
                    FormatSizeMb((long)settings.ThresholdMb * 1024L * 1024L),
                    string.IsNullOrWhiteSpace(lastAdded.Name) ? Strings.AttachmentPromptLastUnknown : lastAdded.Name,
                    FormatSizeMb(lastAdded.SizeBytes));
                if (_attachmentPromptOpen)
                {
                    LogFileLink("Compose attachment prompt skipped (composeKey=" + _composeKey + ", reason=prompt_already_open).");
                    return;
                }

                _attachmentPromptOpen = true;
                ComposeAttachmentPromptDecision decision;
                try
                {
                    decision = ComposeAttachmentPromptForm.ShowPrompt(
                        _owner.TryCreateMailInspectorDialogOwner(_mail),
                        reasonText);
                }
                finally
                {
                    _attachmentPromptOpen = false;
                }

                if (_owner.TryGetAttachmentAutomationGuardState("prompt_action", _composeKey, out guardState))
                {
                    return;
                }

                if (decision == ComposeAttachmentPromptDecision.Share)
                {
                    LogFileLink(
                        "Compose attachment threshold decision (composeKey="
                        + _composeKey
                        + ", decision=share, totalBytes="
                        + totalBytes.ToString(CultureInfo.InvariantCulture)
                        + ", thresholdBytes="
                        + settings.ThresholdBytes.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    StartComposeAttachmentShareFlow("threshold", totalBytes, settings.ThresholdMb, lastAdded);
                    return;
                }

                LogFileLink(
                    "Compose attachment threshold decision (composeKey="
                    + _composeKey
                    + ", decision=remove_last, removeCount="
                    + lastAdded.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                RemoveLastAddedAttachmentBatch(lastAdded);
            }

            private AttachmentAutomationSettings ReadAttachmentAutomationSettings()
            {
                _owner.EnsureSettingsLoaded();

                var settings = _owner._currentSettings ?? new AddinSettings();
                int thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(settings.SharingAttachmentsOfferAboveMb);
                bool alwaysConnector = settings.SharingAttachmentsAlwaysConnector;
                bool offerAboveEnabled = settings.SharingAttachmentsOfferAboveEnabled && !alwaysConnector;

                return new AttachmentAutomationSettings
                {
                    AlwaysConnector = alwaysConnector,
                    OfferAboveEnabled = offerAboveEnabled,
                    ThresholdMb = thresholdMb,
                    ThresholdBytes = (long)thresholdMb * 1024L * 1024L
                };
            }

            private List<AttachmentSnapshot> SnapshotAttachments()
            {
                var snapshots = new List<AttachmentSnapshot>();
                Outlook.Attachments attachments = null;

                try
                {
                    attachments = _mail.Attachments;
                    if (attachments == null)
                    {
                        return snapshots;
                    }

                    int count = attachments.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            if (attachment == null)
                            {
                                continue;
                            }

                            snapshots.Add(new AttachmentSnapshot
                            {
                                Index = index,
                                Name = ReadAttachmentName(attachment),
                                SizeBytes = ReadAttachmentSizeBytes(attachment)
                            });
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.FileLink,
                                "Failed to snapshot compose attachment (composeKey="
                                + _composeKey
                                + ", index="
                                + index.ToString(CultureInfo.InvariantCulture)
                                + ").",
                                ex);
                        }
                        finally
                        {
                            ReleaseComObject(attachment, "compose attachment snapshot");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose attachments (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ReleaseComObject(attachments, "compose attachments collection snapshot");
                }

                return snapshots;
            }

            private static long SumAttachmentBytes(List<AttachmentSnapshot> snapshots)
            {
                long total = 0;
                if (snapshots == null)
                {
                    return 0;
                }

                for (int i = 0; i < snapshots.Count; i++)
                {
                    total += Math.Max(0, snapshots[i].SizeBytes);
                }

                return total;
            }

            private AttachmentBatchInfo BuildLastAddedBatchInfo(List<AttachmentSnapshot> snapshots)
            {
                if (_pendingAddedBatch.Count > 0)
                {
                    long total = 0;
                    for (int i = 0; i < _pendingAddedBatch.Count; i++)
                    {
                        total += Math.Max(0, _pendingAddedBatch[i].SizeBytes);
                    }

                    var info = new AttachmentBatchInfo
                    {
                        Count = _pendingAddedBatch.Count,
                        Name = _pendingAddedBatch[_pendingAddedBatch.Count - 1].Name ?? string.Empty,
                        SizeBytes = total
                    };
                    _pendingAddedBatch.Clear();
                    return info;
                }

                if (snapshots == null || snapshots.Count == 0)
                {
                    return new AttachmentBatchInfo
                    {
                        Count = 0,
                        Name = string.Empty,
                        SizeBytes = 0
                    };
                }

                AttachmentSnapshot latest = snapshots[snapshots.Count - 1];
                return new AttachmentBatchInfo
                {
                    Count = 1,
                    Name = latest.Name ?? string.Empty,
                    SizeBytes = Math.Max(0, latest.SizeBytes)
                };
            }

            private void StartComposeAttachmentShareFlow(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            {
                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("start_flow", _composeKey, out guardState))
                {
                    return;
                }

                var selections = new List<FileLinkSelection>();
                var removeIndices = new List<int>();
                var tempFiles = new List<string>();

                CollectAttachmentSelectionsForShare(selections, removeIndices, tempFiles);
                if (selections.Count == 0)
                {
                    CleanupTemporaryFiles(tempFiles);
                    LogFileLink("Compose attachment flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                    return;
                }

                RemoveAttachmentsByIndices(removeIndices, "share_flow");

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "always" : trigger,
                    AttachmentTotalBytes = Math.Max(0, totalBytes),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = lastAdded != null ? (lastAdded.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = lastAdded != null ? Math.Max(0, lastAdded.SizeBytes) : 0
                };
                for (int i = 0; i < selections.Count; i++)
                {
                    launchOptions.InitialSelections.Add(new FileLinkSelection(selections[i].SelectionType, selections[i].LocalPath));
                }

                try
                {
                    bool wizardAccepted = _owner.RunFileLinkWizardForMail(_mail, launchOptions);
                    LogFileLink(
                        "Compose attachment flow completed (composeKey="
                        + _composeKey
                        + ", trigger="
                        + launchOptions.AttachmentTrigger
                        + ", queued="
                        + selections.Count.ToString(CultureInfo.InvariantCulture)
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                finally
                {
                    CleanupTemporaryFiles(tempFiles);
                }
            }

            private void CollectAttachmentSelectionsForShare(List<FileLinkSelection> selections, List<int> removeIndices, List<string> temporaryFiles)
            {
                Outlook.Attachments attachments = null;
                try
                {
                    attachments = _mail.Attachments;
                    if (attachments == null)
                    {
                        return;
                    }

                    int count = attachments.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            if (attachment == null)
                            {
                                continue;
                            }

                            string attachmentName = ReadAttachmentName(attachment);
                            string localPath;
                            if (!TryResolveAttachmentLocalPath(attachment, attachmentName, temporaryFiles, out localPath))
                            {
                                continue;
                            }

                            selections.Add(new FileLinkSelection(FileLinkSelectionType.File, localPath));
                            removeIndices.Add(index);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.FileLink,
                                "Failed to collect compose attachment for sharing (composeKey="
                                + _composeKey
                                + ", index="
                                + index.ToString(CultureInfo.InvariantCulture)
                                + ").",
                                ex);
                        }
                        finally
                        {
                            ReleaseComObject(attachment, "compose attachment collect");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to collect compose attachments for sharing (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ReleaseComObject(attachments, "compose attachments collection collect");
                }
            }

            private bool TryResolveAttachmentLocalPath(Outlook.Attachment attachment, string attachmentName, List<string> temporaryFiles, out string localPath)
            {
                localPath = string.Empty;
                if (attachment == null)
                {
                    return false;
                }

                string pathName = ReadAttachmentPathName(attachment);
                if (!string.IsNullOrWhiteSpace(pathName) && File.Exists(pathName))
                {
                    localPath = pathName;
                    return true;
                }

                string safeName = FileLinkService.SanitizeComponent(attachmentName);
                if (string.IsNullOrWhiteSpace(safeName))
                {
                    safeName = "attachment.bin";
                }

                string tempRoot = Path.Combine(Path.GetTempPath(), "NCConnectorOutlook", "Attachments", _composeKey);
                try
                {
                    Directory.CreateDirectory(tempRoot);
                    string targetPath = BuildUniqueFilePath(tempRoot, safeName);
                    attachment.SaveAsFile(targetPath);
                    if (!File.Exists(targetPath))
                    {
                        return false;
                    }

                    localPath = targetPath;
                    if (temporaryFiles != null)
                    {
                        temporaryFiles.Add(targetPath);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to materialize compose attachment file (composeKey="
                        + _composeKey
                        + ", attachment="
                        + safeName
                        + ").",
                        ex);
                    return false;
                }
            }

            private static string BuildUniqueFilePath(string directory, string fileName)
            {
                string candidate = Path.Combine(directory, fileName);
                if (!File.Exists(candidate))
                {
                    return candidate;
                }

                for (int suffix = 1; suffix < 1000; suffix++)
                {
                    string slotDirectory = Path.Combine(
                        directory,
                        "dup_" + suffix.ToString(CultureInfo.InvariantCulture));
                    candidate = Path.Combine(slotDirectory, fileName);
                    if (!File.Exists(candidate))
                    {
                        Directory.CreateDirectory(slotDirectory);
                        return candidate;
                    }
                }

                string fallbackDirectory = Path.Combine(directory, Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(fallbackDirectory);
                return Path.Combine(fallbackDirectory, fileName);
            }

            private void RemoveAttachmentsByIndices(List<int> indices, string reason)
            {
                if (indices == null || indices.Count == 0)
                {
                    return;
                }

                indices.Sort();

                Outlook.Attachments attachments = null;
                int removed = 0;
                _attachmentSuppressed = true;
                _pendingAddedBatch.Clear();

                try
                {
                    attachments = _mail.Attachments;
                    if (attachments == null)
                    {
                        return;
                    }

                    for (int index = indices.Count - 1; index >= 0; index--)
                    {
                        int attachmentIndex = indices[index];
                        if (attachmentIndex <= 0 || attachmentIndex > attachments.Count)
                        {
                            continue;
                        }

                        attachments.Remove(attachmentIndex);
                        removed++;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to remove compose attachments (composeKey="
                        + _composeKey
                        + ", reason="
                        + (reason ?? string.Empty)
                        + ").",
                        ex);
                }
                finally
                {
                    ReleaseComObject(attachments, "compose attachments collection remove");
                    _attachmentSuppressed = false;
                }

                LogFileLink(
                    "Compose attachments removed (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", requested="
                    + indices.Count.ToString(CultureInfo.InvariantCulture)
                    + ", removed="
                    + removed.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void RemoveLastAddedAttachmentBatch(AttachmentBatchInfo lastAdded)
            {
                int removeCount = lastAdded != null ? Math.Max(1, lastAdded.Count) : 1;

                Outlook.Attachments attachments = null;
                int removed = 0;
                _attachmentSuppressed = true;
                _pendingAddedBatch.Clear();

                try
                {
                    attachments = _mail.Attachments;
                    if (attachments == null || attachments.Count <= 0)
                    {
                        return;
                    }

                    int totalCount = attachments.Count;
                    int effectiveCount = Math.Min(totalCount, removeCount);
                    for (int i = 0; i < effectiveCount; i++)
                    {
                        attachments.Remove(attachments.Count);
                        removed++;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to remove last attachment batch (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ReleaseComObject(attachments, "compose attachments collection remove_last");
                    _attachmentSuppressed = false;
                }

                LogFileLink(
                    "Compose attachment batch removed (composeKey="
                    + _composeKey
                    + ", requested="
                    + removeCount.ToString(CultureInfo.InvariantCulture)
                    + ", removed="
                    + removed.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void CleanupTemporaryFiles(List<string> temporaryFiles)
            {
                if (temporaryFiles == null || temporaryFiles.Count == 0)
                {
                    return;
                }

                for (int i = 0; i < temporaryFiles.Count; i++)
                {
                    string path = temporaryFiles[i];
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    try
                    {
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to delete temporary compose attachment file '" + path + "'.",
                            ex);
                    }
                }
            }

            private void OnSend(ref bool cancel)
            {
                if (_disposed)
                {
                    return;
                }

                if (cancel)
                {
                    LogFileLink("Compose send cancelled before dispatch handling (composeKey=" + _composeKey + ").");
                    return;
                }

                _sendPending = true;
                _sendPendingAtUtc = DateTime.UtcNow;
                _cleanupGraceTimer.Stop();
                _awaitingGraceCloseResolution = false;

                CapturePasswordDispatchRecipients();
                LogFileLink(
                    "Compose send state updated (composeKey="
                    + _composeKey
                    + ", sendPending="
                    + _sendPending.ToString(CultureInfo.InvariantCulture)
                    + ", cleanupArmedCount="
                    + _cleanupEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void OnClose(ref bool cancel)
            {
                if (_disposed)
                {
                    return;
                }

                ComposeSendState sendState = EvaluateMailSendState();
                bool hasPendingPostSendWork = _cleanupEntries.Count > 0 || _passwordDispatchQueue.Count > 0;
                int delayMs = 0;
                if (_sendPending)
                {
                    double elapsedMs = (DateTime.UtcNow - _sendPendingAtUtc).TotalMilliseconds;
                    delayMs = (int)Math.Max(0, ComposeShareCleanupSendGraceMs - elapsedMs);
                }

                if (sendState == ComposeSendState.Sent)
                {
                    ClearShareCleanupEntries("after_send_success");
                    DispatchSeparatePasswordQueue("after_send_success");
                    Dispose();
                    return;
                }

                if (sendState == ComposeSendState.UnavailableAfterSend)
                {
                    if (hasPendingPostSendWork && delayMs > 0)
                    {
                        _awaitingGraceCloseResolution = true;
                        _cleanupGraceTimer.Interval = Math.Max(250, delayMs);
                        _cleanupGraceTimer.Start();

                        LogFileLink(
                            "Compose share cleanup delayed (composeKey="
                            + _composeKey
                            + ", delayMs="
                            + delayMs.ToString(CultureInfo.InvariantCulture)
                            + ", reason=close_send_state_unavailable).");
                        return;
                    }

                    ClearShareCleanupEntries("after_send_state_unavailable");
                    DispatchSeparatePasswordQueue("after_send_state_unavailable");
                    Dispose();
                    return;
                }

                if (hasPendingPostSendWork && _sendPending && delayMs > 0)
                {
                    _awaitingGraceCloseResolution = true;
                    _cleanupGraceTimer.Interval = Math.Max(250, delayMs);
                    _cleanupGraceTimer.Start();

                    LogFileLink(
                        "Compose share cleanup delayed (composeKey="
                        + _composeKey
                        + ", delayMs="
                        + delayMs.ToString(CultureInfo.InvariantCulture)
                        + ", reason=close_send_pending).");
                    return;
                }

                if (hasPendingPostSendWork && _sendPending)
                {
                    LogFileLink(
                        "Compose send state not confirmed after grace; applying unsent cleanup path (composeKey="
                        + _composeKey
                        + ", reason=close_send_pending_timeout).");
                    DeleteShareCleanupEntries("close_send_pending_timeout_without_successful_send");
                    ClearSeparatePasswordDispatchQueue("close_send_pending_timeout_without_successful_send");
                    Dispose();
                    return;
                }

                if (_cleanupEntries.Count > 0)
                {
                    DeleteShareCleanupEntries("close_without_successful_send");
                }

                if (_passwordDispatchQueue.Count > 0)
                {
                    ClearSeparatePasswordDispatchQueue("close_without_successful_send");
                }

                Dispose();
            }

            private void OnCleanupGraceTimerTick(object sender, EventArgs e)
            {
                _cleanupGraceTimer.Stop();

                ComposeSendState sendState = EvaluateMailSendState();
                if (sendState == ComposeSendState.Sent || sendState == ComposeSendState.UnavailableAfterSend)
                {
                    string reason = sendState == ComposeSendState.Sent
                        ? "delayed_after_send_success"
                        : "delayed_after_send_state_unavailable";
                    ClearShareCleanupEntries(reason);
                    DispatchSeparatePasswordQueue(reason);
                }
                else
                {
                    if (_sendPending)
                    {
                        LogFileLink(
                            "Compose send state not confirmed after delayed grace; applying unsent cleanup path (composeKey="
                            + _composeKey
                            + ", reason=delayed_send_pending_timeout).");
                        DeleteShareCleanupEntries("delayed_send_pending_timeout_without_successful_send");
                        ClearSeparatePasswordDispatchQueue("delayed_send_pending_timeout_without_successful_send");
                    }
                    else
                    {
                        DeleteShareCleanupEntries("delayed_close_without_successful_send");
                        ClearSeparatePasswordDispatchQueue("delayed_close_without_successful_send");
                    }
                }

                Dispose();
            }

            private enum ComposeSendState
            {
                NotSent,
                Sent,
                UnavailableAfterSend
            }

            private ComposeSendState EvaluateMailSendState()
            {
                try
                {
                    return _mail != null && _mail.Sent ? ComposeSendState.Sent : ComposeSendState.NotSent;
                }
                catch (Exception ex)
                {
                    if (_sendPending && IsMailSentUnavailableAfterSend(ex))
                    {
                        LogFileLink(
                            "Compose send state unavailable after send (composeKey="
                            + _composeKey
                            + ", hresult="
                            + ToHResultHex(ex)
                            + ").");
                        return ComposeSendState.UnavailableAfterSend;
                    }

                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read MailItem.Sent (composeKey=" + _composeKey + ").",
                        ex);
                    return ComposeSendState.NotSent;
                }
            }

            private static bool IsMailSentUnavailableAfterSend(Exception ex)
            {
                var comException = ex as COMException;
                if (comException == null)
                {
                    return false;
                }

                uint errorCode = unchecked((uint)comException.ErrorCode);
                return (errorCode & 0xFFFFu) == 0x010Au;
            }

            private static string ToHResultHex(Exception ex)
            {
                if (ex == null)
                {
                    return "0x00000000";
                }

                return "0x" + unchecked((uint)ex.HResult).ToString("X8", CultureInfo.InvariantCulture);
            }

            private void ClearShareCleanupEntries(string reason)
            {
                if (_cleanupEntries.Count == 0)
                {
                    return;
                }

                for (int i = 0; i < _cleanupEntries.Count; i++)
                {
                    ComposeShareCleanupEntry entry = _cleanupEntries[i];
                    LogFileLink(
                        "Compose share cleanup cleared (composeKey="
                        + _composeKey
                        + ", reason="
                        + (reason ?? string.Empty)
                        + ", relativeFolder="
                        + (entry.RelativeFolder ?? string.Empty)
                        + ", shareId="
                        + (entry.ShareId ?? string.Empty)
                        + ", shareLabel="
                        + (entry.ShareLabel ?? string.Empty)
                        + ").");
                }

                _cleanupEntries.Clear();
                _sendPending = false;
                _awaitingGraceCloseResolution = false;
            }

            private void DeleteShareCleanupEntries(string reason)
            {
                if (_cleanupEntries.Count == 0)
                {
                    return;
                }

                List<ComposeShareCleanupEntry> entries = new List<ComposeShareCleanupEntry>(_cleanupEntries);
                _cleanupEntries.Clear();
                _sendPending = false;
                _awaitingGraceCloseResolution = false;

                for (int i = 0; i < entries.Count; i++)
                {
                    ComposeShareCleanupEntry entry = entries[i];
                    _owner.TryDeleteComposeShareFolder(
                        entry.RelativeFolder,
                        reason,
                        entry.ShareId,
                        entry.ShareLabel);
                }
            }

            private void ClearSeparatePasswordDispatchQueue(string reason)
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                _passwordDispatchQueue.Clear();
                LogFileLink(
                    "Separate password dispatch cleared (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ").");
            }

            private void CapturePasswordDispatchRecipients()
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                string to;
                string cc;
                string bcc;
                bool capturedFromRecipients = TryCaptureRecipientListsFromRecipientsCollection(out to, out cc, out bcc);
                if (!capturedFromRecipients)
                {
                    to = BuildNormalizedRecipientCsv(ReadMailRecipientList("To"));
                    cc = BuildNormalizedRecipientCsv(ReadMailRecipientList("CC"));
                    bcc = BuildNormalizedRecipientCsv(ReadMailRecipientList("BCC"));
                }

                for (int i = 0; i < _passwordDispatchQueue.Count; i++)
                {
                    _passwordDispatchQueue[i].To = to;
                    _passwordDispatchQueue[i].Cc = cc;
                    _passwordDispatchQueue[i].Bcc = bcc;
                }

                LogFileLink(
                    "Separate password recipients captured (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", to="
                    + CountRecipients(to).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + CountRecipients(cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + CountRecipients(bcc).ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (capturedFromRecipients ? "recipients_collection" : "mail_fields")
                    + ").");
            }

            private bool TryCaptureRecipientListsFromRecipientsCollection(out string to, out string cc, out string bcc)
            {
                to = string.Empty;
                cc = string.Empty;
                bcc = string.Empty;

                if (_mail == null)
                {
                    return false;
                }

                var toRecipients = new List<string>();
                var ccRecipients = new List<string>();
                var bccRecipients = new List<string>();
                Outlook.Recipients recipients = null;
                try
                {
                    recipients = _mail.Recipients;
                    if (recipients == null)
                    {
                        return false;
                    }

                    int count = 0;
                    try
                    {
                        count = recipients.Count;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to read compose Recipients.Count (composeKey=" + _composeKey + ").",
                            ex);
                        count = 0;
                    }

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

                            string address = TryGetRecipientSmtpAddress(recipient);
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                try
                                {
                                    address = NormalizeRecipientAddress(recipient.Address);
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(
                                        LogCategories.FileLink,
                                        "Failed to read compose recipient.Address (composeKey=" + _composeKey + ").",
                                        ex);
                                    address = string.Empty;
                                }
                            }

                            if (string.IsNullOrWhiteSpace(address))
                            {
                                continue;
                            }

                            int recipientType = 1;
                            try
                            {
                                recipientType = recipient.Type;
                            }
                            catch (Exception ex)
                            {
                                DiagnosticsLogger.LogException(
                                    LogCategories.FileLink,
                                    "Failed to read compose recipient.Type (composeKey=" + _composeKey + ").",
                                    ex);
                                recipientType = 1;
                            }

                            if (recipientType == (int)Outlook.OlMailRecipientType.olCC)
                            {
                                AddUniqueRecipient(ccRecipients, address);
                            }
                            else if (recipientType == (int)Outlook.OlMailRecipientType.olBCC)
                            {
                                AddUniqueRecipient(bccRecipients, address);
                            }
                            else
                            {
                                AddUniqueRecipient(toRecipients, address);
                            }
                        }
                        finally
                        {
                            if (recipient != null && Marshal.IsComObject(recipient))
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(recipient);
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release compose Recipient COM object.", ex);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to capture compose recipients from Recipients collection (composeKey=" + _composeKey + ").",
                        ex);
                    return false;
                }
                finally
                {
                    if (recipients != null && Marshal.IsComObject(recipients))
                    {
                        try
                        {
                            Marshal.ReleaseComObject(recipients);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to release compose Recipients COM object.", ex);
                        }
                    }
                }

                to = toRecipients.Count == 0 ? string.Empty : string.Join("; ", toRecipients.ToArray());
                cc = ccRecipients.Count == 0 ? string.Empty : string.Join("; ", ccRecipients.ToArray());
                bcc = bccRecipients.Count == 0 ? string.Empty : string.Join("; ", bccRecipients.ToArray());
                return toRecipients.Count + ccRecipients.Count + bccRecipients.Count > 0;
            }

            private void DispatchSeparatePasswordQueue(string reason)
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                var queue = new List<SeparatePasswordDispatchEntry>(_passwordDispatchQueue);
                _passwordDispatchQueue.Clear();
                LogFileLink(
                    "Separate password dispatch taken (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", queued="
                    + queue.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                _owner.DispatchSeparatePasswordMailQueue(_composeKey, queue);
            }

            private string ReadMailRecipientList(string fieldName)
            {
                try
                {
                    if (_mail == null)
                    {
                        return string.Empty;
                    }

                    switch ((fieldName ?? string.Empty).Trim().ToUpperInvariant())
                    {
                        case "TO":
                            return _mail.To ?? string.Empty;
                        case "CC":
                            return _mail.CC ?? string.Empty;
                        case "BCC":
                            return _mail.BCC ?? string.Empty;
                        default:
                            return string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose recipient field '" + (fieldName ?? string.Empty) + "' (composeKey=" + _composeKey + ").",
                        ex);
                    return string.Empty;
                }
            }

            private static int CountRecipients(string csv)
            {
                return CountRecipientsInCsv(csv);
            }

            private static string FormatSizeMb(long bytes)
            {
                decimal value = Math.Max(0, bytes) / (1024m * 1024m);
                return string.Format(CultureInfo.CurrentCulture, "{0:0.0} MB", value);
            }

            private static string ReadAttachmentName(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return string.Empty;
                }

                try
                {
                    string fileName = attachment.FileName;
                    if (!string.IsNullOrWhiteSpace(fileName))
                    {
                        return fileName.Trim();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.FileName.", ex);
                }

                try
                {
                    string displayName = attachment.DisplayName;
                    if (!string.IsNullOrWhiteSpace(displayName))
                    {
                        return displayName.Trim();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.DisplayName.", ex);
                }

                return string.Empty;
            }

            private static long ReadAttachmentSizeBytes(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return 0;
                }

                try
                {
                    return Math.Max(0, attachment.Size);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.Size.", ex);
                    return 0;
                }
            }

            private static string ReadAttachmentPathName(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return string.Empty;
                }

                try
                {
                    return attachment.PathName ?? string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.PathName.", ex);
                    return string.Empty;
                }
            }

            private static void ReleaseComObject(object comObject, string description)
            {
                if (comObject == null || !Marshal.IsComObject(comObject))
                {
                    return;
                }

                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to release COM object (" + (description ?? string.Empty) + ").",
                        ex);
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _attachmentEvalTimer.Stop();
                _cleanupGraceTimer.Stop();

                try
                {
                    _attachmentEvalTimer.Tick -= OnAttachmentEvalTimerTick;
                    _cleanupGraceTimer.Tick -= OnCleanupGraceTimerTick;
                    _attachmentEvalTimer.Dispose();
                    _cleanupGraceTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose compose timers.", ex);
                }

                if (_events != null)
                {
                    try
                    {
                        _events.AttachmentAdd -= OnAttachmentAdd;
                        _events.PropertyChange -= OnPropertyChange;
                        _events.Send -= OnSend;
                        _events.Close -= OnClose;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to detach compose event handlers.", ex);
                    }
                }

                _owner.RemoveMailComposeSubscription(this);
                _disposed = true;
                LogFileLink(
                    "Compose subscription disposed (composeKey="
                    + _composeKey
                    + ", hadCleanup="
                    + (_cleanupEntries.Count > 0).ToString(CultureInfo.InvariantCulture)
                    + ", hadPasswordDispatch="
                    + (_passwordDispatchQueue.Count > 0).ToString(CultureInfo.InvariantCulture)
                    + ", delayed="
                    + _awaitingGraceCloseResolution.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }
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
            private bool _unsavedCloseCleanupPending;
            private int _unsavedCloseCleanupAttempts;
            private System.Windows.Forms.Timer _unsavedCloseCleanupTimer;
            private string _entryId;
            private const int UnsavedCloseCleanupMaxAttempts = 90;

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
                _lastLobbyTimer = GetIcalStartEpochOrNull(appointment);
                _entryId = entryId;
                _events = appointment as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.BeforeDelete += OnBeforeDelete;
                    _events.Write += OnWrite;
                    _events.Close += OnClose;
                }

                LogTalk("Subscription registered (token=" + _roomToken + ", lobby=" + _lobbyEnabled + ", event=" + _isEventConversation + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }

            private void OnWrite(ref bool cancel)
            {
                if (_unsavedCloseCleanupPending)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("OnWrite canceled pending unsaved cleanup (token=" + _roomToken + ").");
                }

                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("OnWrite ignored (no token).");
                    return;
                }

                long currentStartEpoch = 0;
                bool hasStampedStartEpoch = _owner.TryStampIcalStartEpoch(_appointment, _roomToken, out currentStartEpoch);

                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnWrite ignored (not organizer, token=" + _roomToken + ", stampedStart=" + hasStampedStartEpoch + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                string delegateId;
                if (_owner.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("OnWrite skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                LogTalk("OnWrite for appointment (token=" + _roomToken + ").");

                bool effectiveLobbyKnown;
                bool effectiveLobbyEnabled;
                bool effectiveIsEventConversation;
                _owner.ResolveRuntimeRoomTraits(_appointment, _roomToken, _lobbyEnabled, _isEventConversation, out effectiveLobbyKnown, out effectiveLobbyEnabled, out effectiveIsEventConversation);
                LogTalk("OnWrite traits resolved (token=" + _roomToken + ", lobbyKnown=" + effectiveLobbyKnown + ", lobby=" + effectiveLobbyEnabled + ", event=" + effectiveIsEventConversation + ").");

                string pendingDelegateId;
                bool delegationPending = _owner.IsDelegationPending(_appointment, out pendingDelegateId);
                if (delegationPending)
                {
                    LogTalk("OnWrite delegation-pending path (token=" + _roomToken + ", delegate=" + pendingDelegateId + ").");
                }

                bool roomNameSynced = false;
                bool lobbySynced = true;
                bool descriptionSynced = false;
                bool participantsSynced;

                LogTalk("Updating room name during OnWrite (token=" + _roomToken + ").");
                roomNameSynced = _owner.TryUpdateRoomName(_appointment, _roomToken);

                bool shouldAttemptLobbyUpdate = effectiveLobbyEnabled || !effectiveLobbyKnown;
                if (shouldAttemptLobbyUpdate)
                {
                    if (!hasStampedStartEpoch)
                    {
                        LogTalk("Lobby update skipped: X-NCTALK-START is unavailable after stamping attempt (token=" + _roomToken + ").");
                        lobbySynced = false;
                    }
                    else if (!_lastLobbyTimer.HasValue || currentStartEpoch != _lastLobbyTimer.Value)
                    {
                        LogTalk("Attempting lobby update during OnWrite (token=" + _roomToken + ", startEpoch=" + currentStartEpoch.ToString(CultureInfo.InvariantCulture) + ", lobbyKnown=" + effectiveLobbyKnown + ").");
                        if (_owner.TryUpdateLobby(_appointment, _roomToken, effectiveIsEventConversation))
                        {
                            _lastLobbyTimer = currentStartEpoch;
                            if (!effectiveLobbyKnown)
                            {
                                _owner.PersistLobbyTraits(_appointment, _roomToken, true);
                            }

                            LogTalk("Lobby update successful (token=" + _roomToken + ").");
                        }
                        else
                        {
                            LogTalk("Lobby update failed (token=" + _roomToken + ").");
                            lobbySynced = false;
                        }
                    }
                }

                LogTalk("Updating room description during OnWrite (token=" + _roomToken + ").");
                descriptionSynced = _owner.TryUpdateRoomDescription(_appointment, _roomToken, effectiveIsEventConversation);

                participantsSynced = _owner.TrySyncRoomParticipants(_appointment, _roomToken, effectiveIsEventConversation);
                LogTalk(
                    "OnWrite pre-delegation sync result (token="
                    + _roomToken
                    + ", roomName="
                    + roomNameSynced
                    + ", lobby="
                    + lobbySynced
                    + ", description="
                    + descriptionSynced
                    + ", participants="
                    + participantsSynced
                    + ").");
                if (delegationPending)
                {
                    // Delegation is always executed when pending; pre-step failures are logged with
                    // explicit per-step status to keep runtime behavior transparent.
                    if (!roomNameSynced || !lobbySynced || !descriptionSynced || !participantsSynced)
                    {
                        LogTalk(
                            "Delegation continues despite pre-delegation sync failures (token="
                            + _roomToken
                            + ", roomName="
                            + roomNameSynced
                            + ", lobby="
                            + lobbySynced
                            + ", description="
                            + descriptionSynced
                            + ", participants="
                            + participantsSynced
                            + ").");
                    }

                    _owner.TryApplyDelegation(_appointment, _roomToken);
                }
                _owner.RefreshEntryBinding(this);
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("BeforeDelete ignored (not organizer, token=" + _roomToken + ").");
                    return;
                }

                LogTalk("BeforeDelete -> EnsureRoomDeleted (Token=" + _roomToken + ").");
                EnsureRoomDeleted();
            }

            private void OnClose(ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnClose without organizer (token=" + _roomToken + ").");
                    Dispose();
                    return;
                }

                if (!_roomDeleted && _appointment != null && !_appointment.Saved)
                {
                    LogTalk("OnClose unsaved state observed (token=" + _roomToken + ", cancel=" + cancel + ").");
                    if (!cancel)
                    {
                        ScheduleUnsavedCloseCleanup();
                    }

                    return;
                }

                LogTalk("OnClose completed (token=" + _roomToken + ", deleted=" + _roomDeleted + ").");
            }

            private void ScheduleUnsavedCloseCleanup()
            {
                if (_unsavedCloseCleanupPending)
                {
                    return;
                }

                _unsavedCloseCleanupPending = true;
                _unsavedCloseCleanupAttempts = 0;
                _unsavedCloseCleanupTimer = new System.Windows.Forms.Timer();
                _unsavedCloseCleanupTimer.Interval = 1000;
                _unsavedCloseCleanupTimer.Tick += OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Start();
                LogTalk("OnClose unsaved cleanup scheduled (token=" + _roomToken + ").");
            }

            private void OnUnsavedCloseCleanupTick(object sender, EventArgs e)
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    return;
                }

                _unsavedCloseCleanupAttempts++;
                bool saved;
                bool hasSavedState = TryGetAppointmentSaved(_appointment, out saved);
                bool inspectorOpen = IsAppointmentOpenInAnyInspector(_appointment);
                if (saved)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True).");
                    return;
                }

                if (inspectorOpen)
                {
                    if (_unsavedCloseCleanupAttempts >= UnsavedCloseCleanupMaxAttempts)
                    {
                        _unsavedCloseCleanupPending = false;
                        _unsavedCloseCleanupAttempts = 0;
                        StopUnsavedCloseCleanupTimer();
                        LogTalk("Deferred OnClose cleanup skipped after timeout (token=" + _roomToken + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ", inspectorOpen=True).");
                        return;
                    }

                    if (_unsavedCloseCleanupAttempts == 1 || (_unsavedCloseCleanupAttempts % 5) == 0)
                    {
                        LogTalk("Deferred OnClose cleanup waiting for inspector close (token=" + _roomToken + ", attempt=" + _unsavedCloseCleanupAttempts.ToString(CultureInfo.InvariantCulture) + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ").");
                    }

                    return;
                }

                _unsavedCloseCleanupPending = false;
                _unsavedCloseCleanupAttempts = 0;
                StopUnsavedCloseCleanupTimer();

                if (!saved)
                {
                    LogTalk("Deferred OnClose cleanup deleting room (token=" + _roomToken + ").");
                    EnsureRoomDeleted();
                    return;
                }

                LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True, hasSavedState=" + hasSavedState + ", inspectorOpen=False).");
            }

            private static bool TryGetAppointmentSaved(Outlook.AppointmentItem appointment, out bool saved)
            {
                saved = false;
                if (appointment == null)
                {
                    return false;
                }

                try
                {
                    saved = appointment.Saved;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private void StopUnsavedCloseCleanupTimer()
            {
                if (_unsavedCloseCleanupTimer == null)
                {
                    return;
                }

                _unsavedCloseCleanupTimer.Stop();
                _unsavedCloseCleanupTimer.Tick -= OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Dispose();
                _unsavedCloseCleanupTimer = null;
            }

            private bool IsAppointmentOpenInAnyInspector(Outlook.AppointmentItem appointment)
            {
                if (appointment == null || _owner == null || _owner._inspectors == null)
                {
                    return false;
                }

                string appointmentEntryId = null;
                try
                {
                    appointmentEntryId = appointment.EntryID;
                }
                catch
                {
                }

                int inspectorCount = 0;
                try
                {
                    inspectorCount = _owner._inspectors.Count;
                }
                catch
                {
                    return false;
                }

                for (int i = 1; i <= inspectorCount; i++)
                {
                    Outlook.Inspector inspector = null;
                    object currentItem = null;
                    Outlook.AppointmentItem currentAppointment = null;
                    try
                    {
                        inspector = _owner._inspectors[i];
                        if (inspector == null)
                        {
                            continue;
                        }

                        currentItem = inspector.CurrentItem;
                        currentAppointment = currentItem as Outlook.AppointmentItem;
                        if (currentAppointment == null)
                        {
                            continue;
                        }

                        if (currentAppointment == appointment || ReferenceEquals(currentAppointment, appointment))
                        {
                            return true;
                        }

                        if (!string.IsNullOrWhiteSpace(appointmentEntryId))
                        {
                            string currentEntryId = null;
                            try
                            {
                                currentEntryId = currentAppointment.EntryID;
                            }
                            catch
                            {
                            }

                            if (!string.IsNullOrWhiteSpace(currentEntryId)
                                && string.Equals(currentEntryId, appointmentEntryId, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (currentAppointment != null && !ReferenceEquals(currentAppointment, appointment))
                        {
                            try
                            {
                                Marshal.ReleaseComObject(currentAppointment);
                            }
                            catch
                            {
                            }
                        }

                        if (currentItem != null
                            && !ReferenceEquals(currentItem, currentAppointment)
                            && !ReferenceEquals(currentItem, appointment))
                        {
                            try
                            {
                                Marshal.ReleaseComObject(currentItem);
                            }
                            catch
                            {
                            }
                        }

                        if (inspector != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(inspector);
                            }
                            catch
                            {
                            }
                        }
                    }
                }

                return false;
            }

            private void EnsureRoomDeleted()
            {
                if (_roomDeleted || !_owner.IsOrganizer(_appointment))
                {
                    if (_roomDeleted)
                    {
                        LogTalk("EnsureRoomDeleted: room already deleted (token=" + _roomToken + ").");
                    }

                    return;
                }

                string delegateId;
                if (_owner.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("EnsureRoomDeleted skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    Dispose();
                    return;
                }

                if (_owner.TryDeleteRoom(_roomToken, _isEventConversation))
                {
                    _owner.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    LogTalk("EnsureRoomDeleted successful (token=" + _roomToken + ").");
                    Dispose();
                }
                else
                {
                    LogTalk("EnsureRoomDeleted failed (token=" + _roomToken + ").");
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    LogTalk("Subscription.Dispose called again (token=" + _roomToken + ").");
                    return;
                }

                if (_events != null)
                {
                    _events.BeforeDelete -= OnBeforeDelete;
                    _events.Write -= OnWrite;
                    _events.Close -= OnClose;
                }

                StopUnsavedCloseCleanupTimer();

                _owner.UnregisterSubscription(_key, _roomToken, _entryId);
                _disposed = true;
                LogTalk("Subscription.Dispose completed (token=" + _roomToken + ").");
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
                if (appointment == null)
                {
                    return false;
                }

                if (appointment == _appointment || ReferenceEquals(appointment, _appointment))
                {
                    return true;
                }

                string thisEntryId = _entryId;
                if (string.IsNullOrWhiteSpace(thisEntryId))
                {
                    thisEntryId = GetEntryId(_appointment);
                }

                string otherEntryId = GetEntryId(appointment);
                if (!string.IsNullOrWhiteSpace(thisEntryId)
                    && !string.IsNullOrWhiteSpace(otherEntryId)
                    && string.Equals(thisEntryId, otherEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                return false;
            }

            internal bool MatchesToken(string token)
            {
                return string.Equals(_roomToken, token, StringComparison.OrdinalIgnoreCase);
            }

            internal void UpdateEntryId(string entryId)
            {
                _entryId = entryId;
                LogTalk("Subscription EntryId updated (token=" + _roomToken + ", EntryId=" + (_entryId ?? "n/a") + ").");
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
            if (appointment != null)
            {
                EnsureSubscriptionForAppointment(appointment);
            }
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

                var mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                {
                    string inspectorIdentityKey = ResolveComIdentityKey(inspector, "Inspector");
                    EnsureMailComposeSubscription(mail, inspectorIdentityKey);
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
            EnsureSubscriptionForAppointment(appointment, true);
        }

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment, bool allowDeferredRetry)
        {
            if (appointment == null)
            {
                return;
            }

            try
            {
                string roomToken = ResolveRoomTokenForAppointment(appointment);
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
                bool? isEventConversation = roomType.HasValue
                    ? (bool?)(roomType.Value == TalkRoomType.EventConversation)
                    : null;

                bool hasLobbyFlag = HasUserProperty(appointment, IcalLobby) || HasUserProperty(appointment, PropertyLobby);
                bool? lobbyEnabled = hasLobbyFlag
                    ? (bool?)GetUserPropertyBoolPrefer(appointment, IcalLobby, PropertyLobby)
                    : null;

                TryHydrateMissingRoomTraitsFromServer(appointment, roomToken, ref lobbyEnabled, ref isEventConversation);

                RegisterSubscription(
                    appointment,
                    roomToken,
                    lobbyEnabled.HasValue && lobbyEnabled.Value,
                    isEventConversation.HasValue && isEventConversation.Value);
            }
            catch (COMException ex)
            {
                if (IsOutlookEventProcedureRestriction(ex))
                {
                    if (allowDeferredRetry)
                    {
                        if (QueueDeferredAppointmentSubscriptionEnsure(appointment, ex))
                        {
                            return;
                        }

                        return;
                    }

                    LogDeferredAppointmentEnsureRestriction(
                        "Deferred appointment subscription ensure skipped: Outlook still blocks property access in the current event context (hresult=0x" +
                        ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                        ").");
                    return;
                }

                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure subscription for appointment.", ex);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure subscription for appointment.", ex);
            }
        }

        private bool QueueDeferredAppointmentSubscriptionEnsure(Outlook.AppointmentItem appointment, COMException triggerException)
        {
            if (appointment == null)
            {
                return false;
            }

            SynchronizationContext context = _uiSynchronizationContext;
            if (context == null)
            {
                LogDeferredAppointmentEnsureRestriction(
                    "Deferred appointment subscription ensure unavailable: UI synchronization context is missing (hresult=0x" +
                    triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                    ").");
                return false;
            }

            string entryId = TryGetEntryIdForDeferredKey(appointment);
            string globalAppointmentId = TryGetGlobalAppointmentIdForDeferredKey(appointment);
            string ensureKey = !string.IsNullOrWhiteSpace(entryId)
                ? ("entry:" + entryId.Trim())
                : (!string.IsNullOrWhiteSpace(globalAppointmentId)
                    ? ("gid:" + globalAppointmentId.Trim())
                    : ("obj:" + RuntimeHelpers.GetHashCode(appointment).ToString("X8", CultureInfo.InvariantCulture)));

            if (ensureKey.StartsWith("obj:", StringComparison.Ordinal))
            {
                LogDeferredAppointmentEnsureRestriction(
                    "Deferred appointment subscription ensure suppressed: unstable appointment identity during Outlook event restriction (hresult=0x" +
                    triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                    ").");
                return false;
            }

            lock (_pendingAppointmentEnsureSyncRoot)
            {
                if (_pendingAppointmentEnsureKeys.Contains(ensureKey))
                {
                    return true;
                }

                _pendingAppointmentEnsureKeys.Add(ensureKey);
            }

            LogDeferredAppointmentEnsureRestriction(
                "Deferred appointment subscription ensure queued (key=" + ensureKey +
                ", hresult=0x" + triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                ").");

            context.Post(
                _ =>
                {
                    try
                    {
                        EnsureSubscriptionForAppointment(appointment, false);
                    }
                    finally
                    {
                        lock (_pendingAppointmentEnsureSyncRoot)
                        {
                            _pendingAppointmentEnsureKeys.Remove(ensureKey);
                        }
                    }
                },
                null);

            return true;
        }

        private static bool IsOutlookEventProcedureRestriction(COMException ex)
        {
            if (ex == null)
            {
                return false;
            }

            return (ex.ErrorCode & 0xFFFF) == 0x0108;
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
