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
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Controllers;
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
    public sealed partial class NextcloudTalkAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private Outlook.Application _outlookApplication;
        private SettingsStorage _settingsStorage;
        private AddinSettings _currentSettings;
        private readonly Dictionary<string, AppointmentSubscription> _activeSubscriptions = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByToken = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private FreeBusyManager _freeBusyManager;
        private Outlook.Inspectors _inspectors;
        private Outlook.ApplicationEvents_11_Event _applicationEvents;
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByEntryId = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly MailComposeSubscriptionRegistryController _mailComposeSubscriptionRegistry = new MailComposeSubscriptionRegistryController();
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private readonly TalkAppointmentController _talkAppointmentController;
        private readonly ComposeShareLifecycleController _composeShareLifecycleController;
        private readonly HashSet<string> _pendingAppointmentEnsureKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _pendingAppointmentEnsureSyncRoot = new object();
        private DateTime _lastDeferredAppointmentEnsureRestrictionLogUtc = DateTime.MinValue;
        private DateTime _lastDeferredAppointmentEnsureUnstableIdentityLogUtc = DateTime.MinValue;
        private int _deferredAppointmentEnsureRestrictionSuppressedCount;
        private SynchronizationContext _uiSynchronizationContext;
        private IRibbonUI _ribbonUi;
        private const int ComposeAttachmentEvalDebounceMs = 250;
        private const int ComposeShareCleanupSendGraceMs = 15000;

        internal const string PropertyToken = "NcTalkRoomToken";
        internal const string PropertyRoomType = "NcTalkRoomType";
        internal const string PropertyLobby = "NcTalkLobbyEnabled";
        internal const string PropertySearchVisible = "NcTalkSearchVisible";
        internal const string PropertyPasswordSet = "NcTalkPasswordSet";
        internal const string PropertyStartEpoch = "NcTalkStartEpoch";
        internal const string PropertyDataVersion = "NcTalkDataVersion";
        internal const string PropertyAddUsers = "NcTalkAddUsers";
        internal const string PropertyAddGuests = "NcTalkAddGuests";
        internal const string PropertyDelegateId = "NcTalkDelegateId";
        internal const string PropertyDelegated = "NcTalkDelegated";
        internal const string IcalToken = "X-NCTALK-TOKEN";
        internal const string IcalUrl = "X-NCTALK-URL";
        internal const string IcalLobby = "X-NCTALK-LOBBY";
        internal const string IcalStart = "X-NCTALK-START";
        internal const string IcalEvent = "X-NCTALK-EVENT";
        internal const string IcalObjectId = "X-NCTALK-OBJECTID";
        internal const string IcalAddUsers = "X-NCTALK-ADD-USERS";
        internal const string IcalAddGuests = "X-NCTALK-ADD-GUESTS";
        internal const string IcalAddParticipants = "X-NCTALK-ADD-PARTICIPANTS";
        internal const string IcalDelegate = "X-NCTALK-DELEGATE";
        internal const string IcalDelegateName = "X-NCTALK-DELEGATE-NAME";
        internal const string IcalDelegated = "X-NCTALK-DELEGATED";
        internal const string IcalDelegateReady = "X-NCTALK-DELEGATE-READY";
        internal const int PropertyVersionValue = 1;
        private static readonly Regex TalkUrlTokenRegex = new Regex(@"https?://[^\s""'<>]+/call/(?<token>[A-Za-z0-9_-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public NextcloudTalkAddIn()
        {
            _talkAppointmentController = new TalkAppointmentController(this);
            _composeShareLifecycleController = new ComposeShareLifecycleController(this);
        }

        internal AddinSettings CurrentSettings
        {
            get { return _currentSettings; }
        }

        internal SettingsStorage SettingsStorage
        {
            get { return _settingsStorage; }
        }

        internal Outlook.Application OutlookApplication
        {
            get { return _outlookApplication; }
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
            ConfigureDiagnosticsLogger(_currentSettings);
            TryApplyTransportSecurityFromSettings("startup", false);
            TryApplyOfficeUiLanguage();
            LogCore("Add-in connected (Outlook version=" + (_outlookApplication != null ? _outlookApplication.Version : "unknown") + ").");
            if (!string.IsNullOrWhiteSpace(outlookProfileName))
            {
                LogCore("Using Outlook profile settings: " + outlookProfileName + ".");
            }
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_currentSettings != null)
            {
                LogSettings("Settings loaded (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", IfbPort=" + _currentSettings.IfbPort + ", Debug=" + _currentSettings.DebugLoggingEnabled + ", LogAnonymize=" + _currentSettings.LogAnonymizationEnabled + ").");
            }
            _freeBusyManager = new FreeBusyManager(_settingsStorage.DataDirectory);
            _freeBusyManager.Initialize(_outlookApplication);
            EnsureApplicationHook();
            EnsureInspectorHook();
            ApplyIfbSettings();
        }

        private void TryApplyOfficeUiLanguage()
        {
            try
            {
                // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
                if (_outlookApplication == null)
                {
                    return;
                }

                LanguageSettings languageSettings = _outlookApplication.LanguageSettings;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            object session = null;
            try
            {
                session = _outlookApplication.Session;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Core,
                    "Failed to release Outlook session COM object after profile resolution.");
            }
        }

        /**
         * Outlook signals that the add-in is unloading.
         * Cleanup hooks follow once resources are held.
         */
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            TearDownAddInState("disconnect", true);
            LogCore("Add-in disconnected (removeMode=" + removeMode + ").");
        }

        /**
         * Required IDTExtensibility2 callback.
         * Intentionally no-op because runtime wiring is already complete in OnConnection.
         */
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /**
         * Required IDTExtensibility2 callback.
         * Intentionally no-op because startup work is handled in OnConnection.
         */
        public void OnStartupComplete(ref Array custom)
        {
        }

        /**
         * Called when Outlook shuts down; reserved for future cleanup steps.
         */
        public void OnBeginShutdown(ref Array custom)
        {
            TearDownAddInState("shutdown", false);
        }

        /**
         * Centralized teardown used by both OnBeginShutdown and OnDisconnection.
         * This path must be idempotent, because Outlook can call both callbacks.
         */
        private void TearDownAddInState(string origin, bool clearOutlookApplication)
        {
            UnhookApplication();
            UnhookInspector();
            UnhookMailComposeSubscriptions();

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "Failed to disable IFB during add-in " + (origin ?? "teardown") + ".",
                        ex);
                }
            }

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_freeBusyManager != null)
            {
                _freeBusyManager.Dispose();
            }
            _freeBusyManager = null;

            if (clearOutlookApplication)
            {
                _outlookApplication = null;
            }

            _ribbonUi = null;
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
            // Keep a stable handle so future dynamic ribbon refreshes can call Invalidate/InvalidateControl.
            if (!ReferenceEquals(_ribbonUi, ribbonUI))
            {
                _ribbonUi = ribbonUI;
            }
        }

        public async void OnTalkButtonPressed(IRibbonControl control)
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => FetchBackendPolicyStatus(configuration, "talk_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => FetchPasswordPolicyForTalkWizard(configuration));
            await Task.WhenAll(policyStatusTask, passwordPolicyTask);
            BackendPolicyStatus policyStatus = policyStatusTask.Result;
            PasswordPolicyInfo passwordPolicy = passwordPolicyTask.Result;

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

            List<NextcloudUser> userDirectory;
            try
            {
                userDirectory = talkClickAddressbookStatus.Available
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
                policyStatus,
                userDirectory,
                talkClickAddressbookStatus,
                subject,
                start,
                end))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    LogTalk("Talk link dialog cancelled.");
                    return;
                }

                string descriptionLanguage = ResolveTalkDescriptionLanguage(
                    policyStatus,
                    _currentSettings != null ? _currentSettings.EventDescriptionLang : "default");
                string descriptionType = ResolveTalkEventDescriptionType(policyStatus);
                string invitationTemplate = ResolveTalkInvitationTemplate(policyStatus);
                string initialDescription;
                try
                {
                    initialDescription = BuildInitialRoomDescription(
                        dialog.TalkPassword,
                        descriptionLanguage,
                        invitationTemplate);
                }
                catch (Exception ex)
                {
                    LogTalk("Talk invitation template rendering blocked: " + ex.Message);
                    MessageBox.Show(
                        string.Format(CultureInfo.CurrentCulture, Strings.ErrorCreateRoomUnexpected, ex.Message),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
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
                    DescriptionLanguage = descriptionLanguage,
                    DescriptionType = descriptionType,
                    InvitationTemplate = invitationTemplate,
                    Description = initialDescription,
                    AddUsers = dialog.AddUsers,
                    AddGuests = dialog.AddGuests,
                    DelegateModeratorId = dialog.DelegateModeratorId,
                    DelegateModeratorName = dialog.DelegateModeratorName
                };
                LogTalk("Room request prepared (title='" + request.Title + "', type=" + request.RoomType + ", lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ", passwordSet=" + (!string.IsNullOrEmpty(request.Password)) + ").");

                string existingToken = TalkAppointmentController.GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
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

                    var existingType = TalkAppointmentController.GetRoomType(appointment);
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

        public async void OnSettingsButtonPressed(IRibbonControl control)
        {
            EnsureSettingsLoaded();
            LogSettings("Settings dialog opened.");
            var configuration = new TalkServiceConfiguration(
                _currentSettings != null ? _currentSettings.ServerUrl : string.Empty,
                _currentSettings != null ? _currentSettings.Username : string.Empty,
                _currentSettings != null ? _currentSettings.AppPassword : string.Empty);
            BackendPolicyStatus initialPolicyStatus = await Task.Run(() => FetchBackendPolicyStatus(configuration, "settings_open_initial"));
            using (var form = new SettingsForm(_currentSettings, _outlookApplication, initialPolicyStatus))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    AddinSettings previousSettings = (_currentSettings ?? new AddinSettings()).Clone();
                    _currentSettings = form.Result ?? new AddinSettings();
                    ConfigureDiagnosticsLogger(_currentSettings);
                    if (!TryApplyTransportSecurityFromSettings("settings_save", true))
                    {
                        _currentSettings = previousSettings;
                        ConfigureDiagnosticsLogger(_currentSettings);
                        TryApplyTransportSecurityFromSettings("settings_save_revert", false);
                        LogSettings("Settings save aborted because transport security settings could not be applied.");
                        return;
                    }
                    LogSettings("Settings applied (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", IfbPort=" + _currentSettings.IfbPort + ", Debug=" + _currentSettings.DebugLoggingEnabled + ", LogAnonymize=" + _currentSettings.LogAnonymizationEnabled + ").");
                    ApplyIfbSettings();
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null || _currentSettings == null)
            {
                return false;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);
            // Keep this method synchronous for existing callers and prefetch both policies in parallel.
            Task<BackendPolicyStatus> policyStatusTask = Task.Run(() => FetchBackendPolicyStatus(configuration, "sharing_wizard_open"));
            Task<PasswordPolicyInfo> passwordPolicyTask = Task.Run(() => FetchPasswordPolicyForFileLinkWizard(configuration));
            Task.WhenAll(policyStatusTask, passwordPolicyTask).GetAwaiter().GetResult();
            BackendPolicyStatus policyStatus = policyStatusTask.Result;

            string basePath = string.IsNullOrWhiteSpace(_currentSettings.FileLinkBasePath)
                ? "90 Freigaben - extern"
                : _currentSettings.FileLinkBasePath;

            PasswordPolicyInfo passwordPolicy = passwordPolicyTask.Result;

            using (var wizard = new FileLinkWizardForm(_currentSettings ?? new AddinSettings(), configuration, passwordPolicy, policyStatus, basePath, launchOptions))
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (wizard.ShowDialog() == DialogResult.OK && wizard.Result != null)
                {
                    string languageOverride = _currentSettings != null ? _currentSettings.ShareBlockLang : "default";
                    LogFileLink("Share created (folder=\"" + wizard.Result.FolderName + "\").");

                    MailComposeSubscription composeSubscription = EnsureMailComposeSubscription(mail, ResolveActiveInspectorIdentityKey());
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (composeSubscription != null)
                    {
                        composeSubscription.ArmShareCleanup(wizard.Result);
                    }

                    string html;
                    try
                    {
                        html = FileLinkHtmlBuilder.Build(wizard.Result, wizard.RequestSnapshot, languageOverride, policyStatus);
                    }
                    catch (Exception ex)
                    {
                        LogFileLink("Share template rendering blocked: " + ex.Message);
                        MessageBox.Show(
                            string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                            Strings.DialogTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return false;
                    }

                    if (composeSubscription != null
                        && wizard.RequestSnapshot != null
                        && wizard.RequestSnapshot.PasswordSeparateEnabled
                        && !string.IsNullOrWhiteSpace(wizard.Result.Password))
                    {
                        string passwordOnlyHtml;
                        try
                        {
                            passwordOnlyHtml = FileLinkHtmlBuilder.BuildPasswordOnly(wizard.Result, languageOverride, policyStatus);
                        }
                        catch (Exception ex)
                        {
                            LogFileLink("Password-only template rendering blocked: " + ex.Message);
                            MessageBox.Show(
                                string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                                Strings.DialogTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return false;
                        }

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

        internal sealed class SeparatePasswordDispatchEntry
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null)
            {
                return null;
            }

            string mailIdentityKey = ComInteropScope.ResolveIdentityKey(mail, LogCategories.FileLink, "MailItem");
            string inspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityOverride)
                ? ResolveMailInspectorIdentityKey(mail)
                : inspectorIdentityOverride.Trim();

            return _mailComposeSubscriptionRegistry.GetOrCreate(
                mail,
                mailIdentityKey,
                inspectorIdentityKey,
                () => new MailComposeSubscription(this, mail, mailIdentityKey, inspectorIdentityKey));
        }

        private static string ResolveMailInspectorIdentityKey(Outlook.MailItem mail)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                return ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
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
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release compose Inspector COM object.");
            }
        }

        private IWin32Window TryCreateMailInspectorDialogOwner(Outlook.MailItem mail)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release compose prompt owner Inspector COM object.");
            }
        }

        private static int ReadInspectorWindowHandle(Outlook.Inspector inspector)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (inspector == null)
            {
                return 0;
            }

            foreach (string propertyName in new[] { "HWND", "Hwnd" })
            {
                try
                {
                    PropertyInfo property = inspector.GetType().GetProperty(propertyName);
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (property == null)
                    {
                        continue;
                    }

                    object value = property.GetValue(inspector, null);
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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

        private void RemoveMailComposeSubscription(MailComposeSubscription subscription)
        {
            _mailComposeSubscriptionRegistry.Remove(subscription);
        }

        private void UnhookMailComposeSubscriptions()
        {
            _mailComposeSubscriptionRegistry.DisposeAll();
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

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
            return _composeShareLifecycleController.TryDeleteComposeShareFolder(relativeFolder, reason, shareId, shareLabel);
        }

        private void DispatchSeparatePasswordMailQueue(string composeKey, List<SeparatePasswordDispatchEntry> queue)
        {
            _composeShareLifecycleController.DispatchSeparatePasswordMailQueue(composeKey, queue);
        }

        internal void ShowPasswordMailSuccessNotification(int recipientCount)
        {
            if (recipientCount <= 0)
            {
                return;
            }

            SynchronizationContext notificationUiContext = _uiSynchronizationContext ?? SynchronizationContext.Current;
            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
            if (notificationUiContext == null)
            {
                LogFileLink("Separate password notification skipped (UI context unavailable, recipients=" + recipientCount.ToString(CultureInfo.InvariantCulture) + ").");
                return;
            }

            try
            {
                notificationUiContext.Post(
                    _ => ShowPasswordMailSuccessNotificationOnUiContext(recipientCount, notificationUiContext),
                    null);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Separate password notification failed.", ex);
            }
        }

        private void ShowPasswordMailSuccessNotificationOnUiContext(int recipientCount, SynchronizationContext notificationUiContext)
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
                ScheduleNotifyIconDispose(notifyIcon, 7000, notificationUiContext);
                LogFileLink("Separate password notification shown (recipients=" + recipientCount.ToString(CultureInfo.InvariantCulture) + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Separate password notification failed on UI context.", ex);
            }
        }

        private void ScheduleNotifyIconDispose(NotifyIcon notifyIcon, int delayMs, SynchronizationContext notificationUiContext)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (notifyIcon == null)
            {
                return;
            }

            int effectiveDelayMs = Math.Max(0, delayMs);
            Task.Delay(effectiveDelayMs).ContinueWith(
                _ => DisposeNotifyIconOnUiContext(notifyIcon, notificationUiContext),
                CancellationToken.None,
                TaskContinuationOptions.None,
                TaskScheduler.Default);
        }

        private void DisposeNotifyIconOnUiContext(NotifyIcon notifyIcon, SynchronizationContext notificationUiContext)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (notifyIcon == null)
            {
                return;
            }

            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
            if (notificationUiContext == null)
            {
                DisposeNotifyIcon(notifyIcon);
                return;
            }

            try
            {
                notificationUiContext.Post(_ => DisposeNotifyIcon(notifyIcon), null);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to marshal password notification icon dispose onto UI context.", ex);
                DisposeNotifyIcon(notifyIcon);
            }
        }

        private static void DisposeNotifyIcon(NotifyIcon notifyIcon)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (notifyIcon == null)
            {
                return;
            }

            try
            {
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose password notification icon.", ex);
            }
        }

        private Outlook.MailItem GetActiveMailItem()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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

            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (mailItem != null)
                {
                    return mailItem;
                }
            }

            return null;
        }

        private string ResolveActiveInspectorIdentityKey()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = _outlookApplication.ActiveInspector();
                return ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "ActiveInspector");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve active inspector identity key.", ex);
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release active Inspector COM object.");
            }
        }

        private void InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null || string.IsNullOrWhiteSpace(html))
            {
                return;
            }

            if (TryInsertHtmlIntoMailBody(mail, html))
            {
                LogCore("Inserted HTML block into mail (HTMLBody primary).");
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

            LogCore("Failed to insert HTML into mail: all insertion paths exhausted.");
            MessageBox.Show(
                string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, "all insertion paths exhausted"),
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private bool TryInsertHtmlIntoMailBody(Outlook.MailItem mail, string html)
        {
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

                return true;
            }
            catch (Exception ex)
            {
                LogCore("Failed to insert HTML via HTMLBody path: " + ex.Message);
                return false;
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

            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (wordEditor == null)
            {
                return false;
            }

            try
            {
                application = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (application == null)
                {
                    return false;
                }

                selection = application.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, application, null);
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release Word selection COM object.");
                ComInteropScope.TryRelease(application, LogCategories.Core, "Failed to release Word application COM object.");
            }
        }

        internal static bool TryWriteAppointmentHtmlBody(Outlook.AppointmentItem appointment, string html)
        {
            return TryWriteAppointmentHtmlBodyWithRtfBridge(appointment, html);
        }

        private static bool TryWriteAppointmentHtmlBodyWithRtfBridge(Outlook.AppointmentItem appointment, string html)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(html))
            {
                return false;
            }

            Outlook.Application application = null;
            Outlook.MailItem stagingMail = null;

            try
            {
                application = appointment.Application;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (application == null)
                {
                    return false;
                }

                stagingMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (stagingMail == null)
                {
                    return false;
                }

                string bridgeHtml = HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(html);
                if (string.IsNullOrWhiteSpace(bridgeHtml))
                {
                    bridgeHtml = html ?? string.Empty;
                }

                stagingMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                stagingMail.HTMLBody = bridgeHtml;

                var rtfBody = stagingMail.RTFBody as byte[];
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (rtfBody == null || rtfBody.Length == 0)
                {
                    DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML->RTF bridge produced empty RTF body.");
                    return false;
                }

                appointment.RTFBody = rtfBody;
                DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML body written via HTML->RTF bridge.");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to write appointment HTML body via HTML->RTF bridge.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(stagingMail, LogCategories.Talk, "Failed to release staging MailItem COM object.");
                ComInteropScope.TryRelease(application, LogCategories.Talk, "Failed to release Outlook application COM object for appointment HTML->RTF bridge.");
            }
        }

        private Outlook.AppointmentItem GetActiveAppointment()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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

            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (inspector != null)
            {
                return inspector.CurrentItem as Outlook.AppointmentItem;
            }

            return null;
        }

        internal bool SettingsAreComplete()
        {
            return _currentSettings != null
                   && !string.IsNullOrWhiteSpace(_currentSettings.ServerUrl)
                   && !string.IsNullOrWhiteSpace(_currentSettings.Username)
                   && !string.IsNullOrWhiteSpace(_currentSettings.AppPassword);
        }

        private void ApplyIfbSettings()
        {
            // Diagnostics logger configuration is intentionally handled by the caller
            // (startup/settings save) to avoid duplicate reconfiguration on normal settings saves.

            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_freeBusyManager == null || _currentSettings == null)
            {
                return;
            }

            try
            {
                LogCore("Applying IFB (Enabled=" + _currentSettings.IfbEnabled + ", Days=" + _currentSettings.IfbDays + ", Port=" + _currentSettings.IfbPort + ", CacheHours=" + _currentSettings.IfbCacheHours + ").");
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

        internal TalkService CreateTalkService()
        {
            return new TalkService(new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword));
        }

        private BackendPolicyStatus FetchBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            try
            {
                var service = new BackendPolicyService(configuration);
                BackendPolicyStatus status = service.FetchStatus();
                LogCore(
                    "Backend policy status fetched (trigger=" + (trigger ?? "n/a")
                    + ", active=" + (status != null && status.PolicyActive)
                    + ", warningVisible=" + (status != null && status.WarningVisible)
                    + ", mode=" + (status != null ? status.Mode : "local")
                    + ", reason=" + (status != null ? status.Reason : "n/a")
                    + ").");
                return status;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Backend policy status fetch failed (trigger=" + (trigger ?? "n/a") + ").", ex);
                return null;
            }
        }

        private PasswordPolicyInfo FetchPasswordPolicyForTalkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogTalk("Password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        private PasswordPolicyInfo FetchPasswordPolicyForFileLinkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogFileLink("Sharing password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        private static string ResolveTalkDescriptionLanguage(BackendPolicyStatus policyStatus, string fallbackLanguageOverride)
        {
            if (policyStatus != null
                && policyStatus.PolicyActive
                && policyStatus.IsLocked("talk", "language_talk_description"))
            {
                string policyLanguageRaw = policyStatus.GetPolicyString("talk", "language_talk_description");
                if (!string.IsNullOrWhiteSpace(policyLanguageRaw))
                {
                    return NormalizeTalkDescriptionLanguage(policyLanguageRaw);
                }
            }

            return NormalizeTalkDescriptionLanguage(fallbackLanguageOverride);
        }

        private static string ResolveTalkInvitationTemplate(BackendPolicyStatus policyStatus)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (policyStatus == null || !policyStatus.PolicyActive)
            {
                return string.Empty;
            }

            return policyStatus.GetPolicyString("talk", "talk_invitation_template");
        }

        private static string ResolveTalkEventDescriptionType(BackendPolicyStatus policyStatus)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (policyStatus != null && policyStatus.PolicyActive)
            {
                string policyTypeRaw = policyStatus.GetPolicyString("talk", "event_description_type");
                if (!string.IsNullOrWhiteSpace(policyTypeRaw))
                {
                    return NormalizeTalkEventDescriptionType(policyTypeRaw);
                }
            }

            return "plain_text";
        }

        /**
         * Normalize Talk description language while preserving the explicit "custom" mode.
         */
        private static string NormalizeTalkDescriptionLanguage(string languageOverride)
        {
            if (string.IsNullOrWhiteSpace(languageOverride))
            {
                return "default";
            }

            string trimmed = languageOverride.Trim();
            if (string.Equals(trimmed, "custom", StringComparison.OrdinalIgnoreCase))
            {
                return "custom";
            }

            return Strings.NormalizeLanguageOverride(trimmed);
        }

        internal static string NormalizeTalkEventDescriptionType(string descriptionType)
        {
            return string.Equals(descriptionType, "html", StringComparison.OrdinalIgnoreCase)
                ? "html"
                : "plain_text";
        }

        private void ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            _talkAppointmentController.ApplyRoomToAppointment(appointment, request, result);
        }

        internal static string UpdateBodyWithTalkBlock(string existingBody, string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            return TalkDescriptionTemplateController.UpdateBodyWithTalkBlock(existingBody, roomUrl, password, languageOverride, invitationTemplate);
        }

        internal static string UpdateHtmlBodyWithTalkBlock(string existingHtmlBody, string existingBody, string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            return TalkDescriptionTemplateController.UpdateHtmlBodyWithTalkBlock(existingHtmlBody, existingBody, roomUrl, password, languageOverride, invitationTemplate);
        }

        private static string BuildInitialRoomDescription(string password, string languageOverride, string invitationTemplate)
        {
            return TalkDescriptionTemplateController.BuildInitialRoomDescription(password, languageOverride, invitationTemplate);
        }

        /**
         * Outlook appointment bodies are plain text. Convert backend HTML/text
         * templates into a stable plain-text block before inserting them.
         */
        internal static string ConvertHtmlTemplateToPlainText(string value)
        {
            return TalkDescriptionTemplateController.ConvertHtmlTemplateToPlainText(value);
        }

        private bool TryStampIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            return _talkAppointmentController.TryStampIcalStartEpoch(appointment, roomToken, out startEpoch);
        }

        private bool TryReadRequiredIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            return _talkAppointmentController.TryReadRequiredIcalStartEpoch(appointment, roomToken, out startEpoch);
        }

        private static long? GetIcalStartEpochOrNull(Outlook.AppointmentItem appointment)
        {
            return TalkAppointmentController.GetIcalStartEpochOrNull(appointment);
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
            // Deferred ensure path can run after appointment release; null means "no stable key".
            if (appointment == null)
            {
                return null;
            }

            try
            {
                return appointment.GlobalAppointmentID;
            }
            catch
            {
                return null;
            }
        }

        private void UpdateStoredServerVersion(string response)
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
            string roomToken = TalkAppointmentController.GetUserPropertyTextPrefer(appointment, IcalToken, PropertyToken);
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
                TalkAppointmentController.SetUserProperty(appointment, IcalToken, Outlook.OlUserPropertyType.olText, extracted);
                TalkAppointmentController.SetUserProperty(appointment, PropertyToken, Outlook.OlUserPropertyType.olText, extracted);
                LogTalk("Room token bootstrapped from location (token=" + extracted + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to persist bootstrapped room token.", ex);
            }

            return extracted;
        }

        private void RefreshEntryBinding(AppointmentSubscription subscription)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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

        private void ResolveRuntimeRoomTraits(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool fallbackLobbyEnabled,
            bool fallbackIsEventConversation,
            out bool lobbyKnown,
            out bool lobbyEnabled,
            out bool isEventConversation)
        {
            _talkAppointmentController.ResolveRuntimeRoomTraits(
                appointment,
                roomToken,
                fallbackLobbyEnabled,
                fallbackIsEventConversation,
                out lobbyKnown,
                out lobbyEnabled,
                out isEventConversation);
        }

        private void PersistLobbyTraits(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled)
        {
            _talkAppointmentController.PersistLobbyTraits(appointment, roomToken, lobbyEnabled);
        }

        private bool IsOrganizer(Outlook.AppointmentItem appointment)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

        internal void ClearTalkProperties(Outlook.AppointmentItem appointment)
        {
            _talkAppointmentController.ClearTalkProperties(appointment);
        }

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, TalkRoomCreationResult result)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (result == null)
            {
                return;
            }

            RegisterSubscription(appointment, result.RoomToken, result.LobbyEnabled, result.CreatedAsEventConversation);
        }

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled, bool isEventConversation)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
            return _talkAppointmentController.IsDelegatedToOtherUser(appointment, out delegateId);
        }

        private bool IsDelegationPending(Outlook.AppointmentItem appointment, out string delegateId)
        {
            return _talkAppointmentController.IsDelegationPending(appointment, out delegateId);
        }

        // Returns true when participant sync completed without runtime/service failures.
        // This status is used for deterministic pre-delegation logging in OnWrite.
        private bool TrySyncRoomParticipants(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            return _talkAppointmentController.TrySyncRoomParticipants(appointment, roomToken, isEventConversation);
        }

        private void TryApplyDelegation(Outlook.AppointmentItem appointment, string roomToken)
        {
            _talkAppointmentController.TryApplyDelegation(appointment, roomToken);
        }

        internal static List<string> GetAppointmentAttendeeEmails(Outlook.AppointmentItem appointment)
        {
            return OutlookRecipientResolverController.CollectAppointmentAttendeeEmails(appointment);
        }

        private static string TryGetRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            return OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient);
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
            return _talkAppointmentController.TryUpdateLobby(appointment, roomToken, isEventConversation);
        }

        private bool TryUpdateRoomName(Outlook.AppointmentItem appointment, string roomToken)
        {
            return _talkAppointmentController.TryUpdateRoomName(appointment, roomToken);
        }

        private bool TryUpdateRoomDescription(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            return _talkAppointmentController.TryUpdateRoomDescription(appointment, roomToken, isEventConversation);
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

        internal static void ShowWarningDialog(string message)
        {
            ShowWarning(message);
        }

        internal void EnsureSettingsLoaded()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_currentSettings == null)
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_settingsStorage != null)
                {
                    _currentSettings = _settingsStorage.Load();
                }

                // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
                if (_currentSettings == null)
                {
                    _currentSettings = new AddinSettings();
                }
                TryApplyTransportSecurityFromSettings("lazy_load", false);
                EnsureInspectorHook();
            }
        }

        private static void ConfigureDiagnosticsLogger(AddinSettings settings)
        {
            bool debugEnabled = settings != null && settings.DebugLoggingEnabled;
            bool anonymizationEnabled = settings == null || settings.LogAnonymizationEnabled;
            string serverUrl = settings != null ? settings.ServerUrl : string.Empty;

            DiagnosticsLogger.SetEnabled(debugEnabled);
            DiagnosticsLogger.SetAnonymization(anonymizationEnabled, serverUrl);
        }

        private bool TryApplyTransportSecurityFromSettings(string source, bool showWarning)
        {
            try
            {
                TransportSecurityConfigurator.ApplyFromSettings(_currentSettings, source);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply transport security settings (source=" + (source ?? string.Empty) + ").",
                    ex);

                if (showWarning)
                {
                    ShowWarning(string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.TransportTlsApplyFailed,
                        ex.Message));
                }

                return false;
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

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment)
        {
            EnsureSubscriptionForAppointment(appointment, true);
        }

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment, bool allowDeferredRetry)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

                bool lobbyKnown;
                bool lobbyEnabled;
                bool isEventConversation;
                _talkAppointmentController.ResolveRuntimeRoomTraits(
                    appointment,
                    roomToken,
                    false,
                    false,
                    out lobbyKnown,
                    out lobbyEnabled,
                    out isEventConversation);

                RegisterSubscription(appointment, roomToken, lobbyEnabled, isEventConversation);
            }
            catch (COMException ex)
            {
                if (IsOutlookEventProcedureRestriction(ex))
                {
                    if (allowDeferredRetry)
                    {
                        QueueDeferredAppointmentSubscriptionEnsure(appointment, ex);
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
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return false;
            }

            SynchronizationContext context = _uiSynchronizationContext;
            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
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
                bool shouldEmitIdentityRestriction;
                lock (_pendingAppointmentEnsureSyncRoot)
                {
                    DateTime nowUtc = DateTime.UtcNow;
                    shouldEmitIdentityRestriction = _lastDeferredAppointmentEnsureUnstableIdentityLogUtc == DateTime.MinValue ||
                        (nowUtc - _lastDeferredAppointmentEnsureUnstableIdentityLogUtc).TotalSeconds >= 60;
                    if (shouldEmitIdentityRestriction)
                    {
                        _lastDeferredAppointmentEnsureUnstableIdentityLogUtc = nowUtc;
                    }
                }

                if (shouldEmitIdentityRestriction)
                {
                    LogDeferredAppointmentEnsureRestriction(
                        "Deferred appointment subscription ensure suppressed: unstable appointment identity during Outlook event restriction (hresult=0x" +
                        triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                        ").");
                }
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
            // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
            if (ex == null)
            {
                return false;
            }

            return (ex.ErrorCode & 0xFFFF) == 0x0108;
        }

    }
}
