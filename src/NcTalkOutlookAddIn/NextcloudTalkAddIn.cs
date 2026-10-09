// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Entry point and ribbon implementation for NC Connector for Outlook.
    // Registers as a classic COM add-in via IDTExtensibility2 and provides the
    // ribbon XML for appointment windows.
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
        private Outlook.Explorers _explorers;
        private Outlook.ExplorersEvents_Event _explorersEvents;
        private readonly Dictionary<string, Outlook.Explorer> _hookedExplorers = new Dictionary<string, Outlook.Explorer>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Outlook.ExplorerEvents_10_Event> _hookedExplorerEvents = new Dictionary<string, Outlook.ExplorerEvents_10_Event>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Outlook.ExplorerEvents_10_InlineResponseEventHandler> _inlineResponseHandlers = new Dictionary<string, Outlook.ExplorerEvents_10_InlineResponseEventHandler>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler> _inlineResponseCloseHandlers = new Dictionary<string, Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Outlook.ExplorerEvents_10_SelectionChangeEventHandler> _explorerSelectionChangeHandlers = new Dictionary<string, Outlook.ExplorerEvents_10_SelectionChangeEventHandler>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MailComposeSubscription> _inlineResponseSubscriptions = new Dictionary<string, MailComposeSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, AppointmentSubscription> _subscriptionByEntryId = new Dictionary<string, AppointmentSubscription>(StringComparer.OrdinalIgnoreCase);
        private readonly MailComposeSubscriptionRegistryController _mailComposeSubscriptionRegistry = new MailComposeSubscriptionRegistryController();
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private readonly TalkAppointmentController _talkAppointmentController;
        private readonly ComposeShareCleanupService _composeShareCleanupService = new ComposeShareCleanupService();
        private readonly SeparatePasswordDeliveryController _separatePasswordDeliveryController;
        private readonly FileLinkLaunchController _fileLinkLaunchController;
        private readonly TalkRibbonController _talkRibbonController;
        private readonly MailInteropController _mailInteropController;
        private readonly MailBodyInsertionController _mailBodyInsertionController;
        private readonly ManagedEmailSignatureController _managedEmailSignatureController;
        private readonly UpdateCheckService _updateCheckService = new UpdateCheckService();
        private readonly DeferredAppointmentEnsureState _deferredAppointmentEnsureState = new DeferredAppointmentEnsureState();
        private OutlookUiSynchronizationContext _uiSynchronizationContext;
        private IRibbonUI _ribbonUi;
        private const int ComposeAttachmentEvalDebounceMs = 250;
        internal const string IcalToken = "X-NCTALK-TOKEN";
        internal const string IcalUrl = "X-NCTALK-URL";
        internal const string IcalLobby = "X-NCTALK-LOBBY";
        internal const string IcalStart = "X-NCTALK-START";
        internal const string IcalEvent = "X-NCTALK-EVENT";
        internal const string IcalObjectId = "X-NCTALK-OBJECTID";
        internal const string IcalAddUsers = "X-NCTALK-ADD-USERS";
        internal const string IcalAddGuests = "X-NCTALK-ADD-GUESTS";
        internal const string IcalDelegate = "X-NCTALK-DELEGATE";
        internal const string IcalDelegateName = "X-NCTALK-DELEGATE-NAME";
        internal const string IcalDelegated = "X-NCTALK-DELEGATED";
        internal const string IcalDelegateReady = "X-NCTALK-DELEGATE-READY";

        public NextcloudTalkAddIn()
        {
            _talkAppointmentController = new TalkAppointmentController(this);
            _separatePasswordDeliveryController =
                new SeparatePasswordDeliveryController(this);
            _fileLinkLaunchController = new FileLinkLaunchController(this);
            _talkRibbonController = new TalkRibbonController(this);
            _mailInteropController = new MailInteropController(this);
            _mailBodyInsertionController = new MailBodyInsertionController(this, _mailInteropController);
            _managedEmailSignatureController = new ManagedEmailSignatureController(this);
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
      <tab id='NcTalkExplorerTab' label='{0}' insertAfterMso='TabMail' getVisible='OnGetMainRibbonTabVisible'>
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
    <contextualTabs>
      <tabSet idMso='TabComposeTools'>
        <tab idMso='TabMessage'>
          <group id='NcTalkInlineMailGroup' label='{1}'>
            <button id='NcTalkInlineFileLinkButton'
                    label='{5}'
                    size='large'
                    getImage='OnGetButtonImage'
                    onAction='OnFileLinkButtonPressed'
                    screentip='{6}'
                    supertip='{7}' />
          </group>
        </tab>
      </tabSet>
    </contextualTabs>
  </ribbon>
</customUI>",
                    EscapeXml(Strings.RibbonExplorerTabLabel),
                    EscapeXml(Strings.RibbonExplorerGroupLabel),
                    EscapeXml(Strings.RibbonSettingsButtonLabel),
                    EscapeXml(Strings.RibbonSettingsScreenTip),
                    EscapeXml(Strings.RibbonSettingsSuperTip),
                    EscapeXml(Strings.RibbonFileLinkButtonLabel),
                    EscapeXml(Strings.RibbonFileLinkButtonScreenTip),
                    EscapeXml(Strings.RibbonFileLinkButtonSuperTip));
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

                // Outlook passes the ribbon handle right after loading.
        // Stores the instance for later refresh operations.
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
            try
            {
                EnsureSettingsLoaded();
                await _talkRibbonController.OnTalkButtonPressedAsync(control);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Talk ribbon handler failed.", ex);
            }
        }

        public async void OnSettingsButtonPressed(IRibbonControl control)
        {
            try
            {
                EnsureSettingsLoaded();
                if (!_currentSettings.ShowMainRibbonTab)
                {
                    return;
                }
                await CreateSettingsWorkflowController().RunAsync();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Settings ribbon handler failed.", ex);
            }
        }

        private SettingsWorkflowController CreateSettingsWorkflowController()
        {
            return new SettingsWorkflowController(
                _outlookApplication,
                () => _currentSettings,
                settings =>
                {
                    _currentSettings = settings;
                    _mailComposeSubscriptionRegistry
                        .RefreshAttachmentAutomationSettings();
                },
                (configuration, trigger) => FetchBackendPolicyStatus(configuration, trigger),
                settings => ConfigureDiagnosticsLogger(settings),
                (settings, source, showWarning) =>
                    TryApplyTransportSecurityFromSettings(
                        settings,
                        source,
                        showWarning),
                () => ApplyIfbSettings(),
                settings =>
                {
                    if (_settingsStorage != null)
                    {
                        _settingsStorage.SaveUserInitiated(settings);
                    }
                },
                callback => RunOnOutlookUiThreadAsync(callback),
                message => LogSettings(message),
                _settingsStorage != null
                    ? _settingsStorage.DataDirectory
                    : string.Empty,
                OutlookProfileScope);
        }

        public bool OnGetMainRibbonTabVisible(IRibbonControl control)
        {
            EnsureSettingsLoaded();
            return _currentSettings.ShowMainRibbonTab;
        }

        internal Task<AddinSettings> EnsureConnectionForActionAsync(
            object originalItem,
            bool allowInteractiveRecovery,
            string context,
            Action<Exception> onFailureObserved = null)
        {
            EnsureSettingsLoaded();
            return CreateSettingsWorkflowController().EnsureConnectionForActionAsync(
                () => IsItemOpenForRibbonAction(originalItem),
                allowInteractiveRecovery,
                context,
                onFailureObserved);
        }

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

        public async void OnFileLinkButtonPressed(IRibbonControl control)
        {
            try
            {
                EnsureSettingsLoaded();
                await _fileLinkLaunchController.OnFileLinkButtonPressedAsync(control);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "FileLink ribbon handler failed.", ex);
            }
        }

        internal Task<bool> RunFileLinkWizardForMailAsync(Outlook.MailItem mail, FileLinkWizardLaunchOptions launchOptions)
        {
            return _fileLinkLaunchController.RunFileLinkWizardForMailAsync(mail, launchOptions);
        }

        internal Task<T> RunOnOutlookUiThreadAsync<T>(Func<T> callback)
        {
            OutlookUiSynchronizationContext context = _uiSynchronizationContext;
            if (context == null)
            {
                var completion = new TaskCompletionSource<T>();
                completion.SetException(new InvalidOperationException("The Outlook UI synchronization context is unavailable."));
                return completion.Task;
            }

            return context.InvokeAsync(callback);
        }

        internal Task RunOnOutlookUiThreadAsync(Action callback)
        {
            if (callback == null)
            {
                throw new ArgumentNullException("callback");
            }

            return RunOnOutlookUiThreadAsync(
                () =>
                {
                    callback();
                    return true;
                });
        }

        internal MailComposeSubscription EnsureMailComposeSubscription(
            Outlook.MailItem mail,
            string inspectorIdentityOverride = null,
            bool isInlineResponse = false,
            string inlineExplorerIdentityOverride = null,
            Outlook.Inspector inspector = null)
        {
            if (mail == null)
            {
                return null;
            }
            if (!IsMailComposeCandidate(mail, "ensure_subscription"))
            {
                return null;
            }
            string mailIdentityKey = ComInteropScope.ResolveIdentityKey(mail, LogCategories.FileLink, "MailItem");
            string inspectorIdentityKey = isInlineResponse
                ? string.Empty
                : (string.IsNullOrWhiteSpace(inspectorIdentityOverride)
                    ? MailInteropController.ResolveMailInspectorIdentityKey(mail)
                    : inspectorIdentityOverride.Trim());
            string inlineExplorerIdentityKey = isInlineResponse && !string.IsNullOrWhiteSpace(inlineExplorerIdentityOverride)
                ? inlineExplorerIdentityOverride.Trim()
                : string.Empty;

            MailComposeSubscription subscription = _mailComposeSubscriptionRegistry.GetOrCreate(
                mail,
                mailIdentityKey,
                inspectorIdentityKey,
                () => new MailComposeSubscription(
                    this,
                    mail,
                    mailIdentityKey,
                    inspectorIdentityKey,
                    isInlineResponse,
                    inlineExplorerIdentityKey));
            if (subscription == null)
            {
                return null;
            }
            if (isInlineResponse)
            {
                subscription.MarkInlineResponse(inlineExplorerIdentityKey);
            }
            else
            {
                subscription.MarkInspector(inspectorIdentityKey);
                subscription.BindInspectorLifecycle(inspector);
            }
            return subscription;
        }

        private void RemoveMailComposeSubscription(MailComposeSubscription subscription)
        {
            _mailComposeSubscriptionRegistry.Remove(subscription);
            if (subscription == null || _inlineResponseSubscriptions.Count == 0)
            {
                return;
            }

            var explorerKeys = new List<string>();
            foreach (var pair in _inlineResponseSubscriptions)
            {
                if (ReferenceEquals(pair.Value, subscription))
                {
                    explorerKeys.Add(pair.Key);
                }
            }
            for (int i = 0; i < explorerKeys.Count; i++)
            {
                _inlineResponseSubscriptions.Remove(explorerKeys[i]);
            }
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

        internal void ShowPasswordMailSuccessNotification(int recipientCount)
        {
            if (recipientCount <= 0)
            {
                return;
            }

            SynchronizationContext notificationUiContext = _uiSynchronizationContext ?? SynchronizationContext.Current;
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
            ShowComposeNotificationOnUiContext(string.Format(
                CultureInfo.CurrentCulture, Strings.SharingPasswordMailNotificationSuccess,
                recipientCount.ToString(CultureInfo.CurrentCulture)), ToolTipIcon.Info, notificationUiContext);
        }

        internal void ShowComposeWarning(string message)
        {
            SynchronizationContext context = _uiSynchronizationContext ?? SynchronizationContext.Current;
            if (context == null || string.IsNullOrWhiteSpace(message))
            {
                return;
            }
            try
            {
                context.Post(_ => ShowComposeNotificationOnUiContext(message, ToolTipIcon.Warning, context), null);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Compose warning could not be displayed.", ex);
            }
        }

        private void ShowComposeNotificationOnUiContext(string message, ToolTipIcon icon, SynchronizationContext notificationUiContext)
        {
            try
            {
                var notifyIcon = new NotifyIcon();
                notifyIcon.Icon = BrandingAssets.GetAppIcon(32);
                notifyIcon.Visible = true;
                notifyIcon.BalloonTipTitle = Strings.SharingPasswordMailNotificationTitle;
                notifyIcon.BalloonTipIcon = icon;
                notifyIcon.BalloonTipText = message;
                notifyIcon.ShowBalloonTip(5000);
                ScheduleNotifyIconDispose(notifyIcon, 7000, notificationUiContext);
                LogCore("Compose notification shown (icon=" + icon + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Compose notification failed on UI context.", ex);
            }
        }

        private void ScheduleNotifyIconDispose(NotifyIcon notifyIcon, int delayMs, SynchronizationContext notificationUiContext)
        {
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
            if (notifyIcon == null)
            {
                return;
            }
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

        internal Outlook.MailItem GetActiveMailItem()
        {
            return _mailInteropController.GetActiveMailItem();
        }

        internal bool IsItemOpenForRibbonAction(object item)
        {
            return _mailInteropController.IsItemOpenForRibbonAction(item);
        }

        internal string ResolveActiveInspectorIdentityKey()
        {
            return _mailInteropController.ResolveActiveInspectorIdentityKey();
        }

        internal bool IsActiveInlineResponse(Outlook.MailItem mail)
        {
            return _mailInteropController.IsActiveInlineResponse(mail);
        }

        internal static bool TryWriteAppointmentHtmlBody(Outlook.AppointmentItem appointment, string html)
        {
            return AppointmentHtmlBodyWriter.TryWriteAppointmentHtmlBody(appointment, html);
        }

        internal Outlook.AppointmentItem GetActiveAppointment()
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
            if (_freeBusyManager == null || _currentSettings == null)
            {
                return;
            }
            try
            {
                LogCore("Applying IFB (Enabled=" + _currentSettings.IfbEnabled + ", Days=" + _currentSettings.IfbDays + ", Port=" + _currentSettings.IfbPort + ", CacheHours=" + _currentSettings.IfbCacheHours + ").");
                string legacyFreeBusyPath = _currentSettings.IfbPreviousFreeBusyPath;
                _freeBusyManager.ApplySettings(_currentSettings);
                if (!string.IsNullOrWhiteSpace(legacyFreeBusyPath)
                    && string.IsNullOrWhiteSpace(_currentSettings.IfbPreviousFreeBusyPath)
                    && _settingsStorage != null)
                {
                    _settingsStorage.Save(_currentSettings);
                }
            }
            catch (Exception ex)
            {
                LogCore("Failed to start IFB: " + ex.Message);
                ShowWarning(string.Format(Strings.WarningIfbStartFailed, ex.Message));
            }
        }

        internal TalkService CreateTalkService()
        {
            return new TalkService(new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword));
        }

        internal bool ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            return _talkAppointmentController.ApplyRoomToAppointment(appointment, request, result);
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

        private string ResolveRoomTokenForAppointment(Outlook.AppointmentItem appointment)
        {
            string roomToken = TalkAppointmentController.GetUserPropertyText(appointment, IcalToken);
            if (!string.IsNullOrWhiteSpace(roomToken))
            {
                return roomToken.Trim();
            }

            return null;
        }

        internal bool ShouldDeleteTalkRoomOnSavedEventDelete()
        {
            bool localEnabled = _currentSettings != null && _currentSettings.TalkDeleteRoomOnEventDelete;
            if (!SettingsAreComplete())
            {
                return localEnabled;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);
            BackendPolicyStatus policyStatus = FetchBackendPolicyStatus(configuration, "talk_delete_room_on_event_delete");
            return _currentSettings.ResolvePolicyDefaults(policyStatus).TalkDeleteRoomOnEventDelete;
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

        private bool IsOrganizer(Outlook.AppointmentItem appointment)
        {
            if (appointment == null)
            {
                return false;
            }

            try
            {
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
            catch (COMException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment organizer state.", ex);
                return false;
            }
        }

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, TalkRoomCreationResult result)
        {
            if (result == null)
            {
                return;
            }

            RegisterSubscription(appointment, result.RoomToken, result.RoomUrl, result.LobbyEnabled, result.CreatedAsEventConversation);
        }

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled, bool isEventConversation)
        {
            string roomUrl = TalkAppointmentController.GetUserPropertyText(appointment, IcalUrl);
            RegisterSubscription(appointment, roomToken, roomUrl, lobbyEnabled, isEventConversation);
        }

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, string roomToken, string roomUrl, bool lobbyEnabled, bool isEventConversation)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }
            string normalizedRoomToken = roomToken.Trim();
            string normalizedRoomUrl = !string.IsNullOrWhiteSpace(roomUrl) ? roomUrl.Trim() : string.Empty;
            LogTalk("Registering appointment subscription (token=" + normalizedRoomToken + ", lobby=" + lobbyEnabled + ", event=" + isEventConversation + ", urlSet=" + !string.IsNullOrWhiteSpace(normalizedRoomUrl) + ").");

            string entryId = GetEntryId(appointment);
            if (!string.IsNullOrEmpty(entryId))
            {
                AppointmentSubscription existingByEntry;
                if (_subscriptionByEntryId.TryGetValue(entryId, out existingByEntry))
                {
                    if (existingByEntry.IsFor(appointment)
                        && existingByEntry.MatchesToken(normalizedRoomToken))
                    {
                        return;
                    }
                }
            }

            AppointmentSubscription existingByToken;
            if (_subscriptionByToken.TryGetValue(normalizedRoomToken, out existingByToken))
            {
                if (existingByToken.IsFor(appointment))
                {
                    return;
                }

                existingByToken.Dispose();
            }
            var key = Guid.NewGuid().ToString("N");
            var subscription = new AppointmentSubscription(this, appointment, key, normalizedRoomToken, normalizedRoomUrl, lobbyEnabled, isEventConversation, entryId);

            AppointmentSubscription staleByEntry;
            if (!string.IsNullOrEmpty(entryId)
                && _subscriptionByEntryId.TryGetValue(entryId, out staleByEntry)
                && staleByEntry != subscription)
            {
                staleByEntry.Dispose();
            }

            _activeSubscriptions[key] = subscription;
            _subscriptionByToken[normalizedRoomToken] = subscription;

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

        internal static List<string> GetAppointmentAttendeeEmails(Outlook.AppointmentItem appointment)
        {
            return OutlookRecipientResolverController.CollectAppointmentAttendeeEmails(appointment);
        }

        private static string TryGetRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            return OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient);
        }

        internal bool TryDeleteRoom(string roomToken, bool isEventConversation)
        {
            return TryDeleteRoom(roomToken, isEventConversation, true);
        }

        internal bool TryDeleteRoom(string roomToken, bool isEventConversation, bool showWarning)
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
                if (showWarning)
                {
                    ShowWarning(string.Format(Strings.WarningRoomDeleteFailed, ex.Message));
                }
            }
            catch (Exception ex)
            {
                LogTalk("Unexpected error while deleting room: " + ex.Message);
                if (showWarning)
                {
                    ShowWarning(string.Format(Strings.WarningRoomDeleteFailed, ex.Message));
                }
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

        internal static void ShowWarningDialog(string message)
        {
            ShowWarning(message);
        }

        internal void EnsureSettingsLoaded()
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
            if (settings != null && !settings.IsManagedLoggingValid)
            {
                DiagnosticsLogger.LogException(LogCategories.Core,
                    "Invalid managed logging value. Invalid DebugLoggingEnabled defaults to false; invalid LogAnonymizationEnabled defaults to true.", null);
            }
            if (settings != null && settings.HasManagedSendPolicyFailureMode && !settings.IsManagedSendPolicyFailureModeValid)
            {
                DiagnosticsLogger.LogException(LogCategories.Core,
                    "Invalid managed SendPolicyFailureMode; failopen applies (source="
                    + settings.ManagedSendPolicyFailureModeSource + ").", null);
            }
        }

        private bool TryApplyTransportSecurityFromSettings(string source, bool showWarning)
        {
            return TryApplyTransportSecurityFromSettings(
                _currentSettings,
                source,
                showWarning);
        }

        private bool TryApplyTransportSecurityFromSettings(
            AddinSettings settings,
            string source,
            bool showWarning)
        {
            try
            {
                TransportSecurityConfigurator.ApplyFromSettings(settings, source);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply transport security settings (source=" + (source ?? string.Empty) + ").",
                    ex);

                if (showWarning || (source == "startup" && settings != null && settings.HasManagedTransportTls))
                {
                    ShowWarning(string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.TransportTlsApplyFailed,
                        ex.Message));
                }
                return false;
            }
        }

    }
}


