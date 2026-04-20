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
        private readonly FileLinkLaunchController _fileLinkLaunchController;
        private readonly TalkRibbonController _talkRibbonController;
        private readonly MailInteropController _mailInteropController;
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
            _fileLinkLaunchController = new FileLinkLaunchController(this);
            _talkRibbonController = new TalkRibbonController(this);
            _mailInteropController = new MailInteropController(this);
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
            await _talkRibbonController.OnTalkButtonPressedAsync(control);
        }

        public async void OnSettingsButtonPressed(IRibbonControl control)
        {
            EnsureSettingsLoaded();
            await CreateSettingsWorkflowController().RunAsync();
        }

        private SettingsWorkflowController CreateSettingsWorkflowController()
        {
            return new SettingsWorkflowController(
                _outlookApplication,
                () => _currentSettings,
                settings => _currentSettings = settings,
                (configuration, trigger) => FetchBackendPolicyStatus(configuration, trigger),
                settings => ConfigureDiagnosticsLogger(settings),
                (source, showWarning) => TryApplyTransportSecurityFromSettings(source, showWarning),
                () => ApplyIfbSettings(),
                settings =>
                {
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollstaendigem Runtime-Zustand kontrolliert abbrechen.
                    if (_settingsStorage != null)
                    {
                        _settingsStorage.Save(settings);
                    }
                },
                message => LogSettings(message));
        }

        public stdole.IPictureDisp OnGetButtonImage(IRibbonControl control)
        {
            string resourceName = "NcTalkOutlookAddIn.Resources.app.png";

            using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollstaendigem Runtime-Zustand kontrolliert abbrechen.
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
            _fileLinkLaunchController.OnFileLinkButtonPressed(control);
        }

        internal bool RunFileLinkWizardForMail(Outlook.MailItem mail, FileLinkWizardLaunchOptions launchOptions)
        {
            return _fileLinkLaunchController.RunFileLinkWizardForMail(mail, launchOptions);
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

        internal MailComposeSubscription EnsureMailComposeSubscription(Outlook.MailItem mail, string inspectorIdentityOverride = null)
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
            return MailInteropController.ResolveMailInspectorIdentityKey(mail);
        }

        private IWin32Window TryCreateMailInspectorDialogOwner(Outlook.MailItem mail)
        {
            return _mailInteropController.TryCreateMailInspectorDialogOwner(mail);
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

            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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
            // Kontext kann auÃŸerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
            if (notifyIcon == null)
            {
                return;
            }

            // Kontext kann auÃŸerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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

        internal string ResolveActiveInspectorIdentityKey()
        {
            return _mailInteropController.ResolveActiveInspectorIdentityKey();
        }

        internal void InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {
            _mailInteropController.InsertHtmlIntoMail(mail, html);
        }

        internal static bool TryWriteAppointmentHtmlBody(Outlook.AppointmentItem appointment, string html)
        {
            return MailInteropController.TryWriteAppointmentHtmlBody(appointment, html);
        }

        internal Outlook.AppointmentItem GetActiveAppointment()
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

        internal TalkService CreateTalkService()
        {
            return new TalkService(new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword));
        }

        internal void ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            _talkAppointmentController.ApplyRoomToAppointment(appointment, request, result);
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

        internal void UpdateStoredServerVersion(string response)
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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

        internal void RegisterSubscription(Outlook.AppointmentItem appointment, TalkRoomCreationResult result)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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
                // Defensiver Null-Guard: dieser Pfad soll bei unvollstÃ¤ndigem Runtime-Zustand kontrolliert abbrechen.
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

    }
}


