// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Add-in lifecycle and bootstrap/teardown flow.
    public sealed partial class NextcloudTalkAddIn
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            _uiSynchronizationContext = SynchronizationContext.Current;
            string outlookProfileName = ResolveCurrentOutlookProfileName();
            _settingsStorage = new NcTalkOutlookAddIn.Settings.SettingsStorage(outlookProfileName);
            _currentSettings = _settingsStorage.Load();
            ConfigureDiagnosticsLogger(_currentSettings);
            TryApplyTransportSecurityFromSettings("startup", false);
            TryApplyOfficeUiLanguage();
            LogCore("Add-in connected (Outlook version=" + (_outlookApplication != null ? _outlookApplication.Version : "unknown") + ").");
            if (!string.IsNullOrWhiteSpace(outlookProfileName))
            {
                LogCore("Using Outlook profile settings: " + outlookProfileName + ".");
            }
            if (_currentSettings != null)
            {
                LogSettings("Settings loaded (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", IfbPort=" + _currentSettings.IfbPort + ", Debug=" + _currentSettings.DebugLoggingEnabled + ", LogAnonymize=" + _currentSettings.LogAnonymizationEnabled + ").");
            }

            _freeBusyManager = new FreeBusyManager(_settingsStorage.DataDirectory);
            _freeBusyManager.Initialize(_outlookApplication);
            EnsureApplicationHook();
            EnsureInspectorHook();
            ApplyIfbSettings();
            StartUpdateCheckIfDue();
        }

        private void StartUpdateCheckIfDue()
        {
            if (_currentSettings == null || _settingsStorage == null)
            {
                return;
            }

            bool hadInstallId = !string.IsNullOrWhiteSpace(_currentSettings.UpdateInstallId);
            UpdateCheckService.EnsureInstallId(_currentSettings);
            if (!hadInstallId)
            {
                _settingsStorage.Save(_currentSettings);
            }

            AddinSettings updateSettings = _currentSettings.Clone();
            Task.Run(async () =>
            {
                try
                {
                    UpdateCheckResult result = await _updateCheckService.CheckAsync(updateSettings, false).ConfigureAwait(false);
                    StoreUpdateCheckSettings(updateSettings);
                    if (UpdateCheckService.ShouldNotify(updateSettings, result))
                    {
                        PostUpdateNotification(result);
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Update check failed.", ex);
                }
            });
        }

        private void StoreUpdateCheckSettings(AddinSettings updateSettings)
        {
            if (updateSettings == null || _currentSettings == null || _settingsStorage == null)
            {
                return;
            }

            _currentSettings.UpdateInstallId = updateSettings.UpdateInstallId ?? string.Empty;
            _currentSettings.UpdateLastCheckedAtUtc = updateSettings.UpdateLastCheckedAtUtc ?? string.Empty;
            _currentSettings.UpdateLatestVersion = updateSettings.UpdateLatestVersion ?? string.Empty;
            _currentSettings.UpdateReleaseUrl = updateSettings.UpdateReleaseUrl ?? string.Empty;
            _currentSettings.UpdateDownloadUrl = updateSettings.UpdateDownloadUrl ?? string.Empty;
            _currentSettings.UpdatePublishedAt = updateSettings.UpdatePublishedAt ?? string.Empty;
            _currentSettings.UpdateChangelogTitle = updateSettings.UpdateChangelogTitle ?? string.Empty;
            _currentSettings.UpdateChangelogText = updateSettings.UpdateChangelogText ?? string.Empty;
            _settingsStorage.Save(_currentSettings);
        }

        private void PostUpdateNotification(UpdateCheckResult result)
        {
            SynchronizationContext context = _uiSynchronizationContext;
            if (context == null)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Update notification skipped because no UI context is available.");
                return;
            }

            context.Post(_ => ShowUpdateNotification(result), null);
        }

        private void ShowUpdateNotification(UpdateCheckResult result)
        {
            try
            {
                if (_currentSettings == null || !UpdateCheckService.ShouldNotify(_currentSettings, result))
                {
                    return;
                }

                UpdateCheckService.MarkNotified(_currentSettings, result);
                if (_settingsStorage != null)
                {
                    _settingsStorage.Save(_currentSettings);
                }

                DialogResult answer = MessageBox.Show(
                    UpdateCheckService.BuildNotificationMessage(result),
                    Strings.UpdateAvailableTitle,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (answer != DialogResult.Yes)
                {
                    return;
                }

                string url = UpdateCheckService.GetPreferredOpenUrl(result);
                if (!string.IsNullOrWhiteSpace(url))
                {
                    BrowserLauncher.OpenUrl(
                        url,
                        LogCategories.Core,
                        "Failed to open update download URL.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Update notification failed.", ex);
            }
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
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Core,
                    "Failed to release Outlook session COM object after profile resolution.");
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            TearDownAddInState("disconnect", true);
            LogCore("Add-in disconnected (removeMode=" + removeMode + ").");
        }

        // IDTExtensibility2 requires this callback; runtime wiring happens in OnConnection
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        // IDTExtensibility2 requires this callback; startup work happens in OnConnection
        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            TearDownAddInState("shutdown", false);
        }

        // Outlook can call both shutdown callbacks, so teardown must stay idempotent
        private void TearDownAddInState(string origin, bool clearOutlookApplication)
        {
            UnhookApplication();
            UnhookInspector();
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
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "Failed to disable IFB during add-in " + (origin ?? "teardown") + ".",
                        ex);
                }
            }
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
            _deferredAppointmentEnsureState.ClearPendingKeys();
        }
    }
}
