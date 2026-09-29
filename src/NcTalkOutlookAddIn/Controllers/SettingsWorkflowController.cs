// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Encapsulates the full settings dialog workflow (open, apply, revert on TLS failures, persist).
    // Keeps the add-in host focused on orchestration and COM lifecycle code.
    internal sealed class SettingsWorkflowController
    {
        private readonly Outlook.Application _outlookApplication;
        private readonly Func<AddinSettings> _getCurrentSettings;
        private readonly Action<AddinSettings> _setCurrentSettings;
        private readonly Func<TalkServiceConfiguration, string, BackendPolicyStatus> _fetchBackendPolicyStatus;
        private readonly Action<AddinSettings> _configureDiagnostics;
        private readonly Func<AddinSettings, string, bool, bool> _applyTransportSecurityFromSettings;
        private readonly Action _applyIfbSettings;
        private readonly Action<AddinSettings> _persistSettings;
        private readonly Func<Action, Task> _runOnOutlookUiThreadAsync;
        private readonly Action<string> _logSettings;
        private readonly string _dataDirectory;
        private readonly string _outlookProfileScope;

        internal SettingsWorkflowController(
            Outlook.Application outlookApplication,
            Func<AddinSettings> getCurrentSettings,
            Action<AddinSettings> setCurrentSettings,
            Func<TalkServiceConfiguration, string, BackendPolicyStatus> fetchBackendPolicyStatus,
            Action<AddinSettings> configureDiagnostics,
            Func<AddinSettings, string, bool, bool> applyTransportSecurityFromSettings,
            Action applyIfbSettings,
            Action<AddinSettings> persistSettings,
            Func<Action, Task> runOnOutlookUiThreadAsync,
            Action<string> logSettings,
            string dataDirectory,
            string outlookProfileScope)
        {
            _outlookApplication = outlookApplication;
            _getCurrentSettings = getCurrentSettings;
            _setCurrentSettings = setCurrentSettings;
            _fetchBackendPolicyStatus = fetchBackendPolicyStatus;
            _configureDiagnostics = configureDiagnostics;
            _applyTransportSecurityFromSettings = applyTransportSecurityFromSettings;
            _applyIfbSettings = applyIfbSettings;
            _persistSettings = persistSettings;
            _runOnOutlookUiThreadAsync = runOnOutlookUiThreadAsync;
            _logSettings = logSettings;
            _dataDirectory = dataDirectory ?? string.Empty;
            _outlookProfileScope = outlookProfileScope ?? string.Empty;
        }

        internal async Task<bool> RunAsync(bool requireAuthentication = false, bool authenticationRejected = false)
        {
            AddinSettings currentSettings = ((_getCurrentSettings != null ? _getCurrentSettings() : null) ?? new AddinSettings()).Clone();
            currentSettings.ApplyManagedSetupPolicy(ManagedSetupPolicy.Load());
            _logSettings("Settings dialog opened.");
            if (currentSettings.HasManagedNextcloudUrl)
            {
                _logSettings(
                    "Managed Nextcloud URL policy active (source=" + currentSettings.ManagedNextcloudUrlSource
                    + ", locked=" + currentSettings.ManagedNextcloudUrlLocked + ").");
            }

            var configuration = new TalkServiceConfiguration(
                currentSettings.ServerUrl ?? string.Empty,
                currentSettings.Username ?? string.Empty,
                currentSettings.AppPassword ?? string.Empty);

            var addressBookCache = new IfbAddressBookCache(
                _dataDirectory,
                _outlookProfileScope);
            Task<BackendPolicyStatus> policyStatusTask =
                configuration.IsComplete() && !requireAuthentication && _fetchBackendPolicyStatus != null
                    ? Task.Run(
                        () => _fetchBackendPolicyStatus(
                            configuration,
                            "settings_open_initial"))
                    : Task.FromResult<BackendPolicyStatus>(null);
            Task<IfbAddressBookCache.SystemAddressbookStatus> addressbookStatusTask =
                configuration.IsComplete() && !requireAuthentication ? Task.Run(
                    () => addressBookCache.GetSystemAddressbookStatus(
                        configuration,
                        currentSettings.IfbCacheHours,
                        false)) : Task.FromResult<IfbAddressBookCache.SystemAddressbookStatus>(null);
            await Task.WhenAll(
                    policyStatusTask,
                    addressbookStatusTask)
                .ConfigureAwait(false);
            BackendPolicyStatus initialPolicyStatus =
                await policyStatusTask.ConfigureAwait(false);
            IfbAddressBookCache.SystemAddressbookStatus initialAddressbookStatus =
                await addressbookStatusTask.ConfigureAwait(false);

            if (_runOnOutlookUiThreadAsync == null)
            {
                throw new InvalidOperationException("The Outlook UI-thread dispatcher is unavailable.");
            }

            bool saved = false;
            await _runOnOutlookUiThreadAsync(
                () => saved = RunSettingsDialogOnUiThread(
                    currentSettings,
                    initialPolicyStatus,
                    addressBookCache,
                    initialAddressbookStatus,
                    requireAuthentication,
                    authenticationRejected)).ConfigureAwait(false);
            return saved;
        }

        private bool RunSettingsDialogOnUiThread(
            AddinSettings currentSettings,
            BackendPolicyStatus initialPolicyStatus,
            IfbAddressBookCache addressBookCache,
            IfbAddressBookCache.SystemAddressbookStatus initialAddressbookStatus,
            bool requireAuthentication,
            bool authenticationRejected)
        {
            // SettingsForm owns Outlook COM references and async WinForms handlers, so its complete
            // modal lifetime must begin on the Outlook STA thread captured during add-in startup.
            using (var form = new SettingsForm(
                currentSettings,
                _outlookApplication,
                initialPolicyStatus,
                addressBookCache,
                initialAddressbookStatus))
            {
                if (requireAuthentication)
                {
                    form.BeginAuthentication(authenticationRejected);
                }
                if (form.ShowDialog() == DialogResult.OK)
                {
                    AddinSettings previousSettings = currentSettings.Clone();
                    AddinSettings nextSettings = (form.Result ?? new AddinSettings()).Clone();

                    if (!ValidateTransportSecurityBeforeSave(previousSettings, nextSettings))
                    {
                        _logSettings("Settings save aborted because transport security settings could not be applied.");
                        return false;
                    }

                    if (_persistSettings == null)
                    {
                        throw new InvalidOperationException("The settings persistence service is unavailable.");
                    }

                    try
                    {
                        _persistSettings(nextSettings);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Settings save failed; runtime settings were not changed.",
                            ex);
                        MessageBox.Show(
                            form,
                            Strings.SettingsSaveFailed
                            + Environment.NewLine
                            + Environment.NewLine
                            + (ex.Message ?? string.Empty),
                            Strings.SettingsFormTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return false;
                    }

                    ApplyRuntimeSettings(nextSettings);

                    if (_applyTransportSecurityFromSettings != null
                        && !_applyTransportSecurityFromSettings(
                            nextSettings,
                            "settings_save_commit",
                            false))
                    {
                        try
                        {
                            _persistSettings(previousSettings);
                        }
                        finally
                        {
                            ApplyRuntimeSettings(previousSettings);
                            _applyTransportSecurityFromSettings(
                                previousSettings,
                                "settings_save_commit_revert",
                                false);
                        }
                        _logSettings("Settings save reverted because transport security settings could not be committed.");
                        return false;
                    }

                    if (_applyIfbSettings != null)
                    {
                        _applyIfbSettings();
                    }

                    _logSettings(
                        "Settings applied (AuthMode=" + nextSettings.AuthMode
                        + ", IFB=" + nextSettings.IfbEnabled
                        + ", IfbPort=" + nextSettings.IfbPort
                        + ", Debug=" + nextSettings.DebugLoggingEnabled
                        + ", LogAnonymize=" + nextSettings.LogAnonymizationEnabled
                        + ").");
                    return true;
                }
                else
                {
                    _logSettings("Settings dialog closed without changes.");
                }
            }
            return false;
        }

        private bool ValidateTransportSecurityBeforeSave(
            AddinSettings previousSettings,
            AddinSettings nextSettings)
        {
            if (_applyTransportSecurityFromSettings == null)
            {
                return true;
            }

            bool valid = _applyTransportSecurityFromSettings(
                nextSettings,
                "settings_save_validate",
                true);
            _applyTransportSecurityFromSettings(
                previousSettings,
                "settings_save_validate_revert",
                false);
            return valid;
        }

        private void ApplyRuntimeSettings(AddinSettings settings)
        {
            if (_setCurrentSettings != null)
            {
                _setCurrentSettings(settings);
            }
            if (_configureDiagnostics != null)
            {
                _configureDiagnostics(settings);
            }
        }
    }
}
