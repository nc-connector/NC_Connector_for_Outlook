// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
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
        private readonly Action<AddinSettings> _removeSavedCredentials;
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
            string outlookProfileScope,
            Action<AddinSettings> removeSavedCredentials = null)
        {
            _outlookApplication = outlookApplication;
            _getCurrentSettings = getCurrentSettings;
            _setCurrentSettings = setCurrentSettings;
            _fetchBackendPolicyStatus = fetchBackendPolicyStatus;
            _configureDiagnostics = configureDiagnostics;
            _applyTransportSecurityFromSettings = applyTransportSecurityFromSettings;
            _applyIfbSettings = applyIfbSettings;
            _persistSettings = persistSettings;
            _removeSavedCredentials = removeSavedCredentials;
            _runOnOutlookUiThreadAsync = runOnOutlookUiThreadAsync;
            _logSettings = logSettings;
            _dataDirectory = dataDirectory ?? string.Empty;
            _outlookProfileScope = outlookProfileScope ?? string.Empty;
        }

        internal async Task<bool> RunAsync(
            bool requireAuthentication = false,
            bool authenticationRejected = false,
            Func<bool> canOpenDialog = null)
        {
            if (_runOnOutlookUiThreadAsync == null)
            {
                throw new InvalidOperationException("The Outlook UI-thread dispatcher is unavailable.");
            }
            AddinSettings currentSettings = null;
            await _runOnOutlookUiThreadAsync(() =>
            {
                if (canOpenDialog == null || canOpenDialog())
                {
                    currentSettings = ((_getCurrentSettings != null ? _getCurrentSettings() : null)
                        ?? new AddinSettings()).Clone();
                }
            }).ConfigureAwait(false);
            if (currentSettings == null) { return false; }
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

            bool saved = false;
            await _runOnOutlookUiThreadAsync(
                () => saved = (canOpenDialog == null || canOpenDialog()) && RunSettingsDialogOnUiThread(
                    currentSettings,
                    initialPolicyStatus,
                    addressBookCache,
                    initialAddressbookStatus,
                    requireAuthentication,
                    authenticationRejected)).ConfigureAwait(false);
            return saved;
        }

        internal async Task<AddinSettings> EnsureConnectionForActionAsync(
            Func<bool> isOriginalItemOpen,
            bool allowInteractiveRecovery,
            string context,
            Action<Exception> onFailureObserved = null)
        {
            if (_runOnOutlookUiThreadAsync == null)
            {
                throw new InvalidOperationException("The Outlook UI-thread dispatcher is unavailable.");
            }
            for (int attempt = 0; attempt < 2; attempt++)
            {
                AddinSettings settings = null;
                await _runOnOutlookUiThreadAsync(() =>
                {
                    if (isOriginalItemOpen != null && isOriginalItemOpen())
                    {
                        AddinSettings current = _getCurrentSettings != null ? _getCurrentSettings() : null;
                        settings = current != null ? current.Clone() : null;
                    }
                }).ConfigureAwait(false);
                if (settings == null)
                {
                    _logSettings(context + " connection check ended: original item unavailable.");
                    return null;
                }

                var configuration = new TalkServiceConfiguration(
                    settings.ServerUrl, settings.Username, settings.AppPassword);
                Func<bool> canOpenSettings = () => isOriginalItemOpen != null
                    && isOriginalItemOpen() && ConnectionSettingsMatch(settings);
                if (!configuration.IsComplete() && allowInteractiveRecovery && attempt == 0)
                {
                    if (!await RunAsync(true, false, canOpenSettings).ConfigureAwait(false)) { return null; }
                    continue;
                }

                Exception failure = null;
                try
                {
                    string response = string.Empty;
                    _logSettings(context + " fresh connection check started.");
                    bool verified = await Task.Run(() =>
                        new TalkService(configuration).VerifyConnection(out response)).ConfigureAwait(false);
                    if (!verified)
                    {
                        failure = new TalkServiceException(
                            string.IsNullOrWhiteSpace(response) ? Strings.ErrorCredentialsNotVerified : response,
                            false, 0, null);
                    }
                }
                catch (OperationCanceledException)
                {
                    return null;
                }
                catch (Exception ex)
                {
                    failure = ex;
                }

                bool canContinue = false;
                await _runOnOutlookUiThreadAsync(() =>
                    canContinue = isOriginalItemOpen != null && isOriginalItemOpen()
                        && ConnectionSettingsMatch(settings)).ConfigureAwait(false);
                if (!canContinue)
                {
                    _logSettings(context + " connection check discarded: original item or credentials changed.");
                    return null;
                }
                if (failure == null)
                {
                    _logSettings(context + " fresh connection check succeeded.");
                    return settings;
                }
                if (onFailureObserved != null) { onFailureObserved(failure); }
                _logSettings(context + " fresh connection check failed: " + failure.Message);

                var serviceFailure = failure as TalkServiceException;
                bool authenticationRejected = serviceFailure != null && serviceFailure.IsAuthenticationError;
                bool canRecover = allowInteractiveRecovery && attempt == 0;
                if (authenticationRejected && canRecover)
                {
                    if (!await RunAsync(true, true, canOpenSettings).ConfigureAwait(false)) { return null; }
                    continue;
                }

                string message = GetActionConnectionFailureMessage(failure);
                DialogResult result = DialogResult.Cancel;
                await _runOnOutlookUiThreadAsync(() =>
                {
                    if (isOriginalItemOpen == null || !isOriginalItemOpen() || !ConnectionSettingsMatch(settings))
                    {
                        return;
                    }
                    result = MessageBox.Show(
                        canRecover ? string.Format(CultureInfo.CurrentCulture, Strings.PromptOpenSettings, message) : message,
                        Strings.DialogTitle,
                        canRecover ? MessageBoxButtons.YesNo : MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }).ConfigureAwait(false);
                if (result != DialogResult.Yes
                    || !await RunAsync(true, false, canOpenSettings).ConfigureAwait(false)) { return null; }
            }
            return null;
        }

        internal static string GetActionConnectionFailureMessage(Exception failure)
        {
            string connectionNotice = PolicyUiHelper.GetConnectionFailureMessage(failure);
            if (!string.IsNullOrEmpty(connectionNotice)) { return connectionNotice; }
            var serviceFailure = failure as TalkServiceException;
            if (serviceFailure != null && serviceFailure.IsTransportError)
            {
                return Strings.ErrorServerUnavailable;
            }
            return string.Format(CultureInfo.CurrentCulture, Strings.ErrorConnectionFailed, failure.Message);
        }

        private bool ConnectionSettingsMatch(AddinSettings verifiedSettings)
        {
            AddinSettings current = _getCurrentSettings != null ? _getCurrentSettings() : null;
            return ConnectionSettingsMatch(current, verifiedSettings);
        }

        internal static bool ConnectionSettingsMatch(AddinSettings current, AddinSettings verifiedSettings)
        {
            return current != null && verifiedSettings != null
                && string.Equals(current.ServerUrl, verifiedSettings.ServerUrl, StringComparison.Ordinal)
                && string.Equals(current.Username, verifiedSettings.Username, StringComparison.Ordinal)
                && string.Equals(current.AppPassword, verifiedSettings.AppPassword, StringComparison.Ordinal);
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
                initialAddressbookStatus,
                _fetchBackendPolicyStatus,
                _removeSavedCredentials != null ? new Func<AddinSettings>(RemoveSavedCredentials) : null))
            {
                if (requireAuthentication)
                {
                    form.BeginAuthentication(authenticationRejected);
                }
                if (form.ShowDialog() == DialogResult.OK)
                {
                    AddinSettings previousSettings = ((_getCurrentSettings != null ? _getCurrentSettings() : null)
                        ?? currentSettings).Clone();
                    AddinSettings nextSettings = (form.Result ?? new AddinSettings()).Clone();
                    var configuration = new TalkServiceConfiguration(
                        nextSettings.ServerUrl, nextSettings.Username, nextSettings.AppPassword);
                    bool verified = NextcloudConnectionState.HasFreshVerification(configuration);
                    bool verificationRequired = requireAuthentication
                        || !ConnectionSettingsMatch(previousSettings, nextSettings)
                        || string.Equals(NextcloudConnectionState.GetStatus(configuration).Reason,
                            "auth_required", StringComparison.Ordinal);
                    if (verificationRequired && !verified)
                    {
                        MessageBox.Show(Strings.ErrorCredentialsNotVerified, Strings.SettingsFormTitle,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }

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

                    try
                    {
                        ApplyRuntimeSettings(nextSettings);
                        if (verified) { NextcloudConnectionState.CommitVerifiedSavedConnection(configuration); }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Verified connection could not be committed.", ex);
                        try { _persistSettings(previousSettings); }
                        finally
                        {
                            ApplyRuntimeSettings(previousSettings);
                            if (_applyTransportSecurityFromSettings != null)
                            {
                                _applyTransportSecurityFromSettings(previousSettings, "settings_connection_commit_revert", false);
                            }
                        }
                        MessageBox.Show(Strings.SettingsSaveFailed, Strings.SettingsFormTitle,
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private AddinSettings RemoveSavedCredentials()
        {
            if (_removeSavedCredentials == null)
            {
                throw new InvalidOperationException("Credential removal is unavailable.");
            }
            AddinSettings nextSettings = ((_getCurrentSettings != null ? _getCurrentSettings() : null)
                ?? new AddinSettings()).Clone();
            nextSettings.Username = string.Empty;
            nextSettings.AppPassword = string.Empty;
            _removeSavedCredentials(nextSettings);
            try
            {
                NextcloudConnectionState.RemoveSavedCredentials(new TalkServiceConfiguration(
                    nextSettings.ServerUrl, string.Empty, string.Empty));
            }
            finally
            {
                ApplyRuntimeSettings(nextSettings);
            }
            if (_applyIfbSettings != null) { _applyIfbSettings(); }
            _logSettings("Saved credentials removed; other settings and pending cleanup retained.");
            return nextSettings;
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
