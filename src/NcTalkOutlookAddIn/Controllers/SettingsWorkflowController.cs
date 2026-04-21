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
        private readonly Func<string, bool, bool> _applyTransportSecurityFromSettings;
        private readonly Action _applyIfbSettings;
        private readonly Action<AddinSettings> _persistSettings;
        private readonly Action<string> _logSettings;

        internal SettingsWorkflowController(
            Outlook.Application outlookApplication,
            Func<AddinSettings> getCurrentSettings,
            Action<AddinSettings> setCurrentSettings,
            Func<TalkServiceConfiguration, string, BackendPolicyStatus> fetchBackendPolicyStatus,
            Action<AddinSettings> configureDiagnostics,
            Func<string, bool, bool> applyTransportSecurityFromSettings,
            Action applyIfbSettings,
            Action<AddinSettings> persistSettings,
            Action<string> logSettings)
        {
            _outlookApplication = outlookApplication;
            _getCurrentSettings = getCurrentSettings;
            _setCurrentSettings = setCurrentSettings;
            _fetchBackendPolicyStatus = fetchBackendPolicyStatus;
            _configureDiagnostics = configureDiagnostics;
            _applyTransportSecurityFromSettings = applyTransportSecurityFromSettings;
            _applyIfbSettings = applyIfbSettings;
            _persistSettings = persistSettings;
            _logSettings = logSettings;
        }

        internal async Task RunAsync()
        {
            AddinSettings currentSettings = (_getCurrentSettings != null ? _getCurrentSettings() : null) ?? new AddinSettings();
            _logSettings("Settings dialog opened.");

            var configuration = new TalkServiceConfiguration(
                currentSettings.ServerUrl ?? string.Empty,
                currentSettings.Username ?? string.Empty,
                currentSettings.AppPassword ?? string.Empty);

            BackendPolicyStatus initialPolicyStatus = null;
            if (_fetchBackendPolicyStatus != null)
            {
                initialPolicyStatus = await Task.Run(() =>
                    _fetchBackendPolicyStatus(configuration, "settings_open_initial"));
            }

            using (var form = new SettingsForm(currentSettings, _outlookApplication, initialPolicyStatus))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    AddinSettings previousSettings = currentSettings.Clone();
                    AddinSettings nextSettings = form.Result ?? new AddinSettings();

                    if (_setCurrentSettings != null)
                    {
                        _setCurrentSettings(nextSettings);
                    }
                    if (_configureDiagnostics != null)
                    {
                        _configureDiagnostics(nextSettings);
                    }
                    if (_applyTransportSecurityFromSettings != null
                        && !_applyTransportSecurityFromSettings("settings_save", true))
                    {
                        if (_setCurrentSettings != null)
                        {
                            _setCurrentSettings(previousSettings);
                        }
                        if (_configureDiagnostics != null)
                        {
                            _configureDiagnostics(previousSettings);
                        }
                        if (_applyTransportSecurityFromSettings != null)
                        {
                            _applyTransportSecurityFromSettings("settings_save_revert", false);
                        }

                        _logSettings("Settings save aborted because transport security settings could not be applied.");
                        return;
                    }

                    _logSettings(
                        "Settings applied (AuthMode=" + nextSettings.AuthMode
                        + ", IFB=" + nextSettings.IfbEnabled
                        + ", IfbPort=" + nextSettings.IfbPort
                        + ", Debug=" + nextSettings.DebugLoggingEnabled
                        + ", LogAnonymize=" + nextSettings.LogAnonymizationEnabled
                        + ").");

                    if (_applyIfbSettings != null)
                    {
                        _applyIfbSettings();
                    }
                    if (_persistSettings != null)
                    {
                        _persistSettings(nextSettings);
                    }
                }
                else
                {
                    _logSettings("Settings dialog closed without changes.");
                }
            }
        }
    }
}
