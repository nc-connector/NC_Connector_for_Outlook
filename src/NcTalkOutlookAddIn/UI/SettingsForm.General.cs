// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private void InitializeGeneralTab()
        {
            _generalTab.AutoScroll = true;
            _generalTab.Resize += (s, e) => ApplyGeneralTabFieldSizing();

            _serverUrlTextBox.TextChanged += OnGeneralValueChanged;
            _usernameTextBox.TextChanged += OnGeneralValueChanged;
            _appPasswordTextBox.TextChanged += OnGeneralValueChanged;

            var serverLabel = new Label
            {
                Text = Strings.LabelServerUrl,
                Location = new Point(15, 20),
                AutoSize = true
            };
            _generalTab.Controls.Add(serverLabel);

            _serverUrlTextBox.Location = new Point(150, 16);
            _serverUrlTextBox.Width = 280;
            _serverUrlTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _generalTab.Controls.Add(_serverUrlTextBox);

            var authGroup = new GroupBox
            {
                Text = Strings.GroupAuthentication,
                Location = new Point(18, 55),
                Size = new Size(440, 120),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _generalTab.Controls.Add(authGroup);

            _manualRadio.Text = Strings.RadioManual;
            _manualRadio.Location = new Point(12, 25);
            _manualRadio.AutoSize = true;
            _manualRadio.CheckedChanged += OnGeneralValueChanged;
            authGroup.Controls.Add(_manualRadio);

            _loginFlowRadio.Text = Strings.RadioLoginFlow;
            _loginFlowRadio.Location = new Point(12, 55);
            _loginFlowRadio.AutoSize = true;
            _loginFlowRadio.CheckedChanged += OnGeneralValueChanged;
            authGroup.Controls.Add(_loginFlowRadio);

            _loginFlowButton.Text = Strings.ButtonLoginFlow;
            _loginFlowButton.Location = new Point(12, 85);
            _loginFlowButton.Width = 200;
            _loginFlowButton.Click += OnLoginFlowButtonClick;
            authGroup.Controls.Add(_loginFlowButton);

            _testButton.Text = Strings.ButtonTestConnection;
            _testButton.Location = new Point(224, 85);
            _testButton.Width = 200;
            _testButton.Click += OnTestButtonClick;
            authGroup.Controls.Add(_testButton);

            var userLabel = new Label
            {
                Text = Strings.LabelUsername,
                Location = new Point(15, 200),
                AutoSize = true
            };
            _generalTab.Controls.Add(userLabel);

            _usernameTextBox.Location = new Point(150, 196);
            _usernameTextBox.Width = 280;
            _usernameTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _generalTab.Controls.Add(_usernameTextBox);

            var passwordLabel = new Label
            {
                Text = Strings.LabelAppPassword,
                Location = new Point(15, 235),
                AutoSize = true
            };
            _generalTab.Controls.Add(passwordLabel);

            _appPasswordTextBox.Location = new Point(150, 231);
            _appPasswordTextBox.Width = 280;
            _appPasswordTextBox.UseSystemPasswordChar = true;
            _appPasswordTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _generalTab.Controls.Add(_appPasswordTextBox);

            ApplyGeneralTabFieldSizing();
        }

        private void ApplyGeneralTabFieldSizing()
        {
            const int fieldLeft = 150;
            const int rightMargin = 24;
            const int minWidth = 260;

            int availableWidth = _generalTab.ClientSize.Width - fieldLeft - rightMargin;
            int width = Math.Max(minWidth, availableWidth);

            _serverUrlTextBox.Width = width;
            _usernameTextBox.Width = width;
            _appPasswordTextBox.Width = width;
        }

        private async void OnLoginFlowButtonClick(object sender, EventArgs e)
        {
            await StartLoginFlowAsync();
        }

        private bool ShouldStartManagedLoginFlow()
        {
            if (Result == null || !Result.HasManagedAuthMode || !Result.IsManagedAuthModeValid
                || Result.AuthMode != AuthenticationMode.LoginFlow || !Result.HasManagedNextcloudUrl)
            {
                return false;
            }
            string normalizedUrl;
            return NextcloudUriValidator.TryNormalizeBaseUrl(_serverUrlTextBox.Text, out normalizedUrl)
                && string.Equals(normalizedUrl, Result.ManagedNextcloudUrl, StringComparison.Ordinal)
                && !new TalkServiceConfiguration(normalizedUrl, _usernameTextBox.Text, _appPasswordTextBox.Text).IsComplete();
        }

        private async Task StartLoginFlowAsync()
        {
            if (_isBusy || IsDisposed || Disposing || _manualRadio.Checked)
            {
                return;
            }
            string baseUrl = _serverUrlTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                SetStatus(Strings.StatusServerUrlRequired, true);
                return;
            }
            var normalizedUrl = new TalkServiceConfiguration(baseUrl, string.Empty, string.Empty).GetNormalizedBaseUrl();
            if (string.IsNullOrWhiteSpace(normalizedUrl))
            {
                SetStatus(Strings.StatusInvalidServerUrl, true);
                return;
            }

            _serverUrlTextBox.Text = normalizedUrl;

            SetBusy(true);
            SetStatus(Strings.StatusLoginFlowStarting, false);
            SecurityProtocolType previousSecurityProtocol = ServicePointManager.SecurityProtocol;
            bool temporaryTlsApplied = false;
            bool loginVerified = false;

            try
            {
                previousSecurityProtocol = ApplySelectedTransportSecurity("settings_login_flow");
                temporaryTlsApplied = true;

                var flowService = new TalkLoginFlowService(normalizedUrl);
                var startInfo = await Task.Run(() => flowService.StartLoginFlow());
                if (IsDisposed || Disposing)
                {
                    return;
                }

                BrowserLauncher.OpenUrl(
                    startInfo.LoginUrl,
                    LogCategories.Core,
                    "Failed to open browser for login flow URL.");
                SetStatus(Strings.StatusLoginFlowBrowser, false);

                var credentials = await Task.Run(() => flowService.CompleteLoginFlow(startInfo, TimeSpan.FromMinutes(2), TimeSpan.FromSeconds(2)));
                if (IsDisposed || Disposing)
                {
                    return;
                }
                _usernameTextBox.Text = credentials.LoginName ?? string.Empty;
                _appPasswordTextBox.Text = credentials.AppPassword ?? string.Empty;
                _loginFlowRadio.Checked = true;
                var verificationService = new TalkService(new TalkServiceConfiguration(
                    normalizedUrl,
                    _usernameTextBox.Text,
                    _appPasswordTextBox.Text));
                string versionResponse = string.Empty;
                bool verified = await Task.Run(() => verificationService.VerifyConnection(out versionResponse));
                if (IsDisposed || Disposing)
                {
                    return;
                }
                if (!verified)
                {
                    string failureMessage = string.IsNullOrWhiteSpace(versionResponse)
                        ? Strings.ErrorCredentialsNotVerified
                        : versionResponse;
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Login flow credential verification failed: "
                        + failureMessage);
                    SetStatus(
                        string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.StatusLoginFlowFailure,
                            failureMessage),
                        true);
                    return;
                }

                _connectionSetupPending = false;
                _authenticationRejected = false;
                ApplyBackendPolicyStatus("login_verified");
                SetStatus(Strings.StatusLoginFlowSuccess, false);
                loginVerified = true;
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Login flow failed.", ex);
                HandleServiceFailure(Strings.StatusLoginFlowFailure, ex);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Login flow failed unexpectedly.", ex);
                SetStatus(string.Format(Strings.StatusLoginFlowFailure, ex.Message), true);
            }
            finally
            {
                if (temporaryTlsApplied)
                {
                    RestoreTemporaryTls(previousSecurityProtocol, "settings_login_flow");
                }
                SetBusy(false);
            }

            if (loginVerified && !IsDisposed && !Disposing && _authenticationRequired
                && Result != null && Result.HasManagedAuthMode && Result.IsManagedAuthModeValid
                && Result.AuthMode == AuthenticationMode.LoginFlow)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Saving verified managed login credentials for the pending action.");
                await SaveSettingsWithErrorHandlingAsync();
            }
        }

        // WinForms event handlers must stay async void; keep awaited flow inside this method-level try/catch.
        private async void OnTestButtonClick(object sender, EventArgs e)
        {
            await TestConnectionAsync();
        }

        private async Task<bool> TestConnectionAsync()
        {
            if (_isBusy)
            {
                return false;
            }
            string baseUrl = _serverUrlTextBox.Text.Trim();
            string user = _usernameTextBox.Text.Trim();
            string appPassword = _appPasswordTextBox.Text;

            if (string.IsNullOrWhiteSpace(baseUrl) ||
                string.IsNullOrWhiteSpace(user) ||
                string.IsNullOrEmpty(appPassword))
            {
                SetStatus(Strings.StatusMissingFields, true);
                return false;
            }
            string normalizedUrl;
            if (!NextcloudUriValidator.TryNormalizeBaseUrl(baseUrl, out normalizedUrl))
            {
                SetStatus(Strings.StatusInvalidServerUrl, true);
                return false;
            }
            _serverUrlTextBox.Text = normalizedUrl;

            SetBusy(true);
            SetStatus(Strings.StatusTestRunning, false);
            DiagnosticsLogger.Log(LogCategories.Core, "Connection test started (Server=" + normalizedUrl + ", User=" + user + ").");
            SecurityProtocolType previousSecurityProtocol = ServicePointManager.SecurityProtocol;
            bool temporaryTlsApplied = false;

            try
            {
                previousSecurityProtocol = ApplySelectedTransportSecurity("settings_connection_test");
                temporaryTlsApplied = true;

                var service = new TalkService(new TalkServiceConfiguration(normalizedUrl, user, appPassword));
                string responseMessage = string.Empty;
                bool success = await Task.Run(() => service.VerifyConnection(out responseMessage));
                if (IsDisposed || Disposing)
                {
                    return false;
                }
                if (success)
                {
                    _connectionSetupPending = false;
                    _authenticationRejected = false;
                    ApplyBackendPolicyStatus("connection_verified");
                    DiagnosticsLogger.Log(LogCategories.Core, "Connection test succeeded (Response=" + (string.IsNullOrEmpty(responseMessage) ? "OK" : responseMessage) + ").");
                    string suffix = string.IsNullOrEmpty(responseMessage)
                        ? string.Empty
                        : " (" + string.Format(Strings.StatusTestSuccessVersionFormat, responseMessage) + ")";
                    SetStatus(string.Format(Strings.StatusTestSuccessFormat, suffix), false);
                    return true;
                }
                else
                {
                    var failureMessage = string.IsNullOrEmpty(responseMessage)
                        ? Strings.StatusTestFailureUnknown
                        : responseMessage;
                    DiagnosticsLogger.Log(LogCategories.Core, "Connection test failed: " + failureMessage);
                    SetStatus(string.Format(Strings.StatusTestFailure, failureMessage), true);
                }
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Connection test failed with service error.", ex);
                HandleServiceFailure(Strings.StatusTestFailure, ex);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Connection test failed unexpectedly.", ex);
                SetStatus(string.Format(Strings.StatusTestFailure, ex.Message), true);
            }
            finally
            {
                if (temporaryTlsApplied)
                {
                    RestoreTemporaryTls(previousSecurityProtocol, "settings_connection_test");
                }
                SetBusy(false);
            }
            return false;
        }

        private SecurityProtocolType ApplySelectedTransportSecurity(string source)
        {
            SecurityProtocolType previous = ServicePointManager.SecurityProtocol;
            try
            {
                AddinSettings selected = Result.Clone();
                if (!selected.HasManagedTransportTls)
                {
                    selected.TransportTlsUseSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
                    selected.TransportTlsEnable12 = _tlsEnable12CheckBox.Checked;
                    selected.TransportTlsEnable13 = _tlsEnable13CheckBox.Checked;
                }
                TransportSecurityConfigurator.ApplyFromSettings(selected, source);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply temporary TLS settings (source=" + (source ?? string.Empty) + ").",
                    ex);
                throw;
            }
            return previous;
        }

        private static void RestoreTemporaryTls(SecurityProtocolType previous, string source)
        {
            TransportSecurityConfigurator.Restore(previous);
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Transport security restored after temporary settings operation (source="
                + (source ?? string.Empty)
                + ", securityProtocol="
                + previous
                + ").");
        }

        private void HandleServiceFailure(string statusFormat, TalkServiceException ex)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }
            if (ex != null && ex.IsAuthenticationError)
            {
                _connectionSetupPending = true;
                _authenticationRejected = true;
                _backendPolicyStatus = null;
                ApplyBackendPolicyStatus("authentication_rejected");
                SetStatus(Strings.ConnectionSignInRequired, true);
                return;
            }
            string message = ex != null && !string.IsNullOrWhiteSpace(ex.Message)
                ? ex.Message.Trim()
                : Strings.StatusTestFailureUnknown;

            string summary = ExtractFirstLine(message);
            SetStatus(string.Format(statusFormat, summary), true);
            if (ex != null && ex.IsTransportError)
            {
                MessageBox.Show(
                    this,
                    message,
                    Strings.ConnectionDiagnosticsDialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private static string ExtractFirstLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return Strings.StatusTestFailureUnknown;
            }
            string[] lines = message.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
            {
                return Strings.StatusTestFailureUnknown;
            }
            return lines[0].Trim();
        }

        private void OnGeneralValueChanged(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }
            if (!_suppressImmediateTlsApply && Result != null
                && (sender == _serverUrlTextBox || sender == _usernameTextBox || sender == _appPasswordTextBox))
            {
                _connectionSetupPending = true;
                _authenticationRejected = false;
                _backendPolicyStatus = null;
                ApplyBackendPolicyStatus("credentials_changed");
                return;
            }
            UpdateControlState();
        }

    }
}
