/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * WinForms-Dialog fuer alle Add-in Einstellungen (Authentifizierung, Filelink, IFB, Debug etc.).
     * Kapselt UI-Logik inklusive Login-Flow-Start, Verbindungstest und Statusmeldungen.
     */
    internal sealed class SettingsForm : Form
    {
        private readonly Outlook.Application _outlookApplication;
        private readonly Panel _headerPanel = new Panel();
        private readonly PictureBox _headerLogo = new PictureBox();
        private const int HeaderHeight = 34;

        private readonly TabControl _tabControl = new TabControl();
        private readonly TabPage _generalTab = new TabPage(Strings.TabGeneral);
        private readonly TabPage _ifbTab = new TabPage(Strings.TabIfb);
        private readonly TabPage _invitationsTab = new TabPage(Strings.TabInvitations);
        private readonly TabPage _advancedTab = new TabPage(Strings.TabAdvanced);
        private readonly TabPage _debugTab = new TabPage(Strings.TabDebug);
        private readonly TabPage _aboutTab = new TabPage(Strings.TabAbout);
        private readonly TabPage _fileLinkTab = new TabPage(Strings.TabFileLink);

        private readonly TextBox _serverUrlTextBox = new TextBox();
        private readonly TextBox _usernameTextBox = new TextBox();
        private readonly TextBox _appPasswordTextBox = new TextBox();
        private readonly RadioButton _manualRadio = new RadioButton();
        private readonly RadioButton _loginFlowRadio = new RadioButton();
        private readonly Button _loginFlowButton = new Button();
        private readonly Button _testButton = new Button();

        private readonly CheckBox _muzzleCheckBox = new CheckBox();
        private readonly CheckedListBox _muzzleAccountsList = new CheckedListBox();
        private readonly Label _muzzleAccountsLabel = new Label();
        private readonly Label _muzzleAccountsInfoLabel = new Label();
        private readonly CheckBox _ifbEnabledCheckBox = new CheckBox();
        private readonly ComboBox _ifbDaysCombo = new ComboBox();
        private readonly Label _ifbDaysLabel = new Label();
        private readonly ComboBox _ifbCacheHoursCombo = new ComboBox();
        private readonly Label _ifbCacheHoursLabel = new Label();
        private readonly CheckBox _debugLogCheckBox = new CheckBox();
        private readonly Label _debugPathLabel = new Label();
        private readonly LinkLabel _debugOpenLink = new LinkLabel();
        private readonly Label _aboutVersionLabel = new Label();
        private readonly Label _aboutCopyrightLabel = new Label();
        private readonly LinkLabel _aboutLicenseLink = new LinkLabel();
        private readonly TextBox _fileLinkBaseTextBox = new TextBox();
        private readonly Label _fileLinkBaseHintLabel = new Label();

        private readonly Label _statusLabel = new Label();
        private readonly Button _saveButton = new Button();
        private readonly Button _cancelButton = new Button();

        private readonly List<MuzzleAccountOption> _muzzleAccountOptions = new List<MuzzleAccountOption>();
        private bool _isBusy;
        private AddinSettings _result;
        private string _lastKnownServerVersion = string.Empty;
        private bool _initialIfbEnabled;
        private bool _ifbDefaultApplied;

        internal AddinSettings Result
        {
            get { return _result; }
            private set { _result = value; }
        }

        internal SettingsForm(AddinSettings settings, Outlook.Application outlookApplication)
        {
            _outlookApplication = outlookApplication;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            Text = Strings.SettingsFormTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(560, 400);

            InitializeHeader();
            InitializeComponents();
            ApplySettings(settings);
            UpdateAboutTab();
            UpdateControlState();
        }

        private void InitializeComponents()
        {
            int tabTop = HeaderHeight + 12;
            _tabControl.Location = new Point(12, tabTop);
            _tabControl.Size = new Size(536, 296 - HeaderHeight);
            _tabControl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _tabControl.TabPages.Add(_generalTab);
            _tabControl.TabPages.Add(_fileLinkTab);
            _tabControl.TabPages.Add(_ifbTab);
            _tabControl.TabPages.Add(_invitationsTab);
            _tabControl.TabPages.Add(_advancedTab);
            _tabControl.TabPages.Add(_debugTab);
            _tabControl.TabPages.Add(_aboutTab);
            Controls.Add(_tabControl);

            InitializeGeneralTab();
            InitializeIfbTab();
            InitializeInvitationsTab();
            InitializeAdvancedTab();
            InitializeDebugTab();
            InitializeAboutTab();
            InitializeFileLinkTab();

            _statusLabel.AutoSize = false;
            _statusLabel.Location = new Point(18, 316);
            _statusLabel.Size = new Size(524, 32);
            _statusLabel.ForeColor = Color.Black;
            Controls.Add(_statusLabel);

            _saveButton.Text = Strings.ButtonSave;
            _saveButton.Location = new Point(320, 360);
            _saveButton.Width = 100;
            _saveButton.DialogResult = DialogResult.OK;
            _saveButton.Click += OnSaveButtonClick;
            Controls.Add(_saveButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.Location = new Point(430, 360);
            _cancelButton.Width = 100;
            _cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(_cancelButton);

            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
        }

        private void InitializeGeneralTab()
        {
            _generalTab.AutoScroll = true;

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
            _manualRadio.CheckedChanged += OnAuthModeChanged;
            authGroup.Controls.Add(_manualRadio);

            _loginFlowRadio.Text = Strings.RadioLoginFlow;
            _loginFlowRadio.Location = new Point(12, 55);
            _loginFlowRadio.AutoSize = true;
            _loginFlowRadio.CheckedChanged += OnAuthModeChanged;
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
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = HeaderHeight;
            _headerPanel.Dock = DockStyle.Top;
            _headerPanel.BackColor = Color.FromArgb(0, 120, 212);

            _headerLogo.Size = new Size(22, 22);
            _headerLogo.SizeMode = PictureBoxSizeMode.Zoom;
            _headerLogo.Image = LoadEmbeddedLogo("NcTalkOutlookAddIn.Resources.logo-nextcloud.png");
            _headerPanel.Controls.Add(_headerLogo);
            _headerPanel.Resize += (s, e) => CenterHeaderLogo();

            Controls.Add(_headerPanel);
            CenterHeaderLogo();
        }

        private void InitializeIfbTab()
        {
            _ifbTab.AutoScroll = true;
            _ifbTab.Padding = new Padding(12);

            _ifbEnabledCheckBox.Text = Strings.CheckIfbEnabled;
            _ifbEnabledCheckBox.AutoSize = true;
            _ifbEnabledCheckBox.Location = new Point(24, 20);
            _ifbEnabledCheckBox.CheckedChanged += OnIfbEnabledChanged;
            _ifbTab.Controls.Add(_ifbEnabledCheckBox);

            _ifbDaysLabel.Text = Strings.LabelIfbDays;
            _ifbDaysLabel.Location = new Point(24, 60);
            _ifbDaysLabel.Size = new Size(160, 20);
            _ifbTab.Controls.Add(_ifbDaysLabel);

            _ifbDaysCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbDaysCombo.Location = new Point(200, 58);
            _ifbDaysCombo.Width = 100;
            _ifbDaysCombo.Items.AddRange(new object[] { "10", "30", "60", "90" });
            _ifbTab.Controls.Add(_ifbDaysCombo);
        }

        private void InitializeInvitationsTab()
        {
            _invitationsTab.AutoScroll = true;
            _invitationsTab.Padding = new Padding(12);

            var muzzleGroup = new GroupBox
            {
                Text = Strings.GroupMuzzle,
                Location = new Point(18, 20),
                Size = new Size(440, 230),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _invitationsTab.Controls.Add(muzzleGroup);

            _muzzleCheckBox.Text = Strings.CheckMuzzle;
            _muzzleCheckBox.AutoSize = true;
            _muzzleCheckBox.Location = new Point(12, 25);
            _muzzleCheckBox.Checked = true;
            _muzzleCheckBox.CheckedChanged += OnMuzzleCheckChanged;
            muzzleGroup.Controls.Add(_muzzleCheckBox);

            _muzzleAccountsLabel.Text = Strings.LabelMuzzleAccounts;
            _muzzleAccountsLabel.Location = new Point(12, 55);
            _muzzleAccountsLabel.AutoSize = true;
            muzzleGroup.Controls.Add(_muzzleAccountsLabel);

            _muzzleAccountsList.Location = new Point(12, 75);
            _muzzleAccountsList.Size = new Size(400, 90);
            _muzzleAccountsList.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _muzzleAccountsList.CheckOnClick = true;
            _muzzleAccountsList.HorizontalScrollbar = true;
            muzzleGroup.Controls.Add(_muzzleAccountsList);

            _muzzleAccountsInfoLabel.Text = Strings.LabelMuzzleNoAccounts;
            _muzzleAccountsInfoLabel.Location = new Point(12, 75);
            _muzzleAccountsInfoLabel.Size = new Size(400, 40);
            _muzzleAccountsInfoLabel.Visible = false;
            muzzleGroup.Controls.Add(_muzzleAccountsInfoLabel);

            var muzzleHintLabel = new Label
            {
                Text = Strings.LabelMuzzleAccountsHint + Environment.NewLine + Strings.LabelMuzzleHint,
                Location = new Point(12, 170),
                Size = new Size(400, 50)
            };
            muzzleGroup.Controls.Add(muzzleHintLabel);
        }

        private void InitializeAdvancedTab()
        {
            _advancedTab.AutoScroll = true;
            _advancedTab.Padding = new Padding(12);

            _ifbCacheHoursLabel.Text = Strings.LabelIfbCacheHours;
            _ifbCacheHoursLabel.Location = new Point(24, 24);
            _ifbCacheHoursLabel.Size = new Size(220, 20);
            _advancedTab.Controls.Add(_ifbCacheHoursLabel);

            _ifbCacheHoursCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbCacheHoursCombo.Location = new Point(260, 22);
            _ifbCacheHoursCombo.Width = 80;
            _ifbCacheHoursCombo.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            for (int i = 1; i <= 24; i++)
            {
                _ifbCacheHoursCombo.Items.Add(i.ToString());
            }
            _advancedTab.Controls.Add(_ifbCacheHoursCombo);
        }

        private void InitializeDebugTab()
        {
            _debugTab.AutoScroll = true;
            _debugTab.Padding = new Padding(12);

            _debugLogCheckBox.Text = Strings.DebugCheckbox;
            _debugLogCheckBox.AutoSize = true;
            _debugLogCheckBox.Location = new Point(24, 20);
            _debugTab.Controls.Add(_debugLogCheckBox);

            _debugPathLabel.AutoSize = false;
            _debugPathLabel.Location = new Point(24, 60);
            _debugPathLabel.Size = new Size(420, 40);
            _debugTab.Controls.Add(_debugPathLabel);

            _debugOpenLink.Text = Strings.DebugOpenLog;
            _debugOpenLink.Location = new Point(24, 110);
            _debugOpenLink.AutoSize = true;
            _debugOpenLink.LinkClicked += OnDebugOpenLinkClicked;
            _debugTab.Controls.Add(_debugOpenLink);

            UpdateDebugPathLabel();
        }

        private void InitializeAboutTab()
        {
            _aboutTab.Padding = new Padding(12);

            _aboutVersionLabel.AutoSize = true;
            _aboutVersionLabel.Location = new Point(18, 20);
            _aboutTab.Controls.Add(_aboutVersionLabel);

            _aboutCopyrightLabel.AutoSize = true;
            _aboutCopyrightLabel.Location = new Point(18, 50);
            _aboutTab.Controls.Add(_aboutCopyrightLabel);

            var licenseLabel = new Label
            {
                Text = Strings.AboutLicenseLabel,
                Location = new Point(18, 80),
                AutoSize = true
            };
            _aboutTab.Controls.Add(licenseLabel);

            _aboutLicenseLink.AutoSize = true;
            _aboutLicenseLink.Location = new Point(80, 80);
            _aboutLicenseLink.Text = Strings.AboutLicenseLink;
            _aboutLicenseLink.LinkClicked += OnAboutLicenseLinkClicked;
            _aboutTab.Controls.Add(_aboutLicenseLink);
        }

        private void UpdateDebugPathLabel()
        {
            string path = DiagnosticsLogger.LogFileFullPath ?? string.Empty;
            _debugPathLabel.Text = Strings.DebugPathPrefix + path;
        }

        private void UpdateAboutTab()
        {
            string version = GetAddinVersion() ?? "n/a";
            _aboutVersionLabel.Text = string.Format(Strings.AboutVersionFormat, version);
            _aboutCopyrightLabel.Text = Strings.AboutCopyright;
        }

        private static string GetAddinVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                if (version != null)
                {
                    return version.ToString();
                }
            }
            catch
            {
            }

            try
            {
                var info = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(info.ProductVersion))
                {
                    return info.ProductVersion;
                }
            }
            catch
            {
            }

            return null;
        }

        private void OnAboutLicenseLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string licensePath = ResolveLicensePath();
            try
            {
                if (File.Exists(licensePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = licensePath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    MessageBox.Show(
                        string.Format(Strings.LicenseFileMissingMessage, licensePath),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format(Strings.LicenseFileOpenErrorMessage, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private static string ResolveLicensePath()
        {
            try
            {
                string candidate;

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(assemblyPath))
                {
                    string baseDir = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
                    candidate = Path.Combine(baseDir, "License.txt");
                    if (!string.IsNullOrEmpty(candidate))
                    {
                        return candidate;
                    }
                }

                string appBase = AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
                candidate = Path.Combine(appBase, "License.txt");
                if (!string.IsNullOrEmpty(candidate))
                {
                    return candidate;
                }

                return "License.txt";
            }
            catch
            {
                return "License.txt";
            }
        }

        private void ApplySettings(AddinSettings settings)
        {
            Result = settings.Clone();
            _initialIfbEnabled = Result.IfbEnabled;
            _ifbDefaultApplied = _initialIfbEnabled || !string.IsNullOrEmpty(Result.IfbPreviousFreeBusyPath);
            _lastKnownServerVersion = Result.LastKnownServerVersion ?? string.Empty;
            _serverUrlTextBox.Text = Result.ServerUrl;
            _usernameTextBox.Text = Result.Username;
            _appPasswordTextBox.Text = Result.AppPassword;
            _manualRadio.Checked = Result.AuthMode == AuthenticationMode.Manual;
            _loginFlowRadio.Checked = !_manualRadio.Checked;
            _muzzleCheckBox.Checked = Result.OutlookMuzzleEnabled;
            LoadMuzzleAccounts();
            _ifbEnabledCheckBox.Checked = Result.IfbEnabled;
            SelectComboValue(_ifbDaysCombo, Result.IfbDays, 30);
            SelectComboValue(_ifbCacheHoursCombo, Result.IfbCacheHours, 24);
            _debugLogCheckBox.Checked = Result.DebugLoggingEnabled;
            _fileLinkBaseTextBox.Text = Result.FileLinkBasePath ?? string.Empty;
            UpdateDebugPathLabel();
            UpdateAboutTab();
        }

        private void LoadMuzzleAccounts()
        {
            _muzzleAccountOptions.Clear();
            _muzzleAccountsList.Items.Clear();

            if (_outlookApplication == null)
            {
                ShowMuzzleAccountsInfo(Strings.LabelMuzzleNoAccounts);
                UpdateControlState();
                return;
            }

            try
            {
                var session = _outlookApplication.Session;
                if (session == null)
                {
                    DiagnosticsLogger.Log("Muzzle", "Settings: Outlook session not available.");
                    ShowMuzzleAccountsInfo(Strings.LabelMuzzleNoAccounts);
                    UpdateControlState();
                    return;
                }

                var accounts = session.Accounts;
                if (accounts == null || accounts.Count == 0)
                {
                    DiagnosticsLogger.Log("Muzzle", "Settings: keine Outlook-Accounts gefunden.");
                    ShowMuzzleAccountsInfo(Strings.LabelMuzzleNoAccounts);
                    UpdateControlState();
                    return;
                }

                for (int i = 1; i <= accounts.Count; i++)
                {
                    Outlook.Account account = null;
                    try
                    {
                        account = accounts[i];
                        var option = new MuzzleAccountOption(
                            OutlookAccountHelper.NormalizeAccountIdentifier(account),
                            BuildAccountDisplay(account));
                        _muzzleAccountOptions.Add(option);
                        int index = _muzzleAccountsList.Items.Add(option);
                        bool enabled = ResolveMuzzleAccountState(option.Key);
                        _muzzleAccountsList.SetItemChecked(index, enabled);
                        DiagnosticsLogger.Log("Muzzle", "Settings: Account " + option.Key + " loaded (enabled=" + enabled + ").");
                    }
                    catch
                    {
                    }
                    finally
                    {
                        ReleaseCom(account);
                    }
                }

                bool hasAccounts = _muzzleAccountOptions.Count > 0;
                _muzzleAccountsInfoLabel.Visible = !hasAccounts;
                _muzzleAccountsList.Visible = hasAccounts;
                DiagnosticsLogger.Log("Muzzle", "Settings: Gesamtanzahl Accounts = " + _muzzleAccountOptions.Count + ".");
            }
            catch
            {
                ShowMuzzleAccountsInfo(Strings.LabelMuzzleNoAccounts);
            }

            UpdateControlState();
        }

        private void ShowMuzzleAccountsInfo(string message)
        {
            _muzzleAccountsInfoLabel.Text = message;
            _muzzleAccountsInfoLabel.Visible = true;
            _muzzleAccountsList.Visible = false;
        }

        private bool ResolveMuzzleAccountState(string key)
        {
            if (Result == null)
            {
                return true;
            }

            return Result.IsMuzzleEnabledForAccount(key);
        }

        private Dictionary<string, bool> CollectMuzzleAccountStates()
        {
            var map = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _muzzleAccountOptions.Count && i < _muzzleAccountsList.Items.Count; i++)
            {
                var option = _muzzleAccountOptions[i];
                if (string.IsNullOrWhiteSpace(option.Key))
                {
                    continue;
                }

                map[option.Key] = _muzzleAccountsList.GetItemChecked(i);
            }

            return map;
        }

        private static string BuildAccountDisplay(Outlook.Account account)
        {
            if (account == null)
            {
                return Strings.LabelMuzzleUnknownAccount;
            }

            string smtp = TryGetAccountProperty(() => account.SmtpAddress);
            string display = TryGetAccountProperty(() => account.DisplayName);
            string user = TryGetAccountProperty(() => account.UserName);

            if (!string.IsNullOrWhiteSpace(display) &&
                !string.IsNullOrWhiteSpace(smtp) &&
                !string.Equals(display, smtp, StringComparison.OrdinalIgnoreCase))
            {
                return display + " (" + smtp + ")";
            }

            if (!string.IsNullOrWhiteSpace(smtp))
            {
                return smtp;
            }

            if (!string.IsNullOrWhiteSpace(display))
            {
                return display;
            }

            if (!string.IsNullOrWhiteSpace(user))
            {
                return user;
            }

            return Strings.LabelMuzzleUnknownAccount;
        }

        private static string TryGetAccountProperty(Func<string> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void ReleaseCom(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch
            {
            }
        }

        private sealed class MuzzleAccountOption
        {
            internal MuzzleAccountOption(string key, string display)
            {
                Key = key ?? string.Empty;
                Display = display ?? string.Empty;
            }

            internal string Key { get; private set; }

            private string Display { get; set; }

            public override string ToString()
            {
                return Display;
            }
        }

        private void OnSaveButtonClick(object sender, EventArgs e)
        {
            Result.ServerUrl = _serverUrlTextBox.Text.Trim();
            Result.Username = _usernameTextBox.Text.Trim();
            Result.AppPassword = _appPasswordTextBox.Text;
            Result.AuthMode = _loginFlowRadio.Checked ? AuthenticationMode.LoginFlow : AuthenticationMode.Manual;
            Result.OutlookMuzzleEnabled = _muzzleCheckBox.Checked;
            Result.OutlookMuzzleAccounts = CollectMuzzleAccountStates();
            Result.IfbEnabled = _ifbEnabledCheckBox.Checked;
            Result.IfbDays = ParseComboValue(_ifbDaysCombo, 30);
            Result.IfbCacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            Result.DebugLoggingEnabled = _debugLogCheckBox.Checked;
            Result.LastKnownServerVersion = _lastKnownServerVersion ?? string.Empty;
            Result.FileLinkBasePath = _fileLinkBaseTextBox.Text.Trim();
        }

        private async void OnLoginFlowButtonClick(object sender, EventArgs e)
        {
            if (_isBusy)
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

            try
            {
                var flowService = new TalkLoginFlowService(normalizedUrl);
                var startInfo = flowService.StartLoginFlow();

                OpenBrowser(startInfo.LoginUrl);
                SetStatus(Strings.StatusLoginFlowBrowser, false);

                var credentials = await Task.Run(() => flowService.CompleteLoginFlow(startInfo, TimeSpan.FromMinutes(2), TimeSpan.FromSeconds(2)));
                _usernameTextBox.Text = credentials.LoginName ?? string.Empty;
                _appPasswordTextBox.Text = credentials.AppPassword ?? string.Empty;
                _loginFlowRadio.Checked = true;
                try
                {
                    var verificationService = new TalkService(new TalkServiceConfiguration(
                        normalizedUrl,
                        _usernameTextBox.Text,
                        _appPasswordTextBox.Text));
                    string versionResponse = string.Empty;
                    bool ok = await Task.Run(() => verificationService.VerifyConnection(out versionResponse));
                    if (ok)
                    {
                        UpdateKnownServerVersion(versionResponse);
                    }
                }
                catch
                {
                    // Wir ignorieren Fehler beim optionalen Versionsabruf.
                }
                SetStatus(Strings.StatusLoginFlowSuccess, false);
            }
            catch (TalkServiceException ex)
            {
                SetStatus(string.Format(Strings.StatusLoginFlowFailure, ex.Message), true);
            }
            catch (Exception ex)
            {
                SetStatus(string.Format(Strings.StatusLoginFlowFailure, ex.Message), true);
            }
            finally
            {
                SetBusy(false);
                UpdateControlState();
            }
        }

        private async void OnTestButtonClick(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            string baseUrl = _serverUrlTextBox.Text.Trim();
            string user = _usernameTextBox.Text.Trim();
            string appPassword = _appPasswordTextBox.Text;

            if (string.IsNullOrWhiteSpace(baseUrl) ||
                string.IsNullOrWhiteSpace(user) ||
                string.IsNullOrEmpty(appPassword))
            {
                SetStatus(Strings.StatusMissingFields, true);
                return;
            }

            SetBusy(true);
            SetStatus(Strings.StatusTestRunning, false);
            DiagnosticsLogger.Log("Settings", "Verbindungstest gestartet (Server=" + baseUrl + ", Benutzer=" + user + ").");

            try
            {
                var service = new TalkService(new TalkServiceConfiguration(baseUrl, user, appPassword));
                string responseMessage = string.Empty;
                bool success = await Task.Run(() => service.VerifyConnection(out responseMessage));
                if (success)
                {
                    UpdateKnownServerVersion(responseMessage);
                    DiagnosticsLogger.Log("Settings", "Verbindungstest erfolgreich (Antwort=" + (string.IsNullOrEmpty(responseMessage) ? "OK" : responseMessage) + ").");
                    string suffix = string.IsNullOrEmpty(responseMessage)
                        ? string.Empty
                        : " (" + string.Format(Strings.StatusTestSuccessVersionFormat, responseMessage) + ")";
                    SetStatus(string.Format(Strings.StatusTestSuccessFormat, suffix), false);
                }
                else
                {
                    var failureMessage = string.IsNullOrEmpty(responseMessage)
                        ? Strings.StatusTestFailureUnknown
                        : responseMessage;
                    DiagnosticsLogger.Log("Settings", "Verbindungstest fehlgeschlagen: " + failureMessage);
                    SetStatus(string.Format(Strings.StatusTestFailure, failureMessage), true);
                }
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.Log("Settings", "Verbindungstest TalkServiceException: " + ex.Message);
                SetStatus(string.Format(Strings.StatusTestFailure, ex.Message), true);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log("Settings", "Verbindungstest Exception: " + ex.Message);
                SetStatus(string.Format(Strings.StatusTestFailure, ex.Message), true);
            }
            finally
            {
                SetBusy(false);
                UpdateControlState();
            }
        }

        private void OnAuthModeChanged(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            UpdateControlState();
        }

        private void OnIfbEnabledChanged(object sender, EventArgs e)
        {
            _ifbDefaultApplied = true;

            if (_isBusy)
            {
                return;
            }

            UpdateControlState();
        }

        private void OnMuzzleCheckChanged(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            UpdateControlState();
        }

        private void OnGeneralValueChanged(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            UpdateControlState();
        }

        private static void SelectComboValue(ComboBox combo, int value, int fallback)
        {
            if (combo == null)
            {
                return;
            }

            string text = value.ToString();
            if (combo.Items.Contains(text))
            {
                combo.SelectedItem = text;
            }
            else if (combo.Items.Contains(fallback.ToString()))
            {
                combo.SelectedItem = fallback.ToString();
            }
            else if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
        }

        private static int ParseComboValue(ComboBox combo, int fallback)
        {
            if (combo == null)
            {
                return fallback;
            }

            string selected = combo.SelectedItem as string;
            int parsed;
            if (selected != null && int.TryParse(selected, out parsed))
            {
                return parsed;
            }

            int parsedText;
            if (int.TryParse(combo.Text, out parsedText))
            {
                return parsedText;
            }

            return fallback;
        }

        private void OnDebugOpenLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string path = DiagnosticsLogger.LogFileFullPath ?? string.Empty;
            try
            {
                if (File.Exists(path))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                else
                {
                    string directory = Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = directory,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        MessageBox.Show(
                            Strings.DebugLogMissingMessage,
                            Strings.DialogTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch
            {
                MessageBox.Show(
                    Strings.DebugLogOpenErrorMessage,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void UpdateControlState()
        {
            bool manual = _manualRadio.Checked;

            _usernameTextBox.Enabled = manual && !_isBusy;
            _appPasswordTextBox.Enabled = manual && !_isBusy;
            _loginFlowButton.Enabled = !manual && !_isBusy;
            _testButton.Enabled = !_isBusy;
            _muzzleCheckBox.Enabled = !_isBusy;
            bool muzzleAccountsAvailable = _muzzleAccountOptions.Count > 0 && !_muzzleAccountsInfoLabel.Visible;
            bool muzzleControlsEnabled = _muzzleCheckBox.Checked && !_isBusy && muzzleAccountsAvailable;
            _muzzleAccountsLabel.Enabled = muzzleControlsEnabled || (_muzzleAccountOptions.Count > 0 && _muzzleCheckBox.Checked && !_isBusy);
            _muzzleAccountsList.Enabled = muzzleControlsEnabled;
            _muzzleAccountsInfoLabel.Enabled = _muzzleCheckBox.Checked && !_isBusy;

            bool credentialsAvailable =
                !string.IsNullOrWhiteSpace(_serverUrlTextBox.Text) &&
                !string.IsNullOrWhiteSpace(_usernameTextBox.Text) &&
                !string.IsNullOrEmpty(_appPasswordTextBox.Text);

            if (credentialsAvailable && !_ifbDefaultApplied && !_initialIfbEnabled && !_ifbEnabledCheckBox.Checked)
            {
                _ifbEnabledCheckBox.Checked = true;
                _ifbDefaultApplied = true;
            }

            if (!credentialsAvailable && _ifbEnabledCheckBox.Checked)
            {
                _ifbEnabledCheckBox.Checked = false;
            }

            _ifbEnabledCheckBox.Enabled = credentialsAvailable && !_isBusy;

            bool showDays = _ifbEnabledCheckBox.Checked;
            _ifbDaysCombo.Visible = showDays;
            _ifbDaysLabel.Visible = showDays;
            _ifbDaysCombo.Enabled = showDays && !_isBusy && _ifbEnabledCheckBox.Enabled;
            _ifbDaysLabel.Enabled = showDays && !_isBusy && _ifbEnabledCheckBox.Enabled;

            _ifbCacheHoursCombo.Enabled = !_isBusy;
            _ifbCacheHoursLabel.Enabled = !_isBusy;
            _debugLogCheckBox.Enabled = !_isBusy;
            _debugOpenLink.Enabled = !_isBusy;
        }

        private void SetBusy(bool busy)
        {
            _isBusy = busy;
            Cursor.Current = busy ? Cursors.WaitCursor : Cursors.Default;
            _saveButton.Enabled = !busy;
            _cancelButton.Enabled = !busy;
            _debugOpenLink.Enabled = !busy;
            UpdateControlState();
        }

        private void SetStatus(string message, bool isError)
        {
            _statusLabel.Text = message;
            _statusLabel.ForeColor = isError ? Color.DarkRed : Color.DarkGreen;
        }

        private static void OpenBrowser(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Benutzer erhaelt spaeter Fehlermeldung ueber Timeout/Abbruch.
            }
        }

        private void InitializeFileLinkTab()
        {
            _fileLinkTab.Padding = new Padding(12);
            _fileLinkTab.AutoScroll = true;

            var baseLabel = new Label
            {
                Text = Strings.FileLinkBaseLabel,
                Location = new Point(18, 24),
                AutoSize = true
            };
            _fileLinkTab.Controls.Add(baseLabel);

            _fileLinkBaseTextBox.Location = new Point(18, 48);
            _fileLinkBaseTextBox.Width = 360;
            _fileLinkTab.Controls.Add(_fileLinkBaseTextBox);

            _fileLinkBaseHintLabel.Text = Strings.FileLinkBaseHint;
            _fileLinkBaseHintLabel.Location = new Point(18, 82);
            _fileLinkBaseHintLabel.Size = new Size(420, 40);
            _fileLinkBaseHintLabel.AutoSize = false;
            _fileLinkBaseHintLabel.ForeColor = Color.DimGray;
            _fileLinkTab.Controls.Add(_fileLinkBaseHintLabel);
        }

        private void CenterHeaderLogo()
        {
            if (_headerPanel == null || _headerLogo == null)
            {
                return;
            }

            int x = (_headerPanel.Width - _headerLogo.Width) / 2;
            int y = (_headerPanel.Height - _headerLogo.Height) / 2;
            _headerLogo.Location = new Point(Math.Max(0, x), Math.Max(0, y));
        }

        private static Image LoadEmbeddedLogo(string resourceName)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                return Image.FromStream(stream);
            }
        }

        private void UpdateKnownServerVersion(string candidate)
        {
            Version parsed;
            if (NextcloudVersionHelper.TryParse(candidate, out parsed))
            {
                _lastKnownServerVersion = parsed.ToString();
            }
        }
    }
}
