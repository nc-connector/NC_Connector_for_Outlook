/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * WinForms dialog for all add-in settings (authentication, sharing, IFB, debug, ...).
     * Encapsulates UI logic including starting the login flow, connection tests, and status messages.
     */
    internal sealed class SettingsForm : Form
    {
        private readonly Outlook.Application _outlookApplication;
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private const int HeaderHeight = 48;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

        private readonly TabControl _tabControl = new SettingsTabControl();
        private readonly TabPage _generalTab = new TabPage(Strings.TabGeneral);
        private readonly TabPage _ifbTab = new TabPage(Strings.TabIfb);
        private readonly TabPage _advancedTab = new TabPage(Strings.TabAdvanced);
        private readonly TabPage _debugTab = new TabPage(Strings.TabDebug);
        private readonly TabPage _aboutTab = new TabPage(Strings.TabAbout);
        private readonly TabPage _fileLinkTab = new TabPage(Strings.TabFileLink);
        private readonly TabPage _talkTab = new TabPage(Strings.TabTalkLink);
        private readonly ToolTip _toolTip = new ToolTip();

        private readonly TextBox _serverUrlTextBox = new TextBox();
        private readonly TextBox _usernameTextBox = new TextBox();
        private readonly TextBox _appPasswordTextBox = new TextBox();
        private readonly RadioButton _manualRadio = new RadioButton();
        private readonly RadioButton _loginFlowRadio = new RadioButton();
        private readonly Button _loginFlowButton = new Button();
        private readonly Button _testButton = new Button();
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
        private readonly GroupBox _sharingDefaultsGroup = new GroupBox();
        private readonly Label _sharingDefaultShareNameLabel = new Label();
        private readonly TextBox _sharingDefaultShareNameTextBox = new TextBox();
        private readonly Label _sharingDefaultPermissionsLabel = new Label();
        private readonly CheckBox _sharingDefaultPermCreateCheckBox = new CheckBox();
        private readonly CheckBox _sharingDefaultPermWriteCheckBox = new CheckBox();
        private readonly CheckBox _sharingDefaultPermDeleteCheckBox = new CheckBox();
        private readonly CheckBox _sharingDefaultPasswordCheckBox = new CheckBox();
        private readonly Label _sharingDefaultExpireDaysLabel = new Label();
        private readonly NumericUpDown _sharingDefaultExpireDaysUpDown = new NumericUpDown();
        private readonly GroupBox _talkDefaultsGroup = new GroupBox();
        private readonly Label _talkDefaultRoomTypeLabel = new Label();
        private readonly ComboBox _talkDefaultRoomTypeCombo = new ComboBox();
        private readonly CheckBox _talkDefaultPasswordCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddUsersCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddGuestsCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultLobbyCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultSearchCheckBox = new CheckBox();
        private readonly Label _shareBlockLangLabel = new Label();
        private readonly ComboBox _shareBlockLangCombo = new ComboBox();
        private readonly Label _eventDescriptionLangLabel = new Label();
        private readonly ComboBox _eventDescriptionLangCombo = new ComboBox();

        private readonly Label _statusLabel = new Label();
        private readonly Button _saveButton = new Button();
        private readonly Button _cancelButton = new Button();
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
            ClientSize = new Size(720, 620);
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeComponents();
            ApplySettings(settings);
            UpdateAboutTab();
            UpdateControlState();

            UiThemeManager.ApplyToForm(this, _toolTip);
        }

        private void InitializeComponents()
        {
            const int outerPadding = 12;
            const int buttonRowHeight = 32;
            const int statusHeight = 36;
            const int footerGap = 8;

            int tabTop = HeaderHeight + outerPadding;
            int buttonTop = ClientSize.Height - outerPadding - buttonRowHeight;
            int statusTop = buttonTop - footerGap - statusHeight;

            _tabControl.Location = new Point(outerPadding, tabTop);
            _tabControl.Size = new Size(ClientSize.Width - (outerPadding * 2), statusTop - tabTop - outerPadding);
            _tabControl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _tabControl.TabPages.Add(_generalTab);
            _tabControl.TabPages.Add(_fileLinkTab);
            _tabControl.TabPages.Add(_talkTab);
            _tabControl.TabPages.Add(_ifbTab);
            _tabControl.TabPages.Add(_advancedTab);
            _tabControl.TabPages.Add(_debugTab);
            _tabControl.TabPages.Add(_aboutTab);
            Controls.Add(_tabControl);

            InitializeGeneralTab();
            InitializeTalkTab();
            InitializeIfbTab();
            InitializeAdvancedTab();
            InitializeDebugTab();
            InitializeAboutTab();
            InitializeFileLinkTab();

            _statusLabel.AutoSize = false;
            _statusLabel.Location = new Point(outerPadding, statusTop);
            _statusLabel.Size = new Size(ClientSize.Width - (outerPadding * 2), statusHeight);
            _statusLabel.ForeColor = Color.Black;
            _statusLabel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_statusLabel);

            _saveButton.Text = Strings.ButtonSave;
            _saveButton.Size = new Size(120, buttonRowHeight);
            _saveButton.Location = new Point(ClientSize.Width - outerPadding - 120 - 120 - 10, buttonTop);
            _saveButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _saveButton.DialogResult = DialogResult.OK;
            _saveButton.Click += OnSaveButtonClick;
            Controls.Add(_saveButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.Size = new Size(120, buttonRowHeight);
            _cancelButton.Location = new Point(ClientSize.Width - outerPadding - 120, buttonTop);
            _cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(_cancelButton);

            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
        }

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

            ApplyGeneralTabFieldSizing();
        }

        private void ApplyGeneralTabFieldSizing()
        {
            try
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to apply settings General tab field sizing.", ex);
            }
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = HeaderHeight;
            _headerPanel.Dock = DockStyle.Top;

            Controls.Add(_headerPanel);
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

            int langTop = 70;

            _shareBlockLangLabel.Text = Strings.AdvancedShareBlockLangLabel;
            _shareBlockLangLabel.Location = new Point(24, langTop);
            _shareBlockLangLabel.Size = new Size(220, 20);
            _advancedTab.Controls.Add(_shareBlockLangLabel);

            _shareBlockLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _shareBlockLangCombo.Location = new Point(260, langTop - 2);
            _shareBlockLangCombo.Width = 240;
            PopulateLanguageOverrideCombo(_shareBlockLangCombo);
            _advancedTab.Controls.Add(_shareBlockLangCombo);

            _eventDescriptionLangLabel.Text = Strings.AdvancedEventDescriptionLangLabel;
            _eventDescriptionLangLabel.Location = new Point(24, langTop + 32);
            _eventDescriptionLangLabel.Size = new Size(220, 20);
            _advancedTab.Controls.Add(_eventDescriptionLangLabel);

            _eventDescriptionLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _eventDescriptionLangCombo.Location = new Point(260, langTop + 30);
            _eventDescriptionLangCombo.Width = 240;
            PopulateLanguageOverrideCombo(_eventDescriptionLangCombo);
            _advancedTab.Controls.Add(_eventDescriptionLangCombo);
        }

        private void InitializeTalkTab()
        {
            _talkTab.AutoScroll = true;
            _talkTab.Padding = new Padding(12);

            _talkDefaultsGroup.Text = Strings.SettingsTalkDefaultsGroup;
            _talkDefaultsGroup.Location = new Point(24, 20);
            _talkDefaultsGroup.Size = new Size(480, 180);
            _talkTab.Controls.Add(_talkDefaultsGroup);

            _talkDefaultRoomTypeLabel.Text = Strings.TalkRoomGroup;
            _talkDefaultRoomTypeLabel.Location = new Point(12, 28);
            _talkDefaultRoomTypeLabel.Size = new Size(160, 20);
            _talkDefaultsGroup.Controls.Add(_talkDefaultRoomTypeLabel);

            _talkDefaultRoomTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _talkDefaultRoomTypeCombo.Location = new Point(200, 26);
            _talkDefaultRoomTypeCombo.Width = 240;
            _talkDefaultRoomTypeCombo.Items.Add(new TalkRoomTypeOption(TalkRoomType.EventConversation, Strings.TalkEventRadio));
            _talkDefaultRoomTypeCombo.Items.Add(new TalkRoomTypeOption(TalkRoomType.StandardRoom, Strings.TalkStandardRadio));
            _talkDefaultRoomTypeCombo.SelectedIndexChanged += (s, e) => UpdateTalkRoomTypeTooltip();
            _talkDefaultsGroup.Controls.Add(_talkDefaultRoomTypeCombo);

            _talkDefaultPasswordCheckBox.Text = Strings.TalkPasswordSetCheck;
            _talkDefaultPasswordCheckBox.AutoSize = true;
            _talkDefaultPasswordCheckBox.Location = new Point(12, 58);
            _talkDefaultsGroup.Controls.Add(_talkDefaultPasswordCheckBox);

            _talkDefaultAddUsersCheckBox.Text = Strings.TalkAddUsersCheck;
            _talkDefaultAddUsersCheckBox.AutoSize = true;
            _talkDefaultAddUsersCheckBox.Location = new Point(12, 84);
            _talkDefaultsGroup.Controls.Add(_talkDefaultAddUsersCheckBox);

            _talkDefaultAddGuestsCheckBox.Text = Strings.TalkAddGuestsCheck;
            _talkDefaultAddGuestsCheckBox.AutoSize = true;
            _talkDefaultAddGuestsCheckBox.Location = new Point(12, 108);
            _talkDefaultsGroup.Controls.Add(_talkDefaultAddGuestsCheckBox);

            _talkDefaultLobbyCheckBox.Text = Strings.TalkLobbyCheck;
            _talkDefaultLobbyCheckBox.AutoSize = true;
            _talkDefaultLobbyCheckBox.Location = new Point(12, 132);
            _talkDefaultsGroup.Controls.Add(_talkDefaultLobbyCheckBox);

            _talkDefaultSearchCheckBox.Text = Strings.TalkSearchCheck;
            _talkDefaultSearchCheckBox.AutoSize = true;
            _talkDefaultSearchCheckBox.Location = new Point(12, 156);
            _talkDefaultsGroup.Controls.Add(_talkDefaultSearchCheckBox);

            _toolTip.SetToolTip(_talkDefaultAddUsersCheckBox, Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_talkDefaultAddGuestsCheckBox, Strings.TooltipAddGuests);
            _toolTip.SetToolTip(_talkDefaultLobbyCheckBox, Strings.TooltipLobby);
            _toolTip.SetToolTip(_talkDefaultSearchCheckBox, Strings.TooltipSearchVisible);
            UpdateTalkRoomTypeTooltip();
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

            var supportNoteLabel = new Label
            {
                Text = Strings.AboutSupportNote,
                Location = new Point(18, 110),
                MaximumSize = new Size(640, 0),
                AutoSize = true,
                ForeColor = Color.DimGray
            };
            supportNoteLabel.Size = supportNoteLabel.PreferredSize;
            _aboutTab.Controls.Add(supportNoteLabel);

            var supportHeadingLabel = new Label
            {
                Text = Strings.AboutSupportHeading,
                Location = new Point(18, supportNoteLabel.Bottom + 12),
                AutoSize = true
            };
            supportHeadingLabel.Font = new Font(supportHeadingLabel.Font, FontStyle.Bold);
            _aboutTab.Controls.Add(supportHeadingLabel);

            var paypalLink = new LinkLabel
            {
                Text = Strings.AboutSupportLink,
                Location = new Point(18, supportHeadingLabel.Bottom + 8),
                AutoSize = true
            };
            paypalLink.LinkClicked += (sender, args) => OpenBrowser("https://www.paypal.com/donate/?hosted_button_id=FTZWNRNKVKUN6");
            _aboutTab.Controls.Add(paypalLink);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read assembly version.", ex);
            }

            try
            {
                var info = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(info.ProductVersion))
                {
                    return info.ProductVersion;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read file version info.", ex);
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
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to open license file.", ex);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve license path.", ex);
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
            _ifbEnabledCheckBox.Checked = Result.IfbEnabled;
            SelectComboValue(_ifbDaysCombo, Result.IfbDays, 30);
            SelectComboValue(_ifbCacheHoursCombo, Result.IfbCacheHours, 24);
            _debugLogCheckBox.Checked = Result.DebugLoggingEnabled;
            _fileLinkBaseTextBox.Text = Result.FileLinkBasePath ?? string.Empty;
            _sharingDefaultShareNameTextBox.Text = Result.SharingDefaultShareName ?? string.Empty;
            _sharingDefaultPermCreateCheckBox.Checked = Result.SharingDefaultPermCreate;
            _sharingDefaultPermWriteCheckBox.Checked = Result.SharingDefaultPermWrite;
            _sharingDefaultPermDeleteCheckBox.Checked = Result.SharingDefaultPermDelete;
            _sharingDefaultPasswordCheckBox.Checked = Result.SharingDefaultPasswordEnabled;
            int expireDays = Result.SharingDefaultExpireDays;
            if (expireDays <= 0)
            {
                expireDays = 7;
            }
            if (expireDays > 3650)
            {
                expireDays = 3650;
            }
            _sharingDefaultExpireDaysUpDown.Value = expireDays;
            _talkDefaultPasswordCheckBox.Checked = Result.TalkDefaultPasswordEnabled;
            _talkDefaultAddUsersCheckBox.Checked = Result.TalkDefaultAddUsers;
            _talkDefaultAddGuestsCheckBox.Checked = Result.TalkDefaultAddGuests;
            _talkDefaultLobbyCheckBox.Checked = Result.TalkDefaultLobbyEnabled;
            _talkDefaultSearchCheckBox.Checked = Result.TalkDefaultSearchVisible;
            SelectTalkRoomType(Result.TalkDefaultRoomType);
            UpdateTalkRoomTypeTooltip();
            SelectLanguageChoice(_shareBlockLangCombo, Result.ShareBlockLang);
            SelectLanguageChoice(_eventDescriptionLangCombo, Result.EventDescriptionLang);
            UpdateDebugPathLabel();
            UpdateAboutTab();
        }

        private void OnSaveButtonClick(object sender, EventArgs e)
        {
            Result.ServerUrl = _serverUrlTextBox.Text.Trim();
            Result.Username = _usernameTextBox.Text.Trim();
            Result.AppPassword = _appPasswordTextBox.Text;
            Result.AuthMode = _loginFlowRadio.Checked ? AuthenticationMode.LoginFlow : AuthenticationMode.Manual;
            Result.IfbEnabled = _ifbEnabledCheckBox.Checked;
            Result.IfbDays = ParseComboValue(_ifbDaysCombo, 30);
            Result.IfbCacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            Result.DebugLoggingEnabled = _debugLogCheckBox.Checked;
            Result.LastKnownServerVersion = _lastKnownServerVersion ?? string.Empty;
            Result.FileLinkBasePath = _fileLinkBaseTextBox.Text.Trim();
            Result.SharingDefaultShareName = _sharingDefaultShareNameTextBox.Text.Trim();
            Result.SharingDefaultPermCreate = _sharingDefaultPermCreateCheckBox.Checked;
            Result.SharingDefaultPermWrite = _sharingDefaultPermWriteCheckBox.Checked;
            Result.SharingDefaultPermDelete = _sharingDefaultPermDeleteCheckBox.Checked;
            Result.SharingDefaultPasswordEnabled = _sharingDefaultPasswordCheckBox.Checked;
            Result.SharingDefaultExpireDays = (int)_sharingDefaultExpireDaysUpDown.Value;
            Result.TalkDefaultPasswordEnabled = _talkDefaultPasswordCheckBox.Checked;
            Result.TalkDefaultAddUsers = _talkDefaultAddUsersCheckBox.Checked;
            Result.TalkDefaultAddGuests = _talkDefaultAddGuestsCheckBox.Checked;
            Result.TalkDefaultLobbyEnabled = _talkDefaultLobbyCheckBox.Checked;
            Result.TalkDefaultSearchVisible = _talkDefaultSearchCheckBox.Checked;
            Result.TalkDefaultRoomType = GetSelectedTalkRoomType();
            Result.ShareBlockLang = GetSelectedLanguageChoice(_shareBlockLangCombo);
            Result.EventDescriptionLang = GetSelectedLanguageChoice(_eventDescriptionLangCombo);
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
                catch (Exception ex)
                {
                    // Ignore errors when fetching the optional version hint.
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to fetch optional server version hint.", ex);
                }
                SetStatus(Strings.StatusLoginFlowSuccess, false);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Login flow failed.", ex);
                SetStatus(string.Format(Strings.StatusLoginFlowFailure, ex.Message), true);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Login flow failed unexpectedly.", ex);
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
            DiagnosticsLogger.Log(LogCategories.Core, "Connection test started (Server=" + baseUrl + ", User=" + user + ").");

            try
            {
                var service = new TalkService(new TalkServiceConfiguration(baseUrl, user, appPassword));
                string responseMessage = string.Empty;
                bool success = await Task.Run(() => service.VerifyConnection(out responseMessage));
                if (success)
                {
                    UpdateKnownServerVersion(responseMessage);
                    DiagnosticsLogger.Log(LogCategories.Core, "Connection test succeeded (Response=" + (string.IsNullOrEmpty(responseMessage) ? "OK" : responseMessage) + ").");
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
                    DiagnosticsLogger.Log(LogCategories.Core, "Connection test failed: " + failureMessage);
                    SetStatus(string.Format(Strings.StatusTestFailure, failureMessage), true);
                }
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Connection test failed with service error.", ex);
                SetStatus(string.Format(Strings.StatusTestFailure, ex.Message), true);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Connection test failed unexpectedly.", ex);
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

        private sealed class LanguageOption
        {
            internal LanguageOption(string value, string label)
            {
                Value = value ?? string.Empty;
                Label = label ?? value ?? string.Empty;
            }

            internal string Value { get; private set; }

            internal string Label { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }

        private static string NormalizeLanguageChoice(string value)
        {
            return Strings.NormalizeLanguageOverride(value);
        }

        private static string GetLanguageLabel(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return string.Empty;
            }

            try
            {
                string cultureCode = code.Replace('_', '-');
                return CultureInfo.GetCultureInfo(cultureCode).NativeName;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve language label for '" + code + "'.", ex);
                return code;
            }
        }

        private void PopulateLanguageOverrideCombo(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            combo.Items.Clear();
            foreach (string code in Strings.SupportedLanguageOverrideCodes)
            {
                if (string.Equals(code, "default", StringComparison.OrdinalIgnoreCase))
                {
                    combo.Items.Add(new LanguageOption("default", Strings.LanguageOverrideDefaultOption));
                    continue;
                }

                combo.Items.Add(new LanguageOption(code, GetLanguageLabel(code)));
            }
            if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
        }

        private static void SelectLanguageChoice(ComboBox combo, string value)
        {
            if (combo == null)
            {
                return;
            }

            string normalized = NormalizeLanguageChoice(value);
            foreach (var item in combo.Items)
            {
                var option = item as LanguageOption;
                if (option != null && string.Equals(option.Value, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = option;
                    return;
                }
            }

            if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
        }

        private static string GetSelectedLanguageChoice(ComboBox combo)
        {
            if (combo == null)
            {
                return "default";
            }

            var selected = combo.SelectedItem as LanguageOption;
            return selected != null ? NormalizeLanguageChoice(selected.Value) : "default";
        }

        private sealed class TalkRoomTypeOption
        {
            internal TalkRoomTypeOption(TalkRoomType value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }

            internal TalkRoomType Value { get; private set; }

            internal string Label { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }

        private void SelectTalkRoomType(TalkRoomType value)
        {
            foreach (var item in _talkDefaultRoomTypeCombo.Items)
            {
                var option = item as TalkRoomTypeOption;
                if (option != null && option.Value == value)
                {
                    _talkDefaultRoomTypeCombo.SelectedItem = option;
                    return;
                }
            }

            if (_talkDefaultRoomTypeCombo.Items.Count > 0)
            {
                _talkDefaultRoomTypeCombo.SelectedIndex = 0;
            }
        }

        private TalkRoomType GetSelectedTalkRoomType()
        {
            var selected = _talkDefaultRoomTypeCombo.SelectedItem as TalkRoomTypeOption;
            return selected != null ? selected.Value : TalkRoomType.StandardRoom;
        }

        private void UpdateTalkRoomTypeTooltip()
        {
            var selected = _talkDefaultRoomTypeCombo.SelectedItem as TalkRoomTypeOption;
            TalkRoomType roomType = selected != null ? selected.Value : TalkRoomType.EventConversation;
            _toolTip.SetToolTip(
                _talkDefaultRoomTypeCombo,
                roomType == TalkRoomType.EventConversation ? Strings.TooltipRoomTypeEvent : Strings.TooltipRoomTypeStandard);
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to open debug log path.", ex);
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
            _statusLabel.ForeColor = isError ? _themePalette.ErrorText : _themePalette.SuccessText;
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
            catch (Exception ex)
            {
                // The user will get an error later via timeout/cancellation.
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to open browser for URL '" + url + "'.", ex);
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

            _sharingDefaultsGroup.Text = Strings.SharingDefaultsHeading;
            _sharingDefaultsGroup.Location = new Point(18, 134);
            _sharingDefaultsGroup.Size = new Size(480, 210);
            _fileLinkTab.Controls.Add(_sharingDefaultsGroup);

            _sharingDefaultShareNameLabel.Text = Strings.SharingDefaultShareNameLabel;
            _sharingDefaultShareNameLabel.Location = new Point(12, 28);
            _sharingDefaultShareNameLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultShareNameLabel);

            _sharingDefaultShareNameTextBox.Location = new Point(12, 52);
            _sharingDefaultShareNameTextBox.Width = 320;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultShareNameTextBox);

            _sharingDefaultPermissionsLabel.Text = Strings.SharingDefaultPermissionsLabel;
            _sharingDefaultPermissionsLabel.Location = new Point(12, 90);
            _sharingDefaultPermissionsLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPermissionsLabel);

            _sharingDefaultPermCreateCheckBox.Text = Strings.SharingDefaultPermCreateLabel;
            _sharingDefaultPermCreateCheckBox.Location = new Point(18, 114);
            _sharingDefaultPermCreateCheckBox.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPermCreateCheckBox);

            _sharingDefaultPermWriteCheckBox.Text = Strings.SharingDefaultPermWriteLabel;
            _sharingDefaultPermWriteCheckBox.Location = new Point(18, 138);
            _sharingDefaultPermWriteCheckBox.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPermWriteCheckBox);

            _sharingDefaultPermDeleteCheckBox.Text = Strings.SharingDefaultPermDeleteLabel;
            _sharingDefaultPermDeleteCheckBox.Location = new Point(18, 162);
            _sharingDefaultPermDeleteCheckBox.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPermDeleteCheckBox);

            _sharingDefaultPasswordCheckBox.Text = Strings.SharingDefaultPasswordLabel;
            _sharingDefaultPasswordCheckBox.Location = new Point(260, 114);
            _sharingDefaultPasswordCheckBox.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPasswordCheckBox);

            _sharingDefaultExpireDaysLabel.Text = Strings.SharingDefaultExpireDaysLabel;
            _sharingDefaultExpireDaysLabel.Location = new Point(260, 144);
            _sharingDefaultExpireDaysLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysLabel);

            _sharingDefaultExpireDaysUpDown.Minimum = 1;
            _sharingDefaultExpireDaysUpDown.Maximum = 3650;
            _sharingDefaultExpireDaysUpDown.Location = new Point(260, 168);
            _sharingDefaultExpireDaysUpDown.Width = 90;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysUpDown);
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
