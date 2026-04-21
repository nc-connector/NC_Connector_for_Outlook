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
using System.Net;
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
    internal sealed partial class SettingsForm : Form
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
        private const string NcConnectorHomepageUrl = "https://nc-connector.de";
        private readonly TabPage _fileLinkTab = new TabPage(Strings.TabFileLink);
        private readonly TabPage _talkTab = new TabPage(Strings.TabTalkLink);
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly DisabledControlTooltipHintHelper _disabledTooltipHints;
        private readonly Panel _policyWarningPanel = new Panel();
        private readonly Label _policyWarningTitleLabel = new Label();
        private readonly Label _policyWarningTextLabel = new Label();
        private readonly LinkLabel _policyWarningLinkLabel = new LinkLabel();

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
        private readonly Label _ifbPortLabel = new Label();
        private readonly NumericUpDown _ifbPortUpDown = new NumericUpDown();
        private readonly ComboBox _ifbCacheHoursCombo = new ComboBox();
        private readonly Label _ifbCacheHoursLabel = new Label();
        private readonly CheckBox _debugLogCheckBox = new CheckBox();
        private readonly CheckBox _debugAnonymizeCheckBox = new CheckBox();
        private readonly Label _debugPathLabel = new Label();
        private readonly LinkLabel _debugOpenLink = new LinkLabel();
        private readonly Label _aboutVersionLabel = new Label();
        private readonly Label _aboutCopyrightLabel = new Label();
        private readonly Label _aboutLicenseLabel = new Label();
        private readonly LinkLabel _aboutLicenseLink = new LinkLabel();
        private readonly Label _aboutHomepageLabel = new Label();
        private readonly LinkLabel _aboutHomepageLink = new LinkLabel();
        private readonly Label _aboutOverviewLabel = new Label();
        private readonly Label _aboutMoreInfoLabel = new Label();
        private readonly LinkLabel _aboutMoreInfoLink = new LinkLabel();
        private readonly Label _aboutSupportNoteLabel = new Label();
        private readonly Label _aboutSupportHeadingLabel = new Label();
        private readonly LinkLabel _aboutSupportLink = new LinkLabel();
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
        private readonly CheckBox _sharingDefaultPasswordSeparateCheckBox = new CheckBox();
        private readonly Label _sharingDefaultExpireDaysLabel = new Label();
        private readonly NumericUpDown _sharingDefaultExpireDaysUpDown = new NumericUpDown();
        private readonly GroupBox _sharingAttachmentAutomationGroup = new GroupBox();
        private readonly Label _sharingAttachmentLockHintLabel = new Label();
        private readonly Label _sharingAttachmentLockStepsLabel = new Label();
        private readonly CheckBox _sharingAttachmentsAlwaysCheckBox = new CheckBox();
        private readonly CheckBox _sharingAttachmentsOfferAboveCheckBox = new CheckBox();
        private readonly NumericUpDown _sharingAttachmentsOfferAboveMbUpDown = new NumericUpDown();
        private readonly Label _sharingAttachmentsOfferAboveUnitLabel = new Label();
        private readonly GroupBox _talkDefaultsGroup = new GroupBox();
        private readonly Label _talkDefaultRoomTypeLabel = new Label();
        private readonly ComboBox _talkDefaultRoomTypeCombo = new ComboBox();
        private readonly CheckBox _talkDefaultPasswordCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddUsersCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddGuestsCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultLobbyCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultSearchCheckBox = new CheckBox();
        private readonly Panel _talkAddressbookWarningPanel = new Panel();
        private readonly Label _talkAddressbookWarningTitleLabel = new Label();
        private readonly Label _talkAddressbookWarningTextLabel = new Label();
        private readonly LinkLabel _talkAddressbookWarningLinkLabel = new LinkLabel();
        private readonly Label _shareBlockLangLabel = new Label();
        private readonly ComboBox _shareBlockLangCombo = new ComboBox();
        private readonly Label _eventDescriptionLangLabel = new Label();
        private readonly ComboBox _eventDescriptionLangCombo = new ComboBox();
        private readonly GroupBox _tlsSettingsGroup = new GroupBox();
        private readonly CheckBox _tlsUseSystemDefaultCheckBox = new CheckBox();
        private readonly CheckBox _tlsEnable12CheckBox = new CheckBox();
        private readonly CheckBox _tlsEnable13CheckBox = new CheckBox();
        private readonly Label _tlsHintLabel = new Label();

        private readonly Label _statusLabel = new Label();
        private readonly Button _saveButton = new Button();
        private readonly Button _cancelButton = new Button();
        private bool _isBusy;
        private AddinSettings _result;
        private string _lastKnownServerVersion = string.Empty;
        private bool _initialIfbEnabled;
        private bool _ifbDefaultApplied;
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private bool _sharingAttachmentLockActive;
        private int _sharingAttachmentLockThresholdMb = 5;
        private bool _talkAddressbookLockActive;
        private string _talkAddressbookLockDetail = string.Empty;
        private bool _layoutApplying;
        private bool _suppressImmediateTlsApply;
        private readonly SecurityProtocolType _runtimeSecurityProtocolAtOpen;
        private BackendPolicyStatus _backendPolicyStatus;

        internal AddinSettings Result
        {
            get { return _result; }
            private set { _result = value; }
        }

        internal SettingsForm(AddinSettings settings, Outlook.Application outlookApplication, BackendPolicyStatus initialPolicyStatus)
        {
            _outlookApplication = outlookApplication;
            _backendPolicyStatus = initialPolicyStatus;
            _runtimeSecurityProtocolAtOpen = ServicePointManager.SecurityProtocol;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            Text = Strings.SettingsFormTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(720, 620);
            MinimumSize = new Size(ScaleLogical(780), ScaleLogical(680));
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeComponents();
            ApplySettings(settings);
            UpdateAboutTab();
            UpdateControlState();
            ApplyResponsiveLayout(true);

            UiThemeManager.ApplyToForm(this, _toolTip);
            FormClosed += OnSettingsFormClosed;
        }

        private void InitializeComponents()
        {
            _tabControl.Location = new Point(12, HeaderHeight + 12);
            _tabControl.Size = new Size(ClientSize.Width - 24, ClientSize.Height - HeaderHeight - 110);
            _tabControl.Anchor = AnchorStyles.None;
            _tabControl.TabPages.Add(_generalTab);
            _tabControl.TabPages.Add(_fileLinkTab);
            _tabControl.TabPages.Add(_talkTab);
            _tabControl.TabPages.Add(_ifbTab);
            _tabControl.TabPages.Add(_advancedTab);
            _tabControl.TabPages.Add(_debugTab);
            _tabControl.TabPages.Add(_aboutTab);
            _tabControl.SelectedIndexChanged += OnSelectedTabChanged;
            Controls.Add(_tabControl);
            InitializePolicyWarningPanel();

            InitializeGeneralTab();
            InitializeTalkTab();
            InitializeIfbTab();
            InitializeAdvancedTab();
            InitializeDebugTab();
            InitializeAboutTab();
            InitializeFileLinkTab();

            _statusLabel.AutoSize = false;
            _statusLabel.Location = new Point(12, ClientSize.Height - 80);
            _statusLabel.Size = new Size(ClientSize.Width - 24, 36);
            _statusLabel.ForeColor = Color.Black;
            _statusLabel.Anchor = AnchorStyles.None;
            Controls.Add(_statusLabel);

            _saveButton.Text = Strings.ButtonSave;
            _saveButton.Size = new Size(120, 32);
            _saveButton.Location = new Point(ClientSize.Width - 262, ClientSize.Height - 44);
            _saveButton.Anchor = AnchorStyles.None;
            _saveButton.DialogResult = DialogResult.OK;
            _saveButton.Click += OnSaveButtonClick;
            Controls.Add(_saveButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.Size = new Size(120, 32);
            _cancelButton.Location = new Point(ClientSize.Width - 132, ClientSize.Height - 44);
            _cancelButton.Anchor = AnchorStyles.None;
            _cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(_cancelButton);

            Resize += (s, e) => ApplyResponsiveLayout(false);
            AttachResponsiveResizeHandlers(_generalTab, _fileLinkTab, _talkTab);

            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
        }

        private void AttachResponsiveResizeHandlers(params Control[] controls)
        {            if (controls == null)
            {
                return;
            }
            for (int i = 0; i < controls.Length; i++)
            {
                Control control = controls[i];                if (control == null)
                {
                    continue;
                }

                control.Resize += (s, e) => ApplyResponsiveLayout(false);
            }
        }

        private void InitializePolicyWarningPanel()
        {
            _policyWarningPanel.Visible = false;
            _policyWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _policyWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _policyWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };
            Controls.Add(_policyWarningPanel);

            _policyWarningTitleLabel.AutoSize = true;
            _policyWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _policyWarningTitleLabel.Font = new Font(_policyWarningTitleLabel.Font, FontStyle.Bold);
            _policyWarningTitleLabel.Text = "\u26a0 " + Strings.PolicyWarningTitle;
            _policyWarningPanel.Controls.Add(_policyWarningTitleLabel);

            _policyWarningTextLabel.AutoSize = true;
            _policyWarningTextLabel.Text = string.Empty;
            _policyWarningPanel.Controls.Add(_policyWarningTextLabel);

            _policyWarningLinkLabel.AutoSize = true;
            _policyWarningLinkLabel.Text = Strings.PolicyWarningAdminLinkLabel;
            _policyWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _policyWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.LinkClicked += (s, e) =>
                BrowserLauncher.OpenUrl(
                    Strings.PolicyAdminGuideUrl,
                    LogCategories.Core,
                    "Failed to open policy admin guide URL.");
            _policyWarningPanel.Controls.Add(_policyWarningLinkLabel);
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
            const int fieldLeft = 150;
            const int rightMargin = 24;
            const int minWidth = 260;

            int availableWidth = _generalTab.ClientSize.Width - fieldLeft - rightMargin;
            int width = Math.Max(minWidth, availableWidth);

            _serverUrlTextBox.Width = width;
            _usernameTextBox.Width = width;
            _appPasswordTextBox.Width = width;
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = HeaderHeight;
            _headerPanel.Dock = DockStyle.Top;

            Controls.Add(_headerPanel);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ApplyResponsiveLayout(true);
            RefreshBackendPolicyStatus("settings_open");
        }

        private void ApplyResponsiveLayout(bool ensureClientWidth)
        {
            if (_layoutApplying || IsDisposed || Disposing)
            {
                return;
            }

            _layoutApplying = true;
            try
            {
                int outerPadding = ScaleLogical(12);
                int footerGap = ScaleLogical(4);
                int footerBottomPadding = ScaleLogical(8);
                var footerButtons = new List<Button> { _saveButton, _cancelButton };

                int requiredClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    footerBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);
                if (ensureClientWidth && requiredClientWidth > ClientSize.Width)
                {
                    ClientSize = new Size(requiredClientWidth, ClientSize.Height);
                }

                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    footerBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);
                int buttonTop = Math.Min(_saveButton.Top, _cancelButton.Top);

                bool hasStatus = !string.IsNullOrWhiteSpace(_statusLabel.Text);
                int statusTop = buttonTop - ScaleLogical(2);
                if (hasStatus)
                {
                    int statusHeight = Math.Max(ScaleLogical(22), _statusLabel.Font.Height + ScaleLogical(8));
                    statusTop = Math.Max(outerPadding, buttonTop - footerGap - statusHeight);
                    _statusLabel.SetBounds(
                        outerPadding,
                        statusTop,
                        Math.Max(1, ClientSize.Width - (outerPadding * 2)),
                        statusHeight);
                    _statusLabel.Visible = true;
                }
                else
                {
                    _statusLabel.Visible = false;
                    _statusLabel.SetBounds(outerPadding, statusTop, 0, 0);
                }
                int tabTop = HeaderHeight + outerPadding;
                if (_policyWarningPanel.Visible)
                {
                    int warningPadding = ScaleLogical(8);
                    int panelWidth = Math.Max(ScaleLogical(420), ClientSize.Width - (outerPadding * 2));
                    int warningTextWidth = Math.Max(ScaleLogical(220), panelWidth - (warningPadding * 2));

                    _policyWarningTitleLabel.Location = new Point(warningPadding, warningPadding);
                    _policyWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningTextTop = _policyWarningTitleLabel.Bottom + ScaleLogical(4);
                    _policyWarningTextLabel.Location = new Point(warningPadding, warningTextTop);
                    _policyWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningLinkTop = _policyWarningTextLabel.Bottom + ScaleLogical(6);
                    _policyWarningLinkLabel.Location = new Point(warningPadding, warningLinkTop);

                    int panelHeight = _policyWarningLinkLabel.Bottom + warningPadding;
                    _policyWarningPanel.SetBounds(outerPadding, tabTop, panelWidth, panelHeight);
                    tabTop = _policyWarningPanel.Bottom + ScaleLogical(8);
                }
                else
                {
                    _policyWarningPanel.SetBounds(outerPadding, tabTop, Math.Max(ScaleLogical(420), ClientSize.Width - (outerPadding * 2)), 0);
                }
                int tabBottom = hasStatus
                    ? Math.Max(tabTop + ScaleLogical(220), statusTop - ScaleLogical(4))
                    : Math.Max(tabTop + ScaleLogical(220), buttonTop - ScaleLogical(4));
                int tabHeight = Math.Max(ScaleLogical(220), tabBottom - tabTop);
                _tabControl.SetBounds(
                    outerPadding,
                    tabTop,
                    Math.Max(ScaleLogical(420), ClientSize.Width - (outerPadding * 2)),
                    tabHeight);

                ApplyGeneralTabFieldSizing();
                ApplyTalkDefaultsTabLayout();
                ApplyIfbTabLayout();
                ApplyAdvancedTabLayout();
                ApplyDebugTabLayout();
                ApplyAboutTabLayout();
                ApplyFileLinkTabLayout();
            }
            finally
            {
                _layoutApplying = false;
            }
        }

        private void ApplyTalkDefaultsTabLayout()
        {
            int groupLeft = ScaleLogical(24);
            int groupWidth = Math.Max(ScaleLogical(360), _talkTab.ClientSize.Width - (groupLeft * 2));
            _talkDefaultsGroup.SetBounds(groupLeft, _talkDefaultsGroup.Top, groupWidth, _talkDefaultsGroup.Height);

            int innerPadding = ScaleLogical(12);
            int y = ScaleLogical(26);
            int rowGap = ScaleLogical(10);
            int checkGap = ScaleLogical(6);
            int contentWidth = Math.Max(ScaleLogical(180), _talkDefaultsGroup.ClientSize.Width - (innerPadding * 2));

            int comboLeft = Math.Max(ScaleLogical(180), _talkDefaultRoomTypeLabel.PreferredSize.Width + ScaleLogical(28));
            int comboWidth = Math.Max(ScaleLogical(160), _talkDefaultsGroup.ClientSize.Width - comboLeft - ScaleLogical(12));
            int comboHeight = Math.Max(_talkDefaultRoomTypeCombo.Height, _talkDefaultRoomTypeCombo.PreferredHeight + ScaleLogical(2));
            _talkDefaultRoomTypeCombo.SetBounds(comboLeft, y - ScaleLogical(2), comboWidth, comboHeight);
            _talkDefaultRoomTypeLabel.Location = new Point(innerPadding, _talkDefaultRoomTypeCombo.Top + Math.Max(0, (comboHeight - _talkDefaultRoomTypeLabel.PreferredHeight) / 2));
            y = Math.Max(_talkDefaultRoomTypeLabel.Bottom, _talkDefaultRoomTypeCombo.Bottom) + rowGap;

            _talkDefaultPasswordCheckBox.Location = new Point(innerPadding, y);
            y = _talkDefaultPasswordCheckBox.Bottom + checkGap;

            _talkDefaultAddUsersCheckBox.Location = new Point(innerPadding, y);
            y = _talkDefaultAddUsersCheckBox.Bottom + checkGap;

            _talkDefaultAddGuestsCheckBox.Location = new Point(innerPadding, y);
            y = _talkDefaultAddGuestsCheckBox.Bottom + checkGap;

            _talkDefaultLobbyCheckBox.Location = new Point(innerPadding, y);
            y = _talkDefaultLobbyCheckBox.Bottom + checkGap;

            _talkDefaultSearchCheckBox.Location = new Point(innerPadding, y);
            y = _talkDefaultSearchCheckBox.Bottom + rowGap;

            if (_talkAddressbookWarningPanel.Visible)
            {
                int warningPadding = ScaleLogical(8);
                int warningWidth = Math.Max(ScaleLogical(160), contentWidth);
                int warningTextWidth = Math.Max(ScaleLogical(120), warningWidth - (warningPadding * 2));

                _talkAddressbookWarningTitleLabel.Location = new Point(warningPadding, warningPadding);
                _talkAddressbookWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningTextTop = _talkAddressbookWarningTitleLabel.Bottom + ScaleLogical(4);
                _talkAddressbookWarningTextLabel.Location = new Point(warningPadding, warningTextTop);
                _talkAddressbookWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningLinkTop = _talkAddressbookWarningTextLabel.Bottom + ScaleLogical(6);
                _talkAddressbookWarningLinkLabel.Location = new Point(warningPadding, warningLinkTop);

                int warningHeight = _talkAddressbookWarningLinkLabel.Bottom + warningPadding;
                _talkAddressbookWarningPanel.SetBounds(innerPadding, y, warningWidth, warningHeight);
                y = _talkAddressbookWarningPanel.Bottom + innerPadding;
            }
            else
            {
                _talkAddressbookWarningPanel.SetBounds(innerPadding, y, contentWidth, 0);
                y += innerPadding;
            }

            _talkDefaultsGroup.Height = Math.Max(ScaleLogical(200), y);
        }

        private void ApplyIfbTabLayout()
        {
            int left = ScaleLogical(24);
            int top = ScaleLogical(20);
            int rowGap = ScaleLogical(18);
            int labelToComboGap = ScaleLogical(14);

            _ifbEnabledCheckBox.Location = new Point(left, top);

            int daysLabelTop = _ifbEnabledCheckBox.Bottom + rowGap;
            _ifbDaysLabel.Location = new Point(left, daysLabelTop);

            int portLabelTop = _ifbDaysLabel.Bottom + rowGap;
            _ifbPortLabel.Location = new Point(left, portLabelTop);

            int comboLeft = Math.Max(_ifbDaysLabel.Right, _ifbPortLabel.Right) + labelToComboGap;
            int comboHeight = Math.Max(_ifbDaysCombo.Height, _ifbDaysCombo.PreferredHeight + ScaleLogical(2));
            _ifbDaysCombo.SetBounds(comboLeft, daysLabelTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbDaysCombo.Width), comboHeight);

            int portHeight = Math.Max(_ifbPortUpDown.Height, _ifbPortUpDown.PreferredHeight + ScaleLogical(2));
            _ifbPortUpDown.SetBounds(comboLeft, portLabelTop - ScaleLogical(2), Math.Max(ScaleLogical(110), _ifbPortUpDown.Width), portHeight);
        }

        private void ApplyAdvancedTabLayout()
        {
            int left = ScaleLogical(24);
            int labelToComboGap = ScaleLogical(16);
            int comboLeft = left + Math.Max(_ifbCacheHoursLabel.PreferredSize.Width, Math.Max(_shareBlockLangLabel.PreferredSize.Width, _eventDescriptionLangLabel.PreferredSize.Width)) + labelToComboGap;
            int rightMargin = ScaleLogical(24);
            int comboWidth = Math.Max(ScaleLogical(160), _advancedTab.ClientSize.Width - comboLeft - rightMargin);
            int rowTop = ScaleLogical(24);
            int rowGap = ScaleLogical(34);

            int ifbComboHeight = Math.Max(_ifbCacheHoursCombo.Height, _ifbCacheHoursCombo.PreferredHeight + ScaleLogical(2));
            _ifbCacheHoursLabel.Location = new Point(left, rowTop);
            _ifbCacheHoursCombo.SetBounds(comboLeft, rowTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbCacheHoursCombo.Width), ifbComboHeight);

            int shareLabelTop = rowTop + rowGap + ScaleLogical(12);
            int shareComboHeight = Math.Max(_shareBlockLangCombo.Height, _shareBlockLangCombo.PreferredHeight + ScaleLogical(2));
            _shareBlockLangLabel.Location = new Point(left, shareLabelTop);
            _shareBlockLangCombo.SetBounds(comboLeft, shareLabelTop - ScaleLogical(2), comboWidth, shareComboHeight);

            int eventLabelTop = _shareBlockLangLabel.Bottom + rowGap;
            int eventComboHeight = Math.Max(_eventDescriptionLangCombo.Height, _eventDescriptionLangCombo.PreferredHeight + ScaleLogical(2));
            _eventDescriptionLangLabel.Location = new Point(left, eventLabelTop);
            _eventDescriptionLangCombo.SetBounds(comboLeft, eventLabelTop - ScaleLogical(2), comboWidth, eventComboHeight);

            int groupTop = _eventDescriptionLangCombo.Bottom + ScaleLogical(14);
            int groupWidth = Math.Max(ScaleLogical(320), _advancedTab.ClientSize.Width - left - rightMargin);
            _tlsSettingsGroup.SetBounds(left, groupTop, groupWidth, ScaleLogical(134));
            _tlsHintLabel.MaximumSize = new Size(Math.Max(ScaleLogical(200), _tlsSettingsGroup.ClientSize.Width - ScaleLogical(20)), 0);
            _tlsHintLabel.AutoSize = true;
        }

        private void ApplyDebugTabLayout()
        {
            int rightMargin = ScaleLogical(24);
            _debugAnonymizeCheckBox.Location = new Point(_debugLogCheckBox.Left, _debugLogCheckBox.Bottom + ScaleLogical(8));
            _debugPathLabel.Location = new Point(_debugPathLabel.Left, _debugAnonymizeCheckBox.Bottom + ScaleLogical(12));
            int width = Math.Max(ScaleLogical(220), _debugTab.ClientSize.Width - rightMargin - _debugPathLabel.Left);
            _debugPathLabel.MaximumSize = new Size(width, 0);
            _debugPathLabel.AutoSize = true;
            _debugOpenLink.Location = new Point(_debugOpenLink.Left, _debugPathLabel.Bottom + ScaleLogical(10));
        }

        private void ApplyAboutTabLayout()
        {
            int left = ScaleLogical(18);
            int right = ScaleLogical(18);
            int top = ScaleLogical(20);
            int gap = ScaleLogical(10);
            int contentWidth = Math.Max(ScaleLogical(220), _aboutTab.ClientSize.Width - left - right);

            _aboutVersionLabel.Location = new Point(left, top);
            _aboutCopyrightLabel.Location = new Point(left, _aboutVersionLabel.Bottom + gap);

            _aboutLicenseLabel.Location = new Point(left, _aboutCopyrightLabel.Bottom + gap);
            _aboutLicenseLink.Location = new Point(_aboutLicenseLabel.Right + ScaleLogical(8), _aboutLicenseLabel.Top);

            _aboutHomepageLabel.Location = new Point(left, _aboutLicenseLabel.Bottom + gap);
            _aboutHomepageLink.Location = new Point(_aboutHomepageLabel.Right + ScaleLogical(8), _aboutHomepageLabel.Top);

            _aboutOverviewLabel.Location = new Point(left, _aboutHomepageLabel.Bottom + ScaleLogical(16));
            _aboutOverviewLabel.MaximumSize = new Size(contentWidth, 0);
            _aboutOverviewLabel.AutoSize = true;

            _aboutMoreInfoLabel.Location = new Point(left, _aboutOverviewLabel.Bottom + gap);
            _aboutMoreInfoLink.Location = new Point(left, _aboutMoreInfoLabel.Bottom + ScaleLogical(6));

            _aboutSupportNoteLabel.Location = new Point(left, _aboutMoreInfoLink.Bottom + ScaleLogical(16));
            _aboutSupportNoteLabel.MaximumSize = new Size(contentWidth, 0);
            _aboutSupportNoteLabel.AutoSize = true;

            _aboutSupportHeadingLabel.Location = new Point(left, _aboutSupportNoteLabel.Bottom + gap);
            _aboutSupportLink.Location = new Point(left, _aboutSupportHeadingLabel.Bottom + ScaleLogical(8));
        }

        private void ApplyFileLinkTabLayout()
        {
            int left = ScaleLogical(18);
            int tabContentWidth = Math.Max(ScaleLogical(420), _fileLinkTab.ClientSize.Width - (left * 2));

            _fileLinkBaseTextBox.Width = tabContentWidth;
            _fileLinkBaseHintLabel.MaximumSize = new Size(tabContentWidth, 0);
            _fileLinkBaseHintLabel.AutoSize = true;

            int sharingGroupTop = _fileLinkBaseHintLabel.Bottom + ScaleLogical(12);
            _sharingDefaultsGroup.SetBounds(left, sharingGroupTop, tabContentWidth, _sharingDefaultsGroup.Height);

            int groupWidth = _sharingDefaultsGroup.ClientSize.Width;
            _sharingDefaultShareNameTextBox.Width = Math.Max(ScaleLogical(220), groupWidth - ScaleLogical(24));

            int leftColumnX = ScaleLogical(18);
            int rightColumnX = Math.Max(ScaleLogical(260), (groupWidth / 2) + ScaleLogical(8));
            int rightColumnRequired = Math.Max(
                Math.Max(_sharingDefaultPasswordCheckBox.PreferredSize.Width, _sharingDefaultPasswordSeparateCheckBox.PreferredSize.Width),
                _sharingDefaultExpireDaysLabel.PreferredSize.Width + _sharingDefaultExpireDaysUpDown.Width + ScaleLogical(12));
            bool stackRightColumn = rightColumnX + rightColumnRequired + ScaleLogical(12) > groupWidth;
            int rightColumnTop = ScaleLogical(114);
            if (stackRightColumn)
            {
                rightColumnX = leftColumnX;
                rightColumnTop = _sharingDefaultPermDeleteCheckBox.Bottom + ScaleLogical(14);
            }

            _sharingDefaultPasswordCheckBox.Location = new Point(rightColumnX, rightColumnTop);
            _sharingDefaultPasswordSeparateCheckBox.Location = new Point(rightColumnX, _sharingDefaultPasswordCheckBox.Bottom + ScaleLogical(12));
            _sharingDefaultExpireDaysLabel.Location = new Point(rightColumnX, _sharingDefaultPasswordSeparateCheckBox.Bottom + ScaleLogical(12));
            _sharingDefaultExpireDaysUpDown.Location = new Point(rightColumnX, _sharingDefaultExpireDaysLabel.Bottom + ScaleLogical(6));

            int columnsBottom = Math.Max(_sharingDefaultPermDeleteCheckBox.Bottom, _sharingDefaultExpireDaysUpDown.Bottom);
            int automationTop = columnsBottom + ScaleLogical(18);
            int automationWidth = Math.Max(ScaleLogical(320), groupWidth - ScaleLogical(24));
            _sharingAttachmentAutomationGroup.SetBounds(ScaleLogical(12), automationTop, automationWidth, _sharingAttachmentAutomationGroup.Height);
            int lockTextWidth = Math.Max(ScaleLogical(180), automationWidth - ScaleLogical(24));
            _sharingAttachmentLockHintLabel.MaximumSize = new Size(lockTextWidth, 0);
            _sharingAttachmentLockStepsLabel.MaximumSize = new Size(lockTextWidth, 0);
            _sharingAttachmentLockHintLabel.AutoSize = true;
            _sharingAttachmentLockStepsLabel.AutoSize = true;

            UpdateSharingAttachmentOptionsState();
            int sharingDefaultsHeight = _sharingAttachmentAutomationGroup.Bottom + ScaleLogical(12);
            _sharingDefaultsGroup.Height = Math.Max(ScaleLogical(300), sharingDefaultsHeight);
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
            _ifbDaysLabel.AutoSize = true;
            _ifbTab.Controls.Add(_ifbDaysLabel);

            _ifbDaysCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbDaysCombo.IntegralHeight = false;
            _ifbDaysCombo.Location = new Point(200, 58);
            _ifbDaysCombo.Width = 100;
            _ifbDaysCombo.Items.AddRange(new object[] { "10", "30", "60", "90" });
            _ifbTab.Controls.Add(_ifbDaysCombo);

            _ifbPortLabel.Text = Strings.LabelIfbPort;
            _ifbPortLabel.Location = new Point(24, 94);
            _ifbPortLabel.AutoSize = true;
            _ifbTab.Controls.Add(_ifbPortLabel);

            _ifbPortUpDown.Location = new Point(200, 92);
            _ifbPortUpDown.Width = 110;
            _ifbPortUpDown.Minimum = AddinSettings.MinIfbPort;
            _ifbPortUpDown.Maximum = AddinSettings.MaxIfbPort;
            _ifbPortUpDown.Value = AddinSettings.DefaultIfbPort;
            _ifbTab.Controls.Add(_ifbPortUpDown);
        }

        private void InitializeAdvancedTab()
        {
            _advancedTab.AutoScroll = true;
            _advancedTab.Padding = new Padding(12);

            _ifbCacheHoursLabel.Text = Strings.LabelIfbCacheHours;
            _ifbCacheHoursLabel.Location = new Point(24, 24);
            _ifbCacheHoursLabel.AutoSize = true;
            _advancedTab.Controls.Add(_ifbCacheHoursLabel);

            _ifbCacheHoursCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbCacheHoursCombo.IntegralHeight = false;
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
            _shareBlockLangLabel.AutoSize = true;
            _advancedTab.Controls.Add(_shareBlockLangLabel);

            _shareBlockLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _shareBlockLangCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _shareBlockLangCombo.IntegralHeight = false;
            _shareBlockLangCombo.Location = new Point(260, langTop - 2);
            _shareBlockLangCombo.Width = 240;
            _shareBlockLangCombo.DrawItem += HandleLanguageComboDrawItem;
            _shareBlockLangCombo.SelectionChangeCommitted += HandleLanguageComboSelectionCommitted;
            PopulateLanguageOverrideCombo(_shareBlockLangCombo, "share");
            _advancedTab.Controls.Add(_shareBlockLangCombo);

            _eventDescriptionLangLabel.Text = Strings.AdvancedEventDescriptionLangLabel;
            _eventDescriptionLangLabel.Location = new Point(24, langTop + 32);
            _eventDescriptionLangLabel.AutoSize = true;
            _advancedTab.Controls.Add(_eventDescriptionLangLabel);

            _eventDescriptionLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _eventDescriptionLangCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _eventDescriptionLangCombo.IntegralHeight = false;
            _eventDescriptionLangCombo.Location = new Point(260, langTop + 30);
            _eventDescriptionLangCombo.Width = 240;
            _eventDescriptionLangCombo.DrawItem += HandleLanguageComboDrawItem;
            _eventDescriptionLangCombo.SelectionChangeCommitted += HandleLanguageComboSelectionCommitted;
            PopulateLanguageOverrideCombo(_eventDescriptionLangCombo, "talk");
            _advancedTab.Controls.Add(_eventDescriptionLangCombo);

            _tlsSettingsGroup.Text = Strings.AdvancedTlsHeading;
            _tlsSettingsGroup.Location = new Point(24, langTop + 72);
            _tlsSettingsGroup.Size = new Size(520, 132);
            _advancedTab.Controls.Add(_tlsSettingsGroup);

            _tlsUseSystemDefaultCheckBox.Text = Strings.AdvancedTlsUseSystemDefaultLabel;
            _tlsUseSystemDefaultCheckBox.AutoSize = true;
            _tlsUseSystemDefaultCheckBox.Location = new Point(12, 24);
            _tlsUseSystemDefaultCheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsUseSystemDefaultCheckBox);

            _tlsEnable12CheckBox.Text = Strings.AdvancedTlsEnable12Label;
            _tlsEnable12CheckBox.AutoSize = true;
            _tlsEnable12CheckBox.Location = new Point(12, 48);
            _tlsEnable12CheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsEnable12CheckBox);

            _tlsEnable13CheckBox.Text = Strings.AdvancedTlsEnable13Label;
            _tlsEnable13CheckBox.AutoSize = true;
            _tlsEnable13CheckBox.Location = new Point(12, 72);
            _tlsEnable13CheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsEnable13CheckBox);

            _tlsHintLabel.Text = Strings.AdvancedTlsHint;
            _tlsHintLabel.AutoSize = true;
            _tlsHintLabel.MaximumSize = new Size(492, 0);
            _tlsHintLabel.Location = new Point(12, 94);
            _tlsHintLabel.ForeColor = Color.DimGray;
            _tlsSettingsGroup.Controls.Add(_tlsHintLabel);
        }

        private void InitializeTalkTab()
        {
            _talkTab.AutoScroll = true;
            _talkTab.Padding = new Padding(12);

            _talkDefaultsGroup.Text = Strings.SettingsTalkDefaultsGroup;
            _talkDefaultsGroup.Location = new Point(24, 20);
            _talkDefaultsGroup.Size = new Size(480, 248);
            _talkTab.Controls.Add(_talkDefaultsGroup);

            _talkDefaultRoomTypeLabel.Text = Strings.TalkRoomGroup;
            _talkDefaultRoomTypeLabel.Location = new Point(12, 28);
            _talkDefaultRoomTypeLabel.AutoSize = true;
            _talkDefaultsGroup.Controls.Add(_talkDefaultRoomTypeLabel);

            _talkDefaultRoomTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _talkDefaultRoomTypeCombo.IntegralHeight = false;
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

            _talkAddressbookWarningPanel.Visible = false;
            _talkAddressbookWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _talkAddressbookWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _talkAddressbookWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };
            _talkDefaultsGroup.Controls.Add(_talkAddressbookWarningPanel);

            _talkAddressbookWarningTitleLabel.AutoSize = true;
            _talkAddressbookWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _talkAddressbookWarningTitleLabel.Font = new Font(
                _talkAddressbookWarningTitleLabel.Font,
                FontStyle.Bold);
            _talkAddressbookWarningTitleLabel.Text = "\u26a0 " + Strings.TalkSystemAddressbookRequiredShort;
            _talkAddressbookWarningPanel.Controls.Add(_talkAddressbookWarningTitleLabel);

            _talkAddressbookWarningTextLabel.AutoSize = true;
            _talkAddressbookWarningTextLabel.Text = Strings.TalkSystemAddressbookRequiredMessage;
            _talkAddressbookWarningPanel.Controls.Add(_talkAddressbookWarningTextLabel);

            _talkAddressbookWarningLinkLabel.AutoSize = true;
            _talkAddressbookWarningLinkLabel.Text = Strings.TalkSystemAddressbookAdminLinkLabel;
            _talkAddressbookWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _talkAddressbookWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _talkAddressbookWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _talkAddressbookWarningLinkLabel.LinkClicked += (s, e) =>
                BrowserLauncher.OpenUrl(
                    Strings.TalkSystemAddressbookAdminGuideUrl,
                    LogCategories.Core,
                    "Failed to open system address book admin guide URL.");
            _talkAddressbookWarningPanel.Controls.Add(_talkAddressbookWarningLinkLabel);

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

            _debugAnonymizeCheckBox.Text = Strings.DebugAnonymizeCheckbox;
            _debugAnonymizeCheckBox.AutoSize = true;
            _debugAnonymizeCheckBox.Location = new Point(24, 50);
            _debugTab.Controls.Add(_debugAnonymizeCheckBox);

            _debugPathLabel.AutoSize = true;
            _debugPathLabel.Location = new Point(24, 90);
            _debugPathLabel.MaximumSize = new Size(420, 0);
            _debugTab.Controls.Add(_debugPathLabel);

            _debugOpenLink.Text = Strings.DebugOpenLog;
            _debugOpenLink.Location = new Point(24, 140);
            _debugOpenLink.AutoSize = true;
            _debugOpenLink.LinkClicked += OnDebugOpenLinkClicked;
            _debugTab.Controls.Add(_debugOpenLink);

            UpdateDebugPathLabel();
        }

        private void InitializeAboutTab()
        {
            _aboutTab.Padding = new Padding(12);
            _aboutTab.AutoScroll = true;

            _aboutVersionLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutVersionLabel);

            _aboutCopyrightLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutCopyrightLabel);

            _aboutLicenseLabel.Text = Strings.AboutLicenseLabel;
            _aboutLicenseLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutLicenseLabel);

            _aboutLicenseLink.AutoSize = true;
            _aboutLicenseLink.Text = Strings.AboutLicenseLink;
            _aboutLicenseLink.Padding = new Padding(1, 0, 0, 0);
            _aboutLicenseLink.LinkClicked += OnAboutLicenseLinkClicked;
            _aboutTab.Controls.Add(_aboutLicenseLink);

            _aboutHomepageLabel.Text = Strings.AboutHomepageLabel;
            _aboutHomepageLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutHomepageLabel);

            _aboutHomepageLink.Text = Strings.AboutHomepageLink;
            _aboutHomepageLink.AutoSize = true;
            _aboutHomepageLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    NcConnectorHomepageUrl,
                    LogCategories.Core,
                    "Failed to open NC Connector homepage URL.");
            _aboutTab.Controls.Add(_aboutHomepageLink);

            _aboutOverviewLabel.Text = Strings.AboutOverviewText;
            _aboutOverviewLabel.AutoSize = true;
            _aboutOverviewLabel.MaximumSize = new Size(Math.Max(ScaleLogical(220), _aboutTab.ClientSize.Width - ScaleLogical(36)), 0);
            _aboutTab.Controls.Add(_aboutOverviewLabel);

            _aboutMoreInfoLabel.Text = Strings.AboutMoreInfoLabel;
            _aboutMoreInfoLabel.AutoSize = true;
            _aboutMoreInfoLabel.Font = new Font(_aboutMoreInfoLabel.Font, FontStyle.Bold);
            _aboutTab.Controls.Add(_aboutMoreInfoLabel);

            _aboutMoreInfoLink.Text = Strings.AboutMoreInfoLink;
            _aboutMoreInfoLink.AutoSize = true;
            _aboutMoreInfoLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    NcConnectorHomepageUrl,
                    LogCategories.Core,
                    "Failed to open NC Connector information URL.");
            _aboutTab.Controls.Add(_aboutMoreInfoLink);

            _aboutSupportNoteLabel.Text = Strings.AboutSupportNote;
            _aboutSupportNoteLabel.AutoSize = true;
            _aboutSupportNoteLabel.ForeColor = Color.DimGray;
            _aboutTab.Controls.Add(_aboutSupportNoteLabel);

            _aboutSupportHeadingLabel.Text = Strings.AboutSupportHeading;
            _aboutSupportHeadingLabel.AutoSize = true;
            _aboutSupportHeadingLabel.Font = new Font(_aboutSupportHeadingLabel.Font, FontStyle.Bold);
            _aboutTab.Controls.Add(_aboutSupportHeadingLabel);

            _aboutSupportLink.Text = Strings.AboutSupportLink;
            _aboutSupportLink.AutoSize = true;
            _aboutSupportLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    "https://www.paypal.com/donate/?hosted_button_id=FTZWNRNKVKUN6",
                    LogCategories.Core,
                    "Failed to open support donation URL.");
            _aboutTab.Controls.Add(_aboutSupportLink);

            ApplyAboutTabLayout();
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
                var version = assembly.GetName().Version;                if (version != null)
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
            if (File.Exists(licensePath))
            {
                string openError;
                bool opened = BrowserLauncher.OpenTarget(
                    licensePath,
                    LogCategories.Core,
                    "Failed to open license file.",
                    out openError);
                if (!opened)
                {
                    MessageBox.Show(
                        string.Format(Strings.LicenseFileOpenErrorMessage, openError ?? string.Empty),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return;
            }

            MessageBox.Show(
                string.Format(Strings.LicenseFileMissingMessage, licensePath),
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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
            _suppressImmediateTlsApply = true;
            try
            {
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
                _ifbPortUpDown.Value = Math.Max(
                    _ifbPortUpDown.Minimum,
                    Math.Min(_ifbPortUpDown.Maximum, AddinSettings.NormalizeIfbPort(Result.IfbPort)));
                SelectComboValue(_ifbCacheHoursCombo, Result.IfbCacheHours, 24);
                _debugLogCheckBox.Checked = Result.DebugLoggingEnabled;
                _debugAnonymizeCheckBox.Checked = Result.LogAnonymizationEnabled;
                _tlsUseSystemDefaultCheckBox.Checked = Result.TransportTlsUseSystemDefault;
                _tlsEnable12CheckBox.Checked = Result.TransportTlsEnable12;
                _tlsEnable13CheckBox.Checked = Result.TransportTlsEnable13;
                _fileLinkBaseTextBox.Text = Result.FileLinkBasePath ?? string.Empty;
                _sharingDefaultShareNameTextBox.Text = Result.SharingDefaultShareName ?? string.Empty;
                _sharingDefaultPermCreateCheckBox.Checked = Result.SharingDefaultPermCreate;
                _sharingDefaultPermWriteCheckBox.Checked = Result.SharingDefaultPermWrite;
                _sharingDefaultPermDeleteCheckBox.Checked = Result.SharingDefaultPermDelete;
                _sharingDefaultPasswordCheckBox.Checked = Result.SharingDefaultPasswordEnabled;
                _sharingDefaultPasswordSeparateCheckBox.Checked = HasBackendSeatEntitlement() && Result.SharingDefaultPasswordSeparateEnabled;
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
                _sharingAttachmentsAlwaysCheckBox.Checked = Result.SharingAttachmentsAlwaysConnector;
                _sharingAttachmentsOfferAboveCheckBox.Checked = Result.SharingAttachmentsOfferAboveEnabled;
                int offerAboveMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(Result.SharingAttachmentsOfferAboveMb);
                decimal clampedOfferAbove = Math.Max(
                    _sharingAttachmentsOfferAboveMbUpDown.Minimum,
                    Math.Min(_sharingAttachmentsOfferAboveMbUpDown.Maximum, (decimal)offerAboveMb));
                _sharingAttachmentsOfferAboveMbUpDown.Value = clampedOfferAbove;
                _talkDefaultPasswordCheckBox.Checked = Result.TalkDefaultPasswordEnabled;
                _talkDefaultAddUsersCheckBox.Checked = Result.TalkDefaultAddUsers;
                _talkDefaultAddGuestsCheckBox.Checked = Result.TalkDefaultAddGuests;
                _talkDefaultLobbyCheckBox.Checked = Result.TalkDefaultLobbyEnabled;
                _talkDefaultSearchCheckBox.Checked = Result.TalkDefaultSearchVisible;
                SelectTalkRoomType(Result.TalkDefaultRoomType);
                UpdateTalkRoomTypeTooltip();
                RefreshLanguageOverrideCombos(Result.ShareBlockLang, Result.EventDescriptionLang);
                UpdateDebugPathLabel();
                UpdateAboutTab();
                RefreshSharingAttachmentLockState();
                UpdateSharingAttachmentOptionsState();
                UpdateTlsOptionsState();
                ApplyBackendPolicyStatus("settings_init");
                RefreshTalkSystemAddressbookState(true, "settings_open");
            }
            finally
            {
                _suppressImmediateTlsApply = false;
            }
        }

        private void OnSaveButtonClick(object sender, EventArgs e)
        {
            RefreshBackendPolicyStatus("settings_save");
            RefreshTalkSystemAddressbookState(true, "settings_save");

            if (!_tlsUseSystemDefaultCheckBox.Checked
                && !_tlsEnable12CheckBox.Checked
                && !_tlsEnable13CheckBox.Checked)
            {
                MessageBox.Show(
                    Strings.TransportTlsSelectionRequired,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Result.ServerUrl = _serverUrlTextBox.Text.Trim();
            Result.Username = _usernameTextBox.Text.Trim();
            Result.AppPassword = _appPasswordTextBox.Text;
            Result.AuthMode = _loginFlowRadio.Checked ? AuthenticationMode.LoginFlow : AuthenticationMode.Manual;
            Result.IfbEnabled = _ifbEnabledCheckBox.Checked;
            Result.IfbDays = ParseComboValue(_ifbDaysCombo, 30);
            Result.IfbPort = AddinSettings.NormalizeIfbPort((int)_ifbPortUpDown.Value);
            Result.IfbCacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            Result.DebugLoggingEnabled = _debugLogCheckBox.Checked;
            Result.LogAnonymizationEnabled = _debugAnonymizeCheckBox.Checked;
            Result.TransportTlsUseSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
            Result.TransportTlsEnable12 = _tlsEnable12CheckBox.Checked;
            Result.TransportTlsEnable13 = _tlsEnable13CheckBox.Checked;
            Result.LastKnownServerVersion = _lastKnownServerVersion ?? string.Empty;
            Result.FileLinkBasePath = _fileLinkBaseTextBox.Text.Trim();
            Result.SharingDefaultShareName = _sharingDefaultShareNameTextBox.Text.Trim();
            Result.SharingDefaultPermCreate = _sharingDefaultPermCreateCheckBox.Checked;
            Result.SharingDefaultPermWrite = _sharingDefaultPermWriteCheckBox.Checked;
            Result.SharingDefaultPermDelete = _sharingDefaultPermDeleteCheckBox.Checked;
            Result.SharingDefaultPasswordEnabled = _sharingDefaultPasswordCheckBox.Checked;
            Result.SharingDefaultPasswordSeparateEnabled =
                HasBackendSeatEntitlement() && _sharingDefaultPasswordSeparateCheckBox.Checked;
            Result.SharingDefaultExpireDays = (int)_sharingDefaultExpireDaysUpDown.Value;
            Result.SharingAttachmentsAlwaysConnector = _sharingAttachmentsAlwaysCheckBox.Checked;
            Result.SharingAttachmentsOfferAboveEnabled = _sharingAttachmentsOfferAboveCheckBox.Checked;
            Result.SharingAttachmentsOfferAboveMb = (int)_sharingAttachmentsOfferAboveMbUpDown.Value;
            Result.TalkDefaultPasswordEnabled = _talkDefaultPasswordCheckBox.Checked;
            Result.TalkDefaultAddUsers = _talkDefaultAddUsersCheckBox.Checked;
            Result.TalkDefaultAddGuests = _talkDefaultAddGuestsCheckBox.Checked;
            Result.TalkDefaultLobbyEnabled = _talkDefaultLobbyCheckBox.Checked;
            Result.TalkDefaultSearchVisible = _talkDefaultSearchCheckBox.Checked;
            Result.TalkDefaultRoomType = GetSelectedTalkRoomType();
            Result.ShareBlockLang = GetSelectedLanguageChoice(_shareBlockLangCombo);
            Result.EventDescriptionLang = GetSelectedLanguageChoice(_eventDescriptionLangCombo);
        }

        // WinForms event handlers must stay async void; keep awaited flow inside this method-level try/catch.
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
            SecurityProtocolType previousSecurityProtocol = ServicePointManager.SecurityProtocol;
            bool temporaryTlsApplied = false;

            try
            {
                previousSecurityProtocol = ApplyTemporaryTlsForConnectivity("settings_login_flow");
                temporaryTlsApplied = true;

                var flowService = new TalkLoginFlowService(normalizedUrl);
                var startInfo = flowService.StartLoginFlow();

                BrowserLauncher.OpenUrl(
                    startInfo.LoginUrl,
                    LogCategories.Core,
                    "Failed to open browser for login flow URL.");
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
                UpdateControlState();
            }
        }

        // WinForms event handlers must stay async void; keep awaited flow inside this method-level try/catch.
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
            SecurityProtocolType previousSecurityProtocol = ServicePointManager.SecurityProtocol;
            bool temporaryTlsApplied = false;

            try
            {
                previousSecurityProtocol = ApplyTemporaryTlsForConnectivity("settings_connection_test");
                temporaryTlsApplied = true;

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
                UpdateControlState();
            }
        }

        private SecurityProtocolType ApplyTemporaryTlsForConnectivity(string source)
        {
            SecurityProtocolType previous = ServicePointManager.SecurityProtocol;
            try
            {
                TransportSecurityConfigurator.Apply(
                    _tlsUseSystemDefaultCheckBox.Checked,
                    _tlsEnable12CheckBox.Checked,
                    _tlsEnable13CheckBox.Checked,
                    source);
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
            ServicePointManager.SecurityProtocol = previous;
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
            string message = ex != null && !string.IsNullOrWhiteSpace(ex.Message)
                ? ex.Message.Trim()
                : Strings.StatusTestFailureUnknown;

            string summary = ExtractFirstLine(message);
            SetStatus(string.Format(statusFormat, summary), true);            if (ex != null && ex.IsTransportError)
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

        private void OnSelectedTabChanged(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == _fileLinkTab)
            {
                RefreshSharingAttachmentLockState();
                UpdateSharingAttachmentOptionsState();
                return;
            }
            if (_tabControl.SelectedTab == _talkTab)
            {
                RefreshTalkSystemAddressbookState(true, "settings_tab_talk");
            }
        }

        private bool IsPolicyActive()
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.PolicyActive;
        }

        private bool IsPolicyLocked(string domain, string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked(domain, key);
        }

        /**
         * Return true when the backend endpoint exists.
         */
        private bool IsBackendEndpointAvailable()
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.EndpointAvailable;
        }

        /**
         * Return true when the current user has an active backend seat.
         * This gate controls premium-only UI features independently from policy locks.
         */
        private bool HasBackendSeatEntitlement()
        {
            return _backendPolicyStatus != null
                   && _backendPolicyStatus.EndpointAvailable
                   && _backendPolicyStatus.SeatAssigned
                   && _backendPolicyStatus.IsValid
                   && string.Equals(_backendPolicyStatus.SeatState, "active", StringComparison.OrdinalIgnoreCase);
        }

        /**
         * Return the tooltip shown when separate password delivery is unavailable.
         */
        private string GetSeparatePasswordUnavailableTooltip()
        {            if (_backendPolicyStatus == null || !_backendPolicyStatus.EndpointAvailable)
            {
                return Strings.SharingPasswordSeparateBackendRequiredTooltip;
            }
            if (!_backendPolicyStatus.SeatAssigned)
            {
                return Strings.SharingPasswordSeparateNoSeatTooltip;
            }
            if (!_backendPolicyStatus.IsValid
                || !string.Equals(_backendPolicyStatus.SeatState, "active", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.SharingPasswordSeparatePausedTooltip;
            }
            return string.Empty;
        }

        private void RefreshBackendPolicyStatus(string trigger)
        {
            string serverUrl = _serverUrlTextBox.Text.Trim();
            string username = _usernameTextBox.Text.Trim();
            string appPassword = _appPasswordTextBox.Text ?? string.Empty;
            var configuration = new TalkServiceConfiguration(serverUrl, username, appPassword);

            try
            {
                var service = new BackendPolicyService(configuration);
                _backendPolicyStatus = service.FetchStatus();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Backend policy status check failed in settings.", ex);
                _backendPolicyStatus = null;
            }

            ApplyBackendPolicyStatus(trigger);
        }

        private void ApplyBackendPolicyStatus(string trigger)
        {
            bool warningVisible = _backendPolicyStatus != null
                                  && _backendPolicyStatus.WarningVisible
                                  && !string.IsNullOrWhiteSpace(_backendPolicyStatus.WarningMessage);
            string currentShareLanguage = GetSelectedLanguageChoice(_shareBlockLangCombo);
            string currentTalkLanguage = GetSelectedLanguageChoice(_eventDescriptionLangCombo);
            _policyWarningPanel.Visible = warningVisible;
            _policyWarningTextLabel.Text = warningVisible ? _backendPolicyStatus.WarningMessage : string.Empty;
            RefreshLanguageOverrideCombos(currentShareLanguage, currentTalkLanguage);

            if (IsPolicyActive())
            {
                ApplyPolicyDefaultsToControls();
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Policy status applied in settings (trigger=" + (trigger ?? "n/a")
                + ", active=" + IsPolicyActive().ToString(CultureInfo.InvariantCulture)
                + ", warningVisible=" + warningVisible.ToString(CultureInfo.InvariantCulture)
                + ", mode=" + (_backendPolicyStatus != null ? _backendPolicyStatus.Mode : "local")
                + ", reason=" + (_backendPolicyStatus != null ? _backendPolicyStatus.Reason : "n/a")
                + ").");

            UpdateControlState();
            ApplyResponsiveLayout(false);
        }

        private void ApplyPolicyDefaultsToControls()
        {
            if (!IsPolicyActive())
            {
                return;
            }
            bool policyBool;
            int policyInt;
            string policyString;

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_base_directory");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _fileLinkBaseTextBox.Text = policyString;
            }

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_name_template");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _sharingDefaultShareNameTextBox.Text = policyString;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_upload", out policyBool))
            {
                _sharingDefaultPermCreateCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_edit", out policyBool))
            {
                _sharingDefaultPermWriteCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_delete", out policyBool))
            {
                _sharingDefaultPermDeleteCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_set_password", out policyBool))
            {
                _sharingDefaultPasswordCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_send_password_separately", out policyBool))
            {
                _sharingDefaultPasswordSeparateCheckBox.Checked = policyBool;
            }
            if (!HasBackendSeatEntitlement())
            {
                _sharingDefaultPasswordSeparateCheckBox.Checked = false;
            }
            if (_backendPolicyStatus.TryGetPolicyInt("share", "share_expire_days", out policyInt))
            {
                decimal clamped = Math.Max(_sharingDefaultExpireDaysUpDown.Minimum, Math.Min(_sharingDefaultExpireDaysUpDown.Maximum, policyInt));
                _sharingDefaultExpireDaysUpDown.Value = clamped;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "attachments_always_via_ncconnector", out policyBool))
            {
                _sharingAttachmentsAlwaysCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyInt("share", "attachments_min_size_mb", out policyInt))
            {
                int normalizedThreshold = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(policyInt);
                decimal clampedThreshold = Math.Max(
                    _sharingAttachmentsOfferAboveMbUpDown.Minimum,
                    Math.Min(_sharingAttachmentsOfferAboveMbUpDown.Maximum, normalizedThreshold));
                _sharingAttachmentsOfferAboveMbUpDown.Value = clampedThreshold;
                _sharingAttachmentsOfferAboveCheckBox.Checked = true;
            }
            else if (_backendPolicyStatus.HasPolicyKey("share", "attachments_min_size_mb"))
            {
                _sharingAttachmentsOfferAboveCheckBox.Checked = false;
            }

            policyString = _backendPolicyStatus.GetPolicyString("share", "language_share_html_block");
            if (IsPolicyLocked("share", "language_share_html_block")
                && !string.IsNullOrWhiteSpace(policyString))
            {
                SelectLanguageChoice(_shareBlockLangCombo, policyString);
            }
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_set_password", out policyBool))
            {
                _talkDefaultPasswordCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_users", out policyBool))
            {
                _talkDefaultAddUsersCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_guests", out policyBool))
            {
                _talkDefaultAddGuestsCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_lobby_active", out policyBool))
            {
                _talkDefaultLobbyCheckBox.Checked = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_show_in_search", out policyBool))
            {
                _talkDefaultSearchCheckBox.Checked = policyBool;
            }

            policyString = _backendPolicyStatus.GetPolicyString("talk", "talk_room_type");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                SelectTalkRoomType(
                    string.Equals(policyString.Trim(), "event", StringComparison.OrdinalIgnoreCase)
                    ? TalkRoomType.EventConversation
                    : TalkRoomType.StandardRoom);
            }

            policyString = _backendPolicyStatus.GetPolicyString("talk", "language_talk_description");
            if (IsPolicyLocked("talk", "language_talk_description")
                && !string.IsNullOrWhiteSpace(policyString))
            {
                SelectLanguageChoice(_eventDescriptionLangCombo, policyString);
            }
        }

        private void RefreshSharingAttachmentLockState()
        {
            try
            {
                var state = _attachmentGuardService.ReadLiveState();
                _sharingAttachmentLockActive = state != null && state.LockActive;
                _sharingAttachmentLockThresholdMb = state != null
                    ? OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(state.ThresholdMb)
                    : 5;
            }
            catch (Exception ex)
            {
                _sharingAttachmentLockActive = false;
                _sharingAttachmentLockThresholdMb = 5;
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read live host attachment automation guard state.", ex);
            }
        }

        private void UpdateSharingAttachmentOptionsState()
        {
            bool alwaysConnector = _sharingAttachmentsAlwaysCheckBox.Checked;
            bool lockActive = _sharingAttachmentLockActive;
            bool policyLockAlways = IsPolicyLocked("share", "attachments_always_via_ncconnector");
            bool policyLockThreshold = IsPolicyLocked("share", "attachments_min_size_mb");
            bool uiBusy = _isBusy;
            bool effectiveAlwaysLock = lockActive || policyLockAlways;
            bool effectiveThresholdLock = lockActive || policyLockThreshold;

            _sharingAttachmentLockHintLabel.Visible = lockActive;
            _sharingAttachmentLockStepsLabel.Visible = lockActive;
            int innerPadding = ScaleLogical(12);
            int y = ScaleLogical(24);
            int rowGap = ScaleLogical(8);
            int lockTextWidth = Math.Max(ScaleLogical(180), _sharingAttachmentAutomationGroup.ClientSize.Width - (innerPadding * 2));

            if (lockActive)
            {
                _sharingAttachmentLockHintLabel.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.SharingAttachmentsLockText,
                    _sharingAttachmentLockThresholdMb.ToString(CultureInfo.CurrentCulture));
                _sharingAttachmentLockStepsLabel.Text = string.Join(
                    Environment.NewLine,
                    Strings.SharingAttachmentsLockStep1,
                    Strings.SharingAttachmentsLockStep2,
                    Strings.SharingAttachmentsLockStep3);

                _sharingAttachmentLockHintLabel.MaximumSize = new Size(lockTextWidth, 0);
                _sharingAttachmentLockHintLabel.Location = new Point(innerPadding, y);
                y = _sharingAttachmentLockHintLabel.Bottom + rowGap;

                _sharingAttachmentLockStepsLabel.MaximumSize = new Size(lockTextWidth, 0);
                _sharingAttachmentLockStepsLabel.Location = new Point(innerPadding, y);
                y = _sharingAttachmentLockStepsLabel.Bottom + ScaleLogical(12);
            }
            else
            {
                _sharingAttachmentLockHintLabel.Text = string.Empty;
                _sharingAttachmentLockStepsLabel.Text = string.Empty;
            }

            _sharingAttachmentsAlwaysCheckBox.Location = new Point(innerPadding, y);
            int offerTop = _sharingAttachmentsAlwaysCheckBox.Bottom + rowGap;
            _sharingAttachmentsOfferAboveCheckBox.Location = new Point(innerPadding, offerTop);

            int spinnerLeft = _sharingAttachmentsOfferAboveCheckBox.Right + ScaleLogical(10);
            int spinnerTop = offerTop - ScaleLogical(2);
            int unitWidth = _sharingAttachmentsOfferAboveUnitLabel.PreferredSize.Width;
            int rightLimit = _sharingAttachmentAutomationGroup.ClientSize.Width - innerPadding;
            int inlineRequiredRight = spinnerLeft + _sharingAttachmentsOfferAboveMbUpDown.Width + ScaleLogical(8) + unitWidth;
            if (inlineRequiredRight > rightLimit)
            {
                spinnerLeft = innerPadding + ScaleLogical(24);
                spinnerTop = _sharingAttachmentsOfferAboveCheckBox.Bottom + ScaleLogical(6);
            }

            _sharingAttachmentsOfferAboveMbUpDown.Location = new Point(spinnerLeft, spinnerTop);
            _sharingAttachmentsOfferAboveUnitLabel.Location = new Point(_sharingAttachmentsOfferAboveMbUpDown.Right + ScaleLogical(8), spinnerTop + ScaleLogical(2));

            int contentBottom = Math.Max(
                _sharingAttachmentsOfferAboveCheckBox.Bottom,
                Math.Max(_sharingAttachmentsOfferAboveMbUpDown.Bottom, _sharingAttachmentsOfferAboveUnitLabel.Bottom));
            if (lockActive)
            {
                contentBottom = Math.Max(contentBottom, _sharingAttachmentLockStepsLabel.Bottom);
            }
            int requiredGroupHeight = Math.Max(ScaleLogical(110), contentBottom + innerPadding);
            if (_sharingAttachmentAutomationGroup.Height != requiredGroupHeight)
            {
                _sharingAttachmentAutomationGroup.Height = requiredGroupHeight;
            }

            _sharingAttachmentsAlwaysCheckBox.Enabled = !effectiveAlwaysLock && !uiBusy;
            _sharingAttachmentsOfferAboveCheckBox.Enabled = !effectiveThresholdLock && !alwaysConnector && !uiBusy;

            bool thresholdInputEnabled = !effectiveThresholdLock && !alwaysConnector && _sharingAttachmentsOfferAboveCheckBox.Checked && !uiBusy;
            _sharingAttachmentsOfferAboveMbUpDown.Enabled = thresholdInputEnabled;
            _sharingAttachmentsOfferAboveUnitLabel.Enabled = !effectiveThresholdLock && !alwaysConnector && !uiBusy;

            _disabledTooltipHints.Apply(
                _sharingAttachmentsAlwaysCheckBox,
                lockActive
                    ? _sharingAttachmentLockHintLabel.Text
                    : (policyLockAlways ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSharingAttachmentsAlways),
                effectiveAlwaysLock,
                _sharingAttachmentLockHintLabel);
            _disabledTooltipHints.Apply(
                _sharingAttachmentsOfferAboveCheckBox,
                lockActive
                    ? _sharingAttachmentLockHintLabel.Text
                    : (policyLockThreshold ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSharingAttachmentsOffer),
                effectiveThresholdLock,
                _sharingAttachmentsOfferAboveUnitLabel,
                _sharingAttachmentLockHintLabel,
                _sharingAttachmentsOfferAboveUnitLabel,
                _sharingAttachmentsOfferAboveMbUpDown);
            _disabledTooltipHints.Apply(
                _sharingAttachmentsOfferAboveMbUpDown,
                effectiveThresholdLock ? Strings.PolicyAdminControlledTooltip : string.Empty,
                false,
                _sharingAttachmentsOfferAboveUnitLabel);
        }

        private void RefreshTalkSystemAddressbookState(bool forceRefresh, string trigger)
        {
            string serverUrl = _serverUrlTextBox.Text.Trim();
            string username = _usernameTextBox.Text.Trim();
            string appPassword = _appPasswordTextBox.Text ?? string.Empty;
            int cacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            var configuration = new TalkServiceConfiguration(serverUrl, username, appPassword);
            var cache = new IfbAddressBookCache(AppDataPaths.EnsureLocalRootDirectory());

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "System address book status check requested from settings (trigger=" + (trigger ?? "n/a") +
                ", forceRefresh=" + forceRefresh + ").");

            var status = cache.GetSystemAddressbookStatus(configuration, cacheHours, forceRefresh);
            bool lockActive = !status.Available;
            string detail = lockActive ? Strings.TalkSystemAddressbookRequiredMessage : string.Empty;
            if (lockActive && !string.IsNullOrWhiteSpace(status.Error))
            {
                DiagnosticsLogger.Log(
                    LogCategories.Talk,
                    "System address book unavailable in settings (trigger=" + (trigger ?? "n/a") + ", error=" + status.Error + ").");
            }

            ApplyTalkSystemAddressbookLockState(lockActive, detail, trigger, status);
        }

        private void ApplyTalkSystemAddressbookLockState(
            bool lockActive,
            string detail,
            string trigger,
            IfbAddressBookCache.SystemAddressbookStatus status)
        {
            _talkAddressbookLockActive = lockActive;
            _talkAddressbookLockDetail = lockActive ? (detail ?? Strings.TalkSystemAddressbookRequiredMessage) : string.Empty;

            bool usersPolicyLocked = IsPolicyLocked("talk", "talk_add_users");
            bool guestsPolicyLocked = IsPolicyLocked("talk", "talk_add_guests");

            _talkDefaultAddUsersCheckBox.Enabled = !lockActive && !usersPolicyLocked && !_isBusy;
            _talkDefaultAddGuestsCheckBox.Enabled = !lockActive && !guestsPolicyLocked && !_isBusy;
            _talkAddressbookWarningPanel.Visible = lockActive;
            _talkAddressbookWarningTextLabel.Text = lockActive ? _talkAddressbookLockDetail : string.Empty;

            _disabledTooltipHints.Apply(
                _talkDefaultAddUsersCheckBox,
                lockActive ? Strings.TooltipAddUsersLocked : (usersPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddUsers),
                lockActive || usersPolicyLocked,
                _talkAddressbookWarningPanel);
            _disabledTooltipHints.Apply(
                _talkDefaultAddGuestsCheckBox,
                lockActive ? Strings.TooltipAddGuestsLocked : (guestsPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddGuests),
                lockActive || guestsPolicyLocked,
                _talkAddressbookWarningPanel);

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "System address book lock state applied in settings (trigger=" + (trigger ?? "n/a") +
                ", locked=" + lockActive +
                ", available=" + (status != null && status.Available) +
                ", count=" + (status != null ? status.Count : 0) +
                ", hasError=" + (status != null && !string.IsNullOrWhiteSpace(status.Error)) + ").");

            if (!_layoutApplying)
            {
                ApplyTalkDefaultsTabLayout();
            }
        }

        private static void SelectComboValue(ComboBox combo, int value, int fallback)
        {            if (combo == null)
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
        {            if (combo == null)
            {
                return fallback;
            }
            string selected = combo.SelectedItem as string;
            int parsed;            if (selected != null && int.TryParse(selected, out parsed))
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
            if (File.Exists(path))
            {
                bool opened = BrowserLauncher.OpenTarget(
                    path,
                    LogCategories.Core,
                    "Failed to open debug log file.");
                if (!opened)
                {
                    MessageBox.Show(
                        Strings.DebugLogOpenErrorMessage,
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return;
            }
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
            {
                bool opened = BrowserLauncher.OpenTarget(
                    directory,
                    LogCategories.Core,
                    "Failed to open debug log directory.");
                if (!opened)
                {
                    MessageBox.Show(
                        Strings.DebugLogOpenErrorMessage,
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return;
            }

            MessageBox.Show(
                Strings.DebugLogMissingMessage,
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void UpdateTlsOptionsState()
        {
            bool useSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
            bool allowCustom = !useSystemDefault && !_isBusy;

            _tlsEnable12CheckBox.Enabled = allowCustom;
            _tlsEnable13CheckBox.Enabled = allowCustom;
        }

        private void OnTlsSelectionChanged(object sender, EventArgs e)
        {
            UpdateTlsOptionsState();
            ApplyTlsRuntimePreview("settings_tls_changed");
        }

        private void ApplyTlsRuntimePreview(string source)
        {
            if (_suppressImmediateTlsApply)
            {
                return;
            }
            if (!_tlsUseSystemDefaultCheckBox.Checked
                && !_tlsEnable12CheckBox.Checked
                && !_tlsEnable13CheckBox.Checked)
            {
                SetStatus(Strings.TransportTlsSelectionRequired, true);
                return;
            }
            try
            {
                TransportSecurityConfigurator.Apply(
                    _tlsUseSystemDefaultCheckBox.Checked,
                    _tlsEnable12CheckBox.Checked,
                    _tlsEnable13CheckBox.Checked,
                    source);
                SetStatus(string.Empty, false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply live TLS runtime preview from settings UI.",
                    ex);
                SetStatus(string.Format(Strings.TransportTlsApplyFailed, ex.Message), true);
            }
        }

        private void OnSettingsFormClosed(object sender, FormClosedEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                return;
            }
            try
            {
                ServicePointManager.SecurityProtocol = _runtimeSecurityProtocolAtOpen;
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Transport security restored after settings dialog cancel/close (securityProtocol="
                    + _runtimeSecurityProtocolAtOpen
                    + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to restore transport security after settings dialog cancel/close.",
                    ex);
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
            _ifbPortUpDown.Visible = showDays;
            _ifbPortLabel.Visible = showDays;
            _ifbPortUpDown.Enabled = showDays && !_isBusy && _ifbEnabledCheckBox.Enabled;
            _ifbPortLabel.Enabled = showDays && !_isBusy && _ifbEnabledCheckBox.Enabled;

            _ifbCacheHoursCombo.Enabled = !_isBusy;
            _ifbCacheHoursLabel.Enabled = !_isBusy;
            _debugLogCheckBox.Enabled = !_isBusy;
            _debugAnonymizeCheckBox.Enabled = !_isBusy;
            _debugOpenLink.Enabled = !_isBusy;
            _tlsUseSystemDefaultCheckBox.Enabled = !_isBusy;

            bool lockShareBase = IsPolicyLocked("share", "share_base_directory");
            bool lockShareName = IsPolicyLocked("share", "share_name_template");
            bool lockSharePermCreate = IsPolicyLocked("share", "share_permission_upload");
            bool lockSharePermWrite = IsPolicyLocked("share", "share_permission_edit");
            bool lockSharePermDelete = IsPolicyLocked("share", "share_permission_delete");
            bool lockSharePassword = IsPolicyLocked("share", "share_set_password");
            bool lockSharePasswordSeparate = IsPolicyLocked("share", "share_send_password_separately");
            bool lockShareExpire = IsPolicyLocked("share", "share_expire_days");
            bool lockShareLang = IsPolicyLocked("share", "language_share_html_block");
            bool lockTalkPassword = IsPolicyLocked("talk", "talk_set_password");
            bool lockTalkLobby = IsPolicyLocked("talk", "talk_lobby_active");
            bool lockTalkSearch = IsPolicyLocked("talk", "talk_show_in_search");
            bool lockTalkRoomType = IsPolicyLocked("talk", "talk_room_type");
            bool lockTalkLang = IsPolicyLocked("talk", "language_talk_description");
            bool lockTalkUsers = IsPolicyLocked("talk", "talk_add_users");
            bool lockTalkGuests = IsPolicyLocked("talk", "talk_add_guests");
            bool separatePasswordAvailable = HasBackendSeatEntitlement();
            string separatePasswordUnavailableTooltip = GetSeparatePasswordUnavailableTooltip();

            _fileLinkBaseTextBox.Enabled = !lockShareBase && !_isBusy;
            _sharingDefaultShareNameTextBox.Enabled = !lockShareName && !_isBusy;
            _sharingDefaultPermCreateCheckBox.Enabled = !lockSharePermCreate && !_isBusy;
            _sharingDefaultPermWriteCheckBox.Enabled = !lockSharePermWrite && !_isBusy;
            _sharingDefaultPermDeleteCheckBox.Enabled = !lockSharePermDelete && !_isBusy;
            _sharingDefaultPasswordCheckBox.Enabled = !lockSharePassword && !_isBusy;

            bool separateEnabled = separatePasswordAvailable
                                   && _sharingDefaultPasswordCheckBox.Checked
                                   && !lockSharePasswordSeparate
                                   && !_isBusy;
            _sharingDefaultPasswordSeparateCheckBox.Enabled = separateEnabled;
            if (!separatePasswordAvailable || !_sharingDefaultPasswordCheckBox.Checked)
            {
                _sharingDefaultPasswordSeparateCheckBox.Checked = false;
            }
            _sharingDefaultExpireDaysUpDown.Enabled = !lockShareExpire && !_isBusy;
            _shareBlockLangCombo.Enabled = !lockShareLang && !_isBusy;

            _talkDefaultPasswordCheckBox.Enabled = !lockTalkPassword && !_isBusy;
            _talkDefaultLobbyCheckBox.Enabled = !lockTalkLobby && !_isBusy;
            _talkDefaultSearchCheckBox.Enabled = !lockTalkSearch && !_isBusy;
            _talkDefaultRoomTypeCombo.Enabled = !lockTalkRoomType && !_isBusy;
            _eventDescriptionLangCombo.Enabled = !lockTalkLang && !_isBusy;
            _talkDefaultAddUsersCheckBox.Enabled = !_talkAddressbookLockActive && !lockTalkUsers && !_isBusy;
            _talkDefaultAddGuestsCheckBox.Enabled = !_talkAddressbookLockActive && !lockTalkGuests && !_isBusy;

            SetTooltipWithFallback(_fileLinkBaseTextBox, lockShareBase ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareBase, _fileLinkBaseHintLabel);
            SetTooltipWithFallback(_sharingDefaultShareNameTextBox, lockShareName ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareName, _sharingDefaultShareNameLabel);
            SetTooltipWithFallback(_sharingDefaultPermCreateCheckBox, lockSharePermCreate ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermCreate, _sharingDefaultPermissionsLabel);
            SetTooltipWithFallback(_sharingDefaultPermWriteCheckBox, lockSharePermWrite ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermWrite, _sharingDefaultPermissionsLabel);
            SetTooltipWithFallback(_sharingDefaultPermDeleteCheckBox, lockSharePermDelete ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermDelete, _sharingDefaultPermissionsLabel);
            SetTooltipWithFallback(_sharingDefaultPasswordCheckBox, lockSharePassword ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePassword);
            SetTooltipWithFallback(
                _sharingDefaultPasswordSeparateCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockSharePasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockSharePasswordSeparate);
            SetTooltipWithFallback(_sharingDefaultExpireDaysUpDown, lockShareExpire ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareExpire, _sharingDefaultExpireDaysLabel);
            SetTooltipWithFallback(_shareBlockLangCombo, lockShareLang ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareLang, _shareBlockLangLabel);
            SetTooltipWithFallback(_talkDefaultPasswordCheckBox, lockTalkPassword ? Strings.PolicyAdminControlledTooltip : string.Empty, lockTalkPassword);
            SetTooltipWithFallback(_talkDefaultLobbyCheckBox, lockTalkLobby ? Strings.PolicyAdminControlledTooltip : Strings.TooltipLobby, lockTalkLobby);
            SetTooltipWithFallback(_talkDefaultSearchCheckBox, lockTalkSearch ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSearchVisible, lockTalkSearch);
            TalkRoomTypeOption selectedTalkRoomTypeOption = _talkDefaultRoomTypeCombo.SelectedItem as TalkRoomTypeOption;
            bool standardTalkRoomTypeSelected =
                selectedTalkRoomTypeOption != null &&
                selectedTalkRoomTypeOption.Value == TalkRoomType.StandardRoom;
            SetTooltipWithFallback(
                _talkDefaultRoomTypeCombo,
                lockTalkRoomType
                    ? Strings.PolicyAdminControlledTooltip
                    : (standardTalkRoomTypeSelected
                        ? Strings.TooltipRoomTypeStandard
                        : Strings.TooltipRoomTypeEvent),
                lockTalkRoomType,
                _talkDefaultRoomTypeLabel);
            SetTooltipWithFallback(_eventDescriptionLangCombo, lockTalkLang ? Strings.PolicyAdminControlledTooltip : string.Empty, lockTalkLang, _eventDescriptionLangLabel);
            UpdateSharingAttachmentOptionsState();
            UpdateTlsOptionsState();
        }

        private void SetTooltipWithFallback(Control primary, string text, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, fallbackTargets);
        }

        private void SetTooltipWithFallback(Control primary, string text, bool showHint, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, showHint, fallbackTargets);
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
            ApplyResponsiveLayout(false);
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
            _fileLinkBaseHintLabel.MaximumSize = new Size(420, 0);
            _fileLinkBaseHintLabel.AutoSize = true;
            _fileLinkBaseHintLabel.ForeColor = Color.DimGray;
            _fileLinkTab.Controls.Add(_fileLinkBaseHintLabel);

            _sharingDefaultsGroup.Text = Strings.SharingDefaultsHeading;
            _sharingDefaultsGroup.Location = new Point(18, 134);
            _sharingDefaultsGroup.Size = new Size(500, 394);
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

            _sharingDefaultPasswordSeparateCheckBox.Text = Strings.SharingDefaultPasswordSeparateLabel;
            _sharingDefaultPasswordSeparateCheckBox.Location = new Point(260, 138);
            _sharingDefaultPasswordSeparateCheckBox.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPasswordSeparateCheckBox);

            _sharingDefaultExpireDaysLabel.Text = Strings.SharingDefaultExpireDaysLabel;
            _sharingDefaultExpireDaysLabel.Location = new Point(260, 170);
            _sharingDefaultExpireDaysLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysLabel);

            _sharingDefaultExpireDaysUpDown.Minimum = 1;
            _sharingDefaultExpireDaysUpDown.Maximum = 3650;
            _sharingDefaultExpireDaysUpDown.Location = new Point(260, 194);
            _sharingDefaultExpireDaysUpDown.Width = 90;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysUpDown);

            _sharingAttachmentAutomationGroup.Text = Strings.SharingAttachmentAutomationHeading;
            _sharingAttachmentAutomationGroup.Location = new Point(12, 230);
            _sharingAttachmentAutomationGroup.Size = new Size(472, 150);
            _sharingDefaultsGroup.Controls.Add(_sharingAttachmentAutomationGroup);

            _sharingAttachmentLockHintLabel.AutoSize = true;
            _sharingAttachmentLockHintLabel.Location = new Point(12, 20);
            _sharingAttachmentLockHintLabel.MaximumSize = new Size(448, 0);
            _sharingAttachmentLockHintLabel.ForeColor = Color.Maroon;
            _sharingAttachmentLockHintLabel.Visible = false;
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentLockHintLabel);

            _sharingAttachmentLockStepsLabel.AutoSize = true;
            _sharingAttachmentLockStepsLabel.Location = new Point(12, 52);
            _sharingAttachmentLockStepsLabel.MaximumSize = new Size(448, 0);
            _sharingAttachmentLockStepsLabel.ForeColor = Color.DimGray;
            _sharingAttachmentLockStepsLabel.Visible = false;
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentLockStepsLabel);

            _sharingAttachmentsAlwaysCheckBox.Text = Strings.SharingAttachmentsAlwaysConnectorLabel;
            _sharingAttachmentsAlwaysCheckBox.Location = new Point(12, 24);
            _sharingAttachmentsAlwaysCheckBox.AutoSize = true;
            _sharingAttachmentsAlwaysCheckBox.CheckedChanged += (s, e) => UpdateSharingAttachmentOptionsState();
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentsAlwaysCheckBox);

            _sharingAttachmentsOfferAboveCheckBox.Text = Strings.SharingAttachmentsOfferAboveLabel;
            _sharingAttachmentsOfferAboveCheckBox.Location = new Point(12, 52);
            _sharingAttachmentsOfferAboveCheckBox.AutoSize = true;
            _sharingAttachmentsOfferAboveCheckBox.CheckedChanged += (s, e) => UpdateSharingAttachmentOptionsState();
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentsOfferAboveCheckBox);

            _sharingAttachmentsOfferAboveMbUpDown.Minimum = 1;
            _sharingAttachmentsOfferAboveMbUpDown.Maximum = 10240;
            _sharingAttachmentsOfferAboveMbUpDown.Location = new Point(236, 50);
            _sharingAttachmentsOfferAboveMbUpDown.Width = 72;
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentsOfferAboveMbUpDown);

            _sharingAttachmentsOfferAboveUnitLabel.Text = Strings.SharingAttachmentsOfferAboveUnit;
            _sharingAttachmentsOfferAboveUnitLabel.Location = new Point(316, 52);
            _sharingAttachmentsOfferAboveUnitLabel.AutoSize = true;
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentsOfferAboveUnitLabel);

            _toolTip.SetToolTip(_sharingDefaultPermissionsLabel, Strings.TooltipSharingPermissions);
            _toolTip.SetToolTip(_sharingDefaultPasswordSeparateCheckBox, string.Empty);
            _toolTip.SetToolTip(_sharingAttachmentsAlwaysCheckBox, Strings.TooltipSharingAttachmentsAlways);
            _toolTip.SetToolTip(_sharingAttachmentsOfferAboveCheckBox, Strings.TooltipSharingAttachmentsOffer);
        }

        private void UpdateKnownServerVersion(string candidate)
        {
            Version parsed;
            if (NextcloudVersionHelper.TryParse(candidate, out parsed))
            {
                _lastKnownServerVersion = parsed.ToString();
            }
        }

        private int ScaleLogical(int value)
        {
            int dpi = DeviceDpi > 0 ? DeviceDpi : 96;
            return (int)Math.Round(value * (dpi / 96f));
        }
    }
}

