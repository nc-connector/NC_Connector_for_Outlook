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
        private const string NcConnectorHomepageUrl = "https://nc-connector.de";
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
            _generalTab.Resize += (s, e) => ApplyResponsiveLayout(false);
            _fileLinkTab.Resize += (s, e) => ApplyResponsiveLayout(false);
            _talkTab.Resize += (s, e) => ApplyResponsiveLayout(false);

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

            int comboLeft = _ifbDaysLabel.Right + labelToComboGap;
            int comboHeight = Math.Max(_ifbDaysCombo.Height, _ifbDaysCombo.PreferredHeight + ScaleLogical(2));
            _ifbDaysCombo.SetBounds(comboLeft, daysLabelTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbDaysCombo.Width), comboHeight);
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
        }

        private void ApplyDebugTabLayout()
        {
            int rightMargin = ScaleLogical(24);
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
            _shareBlockLangCombo.IntegralHeight = false;
            _shareBlockLangCombo.Location = new Point(260, langTop - 2);
            _shareBlockLangCombo.Width = 240;
            PopulateLanguageOverrideCombo(_shareBlockLangCombo);
            _advancedTab.Controls.Add(_shareBlockLangCombo);

            _eventDescriptionLangLabel.Text = Strings.AdvancedEventDescriptionLangLabel;
            _eventDescriptionLangLabel.Location = new Point(24, langTop + 32);
            _eventDescriptionLangLabel.AutoSize = true;
            _advancedTab.Controls.Add(_eventDescriptionLangLabel);

            _eventDescriptionLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _eventDescriptionLangCombo.IntegralHeight = false;
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
            _talkAddressbookWarningLinkLabel.LinkClicked += (s, e) => OpenBrowser(Strings.TalkSystemAddressbookAdminGuideUrl);
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

            _debugPathLabel.AutoSize = true;
            _debugPathLabel.Location = new Point(24, 60);
            _debugPathLabel.MaximumSize = new Size(420, 0);
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
            _aboutHomepageLink.LinkClicked += (sender, args) => OpenBrowser(NcConnectorHomepageUrl);
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
            _aboutMoreInfoLink.LinkClicked += (sender, args) => OpenBrowser(NcConnectorHomepageUrl);
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
            _aboutSupportLink.LinkClicked += (sender, args) => OpenBrowser("https://www.paypal.com/donate/?hosted_button_id=FTZWNRNKVKUN6");
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
            _sharingDefaultPasswordSeparateCheckBox.Checked =
                AddinSettings.SeparatePasswordFeatureEnabled
                && Result.SharingDefaultPasswordSeparateEnabled;
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
            SelectLanguageChoice(_shareBlockLangCombo, Result.ShareBlockLang);
            SelectLanguageChoice(_eventDescriptionLangCombo, Result.EventDescriptionLang);
            UpdateDebugPathLabel();
            UpdateAboutTab();
            RefreshSharingAttachmentLockState();
            UpdateSharingAttachmentOptionsState();
            ApplySharingPasswordSeparateAvailability();
            RefreshTalkSystemAddressbookState(true, "settings_open");
        }

        private void OnSaveButtonClick(object sender, EventArgs e)
        {
            RefreshTalkSystemAddressbookState(true, "settings_save");

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
            Result.SharingDefaultPasswordSeparateEnabled =
                AddinSettings.SeparatePasswordFeatureEnabled
                && _sharingDefaultPasswordCheckBox.Checked
                && _sharingDefaultPasswordSeparateCheckBox.Checked;
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
            bool uiBusy = _isBusy;

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

            _sharingAttachmentsAlwaysCheckBox.Enabled = !lockActive && !uiBusy;
            _sharingAttachmentsOfferAboveCheckBox.Enabled = !lockActive && !alwaysConnector && !uiBusy;

            bool thresholdInputEnabled = !lockActive && !alwaysConnector && _sharingAttachmentsOfferAboveCheckBox.Checked && !uiBusy;
            _sharingAttachmentsOfferAboveMbUpDown.Enabled = thresholdInputEnabled;
            _sharingAttachmentsOfferAboveUnitLabel.Enabled = !lockActive && !alwaysConnector && !uiBusy;
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

            _talkDefaultAddUsersCheckBox.Enabled = !lockActive && !_isBusy;
            _talkDefaultAddGuestsCheckBox.Enabled = !lockActive && !_isBusy;
            _talkAddressbookWarningPanel.Visible = lockActive;
            _talkAddressbookWarningTextLabel.Text = lockActive ? _talkAddressbookLockDetail : string.Empty;

            _toolTip.SetToolTip(_talkDefaultAddUsersCheckBox, lockActive ? Strings.TooltipAddUsersLocked : Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_talkDefaultAddGuestsCheckBox, lockActive ? Strings.TooltipAddGuestsLocked : Strings.TooltipAddGuests);

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
            _talkDefaultAddUsersCheckBox.Enabled = !_talkAddressbookLockActive && !_isBusy;
            _talkDefaultAddGuestsCheckBox.Enabled = !_talkAddressbookLockActive && !_isBusy;
            UpdateSharingAttachmentOptionsState();
            ApplySharingPasswordSeparateAvailability();
        }

        private void ApplySharingPasswordSeparateAvailability()
        {
            bool featureEnabled = AddinSettings.SeparatePasswordFeatureEnabled;
            bool interactive = featureEnabled && _sharingDefaultPasswordCheckBox.Checked && !_isBusy;

            _sharingDefaultPasswordSeparateCheckBox.AutoCheck = interactive;
            _sharingDefaultPasswordSeparateCheckBox.TabStop = interactive;
            _sharingDefaultPasswordSeparateCheckBox.ForeColor = interactive ? _themePalette.Text : _themePalette.DisabledText;

            if (!interactive)
            {
                _sharingDefaultPasswordSeparateCheckBox.Checked = false;
            }
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
            _sharingDefaultPasswordCheckBox.CheckedChanged += (s, e) => ApplySharingPasswordSeparateAvailability();
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
            _toolTip.SetToolTip(_sharingDefaultPasswordSeparateCheckBox, Strings.TooltipSharingPasswordSeparate);
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
