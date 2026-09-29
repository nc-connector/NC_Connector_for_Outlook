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
    internal sealed partial class SettingsForm : ScaledForm
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
        private readonly TabPage _signatureTab = new TabPage(Strings.TabSignature);
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
        private readonly Label _sharingPasswordDeliveryModeLabel = new Label();
        private readonly ComboBox _sharingPasswordDeliveryModeCombo = new ComboBox();
        private readonly Label _sharingDefaultExpireDaysLabel = new Label();
        private readonly NumericUpDown _sharingDefaultExpireDaysUpDown = new NumericUpDown();
        private readonly GroupBox _sharingAttachmentAutomationGroup = new GroupBox();
        private readonly Label _sharingAttachmentLockHintLabel = new Label();
        private readonly Label _sharingAttachmentLockStepsLabel = new Label();
        private readonly CheckBox _sharingAttachmentsAlwaysCheckBox = new CheckBox();
        private readonly CheckBox _sharingAttachmentsOfferAboveCheckBox = new CheckBox();
        private readonly NumericUpDown _sharingAttachmentsOfferAboveMbUpDown = new NumericUpDown();
        private readonly Label _sharingAttachmentsOfferAboveUnitLabel = new Label();
        private readonly Label _sharingAttachmentLinkTargetLabel = new Label();
        private readonly ComboBox _sharingAttachmentLinkTargetCombo = new ComboBox();
        private readonly GroupBox _talkDefaultsGroup = new GroupBox();
        private readonly Label _talkDefaultRoomTypeLabel = new Label();
        private readonly ComboBox _talkDefaultRoomTypeCombo = new ComboBox();
        private readonly CheckBox _talkDefaultPasswordCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddUsersCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultAddGuestsCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultLobbyCheckBox = new CheckBox();
        private readonly CheckBox _talkDefaultSearchCheckBox = new CheckBox();
        private readonly CheckBox _talkDeleteRoomOnEventDeleteCheckBox = new CheckBox();
        private readonly Panel _talkAddressbookWarningPanel = new Panel();
        private readonly Label _talkAddressbookWarningTitleLabel = new Label();
        private readonly Label _talkAddressbookWarningTextLabel = new Label();
        private readonly LinkLabel _talkAddressbookWarningLinkLabel = new LinkLabel();
        private readonly Label _shareBlockLangLabel = new Label();
        private readonly ComboBox _shareBlockLangCombo = new ComboBox();
        private readonly Label _eventDescriptionLangLabel = new Label();
        private readonly ComboBox _eventDescriptionLangCombo = new ComboBox();
        private readonly GroupBox _signatureDefaultsGroup = new GroupBox();
        private readonly Label _signatureIdentityHintLabel = new Label();
        private readonly CheckBox _emailSignatureOnComposeCheckBox = new CheckBox();
        private readonly CheckBox _emailSignatureOnReplyCheckBox = new CheckBox();
        private readonly CheckBox _emailSignatureOnForwardCheckBox = new CheckBox();
        private readonly GroupBox _tlsSettingsGroup = new GroupBox();
        private readonly GroupBox _updateSettingsGroup = new GroupBox();
        private readonly CheckBox _updateNotifyCheckBox = new CheckBox();
        private readonly Label _updateInstalledVersionLabel = new Label();
        private readonly Label _updateLatestVersionLabel = new Label();
        private readonly Label _updateLastCheckedLabel = new Label();
        private readonly Button _updateCheckButton = new Button();
        private readonly LinkLabel _updateDownloadLink = new LinkLabel();
        private readonly Label _updateChangelogLabel = new Label();
        private readonly TextBox _updateChangelogTextBox = new TextBox();
        private readonly CheckBox _tlsUseSystemDefaultCheckBox = new CheckBox();
        private readonly CheckBox _tlsEnable12CheckBox = new CheckBox();
        private readonly CheckBox _tlsEnable13CheckBox = new CheckBox();
        private readonly Label _tlsHintLabel = new Label();

        private readonly Label _statusLabel = new Label();
        private readonly Button _saveButton = new Button();
        private readonly Button _cancelButton = new Button();
        private bool _isBusy;
        private bool _authenticationRequired;
        private bool _authenticationRejected;
        private bool _connectionSetupPending;
        private AddinSettings _result;
        private bool _applyingPolicyDefaults;
        private bool _initialIfbEnabled;
        private bool _ifbDefaultApplied;
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private bool _sharingAttachmentLockActive;
        private int _sharingAttachmentLockThresholdMb = 5;
        private bool _talkAddressbookLockActive;
        private string _talkAddressbookLockDetail = string.Empty;
        private int _addressbookRefreshGeneration;
        private bool _layoutApplying;
        private bool _suppressImmediateTlsApply;
        private readonly SecurityProtocolType _runtimeSecurityProtocolAtOpen;
        private readonly IfbAddressBookCache _addressBookCache;
        private readonly IfbAddressBookCache.SystemAddressbookStatus _initialAddressbookStatus;
        private BackendPolicyStatus _backendPolicyStatus;
        private string _updateOpenUrl = string.Empty;

        internal AddinSettings Result
        {
            get { return _result; }
            private set { _result = value; }
        }

        internal void BeginAuthentication(bool authenticationRejected)
        {
            _authenticationRequired = true;
            _authenticationRejected = authenticationRejected;
            _connectionSetupPending = true;
            ApplyBackendPolicyStatus("authentication_open");
        }

        internal SettingsForm(
            AddinSettings settings,
            Outlook.Application outlookApplication,
            BackendPolicyStatus initialPolicyStatus,
            IfbAddressBookCache addressBookCache,
            IfbAddressBookCache.SystemAddressbookStatus initialAddressbookStatus)
        {
            _outlookApplication = outlookApplication;
            _backendPolicyStatus = initialPolicyStatus;
            _addressBookCache = addressBookCache;
            _initialAddressbookStatus = initialAddressbookStatus;
            _runtimeSecurityProtocolAtOpen = ServicePointManager.SecurityProtocol;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            Text = Strings.SettingsFormTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(820, 720);
            MinimumSize = new Size(ScaleLogical(800), ScaleLogical(700));
            Icon = BrandingAssets.GetAppIcon(32);

            BrandedHeader.AttachToParent(_headerPanel, Controls, HeaderHeight);
            InitializeComponents();
            ApplySettings(settings);
            if (!Result.ShowMainRibbonTab)
            {
                // Keep the existing authentication controls, not the normal settings menu.
                _tabControl.TabPages.Clear();
                _tabControl.TabPages.Add(_generalTab);
                Disposed += (sender, args) =>
                {
                    foreach (TabPage page in new[] { _fileLinkTab, _talkTab, _signatureTab,
                        _ifbTab, _advancedTab, _debugTab, _aboutTab })
                    {
                        page.Dispose();
                    }
                };
            }
            TrackLocalPolicyChoices();
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
            _tabControl.TabPages.Add(_signatureTab);
            _tabControl.TabPages.Add(_ifbTab);
            _tabControl.TabPages.Add(_advancedTab);
            _tabControl.TabPages.Add(_debugTab);
            _tabControl.TabPages.Add(_aboutTab);
            _tabControl.SelectedIndexChanged += OnSelectedTabChanged;
            Controls.Add(_tabControl);
            InitializePolicyWarningPanel();

            InitializeGeneralTab();
            InitializeTalkTab();
            InitializeSignatureTab();
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
            _saveButton.DialogResult = DialogResult.None;
            _saveButton.Click += OnSaveButtonClick;
            Controls.Add(_saveButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.Size = new Size(120, 32);
            _cancelButton.Location = new Point(ClientSize.Width - 132, ClientSize.Height - 44);
            _cancelButton.Anchor = AnchorStyles.None;
            _cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(_cancelButton);

            Resize += (s, e) => ApplyResponsiveLayout(false);
            AttachResponsiveResizeHandlers(_generalTab, _fileLinkTab, _talkTab, _signatureTab);

            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
        }

        private void AttachResponsiveResizeHandlers(params Control[] controls)
        {
            if (controls == null)
            {
                return;
            }
            for (int i = 0; i < controls.Length; i++)
            {
                Control control = controls[i];
                if (control == null)
                {
                    continue;
                }

                control.Resize += (s, e) => ApplyResponsiveLayout(false);
            }
        }

        private void InitializePolicyWarningPanel()
        {
            WarningPanelUiHelper.Initialize(
                _policyWarningPanel,
                _policyWarningTitleLabel,
                _policyWarningTextLabel,
                _policyWarningLinkLabel,
                "\u26a0 " + Strings.PolicyWarningTitle,
                string.Empty,
                Strings.PolicyWarningAdminLinkLabel);
            Controls.Add(_policyWarningPanel);
            _policyWarningLinkLabel.LinkClicked += (s, e) =>
                PolicyUiHelper.OpenLicenseAdministration(_policyWarningLinkLabel, LogCategories.Core);
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
                int policyWarningHeight = WarningPanelUiHelper.Layout(
                    _policyWarningPanel,
                    _policyWarningTitleLabel,
                    _policyWarningTextLabel,
                    _policyWarningLinkLabel,
                    outerPadding,
                    tabTop,
                    Math.Max(ScaleLogical(420), ClientSize.Width - (outerPadding * 2)),
                    ScaleLogical(8),
                    ScaleLogical(220),
                    ScaleLogical(4),
                    ScaleLogical(6));
                if (policyWarningHeight > 0)
                {
                    tabTop = _policyWarningPanel.Bottom + ScaleLogical(8);
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
                ApplySignatureTabLayout();
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

        private void ApplySettings(AddinSettings settings)
        {
            Result = settings.Clone();
            _suppressImmediateTlsApply = true;
            try
            {
                _initialIfbEnabled = Result.IfbEnabled;
                _ifbDefaultApplied = _initialIfbEnabled || Result.IfbUserDecisionRecorded;
                _serverUrlTextBox.Text = Result.ManagedNextcloudUrlLocked ? Result.ManagedNextcloudUrl : Result.ServerUrl;
                _usernameTextBox.Text = Result.Username;
                _appPasswordTextBox.Text = Result.AppPassword;
                _manualRadio.Checked = Result.AuthMode == AuthenticationMode.Manual;
                _loginFlowRadio.Checked = !_manualRadio.Checked;
                _connectionSetupPending = !new TalkServiceConfiguration(
                    _serverUrlTextBox.Text, _usernameTextBox.Text, _appPasswordTextBox.Text).IsComplete();
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
                _updateNotifyCheckBox.Checked = Result.UpdateNotifyEnabled;
                UpdateDebugPathLabel();
                UpdateAboutTab();
                UpdateUpdateCheckSection();
                RefreshSharingAttachmentLockState();
                UpdateSharingAttachmentOptionsState();
                UpdateTlsOptionsState();
                ApplyBackendPolicyStatus("settings_init");
                ApplyInitialTalkSystemAddressbookState();
            }
            finally
            {
                _suppressImmediateTlsApply = false;
            }
        }

        private void TrackLocalPolicyChoices()
        {
            TrackLocalPolicyChoice(_fileLinkBaseTextBox, "share", "share_base_directory", () => Result.FileLinkBasePath = _fileLinkBaseTextBox.Text.Trim());
            TrackLocalPolicyChoice(_sharingDefaultShareNameTextBox, "share", "share_name_template", () => Result.SharingDefaultShareName = _sharingDefaultShareNameTextBox.Text.Trim());
            TrackLocalPolicyChoice(_sharingDefaultPermCreateCheckBox, "share", "share_permission_upload", () => Result.SharingDefaultPermCreate = _sharingDefaultPermCreateCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingDefaultPermWriteCheckBox, "share", "share_permission_edit", () => Result.SharingDefaultPermWrite = _sharingDefaultPermWriteCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingDefaultPermDeleteCheckBox, "share", "share_permission_delete", () => Result.SharingDefaultPermDelete = _sharingDefaultPermDeleteCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingDefaultPasswordCheckBox, "share", "share_set_password", () => Result.SharingDefaultPasswordEnabled = _sharingDefaultPasswordCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingDefaultPasswordSeparateCheckBox, "share", "share_send_password_separately", () => Result.SharingDefaultPasswordSeparateEnabled = _sharingDefaultPasswordSeparateCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingPasswordDeliveryModeCombo, "share", "share_send_password_mode", () => Result.SharingDefaultPasswordDeliveryMode = SharePasswordDeliveryModeComboHelper.GetSelected(_sharingPasswordDeliveryModeCombo));
            TrackLocalPolicyChoice(_sharingDefaultExpireDaysUpDown, "share", "share_expire_days", () => Result.SharingDefaultExpireDays = (int)_sharingDefaultExpireDaysUpDown.Value);
            TrackLocalPolicyChoice(_sharingAttachmentsAlwaysCheckBox, "share", "attachments_always_via_ncconnector", () => Result.SharingAttachmentsAlwaysConnector = _sharingAttachmentsAlwaysCheckBox.Checked);
            TrackLocalPolicyChoice(_sharingAttachmentsOfferAboveCheckBox, "share", "attachments_min_size_mb", () => RecordLocalAttachmentThreshold());
            TrackLocalPolicyChoice(_sharingAttachmentsOfferAboveMbUpDown, "share", "attachments_min_size_mb", () => RecordLocalAttachmentThreshold());
            TrackLocalPolicyChoice(_sharingAttachmentLinkTargetCombo, "share", "attachment_link_target", () => Result.SharingAttachmentLinkTarget = GetSelectedAttachmentLinkTarget());
            TrackLocalPolicyChoice(_shareBlockLangCombo, "share", "language_share_html_block", () => Result.ShareBlockLang = GetSelectedLanguageChoice(_shareBlockLangCombo));
            TrackLocalPolicyChoice(_talkDefaultPasswordCheckBox, "talk", "talk_set_password", () => Result.TalkDefaultPasswordEnabled = _talkDefaultPasswordCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDefaultAddUsersCheckBox, "talk", "talk_add_users", () => Result.TalkDefaultAddUsers = _talkDefaultAddUsersCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDefaultAddGuestsCheckBox, "talk", "talk_add_guests", () => Result.TalkDefaultAddGuests = _talkDefaultAddGuestsCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDefaultLobbyCheckBox, "talk", "talk_lobby_active", () => Result.TalkDefaultLobbyEnabled = _talkDefaultLobbyCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDefaultSearchCheckBox, "talk", "talk_show_in_search", () => Result.TalkDefaultSearchVisible = _talkDefaultSearchCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDeleteRoomOnEventDeleteCheckBox, "talk", "talk_delete_room_on_event_delete", () => Result.TalkDeleteRoomOnEventDelete = _talkDeleteRoomOnEventDeleteCheckBox.Checked);
            TrackLocalPolicyChoice(_talkDefaultRoomTypeCombo, "talk", "talk_room_type", () => Result.TalkDefaultRoomType = TalkRoomTypeComboHelper.GetSelected(_talkDefaultRoomTypeCombo, TalkRoomType.EventConversation));
            TrackLocalPolicyChoice(_eventDescriptionLangCombo, "talk", "language_talk_description", () => Result.EventDescriptionLang = GetSelectedLanguageChoice(_eventDescriptionLangCombo));
            TrackLocalPolicyChoice(_emailSignatureOnComposeCheckBox, "email_signature", "email_signature_on_compose", () => Result.EmailSignatureOnCompose = _emailSignatureOnComposeCheckBox.Checked);
            TrackLocalPolicyChoice(_emailSignatureOnReplyCheckBox, "email_signature", "email_signature_on_reply", () => Result.EmailSignatureOnReply = _emailSignatureOnReplyCheckBox.Checked);
            TrackLocalPolicyChoice(_emailSignatureOnForwardCheckBox, "email_signature", "email_signature_on_forward", () => Result.EmailSignatureOnForward = _emailSignatureOnForwardCheckBox.Checked);
        }

        private void RecordLocalAttachmentThreshold()
        {
            Result.SharingAttachmentsOfferAboveEnabled = _sharingAttachmentsOfferAboveCheckBox.Checked;
            Result.SharingAttachmentsOfferAboveMb = (int)_sharingAttachmentsOfferAboveMbUpDown.Value;
        }

        private void TrackLocalPolicyChoice(Control control, string domain, string key, Action recordChoice)
        {
            EventHandler changed = (sender, args) =>
            {
                if (!_applyingPolicyDefaults && !_isBusy && control.Enabled && !IsPolicyLocked(domain, key))
                {
                    recordChoice();
                }
            };
            CheckBox checkBox = control as CheckBox;
            TextBox textBox = control as TextBox;
            ComboBox comboBox = control as ComboBox;
            NumericUpDown numeric = control as NumericUpDown;
            if (checkBox != null) checkBox.CheckedChanged += changed;
            else if (textBox != null) textBox.TextChanged += changed;
            else if (comboBox != null) comboBox.SelectionChangeCommitted += changed;
            else if (numeric != null) numeric.ValueChanged += changed;
        }

        private async void OnSaveButtonClick(object sender, EventArgs e)
        {
            try
            {
                await SaveSettingsAsync();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Settings form save failed.",
                    ex);
                if (!IsDisposed && !Disposing)
                {
                    SetBusy(false);
                    SetStatus(Strings.SettingsSaveFailed, true);
                }
            }
        }

        private async Task SaveSettingsAsync()
        {
            if (_isBusy)
            {
                return;
            }
            if (!Result.IsManagedTransportTlsValid)
            {
                SetStatus(Strings.ManagedTlsPolicyInvalid, true);
                return;
            }
            string requestedServerUrl = Result.ManagedNextcloudUrlLocked
                ? Result.ManagedNextcloudUrl
                : _serverUrlTextBox.Text.Trim();
            string normalizedServerUrl = string.Empty;
            if (!string.IsNullOrWhiteSpace(requestedServerUrl)
                && !NextcloudUriValidator.TryNormalizeBaseUrl(requestedServerUrl, out normalizedServerUrl))
            {
                SetStatus(Strings.StatusInvalidServerUrl, true);
                return;
            }

            _serverUrlTextBox.Text = normalizedServerUrl;
            var configuration = new TalkServiceConfiguration(
                normalizedServerUrl,
                _usernameTextBox.Text.Trim(),
                _appPasswordTextBox.Text ?? string.Empty);
            if ((_authenticationRequired || !Result.ShowMainRibbonTab) && !configuration.IsComplete())
            {
                ApplyBackendPolicyStatus("authentication_incomplete");
                _tabControl.SelectedTab = _generalTab;
                return;
            }
            if (_authenticationRequired && _connectionSetupPending
                && !await TestConnectionAsync())
            {
                return;
            }
            if (IsDisposed || Disposing)
            {
                return;
            }
            if (!await RefreshSettingsServerStateAsync(
                    configuration,
                    true,
                    "settings_save"))
            {
                return;
            }
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

            Result.ServerUrl = normalizedServerUrl;
            Result.Username = _usernameTextBox.Text.Trim();
            Result.AppPassword = _appPasswordTextBox.Text;
            Result.AuthMode = _loginFlowRadio.Checked ? AuthenticationMode.LoginFlow : AuthenticationMode.Manual;
            if (!Result.ShowMainRibbonTab)
            {
                DialogResult = DialogResult.OK;
                Close();
                return;
            }
            Result.IfbEnabled = _ifbEnabledCheckBox.Checked;
            Result.IfbUserDecisionRecorded = _ifbDefaultApplied;
            Result.IfbDays = ParseComboValue(_ifbDaysCombo, 30);
            Result.IfbPort = AddinSettings.NormalizeIfbPort((int)_ifbPortUpDown.Value);
            Result.IfbCacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            Result.DebugLoggingEnabled = _debugLogCheckBox.Checked;
            Result.LogAnonymizationEnabled = _debugAnonymizeCheckBox.Checked;
            if (!Result.HasManagedTransportTls)
            {
                Result.TransportTlsUseSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
                Result.TransportTlsEnable12 = _tlsEnable12CheckBox.Checked;
                Result.TransportTlsEnable13 = _tlsEnable13CheckBox.Checked;
            }
            Result.UpdateNotifyEnabled = _updateNotifyCheckBox.Checked;
            DialogResult = DialogResult.OK;
            Close();
        }

        // WinForms event handlers must stay async void; keep awaited flow inside this method-level try/catch.
        private async void OnSelectedTabChanged(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == _fileLinkTab)
            {
                RefreshSharingAttachmentLockState();
                UpdateSharingAttachmentOptionsState();
                return;
            }
            if (_tabControl.SelectedTab == _talkTab && !_isBusy)
            {
                var configuration = new TalkServiceConfiguration(
                    _serverUrlTextBox.Text.Trim(),
                    _usernameTextBox.Text.Trim(),
                    _appPasswordTextBox.Text ?? string.Empty);
                try
                {
                    await RefreshTalkSystemAddressbookStateAsync(
                        configuration,
                        false,
                        "settings_tab_talk");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "System address book refresh failed in settings.",
                        ex);
                }
            }
        }

        private bool IsPolicyLocked(string domain, string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked(domain, key);
        }

        private async Task<bool> RefreshSettingsServerStateAsync(
            TalkServiceConfiguration configuration,
            bool forceAddressbookRefresh,
            string trigger)
        {
            SetBusy(true);
            int addressbookGeneration = ++_addressbookRefreshGeneration;
            int cacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            try
            {
                Task<BackendPolicyStatus> policyTask = Task.Run(
                    () => new BackendPolicyService(configuration).FetchStatus());
                Task<IfbAddressBookCache.SystemAddressbookStatus> addressbookTask =
                    _addressBookCache != null
                        ? Task.Run(
                            () => _addressBookCache.GetSystemAddressbookStatus(
                                configuration,
                                cacheHours,
                                forceAddressbookRefresh))
                        : Task.FromResult(_initialAddressbookStatus);

                await Task.WhenAll(policyTask, addressbookTask);
                BackendPolicyStatus policyStatus = await policyTask;
                IfbAddressBookCache.SystemAddressbookStatus addressbookStatus =
                    await addressbookTask;
                if (IsDisposed || Disposing)
                {
                    return false;
                }

                _backendPolicyStatus = policyStatus;
                if (configuration != null && configuration.IsComplete())
                {
                    _connectionSetupPending = false;
                }
                ApplyBackendPolicyStatus(trigger);
                if (addressbookGeneration == _addressbookRefreshGeneration)
                {
                    ApplyTalkSystemAddressbookStatus(
                        addressbookStatus,
                        trigger);
                }

                if (configuration != null
                    && configuration.IsComplete()
                    && (policyStatus == null || !policyStatus.FetchSucceeded))
                {
                    SetStatus(
                        string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.StatusTestFailure,
                            Strings.StatusTestFailureUnknown),
                        true);
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Settings server state refresh failed.",
                    ex);
                if (!IsDisposed && !Disposing)
                {
                    SetStatus(
                        string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.StatusTestFailure,
                            ex.Message ?? Strings.StatusTestFailureUnknown),
                        true);
                }
                return true;
            }
            finally
            {
                if (!IsDisposed && !Disposing)
                {
                    SetBusy(false);
                }
            }
        }

        private void ApplyBackendPolicyStatus(string trigger)
        {
            _policyWarningTitleLabel.Text = "\u26a0 " + Strings.PolicyWarningTitle;
            bool warningVisible = PolicyUiHelper.ApplyPolicyWarningState(
                _backendPolicyStatus,
                _policyWarningPanel,
                _policyWarningTextLabel,
                _policyWarningTitleLabel,
                _policyWarningLinkLabel,
                _serverUrlTextBox.Text);
            bool credentialsMissing = !new TalkServiceConfiguration(
                _serverUrlTextBox.Text, _usernameTextBox.Text, _appPasswordTextBox.Text).IsComplete();
            if (!Result.IsManagedTransportTlsValid)
            {
                warningVisible = true;
                _policyWarningPanel.Visible = true;
                _policyWarningTitleLabel.ForeColor = _themePalette.ErrorText;
                _policyWarningPanel.BackColor = Color.FromArgb(20, _themePalette.ErrorText);
                _policyWarningTextLabel.Text = Strings.ManagedTlsPolicyInvalid;
                _policyWarningLinkLabel.Visible = false;
                _policyWarningLinkLabel.Tag = null;
            }
            else if (_connectionSetupPending || credentialsMissing)
            {
                warningVisible = true;
                _policyWarningPanel.Visible = true;
                _policyWarningTitleLabel.Text = Strings.ConnectionSetupTitle;
                _policyWarningTitleLabel.ForeColor = _authenticationRejected
                    ? Color.FromArgb(156, 108, 0) : _themePalette.LinkText;
                _policyWarningPanel.BackColor = Color.FromArgb(20, _policyWarningTitleLabel.ForeColor);
                _policyWarningTextLabel.Text = _authenticationRejected
                    ? Strings.ConnectionSignInRequired : Strings.ConnectionSetupMessage;
                _policyWarningLinkLabel.Visible = false;
                _policyWarningLinkLabel.Tag = null;
            }
            else if (Result.IsEnterpriseRollout)
            {
                string notice = _backendPolicyStatus == null
                    ? string.Empty
                    : PolicyUiHelper.GetEnterpriseRolloutNotice(Result, _backendPolicyStatus);
                warningVisible = !string.IsNullOrEmpty(notice);
                _policyWarningPanel.Visible = warningVisible;
                _policyWarningTextLabel.Text = notice;
                _policyWarningLinkLabel.Visible = false;
                _policyWarningLinkLabel.Tag = null;
            }
            ApplyPolicyDefaultsToControls();

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Policy status applied in settings (trigger=" + (trigger ?? "n/a")
                + ", active=" + PolicyUiHelper.IsPolicyActive(_backendPolicyStatus).ToString(CultureInfo.InvariantCulture)
                + ", share=" + PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "share").ToString(CultureInfo.InvariantCulture)
                + ", talk=" + PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "talk").ToString(CultureInfo.InvariantCulture)
                + ", emailSignature=" + PolicyUiHelper.IsPolicyDomainActive(_backendPolicyStatus, "email_signature").ToString(CultureInfo.InvariantCulture)
                + ", warningVisible=" + warningVisible.ToString(CultureInfo.InvariantCulture)
                + ", accessStatus=" + (_backendPolicyStatus != null ? _backendPolicyStatus.AccessStatus : "n/a")
                + ", mode=" + (_backendPolicyStatus != null ? _backendPolicyStatus.Mode : "local")
                + ", reason=" + (_backendPolicyStatus != null ? _backendPolicyStatus.Reason : "n/a")
                + ").");

            UpdateControlState();
            ApplyResponsiveLayout(false);
        }

        private void ApplyPolicyDefaultsToControls()
        {
            AddinSettings effective = Result.ResolvePolicyDefaults(_backendPolicyStatus);
            bool previousApplying = _applyingPolicyDefaults;
            _applyingPolicyDefaults = true;
            try
            {
                _fileLinkBaseTextBox.Text = effective.FileLinkBasePath ?? string.Empty;
                _sharingDefaultShareNameTextBox.Text = effective.SharingDefaultShareName ?? string.Empty;
                _sharingDefaultPermCreateCheckBox.Checked = effective.SharingDefaultPermCreate;
                _sharingDefaultPermWriteCheckBox.Checked = effective.SharingDefaultPermWrite;
                _sharingDefaultPermDeleteCheckBox.Checked = effective.SharingDefaultPermDelete;
                _sharingDefaultPasswordCheckBox.Checked = effective.SharingDefaultPasswordEnabled;
                _sharingDefaultPasswordSeparateCheckBox.Checked =
                    PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus) && effective.SharingDefaultPasswordSeparateEnabled;
                SharePasswordDeliveryModeComboHelper.Select(_sharingPasswordDeliveryModeCombo, effective.SharingDefaultPasswordDeliveryMode);
                int expireDays = effective.SharingDefaultExpireDays;
                if (expireDays <= 0)
                {
                    expireDays = 7;
                }
                if (expireDays > 3650)
                {
                    expireDays = 3650;
                }
                _sharingDefaultExpireDaysUpDown.Value = expireDays;
                _sharingAttachmentsAlwaysCheckBox.Checked = effective.SharingAttachmentsAlwaysConnector;
                _sharingAttachmentsOfferAboveCheckBox.Checked = effective.SharingAttachmentsOfferAboveEnabled;
                int offerAboveMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(effective.SharingAttachmentsOfferAboveMb);
                decimal clampedOfferAbove = Math.Max(
                    _sharingAttachmentsOfferAboveMbUpDown.Minimum,
                    Math.Min(_sharingAttachmentsOfferAboveMbUpDown.Maximum, (decimal)offerAboveMb));
                _sharingAttachmentsOfferAboveMbUpDown.Value = clampedOfferAbove;
                SelectAttachmentLinkTarget(AttachmentLinkTargetPolicy.Resolve(
                    Result.SharingAttachmentLinkTarget,
                    _backendPolicyStatus));
                _talkDefaultPasswordCheckBox.Checked = effective.TalkDefaultPasswordEnabled;
                _talkDefaultAddUsersCheckBox.Checked = effective.TalkDefaultAddUsers;
                _talkDefaultAddGuestsCheckBox.Checked = effective.TalkDefaultAddGuests;
                _talkDefaultLobbyCheckBox.Checked = effective.TalkDefaultLobbyEnabled;
                _talkDefaultSearchCheckBox.Checked = effective.TalkDefaultSearchVisible;
                _talkDeleteRoomOnEventDeleteCheckBox.Checked = effective.TalkDeleteRoomOnEventDelete;
                _emailSignatureOnComposeCheckBox.Checked = EmailSignaturePolicyService.ResolveFlag(
                    _backendPolicyStatus,
                    "email_signature_on_compose",
                    Result.EmailSignatureOnCompose);
                _emailSignatureOnReplyCheckBox.Checked = EmailSignaturePolicyService.ResolveFlag(
                    _backendPolicyStatus,
                    "email_signature_on_reply",
                    Result.EmailSignatureOnReply);
                _emailSignatureOnForwardCheckBox.Checked = EmailSignaturePolicyService.ResolveFlag(
                    _backendPolicyStatus,
                    "email_signature_on_forward",
                    Result.EmailSignatureOnForward);
                TalkRoomTypeComboHelper.Select(
                    _talkDefaultRoomTypeCombo,
                    effective.TalkDefaultRoomType);
                UpdateTalkRoomTypeTooltip();
                RefreshLanguageOverrideCombos(effective.ShareBlockLang, effective.EventDescriptionLang);
            }
            finally
            {
                _applyingPolicyDefaults = previousApplying;
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

        private void OnSettingsFormClosed(object sender, FormClosedEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                return;
            }
            try
            {
                TransportSecurityConfigurator.Restore(_runtimeSecurityProtocolAtOpen);
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
            bool managedUrlLocked = Result != null && Result.ManagedNextcloudUrlLocked;

            _serverUrlTextBox.Enabled = !managedUrlLocked && !_isBusy;
            _usernameTextBox.Enabled = manual && !_isBusy;
            _appPasswordTextBox.Enabled = manual && !_isBusy;
            _loginFlowButton.Enabled = !manual && !_isBusy;
            _testButton.Enabled = !_isBusy;
            _disabledTooltipHints.Apply(
                _serverUrlTextBox,
                managedUrlLocked ? Strings.TooltipManagedNextcloudUrl : string.Empty,
                managedUrlLocked);

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
            _updateNotifyCheckBox.Enabled = !_isBusy;
            _updateCheckButton.Enabled = !_isBusy;
            _updateDownloadLink.Enabled = !_isBusy && !string.IsNullOrWhiteSpace(_updateOpenUrl);

            bool lockShareBase = IsPolicyLocked("share", "share_base_directory");
            bool lockShareName = IsPolicyLocked("share", "share_name_template");
            bool lockSharePermCreate = IsPolicyLocked("share", "share_permission_upload");
            bool lockSharePermWrite = IsPolicyLocked("share", "share_permission_edit");
            bool lockSharePermDelete = IsPolicyLocked("share", "share_permission_delete");
            bool lockSharePassword = IsPolicyLocked("share", "share_set_password");
            bool lockSharePasswordSeparate = IsPolicyLocked("share", "share_send_password_separately");
            bool lockSharePasswordDeliveryMode = IsPolicyLocked("share", "share_send_password_mode");
            bool lockShareExpire = IsPolicyLocked("share", "share_expire_days");
            bool lockShareLang = IsPolicyLocked("share", "language_share_html_block");
            bool lockTalkPassword = IsPolicyLocked("talk", "talk_set_password");
            bool lockTalkLobby = IsPolicyLocked("talk", "talk_lobby_active");
            bool lockTalkSearch = IsPolicyLocked("talk", "talk_show_in_search");
            bool lockTalkRoomType = IsPolicyLocked("talk", "talk_room_type");
            bool lockTalkLang = IsPolicyLocked("talk", "language_talk_description");
            bool lockTalkUsers = IsPolicyLocked("talk", "talk_add_users");
            bool lockTalkGuests = IsPolicyLocked("talk", "talk_add_guests");
            bool lockTalkDeleteRoomOnEventDelete = IsPolicyLocked("talk", "talk_delete_room_on_event_delete");
            bool lockSignatureOnCompose = IsPolicyLocked("email_signature", "email_signature_on_compose");
            bool lockSignatureOnReply = IsPolicyLocked("email_signature", "email_signature_on_reply");
            bool lockSignatureOnForward = IsPolicyLocked("email_signature", "email_signature_on_forward");
            bool separatePasswordAvailable = PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus);
            string separatePasswordUnavailableTooltip = PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(_backendPolicyStatus);
            bool passwordDeliveryModeAvailable = PolicyUiHelper.HasPasswordDeliveryMode(_backendPolicyStatus);
            string passwordDeliveryUnavailableTooltip = PolicyUiHelper.GetPasswordDeliveryModeUnavailableTooltip(_backendPolicyStatus);
            bool emailSignatureAvailable = IsEmailSignaturePolicyAvailable();
            string emailSignatureUnavailableTooltip = GetEmailSignatureUnavailableTooltip();

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
            bool deliveryModeEnabled =
                passwordDeliveryModeAvailable
                && _sharingDefaultPasswordCheckBox.Checked
                && _sharingDefaultPasswordSeparateCheckBox.Checked
                && !lockSharePasswordDeliveryMode
                && !_isBusy;
            _sharingPasswordDeliveryModeCombo.Enabled = deliveryModeEnabled;
            if (!passwordDeliveryModeAvailable)
            {
                SharePasswordDeliveryModeComboHelper.Select(_sharingPasswordDeliveryModeCombo, SharePasswordDeliveryMode.Plain);
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
            _talkDeleteRoomOnEventDeleteCheckBox.Enabled = !lockTalkDeleteRoomOnEventDelete && !_isBusy;
            _emailSignatureOnComposeCheckBox.Enabled = emailSignatureAvailable && !lockSignatureOnCompose && !_isBusy;
            _emailSignatureOnReplyCheckBox.Enabled = emailSignatureAvailable && _emailSignatureOnComposeCheckBox.Checked && !lockSignatureOnReply && !_isBusy;
            _emailSignatureOnForwardCheckBox.Enabled = emailSignatureAvailable && _emailSignatureOnComposeCheckBox.Checked && !lockSignatureOnForward && !_isBusy;

            _disabledTooltipHints.Apply(_fileLinkBaseTextBox, lockShareBase ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareBase, _fileLinkBaseHintLabel);
            _disabledTooltipHints.Apply(_sharingDefaultShareNameTextBox, lockShareName ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareName, _sharingDefaultShareNameLabel);
            _disabledTooltipHints.Apply(_sharingDefaultPermCreateCheckBox, lockSharePermCreate ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermCreate);
            _disabledTooltipHints.Apply(_sharingDefaultPermWriteCheckBox, lockSharePermWrite ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermWrite);
            _disabledTooltipHints.Apply(_sharingDefaultPermDeleteCheckBox, lockSharePermDelete ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePermDelete);
            _disabledTooltipHints.Apply(_sharingDefaultPasswordCheckBox, lockSharePassword ? Strings.PolicyAdminControlledTooltip : string.Empty, lockSharePassword);
            _disabledTooltipHints.Apply(
                _sharingDefaultPasswordSeparateCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockSharePasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockSharePasswordSeparate);
            _disabledTooltipHints.Apply(
                _sharingPasswordDeliveryModeCombo,
                !passwordDeliveryModeAvailable
                    ? passwordDeliveryUnavailableTooltip
                    : (lockSharePasswordDeliveryMode
                        ? Strings.PolicyAdminControlledTooltip
                        : (!_sharingDefaultPasswordSeparateCheckBox.Checked ? Strings.SharingPasswordDeliveryEnableSeparateTooltip : string.Empty)),
                !passwordDeliveryModeAvailable
                    || lockSharePasswordDeliveryMode
                    || !_sharingDefaultPasswordSeparateCheckBox.Checked,
                _sharingPasswordDeliveryModeLabel);
            _disabledTooltipHints.Apply(_sharingDefaultExpireDaysUpDown, lockShareExpire ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareExpire, _sharingDefaultExpireDaysLabel);
            _disabledTooltipHints.Apply(_shareBlockLangCombo, lockShareLang ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareLang, _shareBlockLangLabel);
            _disabledTooltipHints.Apply(_talkDefaultPasswordCheckBox, lockTalkPassword ? Strings.PolicyAdminControlledTooltip : string.Empty, lockTalkPassword);
            _disabledTooltipHints.Apply(_talkDefaultLobbyCheckBox, lockTalkLobby ? Strings.PolicyAdminControlledTooltip : Strings.TooltipLobby, lockTalkLobby);
            _disabledTooltipHints.Apply(_talkDefaultSearchCheckBox, lockTalkSearch ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSearchVisible, lockTalkSearch);
            _disabledTooltipHints.Apply(_talkDeleteRoomOnEventDeleteCheckBox, lockTalkDeleteRoomOnEventDelete ? Strings.PolicyAdminControlledTooltip : Strings.TooltipDeleteRoomOnEventDelete, lockTalkDeleteRoomOnEventDelete);
            _disabledTooltipHints.Apply(
                _emailSignatureOnComposeCheckBox,
                !emailSignatureAvailable
                    ? emailSignatureUnavailableTooltip
                    : (lockSignatureOnCompose ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !emailSignatureAvailable || lockSignatureOnCompose);
            _disabledTooltipHints.Apply(
                _emailSignatureOnReplyCheckBox,
                !emailSignatureAvailable
                    ? emailSignatureUnavailableTooltip
                    : (lockSignatureOnReply ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !emailSignatureAvailable || lockSignatureOnReply);
            _disabledTooltipHints.Apply(
                _emailSignatureOnForwardCheckBox,
                !emailSignatureAvailable
                    ? emailSignatureUnavailableTooltip
                    : (lockSignatureOnForward ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !emailSignatureAvailable || lockSignatureOnForward);
            bool standardTalkRoomTypeSelected =
                TalkRoomTypeComboHelper.GetSelected(
                    _talkDefaultRoomTypeCombo,
                    TalkRoomType.EventConversation) == TalkRoomType.StandardRoom;
            _disabledTooltipHints.Apply(
                _talkDefaultRoomTypeCombo,
                lockTalkRoomType
                    ? Strings.PolicyAdminControlledTooltip
                    : (standardTalkRoomTypeSelected
                        ? Strings.TooltipRoomTypeStandard
                        : Strings.TooltipRoomTypeEvent),
                lockTalkRoomType,
                _talkDefaultRoomTypeLabel);
            _disabledTooltipHints.Apply(_eventDescriptionLangCombo, lockTalkLang ? Strings.PolicyAdminControlledTooltip : string.Empty, lockTalkLang, _eventDescriptionLangLabel);
            UpdateSharingAttachmentOptionsState();
            UpdateTlsOptionsState();
        }

        private void SetBusy(bool busy)
        {
            _isBusy = busy;
            if (IsDisposed || Disposing)
            {
                return;
            }
            Cursor.Current = busy ? Cursors.WaitCursor : Cursors.Default;
            _saveButton.Enabled = !busy;
            _cancelButton.Enabled = !busy;
            _debugOpenLink.Enabled = !busy;
            UpdateControlState();
        }

        private void SetStatus(string message, bool isError)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }
            _statusLabel.Text = message;
            _statusLabel.ForeColor = isError ? _themePalette.ErrorText : _themePalette.SuccessText;
            ApplyResponsiveLayout(false);
        }

    }
}
