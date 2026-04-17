/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{

    /// <summary>
    /// Dialog for configuring and creating a Nextcloud Talk room for the active appointment.
    /// </summary>
    internal sealed partial class TalkLinkForm : Form
    {
        private static readonly string DefaultTitle = Strings.TalkDefaultTitle;
        private const int DefaultMinPasswordLength = 5;
        private const int ModeratorDropdownMaxRows = 6;
        private const int ModeratorDropdownMargin = 4;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

        private readonly Label _titleLabel = new Label();
        private readonly Label _roomTypeLabel = new Label();
        private readonly Label _passwordLabel = new Label();
        private readonly GroupBox _settingsGroup = new GroupBox();
        private readonly Label _eventSupportHintLabel = new Label();
        private readonly TextBox _titleTextBox = new TextBox();
        private readonly ComboBox _roomTypeComboBox = new ComboBox();
        private readonly CheckBox _passwordToggleCheckBox = new CheckBox();
        private readonly TextBox _passwordTextBox = new TextBox();
        private readonly Button _passwordGenerateButton = new Button();
        private readonly CheckBox _addUsersCheckBox = new CheckBox();
        private readonly CheckBox _addGuestsCheckBox = new CheckBox();
        private readonly CheckBox _lobbyCheckBox = new CheckBox();
        private readonly CheckBox _searchCheckBox = new CheckBox();
        private readonly GroupBox _moderatorGroup = new GroupBox();
        private readonly PictureBox _moderatorAvatarBox = new PictureBox();
        private readonly TextBox _moderatorTextBox = new TextBox();
        private readonly Button _moderatorClearButton = new Button();
        private readonly ListBox _moderatorListBox = new ListBox();
        private readonly Panel _moderatorAddressbookWarningPanel = new Panel();
        private readonly Label _moderatorAddressbookWarningTitleLabel = new Label();
        private readonly Label _moderatorAddressbookWarningTextLabel = new Label();
        private readonly LinkLabel _moderatorAddressbookWarningLinkLabel = new LinkLabel();
        private readonly Panel _policyWarningPanel = new Panel();
        private readonly Label _policyWarningTitleLabel = new Label();
        private readonly Label _policyWarningTextLabel = new Label();
        private readonly LinkLabel _policyWarningLinkLabel = new LinkLabel();
        private readonly Label _moderatorHintLabel = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly DisabledControlTooltipHintHelper _disabledTooltipHints;
        private readonly Button _okButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private const int HeaderHeight = 48;
        private readonly bool _eventConversationsSupported;
        private readonly string _serverVersionHint;
        private readonly PasswordPolicyInfo _passwordPolicy;
        private readonly TalkServiceConfiguration _configuration;
        private readonly BackendPolicyStatus _backendPolicyStatus;
        private readonly List<NextcloudUser> _userDirectory;
        private readonly bool _systemAddressbookAvailable;
        private readonly string _systemAddressbookError;
        private NextcloudUser _selectedModerator;
        private readonly Dictionary<string, Image> _avatarCache = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _avatarLoading = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _avatarLock = new object();
        private bool _layoutApplying;
        private bool _moderatorSearchLockLogged;

        private string _talkTitle;
        private string _talkPassword;
        private bool _lobbyUntilStart;
        private bool _searchVisible;
        private TalkRoomType _selectedRoomType;
        private bool _addUsers;
        private bool _addGuests;
        private string _delegateModeratorId;
        private string _delegateModeratorName;

        internal string TalkTitle
        {
            get { return _talkTitle; }
            private set { _talkTitle = value; }
        }

        internal string TalkPassword
        {
            get { return _talkPassword; }
            private set { _talkPassword = value; }
        }

        internal bool LobbyUntilStart
        {
            get { return _lobbyUntilStart; }
            private set { _lobbyUntilStart = value; }
        }

        internal bool SearchVisible
        {
            get { return _searchVisible; }
            private set { _searchVisible = value; }
        }

        internal TalkRoomType SelectedRoomType
        {
            get { return _selectedRoomType; }
            private set { _selectedRoomType = value; }
        }

        internal bool AddUsers
        {
            get { return _addUsers; }
            private set { _addUsers = value; }
        }

        internal bool AddGuests
        {
            get { return _addGuests; }
            private set { _addGuests = value; }
        }

        internal string DelegateModeratorId
        {
            get { return _delegateModeratorId; }
            private set { _delegateModeratorId = value; }
        }

        internal string DelegateModeratorName
        {
            get { return _delegateModeratorName; }
            private set { _delegateModeratorName = value; }
        }

        internal TalkLinkForm(
            AddinSettings defaults,
            TalkServiceConfiguration configuration,
            PasswordPolicyInfo passwordPolicy,
            BackendPolicyStatus policyStatus,
            List<NextcloudUser> userDirectory,
            IfbAddressBookCache.SystemAddressbookStatus addressbookStatus,
            string appointmentSubject,
            DateTime startTime,
            DateTime endTime)
        {
            _eventConversationsSupported = DetermineEventConversationSupport(defaults, out _serverVersionHint);
            _passwordPolicy = passwordPolicy;
            _configuration = configuration;
            _backendPolicyStatus = policyStatus;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            _userDirectory = userDirectory ?? new List<NextcloudUser>();
            _systemAddressbookAvailable = addressbookStatus != null && addressbookStatus.Available;
            _systemAddressbookError = addressbookStatus != null ? (addressbookStatus.Error ?? string.Empty) : string.Empty;

            Text = Strings.TalkFormTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ControlBox = true;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 520);
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimumSize = new Size(ScaleLogical(580), ScaleLogical(580));
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeComponents();
            ApplyDefaults(defaults, appointmentSubject);
            ApplyDialogLayout(true);

            UiThemeManager.ApplyToForm(this, _toolTip);
        }

        /**
         * Builds all dialog controls (title, password, options, buttons).
         */
        private void InitializeComponents()
        {
            _titleLabel.Text = Strings.TalkTitleLabel;
            _titleLabel.AutoSize = true;

            _roomTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _roomTypeComboBox.Items.Add(new RoomTypeOption(TalkRoomType.EventConversation, Strings.TalkEventRadio));
            _roomTypeComboBox.Items.Add(new RoomTypeOption(TalkRoomType.StandardRoom, Strings.TalkStandardRadio));
            _roomTypeComboBox.SelectedIndexChanged += (s, e) => UpdateRoomTypeTooltip();

            _roomTypeLabel.Text = Strings.TalkRoomGroup;
            _roomTypeLabel.AutoSize = true;

            _passwordToggleCheckBox.Text = Strings.TalkPasswordSetCheck;
            _passwordToggleCheckBox.AutoSize = true;
            _passwordToggleCheckBox.CheckedChanged += (s, e) => UpdatePasswordState();

            _passwordLabel.Text = Strings.TalkPasswordLabel;
            _passwordLabel.AutoSize = true;

            _passwordTextBox.UseSystemPasswordChar = false;

            _passwordGenerateButton.Text = Strings.TalkPasswordGenerate;
            _passwordGenerateButton.AutoSize = false;
            _passwordGenerateButton.TextAlign = ContentAlignment.MiddleCenter;
            _passwordGenerateButton.Click += (s, e) => GeneratePassword();

            _settingsGroup.Text = Strings.TalkSettingsGroup;

            InitializeSettingsOptionCheckBox(_addUsersCheckBox, Strings.TalkAddUsersCheck);
            InitializeSettingsOptionCheckBox(_addGuestsCheckBox, Strings.TalkAddGuestsCheck);
            InitializeSettingsOptionCheckBox(_lobbyCheckBox, Strings.TalkLobbyCheck);
            InitializeSettingsOptionCheckBox(_searchCheckBox, Strings.TalkSearchCheck);

            _moderatorGroup.Text = Strings.TalkModeratorGroup;

            _moderatorAvatarBox.Size = new Size(32, 32);
            _moderatorAvatarBox.SizeMode = PictureBoxSizeMode.Zoom;
            _moderatorAvatarBox.BackColor = Color.Transparent;
            _moderatorGroup.Controls.Add(_moderatorAvatarBox);

            _moderatorTextBox.TextChanged += (s, e) => UpdateModeratorResults();
            _moderatorTextBox.Enter += (s, e) => UpdateModeratorResults();
            _moderatorTextBox.MouseDown += (s, e) => UpdateModeratorResults();
            _moderatorGroup.Controls.Add(_moderatorTextBox);

            _moderatorClearButton.Text = Strings.TalkModeratorClear;
            _moderatorClearButton.AutoSize = false;
            _moderatorClearButton.TextAlign = ContentAlignment.MiddleCenter;
            _moderatorClearButton.Click += (s, e) => ClearModerator();
            _moderatorGroup.Controls.Add(_moderatorClearButton);

            _moderatorAddressbookWarningPanel.Visible = false;
            _moderatorAddressbookWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _moderatorAddressbookWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _moderatorAddressbookWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };
            _moderatorGroup.Controls.Add(_moderatorAddressbookWarningPanel);

            _moderatorAddressbookWarningTitleLabel.AutoSize = true;
            _moderatorAddressbookWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _moderatorAddressbookWarningTitleLabel.Font = new Font(
                _moderatorAddressbookWarningTitleLabel.Font,
                FontStyle.Bold);
            _moderatorAddressbookWarningTitleLabel.Text = "\u26a0 " + Strings.TalkSystemAddressbookRequiredShort;
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningTitleLabel);

            _moderatorAddressbookWarningTextLabel.AutoSize = true;
            _moderatorAddressbookWarningTextLabel.Text = Strings.TalkSystemAddressbookRequiredMessage;
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningTextLabel);

            _moderatorAddressbookWarningLinkLabel.AutoSize = true;
            _moderatorAddressbookWarningLinkLabel.Text = Strings.TalkSystemAddressbookAdminLinkLabel;
            _moderatorAddressbookWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _moderatorAddressbookWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _moderatorAddressbookWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _moderatorAddressbookWarningLinkLabel.LinkClicked += (s, e) => OpenSystemAddressbookSetupGuide();
            _moderatorAddressbookWarningPanel.Controls.Add(_moderatorAddressbookWarningLinkLabel);

            _moderatorListBox.DrawMode = DrawMode.OwnerDrawFixed;
            _moderatorListBox.ItemHeight = 34;
            _moderatorListBox.IntegralHeight = false;
            _moderatorListBox.ScrollAlwaysVisible = true;
            _moderatorListBox.Visible = false;
            _moderatorListBox.DoubleClick += (s, e) => SelectModeratorFromList();
            _moderatorListBox.DrawItem += OnModeratorDrawItem;
            _moderatorTextBox.Leave += (s, e) =>
            {
                BeginInvoke((Action)(() =>
                {
                    if (!_moderatorListBox.Focused)
                    {
                        HideModeratorDropdown();
                    }
                }));
            };
            _moderatorListBox.Leave += (s, e) =>
            {
                BeginInvoke((Action)(() =>
                {
                    if (!_moderatorTextBox.Focused)
                    {
                        HideModeratorDropdown();
                    }
                }));
            };

            _moderatorHintLabel.Text = Strings.TalkModeratorHint;
            _moderatorHintLabel.ForeColor = Color.DimGray;
            _moderatorHintLabel.Click += (s, e) =>
            {
                _moderatorTextBox.Focus();
                UpdateModeratorResults();
            };
            _moderatorGroup.Controls.Add(_moderatorHintLabel);

            _eventSupportHintLabel.AutoSize = true;
            _eventSupportHintLabel.MaximumSize = new Size(ScaleLogical(260), 0);
            _eventSupportHintLabel.ForeColor = Color.DimGray;
            _eventSupportHintLabel.Visible = !_eventConversationsSupported;
            if (_eventSupportHintLabel.Visible)
            {
                var versionInfo = string.IsNullOrEmpty(_serverVersionHint) ? Strings.TalkVersionUnknown : _serverVersionHint;
                _eventSupportHintLabel.Text = string.Format(Strings.TalkEventHint, versionInfo);
            }

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

            _policyWarningTitleLabel.AutoSize = true;
            _policyWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _policyWarningTitleLabel.Font = new Font(
                _policyWarningTitleLabel.Font,
                FontStyle.Bold);
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
                    LogCategories.Talk,
                    "Failed to open policy admin guide URL.");
            _policyWarningPanel.Controls.Add(_policyWarningLinkLabel);

            _okButton.Text = Strings.DialogOk;
            _okButton.AutoSize = false;
            _okButton.DialogResult = DialogResult.OK;
            _okButton.Click += OnOkButtonClick;

            _cancelButton.Text = Strings.DialogCancel;
            _cancelButton.AutoSize = false;
            _cancelButton.DialogResult = DialogResult.Cancel;

            Controls.Add(_titleLabel);
            Controls.Add(_titleTextBox);
            Controls.Add(_roomTypeLabel);
            Controls.Add(_roomTypeComboBox);
            Controls.Add(_eventSupportHintLabel);
            Controls.Add(_passwordToggleCheckBox);
            Controls.Add(_passwordLabel);
            Controls.Add(_passwordTextBox);
            Controls.Add(_passwordGenerateButton);
            Controls.Add(_settingsGroup);
            Controls.Add(_moderatorGroup);
            Controls.Add(_policyWarningPanel);
            Controls.Add(_okButton);
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;

            _toolTip.AutoPopDelay = 20000;
            _toolTip.InitialDelay = 250;
            _toolTip.ReshowDelay = 150;
            _toolTip.SetToolTip(_addUsersCheckBox, Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_addGuestsCheckBox, Strings.TooltipAddGuests);
            _toolTip.SetToolTip(_lobbyCheckBox, Strings.TooltipLobby);
            _toolTip.SetToolTip(_searchCheckBox, Strings.TooltipSearchVisible);
            _toolTip.SetToolTip(_moderatorTextBox, Strings.TooltipModerator);
            _toolTip.SetToolTip(_moderatorClearButton, Strings.TooltipModerator);
            UpdateRoomTypeTooltip();

            Controls.Add(_moderatorListBox);
            ApplyDialogLayout(false);
        }

        private void InitializeSettingsOptionCheckBox(CheckBox checkBox, string text)
        {
            if (checkBox == null)
            {
                return;
            }

            checkBox.Text = text ?? string.Empty;
            checkBox.AutoSize = true;
            _settingsGroup.Controls.Add(checkBox);
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = HeaderHeight;
            _headerPanel.Dock = DockStyle.Top;
            Controls.Add(_headerPanel);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (_layoutApplying)
            {
                return;
            }

            ApplyDialogLayout(false);
        }

        private void ApplyDialogLayout(bool ensureClientHeight)
        {
            if (_layoutApplying || IsDisposed || Disposing)
            {
                return;
            }

            _layoutApplying = true;
            try
            {
                int outerPadding = ScaleLogical(18);
                int labelX = ScaleLogical(16);
                int inputX = ScaleLogical(170);
                int verticalGap = ScaleLogical(12);
                int rowGap = ScaleLogical(14);

                int y = _headerPanel.Bottom + ScaleLogical(16);
                int inputWidth = Math.Max(ScaleLogical(180), ClientSize.Width - inputX - ScaleLogical(18));

                _titleLabel.Location = new Point(labelX, y + ScaleLogical(4));
                _titleTextBox.SetBounds(inputX, y, inputWidth, _titleTextBox.PreferredHeight + ScaleLogical(2));
                y = Math.Max(_titleLabel.Bottom, _titleTextBox.Bottom) + rowGap;

                int roomTypeComboHeight = Math.Max(_roomTypeComboBox.PreferredHeight + ScaleLogical(2), _roomTypeComboBox.ItemHeight + ScaleLogical(8));
                _roomTypeLabel.Location = new Point(labelX, y + Math.Max(0, (roomTypeComboHeight - _roomTypeLabel.PreferredHeight) / 2));
                _roomTypeComboBox.SetBounds(inputX, y, inputWidth, roomTypeComboHeight);
                y = Math.Max(_roomTypeLabel.Bottom, _roomTypeComboBox.Bottom) + rowGap;

                _eventSupportHintLabel.Visible = !_eventConversationsSupported;
                if (_eventSupportHintLabel.Visible)
                {
                    _eventSupportHintLabel.Location = new Point(inputX, y - ScaleLogical(2));
                    _eventSupportHintLabel.MaximumSize = new Size(inputWidth, 0);
                    y = _eventSupportHintLabel.Bottom + verticalGap;
                }

                if (_policyWarningPanel.Visible)
                {
                    int warningPadding = ScaleLogical(8);
                    int panelWidth = Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2));
                    int warningTextWidth = Math.Max(ScaleLogical(160), panelWidth - (warningPadding * 2));

                    _policyWarningTitleLabel.Location = new Point(warningPadding, warningPadding);
                    _policyWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningTextTop = _policyWarningTitleLabel.Bottom + ScaleLogical(4);
                    _policyWarningTextLabel.Location = new Point(warningPadding, warningTextTop);
                    _policyWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                    int warningLinkTop = _policyWarningTextLabel.Bottom + ScaleLogical(6);
                    _policyWarningLinkLabel.Location = new Point(warningPadding, warningLinkTop);

                    int panelHeight = _policyWarningLinkLabel.Bottom + warningPadding;
                    _policyWarningPanel.SetBounds(outerPadding, y, panelWidth, panelHeight);
                    y = _policyWarningPanel.Bottom + verticalGap;
                }
                else
                {
                    _policyWarningPanel.SetBounds(outerPadding, y, Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2)), 0);
                }

                _passwordToggleCheckBox.Location = new Point(outerPadding, y);
                y = _passwordToggleCheckBox.Bottom + ScaleLogical(10);

                int ignoredGenerateMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(_passwordGenerateButton, out ignoredGenerateMinWidth);
                int generateButtonWidth = _passwordGenerateButton.Width;
                int generateButtonHeight = _passwordGenerateButton.Height;
                int passwordWidth = Math.Max(ScaleLogical(120), inputWidth - generateButtonWidth - ScaleLogical(8));

                _passwordLabel.Location = new Point(labelX, y + ScaleLogical(4));
                _passwordTextBox.SetBounds(inputX, y, passwordWidth, _passwordTextBox.PreferredHeight + ScaleLogical(2));
                _passwordGenerateButton.SetBounds(_passwordTextBox.Right + ScaleLogical(8), y - ScaleLogical(2), generateButtonWidth, generateButtonHeight);
                y = Math.Max(_passwordLabel.Bottom, Math.Max(_passwordTextBox.Bottom, _passwordGenerateButton.Bottom)) + ScaleLogical(16);

                int groupWidth = Math.Max(ScaleLogical(260), ClientSize.Width - (outerPadding * 2));

                _settingsGroup.SetBounds(outerPadding, y, groupWidth, ScaleLogical(132));
                int settingsLeft = ScaleLogical(12);
                int settingsTop = ScaleLogical(24);
                int settingsLineGap = ScaleLogical(24);
                _addUsersCheckBox.Location = new Point(settingsLeft, settingsTop);
                _addGuestsCheckBox.Location = new Point(settingsLeft, settingsTop + settingsLineGap);
                _lobbyCheckBox.Location = new Point(settingsLeft, settingsTop + (settingsLineGap * 2));
                _searchCheckBox.Location = new Point(settingsLeft, settingsTop + (settingsLineGap * 3));
                int settingsGroupHeight = _searchCheckBox.Bottom + ScaleLogical(14);
                _settingsGroup.Height = Math.Max(ScaleLogical(108), settingsGroupHeight);

                y = _settingsGroup.Bottom + verticalGap;
                _moderatorGroup.SetBounds(outerPadding, y, groupWidth, ScaleLogical(122));
                int moderatorGroupHeight = LayoutModeratorGroupControls();
                if (moderatorGroupHeight > _moderatorGroup.Height)
                {
                    _moderatorGroup.Height = moderatorGroupHeight;
                    LayoutModeratorGroupControls();
                }

                var footerButtons = new List<Button> { _okButton, _cancelButton };
                int minClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);
                if (ensureClientHeight && minClientWidth > ClientSize.Width)
                {
                    ClientSize = new Size(minClientWidth, ClientSize.Height);
                }

                int buttonHeight = Math.Max(_okButton.Height, _cancelButton.Height);
                int minClientHeight = _moderatorGroup.Bottom + ScaleLogical(16) + buttonHeight + FooterButtonLayoutHelper.DefaultBottomPadding;
                if (ensureClientHeight && minClientHeight > ClientSize.Height)
                {
                    ClientSize = new Size(ClientSize.Width, minClientHeight);
                }

                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    footerButtons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing);

                PositionModeratorDropdown();
            }
            finally
            {
                _layoutApplying = false;
            }
        }

        private int LayoutModeratorGroupControls()
        {
            int innerPadding = ScaleLogical(12);
            int avatarSize = ScaleLogical(32);
            int textTop = ScaleLogical(24);
            int ignoredClearMinWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(_moderatorClearButton, out ignoredClearMinWidth);
            int clearButtonWidth = _moderatorClearButton.Width;
            int clearButtonHeight = _moderatorClearButton.Height;

            _moderatorAvatarBox.SetBounds(innerPadding, ScaleLogical(22), avatarSize, avatarSize);

            int textLeft = _moderatorAvatarBox.Right + ScaleLogical(8);
            int textWidth = Math.Max(
                ScaleLogical(120),
                _moderatorGroup.ClientSize.Width - textLeft - innerPadding - clearButtonWidth - ScaleLogical(8));
            _moderatorTextBox.SetBounds(textLeft, textTop, textWidth, _moderatorTextBox.PreferredHeight + ScaleLogical(2));

            _moderatorClearButton.SetBounds(
                _moderatorTextBox.Right + ScaleLogical(8),
                textTop - ScaleLogical(2),
                clearButtonWidth,
                clearButtonHeight);

            int rowBottom = Math.Max(
                _moderatorAvatarBox.Bottom,
                Math.Max(_moderatorTextBox.Bottom, _moderatorClearButton.Bottom));
            int contentTop = rowBottom + ScaleLogical(8);

            if (_moderatorAddressbookWarningPanel.Visible)
            {
                int panelPadding = ScaleLogical(8);
                int panelWidth = Math.Max(ScaleLogical(160), _moderatorGroup.ClientSize.Width - (innerPadding * 2));
                int warningTextWidth = Math.Max(ScaleLogical(120), panelWidth - (panelPadding * 2));

                _moderatorAddressbookWarningTitleLabel.Location = new Point(panelPadding, panelPadding);
                _moderatorAddressbookWarningTitleLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningTextTop = _moderatorAddressbookWarningTitleLabel.Bottom + ScaleLogical(4);
                _moderatorAddressbookWarningTextLabel.Location = new Point(panelPadding, warningTextTop);
                _moderatorAddressbookWarningTextLabel.MaximumSize = new Size(warningTextWidth, 0);

                int warningLinkTop = _moderatorAddressbookWarningTextLabel.Bottom + ScaleLogical(6);
                _moderatorAddressbookWarningLinkLabel.Location = new Point(panelPadding, warningLinkTop);

                int panelHeight = _moderatorAddressbookWarningLinkLabel.Bottom + panelPadding;
                _moderatorAddressbookWarningPanel.SetBounds(innerPadding, contentTop, panelWidth, panelHeight);
                contentTop = _moderatorAddressbookWarningPanel.Bottom + ScaleLogical(8);
            }

            int hintTop = contentTop;
            int hintHeight = Math.Max(ScaleLogical(38), _moderatorGroup.ClientSize.Height - hintTop - innerPadding);
            _moderatorHintLabel.SetBounds(innerPadding, hintTop, Math.Max(ScaleLogical(120), _moderatorGroup.ClientSize.Width - (innerPadding * 2)), hintHeight);
            return _moderatorHintLabel.Bottom + innerPadding;
        }

        /**
         * Determines whether event conversations are supported (Nextcloud >= 31).
         */
        private static bool DetermineEventConversationSupport(AddinSettings defaults, out string versionText)
        {
            versionText = string.Empty;
            if (defaults == null)
            {
                return true;
            }

            versionText = defaults.LastKnownServerVersion ?? string.Empty;

            Version parsed;
            if (NextcloudVersionHelper.TryParse(versionText, out parsed))
            {
                versionText = parsed.ToString();
                return parsed.Major >= 31;
            }

            return true;
        }

        private void ApplyDefaults(AddinSettings defaults, string appointmentSubject)
        {
            string titleDefault = string.IsNullOrWhiteSpace(appointmentSubject) ? DefaultTitle : appointmentSubject.Trim();
            bool passwordDefault = defaults == null || defaults.TalkDefaultPasswordEnabled;
            bool addUsersDefault = defaults == null || defaults.TalkDefaultAddUsers;
            bool addGuestsDefault = defaults != null && defaults.TalkDefaultAddGuests;
            bool lobbyDefault = defaults == null || defaults.TalkDefaultLobbyEnabled;
            bool searchDefault = defaults == null || defaults.TalkDefaultSearchVisible;
            TalkRoomType roomTypeDefault = defaults != null ? defaults.TalkDefaultRoomType : TalkRoomType.EventConversation;

            if (IsPolicyActive())
            {
                bool policyBool;
                string policyString;

                policyString = _backendPolicyStatus.GetPolicyString("talk", "talk_title");
                if (!string.IsNullOrWhiteSpace(policyString))
                {
                    titleDefault = policyString;
                }

                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_set_password", out policyBool))
                {
                    passwordDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_users", out policyBool))
                {
                    addUsersDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_add_guests", out policyBool))
                {
                    addGuestsDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_lobby_active", out policyBool))
                {
                    lobbyDefault = policyBool;
                }
                if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_show_in_search", out policyBool))
                {
                    searchDefault = policyBool;
                }

                policyString = _backendPolicyStatus.GetPolicyString("talk", "talk_room_type");
                if (!string.IsNullOrWhiteSpace(policyString))
                {
                    roomTypeDefault = string.Equals(policyString.Trim(), "event", StringComparison.OrdinalIgnoreCase)
                        ? TalkRoomType.EventConversation
                        : TalkRoomType.StandardRoom;
                }
            }

            TalkTitle = titleDefault;
            _titleTextBox.Text = TalkTitle;

            int minLength = GetMinPasswordLength();
            _passwordToggleCheckBox.Checked = passwordDefault;
            TalkPassword = string.Empty;
            _passwordTextBox.Text = string.Empty;
            if (_passwordToggleCheckBox.Checked)
            {
                TalkPassword = GeneratePasswordValue(minLength);
                _passwordTextBox.Text = TalkPassword;
            }

            _addUsersCheckBox.Checked = addUsersDefault;
            _addGuestsCheckBox.Checked = addGuestsDefault;
            _lobbyCheckBox.Checked = lobbyDefault;
            _searchCheckBox.Checked = searchDefault;

            if (!_eventConversationsSupported && roomTypeDefault == TalkRoomType.EventConversation)
            {
                roomTypeDefault = TalkRoomType.StandardRoom;
            }

            SelectRoomType(roomTypeDefault);

            LobbyUntilStart = _lobbyCheckBox.Checked;
            SearchVisible = _searchCheckBox.Checked;
            AddUsers = _addUsersCheckBox.Checked;
            AddGuests = _addGuestsCheckBox.Checked;
            DelegateModeratorId = string.Empty;
            DelegateModeratorName = string.Empty;

            ApplyPolicyWarningUi();
            ApplyPolicyLockState();
            ApplySystemAddressbookLockState();
            UpdatePasswordState();
            UpdateRoomTypeTooltip();
            UpdateModeratorHint();
        }

        private bool IsPolicyActive()
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.PolicyActive;
        }

        private bool IsPolicyLocked(string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked("talk", key);
        }

        private bool IsPolicyGeneratePasswordEnabled()
        {
            if (!IsPolicyActive())
            {
                return true;
            }

            bool value;
            if (_backendPolicyStatus.TryGetPolicyBool("talk", "talk_generate_password", out value))
            {
                return value;
            }

            return true;
        }

        private void ApplyPolicyWarningUi()
        {
            bool visible = _backendPolicyStatus != null
                           && _backendPolicyStatus.WarningVisible
                           && !string.IsNullOrWhiteSpace(_backendPolicyStatus.WarningMessage);
            _policyWarningPanel.Visible = visible;
            _policyWarningTextLabel.Text = visible ? _backendPolicyStatus.WarningMessage : string.Empty;
        }

        private void ApplyPolicyLockState()
        {
            bool lockTitle = IsPolicyLocked("talk_title");
            bool lockRoomType = IsPolicyLocked("talk_room_type");
            bool lockPassword = IsPolicyLocked("talk_set_password");
            bool lockLobby = IsPolicyLocked("talk_lobby_active");
            bool lockSearch = IsPolicyLocked("talk_show_in_search");
            bool lockUsers = IsPolicyLocked("talk_add_users");
            bool lockGuests = IsPolicyLocked("talk_add_guests");

            _titleTextBox.Enabled = !lockTitle;
            _roomTypeComboBox.Enabled = !lockRoomType;
            _passwordToggleCheckBox.Enabled = !lockPassword;
            _lobbyCheckBox.Enabled = !lockLobby;
            _searchCheckBox.Enabled = !lockSearch;

            // Address book lock is applied in ApplySystemAddressbookLockState().
            if (!lockUsers)
            {
                _addUsersCheckBox.Enabled = true;
            }
            if (!lockGuests)
            {
                _addGuestsCheckBox.Enabled = true;
            }

            SetTooltipWithFallback(_titleTextBox, lockTitle ? Strings.PolicyAdminControlledTooltip : string.Empty, lockTitle, _titleLabel);
            SetTooltipWithFallback(_roomTypeComboBox, lockRoomType ? Strings.PolicyAdminControlledTooltip : _toolTip.GetToolTip(_roomTypeComboBox), lockRoomType, _roomTypeLabel);
            _disabledTooltipHints.Apply(
                _passwordToggleCheckBox,
                lockPassword ? Strings.PolicyAdminControlledTooltip : string.Empty,
                lockPassword,
                _passwordGenerateButton,
                _passwordLabel,
                _passwordTextBox,
                _passwordGenerateButton);
            SetTooltipWithFallback(_lobbyCheckBox, lockLobby ? Strings.PolicyAdminControlledTooltip : Strings.TooltipLobby, lockLobby);
            SetTooltipWithFallback(_searchCheckBox, lockSearch ? Strings.PolicyAdminControlledTooltip : Strings.TooltipSearchVisible, lockSearch);

            if (lockUsers)
            {
                SetTooltipWithFallback(_addUsersCheckBox, Strings.PolicyAdminControlledTooltip, true);
            }
            if (lockGuests)
            {
                SetTooltipWithFallback(_addGuestsCheckBox, Strings.PolicyAdminControlledTooltip, true);
            }
        }

        private void ApplySystemAddressbookLockState()
        {
            bool lockActive = !_systemAddressbookAvailable;
            bool usersPolicyLocked = IsPolicyLocked("talk_add_users");
            bool guestsPolicyLocked = IsPolicyLocked("talk_add_guests");
            string lockDetail = lockActive
                ? (!string.IsNullOrWhiteSpace(_systemAddressbookError) ? _systemAddressbookError : Strings.TalkSystemAddressbookRequiredMessage)
                : string.Empty;

            if (lockActive)
            {
                _addUsersCheckBox.Checked = false;
                _addGuestsCheckBox.Checked = false;
                _selectedModerator = null;
                _moderatorTextBox.Text = string.Empty;
                _moderatorListBox.Items.Clear();
                HideModeratorDropdown();
                _moderatorAvatarBox.Image = null;
                _moderatorSearchLockLogged = false;
            }

            _addUsersCheckBox.Enabled = !lockActive && !usersPolicyLocked;
            _addGuestsCheckBox.Enabled = !lockActive && !guestsPolicyLocked;
            _moderatorTextBox.Enabled = !lockActive;
            _moderatorClearButton.Enabled = !lockActive;

            _moderatorAddressbookWarningPanel.Visible = lockActive;
            _moderatorAddressbookWarningTextLabel.Text = lockActive
                ? Strings.TalkSystemAddressbookRequiredMessage
                : string.Empty;

            SetTooltipWithFallback(
                _addUsersCheckBox,
                lockActive ? Strings.TooltipAddUsersLocked : (usersPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddUsers),
                lockActive || usersPolicyLocked);
            SetTooltipWithFallback(
                _addGuestsCheckBox,
                lockActive ? Strings.TooltipAddGuestsLocked : (guestsPolicyLocked ? Strings.PolicyAdminControlledTooltip : Strings.TooltipAddGuests),
                lockActive || guestsPolicyLocked);
            SetTooltipWithFallback(_moderatorTextBox, lockActive ? Strings.TooltipModeratorLocked : Strings.TooltipModerator, lockActive, _moderatorHintLabel);
            SetTooltipWithFallback(_moderatorClearButton, lockActive ? Strings.TooltipModeratorLocked : Strings.TooltipModerator, lockActive, _moderatorHintLabel);

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "Talk wizard system address book lock state applied (locked=" + lockActive +
                ", available=" + _systemAddressbookAvailable +
                ", hasError=" + (!string.IsNullOrWhiteSpace(lockDetail)) +
                ", usersPolicyLocked=" + usersPolicyLocked +
                ", guestsPolicyLocked=" + guestsPolicyLocked + ").");
        }

        private static void OpenSystemAddressbookSetupGuide()
        {
            BrowserLauncher.OpenUrl(
                Strings.TalkSystemAddressbookAdminGuideUrl,
                LogCategories.Talk,
                "Failed to open system address book setup guide URL.");
        }

        private void UpdateRoomTypeTooltip()
        {
            if (IsPolicyLocked("talk_room_type"))
            {
                SetTooltipWithFallback(_roomTypeComboBox, Strings.PolicyAdminControlledTooltip, true, _roomTypeLabel);
                return;
            }

            var selected = _roomTypeComboBox.SelectedItem as RoomTypeOption;
            var roomType = selected != null ? selected.Value : TalkRoomType.EventConversation;
            SetTooltipWithFallback(
                _roomTypeComboBox,
                roomType == TalkRoomType.EventConversation ? Strings.TooltipRoomTypeEvent : Strings.TooltipRoomTypeStandard,
                false,
                _roomTypeLabel);
        }

        private void SetTooltipWithFallback(Control primary, string text, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, fallbackTargets);
        }

        private void SetTooltipWithFallback(Control primary, string text, bool showHint, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, showHint, fallbackTargets);
        }

        /**
         * Collects user input, validates the password, and exposes the selection to the caller.
         */
        private void OnOkButtonClick(object sender, EventArgs e)
        {
            TalkTitle = _titleTextBox.Text.Trim();

            bool passwordEnabled = _passwordToggleCheckBox.Checked;
            TalkPassword = passwordEnabled ? _passwordTextBox.Text.Trim() : string.Empty;
            int minLength = GetMinPasswordLength();
            if (passwordEnabled && TalkPassword.Length > 0 && TalkPassword.Length < minLength)
            {
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.TalkPasswordTooShort, minLength),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                _passwordTextBox.Focus();
                _passwordTextBox.SelectAll();
                return;
            }

            LobbyUntilStart = _lobbyCheckBox.Checked;
            SearchVisible = _searchCheckBox.Checked;
            AddUsers = _addUsersCheckBox.Checked;
            AddGuests = _addGuestsCheckBox.Checked;

            var selected = _roomTypeComboBox.SelectedItem as RoomTypeOption;
            SelectedRoomType = selected != null ? selected.Value : TalkRoomType.StandardRoom;

            string moderatorCandidate = _selectedModerator != null ? _selectedModerator.UserId : _moderatorTextBox.Text.Trim();
            DelegateModeratorId = string.IsNullOrWhiteSpace(moderatorCandidate) ? string.Empty : moderatorCandidate.Trim();
            DelegateModeratorName = _selectedModerator != null ? _selectedModerator.DisplayLabel : DelegateModeratorId;
        }

        private void UpdatePasswordState()
        {
            bool enabled = _passwordToggleCheckBox.Checked;
            bool lockPassword = IsPolicyLocked("talk_set_password");
            bool allowGenerate = IsPolicyGeneratePasswordEnabled();
            _passwordTextBox.Enabled = enabled && !lockPassword;
            _passwordGenerateButton.Enabled = enabled && allowGenerate;
            _disabledTooltipHints.Apply(
                _passwordGenerateButton,
                !enabled ? string.Empty : (allowGenerate ? string.Empty : Strings.PolicyAdminControlledTooltip),
                enabled && !allowGenerate,
                _passwordLabel);
        }

        private int GetMinPasswordLength()
        {
            if (_passwordPolicy != null && _passwordPolicy.MinLength > 0)
            {
                return _passwordPolicy.MinLength;
            }

            return DefaultMinPasswordLength;
        }

        private void GeneratePassword()
        {
            if (!_passwordToggleCheckBox.Checked)
            {
                return;
            }

            int minLength = GetMinPasswordLength();
            _passwordTextBox.Text = GeneratePasswordValue(minLength);
        }

        private string GeneratePasswordValue(int minLength)
        {
            string generated = null;

            try
            {
                if (_configuration != null && _passwordPolicy != null && _passwordPolicy.HasPolicy)
                {
                    var policyService = new PasswordPolicyService(_configuration);
                    generated = policyService.GeneratePassword(_passwordPolicy);
                }
            }
            catch (Exception ex)
            {
                generated = null;
                DiagnosticsLogger.LogException(LogCategories.Talk, "Password generation via server policy failed; falling back to local generator.", ex);
            }

            if (string.IsNullOrWhiteSpace(generated) || generated.Trim().Length < minLength)
            {
                generated = PasswordGenerator.GenerateLocalPassword(minLength);
            }

            return generated;
        }

        private void SelectRoomType(TalkRoomType type)
        {
            foreach (var item in _roomTypeComboBox.Items)
            {
                var option = item as RoomTypeOption;
                if (option != null && option.Value == type)
                {
                    _roomTypeComboBox.SelectedItem = option;
                    return;
                }
            }

            if (_roomTypeComboBox.Items.Count > 0)
            {
                _roomTypeComboBox.SelectedIndex = 0;
            }
        }

        private int ScaleLogical(int value)
        {
            int dpi = DeviceDpi > 0 ? DeviceDpi : 96;
            return (int)Math.Round(value * (dpi / 96f));
        }

        private sealed class RoomTypeOption
        {
            internal RoomTypeOption(TalkRoomType value, string label)
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

    }
}
