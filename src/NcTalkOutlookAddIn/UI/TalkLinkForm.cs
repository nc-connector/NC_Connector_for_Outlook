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
    internal sealed class TalkLinkForm : Form
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
        private readonly Label _moderatorHintLabel = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly Button _okButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private const int HeaderHeight = 48;
        private readonly bool _eventConversationsSupported;
        private readonly string _serverVersionHint;
        private readonly PasswordPolicyInfo _passwordPolicy;
        private readonly TalkServiceConfiguration _configuration;
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
            List<NextcloudUser> userDirectory,
            IfbAddressBookCache.SystemAddressbookStatus addressbookStatus,
            string appointmentSubject,
            DateTime startTime,
            DateTime endTime)
        {
            _eventConversationsSupported = DetermineEventConversationSupport(defaults, out _serverVersionHint);
            _passwordPolicy = passwordPolicy;
            _configuration = configuration;
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

            _addUsersCheckBox.Text = Strings.TalkAddUsersCheck;
            _addUsersCheckBox.AutoSize = true;
            _settingsGroup.Controls.Add(_addUsersCheckBox);

            _addGuestsCheckBox.Text = Strings.TalkAddGuestsCheck;
            _addGuestsCheckBox.AutoSize = true;
            _settingsGroup.Controls.Add(_addGuestsCheckBox);

            _lobbyCheckBox.Text = Strings.TalkLobbyCheck;
            _lobbyCheckBox.AutoSize = true;
            _settingsGroup.Controls.Add(_lobbyCheckBox);

            _searchCheckBox.Text = Strings.TalkSearchCheck;
            _searchCheckBox.AutoSize = true;
            _settingsGroup.Controls.Add(_searchCheckBox);

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
            TalkTitle = string.IsNullOrWhiteSpace(appointmentSubject) ? DefaultTitle : appointmentSubject.Trim();
            _titleTextBox.Text = TalkTitle;

            int minLength = GetMinPasswordLength();
            _passwordToggleCheckBox.Checked = defaults == null || defaults.TalkDefaultPasswordEnabled;
            TalkPassword = string.Empty;
            _passwordTextBox.Text = string.Empty;
            if (_passwordToggleCheckBox.Checked)
            {
                TalkPassword = GeneratePasswordValue(minLength);
                _passwordTextBox.Text = TalkPassword;
            }

            _addUsersCheckBox.Checked = defaults == null || defaults.TalkDefaultAddUsers;
            _addGuestsCheckBox.Checked = defaults != null && defaults.TalkDefaultAddGuests;
            _lobbyCheckBox.Checked = defaults == null || defaults.TalkDefaultLobbyEnabled;
            _searchCheckBox.Checked = defaults == null || defaults.TalkDefaultSearchVisible;

            var preferred = defaults != null ? defaults.TalkDefaultRoomType : TalkRoomType.EventConversation;
            if (!_eventConversationsSupported && preferred == TalkRoomType.EventConversation)
            {
                preferred = TalkRoomType.StandardRoom;
            }

            SelectRoomType(preferred);

            LobbyUntilStart = _lobbyCheckBox.Checked;
            SearchVisible = _searchCheckBox.Checked;
            AddUsers = _addUsersCheckBox.Checked;
            AddGuests = _addGuestsCheckBox.Checked;
            DelegateModeratorId = string.Empty;
            DelegateModeratorName = string.Empty;

            UpdatePasswordState();
            UpdateRoomTypeTooltip();
            ApplySystemAddressbookLockState();
            UpdateModeratorHint();
        }

        private void ApplySystemAddressbookLockState()
        {
            bool lockActive = !_systemAddressbookAvailable;
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

            _addUsersCheckBox.Enabled = !lockActive;
            _addGuestsCheckBox.Enabled = !lockActive;
            _moderatorTextBox.Enabled = !lockActive;
            _moderatorClearButton.Enabled = !lockActive;

            _moderatorAddressbookWarningPanel.Visible = lockActive;
            _moderatorAddressbookWarningTextLabel.Text = lockActive
                ? Strings.TalkSystemAddressbookRequiredMessage
                : string.Empty;

            _toolTip.SetToolTip(_addUsersCheckBox, lockActive ? Strings.TooltipAddUsersLocked : Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_addGuestsCheckBox, lockActive ? Strings.TooltipAddGuestsLocked : Strings.TooltipAddGuests);
            _toolTip.SetToolTip(_moderatorTextBox, lockActive ? Strings.TooltipModeratorLocked : Strings.TooltipModerator);
            _toolTip.SetToolTip(_moderatorClearButton, lockActive ? Strings.TooltipModeratorLocked : Strings.TooltipModerator);

            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "Talk wizard system address book lock state applied (locked=" + lockActive +
                ", available=" + _systemAddressbookAvailable +
                ", hasError=" + (!string.IsNullOrWhiteSpace(lockDetail)) + ").");
        }

        private static void OpenSystemAddressbookSetupGuide()
        {
            string url = Strings.TalkSystemAddressbookAdminGuideUrl;
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
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to open system address book setup guide URL.", ex);
            }
        }

        private void UpdateRoomTypeTooltip()
        {
            var selected = _roomTypeComboBox.SelectedItem as RoomTypeOption;
            var roomType = selected != null ? selected.Value : TalkRoomType.EventConversation;
            _toolTip.SetToolTip(
                _roomTypeComboBox,
                roomType == TalkRoomType.EventConversation ? Strings.TooltipRoomTypeEvent : Strings.TooltipRoomTypeStandard);
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
            _passwordTextBox.Enabled = enabled;
            _passwordGenerateButton.Enabled = enabled;
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
                generated = GenerateLocalPassword(minLength);
            }

            return generated;
        }

        private static string GenerateLocalPassword(int minLength)
        {
            const string alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
            int length = Math.Max(8, minLength);
            var chars = new char[length];
            using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
            {
                byte[] data = new byte[4];
                for (int i = 0; i < chars.Length; i++)
                {
                    rng.GetBytes(data);
                    int index = (int)(BitConverter.ToUInt32(data, 0) % alphabet.Length);
                    chars[i] = alphabet[index];
                }
            }

            return new string(chars);
        }

        private void UpdateModeratorResults()
        {
            if (!_systemAddressbookAvailable)
            {
                if (!_moderatorSearchLockLogged)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Talk,
                        "Talk wizard moderator search skipped because system address book is locked.");
                    _moderatorSearchLockLogged = true;
                }

                _moderatorListBox.Items.Clear();
                HideModeratorDropdown();
                UpdateModeratorHint();
                return;
            }

            _moderatorSearchLockLogged = false;

            if (_userDirectory.Count == 0)
            {
                HideModeratorDropdown();
                UpdateModeratorHint();
                return;
            }

            string term = (_moderatorTextBox.Text ?? string.Empty).Trim();

            _moderatorListBox.BeginUpdate();
            _moderatorListBox.Items.Clear();
            int added = 0;
            foreach (var user in _userDirectory)
            {
                if (user == null)
                {
                    continue;
                }

                if (term.Length == 0 ||
                    user.UserId.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (!string.IsNullOrEmpty(user.Email) && user.Email.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    _moderatorListBox.Items.Add(user);
                    added++;
                }
            }
            _moderatorListBox.EndUpdate();

            _moderatorHintLabel.Visible = added == 0;
            _moderatorHintLabel.Text = added == 0 ? Strings.TalkModeratorNoMatches : Strings.TalkModeratorHint;
            if (added > 0)
            {
                _moderatorHintLabel.Visible = false;
                ShowModeratorDropdown();
            }
            else
            {
                HideModeratorDropdown();
            }
        }

        private void SelectModeratorFromList()
        {
            var selected = _moderatorListBox.SelectedItem as NextcloudUser;
            if (selected == null)
            {
                return;
            }

            _selectedModerator = selected;
            _moderatorTextBox.Text = selected.UserId;
            HideModeratorDropdown();
            SetModeratorAvatar(selected);
            UpdateModeratorHint();
        }

        private void ClearModerator()
        {
            _selectedModerator = null;
            _moderatorTextBox.Text = string.Empty;
            _moderatorListBox.Items.Clear();
            HideModeratorDropdown();
            _moderatorAvatarBox.Image = null;
            UpdateModeratorHint();
        }

        private void ShowModeratorDropdown()
        {
            if (_moderatorListBox.Items.Count == 0)
            {
                HideModeratorDropdown();
                return;
            }

            if (_moderatorListBox.SelectedIndex < 0)
            {
                _moderatorListBox.SelectedIndex = 0;
            }

            PositionModeratorDropdown();
            _moderatorListBox.BringToFront();
            _moderatorListBox.Visible = true;
        }

        private void HideModeratorDropdown()
        {
            _moderatorListBox.Visible = false;
        }

        private void PositionModeratorDropdown()
        {
            if (!_systemAddressbookAvailable)
            {
                HideModeratorDropdown();
                return;
            }

            int width = Math.Max(0, _moderatorGroup.Width - 24);
            int rows = Math.Max(1, Math.Min(ModeratorDropdownMaxRows, _moderatorListBox.Items.Count));
            int desiredHeight = (rows * _moderatorListBox.ItemHeight) + 4;

            Point inputClient = PointToClient(_moderatorTextBox.PointToScreen(Point.Empty));
            int x = _moderatorGroup.Left + 12;
            int y = inputClient.Y - desiredHeight - ModeratorDropdownMargin;

            int minY = _headerPanel.Bottom + ModeratorDropdownMargin;
            if (y < minY)
            {
                y = inputClient.Y + _moderatorTextBox.Height + ModeratorDropdownMargin;
            }

            int maxBottom = _okButton != null ? _okButton.Top - ModeratorDropdownMargin : ClientSize.Height;
            if (y + desiredHeight > maxBottom)
            {
                desiredHeight = Math.Max(_moderatorListBox.ItemHeight + 4, maxBottom - y);
            }

            _moderatorListBox.Location = new Point(x, y);
            _moderatorListBox.Size = new Size(width, Math.Max(_moderatorListBox.ItemHeight + 4, desiredHeight));
        }

        private void SetModeratorAvatar(NextcloudUser user)
        {
            if (user == null)
            {
                _moderatorAvatarBox.Image = null;
                return;
            }

            EnsureAvatarLoaded(user.UserId);
            _moderatorAvatarBox.Image = GetCachedAvatar(user.UserId);
        }

        private void OnModeratorDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _moderatorListBox.Items.Count)
            {
                return;
            }

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color backColor;
            Color foreColor;
            if (_themePalette.IsDark)
            {
                backColor = isSelected ? _themePalette.SelectionBackground : _moderatorListBox.BackColor;
                foreColor = isSelected ? _themePalette.SelectionText : _moderatorListBox.ForeColor;
            }
            else
            {
                backColor = isSelected ? SystemColors.Highlight : e.BackColor;
                foreColor = isSelected ? SystemColors.HighlightText : e.ForeColor;
            }

            using (var back = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(back, e.Bounds);
            }

            var user = _moderatorListBox.Items[e.Index] as NextcloudUser;
            string label = user != null ? user.DisplayLabel : (_moderatorListBox.Items[e.Index] != null ? _moderatorListBox.Items[e.Index].ToString() : string.Empty);

            const int avatarSize = 28;
            int avatarX = e.Bounds.Left + 4;
            int avatarY = e.Bounds.Top + ((e.Bounds.Height - avatarSize) / 2);
            var avatarBounds = new Rectangle(avatarX, avatarY, avatarSize, avatarSize);

            Image avatar = user != null ? GetCachedAvatar(user.UserId) : null;
            if (avatar == null && user != null)
            {
                EnsureAvatarLoaded(user.UserId);
                avatar = GetCachedAvatar(user.UserId);
            }
            if (avatar != null)
            {
                e.Graphics.DrawImage(avatar, avatarBounds);
            }
            else
            {
                DrawAvatarPlaceholder(e.Graphics, avatarBounds, user != null ? user.UserId : null);
            }

            int textX = avatarBounds.Right + 8;
            var textBounds = new Rectangle(textX, e.Bounds.Top, e.Bounds.Right - textX - 4, e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, label, e.Font, textBounds, foreColor, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            e.DrawFocusRectangle();
        }

        private Image GetCachedAvatar(string userId)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                return null;
            }

            lock (_avatarLock)
            {
                Image cached;
                if (_avatarCache.TryGetValue(userId, out cached))
                {
                    return cached;
                }
            }

            return null;
        }

        private void EnsureAvatarLoaded(string userId)
        {
            if (string.IsNullOrWhiteSpace(userId) || _configuration == null || !_configuration.IsComplete())
            {
                return;
            }

            lock (_avatarLock)
            {
                if (_avatarCache.ContainsKey(userId) || _avatarLoading.Contains(userId))
                {
                    return;
                }

                _avatarLoading.Add(userId);
            }

            Task.Run(() =>
            {
                Image fetched = null;
                try
                {
                    fetched = FetchAvatar(userId);
                }
                catch (Exception ex)
                {
                    fetched = null;
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to fetch avatar for user '" + userId + "'.", ex);
                }

                lock (_avatarLock)
                {
                    _avatarLoading.Remove(userId);
                    if (fetched != null)
                    {
                        Image existing;
                        if (_avatarCache.TryGetValue(userId, out existing))
                        {
                            fetched.Dispose();
                        }
                        else
                        {
                            _avatarCache[userId] = fetched;
                        }
                    }
                }

                try
                {
                    BeginInvoke((Action)(() =>
                    {
                        _moderatorListBox.Invalidate();
                        if (_selectedModerator != null && string.Equals(_selectedModerator.UserId, userId, StringComparison.OrdinalIgnoreCase))
                        {
                            _moderatorAvatarBox.Image = GetCachedAvatar(userId);
                        }
                    }));
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to update UI after avatar fetch.", ex);
                }
            });
        }

        private Image FetchAvatar(string userId)
        {
            string baseUrl = _configuration != null ? _configuration.GetNormalizedBaseUrl() : string.Empty;
            if (string.IsNullOrEmpty(baseUrl))
            {
                return null;
            }

            string url = baseUrl + "/index.php/avatar/" + Uri.EscapeDataString(userId) + "/64";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Accept = "image/png,image/*;q=0.8,*/*;q=0.5";
            request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
            request.Timeout = 20000;

            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        return null;
                    }

                    using (Stream stream = response.GetResponseStream())
                    {
                        if (stream == null)
                        {
                            return null;
                        }

                        using (var image = Image.FromStream(stream))
                        {
                            return new Bitmap(image);
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                var response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    response.Close();
                }
                DiagnosticsLogger.LogException(LogCategories.Talk, "Avatar request failed for user '" + userId + "'.", ex);
                return null;
            }
        }

        private void DrawAvatarPlaceholder(Graphics graphics, Rectangle bounds, string userId)
        {
            if (graphics == null)
            {
                return;
            }

            using (var fill = new SolidBrush(_themePalette.AvatarPlaceholderFill))
            using (var border = new Pen(_themePalette.AvatarPlaceholderBorder))
            {
                graphics.FillEllipse(fill, bounds);
                graphics.DrawEllipse(border, bounds);
            }

            string initial = string.IsNullOrEmpty(userId) ? "?" : userId.Trim().Substring(0, 1).ToUpperInvariant();
            using (var textBrush = new SolidBrush(_themePalette.AvatarPlaceholderText))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                graphics.DrawString(initial, SystemFonts.DefaultFont, textBrush, bounds, format);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                lock (_avatarLock)
                {
                    foreach (var entry in _avatarCache)
                    {
                        try { entry.Value.Dispose(); }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to dispose cached avatar image for user '" + entry.Key + "'.", ex);
                        }
                    }
                    _avatarCache.Clear();
                    _avatarLoading.Clear();
                }
            }

            base.Dispose(disposing);
        }

        private void UpdateModeratorHint()
        {
            _moderatorHintLabel.Visible = true;
            if (!_systemAddressbookAvailable)
            {
                _moderatorHintLabel.Text = Strings.TalkSystemAddressbookRequiredMessage;
            }
            else if (_userDirectory.Count == 0)
            {
                _moderatorHintLabel.Text = Strings.TalkModeratorHintNoDirectory;
            }
            else
            {
                _moderatorHintLabel.Text = Strings.TalkModeratorHint;
            }
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
