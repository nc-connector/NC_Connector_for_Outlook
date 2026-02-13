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
    internal sealed class TalkLinkForm : Form
    {
        private static readonly string DefaultTitle = Strings.TalkDefaultTitle;
        private const int DefaultMinPasswordLength = 5;
        private const int ModeratorDropdownMaxRows = 6;
        private const int ModeratorDropdownMargin = 4;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

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
        private NextcloudUser _selectedModerator;
        private readonly Dictionary<string, Image> _avatarCache = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _avatarLoading = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly object _avatarLock = new object();

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
            string appointmentSubject,
            DateTime startTime,
            DateTime endTime)
        {
            _eventConversationsSupported = DetermineEventConversationSupport(defaults, out _serverVersionHint);
            _passwordPolicy = passwordPolicy;
            _configuration = configuration;
            _userDirectory = userDirectory ?? new List<NextcloudUser>();

            Text = Strings.TalkFormTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 520);
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeComponents();
            ApplyDefaults(defaults, appointmentSubject);

            UiThemeManager.ApplyToForm(this, _toolTip);
        }

        /**
         * Builds all dialog controls (title, password, options, buttons).
         */
        private void InitializeComponents()
        {
            int topOffset = _headerPanel.Bottom + 15;
            int labelX = 15;
            int inputX = 170;
            int inputWidth = ClientSize.Width - inputX - 25;

            var titleLabel = new Label
            {
                Text = Strings.TalkTitleLabel,
                Location = new Point(labelX, topOffset),
                AutoSize = true
            };

            _titleTextBox.Location = new Point(inputX, topOffset - 4);
            _titleTextBox.Width = inputWidth;

            var roomTypeLabel = new Label
            {
                Text = Strings.TalkRoomGroup,
                Location = new Point(labelX, topOffset + 40),
                AutoSize = true
            };

            _roomTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _roomTypeComboBox.Location = new Point(inputX, topOffset + 36);
            _roomTypeComboBox.Width = inputWidth;
            _roomTypeComboBox.Items.Add(new RoomTypeOption(TalkRoomType.EventConversation, Strings.TalkEventRadio));
            _roomTypeComboBox.Items.Add(new RoomTypeOption(TalkRoomType.StandardRoom, Strings.TalkStandardRadio));
            _roomTypeComboBox.SelectedIndexChanged += (s, e) => UpdateRoomTypeTooltip();

            _passwordToggleCheckBox.Text = Strings.TalkPasswordSetCheck;
            _passwordToggleCheckBox.AutoSize = true;
            _passwordToggleCheckBox.Location = new Point(18, topOffset + 80);
            _passwordToggleCheckBox.CheckedChanged += (s, e) => UpdatePasswordState();

            var passwordLabel = new Label
            {
                Text = Strings.TalkPasswordLabel,
                Location = new Point(labelX, topOffset + 112),
                AutoSize = true
            };

            _passwordTextBox.Location = new Point(inputX, topOffset + 108);
            _passwordTextBox.Width = inputWidth - 110;
            _passwordTextBox.UseSystemPasswordChar = false;

            _passwordGenerateButton.Text = Strings.TalkPasswordGenerate;
            _passwordGenerateButton.Width = 100;
            _passwordGenerateButton.Location = new Point(_passwordTextBox.Right + 10, _passwordTextBox.Top - 2);
            _passwordGenerateButton.Click += (s, e) => GeneratePassword();

            var settingsGroup = new GroupBox
            {
                Text = Strings.TalkSettingsGroup,
                Location = new Point(18, topOffset + 150),
                Size = new Size(ClientSize.Width - 36, 120)
            };

            _addUsersCheckBox.Text = Strings.TalkAddUsersCheck;
            _addUsersCheckBox.Location = new Point(12, 25);
            _addUsersCheckBox.AutoSize = true;
            settingsGroup.Controls.Add(_addUsersCheckBox);

            _addGuestsCheckBox.Text = Strings.TalkAddGuestsCheck;
            _addGuestsCheckBox.Location = new Point(12, 50);
            _addGuestsCheckBox.AutoSize = true;
            settingsGroup.Controls.Add(_addGuestsCheckBox);

            _lobbyCheckBox.Text = Strings.TalkLobbyCheck;
            _lobbyCheckBox.Location = new Point(12, 75);
            _lobbyCheckBox.AutoSize = true;
            settingsGroup.Controls.Add(_lobbyCheckBox);

            _searchCheckBox.Text = Strings.TalkSearchCheck;
            _searchCheckBox.Location = new Point(12, 100);
            _searchCheckBox.AutoSize = true;
            settingsGroup.Controls.Add(_searchCheckBox);

            _moderatorGroup.Text = Strings.TalkModeratorGroup;
            _moderatorGroup.Location = new Point(18, settingsGroup.Bottom + 12);
            _moderatorGroup.Size = new Size(ClientSize.Width - 36, 105);

            _moderatorAvatarBox.Location = new Point(12, 22);
            _moderatorAvatarBox.Size = new Size(32, 32);
            _moderatorAvatarBox.SizeMode = PictureBoxSizeMode.Zoom;
            _moderatorAvatarBox.BackColor = Color.Transparent;
            _moderatorGroup.Controls.Add(_moderatorAvatarBox);

            _moderatorTextBox.Location = new Point(_moderatorAvatarBox.Right + 8, 24);
            _moderatorTextBox.Width = _moderatorGroup.Width - 120 - _moderatorAvatarBox.Width - 8;
            _moderatorTextBox.TextChanged += (s, e) => UpdateModeratorResults();
            _moderatorTextBox.Enter += (s, e) => UpdateModeratorResults();
            _moderatorTextBox.MouseDown += (s, e) => UpdateModeratorResults();
            _moderatorGroup.Controls.Add(_moderatorTextBox);

            _moderatorClearButton.Text = Strings.TalkModeratorClear;
            _moderatorClearButton.Location = new Point(_moderatorTextBox.Right + 8, 22);
            _moderatorClearButton.Width = 90;
            _moderatorClearButton.Click += (s, e) => ClearModerator();
            _moderatorGroup.Controls.Add(_moderatorClearButton);

            _moderatorListBox.Location = new Point(12, 54);
            _moderatorListBox.Size = new Size(_moderatorGroup.Width - 24, 44);
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
            _moderatorHintLabel.Location = new Point(12, 54);
            _moderatorHintLabel.Size = new Size(_moderatorGroup.Width - 24, 44);
            _moderatorHintLabel.ForeColor = Color.DimGray;
            _moderatorHintLabel.Click += (s, e) =>
            {
                _moderatorTextBox.Focus();
                UpdateModeratorResults();
            };
            _moderatorGroup.Controls.Add(_moderatorHintLabel);

            if (!_eventConversationsSupported)
            {
                var versionInfo = string.IsNullOrEmpty(_serverVersionHint) ? Strings.TalkVersionUnknown : _serverVersionHint;
                var hintLabel = new Label
                {
                    Text = string.Format(Strings.TalkEventHint, versionInfo),
                    Location = new Point(inputX, topOffset + 60),
                    Size = new Size(inputWidth, 36),
                    ForeColor = Color.DimGray
                };
                Controls.Add(hintLabel);
            }

            _okButton.Text = Strings.DialogOk;
            _okButton.Location = new Point(ClientSize.Width - 200, ClientSize.Height - 42);
            _okButton.Width = 90;
            _okButton.DialogResult = DialogResult.OK;
            _okButton.Click += OnOkButtonClick;

            _cancelButton.Text = Strings.DialogCancel;
            _cancelButton.Location = new Point(ClientSize.Width - 105, ClientSize.Height - 42);
            _cancelButton.Width = 90;
            _cancelButton.DialogResult = DialogResult.Cancel;

            Controls.Add(titleLabel);
            Controls.Add(_titleTextBox);
            Controls.Add(roomTypeLabel);
            Controls.Add(_roomTypeComboBox);
            Controls.Add(_passwordToggleCheckBox);
            Controls.Add(passwordLabel);
            Controls.Add(_passwordTextBox);
            Controls.Add(_passwordGenerateButton);
            Controls.Add(settingsGroup);
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
            UpdateRoomTypeTooltip();

            Controls.Add(_moderatorListBox);
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = HeaderHeight;
            _headerPanel.Dock = DockStyle.Top;
            Controls.Add(_headerPanel);
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
                TalkPassword = GenerateLocalPassword(minLength);
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
            UpdateModeratorHint();
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

            if (string.IsNullOrWhiteSpace(generated))
            {
                generated = GenerateLocalPassword(minLength);
            }

            _passwordTextBox.Text = generated;
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
            if (_userDirectory.Count == 0)
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
