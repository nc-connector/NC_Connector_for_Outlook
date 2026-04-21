/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * Moderator search dropdown and avatar lifecycle handling.
     */
    internal sealed partial class TalkLinkForm
    {
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
            {                if (user == null)
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
            var selected = _moderatorListBox.SelectedItem as NextcloudUser;            if (selected == null)
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
        {            if (user == null)
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

            Image avatar = user != null ? GetCachedAvatar(user.UserId) : null;            if (avatar == null && user != null)
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
        {            if (string.IsNullOrWhiteSpace(userId) || _configuration == null || !_configuration.IsComplete())
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
                    _avatarLoading.Remove(userId);                    if (fetched != null)
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
                        _moderatorListBox.Invalidate();                        if (_selectedModerator != null && string.Equals(_selectedModerator.UserId, userId, StringComparison.OrdinalIgnoreCase))
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
            var httpClient = new Services.NcHttpClient(_configuration);
            Services.NcHttpResponse response = httpClient.Send(new Services.NcHttpRequestOptions
            {
                Method = "GET",
                Url = url,
                Accept = "image/png,image/*;q=0.8,*/*;q=0.5",
                TimeoutMs = 20000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                ReadResponseAsBytes = true
            });

            if (!response.HasHttpResponse || response.StatusCode != HttpStatusCode.OK)
            {                if (response.TransportException != null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Avatar request failed for user '" + userId + "'.", response.TransportException);
                }
                return null;
            }
            if (response.ResponseBytes == null || response.ResponseBytes.Length == 0)
            {
                return null;
            }

            using (var stream = new MemoryStream(response.ResponseBytes))
            using (var image = Image.FromStream(stream))
            {
                return new Bitmap(image);
            }
        }

        private void DrawAvatarPlaceholder(Graphics graphics, Rectangle bounds, string userId)
        {            if (graphics == null)
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
    }
}

