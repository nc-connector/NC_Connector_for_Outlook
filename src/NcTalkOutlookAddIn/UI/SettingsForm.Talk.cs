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

            _talkDeleteRoomOnEventDeleteCheckBox.Location = new Point(innerPadding, y);
            y = _talkDeleteRoomOnEventDeleteCheckBox.Bottom + rowGap;

            int addressbookWarningHeight = WarningPanelUiHelper.Layout(
                _talkAddressbookWarningPanel,
                _talkAddressbookWarningTitleLabel,
                _talkAddressbookWarningTextLabel,
                _talkAddressbookWarningLinkLabel,
                innerPadding,
                y,
                Math.Max(ScaleLogical(160), contentWidth),
                ScaleLogical(8),
                ScaleLogical(120),
                ScaleLogical(4),
                ScaleLogical(6));
            if (addressbookWarningHeight > 0)
            {
                y = _talkAddressbookWarningPanel.Bottom + innerPadding;
            }
            else
            {
                y += innerPadding;
            }

            _talkDefaultsGroup.Height = Math.Max(ScaleLogical(200), y);
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
            TalkRoomTypeComboHelper.Populate(
                _talkDefaultRoomTypeCombo,
                Strings.TalkEventRadio,
                Strings.TalkStandardRadio);
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

            _talkDeleteRoomOnEventDeleteCheckBox.Text = Strings.TalkDeleteRoomOnEventDeleteCheck;
            _talkDeleteRoomOnEventDeleteCheckBox.AutoSize = true;
            _talkDeleteRoomOnEventDeleteCheckBox.Location = new Point(12, 180);
            _talkDefaultsGroup.Controls.Add(_talkDeleteRoomOnEventDeleteCheckBox);

            WarningPanelUiHelper.Initialize(
                _talkAddressbookWarningPanel,
                _talkAddressbookWarningTitleLabel,
                _talkAddressbookWarningTextLabel,
                _talkAddressbookWarningLinkLabel,
                "\u26a0 " + Strings.TalkSystemAddressbookRequiredShort,
                Strings.TalkSystemAddressbookRequiredMessage,
                Strings.TalkSystemAddressbookAdminLinkLabel);
            _talkDefaultsGroup.Controls.Add(_talkAddressbookWarningPanel);
            _talkAddressbookWarningLinkLabel.LinkClicked += (s, e) =>
                BrowserLauncher.OpenUrl(
                    Strings.TalkSystemAddressbookAdminGuideUrl,
                    LogCategories.Core,
                    "Failed to open system address book admin guide URL.");

            _toolTip.SetToolTip(_talkDefaultAddUsersCheckBox, Strings.TooltipAddUsers);
            _toolTip.SetToolTip(_talkDefaultAddGuestsCheckBox, Strings.TooltipAddGuests);
            _toolTip.SetToolTip(_talkDefaultLobbyCheckBox, Strings.TooltipLobby);
            _toolTip.SetToolTip(_talkDefaultSearchCheckBox, Strings.TooltipSearchVisible);
            _toolTip.SetToolTip(_talkDeleteRoomOnEventDeleteCheckBox, Strings.TooltipDeleteRoomOnEventDelete);
            UpdateTalkRoomTypeTooltip();
        }

        private void ApplyInitialTalkSystemAddressbookState()
        {
            ApplyTalkSystemAddressbookStatus(
                _initialAddressbookStatus,
                "settings_open");
        }

        private async Task RefreshTalkSystemAddressbookStateAsync(
            TalkServiceConfiguration configuration,
            bool forceRefresh,
            string trigger)
        {
            int generation = ++_addressbookRefreshGeneration;
            int cacheHours = ParseComboValue(_ifbCacheHoursCombo, 24);
            DiagnosticsLogger.Log(
                LogCategories.Talk,
                "System address book status check requested from settings (trigger=" + (trigger ?? "n/a") +
                ", forceRefresh=" + forceRefresh + ").");

            IfbAddressBookCache.SystemAddressbookStatus status =
                _addressBookCache != null
                    ? await Task.Run(
                        () => _addressBookCache.GetSystemAddressbookStatus(
                            configuration,
                            cacheHours,
                            forceRefresh))
                    : _initialAddressbookStatus;
            if (generation != _addressbookRefreshGeneration
                || IsDisposed
                || Disposing)
            {
                return;
            }
            ApplyTalkSystemAddressbookStatus(status, trigger);
        }

        private void ApplyTalkSystemAddressbookStatus(
            IfbAddressBookCache.SystemAddressbookStatus status,
            string trigger)
        {
            bool statusAvailable = status != null && status.Available;
            bool lockActive = !statusAvailable;
            string detail = lockActive ? Strings.TalkSystemAddressbookRequiredMessage : string.Empty;
            if (lockActive
                && status != null
                && !string.IsNullOrWhiteSpace(status.Error))
            {
                detail = status.Error;
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

    }
}
