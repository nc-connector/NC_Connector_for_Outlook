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
        private void ApplyAdvancedTabLayout()
        {
            int left = ScaleLogical(24);
            int labelToComboGap = ScaleLogical(16);
            int comboLeft = left + _ifbCacheHoursLabel.PreferredSize.Width + labelToComboGap;
            int rightMargin = ScaleLogical(24);
            int rowTop = ScaleLogical(24);

            int ifbComboHeight = Math.Max(_ifbCacheHoursCombo.Height, _ifbCacheHoursCombo.PreferredHeight + ScaleLogical(2));
            _ifbCacheHoursLabel.Location = new Point(left, rowTop);
            _ifbCacheHoursCombo.SetBounds(comboLeft, rowTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbCacheHoursCombo.Width), ifbComboHeight);

            int groupTop = Math.Max(_ifbCacheHoursLabel.Bottom, _ifbCacheHoursCombo.Bottom) + ScaleLogical(20);
            int groupWidth = Math.Max(ScaleLogical(320), _advancedTab.ClientSize.Width - left - rightMargin);
            _updateSettingsGroup.SetBounds(left, groupTop, groupWidth, ScaleLogical(286));
            int updateInnerWidth = Math.Max(ScaleLogical(220), _updateSettingsGroup.ClientSize.Width - ScaleLogical(24));
            _updateNotifyCheckBox.Location = new Point(ScaleLogical(12), ScaleLogical(24));
            _updateInstalledVersionLabel.SetBounds(ScaleLogical(12), ScaleLogical(54), updateInnerWidth, ScaleLogical(20));
            _updateLatestVersionLabel.SetBounds(ScaleLogical(12), ScaleLogical(78), updateInnerWidth, ScaleLogical(20));
            _updateLastCheckedLabel.SetBounds(ScaleLogical(12), ScaleLogical(102), updateInnerWidth, ScaleLogical(20));
            _updateCheckButton.SetBounds(ScaleLogical(12), ScaleLogical(130), ScaleLogical(120), ScaleLogical(28));
            _updateDownloadLink.SetBounds(_updateCheckButton.Right + ScaleLogical(16), _updateCheckButton.Top + ScaleLogical(6), Math.Max(ScaleLogical(140), updateInnerWidth - _updateCheckButton.Width - ScaleLogical(28)), ScaleLogical(22));
            _updateChangelogLabel.Location = new Point(ScaleLogical(12), ScaleLogical(166));
            _updateChangelogTextBox.SetBounds(ScaleLogical(12), ScaleLogical(186), updateInnerWidth, ScaleLogical(86));

            int tlsTop = _updateSettingsGroup.Bottom + ScaleLogical(14);
            _tlsSettingsGroup.SetBounds(left, tlsTop, groupWidth, ScaleLogical(134));
            _tlsHintLabel.MaximumSize = new Size(Math.Max(ScaleLogical(200), _tlsSettingsGroup.ClientSize.Width - ScaleLogical(20)), 0);
            _tlsHintLabel.AutoSize = true;
            _tlsSettingsGroup.Height = Math.Max(ScaleLogical(134), _tlsHintLabel.Bottom + ScaleLogical(12));
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

            _updateSettingsGroup.Text = Strings.UpdateSettingsHeading;
            _updateSettingsGroup.Location = new Point(24, 70);
            _updateSettingsGroup.Size = new Size(520, 286);
            _advancedTab.Controls.Add(_updateSettingsGroup);

            _updateNotifyCheckBox.Text = Strings.UpdateNotifyLabel;
            _updateNotifyCheckBox.AutoSize = true;
            _updateNotifyCheckBox.Location = new Point(12, 24);
            _updateSettingsGroup.Controls.Add(_updateNotifyCheckBox);

            _updateInstalledVersionLabel.AutoSize = false;
            _updateInstalledVersionLabel.Location = new Point(12, 54);
            _updateInstalledVersionLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateInstalledVersionLabel);

            _updateLatestVersionLabel.AutoSize = false;
            _updateLatestVersionLabel.Location = new Point(12, 78);
            _updateLatestVersionLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateLatestVersionLabel);

            _updateLastCheckedLabel.AutoSize = false;
            _updateLastCheckedLabel.Location = new Point(12, 102);
            _updateLastCheckedLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateLastCheckedLabel);

            _updateCheckButton.Text = Strings.UpdateCheckNowButton;
            _updateCheckButton.Location = new Point(12, 130);
            _updateCheckButton.Size = new Size(120, 28);
            _updateCheckButton.Click += OnUpdateCheckButtonClick;
            _updateSettingsGroup.Controls.Add(_updateCheckButton);

            _updateDownloadLink.Text = Strings.UpdateDownloadLink;
            _updateDownloadLink.Location = new Point(148, 136);
            _updateDownloadLink.AutoSize = false;
            _updateDownloadLink.Size = new Size(320, 22);
            _updateDownloadLink.LinkClicked += OnUpdateDownloadLinkClicked;
            _updateSettingsGroup.Controls.Add(_updateDownloadLink);

            _updateChangelogLabel.Text = Strings.UpdateChangelogHeading;
            _updateChangelogLabel.Location = new Point(12, 166);
            _updateChangelogLabel.AutoSize = true;
            _updateSettingsGroup.Controls.Add(_updateChangelogLabel);

            _updateChangelogTextBox.ReadOnly = true;
            _updateChangelogTextBox.Multiline = true;
            _updateChangelogTextBox.ScrollBars = ScrollBars.Vertical;
            _updateChangelogTextBox.Location = new Point(12, 186);
            _updateChangelogTextBox.Size = new Size(480, 86);
            _updateSettingsGroup.Controls.Add(_updateChangelogTextBox);

            _tlsSettingsGroup.Text = Strings.AdvancedTlsHeading;
            _tlsSettingsGroup.Location = new Point(24, 370);
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

        private void UpdateUpdateCheckSection()
        {
            UpdateCheckResult result = UpdateCheckService.BuildCachedResult(Result);
            string installedVersion = AddinVersionInfo.GetVersion();
            if (string.IsNullOrWhiteSpace(installedVersion))
            {
                installedVersion = Strings.TalkVersionUnknown;
            }
            string latestVersion = string.IsNullOrWhiteSpace(result.LatestVersion)
                ? Strings.UpdateNotChecked
                : result.LatestVersion.Trim();
            string lastChecked = FormatUpdateCheckedAt(Result != null ? Result.UpdateLastCheckedAtUtc : string.Empty);

            _updateInstalledVersionLabel.Text = string.Format(Strings.UpdateInstalledVersionFormat, installedVersion);
            _updateLatestVersionLabel.Text = string.Format(Strings.UpdateLatestVersionFormat, latestVersion);
            _updateLastCheckedLabel.Text = string.Format(Strings.UpdateLastCheckedFormat, lastChecked);

            _updateOpenUrl = result.UpdateAvailable ? UpdateCheckService.GetPreferredOpenUrl(result) : string.Empty;
            _updateDownloadLink.Text = string.IsNullOrWhiteSpace(_updateOpenUrl)
                ? Strings.UpdateNoDownloadLink
                : Strings.UpdateDownloadLink;
            _updateDownloadLink.Enabled = !_isBusy && !string.IsNullOrWhiteSpace(_updateOpenUrl);

            _updateChangelogTextBox.Text = string.IsNullOrWhiteSpace(result.ChangelogText)
                ? Strings.UpdateChangelogEmpty
                : result.ChangelogText;
        }

        private static string FormatUpdateCheckedAt(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Strings.UpdateNotChecked;
            }

            DateTime parsed;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out parsed))
            {
                return parsed.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
            }

            return Strings.UpdateNotChecked;
        }

        private async void OnUpdateCheckButtonClick(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            SetBusy(true);
            SetStatus(Strings.UpdateCheckRunning, false);
            try
            {
                var service = new UpdateCheckService();
                UpdateCheckResult result = await service.CheckAsync(Result, true);
                SetStatus(result != null && result.UpdateAvailable ? string.Format(Strings.UpdateAvailableStatusFormat, result.LatestVersion) : Strings.UpdateNoUpdateAvailable, false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Manual update check failed.", ex);
                SetStatus(string.Format(Strings.UpdateCheckFailedFormat, ex.Message), true);
            }
            finally
            {
                SetBusy(false);
                UpdateUpdateCheckSection();
            }
        }

        private void OnUpdateDownloadLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_updateOpenUrl))
            {
                return;
            }

            BrowserLauncher.OpenUrl(
                _updateOpenUrl,
                LogCategories.Core,
                "Failed to open update download URL.");
        }

        private void UpdateTlsOptionsState()
        {
            bool managed = Result != null && Result.HasManagedTransportTls;
            bool useSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
            bool allowCustom = !managed && !useSystemDefault && !_isBusy;

            _tlsUseSystemDefaultCheckBox.Enabled = !managed && !_isBusy;
            _tlsEnable12CheckBox.Enabled = allowCustom;
            _tlsEnable13CheckBox.Enabled = allowCustom;
            _tlsHintLabel.Text = managed ? Strings.AdvancedTlsManagedHint : Strings.AdvancedTlsHint;
            foreach (Control control in new Control[] { _tlsUseSystemDefaultCheckBox, _tlsEnable12CheckBox, _tlsEnable13CheckBox })
            {
                _disabledTooltipHints.Apply(control, managed ? Strings.AdvancedTlsManagedHint : string.Empty, managed);
            }
            _toolTip.SetToolTip(_tlsHintLabel, managed ? Strings.AdvancedTlsManagedHint : string.Empty);
        }

        private void OnTlsSelectionChanged(object sender, EventArgs e)
        {
            UpdateTlsOptionsState();
            ApplyTlsRuntimePreview("settings_tls_changed");
        }

        private void ApplyTlsRuntimePreview(string source)
        {
            if (_suppressImmediateTlsApply || (Result != null && Result.HasManagedTransportTls))
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
                ApplySelectedTransportSecurity(source);
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

    }
}
