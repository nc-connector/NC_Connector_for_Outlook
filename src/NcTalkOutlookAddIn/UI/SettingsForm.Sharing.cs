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
                Math.Max(
                    Math.Max(_sharingDefaultPasswordCheckBox.PreferredSize.Width, _sharingDefaultPasswordSeparateCheckBox.PreferredSize.Width),
                    _sharingPasswordDeliveryModeLabel.PreferredSize.Width + _sharingPasswordDeliveryModeCombo.Width + ScaleLogical(12)),
                _sharingDefaultExpireDaysLabel.PreferredSize.Width + _sharingDefaultExpireDaysUpDown.Width + ScaleLogical(12));
            bool stackRightColumn = rightColumnX + rightColumnRequired + ScaleLogical(12) > groupWidth;
            int rightColumnTop = ScaleLogical(114);
            if (stackRightColumn)
            {
                rightColumnX = leftColumnX;
                rightColumnTop = _sharingDefaultPermDeleteCheckBox.Bottom + ScaleLogical(14);
            }

            _sharingDefaultPasswordCheckBox.Location = new Point(rightColumnX, rightColumnTop);
            _sharingDefaultExpireDaysLabel.Location = new Point(rightColumnX, _sharingDefaultPasswordCheckBox.Bottom + ScaleLogical(12));
            _sharingDefaultExpireDaysUpDown.Location = new Point(rightColumnX, _sharingDefaultExpireDaysLabel.Bottom + ScaleLogical(6));
            _sharingDefaultPasswordSeparateCheckBox.Location = new Point(rightColumnX, _sharingDefaultExpireDaysUpDown.Bottom + ScaleLogical(12));
            _sharingPasswordDeliveryModeLabel.Location = new Point(rightColumnX, _sharingDefaultPasswordSeparateCheckBox.Bottom + ScaleLogical(12));
            _sharingPasswordDeliveryModeCombo.Location = new Point(rightColumnX, _sharingPasswordDeliveryModeLabel.Bottom + ScaleLogical(6));

            int columnsBottom = Math.Max(_sharingDefaultPermDeleteCheckBox.Bottom, _sharingPasswordDeliveryModeCombo.Bottom);
            int automationWidth = Math.Max(ScaleLogical(320), groupWidth - ScaleLogical(24));
            _shareBlockLangLabel.MaximumSize = new Size(automationWidth, 0);
            _shareBlockLangLabel.Location = new Point(ScaleLogical(12), columnsBottom + ScaleLogical(18));
            _shareBlockLangCombo.SetBounds(
                ScaleLogical(12),
                _shareBlockLangLabel.Bottom + ScaleLogical(6),
                automationWidth,
                Math.Max(_shareBlockLangCombo.Height, _shareBlockLangCombo.PreferredHeight + ScaleLogical(2)));
            int automationTop = _shareBlockLangCombo.Bottom + ScaleLogical(18);
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
            bool policyLockLinkTarget = IsPolicyLocked("share", AttachmentLinkTargetPolicy.Key);
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
            _sharingAttachmentLinkTargetLabel.Location = new Point(innerPadding, contentBottom + ScaleLogical(12));
            _sharingAttachmentLinkTargetCombo.Location = new Point(
                innerPadding,
                _sharingAttachmentLinkTargetLabel.Bottom + ScaleLogical(6));
            _sharingAttachmentLinkTargetCombo.Width = Math.Max(
                ScaleLogical(240),
                _sharingAttachmentAutomationGroup.ClientSize.Width - (innerPadding * 2));
            contentBottom = _sharingAttachmentLinkTargetCombo.Bottom;

            int requiredGroupHeight = Math.Max(ScaleLogical(150), contentBottom + innerPadding);
            if (_sharingAttachmentAutomationGroup.Height != requiredGroupHeight)
            {
                _sharingAttachmentAutomationGroup.Height = requiredGroupHeight;
            }

            _sharingAttachmentsAlwaysCheckBox.Enabled = !effectiveAlwaysLock && !uiBusy;
            _sharingAttachmentsOfferAboveCheckBox.Enabled = !effectiveThresholdLock && !alwaysConnector && !uiBusy;

            bool thresholdInputEnabled = !effectiveThresholdLock && !alwaysConnector && _sharingAttachmentsOfferAboveCheckBox.Checked && !uiBusy;
            _sharingAttachmentsOfferAboveMbUpDown.Enabled = thresholdInputEnabled;
            _sharingAttachmentsOfferAboveUnitLabel.Enabled = !effectiveThresholdLock && !alwaysConnector && !uiBusy;
            _sharingAttachmentLinkTargetCombo.Enabled = !policyLockLinkTarget && !uiBusy;

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
            _disabledTooltipHints.Apply(
                _sharingAttachmentLinkTargetCombo,
                policyLockLinkTarget ? Strings.PolicyAdminControlledTooltip : Strings.SharingAttachmentLinkTargetTooltip,
                policyLockLinkTarget,
                _sharingAttachmentLinkTargetLabel);
        }

        private void SelectAttachmentLinkTarget(AttachmentLinkTarget target)
        {
            if (_sharingAttachmentLinkTargetCombo.Items.Count == 0)
            {
                _sharingAttachmentLinkTargetCombo.Items.Add(Strings.SharingAttachmentLinkTargetZipDownload);
                _sharingAttachmentLinkTargetCombo.Items.Add(Strings.SharingAttachmentLinkTargetSharePage);
            }
            _sharingAttachmentLinkTargetCombo.SelectedIndex =
                target == AttachmentLinkTarget.SharePage ? 1 : 0;
        }

        private AttachmentLinkTarget GetSelectedAttachmentLinkTarget()
        {
            return _sharingAttachmentLinkTargetCombo.SelectedIndex == 1
                ? AttachmentLinkTarget.SharePage
                : AttachmentLinkTarget.ZipDownload;
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
            _sharingDefaultPasswordCheckBox.CheckedChanged += (s, e) => UpdateControlState();
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPasswordCheckBox);

            _sharingDefaultPasswordSeparateCheckBox.Text = Strings.SharingDefaultPasswordSeparateLabel;
            _sharingDefaultPasswordSeparateCheckBox.Location = new Point(260, 138);
            _sharingDefaultPasswordSeparateCheckBox.AutoSize = true;
            _sharingDefaultPasswordSeparateCheckBox.CheckedChanged += (s, e) => UpdateControlState();
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultPasswordSeparateCheckBox);

            _sharingPasswordDeliveryModeLabel.Text = Strings.SharingPasswordDeliveryModeLabel;
            _sharingPasswordDeliveryModeLabel.Location = new Point(260, 162);
            _sharingPasswordDeliveryModeLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingPasswordDeliveryModeLabel);

            _sharingPasswordDeliveryModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _sharingPasswordDeliveryModeCombo.IntegralHeight = false;
            _sharingPasswordDeliveryModeCombo.Location = new Point(260, 184);
            _sharingPasswordDeliveryModeCombo.Width = 220;
            SharePasswordDeliveryModeComboHelper.Populate(_sharingPasswordDeliveryModeCombo);
            _sharingDefaultsGroup.Controls.Add(_sharingPasswordDeliveryModeCombo);

            _sharingDefaultExpireDaysLabel.Text = Strings.SharingDefaultExpireDaysLabel;
            _sharingDefaultExpireDaysLabel.Location = new Point(260, 216);
            _sharingDefaultExpireDaysLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysLabel);

            _sharingDefaultExpireDaysUpDown.Minimum = 1;
            _sharingDefaultExpireDaysUpDown.Maximum = 3650;
            _sharingDefaultExpireDaysUpDown.Location = new Point(260, 240);
            _sharingDefaultExpireDaysUpDown.Width = 90;
            _sharingDefaultsGroup.Controls.Add(_sharingDefaultExpireDaysUpDown);

            _shareBlockLangLabel.Text = Strings.AdvancedShareBlockLangLabel;
            _shareBlockLangLabel.AutoSize = true;
            _sharingDefaultsGroup.Controls.Add(_shareBlockLangLabel);

            _shareBlockLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _shareBlockLangCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _shareBlockLangCombo.IntegralHeight = false;
            _shareBlockLangCombo.DrawItem += HandleLanguageComboDrawItem;
            _shareBlockLangCombo.SelectionChangeCommitted += HandleLanguageComboSelectionCommitted;
            PopulateLanguageOverrideCombo(_shareBlockLangCombo, "share");
            _sharingDefaultsGroup.Controls.Add(_shareBlockLangCombo);

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

            _sharingAttachmentLinkTargetLabel.Text = Strings.SharingAttachmentLinkTargetLabel;
            _sharingAttachmentLinkTargetLabel.Location = new Point(12, 82);
            _sharingAttachmentLinkTargetLabel.AutoSize = true;
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentLinkTargetLabel);

            _sharingAttachmentLinkTargetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _sharingAttachmentLinkTargetCombo.IntegralHeight = false;
            _sharingAttachmentLinkTargetCombo.Location = new Point(12, 104);
            _sharingAttachmentLinkTargetCombo.Width = 448;
            SelectAttachmentLinkTarget(AttachmentLinkTarget.ZipDownload);
            _sharingAttachmentAutomationGroup.Controls.Add(_sharingAttachmentLinkTargetCombo);

            _toolTip.SetToolTip(_sharingDefaultPermissionsLabel, Strings.TooltipSharingPermissions);
            _toolTip.SetToolTip(_sharingDefaultPasswordSeparateCheckBox, string.Empty);
            _toolTip.SetToolTip(_sharingPasswordDeliveryModeCombo, string.Empty);
            _toolTip.SetToolTip(_sharingAttachmentsAlwaysCheckBox, Strings.TooltipSharingAttachmentsAlways);
            _toolTip.SetToolTip(_sharingAttachmentsOfferAboveCheckBox, Strings.TooltipSharingAttachmentsOffer);
            _toolTip.SetToolTip(_sharingAttachmentLinkTargetLabel, Strings.SharingAttachmentLinkTargetTooltip);
            _toolTip.SetToolTip(_sharingAttachmentLinkTargetCombo, Strings.SharingAttachmentLinkTargetTooltip);
        }

    }
}
