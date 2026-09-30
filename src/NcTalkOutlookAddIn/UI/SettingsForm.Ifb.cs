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
        private readonly Label _ifbPolicyHintLabel = new Label();

        private void ApplyIfbTabLayout()
        {
            int left = ScaleLogical(24);
            int top = ScaleLogical(20);
            int rowGap = ScaleLogical(18);
            int labelToComboGap = ScaleLogical(Result != null && Result.HasManagedIfb ? 28 : 14);

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
            _ifbPolicyHintLabel.MaximumSize = new Size(Math.Max(ScaleLogical(220), _ifbTab.ClientSize.Width - 2 * left), 0);
            _ifbPolicyHintLabel.Location = new Point(left, Math.Max(_ifbPortLabel.Bottom, _ifbPortUpDown.Bottom) + rowGap);
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

            _ifbPolicyHintLabel.AutoSize = true;
            _ifbTab.Controls.Add(_ifbPolicyHintLabel);
        }

        private void UpdateIfbOptionsState(bool credentialsAvailable)
        {
            bool managed = Result != null && Result.HasManagedIfb;
            if (!managed)
            {
                if (credentialsAvailable && !_ifbDefaultApplied && !_initialIfbEnabled && !_ifbEnabledCheckBox.Checked)
                {
                    _ifbEnabledCheckBox.Checked = true;
                    _ifbDefaultApplied = true;
                }
                if (!credentialsAvailable && _ifbEnabledCheckBox.Checked)
                {
                    _ifbEnabledCheckBox.Checked = false;
                }
            }

            _ifbEnabledCheckBox.Enabled = !managed && credentialsAvailable && !_isBusy;
            bool showOptions = managed || _ifbEnabledCheckBox.Checked;
            bool editOptions = showOptions && _ifbEnabledCheckBox.Enabled;
            _ifbDaysCombo.Visible = showOptions;
            _ifbDaysLabel.Visible = showOptions;
            _ifbDaysCombo.Enabled = editOptions;
            _ifbDaysLabel.Enabled = editOptions;
            _ifbPortUpDown.Visible = showOptions;
            _ifbPortLabel.Visible = showOptions;
            _ifbPortUpDown.Enabled = editOptions;
            _ifbPortLabel.Enabled = editOptions;
            _ifbCacheHoursCombo.Enabled = !managed && !_isBusy;
            _ifbCacheHoursLabel.Enabled = !managed && !_isBusy;

            bool invalid = managed && !Result.IsManagedIfbValid;
            string hint = managed
                ? (invalid ? Strings.ManagedIfbPolicyInvalid : Strings.PolicyAdminControlledTooltip)
                : string.Empty;
            _ifbPolicyHintLabel.Text = invalid ? hint : string.Empty;
            _ifbPolicyHintLabel.Visible = invalid;
            _ifbPolicyHintLabel.ForeColor = _themePalette.ErrorText;
            _disabledTooltipHints.Apply(_ifbEnabledCheckBox, hint, managed);
            _disabledTooltipHints.Apply(_ifbDaysCombo, hint, managed, _ifbDaysLabel);
            _disabledTooltipHints.Apply(_ifbPortUpDown, hint, managed, _ifbPortLabel);
            _disabledTooltipHints.Apply(_ifbCacheHoursCombo, hint, managed, _ifbCacheHoursLabel);
            ApplyIfbTabLayout();
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

    }
}
