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
        private readonly Label _debugPolicyHintLabel = new Label();

        private void ApplyDebugTabLayout()
        {
            int rightMargin = ScaleLogical(24);
            _debugAnonymizeCheckBox.Location = new Point(_debugLogCheckBox.Left, _debugLogCheckBox.Bottom + ScaleLogical(8));
            int width = Math.Max(ScaleLogical(220), _debugTab.ClientSize.Width - rightMargin - _debugPathLabel.Left);
            _debugPolicyHintLabel.Location = new Point(_debugLogCheckBox.Left, _debugAnonymizeCheckBox.Bottom + ScaleLogical(12));
            _debugPolicyHintLabel.MaximumSize = new Size(width, 0);
            bool managed = Result != null && Result.HasManagedLogging;
            _debugPathLabel.Location = new Point(_debugPathLabel.Left, (managed ? _debugPolicyHintLabel.Bottom : _debugAnonymizeCheckBox.Bottom) + ScaleLogical(12));
            _debugPathLabel.MaximumSize = new Size(width, 0);
            _debugPathLabel.AutoSize = true;
            _debugOpenLink.Location = new Point(_debugOpenLink.Left, _debugPathLabel.Bottom + ScaleLogical(10));
        }

        private void InitializeDebugTab()
        {
            _debugTab.AutoScroll = true;
            _debugTab.Padding = new Padding(12);

            _debugLogCheckBox.Text = Strings.DebugCheckbox;
            _debugLogCheckBox.AutoSize = true;
            _debugLogCheckBox.Location = new Point(24, 20);
            _debugTab.Controls.Add(_debugLogCheckBox);

            _debugAnonymizeCheckBox.Text = Strings.DebugAnonymizeCheckbox;
            _debugAnonymizeCheckBox.AutoSize = true;
            _debugAnonymizeCheckBox.Location = new Point(24, 50);
            _debugTab.Controls.Add(_debugAnonymizeCheckBox);

            _debugPolicyHintLabel.AutoSize = true;
            _debugTab.Controls.Add(_debugPolicyHintLabel);

            _debugPathLabel.AutoSize = true;
            _debugPathLabel.Location = new Point(24, 90);
            _debugPathLabel.MaximumSize = new Size(420, 0);
            _debugTab.Controls.Add(_debugPathLabel);

            _debugOpenLink.Text = Strings.DebugOpenLog;
            _debugOpenLink.Location = new Point(24, 140);
            _debugOpenLink.AutoSize = true;
            _debugOpenLink.LinkClicked += OnDebugOpenLinkClicked;
            _debugTab.Controls.Add(_debugOpenLink);

            UpdateDebugPathLabel();
        }

        private void UpdateLoggingOptionsState()
        {
            bool managed = Result != null && Result.HasManagedLogging;
            string hint = managed
                ? (Result.IsManagedLoggingValid ? Strings.DebugManagedHint : Strings.ManagedLoggingPolicyInvalid)
                : string.Empty;
            _debugLogCheckBox.Enabled = !managed && !_isBusy;
            _debugAnonymizeCheckBox.Enabled = !managed && !_isBusy;
            _debugPolicyHintLabel.Text = hint;
            _debugPolicyHintLabel.Visible = managed;
            _disabledTooltipHints.Apply(_debugLogCheckBox, hint, managed);
            _disabledTooltipHints.Apply(_debugAnonymizeCheckBox, hint, managed);
            ApplyDebugTabLayout();
        }

        private void UpdateDebugPathLabel()
        {
            string path = DiagnosticsLogger.LogFileFullPath ?? string.Empty;
            _debugPathLabel.Text = Strings.DebugPathPrefix + path;
        }

        private void OnDebugOpenLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string path = DiagnosticsLogger.LogFileFullPath ?? string.Empty;
            string target;
            string failureContext;
            if (File.Exists(path))
            {
                target = path;
                failureContext = "Failed to open debug log file.";
            }
            else
            {
                string directory = Path.GetDirectoryName(path);
                if (string.IsNullOrEmpty(directory)
                    || !Directory.Exists(directory))
                {
                    MessageBox.Show(
                        Strings.DebugLogMissingMessage,
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }
                target = directory;
                failureContext = "Failed to open debug log directory.";
            }

            if (!BrowserLauncher.OpenTarget(
                target,
                LogCategories.Core,
                failureContext))
            {
                MessageBox.Show(
                    Strings.DebugLogOpenErrorMessage,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

    }
}
