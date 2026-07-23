// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Resolves upload name collisions through the wizard's rename dialog.
    internal sealed partial class FileLinkWizardForm
    {
        private static string BuildDuplicateRenamePrompt(string originalName)
        {
            string name = string.IsNullOrWhiteSpace(originalName)
                ? Strings.AttachmentPromptLastUnknown
                : originalName.Trim();
            string template = Strings.FileLinkWizardRenameDuplicatePrompt ?? string.Empty;
            if (template.IndexOf("$1", StringComparison.Ordinal) >= 0)
            {
                return template.Replace("$1", name);
            }
            try
            {
                return string.Format(CultureInfo.CurrentCulture, template, name);
            }
            catch (FormatException)
            {
                return template;
            }
        }

        private string HandleDuplicate(FileLinkDuplicateInfo info)
        {
            if (info == null)
            {
                return null;
            }
            if (InvokeRequired)
            {
                string result = null;
                Invoke(new MethodInvoker(() => { result = ShowRenameDialog(info); }));
                return result;
            }
            return ShowRenameDialog(info);
        }

        private string ShowRenameDialog(FileLinkDuplicateInfo info)
        {
            using (var form = new Form())
            {
                form.Text = info.IsDirectory ? Strings.FileLinkWizardRenameFolderTitle : Strings.FileLinkWizardRenameFileTitle;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.AutoScaleDimensions = new SizeF(96f, 96f);
                form.ClientSize = new Size(520, 170);
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.StartPosition = FormStartPosition.CenterParent;
                form.ShowInTaskbar = false;
                form.Icon = BrandingAssets.GetAppIcon(32);

                string promptText = info.IsDirectory
                    ? Strings.FileLinkWizardRenameFolderPrompt
                    : BuildDuplicateRenamePrompt(info.OriginalName);
                var label = new Label
                {
                    AutoSize = true,
                    Text = promptText,
                    Location = new Point(12, 12),
                };
                label.MaximumSize = new Size(Math.Max(260, form.ClientSize.Width - 24), 0);
                form.Controls.Add(label);

                var textBox = new TextBox
                {
                    Location = new Point(12, 62),
                    Width = 496,
                    Text = info.OriginalName
                };
                form.Controls.Add(textBox);

                var okButton = new Button { Text = Strings.DialogOk, DialogResult = DialogResult.OK };
                var cancelButton = new Button { Text = Strings.DialogCancel, DialogResult = DialogResult.Cancel };
                int ignoredOkMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(okButton, out ignoredOkMinWidth);
                int ignoredCancelMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(cancelButton, out ignoredCancelMinWidth);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                Action layoutDialog = () =>
                {
                    int outerPadding = 12;
                    int verticalGap = 12;
                    int rowGap = 10;
                    int textBoxHeight = Math.Max(textBox.PreferredHeight, ScaleLogical(24));

                    label.MaximumSize = new Size(Math.Max(260, form.ClientSize.Width - (outerPadding * 2)), 0);
                    label.Location = new Point(outerPadding, outerPadding);
                    textBox.SetBounds(outerPadding, label.Bottom + rowGap, Math.Max(220, form.ClientSize.Width - (outerPadding * 2)), textBoxHeight);

                    int requiredClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                        form,
                        new[] { okButton, cancelButton },
                        FooterButtonLayoutHelper.DefaultHorizontalPadding,
                        FooterButtonLayoutHelper.DefaultBottomPadding,
                        FooterButtonLayoutHelper.DefaultSpacing,
                        true);
                    if (requiredClientWidth > form.ClientSize.Width)
                    {
                        form.ClientSize = new Size(requiredClientWidth, form.ClientSize.Height);
                        label.MaximumSize = new Size(Math.Max(260, form.ClientSize.Width - (outerPadding * 2)), 0);
                        label.Location = new Point(outerPadding, outerPadding);
                        textBox.SetBounds(outerPadding, label.Bottom + rowGap, Math.Max(220, form.ClientSize.Width - (outerPadding * 2)), textBoxHeight);
                        FooterButtonLayoutHelper.LayoutCentered(
                            form,
                            new[] { okButton, cancelButton },
                            FooterButtonLayoutHelper.DefaultHorizontalPadding,
                            FooterButtonLayoutHelper.DefaultBottomPadding,
                            FooterButtonLayoutHelper.DefaultSpacing,
                            true);
                    }
                    int requiredHeight = textBox.Bottom + verticalGap + okButton.Height + FooterButtonLayoutHelper.DefaultBottomPadding;
                    if (requiredHeight > form.ClientSize.Height)
                    {
                        form.ClientSize = new Size(form.ClientSize.Width, requiredHeight);
                        FooterButtonLayoutHelper.LayoutCentered(
                            form,
                            new[] { okButton, cancelButton },
                            FooterButtonLayoutHelper.DefaultHorizontalPadding,
                            FooterButtonLayoutHelper.DefaultBottomPadding,
                            FooterButtonLayoutHelper.DefaultSpacing,
                            true);
                    }
                };

                UiThemeManager.ApplyToForm(form);
                layoutDialog();

                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    string input = textBox.Text.Trim();
                    if (string.IsNullOrEmpty(input))
                    {
                        return null;
                    }
                    string sanitized = FileLinkPath.SanitizeComponent(input);
                    if (string.IsNullOrEmpty(sanitized))
                    {
                        return null;
                    }
                    if (!info.IsDirectory)
                    {
                        SelectionUploadState state;
                        if (_selectionStates.TryGetValue(info.Selection, out state))
                        {
                            state.RenamedTo = sanitized;
                        }
                    }
                    return sanitized;
                }
            }
            return null;
        }
    }
}
