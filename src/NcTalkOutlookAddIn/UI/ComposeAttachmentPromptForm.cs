/**
 * Copyright (c) 2026 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal enum ComposeAttachmentPromptDecision
    {
        Share,
        RemoveLast
    }

    /**
     * Two-action prompt for threshold-based compose attachment automation.
     * There is intentionally no third cancel action.
     */
    internal sealed class ComposeAttachmentPromptForm : Form
    {
        private readonly Label _reasonLabel = new Label();
        private readonly Button _shareButton = new Button();
        private readonly Button _removeButton = new Button();
        private bool _layoutAdjustingClientSize;

        private ComposeAttachmentPromptDecision _decision = ComposeAttachmentPromptDecision.RemoveLast;

        internal ComposeAttachmentPromptForm(string reasonText)
        {
            Text = Strings.AttachmentPromptTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(540, 210);
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            ShowInTaskbar = false;
            Icon = BrandingAssets.GetAppIcon(32);
            MinimumSize = new Size(ScaleLogical(520), ScaleLogical(190));

            _reasonLabel.AutoSize = true;
            _reasonLabel.MaximumSize = new Size(ClientSize.Width - ScaleLogical(32), 0);
            _reasonLabel.Location = new Point(ScaleLogical(16), ScaleLogical(16));
            _reasonLabel.Text = reasonText ?? string.Empty;
            _reasonLabel.ForeColor = Color.Black;
            Controls.Add(_reasonLabel);

            _removeButton.Text = Strings.AttachmentPromptRemoveLast;
            int ignoredRemoveMinWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(_removeButton, out ignoredRemoveMinWidth);
            _removeButton.Click += (s, e) =>
            {
                _decision = ComposeAttachmentPromptDecision.RemoveLast;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(_removeButton);

            _shareButton.Text = Strings.AttachmentPromptShare;
            int ignoredShareMinWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(_shareButton, out ignoredShareMinWidth);
            _shareButton.Click += (s, e) =>
            {
                _decision = ComposeAttachmentPromptDecision.Share;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(_shareButton);

            AcceptButton = _shareButton;
            CancelButton = _removeButton;

            UiThemeManager.ApplyToForm(this);
            ApplyDialogLayout(true);
        }

        internal static ComposeAttachmentPromptDecision ShowPrompt(IWin32Window owner, string reasonText)
        {
            using (var dialog = new ComposeAttachmentPromptForm(reasonText))
            {
                dialog.ShowDialog(owner);
                return dialog._decision;
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ApplyDialogLayout(true);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (_layoutAdjustingClientSize)
            {
                return;
            }

            ApplyDialogLayout(false);
        }

        private void ApplyDialogLayout(bool ensureClientSize)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }
            int outerPadding = ScaleLogical(16);
            int verticalGap = ScaleLogical(14);
            int bottomPadding = ScaleLogical(12);

            _reasonLabel.MaximumSize = new Size(Math.Max(ScaleLogical(240), ClientSize.Width - (outerPadding * 2)), 0);
            _reasonLabel.Location = new Point(outerPadding, outerPadding);

            var buttons = new[] { _removeButton, _shareButton };
            int requiredClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                this,
                buttons,
                FooterButtonLayoutHelper.DefaultHorizontalPadding,
                FooterButtonLayoutHelper.DefaultBottomPadding,
                FooterButtonLayoutHelper.DefaultSpacing,
                true);
            if (ensureClientSize && requiredClientWidth > ClientSize.Width)
            {
                _layoutAdjustingClientSize = true;
                try
                {
                    ClientSize = new Size(requiredClientWidth, ClientSize.Height);
                }
                finally
                {
                    _layoutAdjustingClientSize = false;
                }

                _reasonLabel.MaximumSize = new Size(Math.Max(ScaleLogical(240), ClientSize.Width - (outerPadding * 2)), 0);
                _reasonLabel.Location = new Point(outerPadding, outerPadding);
                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    buttons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing,
                    true);
            }
            int buttonsTop = Math.Min(_removeButton.Top, _shareButton.Top);
            int requiredHeight = _reasonLabel.Bottom + verticalGap + _removeButton.Height + bottomPadding;
            if (ensureClientSize && requiredHeight > ClientSize.Height)
            {
                _layoutAdjustingClientSize = true;
                try
                {
                    ClientSize = new Size(ClientSize.Width, requiredHeight);
                }
                finally
                {
                    _layoutAdjustingClientSize = false;
                }

                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    buttons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    FooterButtonLayoutHelper.DefaultSpacing,
                    true);
                buttonsTop = Math.Min(_removeButton.Top, _shareButton.Top);
            }
            int maxReasonBottom = buttonsTop - verticalGap;
            if (_reasonLabel.Bottom > maxReasonBottom)
            {
                _reasonLabel.MaximumSize = new Size(
                    Math.Max(ScaleLogical(240), ClientSize.Width - (outerPadding * 2)),
                    Math.Max(ScaleLogical(40), maxReasonBottom - _reasonLabel.Top));
            }
        }

        private int ScaleLogical(int value)
        {
            return DpiScaling.ScaleLogical(this, value);
        }
    }
}
