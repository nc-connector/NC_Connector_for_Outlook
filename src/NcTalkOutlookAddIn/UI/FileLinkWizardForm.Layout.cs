// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Builds and arranges the wizard shell and non-file selection steps.
    internal sealed partial class FileLinkWizardForm
    {
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (_layoutAdjustingClientSize)
            {
                return;
            }

            LayoutBottomButtons();
            LayoutPolicyWarningPanel();
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
            PositionProgressBars();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            AdjustInitialDialogSizeForDisplay();
            ReflowWizardLayout();
        }


        private void InitializeHeader()
        {
            _headerPanel.Height = 48;
            _headerPanel.Dock = DockStyle.Top;
            _headerPanel.Padding = new Padding(0);

            Controls.Add(_headerPanel);
        }

        private void InitializePolicyWarningPanel()
        {
            PolicyUiHelper.InitializePolicyWarningPanel(
                _policyWarningPanel,
                _policyWarningTitleLabel,
                _policyWarningTextLabel,
                _policyWarningLinkLabel);
            Controls.Add(_policyWarningPanel);
            _policyWarningLinkLabel.LinkClicked += (s, e) => OpenPolicyAdminGuide();
        }

        private void InitializeWizardLayout()
        {
            int headerBottom = _headerPanel.Bottom;
            _titleLabel.Location = new Point(20, headerBottom + 12);
            _titleLabel.AutoSize = true;
            _titleLabel.Font = new Font("Segoe UI", 12f, FontStyle.Bold, GraphicsUnit.Point);
            Controls.Add(_titleLabel);

            _stepHost = new Panel
            {
                Location = new Point(20, _titleLabel.Bottom + 16),
                Size = new Size(ClientSize.Width - 40, ClientSize.Height - (_titleLabel.Bottom + 16) - StepHostBottomReservedPixels),
                BorderStyle = BorderStyle.None
            };
            Controls.Add(_stepHost);

            _backButton.Text = Strings.ButtonBack;
            _backButton.AutoSize = false;
            _backButton.Click += async (s, e) =>
                await NavigateAsync(-1);
            Controls.Add(_backButton);

            _uploadButton.Text = Strings.FileLinkWizardUploadButton;
            _uploadButton.AutoSize = false;
            _uploadButton.Enabled = false;
            _uploadButton.Visible = false;
            _uploadButton.Click += async (s, e) => await StartUploadAsync();
            Controls.Add(_uploadButton);

            _nextButton.Text = Strings.ButtonNext;
            _nextButton.AutoSize = false;
            _nextButton.Click += async (s, e) =>
                await NavigateAsync(1);
            Controls.Add(_nextButton);

            _finishButton.Text = Strings.FileLinkWizardFinishButton;
            _finishButton.AutoSize = false;
            _finishButton.Click += async (s, e) => await FinishAsync();
            Controls.Add(_finishButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.AutoSize = false;
            _cancelButton.Click += (s, e) => Close();
            Controls.Add(_cancelButton);

            LayoutBottomButtons();
            LayoutPolicyWarningPanel();
            UpdateStepHostBounds();
            LayoutProgressPanel();
        }

        private void LayoutPolicyWarningPanel()
        {
            int left = ScaleLogical(20);
            int top = _titleLabel.Bottom + ScaleLogical(10);
            int width = Math.Max(ScaleLogical(240), ClientSize.Width - ScaleLogical(40));
            if (!_policyWarningPanel.Visible)
            {
                _policyWarningPanel.SetBounds(left, top, width, 0);
                return;
            }
            int padding = ScaleLogical(8);
            int textWidth = Math.Max(ScaleLogical(180), width - (padding * 2));
            _policyWarningTitleLabel.Location = new Point(padding, padding);
            _policyWarningTitleLabel.MaximumSize = new Size(textWidth, 0);

            int textTop = _policyWarningTitleLabel.Bottom + ScaleLogical(4);
            _policyWarningTextLabel.Location = new Point(padding, textTop);
            _policyWarningTextLabel.MaximumSize = new Size(textWidth, 0);

            int linkTop = _policyWarningTextLabel.Bottom + ScaleLogical(6);
            _policyWarningLinkLabel.Location = new Point(padding, linkTop);

            int height = _policyWarningLinkLabel.Bottom + padding;
            _policyWarningPanel.SetBounds(left, top, width, height);
        }

        private void UpdateStepHostBounds()
        {
            if (_stepHost == null || IsDisposed || Disposing)
            {
                return;
            }
            int top = _policyWarningPanel.Visible
                ? _policyWarningPanel.Bottom + ScaleLogical(10)
                : _titleLabel.Bottom + ScaleLogical(16);
            _stepHost.Location = new Point(ScaleLogical(20), top);

            int stepHostWidth = Math.Max(0, ClientSize.Width - 40);
            int stepHostBottom = GetStepHostBottomLimit();
            int minHeight = ScaleLogical(StepHostMinimumHeightPixels);
            int stepHostHeight = Math.Max(minHeight, stepHostBottom - _stepHost.Top);
            _stepHost.Size = new Size(stepHostWidth, stepHostHeight);
        }

        private void LayoutProgressPanel()
        {
            if (_progressPanel == null || _progressPanel.IsDisposed || _progressPanel.Disposing)
            {
                return;
            }
            int left = _stepHost != null ? _stepHost.Left : ScaleLogical(20);
            int width = _stepHost != null ? _stepHost.Width : Math.Max(ScaleLogical(240), ClientSize.Width - ScaleLogical(40));
            int top = _stepHost != null ? _stepHost.Bottom + ScaleLogical(8) : ScaleLogical(360);
            int labelHeight = Math.Max(
                ScaleLogical(18),
                _progressLabel.PreferredHeight);
            int gap = ScaleLogical(5);
            int barHeight = Math.Max(
                ScaleLogical(14),
                _progressBar.Height);
            int panelHeight = labelHeight + gap + barHeight;
            _progressPanel.SetBounds(left, top, width, panelHeight);

            _progressLabel.SetBounds(
                0,
                0,
                _progressPanel.ClientSize.Width,
                labelHeight);
            _progressLabel.MaximumSize = Size.Empty;
            _progressLabel.AutoEllipsis = true;
            _progressBar.SetBounds(
                0,
                _progressLabel.Bottom + gap,
                _progressPanel.ClientSize.Width,
                barHeight);
        }

        private void SetProgressPanelVisible(bool visible)
        {
            _progressPanel.Visible = visible;
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
        }

        private void LayoutCurrentStep()
        {
            if (_stepHost == null || _stepHost.IsDisposed || _stepHost.Disposing)
            {
                return;
            }

            Size clientSize = _stepHost.ClientSize;
            switch (_currentStepIndex)
            {
                case 0:
                    LayoutGeneralStep(clientSize);
                    break;
                case 1:
                    LayoutExpirationStep(clientSize);
                    break;
                case 2:
                    LayoutFileStep(clientSize);
                    break;
                case 3:
                    LayoutNoteStep(clientSize);
                    break;
            }
        }

        private int GetStepHostBottomLimit()
        {
            int fallbackBottom = ClientSize.Height - StepHostBottomReservedPixels;
            int topMostButton = int.MaxValue;
            var buttons = new[] { _backButton, _uploadButton, _nextButton, _finishButton, _cancelButton };
            foreach (Button button in buttons)
            {
                if (button != null && button.Visible)
                {
                    topMostButton = Math.Min(topMostButton, button.Top);
                }
            }
            if (topMostButton == int.MaxValue)
            {
                return fallbackBottom;
            }
            int safeBottomFromButtons = topMostButton - 12;
            int bottom = Math.Min(fallbackBottom, safeBottomFromButtons);
            if (_progressPanel.Visible)
            {
                int labelHeight = Math.Max(
                    ScaleLogical(18),
                    _progressLabel.PreferredHeight);
                int progressHeight = labelHeight
                                     + ScaleLogical(5)
                                     + Math.Max(
                                         ScaleLogical(14),
                                         _progressBar.Height);
                bottom -= progressHeight + ScaleLogical(8);
            }
            return bottom;
        }

        private void InitializeStepGeneral()
        {
            _generalStepPanel = CreateStepPanel();
            var panel = _generalStepPanel;

            _shareNameLabel.Text = Strings.FileLinkWizardShareNameLabel;
            _shareNameLabel.AutoSize = true;
            _shareNameLabel.MaximumSize = new Size(ScaleLogical(320), 0);
            panel.Controls.Add(_shareNameLabel);

            _shareNameTextBox.Width = 360;
            _shareNameTextBox.Text = string.IsNullOrWhiteSpace(_defaults.SharingDefaultShareName)
                ? Strings.FileLinkWizardFallbackShareName
                : _defaults.SharingDefaultShareName;
            _shareNameTextBox.TextChanged += (s, e) => InvalidateUpload();
            panel.Controls.Add(_shareNameTextBox);

            _permissionsLabel.Text = Strings.FileLinkWizardPermissionsLabel;
            _permissionsLabel.AutoSize = true;
            _permissionsLabel.MaximumSize = new Size(ScaleLogical(320), 0);
            panel.Controls.Add(_permissionsLabel);

            _permissionReadCheckBox.Text = Strings.FileLinkPermissionRead;
            _permissionReadCheckBox.AutoSize = true;
            _permissionReadCheckBox.Checked = true;
            _permissionReadCheckBox.Enabled = false;
            panel.Controls.Add(_permissionReadCheckBox);

            _permissionCreateCheckBox.Text = Strings.FileLinkPermissionCreate;
            _permissionCreateCheckBox.AutoSize = true;
            _permissionCreateCheckBox.Checked = _defaults.SharingDefaultPermCreate;
            panel.Controls.Add(_permissionCreateCheckBox);

            _permissionWriteCheckBox.Text = Strings.FileLinkPermissionWrite;
            _permissionWriteCheckBox.AutoSize = true;
            _permissionWriteCheckBox.Checked = _defaults.SharingDefaultPermWrite;
            panel.Controls.Add(_permissionWriteCheckBox);

            _permissionDeleteCheckBox.Text = Strings.FileLinkPermissionDelete;
            _permissionDeleteCheckBox.AutoSize = true;
            _permissionDeleteCheckBox.Checked = _defaults.SharingDefaultPermDelete;
            panel.Controls.Add(_permissionDeleteCheckBox);

            _passwordToggleCheckBox.Text = Strings.FileLinkWizardPasswordToggle;
            _passwordToggleCheckBox.AutoSize = true;
            _passwordToggleCheckBox.Checked = _defaults.SharingDefaultPasswordEnabled;
            _passwordToggleCheckBox.CheckedChanged += (s, e) => UpdatePasswordState();
            panel.Controls.Add(_passwordToggleCheckBox);

            _passwordTextBox.Width = 220;
            panel.Controls.Add(_passwordTextBox);

            _passwordGenerateButton.Text = Strings.TalkPasswordGenerate;
            _passwordGenerateButton.AutoSize = false;
            int ignoredPasswordGenerateMinWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(_passwordGenerateButton, out ignoredPasswordGenerateMinWidth);
            _passwordGenerateButton.Click += (s, e) => GeneratePassword();
            panel.Controls.Add(_passwordGenerateButton);

            _passwordSeparateToggleCheckBox.Text = Strings.FileLinkWizardPasswordSeparateToggle;
            _passwordSeparateToggleCheckBox.AutoSize = true;
            _passwordSeparateToggleCheckBox.Checked = _defaults.SharingDefaultPasswordSeparateEnabled;
            _passwordSeparateToggleCheckBox.CheckedChanged += (s, e) => UpdatePasswordState();
            panel.Controls.Add(_passwordSeparateToggleCheckBox);

            _passwordDeliveryModeLabel.Text = Strings.SharingPasswordDeliveryModeLabel;
            _passwordDeliveryModeLabel.AutoSize = true;
            panel.Controls.Add(_passwordDeliveryModeLabel);

            _passwordDeliveryModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _passwordDeliveryModeCombo.IntegralHeight = false;
            _passwordDeliveryModeCombo.Width = 220;
            SharePasswordDeliveryModeComboHelper.Populate(_passwordDeliveryModeCombo);
            SharePasswordDeliveryModeComboHelper.Select(_passwordDeliveryModeCombo, _defaults.SharingDefaultPasswordDeliveryMode);
            panel.Controls.Add(_passwordDeliveryModeCombo);

            if (_attachmentMode)
            {
                _permissionCreateCheckBox.Checked = false;
                _permissionWriteCheckBox.Checked = false;
                _permissionDeleteCheckBox.Checked = false;
                _permissionCreateCheckBox.Enabled = false;
                _permissionWriteCheckBox.Enabled = false;
                _permissionDeleteCheckBox.Enabled = false;
                _shareNameTextBox.ReadOnly = true;
            }
            if (_passwordToggleCheckBox.Checked)
            {
                _passwordTextBox.Text = PasswordGenerationHelper.GenerateWithPolicyDefaults(
                    _configuration,
                    _passwordPolicy,
                    DefaultMinPasswordLength,
                    LogCategories.FileLink);
            }

            UpdatePasswordState();
            panel.ClientSizeChanged += (s, e) => LayoutGeneralStep(panel.ClientSize);
            LayoutGeneralStep(panel.ClientSize);

            _steps.Add(panel);
        }

        private void InitializeStepExpiration()
        {
            _expirationStepPanel = CreateStepPanel();
            var panel = _expirationStepPanel;

            _expireToggleCheckBox.Text = Strings.FileLinkWizardExpireToggle;
            _expireToggleCheckBox.AutoSize = true;
            _expireToggleCheckBox.Checked = _defaults.SharingDefaultExpireDays > 0;
            _expireToggleCheckBox.CheckedChanged += (s, e) => UpdateExpireState();
            panel.Controls.Add(_expireToggleCheckBox);

            _expireDatePicker.Width = 160;
            _expireDatePicker.Format = DateTimePickerFormat.Short;
            int expireDays = _defaults.SharingDefaultExpireDays > 0 ? _defaults.SharingDefaultExpireDays : 7;
            _expireDatePicker.Value = DateTime.Today.AddDays(expireDays);
            panel.Controls.Add(_expireDatePicker);

            _expireHintLabel.Text = Strings.FileLinkWizardExpireHint;
            _expireHintLabel.AutoSize = true;
            _expireHintLabel.ForeColor = Color.DimGray;
            _expireHintLabel.MaximumSize = new Size(ScaleLogical(360), 0);
            panel.Controls.Add(_expireHintLabel);

            UpdateExpireState();
            panel.ClientSizeChanged += (s, e) => LayoutExpirationStep(panel.ClientSize);
            LayoutExpirationStep(panel.ClientSize);
            _steps.Add(panel);
        }

        private void LayoutGeneralStep(Size clientSize)
        {
            if (_generalStepPanel == null || _generalStepPanel.IsDisposed || _generalStepPanel.Disposing)
            {
                return;
            }
            int left = ScaleLogical(12);
            int top = ScaleLogical(12);
            int indent = ScaleLogical(6);
            int rowGap = ScaleLogical(8);
            int sectionGap = ScaleLogical(14);
            int contentWidth = Math.Max(ScaleLogical(260), clientSize.Width - (left * 2) - ScaleLogical(12));

            _shareNameLabel.MaximumSize = new Size(contentWidth, 0);
            _shareNameLabel.Location = new Point(left, top);

            int textBoxHeight = Math.Max(ScaleLogical(24), _shareNameTextBox.PreferredHeight + ScaleLogical(2));
            _shareNameTextBox.SetBounds(left, _shareNameLabel.Bottom + ScaleLogical(6), contentWidth, textBoxHeight);

            _permissionsLabel.MaximumSize = new Size(contentWidth, 0);
            _permissionsLabel.Location = new Point(left, _shareNameTextBox.Bottom + sectionGap);

            int permissionLeft = left + indent;
            int permissionTop = _permissionsLabel.Bottom + ScaleLogical(6);
            _permissionReadCheckBox.Location = new Point(permissionLeft, permissionTop);
            _permissionCreateCheckBox.Location = new Point(permissionLeft, _permissionReadCheckBox.Bottom + rowGap);
            _permissionWriteCheckBox.Location = new Point(permissionLeft, _permissionCreateCheckBox.Bottom + rowGap);
            _permissionDeleteCheckBox.Location = new Point(permissionLeft, _permissionWriteCheckBox.Bottom + rowGap);

            int passwordSectionTop = _permissionDeleteCheckBox.Bottom + sectionGap;
            _passwordToggleCheckBox.Location = new Point(left, passwordSectionTop);

            int ignoredGenerateMinWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(_passwordGenerateButton, out ignoredGenerateMinWidth);
            int generateWidth = _passwordGenerateButton.Width;
            int generateHeight = _passwordGenerateButton.Height;
            int passwordHeight = Math.Max(ScaleLogical(24), _passwordTextBox.PreferredHeight + ScaleLogical(2));
            int passwordIndent = ScaleLogical(16);
            int passwordMinWidth = ScaleLogical(180);
            int passwordMaxWidth = ScaleLogical(420);
            int passwordInputLeft = left + passwordIndent;
            int passwordAvailable = (left + contentWidth) - passwordInputLeft;

            int passwordBottom;
            if (passwordAvailable >= passwordMinWidth + ScaleLogical(8) + generateWidth)
            {
                int passwordWidth = Math.Min(passwordMaxWidth, Math.Max(passwordMinWidth, passwordAvailable - generateWidth - ScaleLogical(8)));
                int rowY = _passwordToggleCheckBox.Bottom + ScaleLogical(6);
                _passwordTextBox.SetBounds(passwordInputLeft, rowY, passwordWidth, passwordHeight);
                _passwordGenerateButton.SetBounds(_passwordTextBox.Right + ScaleLogical(8), rowY - ScaleLogical(2), generateWidth, generateHeight);
                passwordBottom = Math.Max(_passwordTextBox.Bottom, _passwordGenerateButton.Bottom);
            }
            else
            {
                int rowY = _passwordToggleCheckBox.Bottom + ScaleLogical(6);
                int passwordWidth = Math.Max(passwordMinWidth, Math.Max(ScaleLogical(120), contentWidth - generateWidth - ScaleLogical(24)));
                _passwordTextBox.SetBounds(passwordInputLeft, rowY, passwordWidth, passwordHeight);
                _passwordGenerateButton.SetBounds(_passwordTextBox.Right + ScaleLogical(8), rowY - ScaleLogical(2), generateWidth, generateHeight);
                passwordBottom = Math.Max(_passwordTextBox.Bottom, _passwordGenerateButton.Bottom);
            }

            _passwordSeparateToggleCheckBox.Location = new Point(left, passwordBottom + sectionGap);

            _passwordDeliveryModeLabel.Location = new Point(passwordInputLeft, _passwordSeparateToggleCheckBox.Bottom + rowGap);
            int deliveryModeWidth = Math.Min(ScaleLogical(260), contentWidth);
            _passwordDeliveryModeCombo.SetBounds(passwordInputLeft, _passwordDeliveryModeLabel.Bottom + ScaleLogical(6), deliveryModeWidth, Math.Max(ScaleLogical(24), _passwordDeliveryModeCombo.PreferredHeight + ScaleLogical(2)));

            int requiredHeight = _passwordDeliveryModeCombo.Bottom + ScaleLogical(16);
            _generalStepPanel.AutoScrollMinSize = new Size(0, requiredHeight);
        }

        private void LayoutExpirationStep(Size clientSize)
        {
            if (_expirationStepPanel == null || _expirationStepPanel.IsDisposed || _expirationStepPanel.Disposing)
            {
                return;
            }
            int left = ScaleLogical(12);
            int top = ScaleLogical(12);
            int rowGap = ScaleLogical(10);
            int sectionGap = ScaleLogical(14);
            int contentWidth = Math.Max(ScaleLogical(260), clientSize.Width - (left * 2) - ScaleLogical(12));

            _expireToggleCheckBox.Location = new Point(left, top);

            int pickerWidth = Math.Max(ScaleLogical(150), _expireDatePicker.Width);
            int pickerHeight = Math.Max(ScaleLogical(24), _expireDatePicker.PreferredHeight + ScaleLogical(2));
            int inlineStartX = _expireToggleCheckBox.Right + ScaleLogical(10);
            int inlineAvailable = (left + contentWidth) - inlineStartX;
            if (inlineAvailable >= pickerWidth)
            {
                _expireDatePicker.SetBounds(inlineStartX, _expireToggleCheckBox.Top - ScaleLogical(2), pickerWidth, pickerHeight);
            }
            else
            {
                _expireDatePicker.SetBounds(left + ScaleLogical(16), _expireToggleCheckBox.Bottom + rowGap, pickerWidth, pickerHeight);
            }
            int hintTop = Math.Max(_expireToggleCheckBox.Bottom, _expireDatePicker.Bottom) + sectionGap;
            _expireHintLabel.MaximumSize = new Size(contentWidth, 0);
            _expireHintLabel.Location = new Point(left, hintTop);

            int requiredHeight = _expireHintLabel.Bottom + ScaleLogical(16);
            _expirationStepPanel.AutoScrollMinSize = new Size(0, requiredHeight);
        }

        private void LayoutNoteStep(Size clientSize)
        {
            if (_noteStepPanel == null || _noteStepPanel.IsDisposed || _noteStepPanel.Disposing)
            {
                return;
            }
            int left = ScaleLogical(12);
            int top = ScaleLogical(12);
            int rowGap = ScaleLogical(8);
            int contentWidth = Math.Max(ScaleLogical(260), clientSize.Width - (left * 2) - ScaleLogical(12));

            _noteToggleCheckBox.Location = new Point(left, top);
            int noteTop = _noteToggleCheckBox.Bottom + rowGap;
            int noteHeight = Math.Max(ScaleLogical(140), clientSize.Height - noteTop - ScaleLogical(12));
            _noteTextBox.SetBounds(left, noteTop, contentWidth, noteHeight);

            int requiredHeight = _noteTextBox.Bottom + ScaleLogical(16);
            _noteStepPanel.AutoScrollMinSize = new Size(0, requiredHeight);
        }


        private void InitializeStepNote()
        {
            _noteStepPanel = CreateStepPanel();
            var panel = _noteStepPanel;

            _noteToggleCheckBox.Text = Strings.FileLinkWizardNoteToggle;
            _noteToggleCheckBox.AutoSize = true;
            _noteToggleCheckBox.CheckedChanged += (s, e) => UpdateNoteState();
            panel.Controls.Add(_noteToggleCheckBox);

            _noteTextBox.Multiline = true;
            _noteTextBox.ScrollBars = ScrollBars.Vertical;
            _noteTextBox.Enabled = false;
            panel.Controls.Add(_noteTextBox);

            panel.ClientSizeChanged += (s, e) => LayoutNoteStep(panel.ClientSize);
            LayoutNoteStep(panel.ClientSize);
            _steps.Add(panel);
        }

        private void InitializeProgressPanel()
        {
            _progressPanel.Visible = false;
            _progressPanel.Location = new Point(20, 360);
            _progressPanel.Size = new Size(600, 48);

            _progressLabel.AutoSize = false;
            _progressLabel.AutoEllipsis = true;
            _progressPanel.Controls.Add(_progressLabel);

            _progressBar.Location = new Point(0, 24);
            _progressBar.Size = new Size(600, 16);
            _progressPanel.Controls.Add(_progressBar);

            Controls.Add(_progressPanel);
            LayoutProgressPanel();
        }

        private Panel CreateStepPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Visible = false,
                AutoScroll = true
            };
            if (_stepHost != null)
            {
                _stepHost.Controls.Add(panel);
                panel.BringToFront();
            }
            return panel;
        }

        private void LayoutBottomButtons()
        {
            if (IsDisposed || Disposing)
            {
                return;
            }
            int spacing = 12;
            var buttons = new List<Button>();

            if (_backButton.Visible)
            {
                buttons.Add(_backButton);
            }
            if (_uploadButton.Visible)
            {
                buttons.Add(_uploadButton);
            }
            if (_nextButton.Visible)
            {
                buttons.Add(_nextButton);
            }
            if (_finishButton.Visible)
            {
                buttons.Add(_finishButton);
            }
            if (_cancelButton.Visible)
            {
                buttons.Add(_cancelButton);
            }
            if (buttons.Count == 0)
            {
                return;
            }
            int requiredClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                this,
                buttons,
                FooterButtonLayoutHelper.DefaultHorizontalPadding,
                FooterButtonLayoutHelper.DefaultBottomPadding,
                spacing,
                true);
            if (requiredClientWidth > ClientSize.Width)
            {
                EnsureDialogWidthForButtons(requiredClientWidth);
                FooterButtonLayoutHelper.LayoutCentered(
                    this,
                    buttons,
                    FooterButtonLayoutHelper.DefaultHorizontalPadding,
                    FooterButtonLayoutHelper.DefaultBottomPadding,
                    spacing,
                    true);
            }
        }

        private void EnsureDialogWidthForButtons(int requiredClientWidth)
        {
            if (requiredClientWidth <= ClientSize.Width || _layoutAdjustingClientSize || IsDisposed || Disposing)
            {
                return;
            }

            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            int maxClientWidth = Math.Max(ClientSize.Width, workingArea.Width - ScaleLogical(32));
            int targetWidth = Math.Min(requiredClientWidth, maxClientWidth);
            if (targetWidth <= ClientSize.Width)
            {
                return;
            }

            _layoutAdjustingClientSize = true;
            try
            {
                ClientSize = new Size(targetWidth, ClientSize.Height);
            }
            finally
            {
                _layoutAdjustingClientSize = false;
            }
        }

        private void ReflowWizardLayout()
        {
            LayoutBottomButtons();
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
            ConfigureFileListViewRowHeight();
            PositionProgressBars();
            Invalidate(true);
        }

        private void AdjustInitialDialogSizeForDisplay()
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            int screenMargin = ScaleLogical(40);
            int maxClientWidth = Math.Max(ScaleLogical(560), workingArea.Width - screenMargin);
            int maxClientHeight = Math.Max(ScaleLogical(420), workingArea.Height - screenMargin);

            int targetWidth = Math.Min(maxClientWidth, Math.Max(ClientSize.Width, ScaleLogical(760)));
            int targetHeight = Math.Min(maxClientHeight, Math.Max(ClientSize.Height, ScaleLogical(620)));

            if (targetWidth == ClientSize.Width && targetHeight == ClientSize.Height)
            {
                return;
            }
            if (_layoutAdjustingClientSize)
            {
                return;
            }

            _layoutAdjustingClientSize = true;
            try
            {
                ClientSize = new Size(targetWidth, targetHeight);
            }
            finally
            {
                _layoutAdjustingClientSize = false;
            }
        }
    }
}
