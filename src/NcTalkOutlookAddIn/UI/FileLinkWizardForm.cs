// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Multi-step Nextcloud sharing wizard.
    internal sealed partial class FileLinkWizardForm : ScaledForm
    {
        private const int DefaultMinPasswordLength = 8;
        private const string AttachmentShareNameBase = "email_attachment";
        private const int PathColumnWheelStepPixels = 48;
        private const int StepHostBottomReservedPixels = 88;
        private const int StepHostMinimumHeightPixels = 180;
        private const int FileStepPaddingPixels = 12;
        private const int FileStepButtonGapPixels = 8;
        private const int FileStepButtonColumnSpacingPixels = 12;
        private const int FileStepButtonColumnMinWidthPixels = 168;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

        private readonly FileLinkService _service;
        private readonly FileLinkRequest _request = new FileLinkRequest();
        private readonly TalkServiceConfiguration _configuration;
        private readonly PasswordPolicyInfo _passwordPolicy;
        private readonly BackendPolicyStatus _backendPolicyStatus;
        private readonly AddinSettings _defaults;
        private readonly FileLinkWizardLaunchOptions _launchOptions;
        private readonly OutlookAttachmentAutomationGuardService _attachmentGuardService = new OutlookAttachmentAutomationGuardService();
        private readonly List<Panel> _steps = new List<Panel>();
        private readonly Label _titleLabel = new Label();
        private readonly Button _backButton = new Button();
        private readonly Button _nextButton = new Button();
        private readonly Button _finishButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly Button _uploadButton = new Button();
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private readonly Panel _policyWarningPanel = new Panel();
        private readonly Label _policyWarningTitleLabel = new Label();
        private readonly Label _policyWarningTextLabel = new Label();
        private readonly LinkLabel _policyWarningLinkLabel = new LinkLabel();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly DisabledControlTooltipHintHelper _disabledTooltipHints;
        private Panel _stepHost;
        private readonly Panel _progressPanel = new Panel();
        private readonly ProgressBar _progressBar = new ProgressBar();
        private readonly Label _progressLabel = new Label();
        private readonly PathScrollableListView _fileListView = new PathScrollableListView();
        private readonly ImageList _fileListRowHeightImageList = new ImageList();
        private readonly Label _basePathLabel = new Label();
        private readonly Label _shareNameLabel = new Label();
        private readonly Label _permissionsLabel = new Label();
        private readonly TableLayoutPanel _fileStepLayout = new TableLayoutPanel();
        private readonly TableLayoutPanel _fileStepContentLayout = new TableLayoutPanel();
        private readonly FlowLayoutPanel _fileStepActionPanel = new FlowLayoutPanel();
        private readonly TextBox _shareNameTextBox = new TextBox();
        private readonly CheckBox _permissionReadCheckBox = new CheckBox();
        private readonly CheckBox _permissionCreateCheckBox = new CheckBox();
        private readonly CheckBox _permissionWriteCheckBox = new CheckBox();
        private readonly CheckBox _permissionDeleteCheckBox = new CheckBox();
        private readonly CheckBox _passwordToggleCheckBox = new CheckBox();
        private readonly TextBox _passwordTextBox = new TextBox();
        private readonly Button _passwordGenerateButton = new Button();
        private readonly CheckBox _passwordSeparateToggleCheckBox = new CheckBox();
        private readonly Label _passwordDeliveryModeLabel = new Label();
        private readonly ComboBox _passwordDeliveryModeCombo = new ComboBox();
        private readonly CheckBox _expireToggleCheckBox = new CheckBox();
        private readonly DateTimePicker _expireDatePicker = new DateTimePicker();
        private readonly Label _expireHintLabel = new Label();
        private readonly Button _addFilesButton = new Button();
        private readonly Button _addFolderButton = new Button();
        private readonly Button _removeItemButton = new Button();
        private readonly Label _attachmentModeInfoLabel = new Label();
        private readonly CheckBox _noteToggleCheckBox = new CheckBox();
        private readonly TextBox _noteTextBox = new TextBox();
        private readonly List<FileLinkSelection> _items = new List<FileLinkSelection>();
        private readonly Dictionary<FileLinkSelection, SelectionUploadState> _selectionStates = new Dictionary<FileLinkSelection, SelectionUploadState>();
        private int _currentStepIndex;
        private CancellationTokenSource _cancellationSource;
        private FileLinkUploadContext _uploadContext;
        private bool _uploadInProgress;
        private bool _uploadCompleted;
        private bool _allowEmptyUpload;
        private bool _shareFinalized;
        private bool _closeAfterCancellation;
        private int _pathColumnHorizontalOffset;
        private int _pathColumnMaxHorizontalOffset;
        private FileLinkRequest _requestSnapshot;
        private readonly bool _attachmentMode;
        private readonly DateTime _shareDate;
        private bool _layoutAdjustingClientSize;
        private Panel _generalStepPanel;
        private Panel _expirationStepPanel;
        private Panel _noteStepPanel;

        internal FileLinkWizardForm(
            AddinSettings defaults,
            TalkServiceConfiguration configuration,
            NextcloudCapabilitiesSnapshot capabilitiesSnapshot,
            PasswordPolicyInfo passwordPolicy,
            BackendPolicyStatus policyStatus,
            string basePath,
            FileLinkWizardLaunchOptions launchOptions)
        {
            _defaults = (defaults ?? new AddinSettings()).Clone();
            _configuration = configuration;
            _passwordPolicy = passwordPolicy;
            _backendPolicyStatus = policyStatus;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            _service = new FileLinkService(
                configuration,
                capabilitiesSnapshot);
            _launchOptions = launchOptions ?? new FileLinkWizardLaunchOptions();
            _attachmentMode = _launchOptions.AttachmentMode;
            _shareDate = DateTime.Now;
            _request.BasePath = basePath ?? string.Empty;
            _request.AttachmentMode = _attachmentMode;
            _request.ShareDate = _shareDate;
            ApplyPolicyDefaultsToSettings();
            _request.AttachmentLinkTarget = AttachmentLinkTargetPolicy.Resolve(
                _defaults.SharingAttachmentLinkTarget,
                _backendPolicyStatus);
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Attachment link target resolved to "
                + AttachmentLinkTargetPolicy.ToStorageValue(_request.AttachmentLinkTarget)
                + ".");

            Text = Strings.FileLinkWizardTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ControlBox = true;
            ShowInTaskbar = true;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(640, 480);
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimumSize = new Size(ScaleLogical(700), ScaleLogical(560));
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeWizardLayout();
            InitializePolicyWarningPanel();
            InitializeStepGeneral();
            InitializeStepExpiration();
            InitializeStepFiles();
            InitializeStepNote();
            InitializeProgressPanel();
            InitializeUploadProgressPump();
            AdjustInitialDialogSizeForDisplay();

            UiThemeManager.ApplyToForm(this);

            LoadInitialSelections();
            if (_attachmentMode)
            {
                ApplyAttachmentModeDefaults();
            }
            else
            {
                ShowStep(0);
            }

            ApplyPolicyWarningUi();
            ApplyPolicyLockState();
        }

        internal FileLinkResult Result { get; private set; }

        internal FileLinkRequest RequestSnapshot
        {
            get { return _requestSnapshot; }
        }



        private void ShowStep(int index)
        {
            if (index < 0 || index >= _steps.Count)
            {
                return;
            }
            foreach (Panel panel in _steps)
            {
                panel.Visible = false;
            }

            _currentStepIndex = index;
            _steps[index].Visible = true;

            string title;
            switch (index)
            {
                case 0:
                    title = Strings.FileLinkWizardStepShare;
                    break;
                case 1:
                    title = Strings.FileLinkWizardStepExpire;
                    break;
                case 2:
                    title = Strings.FileLinkWizardStepFiles;
                    break;
                case 3:
                    title = Strings.FileLinkWizardStepNote;
                    break;
                default:
                    title = string.Empty;
                    break;
            }
            _titleLabel.Text = title;

            UpdateNavigationState();
            UpdateUploadButtonState();
            LayoutCurrentStep();
            LayoutProgressPanel();
            PositionProgressBars();
        }

        private void Navigate(int direction)
        {
            if (_attachmentMode)
            {
                return;
            }
            if (direction > 0 && !ValidateCurrentStep())
            {
                return;
            }
            int newIndex = _currentStepIndex + direction;
            ShowStep(newIndex);
        }

        private bool ValidateCurrentStep()
        {
            if (_attachmentMode)
            {
                if (_currentStepIndex != 2)
                {
                    return true;
                }
                if (_items.Count == 0)
                {
                    MessageBox.Show(Strings.FileLinkWizardSelectFileOrFolder, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                return true;
            }
            if (_currentStepIndex == 0)
            {
                if (string.IsNullOrWhiteSpace(_shareNameTextBox.Text))
                {
                    MessageBox.Show(Strings.FileLinkWizardShareNameRequired, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _shareNameTextBox.Focus();
                    return false;
                }
                if (_passwordToggleCheckBox.Checked && !IsPasswordValid(_passwordTextBox.Text))
                {
                    MessageBox.Show(
                        string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.TalkPasswordTooShort,
                            PasswordGenerationHelper.ResolveMinLength(_passwordPolicy, DefaultMinPasswordLength)),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    _passwordTextBox.Focus();
                    return false;
                }
            }
            else if (_currentStepIndex == 1)
            {
                if (_expireToggleCheckBox.Checked && _expireDatePicker.Value.Date < DateTime.Today)
                {
                    MessageBox.Show(Strings.FileLinkWizardExpireMustBeFuture, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            else if (_currentStepIndex == 2)
            {
                if (_items.Count == 0)
                {
                    if (_permissionCreateCheckBox.Checked)
                    {
                        if (!ConfirmEmptyUploadProceed())
                        {
                            return false;
                        }
                        _allowEmptyUpload = true;
                        _uploadCompleted = true;
                    }
                    else
                    {
                        MessageBox.Show(Strings.FileLinkWizardSelectFileOrFolder, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }
                else
                {
                    _allowEmptyUpload = false;
                }
                if (!_uploadCompleted)
                {
                    MessageBox.Show(Strings.FileLinkWizardUploadFirst, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            return true;
        }


        private async Task FinishAsync()
        {
            if (_attachmentMode)
            {
                if (!ValidateCurrentStep())
                {
                    return;
                }

                ApplyFormData();
                if (!EnsureAttachmentAutomationAllowedForFinalize())
                {
                    return;
                }
                if (!_uploadCompleted)
                {
                    await StartUploadAsync();
                    if (!_uploadCompleted)
                    {
                        return;
                    }
                }
            }
            else
            {
                if (!ValidateCurrentStep())
                {
                    return;
                }

                ApplyFormData();
                if (_allowEmptyUpload && (_uploadContext == null || !_uploadCompleted))
                {
                    if (!await PrepareEmptyUploadAsync())
                    {
                        return;
                    }
                }
            }
            if (_uploadContext == null || !_uploadCompleted)
            {
                MessageBox.Show(
                    Strings.FileLinkWizardUploadFirst,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _backButton.Enabled = false;
            _nextButton.Enabled = false;
            _finishButton.Enabled = false;
            _cancelButton.Enabled = false;
            SetProgressPanelVisible(true);
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressLabel.Text = Strings.FileLinkWizardCreatingShare;

            _cancellationSource = new CancellationTokenSource();

            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UseWaitCursor = true;

                Result = await Task.Run(() => _service.FinalizeShare(_uploadContext, _request, _cancellationSource.Token));
                _shareFinalized = true;
                _requestSnapshot = CloneRequest(_request);

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Share creation failed.", ex);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.FileLinkWizardCreateFailedFormat, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                SetProgressPanelVisible(false);
                _finishButton.Enabled = true;
            }
            catch (OperationCanceledException)
            {
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Share creation cancelled.");
                SetProgressPanelVisible(false);
                _finishButton.Enabled = true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Share creation failed unexpectedly.", ex);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.FileLinkWizardCreateFailedFormat, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                SetProgressPanelVisible(false);
                _finishButton.Enabled = true;
            }
            finally
            {
                UseWaitCursor = false;
                Cursor.Current = previousCursor;
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }
                _progressBar.Style = ProgressBarStyle.Blocks;
                _cancelButton.Enabled = true;
                UpdateNavigationState();
                UpdateUploadButtonState();
                CloseAfterCancellation();
            }
        }

        private void ApplyFormData()
        {
            _request.ShareName = _shareNameTextBox.Text.Trim();
            _request.Permissions = FileLinkPermissionFlags.Read;
            if (!_attachmentMode && _permissionCreateCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Create;
            }
            if (!_attachmentMode && _permissionWriteCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Write;
            }
            if (!_attachmentMode && _permissionDeleteCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Delete;
            }

            _request.PasswordEnabled = _passwordToggleCheckBox.Checked;
            _request.Password = _passwordToggleCheckBox.Checked ? _passwordTextBox.Text : null;
            _request.PasswordSeparateEnabled =
                _passwordToggleCheckBox.Checked
                && PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus)
                && _passwordSeparateToggleCheckBox.Checked;
            _request.PasswordDeliveryMode = SharePasswordDeliveryModeComboHelper.GetSelected(_passwordDeliveryModeCombo);
            _request.ExpireEnabled = _expireToggleCheckBox.Checked;
            _request.ExpireDate = _expireToggleCheckBox.Checked ? _expireDatePicker.Value.Date : (DateTime?)null;
            _request.NoteEnabled = !_attachmentMode && _noteToggleCheckBox.Checked;
            _request.Note = !_attachmentMode && _noteToggleCheckBox.Checked ? _noteTextBox.Text.Trim() : null;
            _request.AttachmentMode = _attachmentMode;
            _request.ShareDate = _shareDate;

        }

        private static FileLinkRequest CloneRequest(FileLinkRequest source)
        {
            var clone = new FileLinkRequest
            {
                BasePath = source.BasePath,
                ShareName = source.ShareName,
                Permissions = source.Permissions,
                PasswordEnabled = source.PasswordEnabled,
                Password = source.Password,
                PasswordSeparateEnabled = source.PasswordSeparateEnabled,
                PasswordDeliveryMode = source.PasswordDeliveryMode,
                ExpireEnabled = source.ExpireEnabled,
                ExpireDate = source.ExpireDate,
                NoteEnabled = source.NoteEnabled,
                Note = source.Note,
                AttachmentMode = source.AttachmentMode,
                AttachmentLinkTarget = source.AttachmentLinkTarget,
                ShareDate = source.ShareDate
            };
            return clone;
        }


        private void LoadInitialSelections()
        {
            if (_launchOptions == null || _launchOptions.InitialSelections == null)
            {
                return;
            }
            var validSelections = new List<FileLinkSelection>();
            foreach (var selection in _launchOptions.InitialSelections)
            {
                if (!SelectionPathExists(selection))
                {
                    continue;
                }

                validSelections.Add(new FileLinkSelection(selection.SelectionType, selection.LocalPath));
            }

            AddSelections(validSelections);
        }

        private void ApplyAttachmentModeDefaults()
        {
            _noteToggleCheckBox.Checked = false;
            _noteToggleCheckBox.Enabled = false;
            _noteTextBox.Text = string.Empty;
            _noteTextBox.Enabled = false;

            string infoText = BuildAttachmentModeInfoText();
            _attachmentModeInfoLabel.Text = infoText;
            _attachmentModeInfoLabel.Visible = !string.IsNullOrWhiteSpace(infoText);

            _shareNameTextBox.Text = AttachmentShareNameBase;

            ShowStep(2);
        }

        private string BuildAttachmentModeInfoText()
        {
            if (_launchOptions == null || string.Equals(_launchOptions.AttachmentTrigger, "always", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.FileLinkWizardAttachmentModeReasonAlways;
            }
            return string.Format(
                CultureInfo.CurrentCulture,
                Strings.FileLinkWizardAttachmentModeReasonThreshold,
                SizeFormatting.FormatMegabytes(_launchOptions.AttachmentTotalBytes),
                Math.Max(1, _launchOptions.AttachmentThresholdMb).ToString(CultureInfo.CurrentCulture) + " MB",
                string.IsNullOrWhiteSpace(_launchOptions.AttachmentLastName) ? Strings.AttachmentPromptLastUnknown : _launchOptions.AttachmentLastName.Trim(),
                SizeFormatting.FormatMegabytes(_launchOptions.AttachmentLastSizeBytes));
        }

        private bool EnsureAttachmentAutomationAllowedForFinalize()
        {
            if (!_attachmentMode)
            {
                return true;
            }
            try
            {
                var state = _attachmentGuardService.ReadLiveState();
                if (state == null || !state.LockActive)
                {
                    return true;
                }
                string message = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.SharingAttachmentAutomationLockedError,
                    state.ThresholdMb.ToString(CultureInfo.CurrentCulture));
                MessageBox.Show(
                    message,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Attachment mode finalize blocked by host setting (thresholdMb="
                    + state.ThresholdMb.ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (state.Source ?? string.Empty)
                    + ").");
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Attachment mode finalize guard check failed.", ex);
                return false;
            }
        }

        private void UpdatePasswordState()
        {
            bool enabled = _passwordToggleCheckBox.Checked;
            bool lockPasswordSeparate = IsPolicyLocked("share_send_password_separately");
            bool lockPasswordDeliveryMode = IsPolicyLocked("share_send_password_mode");
            bool separatePasswordAvailable = PolicyUiHelper.HasBackendSeatEntitlement(_backendPolicyStatus);
            string separatePasswordUnavailableTooltip = PolicyUiHelper.GetSeparatePasswordUnavailableTooltip(_backendPolicyStatus);
            bool passwordDeliveryModeAvailable = PolicyUiHelper.HasPasswordDeliveryMode(_backendPolicyStatus);
            string passwordDeliveryUnavailableTooltip = PolicyUiHelper.GetPasswordDeliveryModeUnavailableTooltip(_backendPolicyStatus);
            _passwordTextBox.Enabled = enabled;
            _passwordGenerateButton.Enabled = enabled;
            _passwordSeparateToggleCheckBox.Enabled = enabled && !lockPasswordSeparate && separatePasswordAvailable;
            _passwordDeliveryModeCombo.Enabled =
                enabled
                && separatePasswordAvailable
                && passwordDeliveryModeAvailable
                && _passwordSeparateToggleCheckBox.Checked
                && !lockPasswordDeliveryMode;
            _disabledTooltipHints.Apply(
                _passwordSeparateToggleCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockPasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockPasswordSeparate,
                (Control)null,
                _passwordTextBox);
            _disabledTooltipHints.Apply(
                _passwordDeliveryModeCombo,
                !passwordDeliveryModeAvailable
                    ? passwordDeliveryUnavailableTooltip
                    : (lockPasswordDeliveryMode
                        ? Strings.PolicyAdminControlledTooltip
                        : (!_passwordSeparateToggleCheckBox.Checked ? Strings.SharingPasswordDeliveryEnableSeparateTooltip : string.Empty)),
                !passwordDeliveryModeAvailable
                    || lockPasswordDeliveryMode
                    || !_passwordSeparateToggleCheckBox.Checked,
                _passwordDeliveryModeLabel);
            if (!enabled || !separatePasswordAvailable)
            {
                _passwordSeparateToggleCheckBox.Checked = false;
            }
            if (!passwordDeliveryModeAvailable)
            {
                SharePasswordDeliveryModeComboHelper.Select(_passwordDeliveryModeCombo, SharePasswordDeliveryMode.Plain);
            }
            if (_generalStepPanel != null)
            {
                LayoutGeneralStep(_generalStepPanel.ClientSize);
            }
        }

        private void UpdateExpireState()
        {
            bool enabled = _expireToggleCheckBox.Checked;
            bool lockExpireDays = IsPolicyLocked("share_expire_days");
            _expireDatePicker.Enabled = enabled && !lockExpireDays;
            _expireHintLabel.Enabled = enabled;
            if (_expirationStepPanel != null)
            {
                LayoutExpirationStep(_expirationStepPanel.ClientSize);
            }
        }

        private void UpdateNoteState()
        {
            _noteTextBox.Enabled = _noteToggleCheckBox.Checked;
            if (_noteStepPanel != null)
            {
                LayoutNoteStep(_noteStepPanel.ClientSize);
            }
        }





        private void UpdateNavigationState()
        {
            bool onFileStep = _currentStepIndex == 2;
            bool onLastStep = _currentStepIndex == _steps.Count - 1;

            if (_attachmentMode)
            {
                _backButton.Visible = false;
                _nextButton.Visible = false;
                _uploadButton.Visible = false;
                _finishButton.Visible = onFileStep;
                _cancelButton.Visible = true;
                _finishButton.Enabled = onFileStep && !_uploadInProgress && (_uploadCompleted || _items.Count > 0);
                LayoutBottomButtons();
                return;
            }

            _backButton.Visible = true;
            _nextButton.Visible = true;
            _finishButton.Visible = true;
            _cancelButton.Visible = true;

            _backButton.Enabled = _currentStepIndex > 0 && !_uploadInProgress;

            bool canAdvance = _currentStepIndex < _steps.Count - 1 && !_uploadInProgress;
            if (onFileStep)
            {
                if (_items.Count == 0 && _permissionCreateCheckBox.Checked)
                {
                    canAdvance = !_uploadInProgress;
                }
                else
                {
                    canAdvance = _uploadCompleted && !_uploadInProgress;
                }
            }
            _nextButton.Enabled = canAdvance;

            bool finishAllowed = _uploadCompleted || (_allowEmptyUpload && _items.Count == 0);
            _finishButton.Enabled = onLastStep && finishAllowed && !_uploadInProgress;
            _uploadButton.Visible = onFileStep;
            LayoutBottomButtons();
            UpdateStepHostBounds();
            LayoutCurrentStep();
        }

        private void UpdateUploadButtonState()
        {
            _uploadButton.Enabled = _uploadButton.Visible && !_uploadInProgress && _items.Count > 0 && !_uploadCompleted;
        }



        private void GeneratePassword()
        {
            if (!_passwordToggleCheckBox.Checked)
            {
                return;
            }
            _passwordTextBox.Text = PasswordGenerationHelper.GenerateWithPolicyDefaults(
                _configuration,
                _passwordPolicy,
                DefaultMinPasswordLength,
                LogCategories.FileLink);
        }

        private bool IsPasswordValid(string password)
        {
            return PasswordGenerationHelper.MeetsMinimumLength(password, _passwordPolicy, DefaultMinPasswordLength);
        }
    }
}

