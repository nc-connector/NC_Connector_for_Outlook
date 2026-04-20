/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

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
    /**
     * Multi-step wizard for creating a Nextcloud share (permissions, expiration date,
     * file/folder selection, note, and upload progress).
     */
    internal sealed partial class FileLinkWizardForm : Form
    {
        private const int DefaultMinPasswordLength = 8;
        private const string AttachmentShareNameBase = "email_attachment";
        private const string AttachmentShareDatePrefixFormat = "yyyyMMdd";
        private const int PathColumnWheelStepPixels = 48;
        private const int UploadProgressFlushIntervalMs = 80;
        private const int StepHostBottomReservedPixels = 88;
        private const int StepHostMinimumHeightPixels = 180;
        private const int FileStepPaddingPixels = 12;
        private const int FileStepButtonGapPixels = 8;
        private const int FileStepButtonColumnSpacingPixels = 12;
        private const int FileStepButtonColumnMinWidthPixels = 168;
        private static readonly Color ActiveQueueItemBackground = BrandingAssets.BrandBlue;
        private static readonly Color ActiveQueueItemText = Color.White;
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
        private readonly System.Windows.Forms.Timer _uploadProgressFlushTimer = new System.Windows.Forms.Timer();
        private readonly object _uploadProgressSync = new object();
        private readonly Dictionary<FileLinkSelection, FileLinkUploadItemProgress> _pendingUploadProgress = new Dictionary<FileLinkSelection, FileLinkUploadItemProgress>();
        private readonly List<FileLinkSelection> _pendingUploadProgressOrder = new List<FileLinkSelection>();
        private int _currentStepIndex;
        private CancellationTokenSource _cancellationSource;
        private FileLinkUploadContext _uploadContext;
        private bool _uploadInProgress;
        private bool _uploadCompleted;
        private bool _allowEmptyUpload;
        private bool _shareFinalized;
        private int _pathColumnHorizontalOffset;
        private int _pathColumnMaxHorizontalOffset;
        private ListViewItem _lastAutoScrolledUploadItem;
        private bool _uploadProgressPumpRequested;
        private FileLinkRequest _requestSnapshot;
        private readonly bool _attachmentMode;
        private readonly DateTime _shareDate;
        private bool _layoutAdjustingClientSize;
        private Panel _generalStepPanel;
        private Panel _expirationStepPanel;
        private Panel _noteStepPanel;

        internal FileLinkWizardForm(AddinSettings defaults, TalkServiceConfiguration configuration, PasswordPolicyInfo passwordPolicy, BackendPolicyStatus policyStatus, string basePath)
            : this(defaults, configuration, passwordPolicy, policyStatus, basePath, null)
        {
        }

        internal FileLinkWizardForm(AddinSettings defaults, TalkServiceConfiguration configuration, PasswordPolicyInfo passwordPolicy, BackendPolicyStatus policyStatus, string basePath, FileLinkWizardLaunchOptions launchOptions)
        {
            _defaults = (defaults ?? new AddinSettings()).Clone();
            _configuration = configuration;
            _passwordPolicy = passwordPolicy;
            _backendPolicyStatus = policyStatus;
            _disabledTooltipHints = new DisabledControlTooltipHintHelper(_toolTip);
            _service = new FileLinkService(configuration);
            _launchOptions = launchOptions ?? new FileLinkWizardLaunchOptions();
            _attachmentMode = _launchOptions.AttachmentMode;
            _shareDate = DateTime.Now;
            _request.BasePath = basePath ?? string.Empty;
            _request.AttachmentMode = _attachmentMode;
            _request.ShareDate = _shareDate;
            _request.ShareDatePrefixFormat = _attachmentMode ? AttachmentShareDatePrefixFormat : "yyyyMMdd";
            ApplyPolicyDefaultsToSettings();

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

        private bool IsPolicyActive()
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.PolicyActive;
        }

        private bool IsPolicyLocked(string key)
        {
            return _backendPolicyStatus != null && _backendPolicyStatus.IsLocked("share", key);
        }

        /**
         * Return true when the current user has an active backend seat.
         * Separate password delivery is gated by this entitlement.
         */
        private bool HasBackendSeatEntitlement()
        {
            return _backendPolicyStatus != null
                   && _backendPolicyStatus.EndpointAvailable
                   && _backendPolicyStatus.SeatAssigned
                   && _backendPolicyStatus.IsValid
                   && string.Equals(_backendPolicyStatus.SeatState, "active", StringComparison.OrdinalIgnoreCase);
        }

        /**
         * Return the tooltip shown when separate password delivery is unavailable.
         */
        private string GetSeparatePasswordUnavailableTooltip()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_backendPolicyStatus == null || !_backendPolicyStatus.EndpointAvailable)
            {
                return Strings.SharingPasswordSeparateBackendRequiredTooltip;
            }

            if (!_backendPolicyStatus.SeatAssigned)
            {
                return Strings.SharingPasswordSeparateNoSeatTooltip;
            }

            if (!_backendPolicyStatus.IsValid
                || !string.Equals(_backendPolicyStatus.SeatState, "active", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.SharingPasswordSeparatePausedTooltip;
            }

            return string.Empty;
        }

        private void ApplyPolicyDefaultsToSettings()
        {
            if (!IsPolicyActive())
            {
                return;
            }

            bool policyBool;
            int policyInt;
            string policyString;

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_base_directory");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _request.BasePath = policyString;
            }

            policyString = _backendPolicyStatus.GetPolicyString("share", "share_name_template");
            if (!string.IsNullOrWhiteSpace(policyString))
            {
                _defaults.SharingDefaultShareName = policyString;
            }

            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_upload", out policyBool))
            {
                _defaults.SharingDefaultPermCreate = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_edit", out policyBool))
            {
                _defaults.SharingDefaultPermWrite = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_permission_delete", out policyBool))
            {
                _defaults.SharingDefaultPermDelete = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_set_password", out policyBool))
            {
                _defaults.SharingDefaultPasswordEnabled = policyBool;
            }
            if (_backendPolicyStatus.TryGetPolicyBool("share", "share_send_password_separately", out policyBool))
            {
                _defaults.SharingDefaultPasswordSeparateEnabled = policyBool;
            }
            if (!HasBackendSeatEntitlement())
            {
                _defaults.SharingDefaultPasswordSeparateEnabled = false;
            }
            if (_backendPolicyStatus.TryGetPolicyInt("share", "share_expire_days", out policyInt))
            {
                _defaults.SharingDefaultExpireDays = Math.Max(1, policyInt);
            }
        }

        private void ApplyPolicyWarningUi()
        {
            bool visible = _backendPolicyStatus != null
                           && _backendPolicyStatus.WarningVisible
                           && !string.IsNullOrWhiteSpace(_backendPolicyStatus.WarningMessage);
            _policyWarningPanel.Visible = visible;
            _policyWarningTextLabel.Text = visible ? _backendPolicyStatus.WarningMessage : string.Empty;
            LayoutPolicyWarningPanel();
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
        }

        private void ApplyPolicyLockState()
        {
            bool lockShareName = IsPolicyLocked("share_name_template");
            bool lockPermCreate = IsPolicyLocked("share_permission_upload");
            bool lockPermWrite = IsPolicyLocked("share_permission_edit");
            bool lockPermDelete = IsPolicyLocked("share_permission_delete");
            bool lockPassword = IsPolicyLocked("share_set_password");
            bool lockPasswordSeparate = IsPolicyLocked("share_send_password_separately");
            bool lockExpireDays = IsPolicyLocked("share_expire_days");
            bool separatePasswordAvailable = HasBackendSeatEntitlement();
            string separatePasswordUnavailableTooltip = GetSeparatePasswordUnavailableTooltip();

            _shareNameTextBox.ReadOnly = _attachmentMode || lockShareName;
            _permissionCreateCheckBox.Enabled = !_attachmentMode && !lockPermCreate;
            _permissionWriteCheckBox.Enabled = !_attachmentMode && !lockPermWrite;
            _permissionDeleteCheckBox.Enabled = !_attachmentMode && !lockPermDelete;
            _passwordToggleCheckBox.Enabled = !lockPassword;
            _expireToggleCheckBox.Enabled = !lockExpireDays;

            SetTooltipWithFallback(_shareNameTextBox, lockShareName ? Strings.PolicyAdminControlledTooltip : string.Empty, lockShareName, _shareNameLabel, _titleLabel);
            SetTooltipWithFallback(_permissionCreateCheckBox, lockPermCreate ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermCreate, _permissionsLabel);
            SetTooltipWithFallback(_permissionWriteCheckBox, lockPermWrite ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermWrite, _permissionsLabel);
            SetTooltipWithFallback(_permissionDeleteCheckBox, lockPermDelete ? Strings.PolicyAdminControlledTooltip : string.Empty, lockPermDelete, _permissionsLabel);
            _disabledTooltipHints.Apply(
                _passwordToggleCheckBox,
                lockPassword ? Strings.PolicyAdminControlledTooltip : string.Empty,
                lockPassword,
                _passwordGenerateButton,
                _passwordGenerateButton,
                _passwordTextBox);
            SetTooltipWithFallback(
                _passwordSeparateToggleCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockPasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockPasswordSeparate,
                _passwordTextBox);
            SetTooltipWithFallback(_expireToggleCheckBox, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, lockExpireDays, _expireHintLabel);
            SetTooltipWithFallback(_expireDatePicker, lockExpireDays ? Strings.PolicyAdminControlledTooltip : string.Empty, false, _expireHintLabel);

            UpdatePasswordState();
            UpdateExpireState();
        }

        private void SetTooltipWithFallback(Control primary, string text, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, fallbackTargets);
        }

        private void SetTooltipWithFallback(Control primary, string text, bool showHint, params Control[] fallbackTargets)
        {
            _disabledTooltipHints.Apply(primary, text, showHint, fallbackTargets);
        }

        private static void OpenPolicyAdminGuide()
        {
            BrowserLauncher.OpenUrl(
                Strings.PolicyAdminGuideUrl,
                LogCategories.FileLink,
                "Failed to open policy admin guide URL.");
        }

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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_uploadInProgress && _cancellationSource != null && !_cancellationSource.IsCancellationRequested)
            {
                _cancellationSource.Cancel();
                e.Cancel = true;
                return;
            }

            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            ResetUploadProgressPump();
            _uploadProgressFlushTimer.Dispose();
            TryCleanupUnfinalizedUploadContext("wizard_closed_without_finalize");
            base.OnFormClosed(e);
        }

        private void InitializeUploadProgressPump()
        {
            _uploadProgressFlushTimer.Interval = UploadProgressFlushIntervalMs;
            _uploadProgressFlushTimer.Tick += (s, e) => FlushBufferedUploadProgress();
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
            _policyWarningPanel.Visible = false;
            _policyWarningPanel.BackColor = Color.FromArgb(20, 176, 0, 32);
            _policyWarningPanel.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    _policyWarningPanel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };
            Controls.Add(_policyWarningPanel);

            _policyWarningTitleLabel.AutoSize = true;
            _policyWarningTitleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            _policyWarningTitleLabel.Font = new Font(_policyWarningTitleLabel.Font, FontStyle.Bold);
            _policyWarningTitleLabel.Text = "\u26a0 " + Strings.PolicyWarningTitle;
            _policyWarningPanel.Controls.Add(_policyWarningTitleLabel);

            _policyWarningTextLabel.AutoSize = true;
            _policyWarningTextLabel.Text = string.Empty;
            _policyWarningPanel.Controls.Add(_policyWarningTextLabel);

            _policyWarningLinkLabel.AutoSize = true;
            _policyWarningLinkLabel.Text = Strings.PolicyWarningAdminLinkLabel;
            _policyWarningLinkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            _policyWarningLinkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            _policyWarningLinkLabel.LinkClicked += (s, e) => OpenPolicyAdminGuide();
            _policyWarningPanel.Controls.Add(_policyWarningLinkLabel);
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
            _backButton.Click += (s, e) => Navigate(-1);
            Controls.Add(_backButton);

            _uploadButton.Text = Strings.FileLinkWizardUploadButton;
            _uploadButton.AutoSize = false;
            _uploadButton.Enabled = false;
            _uploadButton.Visible = false;
            _uploadButton.Click += async (s, e) => await StartUploadAsync();
            Controls.Add(_uploadButton);

            _nextButton.Text = Strings.ButtonNext;
            _nextButton.AutoSize = false;
            _nextButton.Click += (s, e) => Navigate(1);
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_progressPanel == null || _progressPanel.IsDisposed || _progressPanel.Disposing)
            {
                return;
            }

            int left = _stepHost != null ? _stepHost.Left : ScaleLogical(20);
            int width = _stepHost != null ? _stepHost.Width : Math.Max(ScaleLogical(240), ClientSize.Width - ScaleLogical(40));
            int top = _stepHost != null ? _stepHost.Bottom + ScaleLogical(8) : ScaleLogical(360);
            int panelHeight = Math.Max(ScaleLogical(36), _progressBar.Height + ScaleLogical(18));
            _progressPanel.SetBounds(left, top, width, panelHeight);

            int gap = ScaleLogical(8);
            int labelWidth = Math.Max(ScaleLogical(120), Math.Min(width / 3, _progressLabel.PreferredWidth + ScaleLogical(10)));
            int barHeight = Math.Max(ScaleLogical(14), _progressBar.Height);
            int barTop = Math.Max(0, (_progressPanel.ClientSize.Height - barHeight) / 2);
            int barWidth = Math.Max(ScaleLogical(120), _progressPanel.ClientSize.Width - labelWidth - gap);
            _progressBar.SetBounds(0, barTop, barWidth, barHeight);

            int labelTop = Math.Max(0, (_progressPanel.ClientSize.Height - _progressLabel.PreferredHeight) / 2);
            _progressLabel.Location = new Point(_progressBar.Right + gap, labelTop);
            _progressLabel.MaximumSize = new Size(Math.Max(ScaleLogical(100), _progressPanel.ClientSize.Width - _progressBar.Right - gap), 0);
            _progressLabel.AutoEllipsis = true;
        }

        private void LayoutCurrentStep()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
            return Math.Min(fallbackBottom, safeBottomFromButtons);
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
            panel.Controls.Add(_passwordSeparateToggleCheckBox);

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
                _passwordTextBox.Text = GeneratePasswordValue(GetMinPasswordLength());
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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
            int passwordMinWidth = ScaleLogical(140);
            int inlineStartX = _passwordToggleCheckBox.Right + ScaleLogical(10);
            int inlineAvailable = (left + contentWidth) - inlineStartX;

            int passwordBottom;
            if (inlineAvailable >= passwordMinWidth + ScaleLogical(8) + generateWidth)
            {
                int passwordWidth = Math.Max(passwordMinWidth, inlineAvailable - generateWidth - ScaleLogical(8));
                int rowY = _passwordToggleCheckBox.Top - ScaleLogical(2);
                _passwordTextBox.SetBounds(inlineStartX, rowY, passwordWidth, passwordHeight);
                _passwordGenerateButton.SetBounds(_passwordTextBox.Right + ScaleLogical(8), rowY - ScaleLogical(2), generateWidth, generateHeight);
                passwordBottom = Math.Max(_passwordToggleCheckBox.Bottom, Math.Max(_passwordTextBox.Bottom, _passwordGenerateButton.Bottom));
            }
            else
            {
                int rowY = _passwordToggleCheckBox.Bottom + ScaleLogical(6);
                int passwordWidth = Math.Max(passwordMinWidth, Math.Max(ScaleLogical(120), contentWidth - generateWidth - ScaleLogical(24)));
                _passwordTextBox.SetBounds(left + ScaleLogical(16), rowY, passwordWidth, passwordHeight);
                _passwordGenerateButton.SetBounds(_passwordTextBox.Right + ScaleLogical(8), rowY - ScaleLogical(2), generateWidth, generateHeight);
                passwordBottom = Math.Max(_passwordTextBox.Bottom, _passwordGenerateButton.Bottom);
            }

            _passwordSeparateToggleCheckBox.Location = new Point(left, passwordBottom + sectionGap);

            int requiredHeight = _passwordSeparateToggleCheckBox.Bottom + ScaleLogical(16);
            _generalStepPanel.AutoScrollMinSize = new Size(0, requiredHeight);
        }

        private void LayoutExpirationStep(Size clientSize)
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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

        private void InitializeStepFiles()
        {
            var panel = CreateStepPanel();
            panel.SuspendLayout();

            _fileStepLayout.SuspendLayout();
            _fileStepLayout.ColumnCount = 1;
            _fileStepLayout.RowCount = 3;
            _fileStepLayout.Dock = DockStyle.Fill;
            _fileStepLayout.Padding = new Padding(FileStepPaddingPixels);
            _fileStepLayout.Margin = new Padding(0);
            _fileStepLayout.ColumnStyles.Clear();
            _fileStepLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _fileStepLayout.RowStyles.Clear();
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            panel.Controls.Add(_fileStepLayout);

            _basePathLabel.Text = Strings.FileLinkWizardBasePathPrefix + (_request.BasePath ?? string.Empty);
            _basePathLabel.AutoSize = true;
            _basePathLabel.Margin = new Padding(0);
            _fileStepLayout.Controls.Add(_basePathLabel, 0, 0);

            _attachmentModeInfoLabel.AutoSize = true;
            _attachmentModeInfoLabel.ForeColor = Color.DimGray;
            _attachmentModeInfoLabel.Visible = false;
            _attachmentModeInfoLabel.Margin = new Padding(0, 8, 0, 0);
            _fileStepLayout.Controls.Add(_attachmentModeInfoLabel, 0, 1);

            _fileStepContentLayout.SuspendLayout();
            _fileStepContentLayout.ColumnCount = 2;
            _fileStepContentLayout.RowCount = 1;
            _fileStepContentLayout.Dock = DockStyle.Fill;
            _fileStepContentLayout.Margin = new Padding(0, 12, 0, 0);
            _fileStepContentLayout.ColumnStyles.Clear();
            _fileStepContentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _fileStepContentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FileStepButtonColumnMinWidthPixels));
            _fileStepContentLayout.RowStyles.Clear();
            _fileStepContentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            _fileStepLayout.Controls.Add(_fileStepContentLayout, 0, 2);

            _fileListView.Dock = DockStyle.Fill;
            _fileListView.Margin = new Padding(0, 0, FileStepButtonColumnSpacingPixels, 0);
            _fileListView.View = View.Details;
            _fileListView.FullRowSelect = true;
            _fileListView.HideSelection = false;
            _fileListView.Scrollable = true;
            _fileListView.OwnerDraw = true;
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnPath, 240);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnType, 100);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnStatus, 120);
            _fileListView.Resize += (s, e) => PositionProgressBars();
            _fileListView.DrawColumnHeader += HandleFileListViewDrawColumnHeader;
            _fileListView.DrawItem += HandleFileListViewDrawItem;
            _fileListView.DrawSubItem += HandleFileListViewDrawSubItem;
            _fileListView.HorizontalWheelHandler = HandlePathColumnMouseWheel;
            _fileStepContentLayout.Controls.Add(_fileListView, 0, 0);

            _fileStepActionPanel.FlowDirection = FlowDirection.TopDown;
            _fileStepActionPanel.WrapContents = false;
            _fileStepActionPanel.Dock = DockStyle.Fill;
            _fileStepActionPanel.Margin = new Padding(0);
            _fileStepActionPanel.Padding = new Padding(0);
            _fileStepContentLayout.Controls.Add(_fileStepActionPanel, 1, 0);

            _addFilesButton.Text = Strings.FileLinkWizardAddFilesButton;
            _addFilesButton.AutoSize = false;
            _addFilesButton.Size = new Size(150, 28);
            _addFilesButton.Margin = new Padding(0, 0, 0, FileStepButtonGapPixels);
            _addFilesButton.TextAlign = ContentAlignment.MiddleCenter;
            _addFilesButton.Click += (s, e) => AddFiles();
            _fileStepActionPanel.Controls.Add(_addFilesButton);

            _addFolderButton.Text = Strings.FileLinkWizardAddFolderButton;
            _addFolderButton.AutoSize = false;
            _addFolderButton.Size = new Size(150, 28);
            _addFolderButton.Margin = new Padding(0, 0, 0, FileStepButtonGapPixels);
            _addFolderButton.TextAlign = ContentAlignment.MiddleCenter;
            _addFolderButton.Click += (s, e) => AddFolder();
            _fileStepActionPanel.Controls.Add(_addFolderButton);

            _removeItemButton.Text = Strings.FileLinkWizardRemoveButton;
            _removeItemButton.AutoSize = false;
            _removeItemButton.Size = new Size(150, 28);
            _removeItemButton.Margin = new Padding(0);
            _removeItemButton.TextAlign = ContentAlignment.MiddleCenter;
            _removeItemButton.Click += (s, e) => RemoveSelection();
            _fileStepActionPanel.Controls.Add(_removeItemButton);

            AttachFileQueueDropTarget(panel);
            AttachFileQueueDropTarget(_fileStepLayout);
            AttachFileQueueDropTarget(_fileStepContentLayout);
            AttachFileQueueDropTarget(_fileStepActionPanel);
            AttachFileQueueDropTarget(_fileListView);
            AttachFileQueueDropTarget(_addFilesButton);
            AttachFileQueueDropTarget(_addFolderButton);
            AttachFileQueueDropTarget(_removeItemButton);

            _fileStepContentLayout.ResumeLayout(false);
            _fileStepLayout.ResumeLayout(false);
            _fileStepLayout.PerformLayout();

            panel.ResumeLayout(false);
            panel.PerformLayout();

            panel.ClientSizeChanged += (s, e) => LayoutFileStep(panel.ClientSize);
            LayoutFileStep(panel.ClientSize);
            UpdateQueueColumnWidths();
            PositionProgressBars();

            _steps.Add(panel);
        }

        private void LayoutFileStep(Size clientSize)
        {
            if (_fileStepContentLayout.ColumnStyles.Count >= 2)
            {
                int actionColumnWidth = CalculateFileStepButtonColumnWidth();
                _fileStepContentLayout.ColumnStyles[1].Width = actionColumnWidth;
                ApplyFileStepButtonSize(_addFilesButton, actionColumnWidth);
                ApplyFileStepButtonSize(_addFolderButton, actionColumnWidth);
                ApplyFileStepButtonSize(_removeItemButton, actionColumnWidth);
            }

            int maxInfoWidth = Math.Max(120, clientSize.Width - (FileStepPaddingPixels * 2));
            _attachmentModeInfoLabel.MaximumSize = new Size(maxInfoWidth, 0);

            UpdateQueueColumnWidths();
            PositionProgressBars();
        }

        private void ApplyFileStepButtonSize(Button button, int targetWidth)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (button == null)
            {
                return;
            }

            int minWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(button, out minWidth);
            int width = Math.Max(minWidth, Math.Max(ScaleLogical(120), targetWidth));
            button.Size = new Size(width, button.Height);
        }

        private int CalculateFileStepButtonColumnWidth()
        {
            int textPadding = ScaleLogical(40);
            int minWidth = ScaleLogical(FileStepButtonColumnMinWidthPixels);
            int maxTextWidth = Math.Max(
                Math.Max(
                    TextRenderer.MeasureText(_addFilesButton.Text ?? string.Empty, _addFilesButton.Font).Width,
                    TextRenderer.MeasureText(_addFolderButton.Text ?? string.Empty, _addFolderButton.Font).Width),
                TextRenderer.MeasureText(_removeItemButton.Text ?? string.Empty, _removeItemButton.Font).Width);

            return Math.Max(minWidth, maxTextWidth + textPadding);
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

            _progressBar.Location = new Point(0, 16);
            _progressBar.Size = new Size(400, 16);
            _progressPanel.Controls.Add(_progressBar);

            _progressLabel.Location = new Point(410, 16);
            _progressLabel.AutoSize = true;
            _progressPanel.Controls.Add(_progressLabel);

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

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_stepHost != null)
            {
                _stepHost.Controls.Add(panel);
                panel.BringToFront();
            }

            return panel;
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
                        string.Format(CultureInfo.CurrentCulture, Strings.TalkPasswordTooShort, GetMinPasswordLength()),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    _passwordTextBox.Focus();
                    return false;
                }

                if (!EnsureShareFolderAvailable())
                {
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

        private sealed class SelectionUploadState
        {
            internal SelectionUploadState(ListViewItem item)
            {
                Item = item;
                Status = FileLinkUploadStatus.Pending;
            }

            internal ListViewItem Item { get; private set; }

            internal ProgressBar ProgressBar { get; set; }

            internal long TotalBytes { get; set; }

            internal long UploadedBytes { get; set; }

            internal FileLinkUploadStatus Status { get; set; }

            internal string RenamedTo { get; set; }
        }

        private sealed class PathScrollableListView : ListView
        {
            private const int WmMouseWheel = 0x020A;

            internal Func<int, bool> HorizontalWheelHandler { get; set; }

            protected override void WndProc(ref Message m)
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (m.Msg == WmMouseWheel && HorizontalWheelHandler != null)
                {
                    long wParam = m.WParam.ToInt64();
                    int delta = unchecked((short)((wParam >> 16) & 0xffff));
                    if (delta != 0 && HorizontalWheelHandler(delta))
                    {
                        return;
                    }
                }

                base.WndProc(ref m);
            }
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

                // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
                if (_allowEmptyUpload && (_uploadContext == null || !_uploadCompleted))
                {
                    _uploadContext = _service.PrepareUpload(_request, CancellationToken.None);
                    _uploadCompleted = true;
                }
            }

            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
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
            _progressPanel.Visible = true;
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
                _progressPanel.Visible = false;
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
                _progressPanel.Visible = false;
                _finishButton.Enabled = true;
            }
            finally
            {
                UseWaitCursor = false;
                Cursor.Current = previousCursor;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }
                _progressBar.Style = ProgressBarStyle.Blocks;
                _cancelButton.Enabled = true;
                UpdateNavigationState();
                UpdateUploadButtonState();
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
                && HasBackendSeatEntitlement()
                && _passwordSeparateToggleCheckBox.Checked;
            _request.ExpireEnabled = _expireToggleCheckBox.Checked;
            _request.ExpireDate = _expireToggleCheckBox.Checked ? _expireDatePicker.Value.Date : (DateTime?)null;
            _request.NoteEnabled = !_attachmentMode && _noteToggleCheckBox.Checked;
            _request.Note = !_attachmentMode && _noteToggleCheckBox.Checked ? _noteTextBox.Text.Trim() : null;
            _request.AttachmentMode = _attachmentMode;
            _request.ShareDate = _shareDate;
            _request.ShareDatePrefixFormat = _attachmentMode ? AttachmentShareDatePrefixFormat : "yyyyMMdd";

            _request.Items.Clear();
            foreach (var item in _items)
            {
                _request.Items.Add(item);
            }
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
                ExpireEnabled = source.ExpireEnabled,
                ExpireDate = source.ExpireDate,
                NoteEnabled = source.NoteEnabled,
                Note = source.Note,
                AttachmentMode = source.AttachmentMode,
                ShareDate = source.ShareDate,
                ShareDatePrefixFormat = source.ShareDatePrefixFormat
            };
            foreach (var item in source.Items)
            {
                clone.Items.Add(new FileLinkSelection(item.SelectionType, item.LocalPath));
            }
            return clone;
        }

        private bool EnsureShareFolderAvailable()
        {
            string shareNameInput = _shareNameTextBox.Text.Trim();
            string sanitizedShareName = FileLinkService.SanitizeComponent(shareNameInput);
            if (string.IsNullOrWhiteSpace(sanitizedShareName))
            {
                sanitizedShareName = Strings.FileLinkWizardFallbackShareName;
            }

            string folderName = BuildShareFolderName(sanitizedShareName);

            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UseWaitCursor = true;

                if (_service.FolderExists(_request.BasePath, folderName, CancellationToken.None))
                {
                    MessageBox.Show(
                        string.Format(CultureInfo.CurrentCulture, Strings.FileLinkWizardFolderExistsFormat, folderName),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    _shareNameTextBox.Focus();
                    _shareNameTextBox.SelectAll();
                    return false;
                }
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Share folder existence check failed.", ex);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.FileLinkWizardFolderCheckFailedFormat, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    ex.IsAuthenticationError ? MessageBoxIcon.Warning : MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                UseWaitCursor = false;
                Cursor.Current = previousCursor;
            }

            return true;
        }

        private string BuildShareFolderName(string sanitizedShareName)
        {
            string prefixFormat = FileLinkService.NormalizeShareDatePrefixFormat(_request.ShareDatePrefixFormat);
            return _shareDate.ToString(prefixFormat, CultureInfo.InvariantCulture) + "_" + sanitizedShareName;
        }

        private void TryCleanupUnfinalizedUploadContext(string reason)
        {
            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
            if (_shareFinalized || _uploadContext == null)
            {
                return;
            }

            FileLinkUploadContext context = _uploadContext;
            _uploadContext = null;
            _uploadCompleted = false;

            TryCleanupPreparedUploadContext(context, reason);
        }

        private void TryCleanupPreparedUploadContext(FileLinkUploadContext context, string reason)
        {
            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
            if (context == null)
            {
                return;
            }

            string relativeFolderPath = context.RelativeFolderPath ?? string.Empty;
            if (string.IsNullOrWhiteSpace(relativeFolderPath))
            {
                return;
            }

            try
            {
                _service.DeleteShareFolder(relativeFolderPath, CancellationToken.None);
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Wizard upload context cleanup succeeded (reason="
                    + (reason ?? string.Empty)
                    + ", relativeFolder="
                    + relativeFolderPath
                    + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Wizard upload context cleanup failed (reason="
                    + (reason ?? string.Empty)
                    + ", relativeFolder="
                    + relativeFolderPath
                    + ").",
                    ex);
            }
        }

        private void LoadInitialSelections()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
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

            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UseWaitCursor = true;
                ResolveAttachmentShareName();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve attachment-mode share name.", ex);
                MessageBox.Show(
                    ex.Message,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                UseWaitCursor = false;
                Cursor.Current = previousCursor;
            }

            ShowStep(2);
        }

        private string BuildAttachmentModeInfoText()
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

        private void ResolveAttachmentShareName()
        {
            if (!_attachmentMode)
            {
                return;
            }

            for (int suffix = 0; suffix < 1000; suffix++)
            {
                string candidate = suffix == 0
                    ? AttachmentShareNameBase
                    : AttachmentShareNameBase + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                string folderName = BuildShareFolderName(candidate);
                bool exists = _service.FolderExists(_request.BasePath, folderName, CancellationToken.None);
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Attachment mode share name check: candidate="
                    + candidate
                    + ", folder="
                    + folderName
                    + ", exists="
                    + exists.ToString(CultureInfo.InvariantCulture));

                if (!exists)
                {
                    _shareNameTextBox.Text = candidate;
                    return;
                }
            }

            throw new TalkServiceException(Strings.FileLinkWizardFolderExistsFormat, false, 0, null);
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
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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

        private void UpdateProgress(FileLinkProgress progress)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<FileLinkProgress>(UpdateProgress), progress);
                return;
            }

            if (progress.TotalBytes > 0)
            {
                double percent = (double)progress.UploadedBytes / progress.TotalBytes;
                _progressBar.Value = Math.Min(_progressBar.Maximum, (int)(percent * _progressBar.Maximum));
            }
            _progressLabel.Text = Path.GetFileName(progress.CurrentItem ?? string.Empty);
        }

        private void UpdatePasswordState()
        {
            bool enabled = _passwordToggleCheckBox.Checked;
            bool lockPasswordSeparate = IsPolicyLocked("share_send_password_separately");
            bool separatePasswordAvailable = HasBackendSeatEntitlement();
            string separatePasswordUnavailableTooltip = GetSeparatePasswordUnavailableTooltip();
            _passwordTextBox.Enabled = enabled;
            _passwordGenerateButton.Enabled = enabled;
            _passwordSeparateToggleCheckBox.Enabled = enabled && !lockPasswordSeparate && separatePasswordAvailable;
            SetTooltipWithFallback(
                _passwordSeparateToggleCheckBox,
                !separatePasswordAvailable
                    ? separatePasswordUnavailableTooltip
                    : (lockPasswordSeparate ? Strings.PolicyAdminControlledTooltip : string.Empty),
                !separatePasswordAvailable || lockPasswordSeparate,
                _passwordTextBox);
            if (!enabled || !separatePasswordAvailable)
            {
                _passwordSeparateToggleCheckBox.Checked = false;
            }

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_expirationStepPanel != null)
            {
                LayoutExpirationStep(_expirationStepPanel.ClientSize);
            }
        }

        private void UpdateNoteState()
        {
            _noteTextBox.Enabled = _noteToggleCheckBox.Checked;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_noteStepPanel != null)
            {
                LayoutNoteStep(_noteStepPanel.ClientSize);
            }
        }

        private void AddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var selections = new List<FileLinkSelection>();
                    foreach (string file in dialog.FileNames)
                    {
                        selections.Add(new FileLinkSelection(FileLinkSelectionType.File, file));
                    }

                    AddSelections(selections);
                }
            }
        }

        private void AddFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    AddSelections(new[] { new FileLinkSelection(FileLinkSelectionType.Directory, dialog.SelectedPath) });
                }
            }
        }

        private void RemoveSelection()
        {
            if (_fileListView.SelectedItems.Count == 0)
            {
                return;
            }

            foreach (ListViewItem item in _fileListView.SelectedItems)
            {
                if (ReferenceEquals(_lastAutoScrolledUploadItem, item))
                {
                    _lastAutoScrolledUploadItem = null;
                }

                FileLinkSelection selection = item.Tag as FileLinkSelection;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (selection != null)
                {
                    _items.Remove(selection);
                    SelectionUploadState state;
                    if (_selectionStates.TryGetValue(selection, out state))
                    {
                        DisposeStateProgressBar(state);
                        _selectionStates.Remove(selection);
                    }
                }
                _fileListView.Items.Remove(item);
            }

            if (_items.Count == 0)
            {
                _allowEmptyUpload = false;
            }
            UpdateQueueColumnWidths();
            PositionProgressBars();
            InvalidateUpload();
        }

        private void AddSelections(IEnumerable<FileLinkSelection> selections)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (selections == null)
            {
                return;
            }

            var pendingSelections = selections.Where(s => s != null && !string.IsNullOrWhiteSpace(s.LocalPath)).ToList();
            if (pendingSelections.Count == 0)
            {
                return;
            }

            var existingPaths = _attachmentMode
                ? null
                : new HashSet<string>(
                    _items.Select(i => i.LocalPath ?? string.Empty),
                    StringComparer.OrdinalIgnoreCase);

            int requestedCount = pendingSelections.Count;
            int addedCount = 0;

            _fileListView.BeginUpdate();
            try
            {
                foreach (var selection in pendingSelections)
                {
                    if (TryAddSelection(selection, existingPaths))
                    {
                        addedCount++;
                    }
                }
            }
            finally
            {
                _fileListView.EndUpdate();
            }

            if (addedCount == 0)
            {
                return;
            }

            _allowEmptyUpload = false;
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Queue selections added (requested="
                + requestedCount.ToString(CultureInfo.InvariantCulture)
                + ", added="
                + addedCount.ToString(CultureInfo.InvariantCulture)
                + ", total="
                + _items.Count.ToString(CultureInfo.InvariantCulture)
                + ").");

            UpdateQueueColumnWidths();
            PositionProgressBars();
            InvalidateUpload();
        }

        private void UpdateQueueColumnWidths()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_fileListView == null || _fileListView.Columns.Count < 3 || _fileListView.IsDisposed || _fileListView.Disposing)
            {
                return;
            }

            int typeWidth = 110;
            int statusWidth = 180;
            int clientWidth = Math.Max(0, _fileListView.ClientSize.Width);
            int pathWidth = clientWidth - typeWidth - statusWidth - 6;
            if (pathWidth < 120)
            {
                int shortage = 120 - pathWidth;
                int reducibleStatus = Math.Max(0, statusWidth - 150);
                int reduceStatus = Math.Min(shortage, reducibleStatus);
                statusWidth -= reduceStatus;
                shortage -= reduceStatus;

                int reducibleType = Math.Max(0, typeWidth - 90);
                int reduceType = Math.Min(shortage, reducibleType);
                typeWidth -= reduceType;

                pathWidth = Math.Max(90, clientWidth - typeWidth - statusWidth - 6);
            }

            _fileListView.Columns[0].Width = pathWidth;
            _fileListView.Columns[1].Width = typeWidth;
            _fileListView.Columns[2].Width = statusWidth;
            UpdatePathColumnScrollRange();
            _fileListView.Invalidate();
        }

        private bool HandlePathColumnMouseWheel(int delta)
        {
            if (delta == 0)
            {
                return false;
            }

            if (_pathColumnMaxHorizontalOffset <= 0)
            {
                _pathColumnHorizontalOffset = 0;
                return false;
            }

            int steps = Math.Max(1, Math.Abs(delta) / 120);
            int shift = steps * PathColumnWheelStepPixels;
            int nextOffset = _pathColumnHorizontalOffset + (delta < 0 ? shift : -shift);
            if (nextOffset < 0)
            {
                nextOffset = 0;
            }
            else if (nextOffset > _pathColumnMaxHorizontalOffset)
            {
                nextOffset = _pathColumnMaxHorizontalOffset;
            }

            if (nextOffset == _pathColumnHorizontalOffset)
            {
                return true;
            }

            _pathColumnHorizontalOffset = nextOffset;
            _fileListView.Invalidate();
            return true;
        }

        private void UpdatePathColumnScrollRange()
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_fileListView == null || _fileListView.IsDisposed || _fileListView.Disposing || _fileListView.Columns.Count == 0)
            {
                _pathColumnHorizontalOffset = 0;
                _pathColumnMaxHorizontalOffset = 0;
                return;
            }

            int visibleWidth = Math.Max(0, _fileListView.Columns[0].Width - 8);
            if (_fileListView.Items.Count == 0 || visibleWidth <= 0)
            {
                _pathColumnHorizontalOffset = 0;
                _pathColumnMaxHorizontalOffset = 0;
                return;
            }

            int widestPath = 0;
            foreach (ListViewItem item in _fileListView.Items)
            {
                string path = item != null ? (item.Text ?? string.Empty) : string.Empty;
                if (path.Length == 0)
                {
                    continue;
                }

                int width = TextRenderer.MeasureText(
                    path,
                    _fileListView.Font,
                    new Size(int.MaxValue, int.MaxValue),
                    TextFormatFlags.SingleLine | TextFormatFlags.NoPadding).Width;

                if (width > widestPath)
                {
                    widestPath = width;
                }
            }

            _pathColumnMaxHorizontalOffset = Math.Max(0, widestPath - visibleWidth);
            if (_pathColumnHorizontalOffset > _pathColumnMaxHorizontalOffset)
            {
                _pathColumnHorizontalOffset = _pathColumnMaxHorizontalOffset;
            }
        }

        private void HandleFileListViewDrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (e == null)
            {
                return;
            }

            e.DrawDefault = true;
        }

        private void HandleFileListViewDrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (e == null)
            {
                return;
            }

            if (_fileListView.View != View.Details)
            {
                e.DrawDefault = true;
            }
        }

        private void HandleFileListViewDrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (e == null || e.Item == null || e.SubItem == null)
            {
                return;
            }

            Color rowBackColor = e.Item.BackColor.IsEmpty ? _fileListView.BackColor : e.Item.BackColor;
            Color rowTextColor = e.Item.ForeColor.IsEmpty ? _fileListView.ForeColor : e.Item.ForeColor;
            Color backColor = e.SubItem.BackColor.IsEmpty ? rowBackColor : e.SubItem.BackColor;
            Color textColor = e.SubItem.ForeColor.IsEmpty ? rowTextColor : e.SubItem.ForeColor;

            bool selected = e.Item.Selected && (!_fileListView.HideSelection || _fileListView.Focused);
            if (selected)
            {
                backColor = _themePalette != null ? _themePalette.SelectionBackground : SystemColors.Highlight;
                textColor = _themePalette != null ? _themePalette.SelectionText : SystemColors.HighlightText;
            }

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }

            string text = e.SubItem.Text ?? string.Empty;
            Rectangle textBounds = new Rectangle(
                e.Bounds.Left + 4,
                e.Bounds.Top + 1,
                Math.Max(0, e.Bounds.Width - 6),
                Math.Max(0, e.Bounds.Height - 2));

            if (e.ColumnIndex == 0)
            {
                var state = e.Graphics.Save();
                e.Graphics.SetClip(e.Bounds);
                var shiftedTextBounds = new Rectangle(
                    textBounds.Left - _pathColumnHorizontalOffset,
                    textBounds.Top,
                    textBounds.Width + _pathColumnHorizontalOffset,
                    textBounds.Height);
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    _fileListView.Font,
                    shiftedTextBounds,
                    textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
                e.Graphics.Restore(state);
            }
            else
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    _fileListView.Font,
                    textBounds,
                    textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            }

            if (e.ColumnIndex == _fileListView.Columns.Count - 1 && selected)
            {
                Rectangle focusRect = e.Item.Bounds;
                focusRect.Width = Math.Max(0, _fileListView.ClientSize.Width - focusRect.Left);
                ControlPaint.DrawFocusRectangle(e.Graphics, focusRect, textColor, backColor);
            }
        }

        private void ApplyQueueRowStyle(SelectionUploadState state, Color backgroundColor, Color textColor)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (state == null || state.Item == null)
            {
                return;
            }

            ListViewItem item = state.Item;
            item.BackColor = backgroundColor;
            item.ForeColor = textColor;

            for (int i = 0; i < item.SubItems.Count; i++)
            {
                item.SubItems[i].BackColor = backgroundColor;
                if (i != 2)
                {
                    item.SubItems[i].ForeColor = textColor;
                }
            }
        }

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

        private ProgressBar CreateProgressBar()
        {
            var bar = new ProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Style = ProgressBarStyle.Continuous,
                Visible = false
            };
            _fileListView.Controls.Add(bar);
            return bar;
        }

        private void DisposeStateProgressBar(SelectionUploadState state)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (state == null || state.ProgressBar == null)
            {
                return;
            }

            var bar = state.ProgressBar;
            state.ProgressBar = null;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_fileListView != null && !_fileListView.IsDisposed && !_fileListView.Disposing)
            {
                _fileListView.Controls.Remove(bar);
            }

            bar.Dispose();
        }

        private void InvalidateUpload()
        {
            FileLinkUploadContext previousContext = _uploadContext;
            _uploadCompleted = false;
            _uploadContext = null;
            _lastAutoScrolledUploadItem = null;
            ResetUploadProgressPump();
            if (_items.Count > 0)
            {
                _allowEmptyUpload = false;
            }

            // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
            if (!_shareFinalized && previousContext != null)
            {
                TryCleanupPreparedUploadContext(previousContext, "upload_invalidated");
            }

            foreach (var state in _selectionStates.Values)
            {
                state.TotalBytes = 0;
                state.UploadedBytes = 0;
                state.Status = FileLinkUploadStatus.Pending;
                state.RenamedTo = null;
                ApplyQueueRowStyle(state, _themePalette.InputBackground, _themePalette.Text);
                DisposeStateProgressBar(state);
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = string.Empty;
                    state.Item.SubItems[2].ForeColor = _themePalette.Text;
                }
            }

            UpdateNavigationState();
            UpdateUploadButtonState();
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

        /**
         * Positions or hides the per-item progress bars next to each list entry.
         * If items are missing, existing bars are hidden to avoid exceptions.
         */
        private void PositionProgressBars()
        {
            if (_selectionStates.Count == 0 || _fileListView.Columns.Count < 3)
            {
                return;
            }

            if (_fileListView.IsDisposed || _fileListView.Disposing || !_fileListView.IsHandleCreated)
            {
                return;
            }

            if (_fileListView.Items.Count == 0)
            {
                foreach (var state in _selectionStates.Values)
                {
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (state.ProgressBar != null)
                    {
                        state.ProgressBar.Visible = false;
                    }
                }
                return;
            }

            int statusLeft = _fileListView.Columns[0].Width + _fileListView.Columns[1].Width;
            int statusWidth = _fileListView.Columns[2].Width;

            foreach (var state in _selectionStates.Values)
            {
                PositionProgressBar(state, statusLeft, statusWidth);
            }
        }

        private void PositionProgressBar(SelectionUploadState state)
        {
            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_fileListView == null || _fileListView.Columns.Count < 3)
            {
                return;
            }

            int statusLeft = _fileListView.Columns[0].Width + _fileListView.Columns[1].Width;
            int statusWidth = _fileListView.Columns[2].Width;
            PositionProgressBar(state, statusLeft, statusWidth);
        }

        private void PositionProgressBar(SelectionUploadState state, int statusLeft, int statusWidth)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (state == null || state.ProgressBar == null)
            {
                return;
            }

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (state.ProgressBar.IsDisposed || state.ProgressBar.Disposing || state.Item == null)
            {
                state.ProgressBar.Visible = false;
                return;
            }

            if (state.Item.ListView != _fileListView)
            {
                state.ProgressBar.Visible = false;
                return;
            }

            int itemIndex = state.Item.Index;
            Rectangle bounds;
            if (!TryGetListViewItemBounds(itemIndex, out bounds))
            {
                state.ProgressBar.Visible = false;
                return;
            }

            if (bounds.Height <= 0 || bounds.Width <= 0)
            {
                state.ProgressBar.Visible = false;
                return;
            }

            int left = statusLeft + 4;
            int width = Math.Max(12, statusWidth - 8);
            int top = bounds.Top + 3;
            int height = Math.Max(6, bounds.Height - 6);
            state.ProgressBar.SetBounds(left, top, width, height);
            state.ProgressBar.Visible = state.Status == FileLinkUploadStatus.Uploading;
        }

        private bool TryGetListViewItemBounds(int index, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;

            // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
            if (_fileListView == null || _fileListView.IsDisposed || _fileListView.Disposing)
            {
                return false;
            }

            if (index < 0 || index >= _fileListView.Items.Count)
            {
                return false;
            }

            try
            {
                bounds = _fileListView.GetItemRect(index, ItemBoundsPortion.Entire);
                return true;
            }
            catch (ArgumentException ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Skipped progress bar positioning because list item bounds are unavailable (index="
                    + index.ToString(CultureInfo.InvariantCulture)
                    + ").",
                    ex);
                return false;
            }
        }

        private int ScaleLogical(int value)
        {
            int dpi = DeviceDpi > 0 ? DeviceDpi : 96;
            return (int)Math.Round(value * (dpi / 96f));
        }

        private void ReflowWizardLayout()
        {
            LayoutBottomButtons();
            UpdateStepHostBounds();
            LayoutCurrentStep();
            LayoutProgressPanel();
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

        private async Task StartUploadAsync()
        {
            if (_uploadInProgress || _items.Count == 0)
            {
                return;
            }

            ApplyFormData();
            _uploadCompleted = false;
            _progressPanel.Visible = false;
            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressLabel.Text = string.Empty;
            ToggleUpload(true);

            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            UseWaitCursor = true;
            FileLinkUploadContext preparedContext = null;

            try
            {
                ResetUploadProgressPump();
                _lastAutoScrolledUploadItem = null;
                foreach (var state in _selectionStates.Values)
                {
                    state.TotalBytes = 0;
                    state.UploadedBytes = 0;
                    state.Status = FileLinkUploadStatus.Pending;
                    DisposeStateProgressBar(state);
                    if (state.Item.SubItems.Count >= 3)
                    {
                        state.Item.SubItems[2].Text = string.Empty;
                        state.Item.SubItems[2].ForeColor = _themePalette.Text;
                    }
                }
                PositionProgressBars();

                _cancellationSource = new CancellationTokenSource();
                CancellationToken token = _cancellationSource.Token;

                preparedContext = _service.PrepareUpload(_request, token);
                var progress = new Progress<FileLinkUploadItemProgress>(HandleUploadProgress);

                await Task.Run(() =>
                {
                    _service.UploadSelections(preparedContext, _items, progress, HandleDuplicate, token);
                });

                _uploadContext = preparedContext;
                _uploadCompleted = true;
                UpdateNavigationState();
                UpdateUploadButtonState();
            }
            catch (OperationCanceledException)
            {
                DiagnosticsLogger.Log(LogCategories.FileLink, "Upload cancelled.");
                _lastAutoScrolledUploadItem = null;
                foreach (var state in _selectionStates.Values)
                {
                    state.Status = FileLinkUploadStatus.Failed;
                    ApplyQueueRowStyle(state, _themePalette.InputBackground, _themePalette.Text);
                    DisposeStateProgressBar(state);
                    if (state.Item.SubItems.Count >= 3)
                    {
                        state.Item.SubItems[2].Text = Strings.FileLinkWizardStatusCancelled;
                        state.Item.SubItems[2].ForeColor = _themePalette.ErrorText;
                    }
                }
                FlushBufferedUploadProgress();
                ShowUploadError(Strings.FileLinkWizardUploadCancelledMessage);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Upload failed with service error.", ex);
                FlushBufferedUploadProgress();
                ShowUploadError(ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Upload failed unexpectedly.", ex);
                FlushBufferedUploadProgress();
                ShowUploadError(ex.Message);
            }
            finally
            {
                FlushBufferedUploadProgress();
                ResetUploadProgressPump();
                _lastAutoScrolledUploadItem = null;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }

                // Kontext kann außerhalb des UI-Threads fehlen; Guard verhindert Folgefehler im Shutdown/Background-Pfad.
                if (!_uploadCompleted && preparedContext != null)
                {
                    TryCleanupPreparedUploadContext(preparedContext, "upload_not_completed");
                }

                UseWaitCursor = false;
                Cursor.Current = previousCursor;
                ToggleUpload(false);
                PositionProgressBars();
            }
        }

        private void ToggleUpload(bool uploading)
        {
            _uploadInProgress = uploading;
            UpdateNavigationState();
            UpdateUploadButtonState();
            _cancelButton.Enabled = !uploading;
        }

        /**
         * Asks whether to continue without prepared uploads.
         */
        private bool ConfirmEmptyUploadProceed()
        {
            using (var dialog = new Form())
            {
                dialog.Text = Strings.DialogTitle;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.AutoScaleMode = AutoScaleMode.Dpi;
                dialog.AutoScaleDimensions = new SizeF(96f, 96f);
                dialog.ClientSize = new Size(420, 150);
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.Icon = BrandingAssets.GetAppIcon(32);

                var label = new Label
                {
                    Text = Strings.FileLinkNoFilesConfirm,
                    AutoSize = true,
                    Location = new Point(15, 15)
                };
                label.MaximumSize = new Size(Math.Max(260, dialog.ClientSize.Width - 30), 0);
                dialog.Controls.Add(label);

                var continueButton = new Button
                {
                    Text = Strings.ButtonNext,
                    DialogResult = DialogResult.OK
                };
                var backButton = new Button
                {
                    Text = Strings.ButtonBack,
                    DialogResult = DialogResult.Cancel
                };
                int ignoredNextMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(continueButton, out ignoredNextMinWidth);
                int ignoredBackMinWidth;
                FooterButtonLayoutHelper.ApplyButtonSize(backButton, out ignoredBackMinWidth);

                dialog.Controls.Add(continueButton);
                dialog.Controls.Add(backButton);
                dialog.AcceptButton = continueButton;
                dialog.CancelButton = backButton;

                Action layoutDialog = () =>
                {
                    int outerPadding = 15;
                    int verticalGap = 14;
                    label.MaximumSize = new Size(Math.Max(260, dialog.ClientSize.Width - (outerPadding * 2)), 0);
                    label.Location = new Point(outerPadding, outerPadding);

                    int requiredClientWidth = FooterButtonLayoutHelper.LayoutCentered(
                        dialog,
                        new[] { continueButton, backButton },
                        FooterButtonLayoutHelper.DefaultHorizontalPadding,
                        FooterButtonLayoutHelper.DefaultBottomPadding,
                        FooterButtonLayoutHelper.DefaultSpacing,
                        true);
                    if (requiredClientWidth > dialog.ClientSize.Width)
                    {
                        dialog.ClientSize = new Size(requiredClientWidth, dialog.ClientSize.Height);
                        label.MaximumSize = new Size(Math.Max(260, dialog.ClientSize.Width - (outerPadding * 2)), 0);
                        label.Location = new Point(outerPadding, outerPadding);
                        FooterButtonLayoutHelper.LayoutCentered(
                            dialog,
                            new[] { continueButton, backButton },
                            FooterButtonLayoutHelper.DefaultHorizontalPadding,
                            FooterButtonLayoutHelper.DefaultBottomPadding,
                            FooterButtonLayoutHelper.DefaultSpacing,
                            true);
                    }

                    int buttonsTop = Math.Min(continueButton.Top, backButton.Top);
                    int requiredHeight = label.Bottom + verticalGap + continueButton.Height + FooterButtonLayoutHelper.DefaultBottomPadding;
                    if (requiredHeight > dialog.ClientSize.Height)
                    {
                        dialog.ClientSize = new Size(dialog.ClientSize.Width, requiredHeight);
                        FooterButtonLayoutHelper.LayoutCentered(
                            dialog,
                            new[] { continueButton, backButton },
                            FooterButtonLayoutHelper.DefaultHorizontalPadding,
                            FooterButtonLayoutHelper.DefaultBottomPadding,
                            FooterButtonLayoutHelper.DefaultSpacing,
                            true);
                    }
                };

                UiThemeManager.ApplyToForm(dialog);
                layoutDialog();

                bool result = dialog.ShowDialog(this) == DialogResult.OK;
                if (!result)
                {
                    _allowEmptyUpload = false;
                }
                return result;
            }
        }

        private void HandleUploadProgress(FileLinkUploadItemProgress progress)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (progress == null || progress.Selection == null)
            {
                return;
            }

            bool activatePump = false;
            lock (_uploadProgressSync)
            {
                if (_pendingUploadProgress.ContainsKey(progress.Selection))
                {
                    _pendingUploadProgressOrder.Remove(progress.Selection);
                }

                _pendingUploadProgress[progress.Selection] = progress;
                _pendingUploadProgressOrder.Add(progress.Selection);
                if (!_uploadProgressPumpRequested)
                {
                    _uploadProgressPumpRequested = true;
                    activatePump = true;
                }
            }

            if (!activatePump)
            {
                return;
            }

            EnsureUploadProgressPumpRunning();
        }

        private void EnsureUploadProgressPumpRunning()
        {
            if (IsDisposed || Disposing)
            {
                ResetUploadProgressPump();
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(EnsureUploadProgressPumpRunning));
                return;
            }

            if (!_uploadProgressFlushTimer.Enabled)
            {
                _uploadProgressFlushTimer.Start();
            }
        }

        private void FlushBufferedUploadProgress()
        {
            if (IsDisposed || Disposing)
            {
                ResetUploadProgressPump();
                return;
            }

            List<FileLinkUploadItemProgress> snapshot;
            lock (_uploadProgressSync)
            {
                if (_pendingUploadProgress.Count == 0)
                {
                    _pendingUploadProgressOrder.Clear();
                    _uploadProgressPumpRequested = false;
                    if (_uploadProgressFlushTimer.Enabled)
                    {
                        _uploadProgressFlushTimer.Stop();
                    }
                    return;
                }

                snapshot = new List<FileLinkUploadItemProgress>(_pendingUploadProgressOrder.Count);
                foreach (var selection in _pendingUploadProgressOrder)
                {
                    FileLinkUploadItemProgress queuedProgress;
                    if (_pendingUploadProgress.TryGetValue(selection, out queuedProgress))
                    {
                        snapshot.Add(queuedProgress);
                    }
                }

                _pendingUploadProgress.Clear();
                _pendingUploadProgressOrder.Clear();
            }

            foreach (var progress in snapshot)
            {
                ApplyUploadProgress(progress);
            }

            ListViewItem activeUploadItem = null;
            for (int i = snapshot.Count - 1; i >= 0; i--)
            {
                FileLinkUploadItemProgress queued = snapshot[i];
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (queued != null && queued.Status == FileLinkUploadStatus.Uploading)
                {
                    SelectionUploadState activeState;
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (_selectionStates.TryGetValue(queued.Selection, out activeState) && activeState != null)
                    {
                        activeUploadItem = activeState.Item;
                        break;
                    }
                }
            }

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (activeUploadItem != null)
            {
                EnsureUploadItemVisible(activeUploadItem, true);
            }

            lock (_uploadProgressSync)
            {
                if (_pendingUploadProgress.Count == 0)
                {
                    _uploadProgressPumpRequested = false;
                    if (_uploadProgressFlushTimer.Enabled)
                    {
                        _uploadProgressFlushTimer.Stop();
                    }
                }
            }
        }

        private void ResetUploadProgressPump()
        {
            lock (_uploadProgressSync)
            {
                _pendingUploadProgress.Clear();
                _pendingUploadProgressOrder.Clear();
                _uploadProgressPumpRequested = false;
            }

            if (_uploadProgressFlushTimer.Enabled)
            {
                _uploadProgressFlushTimer.Stop();
            }
        }

        private void ApplyUploadProgress(FileLinkUploadItemProgress progress)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (progress == null || progress.Selection == null)
            {
                return;
            }

            SelectionUploadState state;
            if (!_selectionStates.TryGetValue(progress.Selection, out state))
            {
                return;
            }

            state.TotalBytes = progress.TotalBytes;
            state.UploadedBytes = progress.UploadedBytes;
            state.Status = progress.Status;

            if (progress.Status == FileLinkUploadStatus.Uploading)
            {
                ApplyQueueRowStyle(state, ActiveQueueItemBackground, ActiveQueueItemText);
                EnsureUploadItemVisible(state.Item, false);
                int percent = state.TotalBytes > 0 ? (int)Math.Min(100, (state.UploadedBytes * 100L) / state.TotalBytes) : 100;
                percent = Math.Max(0, Math.Min(100, percent));
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (state.ProgressBar == null)
                {
                    state.ProgressBar = CreateProgressBar();
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = true;
                    state.ProgressBar.Value = percent;
                }
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = percent.ToString(CultureInfo.InvariantCulture) + "%";
                    state.Item.SubItems[2].ForeColor = ActiveQueueItemText;
                }

                PositionProgressBar(state);
            }
            else if (progress.Status == FileLinkUploadStatus.Completed)
            {
                if (ReferenceEquals(_lastAutoScrolledUploadItem, state.Item))
                {
                    _lastAutoScrolledUploadItem = null;
                }
                EnsureUploadItemVisible(state.Item, false);
                ApplyQueueRowStyle(state, _themePalette.InputBackground, _themePalette.Text);
                DisposeStateProgressBar(state);
                string statusText = Strings.FileLinkWizardStatusSuccess;
                if (!string.IsNullOrEmpty(state.RenamedTo))
                {
                    statusText += " \u2192 " + state.RenamedTo;
                }
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = statusText;
                    state.Item.SubItems[2].ForeColor = _themePalette.SuccessText;
                }
            }
            else if (progress.Status == FileLinkUploadStatus.Failed)
            {
                if (ReferenceEquals(_lastAutoScrolledUploadItem, state.Item))
                {
                    _lastAutoScrolledUploadItem = null;
                }
                EnsureUploadItemVisible(state.Item, false);
                ApplyQueueRowStyle(state, _themePalette.InputBackground, _themePalette.Text);
                DisposeStateProgressBar(state);
                string message = string.IsNullOrWhiteSpace(progress.Message) ? string.Empty : " (" + progress.Message + ")";
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = Strings.FileLinkWizardStatusError + message;
                    state.Item.SubItems[2].ForeColor = _themePalette.ErrorText;
                }
            }
        }

        private void EnsureUploadItemVisible(ListViewItem item, bool forceScroll)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (item == null || item.ListView != _fileListView || _fileListView.IsDisposed || _fileListView.Disposing)
            {
                return;
            }

            int index = item.Index;
            if (index < 0 || index >= _fileListView.Items.Count)
            {
                return;
            }

            if (!forceScroll && ReferenceEquals(_lastAutoScrolledUploadItem, item))
            {
                return;
            }

            if (!forceScroll && IsQueueItemVisible(index))
            {
                _lastAutoScrolledUploadItem = item;
                return;
            }

            _lastAutoScrolledUploadItem = item;
            _fileListView.EnsureVisible(index);
        }

        private bool IsQueueItemVisible(int itemIndex)
        {
            Rectangle bounds;
            if (!TryGetListViewItemBounds(itemIndex, out bounds))
            {
                return false;
            }

            int top = bounds.Top;
            int bottom = bounds.Bottom;
            int viewTop = 0;
            int viewBottom = _fileListView.ClientSize.Height;
            return top >= viewTop && bottom <= viewBottom;
        }

        private void ShowUploadError(string message)
        {
            _uploadCompleted = false;
            _uploadContext = null;
            _lastAutoScrolledUploadItem = null;
            ResetUploadProgressPump();
            foreach (var state in _selectionStates.Values)
            {
                if (state.Status == FileLinkUploadStatus.Uploading)
                {
                    state.Status = FileLinkUploadStatus.Failed;
                    ApplyQueueRowStyle(state, _themePalette.InputBackground, _themePalette.Text);
                    DisposeStateProgressBar(state);
                    if (state.Item.SubItems.Count >= 3)
                    {
                        state.Item.SubItems[2].Text = Strings.FileLinkWizardStatusError;
                    }
                }
            }
            PositionProgressBars();
            UpdateNavigationState();
            UpdateUploadButtonState();

            string text = string.IsNullOrWhiteSpace(message)
                ? Strings.FileLinkWizardUploadFailed
                : string.Format(CultureInfo.CurrentCulture, Strings.FileLinkWizardUploadFailedFormat, message);

            MessageBox.Show(
                text,
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private string HandleDuplicate(FileLinkDuplicateInfo info)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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

                    string sanitized = FileLinkService.SanitizeComponent(input);
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

        private int GetMinPasswordLength()
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (_passwordPolicy != null && _passwordPolicy.MinLength > 0)
            {
                return _passwordPolicy.MinLength;
            }

            return DefaultMinPasswordLength;
        }

        private void GeneratePassword()
        {
            if (!_passwordToggleCheckBox.Checked)
            {
                return;
            }

            int minLength = GetMinPasswordLength();
            _passwordTextBox.Text = GeneratePasswordValue(minLength);
        }

        private string GeneratePasswordValue(int minLength)
        {
            string generated = null;

            try
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_configuration != null && _passwordPolicy != null && _passwordPolicy.HasPolicy)
                {
                    var policyService = new PasswordPolicyService(_configuration);
                    generated = policyService.GeneratePassword(_passwordPolicy);
                }
            }
            catch (Exception ex)
            {
                generated = null;
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Password generation via server policy failed; falling back to local generator.", ex);
            }

            if (string.IsNullOrWhiteSpace(generated) || generated.Trim().Length < minLength)
            {
                generated = PasswordGenerator.GenerateLocalPassword(minLength);
            }

            return generated;
        }

        private bool IsPasswordValid(string password)
        {
            if (string.IsNullOrWhiteSpace(password))
            {
                return false;
            }

            string trimmed = password.Trim();
            int minLength = GetMinPasswordLength();
            return trimmed.Length >= minLength;
        }
    }
}
