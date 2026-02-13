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
using System.Security.Cryptography;
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
    internal sealed class FileLinkWizardForm : Form
    {
        private const int DefaultMinPasswordLength = 8;
        private readonly UiThemePalette _themePalette = UiThemeManager.DetectPalette();

        private readonly FileLinkService _service;
        private readonly FileLinkRequest _request = new FileLinkRequest();
        private readonly TalkServiceConfiguration _configuration;
        private readonly PasswordPolicyInfo _passwordPolicy;
        private readonly AddinSettings _defaults;
        private readonly List<Panel> _steps = new List<Panel>();
        private readonly Label _titleLabel = new Label();
        private readonly Button _backButton = new Button();
        private readonly Button _nextButton = new Button();
        private readonly Button _finishButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly Button _uploadButton = new Button();
        private readonly BrandedHeader _headerPanel = new BrandedHeader();
        private Panel _stepHost;
        private readonly Panel _progressPanel = new Panel();
        private readonly ProgressBar _progressBar = new ProgressBar();
        private readonly Label _progressLabel = new Label();
        private readonly ListView _fileListView = new ListView();
        private readonly Label _basePathLabel = new Label();
        private readonly TextBox _shareNameTextBox = new TextBox();
        private readonly CheckBox _permissionReadCheckBox = new CheckBox();
        private readonly CheckBox _permissionCreateCheckBox = new CheckBox();
        private readonly CheckBox _permissionWriteCheckBox = new CheckBox();
        private readonly CheckBox _permissionDeleteCheckBox = new CheckBox();
        private readonly CheckBox _passwordToggleCheckBox = new CheckBox();
        private readonly TextBox _passwordTextBox = new TextBox();
        private readonly Button _passwordGenerateButton = new Button();
        private readonly CheckBox _expireToggleCheckBox = new CheckBox();
        private readonly DateTimePicker _expireDatePicker = new DateTimePicker();
        private readonly Label _expireHintLabel = new Label();
        private readonly Button _addFilesButton = new Button();
        private readonly Button _addFolderButton = new Button();
        private readonly Button _removeItemButton = new Button();
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
        private FileLinkRequest _requestSnapshot;

        internal FileLinkWizardForm(AddinSettings defaults, TalkServiceConfiguration configuration, PasswordPolicyInfo passwordPolicy, string basePath)
        {
            _defaults = defaults ?? new AddinSettings();
            _configuration = configuration;
            _passwordPolicy = passwordPolicy;
            _service = new FileLinkService(configuration);
            _request.BasePath = basePath ?? string.Empty;

            Text = Strings.FileLinkWizardTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(640, 480);
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeHeader();
            InitializeWizardLayout();
            InitializeStepGeneral();
            InitializeStepExpiration();
            InitializeStepFiles();
            InitializeStepNote();
            InitializeProgressPanel();

            UiThemeManager.ApplyToForm(this);

            ShowStep(0);
        }

        internal FileLinkResult Result { get; private set; }

        internal FileLinkRequest RequestSnapshot
        {
            get { return _requestSnapshot; }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            LayoutBottomButtons();
            if (_stepHost != null && _currentStepIndex == 2)
            {
                LayoutFileStep(_stepHost.ClientSize);
            }
            if (_stepHost != null)
            {
                _stepHost.Size = new Size(ClientSize.Width - 40, ClientSize.Height - _stepHost.Top - 120);
            }
            PositionProgressBars();
        }

        private void InitializeHeader()
        {
            _headerPanel.Height = 48;
            _headerPanel.Dock = DockStyle.Top;
            _headerPanel.Padding = new Padding(0);

            Controls.Add(_headerPanel);
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
                Size = new Size(ClientSize.Width - 40, ClientSize.Height - (_titleLabel.Bottom + 16) - 120),
                BorderStyle = BorderStyle.None
            };
            Controls.Add(_stepHost);

            _backButton.Text = Strings.ButtonBack;
            _backButton.AutoSize = true;
            _backButton.Click += (s, e) => Navigate(-1);
            Controls.Add(_backButton);

            _uploadButton.Text = Strings.FileLinkWizardUploadButton;
            _uploadButton.AutoSize = true;
            _uploadButton.Enabled = false;
            _uploadButton.Visible = false;
            _uploadButton.Click += async (s, e) => await StartUploadAsync();
            Controls.Add(_uploadButton);

            _nextButton.Text = Strings.ButtonNext;
            _nextButton.AutoSize = true;
            _nextButton.Click += (s, e) => Navigate(1);
            Controls.Add(_nextButton);

            _finishButton.Text = Strings.FileLinkWizardFinishButton;
            _finishButton.AutoSize = true;
            _finishButton.Click += async (s, e) => await FinishAsync();
            Controls.Add(_finishButton);

            _cancelButton.Text = Strings.ButtonCancel;
            _cancelButton.AutoSize = true;
            _cancelButton.Click += (s, e) => Close();
            Controls.Add(_cancelButton);

            LayoutBottomButtons();
        }

        private void InitializeStepGeneral()
        {
            var panel = CreateStepPanel();

            var nameLabel = new Label
            {
                Text = Strings.FileLinkWizardShareNameLabel,
                Location = new Point(12, 12),
                AutoSize = true
            };
            panel.Controls.Add(nameLabel);

            _shareNameTextBox.Location = new Point(12, 36);
            _shareNameTextBox.Width = 360;
            _shareNameTextBox.Text = string.IsNullOrWhiteSpace(_defaults.SharingDefaultShareName)
                ? Strings.FileLinkWizardFallbackShareName
                : _defaults.SharingDefaultShareName;
            _shareNameTextBox.TextChanged += (s, e) => InvalidateUpload();
            panel.Controls.Add(_shareNameTextBox);

            var permissionsLabel = new Label
            {
                Text = Strings.FileLinkWizardPermissionsLabel,
                Location = new Point(12, 76),
                AutoSize = true
            };
            panel.Controls.Add(permissionsLabel);

            _permissionReadCheckBox.Text = Strings.FileLinkPermissionRead;
            _permissionReadCheckBox.Location = new Point(18, 100);
            _permissionReadCheckBox.Checked = true;
            _permissionReadCheckBox.Enabled = false;
            panel.Controls.Add(_permissionReadCheckBox);

            _permissionCreateCheckBox.Text = Strings.FileLinkPermissionCreate;
            _permissionCreateCheckBox.Location = new Point(18, 128);
            _permissionCreateCheckBox.Checked = _defaults.SharingDefaultPermCreate;
            panel.Controls.Add(_permissionCreateCheckBox);

            _permissionWriteCheckBox.Text = Strings.FileLinkPermissionWrite;
            _permissionWriteCheckBox.Location = new Point(18, 156);
            _permissionWriteCheckBox.Checked = _defaults.SharingDefaultPermWrite;
            panel.Controls.Add(_permissionWriteCheckBox);

            _permissionDeleteCheckBox.Text = Strings.FileLinkPermissionDelete;
            _permissionDeleteCheckBox.Location = new Point(18, 184);
            _permissionDeleteCheckBox.Checked = _defaults.SharingDefaultPermDelete;
            panel.Controls.Add(_permissionDeleteCheckBox);

            _passwordToggleCheckBox.Text = Strings.FileLinkWizardPasswordToggle;
            _passwordToggleCheckBox.AutoSize = true;
            _passwordToggleCheckBox.Location = new Point(12, 228);
            _passwordToggleCheckBox.Checked = _defaults.SharingDefaultPasswordEnabled;
            _passwordToggleCheckBox.CheckedChanged += (s, e) => UpdatePasswordState();
            panel.Controls.Add(_passwordToggleCheckBox);

            int toggleWidth = _passwordToggleCheckBox.PreferredSize.Width;
            _passwordTextBox.Width = 220;
            _passwordTextBox.Location = new Point(_passwordToggleCheckBox.Left + toggleWidth + 12, _passwordToggleCheckBox.Top - 2);
            panel.Controls.Add(_passwordTextBox);

            _passwordGenerateButton.Text = Strings.TalkPasswordGenerate;
            _passwordGenerateButton.AutoSize = true;
            _passwordGenerateButton.Location = new Point(_passwordTextBox.Right + 8, _passwordToggleCheckBox.Top - 4);
            _passwordGenerateButton.Click += (s, e) => GeneratePassword();
            panel.Controls.Add(_passwordGenerateButton);

            if (_passwordToggleCheckBox.Checked)
            {
                _passwordTextBox.Text = GenerateLocalPassword(GetMinPasswordLength());
            }

            UpdatePasswordState();

            _steps.Add(panel);
        }

        private void InitializeStepExpiration()
        {
            var panel = CreateStepPanel();

            _expireToggleCheckBox.Text = Strings.FileLinkWizardExpireToggle;
            _expireToggleCheckBox.AutoSize = true;
            _expireToggleCheckBox.Location = new Point(12, 12);
            _expireToggleCheckBox.Checked = _defaults.SharingDefaultExpireDays > 0;
            _expireToggleCheckBox.CheckedChanged += (s, e) => UpdateExpireState();
            panel.Controls.Add(_expireToggleCheckBox);

            _expireDatePicker.Width = 160;
            _expireDatePicker.Format = DateTimePickerFormat.Short;
            int expireDays = _defaults.SharingDefaultExpireDays > 0 ? _defaults.SharingDefaultExpireDays : 7;
            _expireDatePicker.Value = DateTime.Today.AddDays(expireDays);
            _expireDatePicker.Location = new Point(_expireToggleCheckBox.Left + _expireToggleCheckBox.PreferredSize.Width + 12, _expireToggleCheckBox.Top - 2);
            panel.Controls.Add(_expireDatePicker);

            _expireHintLabel.Text = Strings.FileLinkWizardExpireHint;
            _expireHintLabel.Location = new Point(12, _expireToggleCheckBox.Bottom + 24);
            _expireHintLabel.AutoSize = true;
            _expireHintLabel.ForeColor = Color.DimGray;
            panel.Controls.Add(_expireHintLabel);

            UpdateExpireState();
            _steps.Add(panel);
        }

        private void InitializeStepFiles()
        {
            var panel = CreateStepPanel();
            panel.SuspendLayout();

            _basePathLabel.Text = Strings.FileLinkWizardBasePathPrefix + (_request.BasePath ?? string.Empty);
            _basePathLabel.Location = new Point(12, 12);
            _basePathLabel.AutoSize = true;
            panel.Controls.Add(_basePathLabel);

            _fileListView.Location = new Point(12, _basePathLabel.Bottom + 12);
            _fileListView.Size = new Size(360, 240);
            _fileListView.View = View.Details;
            _fileListView.FullRowSelect = true;
            _fileListView.HideSelection = false;
            _fileListView.Scrollable = true;
            _fileListView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnPath, 240);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnType, 100);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnStatus, 120);
            _fileListView.Resize += (s, e) => PositionProgressBars();
            panel.Controls.Add(_fileListView);

            _addFilesButton.Text = Strings.FileLinkWizardAddFilesButton;
            _addFilesButton.Size = new Size(150, 28);
            _addFilesButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _addFilesButton.Click += (s, e) => AddFiles();
            panel.Controls.Add(_addFilesButton);

            _addFolderButton.Text = Strings.FileLinkWizardAddFolderButton;
            _addFolderButton.Size = new Size(150, 28);
            _addFolderButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _addFolderButton.Click += (s, e) => AddFolder();
            panel.Controls.Add(_addFolderButton);

            _removeItemButton.Text = Strings.FileLinkWizardRemoveButton;
            _removeItemButton.Size = new Size(150, 28);
            _removeItemButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _removeItemButton.Click += (s, e) => RemoveSelection();
            panel.Controls.Add(_removeItemButton);

            panel.ResumeLayout(false);
            panel.PerformLayout();

            panel.ClientSizeChanged += (s, e) => LayoutFileStep(panel.ClientSize);
            LayoutFileStep(panel.ClientSize);
            PositionProgressBars();

            _steps.Add(panel);
        }

        private void LayoutFileStep(Size clientSize)
        {
            int width = Math.Max(360, clientSize.Width);
            int height = Math.Max(260, clientSize.Height);

            int listTop = _basePathLabel.Bottom + 12;
            int buttonWidth = _addFilesButton.Width;
            int buttonLeft = Math.Max(_fileListView.Left + 260, width - 12 - buttonWidth);

            int availableWidth = buttonLeft - _fileListView.Left - 12;
            int listWidth = Math.Max(260, availableWidth);
            int listHeight = Math.Max(160, height - listTop - 12);

            _fileListView.Size = new Size(listWidth, listHeight);

            _addFilesButton.Location = new Point(buttonLeft, listTop);
            _addFolderButton.Location = new Point(buttonLeft, _addFilesButton.Bottom + 8);
            _removeItemButton.Location = new Point(buttonLeft, _addFolderButton.Bottom + 8);

            PositionProgressBars();
            LayoutBottomButtons();
        }

        private void InitializeStepNote()
        {
            var panel = CreateStepPanel();

            _noteToggleCheckBox.Text = Strings.FileLinkWizardNoteToggle;
            _noteToggleCheckBox.AutoSize = true;
            _noteToggleCheckBox.Location = new Point(12, 12);
            _noteToggleCheckBox.CheckedChanged += (s, e) => UpdateNoteState();
            panel.Controls.Add(_noteToggleCheckBox);

            _noteTextBox.Location = new Point(12, 44);
            _noteTextBox.Size = new Size(560, 220);
            _noteTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _noteTextBox.Multiline = true;
            _noteTextBox.ScrollBars = ScrollBars.Vertical;
            _noteTextBox.Enabled = false;
            panel.Controls.Add(_noteTextBox);

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
        }

        private Panel CreateStepPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Visible = false
            };

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
            PositionProgressBars();
        }

        private void Navigate(int direction)
        {
            if (direction > 0 && !ValidateCurrentStep())
            {
                return;
            }

            int newIndex = _currentStepIndex + direction;
            ShowStep(newIndex);
        }

        private bool ValidateCurrentStep()
        {
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

        private async Task FinishAsync()
        {
            if (!ValidateCurrentStep())
            {
                return;
            }

            ApplyFormData();

            if (_allowEmptyUpload && (_uploadContext == null || !_uploadCompleted))
            {
                _uploadContext = _service.PrepareUpload(_request, CancellationToken.None);
                _uploadCompleted = true;
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
            if (_permissionCreateCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Create;
            }
            if (_permissionWriteCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Write;
            }
            if (_permissionDeleteCheckBox.Checked)
            {
                _request.Permissions |= FileLinkPermissionFlags.Delete;
            }

            _request.PasswordEnabled = _passwordToggleCheckBox.Checked;
            _request.Password = _passwordToggleCheckBox.Checked ? _passwordTextBox.Text : null;
            _request.ExpireEnabled = _expireToggleCheckBox.Checked;
            _request.ExpireDate = _expireToggleCheckBox.Checked ? _expireDatePicker.Value.Date : (DateTime?)null;
            _request.NoteEnabled = _noteToggleCheckBox.Checked;
            _request.Note = _noteToggleCheckBox.Checked ? _noteTextBox.Text.Trim() : null;

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
                ExpireEnabled = source.ExpireEnabled,
                ExpireDate = source.ExpireDate,
                NoteEnabled = source.NoteEnabled,
                Note = source.Note
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

            string folderName = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_" + sanitizedShareName;

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
            _passwordTextBox.Enabled = enabled;
            _passwordGenerateButton.Enabled = enabled;
        }

        private void UpdateExpireState()
        {
            bool enabled = _expireToggleCheckBox.Checked;
            _expireDatePicker.Enabled = enabled;
            _expireHintLabel.Enabled = enabled;
        }

        private void UpdateNoteState()
        {
            _noteTextBox.Enabled = _noteToggleCheckBox.Checked;
        }

        private void AddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    foreach (string file in dialog.FileNames)
                    {
                        AddSelection(new FileLinkSelection(FileLinkSelectionType.File, file));
                    }
                }
            }
        }

        private void AddFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    AddSelection(new FileLinkSelection(FileLinkSelectionType.Directory, dialog.SelectedPath));
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
                FileLinkSelection selection = item.Tag as FileLinkSelection;
                if (selection != null)
                {
                    _items.Remove(selection);
                    SelectionUploadState state;
                    if (_selectionStates.TryGetValue(selection, out state))
                    {
                        if (state.ProgressBar != null)
                        {
                            _fileListView.Controls.Remove(state.ProgressBar);
                            state.ProgressBar.Dispose();
                        }
                        _selectionStates.Remove(selection);
                    }
                }
                _fileListView.Items.Remove(item);
            }

            if (_items.Count == 0)
            {
                _allowEmptyUpload = false;
            }
            PositionProgressBars();
            InvalidateUpload();
        }

        private void AddSelection(FileLinkSelection selection)
        {
            if (selection == null)
            {
                return;
            }

            if (_items.Any(i => string.Equals(i.LocalPath, selection.LocalPath, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            _items.Add(selection);
            _allowEmptyUpload = false;

            var listViewItem = new ListViewItem(selection.LocalPath)
            {
                Tag = selection
            };
            listViewItem.UseItemStyleForSubItems = false;
            listViewItem.SubItems.Add(selection.SelectionType == FileLinkSelectionType.File ? Strings.FileLinkWizardTypeFile : Strings.FileLinkWizardTypeFolder);
            listViewItem.SubItems.Add(string.Empty);
            _fileListView.Items.Add(listViewItem);

            var state = new SelectionUploadState(listViewItem)
            {
                ProgressBar = CreateProgressBar()
            };
            _selectionStates[selection] = state;

            PositionProgressBars();
            InvalidateUpload();
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

        private void InvalidateUpload()
        {
            _uploadCompleted = false;
            _uploadContext = null;
            if (_items.Count > 0)
            {
                _allowEmptyUpload = false;
            }

            foreach (var state in _selectionStates.Values)
            {
                state.TotalBytes = 0;
                state.UploadedBytes = 0;
                state.Status = FileLinkUploadStatus.Pending;
                state.RenamedTo = null;
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = false;
                    state.ProgressBar.Value = 0;
                }
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
        }

        private void UpdateUploadButtonState()
        {
            _uploadButton.Enabled = _uploadButton.Visible && !_uploadInProgress && _items.Count > 0 && !_uploadCompleted;
        }

        private void LayoutBottomButtons()
        {
            int y = 420;
            int spacing = 12;
            var buttons = new List<Button>();

            buttons.Add(_backButton);
            if (_uploadButton.Visible)
            {
                buttons.Add(_uploadButton);
            }
            buttons.Add(_nextButton);
            buttons.Add(_finishButton);
            buttons.Add(_cancelButton);

            int totalWidth = 0;
            foreach (var button in buttons)
            {
                totalWidth += button.Width;
            }
            if (buttons.Count > 1)
            {
                totalWidth += spacing * (buttons.Count - 1);
            }

            int left = Math.Max(12, (ClientSize.Width - totalWidth) / 2);

            foreach (var button in buttons)
            {
                button.Location = new Point(left, y);
                left += button.Width + spacing;
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

            if (_fileListView.Items.Count == 0)
            {
                foreach (var state in _selectionStates.Values)
                {
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
                if (state.ProgressBar == null)
                {
                    continue;
                }

                int index = state.Item.Index;
                if (index < 0 || index >= _fileListView.Items.Count)
                {
                    state.ProgressBar.Visible = false;
                    continue;
                }

                Rectangle bounds = _fileListView.GetItemRect(index, ItemBoundsPortion.Entire);
                int left = statusLeft + 4;
                int width = Math.Max(12, statusWidth - 8);
                int top = bounds.Top + 3;
                int height = Math.Max(6, bounds.Height - 6);
                state.ProgressBar.SetBounds(left, top, width, height);
                state.ProgressBar.Visible = state.Status == FileLinkUploadStatus.Uploading;
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

            try
            {
                foreach (var state in _selectionStates.Values)
                {
                    state.TotalBytes = 0;
                    state.UploadedBytes = 0;
                    state.Status = FileLinkUploadStatus.Uploading;
                    if (state.ProgressBar != null)
                    {
                        state.ProgressBar.Value = 0;
                        state.ProgressBar.Visible = true;
                    }
                    if (state.Item.SubItems.Count >= 3)
                    {
                        state.Item.SubItems[2].Text = "0%";
                    }
                }
                PositionProgressBars();

                _cancellationSource = new CancellationTokenSource();
                CancellationToken token = _cancellationSource.Token;

                var context = _service.PrepareUpload(_request, token);
                var progress = new Progress<FileLinkUploadItemProgress>(HandleUploadProgress);

                await Task.Run(() =>
                {
                    _service.UploadSelections(context, _items, progress, HandleDuplicate, token);
                });

                _uploadContext = context;
                _uploadCompleted = true;
                UpdateNavigationState();
                UpdateUploadButtonState();
            }
            catch (OperationCanceledException)
            {
                DiagnosticsLogger.Log(LogCategories.FileLink, "Upload cancelled.");
                foreach (var state in _selectionStates.Values)
                {
                    state.Status = FileLinkUploadStatus.Failed;
                    if (state.ProgressBar != null)
                    {
                        state.ProgressBar.Visible = false;
                        state.ProgressBar.Value = 0;
                    }
                    if (state.Item.SubItems.Count >= 3)
                    {
                        state.Item.SubItems[2].Text = Strings.FileLinkWizardStatusCancelled;
                        state.Item.SubItems[2].ForeColor = _themePalette.ErrorText;
                    }
                }
                ShowUploadError(Strings.FileLinkWizardUploadCancelledMessage);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Upload failed with service error.", ex);
                ShowUploadError(ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Upload failed unexpectedly.", ex);
                ShowUploadError(ex.Message);
            }
            finally
            {
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
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
                dialog.ClientSize = new Size(420, 150);
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.Icon = BrandingAssets.GetAppIcon(32);

                var label = new Label
                {
                    Text = Strings.FileLinkNoFilesConfirm,
                    AutoSize = false,
                    Size = new Size(390, 70),
                    Location = new Point(15, 15)
                };
                dialog.Controls.Add(label);

                var continueButton = new Button
                {
                    Text = Strings.ButtonNext,
                    DialogResult = DialogResult.OK,
                    Width = 90,
                    Location = new Point(220, 100)
                };
                var backButton = new Button
                {
                    Text = Strings.ButtonBack,
                    DialogResult = DialogResult.Cancel,
                    Width = 90,
                    Location = new Point(315, 100)
                };

                dialog.Controls.Add(continueButton);
                dialog.Controls.Add(backButton);
                dialog.AcceptButton = continueButton;
                dialog.CancelButton = backButton;

                UiThemeManager.ApplyToForm(dialog);

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
            if (progress == null || progress.Selection == null)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<FileLinkUploadItemProgress>(HandleUploadProgress), progress);
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
                int percent = state.TotalBytes > 0 ? (int)Math.Min(100, (state.UploadedBytes * 100L) / state.TotalBytes) : 100;
                percent = Math.Max(0, Math.Min(100, percent));
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = true;
                    state.ProgressBar.Value = percent;
                }
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = percent.ToString(CultureInfo.InvariantCulture) + "%";
                    state.Item.SubItems[2].ForeColor = _themePalette.MutedText;
                }
            }
            else if (progress.Status == FileLinkUploadStatus.Completed)
            {
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = false;
                    state.ProgressBar.Value = 100;
                }
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
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = false;
                }
                string message = string.IsNullOrWhiteSpace(progress.Message) ? string.Empty : " (" + progress.Message + ")";
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = Strings.FileLinkWizardStatusError + message;
                    state.Item.SubItems[2].ForeColor = _themePalette.ErrorText;
                }
            }

            PositionProgressBars();
        }

        private void ShowUploadError(string message)
        {
            _uploadCompleted = false;
            _uploadContext = null;
            foreach (var state in _selectionStates.Values)
            {
                if (state.Status == FileLinkUploadStatus.Uploading)
                {
                    state.Status = FileLinkUploadStatus.Failed;
                    if (state.ProgressBar != null)
                    {
                        state.ProgressBar.Visible = false;
                        state.ProgressBar.Value = 0;
                    }
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
                form.ClientSize = new Size(360, 140);
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.StartPosition = FormStartPosition.CenterParent;
                form.ShowInTaskbar = false;
                form.Icon = BrandingAssets.GetAppIcon(32);

                var label = new Label
                {
                    AutoSize = true,
                    Text = info.IsDirectory
                        ? Strings.FileLinkWizardRenameFolderPrompt
                        : Strings.FileLinkWizardRenameFilePrompt,
                    Location = new Point(12, 12)
                };
                form.Controls.Add(label);

                var textBox = new TextBox
                {
                    Location = new Point(12, 40),
                    Width = 330,
                    Text = info.OriginalName
                };
                form.Controls.Add(textBox);

                var okButton = new Button { Text = Strings.DialogOk, DialogResult = DialogResult.OK, Location = new Point(188, 88) };
                var cancelButton = new Button { Text = Strings.DialogCancel, DialogResult = DialogResult.Cancel, Location = new Point(270, 88) };
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                UiThemeManager.ApplyToForm(form);

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
            string generated = null;

            try
            {
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
                generated = GenerateLocalPassword(minLength);
            }

            _passwordTextBox.Text = generated;
        }

        private static string GenerateLocalPassword(int minLength)
        {
            const string alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
            int length = Math.Max(8, minLength);
            var chars = new char[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                byte[] data = new byte[4];
                for (int i = 0; i < chars.Length; i++)
                {
                    rng.GetBytes(data);
                    int index = (int)(BitConverter.ToUInt32(data, 0) % alphabet.Length);
                    chars[i] = alphabet[index];
                }
            }

            return new string(chars);
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
