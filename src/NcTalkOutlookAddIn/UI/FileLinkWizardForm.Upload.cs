// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Upload lifecycle, preparation, cancellation, and error handling.
    internal sealed partial class FileLinkWizardForm
    {
        private readonly object _uploadCleanupSync = new object();
        private Task _pendingUploadCleanupTask = Task.CompletedTask;

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_shareFinalized && _cancellationSource != null)
            {
                _closeAfterCancellation = true;
                if (!_cancellationSource.IsCancellationRequested)
                {
                    _cancellationSource.Cancel();
                }
                e.Cancel = true;
                return;
            }

            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            ResetUploadProgressPump();
            _uploadProgressFlushTimer.Dispose();
            _fileListRowHeightImageList.Dispose();
            QueueUnfinalizedUploadContextCleanup(
                "wizard_closed_without_finalize");
            base.OnFormClosed(e);
        }

        private void QueueUnfinalizedUploadContextCleanup(string reason)
        {
            if (_shareFinalized || _uploadContext == null)
            {
                return;
            }

            FileLinkUploadContext context = _uploadContext;
            _uploadContext = null;
            _uploadCompleted = false;

            QueuePreparedUploadContextCleanup(context, reason);
        }

        private Task QueuePreparedUploadContextCleanup(
            FileLinkUploadContext context,
            string reason)
        {
            if (context == null)
            {
                return Task.CompletedTask;
            }

            string relativeFolderPath = context.RelativeFolderPath ?? string.Empty;
            if (string.IsNullOrWhiteSpace(relativeFolderPath))
            {
                return Task.CompletedTask;
            }

            lock (_uploadCleanupSync)
            {
                _pendingUploadCleanupTask =
                    _pendingUploadCleanupTask.ContinueWith(
                        previousTask => CleanupPreparedUploadContext(
                            relativeFolderPath,
                            reason),
                        CancellationToken.None,
                        TaskContinuationOptions.None,
                        TaskScheduler.Default);
                return _pendingUploadCleanupTask;
            }
        }

        private async Task AwaitPendingUploadCleanupAsync()
        {
            while (true)
            {
                Task pendingTask;
                lock (_uploadCleanupSync)
                {
                    pendingTask = _pendingUploadCleanupTask;
                }

                await pendingTask;

                lock (_uploadCleanupSync)
                {
                    if (ReferenceEquals(
                        pendingTask,
                        _pendingUploadCleanupTask))
                    {
                        return;
                    }
                }
            }
        }

        private void CleanupPreparedUploadContext(
            string relativeFolderPath,
            string reason)
        {
            using (var cleanupSource = new CancellationTokenSource())
            {
                cleanupSource.CancelAfter(
                    FileLinkUploadPolicy.CleanupTimeoutMs);
                try
                {
                    _service.DeleteShareFolder(
                        relativeFolderPath,
                        cleanupSource.Token);
                    DiagnosticsLogger.Log(
                        LogCategories.FileLink,
                        "Wizard upload cleanup completed (reason="
                        + (reason ?? string.Empty)
                        + ", relativeFolder="
                        + relativeFolderPath
                        + ").");
                }
                catch (OperationCanceledException)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.FileLink,
                        "Wizard upload cleanup timed out (reason="
                        + (reason ?? string.Empty)
                        + ", relativeFolder="
                        + relativeFolderPath
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Wizard upload cleanup failed (reason="
                        + (reason ?? string.Empty)
                        + ", relativeFolder="
                        + relativeFolderPath
                        + ").",
                        ex);
                }
            }
        }

        private void InvalidateUpload()
        {
            FileLinkUploadContext previousContext = _uploadContext;
            _uploadCompleted = false;
            _uploadContext = null;
            ResetUploadProgressPump();
            if (_progressPanel.Visible)
            {
                SetProgressPanelVisible(false);
            }
            if (_items.Count > 0)
            {
                _allowEmptyUpload = false;
            }
            if (!_shareFinalized && previousContext != null)
            {
                QueuePreparedUploadContextCleanup(
                    previousContext,
                    "upload_invalidated");
            }
            foreach (var state in _selectionStates.Values)
            {
                state.TotalBytes = 0;
                state.UploadedBytes = 0;
                state.Status = FileLinkUploadStatus.Pending;
                state.RenamedTo = null;
                state.UploadStartedUtc = DateTime.MinValue;
                state.UploadSpeedKbps = 0;
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

        private async Task StartUploadAsync()
        {
            if (_uploadInProgress || _items.Count == 0)
            {
                return;
            }

            ApplyFormData();
            _uploadCompleted = false;
            _overallUploadStartedUtc = DateTime.MinValue;
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressLabel.Text = Strings.FileLinkWizardStatusScanning;
            SetProgressPanelVisible(true);
            ToggleUpload(true);

            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            UseWaitCursor = true;
            FileLinkUploadContext preparedContext = null;
            Task cleanupTask = Task.CompletedTask;

            try
            {
                ResetUploadProgressPump();
                foreach (var state in _selectionStates.Values)
                {
                    state.TotalBytes = 0;
                    state.UploadedBytes = 0;
                    state.Status = FileLinkUploadStatus.Pending;
                    state.UploadStartedUtc = DateTime.MinValue;
                    state.UploadSpeedKbps = 0;
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
                await AwaitPendingUploadCleanupAsync();
                token.ThrowIfCancellationRequested();

                var itemProgress = new Progress<FileLinkUploadItemProgress>(
                    HandleUploadProgress);
                var phaseProgress = new Progress<FileLinkUploadPhaseProgress>(
                    HandleUploadPhaseProgress);

                await Task.Run(() =>
                {
                    preparedContext = _service.PrepareUpload(
                        _request,
                        _items,
                        HandleDuplicate,
                        phaseProgress,
                        token);
                    _service.UploadSelections(
                        preparedContext,
                        itemProgress,
                        phaseProgress,
                        token);
                });

                _uploadContext = preparedContext;
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
                    state.UploadStartedUtc = DateTime.MinValue;
                    state.UploadSpeedKbps = 0;
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
                if (!_uploadCompleted && preparedContext != null)
                {
                    cleanupTask = QueuePreparedUploadContextCleanup(
                        preparedContext,
                        "upload_not_completed");
                }
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }

                UseWaitCursor = false;
                Cursor.Current = previousCursor;
                ToggleUpload(false);
                PositionProgressBars();
                CloseAfterCancellation();
            }

            await cleanupTask;
        }

        private async Task<bool> PrepareEmptyUploadAsync()
        {
            FileLinkUploadContext preparedContext = null;
            Task cleanupTask = Task.CompletedTask;
            bool uploadPrepared = false;
            _uploadCompleted = false;
            _overallUploadStartedUtc = DateTime.MinValue;
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressLabel.Text = Strings.FileLinkWizardStatusScanning;
            SetProgressPanelVisible(true);
            ToggleUpload(true);

            Cursor previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            UseWaitCursor = true;
            try
            {
                _cancellationSource = new CancellationTokenSource();
                CancellationToken token = _cancellationSource.Token;
                await AwaitPendingUploadCleanupAsync();
                token.ThrowIfCancellationRequested();
                var phaseProgress =
                    new Progress<FileLinkUploadPhaseProgress>(
                        HandleUploadPhaseProgress);

                preparedContext = await Task.Run(
                    () => _service.PrepareUpload(
                        _request,
                        _items,
                        HandleDuplicate,
                        phaseProgress,
                        token));

                _uploadContext = preparedContext;
                _uploadCompleted = true;
                _progressBar.Style = ProgressBarStyle.Blocks;
                _progressBar.Value = 100;
                _progressLabel.Text = Strings.FileLinkWizardStatusSuccess;
                uploadPrepared = true;
            }
            catch (OperationCanceledException)
            {
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Empty upload preparation cancelled.");
                ShowUploadError(
                    Strings.FileLinkWizardUploadCancelledMessage);
            }
            catch (TalkServiceException ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Empty upload preparation failed with service error.",
                    ex);
                ShowUploadError(ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Empty upload preparation failed unexpectedly.",
                    ex);
                ShowUploadError(ex.Message);
            }
            finally
            {
                if (!_uploadCompleted && preparedContext != null)
                {
                    cleanupTask = QueuePreparedUploadContextCleanup(
                        preparedContext,
                        "empty_upload_prepare_not_completed");
                }
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }

                UseWaitCursor = false;
                Cursor.Current = previousCursor;
                ToggleUpload(false);
                CloseAfterCancellation();
            }

            await cleanupTask;
            return uploadPrepared;
        }

        private void CloseAfterCancellation()
        {
            if (!_closeAfterCancellation || IsDisposed || Disposing)
            {
                return;
            }

            _closeAfterCancellation = false;
            BeginInvoke(new Action(Close));
        }

        private void ToggleUpload(bool uploading)
        {
            _uploadInProgress = uploading;
            UpdateNavigationState();
            UpdateUploadButtonState();
            _cancelButton.Enabled = true;
        }

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

        private void ShowUploadError(string message)
        {
            _uploadCompleted = false;
            _uploadContext = null;
            ResetUploadProgressPump();
            SetProgressPanelVisible(false);
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

    }
}
