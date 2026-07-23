// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Upload progress buffering and per-selection rendering.
    internal sealed partial class FileLinkWizardForm
    {
        private const int UploadProgressFlushIntervalMs = 80;
        private static readonly Color ActiveQueueItemBackground =
            BrandingAssets.BrandBlue;
        private static readonly Color ActiveQueueItemText = Color.White;
        private readonly System.Windows.Forms.Timer _uploadProgressFlushTimer =
            new System.Windows.Forms.Timer();
        private readonly object _uploadProgressSync = new object();
        private readonly Dictionary<
            FileLinkSelection,
            FileLinkUploadItemProgress> _pendingUploadProgress =
                new Dictionary<
                    FileLinkSelection,
                    FileLinkUploadItemProgress>();
        private readonly List<FileLinkSelection> _pendingUploadProgressOrder =
            new List<FileLinkSelection>();
        private bool _uploadProgressPumpRequested;
        private DateTime _overallUploadStartedUtc;

        private void InitializeUploadProgressPump()
        {
            _uploadProgressFlushTimer.Interval =
                UploadProgressFlushIntervalMs;
            _uploadProgressFlushTimer.Tick +=
                (s, e) => FlushBufferedUploadProgress();
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
            if (state == null || state.ProgressBar == null)
            {
                return;
            }
            var bar = state.ProgressBar;
            state.ProgressBar = null;
            if (_fileListView != null
                && !_fileListView.IsDisposed
                && !_fileListView.Disposing)
            {
                _fileListView.Controls.Remove(bar);
            }

            bar.Dispose();
        }

        private void PositionProgressBars()
        {
            if (_selectionStates.Count == 0
                || _fileListView.Columns.Count < 3)
            {
                return;
            }
            if (_fileListView.IsDisposed
                || _fileListView.Disposing
                || !_fileListView.IsHandleCreated)
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
            int statusLeft = _fileListView.Columns[0].Width
                + _fileListView.Columns[1].Width;
            int statusWidth = _fileListView.Columns[2].Width;

            foreach (var state in _selectionStates.Values)
            {
                PositionProgressBar(state, statusLeft, statusWidth);
            }
        }

        private void PositionProgressBar(SelectionUploadState state)
        {
            if (_fileListView == null
                || _fileListView.Columns.Count < 3)
            {
                return;
            }
            int statusLeft = _fileListView.Columns[0].Width
                + _fileListView.Columns[1].Width;
            int statusWidth = _fileListView.Columns[2].Width;
            PositionProgressBar(state, statusLeft, statusWidth);
        }

        private void PositionProgressBar(
            SelectionUploadState state,
            int statusLeft,
            int statusWidth)
        {
            if (state == null || state.ProgressBar == null)
            {
                return;
            }
            if (state.ProgressBar.IsDisposed
                || state.ProgressBar.Disposing
                || state.Item == null)
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
            int topPadding = ScaleLogical(2);
            int bottomPadding = ScaleLogical(2);
            int top = bounds.Top + topPadding;
            int maxHeight = Math.Max(
                2,
                bounds.Height - topPadding - bottomPadding);
            int height = Math.Min(
                Math.Max(ScaleLogical(6), 6),
                maxHeight);
            state.ProgressBar.SetBounds(left, top, width, height);
            state.ProgressBar.Visible =
                state.Status == FileLinkUploadStatus.Uploading;
        }

        private string FormatUploadSpeedKbps(double speedKbps)
        {
            double normalizedSpeed = speedKbps;
            if (double.IsNaN(normalizedSpeed)
                || double.IsInfinity(normalizedSpeed)
                || normalizedSpeed < 0)
            {
                normalizedSpeed = 0;
            }

            long rounded = (long)Math.Round(
                normalizedSpeed,
                MidpointRounding.AwayFromZero);
            string value = rounded.ToString(
                "N0",
                CultureInfo.CurrentCulture);
            string format = Strings.FileLinkWizardStatusSpeedKbpsFormat;
            try
            {
                return string.Format(
                    CultureInfo.CurrentCulture,
                    format,
                    value);
            }
            catch (FormatException)
            {
                return value + " KB/s";
            }
        }

        private bool TryGetListViewItemBounds(
            int index,
            out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            if (_fileListView == null
                || _fileListView.IsDisposed
                || _fileListView.Disposing)
            {
                return false;
            }
            if (index < 0 || index >= _fileListView.Items.Count)
            {
                return false;
            }
            try
            {
                bounds = _fileListView.GetItemRect(
                    index,
                    ItemBoundsPortion.Entire);
                return true;
            }
            catch (ArgumentException ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Skipped progress bar positioning because list item "
                    + "bounds are unavailable (index="
                    + index.ToString(CultureInfo.InvariantCulture)
                    + ").",
                    ex);
                return false;
            }
        }

        private void HandleUploadProgress(
            FileLinkUploadItemProgress progress)
        {
            if (progress == null || progress.Selection == null)
            {
                return;
            }
            bool activatePump = false;
            lock (_uploadProgressSync)
            {
                if (_pendingUploadProgress.ContainsKey(
                    progress.Selection))
                {
                    _pendingUploadProgressOrder.Remove(
                        progress.Selection);
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

        private void HandleUploadPhaseProgress(
            FileLinkUploadPhaseProgress progress)
        {
            if (progress == null || IsDisposed || Disposing)
            {
                return;
            }
            if (InvokeRequired)
            {
                BeginInvoke(
                    new Action<FileLinkUploadPhaseProgress>(
                        HandleUploadPhaseProgress),
                    progress);
                return;
            }
            if (!_progressPanel.Visible)
            {
                SetProgressPanelVisible(true);
            }

            switch (progress.Phase)
            {
                case FileLinkUploadPhase.Scanning:
                    _progressBar.Style = ProgressBarStyle.Marquee;
                    _progressLabel.Text =
                        Strings.FileLinkWizardStatusScanning;
                    break;

                case FileLinkUploadPhase.PreparingFolders:
                    _progressBar.Style = ProgressBarStyle.Blocks;
                    SetOverallProgressValue(
                        CalculateProgressPercent(
                            progress.CompletedFolders,
                            progress.TotalFolders));
                    _progressLabel.Text = string.Format(
                        CultureInfo.CurrentCulture,
                        Strings
                            .FileLinkWizardStatusPreparingFoldersFormat,
                        progress.CompletedFolders,
                        progress.TotalFolders);
                    break;

                case FileLinkUploadPhase.Uploading:
                    if (_overallUploadStartedUtc == DateTime.MinValue)
                    {
                        _overallUploadStartedUtc = DateTime.UtcNow;
                    }
                    _progressBar.Style = ProgressBarStyle.Blocks;
                    SetOverallProgressValue(
                        progress.TotalBytes > 0
                            ? CalculateProgressPercent(
                                progress.UploadedBytes,
                                progress.TotalBytes)
                            : CalculateProgressPercent(
                                progress.CompletedFiles,
                                progress.TotalFiles));
                    double elapsedSeconds = Math.Max(
                        0.001,
                        (DateTime.UtcNow - _overallUploadStartedUtc)
                            .TotalSeconds);
                    long bytesPerSecond = (long)Math.Round(
                        progress.UploadedBytes / elapsedSeconds,
                        MidpointRounding.AwayFromZero);
                    _progressLabel.Text = string.Format(
                        CultureInfo.CurrentCulture,
                        Strings
                            .FileLinkWizardStatusUploadingSummaryFormat,
                        progress.CompletedFiles,
                        progress.TotalFiles,
                        SizeFormatting.FormatBytes(
                            progress.UploadedBytes),
                        SizeFormatting.FormatBytes(
                            progress.TotalBytes),
                        SizeFormatting.FormatBytesPerSecond(
                            bytesPerSecond));
                    break;

                case FileLinkUploadPhase.Completed:
                    _progressBar.Style = ProgressBarStyle.Blocks;
                    SetOverallProgressValue(100);
                    _progressLabel.Text =
                        Strings.FileLinkWizardStatusSuccess;
                    break;
            }
        }

        private void SetOverallProgressValue(int value)
        {
            _progressBar.Value = Math.Max(
                _progressBar.Minimum,
                Math.Min(_progressBar.Maximum, value));
        }

        private static int CalculateProgressPercent(
            long completed,
            long total)
        {
            if (total <= 0)
            {
                return 100;
            }
            return (int)Math.Max(
                0,
                Math.Min(100, (completed * 100L) / total));
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
                BeginInvoke(
                    new Action(EnsureUploadProgressPumpRunning));
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

                snapshot = new List<FileLinkUploadItemProgress>(
                    _pendingUploadProgressOrder.Count);
                foreach (var selection in _pendingUploadProgressOrder)
                {
                    FileLinkUploadItemProgress queuedProgress;
                    if (_pendingUploadProgress.TryGetValue(
                        selection,
                        out queuedProgress))
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

        private void ApplyUploadProgress(
            FileLinkUploadItemProgress progress)
        {
            if (progress == null || progress.Selection == null)
            {
                return;
            }

            SelectionUploadState state;
            if (!_selectionStates.TryGetValue(
                progress.Selection,
                out state))
            {
                return;
            }

            state.TotalBytes = progress.TotalBytes;
            state.UploadedBytes = progress.UploadedBytes;
            state.Status = progress.Status;

            if (progress.Status == FileLinkUploadStatus.Uploading)
            {
                ApplyQueueRowStyle(
                    state,
                    ActiveQueueItemBackground,
                    ActiveQueueItemText);
                int percent = state.TotalBytes > 0
                    ? (int)Math.Min(
                        100,
                        (state.UploadedBytes * 100L)
                            / state.TotalBytes)
                    : 100;
                percent = Math.Max(0, Math.Min(100, percent));
                if (state.ProgressBar == null)
                {
                    state.ProgressBar = CreateProgressBar();
                }
                if (state.UploadStartedUtc == DateTime.MinValue)
                {
                    state.UploadStartedUtc = DateTime.UtcNow;
                }
                double elapsedSeconds = Math.Max(
                    0.001,
                    (DateTime.UtcNow - state.UploadStartedUtc)
                        .TotalSeconds);
                state.UploadSpeedKbps = state.UploadedBytes > 0
                    ? state.UploadedBytes / 1024d / elapsedSeconds
                    : 0;
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = true;
                    state.ProgressBar.Value = percent;
                }
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text =
                        FormatUploadSpeedKbps(
                            state.UploadSpeedKbps);
                    state.Item.SubItems[2].ForeColor =
                        ActiveQueueItemText;
                }

                PositionProgressBar(state);
            }
            else if (progress.Status
                == FileLinkUploadStatus.Completed)
            {
                ApplyQueueRowStyle(
                    state,
                    _themePalette.InputBackground,
                    _themePalette.Text);
                DisposeStateProgressBar(state);
                state.UploadStartedUtc = DateTime.MinValue;
                state.UploadSpeedKbps = 0;
                string statusText =
                    Strings.FileLinkWizardStatusSuccess;
                if (!string.IsNullOrEmpty(state.RenamedTo))
                {
                    statusText += " \u2192 " + state.RenamedTo;
                }
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text = statusText;
                    state.Item.SubItems[2].ForeColor =
                        _themePalette.SuccessText;
                }
            }
            else if (progress.Status
                == FileLinkUploadStatus.Failed)
            {
                ApplyQueueRowStyle(
                    state,
                    _themePalette.InputBackground,
                    _themePalette.Text);
                DisposeStateProgressBar(state);
                state.UploadStartedUtc = DateTime.MinValue;
                state.UploadSpeedKbps = 0;
                if (state.Item.SubItems.Count >= 3)
                {
                    state.Item.SubItems[2].Text =
                        Strings.FileLinkWizardStatusError;
                    state.Item.SubItems[2].ForeColor =
                        _themePalette.ErrorText;
                }
            }
        }
    }
}
