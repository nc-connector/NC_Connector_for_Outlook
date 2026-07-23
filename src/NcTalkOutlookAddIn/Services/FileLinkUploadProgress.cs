// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal static class FileLinkUploadProgress
    {
        internal const int ReportIntervalMs = 100;

        internal static void ReportPhase(
            IProgress<FileLinkUploadPhaseProgress> progress,
            FileLinkUploadPhase phase,
            int completedFolders,
            int totalFolders,
            int completedFiles,
            int totalFiles,
            long uploadedBytes,
            long totalBytes)
        {
            if (progress == null)
            {
                return;
            }

            progress.Report(new FileLinkUploadPhaseProgress(
                phase,
                completedFolders,
                totalFolders,
                completedFiles,
                totalFiles,
                uploadedBytes,
                totalBytes));
        }
    }

    internal sealed class FileLinkFolderProgressReporter
    {
        private readonly object _sync = new object();
        private readonly IProgress<FileLinkUploadPhaseProgress> _progress;
        private readonly int _totalFolders;
        private readonly int _totalFiles;
        private readonly long _totalBytes;
        private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
        private long _lastReportMs = -FileLinkUploadProgress.ReportIntervalMs;

        internal FileLinkFolderProgressReporter(
            IProgress<FileLinkUploadPhaseProgress> progress,
            int totalFolders,
            int totalFiles,
            long totalBytes)
        {
            _progress = progress;
            _totalFolders = totalFolders;
            _totalFiles = totalFiles;
            _totalBytes = totalBytes;
        }

        internal void Report(int completedFolders, bool force)
        {
            lock (_sync)
            {
                long now = _stopwatch.ElapsedMilliseconds;
                if (!force
                    && now - _lastReportMs
                    < FileLinkUploadProgress.ReportIntervalMs)
                {
                    return;
                }

                _lastReportMs = now;
                FileLinkUploadProgress.ReportPhase(
                    _progress,
                    FileLinkUploadPhase.PreparingFolders,
                    completedFolders,
                    _totalFolders,
                    0,
                    _totalFiles,
                    0,
                    _totalBytes);
            }
        }
    }

    internal sealed class FileLinkUploadProgressCoordinator
    {
        private const int ProgressLogIntervalMs = 5000;

        private readonly object _sync = new object();
        private readonly FileLinkUploadPlan _plan;
        private readonly IProgress<FileLinkUploadItemProgress> _itemProgress;
        private readonly IProgress<FileLinkUploadPhaseProgress> _phaseProgress;
        private readonly Dictionary<FileLinkPlannedFile, long> _fileBytes =
            new Dictionary<FileLinkPlannedFile, long>();
        private readonly Dictionary<FileLinkSelection, long> _selectionBytes =
            new Dictionary<FileLinkSelection, long>();
        private readonly Dictionary<FileLinkSelection, int>
            _selectionCompletedFiles =
                new Dictionary<FileLinkSelection, int>();
        private readonly Dictionary<FileLinkSelection, long>
            _selectionLastReportMs =
                new Dictionary<FileLinkSelection, long>();
        private readonly HashSet<FileLinkPlannedFile> _completedFiles =
            new HashSet<FileLinkPlannedFile>();
        private readonly HashSet<FileLinkSelection> _failedSelections =
            new HashSet<FileLinkSelection>();
        private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
        private long _uploadedBytes;
        private int _completedFileCount;
        private long _lastPhaseReportMs =
            -FileLinkUploadProgress.ReportIntervalMs;
        private long _lastLogMs;

        internal FileLinkUploadProgressCoordinator(
            FileLinkUploadPlan plan,
            IProgress<FileLinkUploadItemProgress> itemProgress,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress)
        {
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            _plan = plan;
            _itemProgress = itemProgress;
            _phaseProgress = phaseProgress;

            foreach (FileLinkPlannedFile file in plan.Files)
            {
                _fileBytes[file] = 0;
            }
            foreach (FileLinkSelection selection in plan.SelectionBytes.Keys)
            {
                _selectionBytes[selection] = 0;
                _selectionCompletedFiles[selection] = 0;
                _selectionLastReportMs[selection] =
                    -FileLinkUploadProgress.ReportIntervalMs;
            }
        }

        internal void Begin()
        {
            lock (_sync)
            {
                foreach (FileLinkSelection selection
                    in _plan.SelectionBytes.Keys)
                {
                    ReportItem(
                        selection,
                        FileLinkUploadStatus.Uploading);
                    if (_plan.SelectionFileCounts[selection] == 0)
                    {
                        ReportItem(
                            selection,
                            FileLinkUploadStatus.Completed);
                    }
                }
                ReportOverall(true);
            }
        }

        internal void SetFileBytes(
            FileLinkPlannedFile file,
            long bytes,
            bool force)
        {
            lock (_sync)
            {
                bool changed = ApplyFileBytes(file, bytes);
                if (!changed && !force)
                {
                    return;
                }

                long now = _stopwatch.ElapsedMilliseconds;
                long selectionLast = _selectionLastReportMs[file.Selection];
                if (force
                    || now - selectionLast
                    >= FileLinkUploadProgress.ReportIntervalMs)
                {
                    _selectionLastReportMs[file.Selection] = now;
                    ReportItem(
                        file.Selection,
                        FileLinkUploadStatus.Uploading);
                }
                if (force
                    || now - _lastPhaseReportMs
                    >= FileLinkUploadProgress.ReportIntervalMs)
                {
                    ReportOverall(force);
                }
                if (now - _lastLogMs >= ProgressLogIntervalMs)
                {
                    _lastLogMs = now;
                    DiagnosticsLogger.Log(
                        LogCategories.FileLink,
                        "Upload progress (files="
                        + _completedFileCount.ToString(
                            CultureInfo.InvariantCulture)
                        + "/"
                        + _plan.Files.Count.ToString(
                            CultureInfo.InvariantCulture)
                        + ", bytes="
                        + _uploadedBytes.ToString(
                            CultureInfo.InvariantCulture)
                        + "/"
                        + _plan.TotalBytes.ToString(
                            CultureInfo.InvariantCulture)
                        + ").");
                }
            }
        }

        internal void CompleteFile(FileLinkPlannedFile file)
        {
            lock (_sync)
            {
                ApplyFileBytes(file, file.Length);
                if (!_completedFiles.Add(file))
                {
                    return;
                }

                _completedFileCount++;
                _selectionCompletedFiles[file.Selection]++;
                if (_selectionCompletedFiles[file.Selection]
                    >= _plan.SelectionFileCounts[file.Selection])
                {
                    ReportItem(
                        file.Selection,
                        FileLinkUploadStatus.Completed);
                }
                ReportOverall(false);
            }
        }

        internal void FailFile(FileLinkPlannedFile file)
        {
            if (file == null)
            {
                return;
            }

            lock (_sync)
            {
                _failedSelections.Add(file.Selection);
                ReportItem(
                    file.Selection,
                    FileLinkUploadStatus.Failed);
            }
        }

        internal void FailActiveSelections()
        {
            lock (_sync)
            {
                foreach (FileLinkSelection selection
                    in _plan.SelectionBytes.Keys)
                {
                    if (_failedSelections.Contains(selection))
                    {
                        continue;
                    }
                    if (_selectionCompletedFiles[selection]
                        < _plan.SelectionFileCounts[selection])
                    {
                        ReportItem(
                            selection,
                            FileLinkUploadStatus.Failed);
                    }
                }
            }
        }

        internal void Complete()
        {
            lock (_sync)
            {
                _uploadedBytes = _plan.TotalBytes;
                _completedFileCount = _plan.Files.Count;
                FileLinkUploadProgress.ReportPhase(
                    _phaseProgress,
                    FileLinkUploadPhase.Completed,
                    _plan.DirectoriesToCreate.Count,
                    _plan.DirectoriesToCreate.Count,
                    _completedFileCount,
                    _plan.Files.Count,
                    _uploadedBytes,
                    _plan.TotalBytes);
            }
        }

        private bool ApplyFileBytes(
            FileLinkPlannedFile file,
            long bytes)
        {
            long normalized = Math.Max(
                0,
                Math.Min(file.Length, bytes));
            long previous = _fileBytes[file];
            if (previous == normalized)
            {
                return false;
            }

            long delta = normalized - previous;
            _fileBytes[file] = normalized;
            _uploadedBytes = Math.Max(
                0,
                Math.Min(
                    _plan.TotalBytes,
                    _uploadedBytes + delta));
            _selectionBytes[file.Selection] = Math.Max(
                0,
                Math.Min(
                    _plan.SelectionBytes[file.Selection],
                    _selectionBytes[file.Selection] + delta));
            return true;
        }

        private void ReportItem(
            FileLinkSelection selection,
            FileLinkUploadStatus status)
        {
            if (_itemProgress == null)
            {
                return;
            }

            _itemProgress.Report(new FileLinkUploadItemProgress(
                selection,
                _selectionBytes[selection],
                _plan.SelectionBytes[selection],
                status));
        }

        private void ReportOverall(bool force)
        {
            long now = _stopwatch.ElapsedMilliseconds;
            if (!force
                && now - _lastPhaseReportMs
                < FileLinkUploadProgress.ReportIntervalMs)
            {
                return;
            }

            _lastPhaseReportMs = now;
            FileLinkUploadProgress.ReportPhase(
                _phaseProgress,
                FileLinkUploadPhase.Uploading,
                _plan.DirectoriesToCreate.Count,
                _plan.DirectoriesToCreate.Count,
                _completedFileCount,
                _plan.Files.Count,
                _uploadedBytes,
                _plan.TotalBytes);
        }
    }
}
