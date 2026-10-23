// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class FileLinkSelectionScanResult
    {
        internal FileLinkSelectionScanResult(
            IList<FileLinkPlannedFile> files,
            ISet<string> directories,
            IDictionary<FileLinkSelection, long> selectionBytes,
            IDictionary<FileLinkSelection, int> selectionFileCounts,
            long totalBytes)
        {
            Files = files;
            Directories = directories;
            SelectionBytes = selectionBytes;
            SelectionFileCounts = selectionFileCounts;
            TotalBytes = totalBytes;
        }

        internal IList<FileLinkPlannedFile> Files { get; private set; }

        internal ISet<string> Directories { get; private set; }

        internal IDictionary<FileLinkSelection, long> SelectionBytes
        {
            get;
            private set;
        }

        internal IDictionary<FileLinkSelection, int> SelectionFileCounts
        {
            get;
            private set;
        }

        internal long TotalBytes { get; private set; }
    }

    internal sealed class FileLinkSelectionScanner
    {
        private readonly Func<FileLinkDuplicateInfo, string>
            _duplicateResolver;
        private readonly IDictionary<FileLinkSelection, FileLinkQueueNode>
            _queueSnapshots;
        private readonly CancellationToken _cancellationToken;
        private readonly List<FileLinkPlannedFile> _files =
            new List<FileLinkPlannedFile>();
        private readonly HashSet<string> _directories =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _reservedFiles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _reservedFolders =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<FileLinkSelection, long>
            _selectionBytes =
                new Dictionary<FileLinkSelection, long>();
        private readonly Dictionary<FileLinkSelection, int>
            _selectionFileCounts =
                new Dictionary<FileLinkSelection, int>();
        private long _totalBytes;

        private FileLinkSelectionScanner(
            IDictionary<FileLinkSelection, FileLinkQueueNode>
                queueSnapshots,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken,
            FileLinkUploadContext existingContext = null)
        {
            _queueSnapshots = queueSnapshots;
            _duplicateResolver = duplicateResolver;
            _cancellationToken = cancellationToken;
            if (existingContext != null)
            {
                _reservedFiles.UnionWith(existingContext.ReservedFiles);
                _reservedFolders.UnionWith(existingContext.ReservedDirectories);
            }
        }

        internal static FileLinkSelectionScanResult Scan(
            IList<FileLinkSelection> selections,
            IDictionary<FileLinkSelection, FileLinkQueueNode>
                queueSnapshots,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken,
            FileLinkUploadContext existingContext = null)
        {
            if (selections == null)
            {
                throw new ArgumentNullException("selections");
            }
            if (queueSnapshots == null)
            {
                throw new ArgumentNullException("queueSnapshots");
            }

            return new FileLinkSelectionScanner(
                queueSnapshots,
                duplicateResolver,
                cancellationToken,
                existingContext)
                .ScanSelections(selections);
        }

        private FileLinkSelectionScanResult ScanSelections(
            IEnumerable<FileLinkSelection> selections)
        {
            foreach (FileLinkSelection selection in selections)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                if (selection == null)
                {
                    continue;
                }

                _selectionBytes[selection] = 0;
                _selectionFileCounts[selection] = 0;
                FileLinkQueueNode queueSnapshot;
                if (!_queueSnapshots.TryGetValue(
                        selection,
                        out queueSnapshot)
                    || queueSnapshot == null)
                {
                    throw CreateSourceChangedException();
                }

                if (selection.Source
                    == FileLinkSelectionSource.Nextcloud)
                {
                    AddNextcloudSelection(selection);
                    continue;
                }

                AddLocalSnapshot(
                    selection,
                    queueSnapshot);
            }

            return new FileLinkSelectionScanResult(
                _files,
                _directories,
                _selectionBytes,
                _selectionFileCounts,
                _totalBytes);
        }

        private void AddNextcloudSelection(
            FileLinkSelection selection)
        {
            if (selection.NextcloudEntries == null
                || selection.NextcloudEntries.Count == 0)
            {
                throw CreateSourceChangedException();
            }

            if (selection.SelectionType == FileLinkSelectionType.File)
            {
                NextcloudStorageEntry entry = selection.NextcloudEntries
                    .FirstOrDefault(
                        item => item != null
                                && !item.IsDirectory
                                && string.Equals(
                                    item.RelativePath,
                                    selection.NextcloudPath,
                                    StringComparison.Ordinal));
                if (entry == null)
                {
                    throw CreateSourceChangedException();
                }

                string fileName = FileLinkPath.SanitizeComponent(
                    entry.DisplayName);
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    throw CreateSourceChangedException();
                }
                AddPlannedNextcloudFile(
                    selection,
                    entry,
                    ReserveUniqueName(
                        string.Empty,
                        fileName,
                        selection,
                        false));
                return;
            }

            AddNextcloudDirectory(selection);
        }

        private void AddNextcloudDirectory(
            FileLinkSelection selection)
        {
            NextcloudStorageEntry root = selection.NextcloudEntries
                .FirstOrDefault(
                    item => item != null
                            && item.IsDirectory
                            && string.Equals(
                                item.RelativePath,
                                selection.NextcloudPath,
                                StringComparison.Ordinal));
            if (root == null)
            {
                throw CreateSourceChangedException();
            }

            string rootName = FileLinkPath.SanitizeComponent(
                selection.DisplayName);
            if (string.IsNullOrWhiteSpace(rootName))
            {
                rootName = "Nextcloud";
            }
            string remoteRoot = ReserveUniqueName(
                string.Empty,
                rootName,
                selection,
                true);
            _directories.Add(remoteRoot);

            var directoryTargets = new Dictionary<string, string>(
                StringComparer.Ordinal);
            directoryTargets[root.RelativePath] = remoteRoot;
            IEnumerable<NextcloudStorageEntry> descendants =
                selection.NextcloudEntries
                    .Where(
                        item => item != null
                                && !string.Equals(
                                    item.RelativePath,
                                    root.RelativePath,
                                    StringComparison.Ordinal))
                    .OrderBy(
                        item => NextcloudPath.GetDepth(
                            item.RelativePath))
                    .ThenBy(item => item.IsDirectory ? 0 : 1)
                    .ThenBy(
                        item => item.RelativePath,
                        StringComparer.Ordinal);
            foreach (NextcloudStorageEntry entry in descendants)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                if (!IsDescendantPath(
                    root.RelativePath,
                    entry.RelativePath))
                {
                    throw CreateSourceChangedException();
                }

                string sourceParent = NextcloudPath.GetParent(
                    entry.RelativePath);
                string targetParent;
                if (!directoryTargets.TryGetValue(
                    sourceParent,
                    out targetParent))
                {
                    throw CreateSourceChangedException();
                }
                string name = FileLinkPath.SanitizeComponent(
                    entry.DisplayName);
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = entry.IsDirectory ? "Folder" : "File";
                }
                string uniqueName = ReserveUniqueName(
                    targetParent,
                    name,
                    selection,
                    entry.IsDirectory);
                string targetPath = FileLinkPath.Combine(
                    targetParent,
                    uniqueName);
                if (entry.IsDirectory)
                {
                    directoryTargets[entry.RelativePath] = targetPath;
                    _directories.Add(targetPath);
                }
                else
                {
                    AddPlannedNextcloudFile(
                        selection,
                        entry,
                        targetPath);
                }
            }
        }

        private void AddPlannedNextcloudFile(
            FileLinkSelection selection,
            NextcloudStorageEntry entry,
            string remotePath)
        {
            long length = Math.Max(0, entry.Length);
            _totalBytes = checked(_totalBytes + length);
            _selectionBytes[selection] = checked(
                _selectionBytes[selection] + length);
            _selectionFileCounts[selection] = checked(
                _selectionFileCounts[selection] + 1);
            _files.Add(new FileLinkPlannedFile(
                selection,
                entry.RelativePath,
                remotePath,
                length,
                entry.LastModifiedUtc ?? DateTime.MinValue,
                FileLinkUploadTransport.ServerCopy));
        }

        private static bool IsDescendantPath(
            string parent,
            string candidate)
        {
            string normalizedParent = (parent ?? string.Empty)
                .Trim('/');
            string normalizedCandidate = (candidate ?? string.Empty)
                .Trim('/');
            if (normalizedParent.Length == 0)
            {
                return normalizedCandidate.Length > 0;
            }
            return normalizedCandidate.StartsWith(
                normalizedParent + "/",
                StringComparison.Ordinal);
        }

        private void AddLocalSnapshot(
            FileLinkSelection selection,
            FileLinkQueueNode snapshot)
        {
            bool expectsDirectory =
                selection.SelectionType
                == FileLinkSelectionType.Directory;
            if (snapshot.IsDirectory != expectsDirectory
                || !PathsEqual(
                    selection.LocalPath,
                    snapshot.SourcePath))
            {
                throw CreateSourceChangedException();
            }

            if (!snapshot.IsDirectory)
            {
                string fileName = SanitizeSnapshotName(
                    snapshot,
                    false);
                AddSnapshotFile(
                    selection,
                    snapshot,
                    ReserveUniqueName(
                        string.Empty,
                        fileName,
                        selection,
                        false));
                return;
            }

            string rootName = SanitizeSnapshotName(
                snapshot,
                true);
            string remoteRoot = ReserveUniqueName(
                string.Empty,
                rootName,
                selection,
                true);
            AddSnapshotDirectory(
                selection,
                snapshot,
                remoteRoot);
        }

        private void AddSnapshotDirectory(
            FileLinkSelection selection,
            FileLinkQueueNode directory,
            string remotePath)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            ValidateSnapshotDirectory(directory);
            _directories.Add(remotePath);

            var childDirectories =
                new List<KeyValuePair<FileLinkQueueNode, string>>();
            foreach (FileLinkQueueNode child in directory.Children)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                if (child == null || !child.IsDirectory)
                {
                    continue;
                }
                ValidateDirectChild(directory, child);
                string targetPath = FileLinkPath.Combine(
                    remotePath,
                    ReserveUniqueName(
                        remotePath,
                        SanitizeSnapshotName(child, true),
                        selection,
                        true));
                childDirectories.Add(
                    new KeyValuePair<FileLinkQueueNode, string>(
                        child,
                        targetPath));
            }

            foreach (FileLinkQueueNode child in directory.Children)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                if (child == null || child.IsDirectory)
                {
                    continue;
                }
                ValidateDirectChild(directory, child);
                string targetPath = FileLinkPath.Combine(
                    remotePath,
                    ReserveUniqueName(
                        remotePath,
                        SanitizeSnapshotName(child, false),
                        selection,
                        false));
                AddSnapshotFile(
                    selection,
                    child,
                    targetPath);
            }

            foreach (KeyValuePair<FileLinkQueueNode, string> child
                in childDirectories)
            {
                AddSnapshotDirectory(
                    selection,
                    child.Key,
                    child.Value);
            }
        }

        private void AddSnapshotFile(
            FileLinkSelection selection,
            FileLinkQueueNode snapshot,
            string remotePath)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var fileInfo = new FileInfo(snapshot.SourcePath);
            fileInfo.Refresh();
            if (!fileInfo.Exists
                || snapshot.IsDirectory
                || !snapshot.Length.HasValue
                || !snapshot.LastWriteTimeUtc.HasValue)
            {
                throw CreateSourceChangedException();
            }
            RejectReparsePoint(fileInfo);
            if (fileInfo.Length != snapshot.Length.Value
                || fileInfo.LastWriteTimeUtc
                   != snapshot.LastWriteTimeUtc.Value)
            {
                throw CreateSourceChangedException();
            }

            long length = Math.Max(0, snapshot.Length.Value);
            _totalBytes = checked(_totalBytes + length);
            _selectionBytes[selection] = checked(
                _selectionBytes[selection] + length);
            _selectionFileCounts[selection] = checked(
                _selectionFileCounts[selection] + 1);
            _files.Add(new FileLinkPlannedFile(
                selection,
                fileInfo.FullName,
                remotePath,
                length,
                snapshot.LastWriteTimeUtc.Value));
        }

        private static void ValidateSnapshotDirectory(
            FileLinkQueueNode snapshot)
        {
            if (snapshot == null || !snapshot.IsDirectory)
            {
                throw CreateSourceChangedException();
            }
            var directory = new DirectoryInfo(snapshot.SourcePath);
            directory.Refresh();
            if (!directory.Exists)
            {
                throw CreateSourceChangedException();
            }
            RejectReparsePoint(directory);
        }

        private static void ValidateDirectChild(
            FileLinkQueueNode parent,
            FileLinkQueueNode child)
        {
            string childParent;
            try
            {
                childParent = Path.GetDirectoryName(
                    Path.GetFullPath(child.SourcePath));
            }
            catch
            {
                throw CreateSourceChangedException();
            }
            if (!PathsEqual(parent.SourcePath, childParent))
            {
                throw CreateSourceChangedException();
            }
        }

        private static bool PathsEqual(
            string left,
            string right)
        {
            try
            {
                string normalizedLeft = Path.GetFullPath(
                    left ?? string.Empty)
                    .TrimEnd(
                        Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar);
                string normalizedRight = Path.GetFullPath(
                    right ?? string.Empty)
                    .TrimEnd(
                        Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar);
                return string.Equals(
                    normalizedLeft,
                    normalizedRight,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string SanitizeSnapshotName(
            FileLinkQueueNode snapshot,
            bool isDirectory)
        {
            string name = FileLinkPath.SanitizeComponent(
                snapshot == null
                    ? string.Empty
                    : snapshot.DisplayName);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }
            return isDirectory ? "Folder" : "File";
        }

        private static void RejectReparsePoint(
            FileSystemInfo item)
        {
            if (item != null
                && (item.Attributes & FileAttributes.ReparsePoint)
                == FileAttributes.ReparsePoint)
            {
                throw new TalkServiceException(
                    Strings.FileLinkUploadLinkedItemUnsupported,
                    false,
                    0,
                    null);
            }
        }

        private string ReserveUniqueName(
            string remoteFolder,
            string sanitizedName,
            FileLinkSelection selection,
            bool isDirectory)
        {
            string folderKey = remoteFolder ?? string.Empty;
            string fullPath = FileLinkPath.Combine(
                folderKey,
                sanitizedName);
            ISet<string> primaryReserved = isDirectory
                ? _reservedFolders
                : _reservedFiles;
            ISet<string> secondaryReserved = isDirectory
                ? _reservedFiles
                : _reservedFolders;

            while (primaryReserved.Contains(fullPath)
                || secondaryReserved.Contains(fullPath))
            {
                if (_duplicateResolver == null)
                {
                    throw new TalkServiceException(
                        Strings.FileLinkWizardUploadFailed,
                        false,
                        0,
                        null);
                }

                string replacement = _duplicateResolver(
                    new FileLinkDuplicateInfo(
                        selection,
                        remoteFolder,
                        sanitizedName,
                        isDirectory));
                if (string.IsNullOrWhiteSpace(replacement))
                {
                    throw new OperationCanceledException(
                        Strings.FileLinkWizardUploadCancelledMessage);
                }

                _cancellationToken.ThrowIfCancellationRequested();
                sanitizedName = FileLinkPath.SanitizeComponent(
                    replacement);
                if (string.IsNullOrWhiteSpace(sanitizedName))
                {
                    throw CreateSourceChangedException();
                }
                fullPath = FileLinkPath.Combine(
                    folderKey,
                    sanitizedName);
            }

            primaryReserved.Add(fullPath);
            return sanitizedName;
        }

        private static TalkServiceException
            CreateSourceChangedException()
        {
            return new TalkServiceException(
                Strings.FileLinkUploadSourceChanged,
                false,
                0,
                null);
        }

    }
}
