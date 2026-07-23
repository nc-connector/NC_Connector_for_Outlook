// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal static class FileLinkUploadPlanBuilder
    {
        internal static FileLinkUploadPlan Build(
            IList<FileLinkSelection> selections,
            bool bulkUploadSupported,
            int fixedRequestCount,
            Func<FileLinkDuplicateInfo, string> duplicateResolver,
            CancellationToken cancellationToken)
        {
            FileLinkSelectionScanResult scan =
                FileLinkSelectionScanner.Scan(
                    selections,
                    duplicateResolver,
                    cancellationToken);
            IList<FileLinkPlannedFile> files = scan.Files;
            ISet<string> directories = scan.Directories;

            foreach (FileLinkPlannedFile file in files)
            {
                file.Transport = FileLinkUploadPolicy.ShouldUseChunkedUpload(file.Length)
                    ? FileLinkUploadTransport.Chunked
                    : FileLinkUploadTransport.Direct;
            }

            List<FileLinkPlannedFile> bulkCandidates = files
                .Where(file => FileLinkUploadPolicy.IsBulkCandidate(file.Length))
                .ToList();
            List<FileLinkBulkUploadBatch> bulkBatches = BuildBulkBatches(bulkCandidates);
            List<string> directDirectories = BuildRequiredDirectories(
                directories,
                files,
                files.Where(
                    file =>
                        file.Transport
                        == FileLinkUploadTransport.Chunked),
                files.Where(
                    file =>
                        file.Transport
                        == FileLinkUploadTransport.Direct));
            List<string> bulkDirectories = BuildRequiredDirectories(
                directories,
                files,
                files.Where(
                    file =>
                        file.Transport
                        == FileLinkUploadTransport.Chunked
                        || FileLinkUploadPolicy.IsBulkCandidate(
                            file.Length)),
                files.Where(
                    file =>
                        file.Transport
                        == FileLinkUploadTransport.Direct
                        && !FileLinkUploadPolicy.IsBulkCandidate(
                            file.Length)));
            long totalFileRequestCount = 0;
            foreach (FileLinkPlannedFile file in files)
            {
                totalFileRequestCount = checked(
                    totalFileRequestCount
                    + FileLinkUploadPolicy.GetTransferRequestCount(
                        file.Length));
            }
            bool useBulkUpload = FileLinkUploadPolicy.ShouldUseBulkUpload(
                bulkUploadSupported,
                bulkCandidates.Count,
                totalFileRequestCount,
                Math.Max(0, fixedRequestCount)
                    + directDirectories.Count,
                Math.Max(0, fixedRequestCount)
                    + bulkDirectories.Count,
                bulkBatches.Count);
            List<string> selectedDirectories;
            if (useBulkUpload)
            {
                foreach (FileLinkPlannedFile file in bulkCandidates)
                {
                    file.Transport = FileLinkUploadTransport.Bulk;
                }
                selectedDirectories = bulkDirectories;
            }
            else
            {
                bulkBatches.Clear();
                selectedDirectories = directDirectories;
            }

            return new FileLinkUploadPlan(
                files,
                selectedDirectories,
                bulkBatches,
                scan.SelectionBytes,
                scan.SelectionFileCounts,
                scan.TotalBytes);
        }

        private static List<FileLinkBulkUploadBatch> BuildBulkBatches(
            IList<FileLinkPlannedFile> candidates)
        {
            var result = new List<FileLinkBulkUploadBatch>();
            var current = new List<FileLinkPlannedFile>();
            long currentBytes = 0;

            foreach (FileLinkPlannedFile file in candidates)
            {
                bool batchFull = current.Count >= FileLinkUploadPolicy.BulkBatchFileLimit;
                bool byteLimitReached = current.Count > 0
                                        && currentBytes + file.Length
                                        > FileLinkUploadPolicy.BulkBatchLimitBytes;
                if (batchFull || byteLimitReached)
                {
                    result.Add(new FileLinkBulkUploadBatch(current));
                    current = new List<FileLinkPlannedFile>();
                    currentBytes = 0;
                }

                current.Add(file);
                currentBytes += file.Length;
            }
            if (current.Count > 0)
            {
                result.Add(new FileLinkBulkUploadBatch(current));
            }
            return result;
        }

        private static List<string> BuildRequiredDirectories(
            ISet<string> directories,
            IEnumerable<FileLinkPlannedFile> allFiles,
            IEnumerable<FileLinkPlannedFile> filesRequiringParents,
            IEnumerable<FileLinkPlannedFile> directFiles)
        {
            var knownDirectories = new HashSet<string>(
                directories,
                StringComparer.OrdinalIgnoreCase);
            var directoriesWithChildren = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (string directory in knownDirectories)
            {
                string parent = FileLinkPath.GetParent(directory);
                if (!string.IsNullOrEmpty(parent))
                {
                    directoriesWithChildren.Add(parent);
                }
            }

            var directoriesWithFiles = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (FileLinkPlannedFile file in allFiles)
            {
                string parent = FileLinkPath.GetParent(file.RemotePath);
                if (!string.IsNullOrEmpty(parent))
                {
                    directoriesWithFiles.Add(parent);
                }
            }

            var required = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (string directory in knownDirectories)
            {
                if (!directoriesWithChildren.Contains(directory)
                    && !directoriesWithFiles.Contains(directory))
                {
                    AddDirectoryAndParents(
                        directory,
                        knownDirectories,
                        required);
                }
            }
            foreach (FileLinkPlannedFile file in filesRequiringParents)
            {
                AddDirectoryAndParents(
                    FileLinkPath.GetParent(file.RemotePath),
                    knownDirectories,
                    required);
            }
            AddSharedDirectDirectories(
                directFiles,
                knownDirectories,
                required);

            return required
                .OrderBy(FileLinkPath.GetDepth)
                .ThenBy(
                    path => path,
                    StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void AddSharedDirectDirectories(
            IEnumerable<FileLinkPlannedFile> directFiles,
            ISet<string> knownDirectories,
            ISet<string> requiredDirectories)
        {
            var directFileCounts = new Dictionary<string, int>(
                StringComparer.OrdinalIgnoreCase);
            foreach (FileLinkPlannedFile file in directFiles)
            {
                string current = FileLinkPath.GetParent(
                    file.RemotePath);
                while (!string.IsNullOrEmpty(current)
                       && knownDirectories.Contains(current))
                {
                    int currentCount;
                    directFileCounts.TryGetValue(
                        current,
                        out currentCount);
                    directFileCounts[current] = checked(
                        currentCount + 1);
                    current = FileLinkPath.GetParent(current);
                }
            }

            foreach (KeyValuePair<string, int> pair
                in directFileCounts)
            {
                if (pair.Value >= 2)
                {
                    AddDirectoryAndParents(
                        pair.Key,
                        knownDirectories,
                        requiredDirectories);
                }
            }
        }

        private static void AddDirectoryAndParents(
            string directory,
            ISet<string> knownDirectories,
            ISet<string> requiredDirectories)
        {
            string current = directory ?? string.Empty;
            while (!string.IsNullOrEmpty(current)
                   && knownDirectories.Contains(current))
            {
                requiredDirectories.Add(current);
                current = FileLinkPath.GetParent(current);
            }
        }

    }
}
