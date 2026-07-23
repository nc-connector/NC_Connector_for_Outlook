// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Services
{
    internal enum FileLinkUploadTransport
    {
        Direct,
        Chunked,
        Bulk
    }

    internal sealed class FileLinkPlannedFile
    {
        internal FileLinkPlannedFile(
            FileLinkSelection selection,
            string localPath,
            string remotePath,
            long length,
            DateTime lastWriteTimeUtc)
        {
            Selection = selection;
            LocalPath = localPath ?? string.Empty;
            RemotePath = remotePath ?? string.Empty;
            Length = length;
            LastWriteTimeUtc = lastWriteTimeUtc;
            Transport = FileLinkUploadTransport.Direct;
        }

        internal FileLinkSelection Selection { get; private set; }

        internal string LocalPath { get; private set; }

        internal string RemotePath { get; private set; }

        internal long Length { get; private set; }

        internal DateTime LastWriteTimeUtc { get; private set; }

        internal FileLinkUploadTransport Transport { get; set; }

        internal string BulkChecksum { get; set; }
    }

    internal sealed class FileLinkBulkUploadBatch
    {
        internal FileLinkBulkUploadBatch(
            IEnumerable<FileLinkPlannedFile> files)
        {
            Files = new ReadOnlyCollection<FileLinkPlannedFile>(
                new List<FileLinkPlannedFile>(
                    files
                    ?? Enumerable.Empty<FileLinkPlannedFile>()));
        }

        internal ReadOnlyCollection<FileLinkPlannedFile> Files
        {
            get;
            private set;
        }
    }

    internal sealed class FileLinkUploadPlan
    {
        internal FileLinkUploadPlan(
            IEnumerable<FileLinkPlannedFile> files,
            IEnumerable<string> directoriesToCreate,
            IEnumerable<FileLinkBulkUploadBatch> bulkBatches,
            IDictionary<FileLinkSelection, long> selectionBytes,
            IDictionary<FileLinkSelection, int> selectionFileCounts,
            long totalBytes)
        {
            Files = new ReadOnlyCollection<FileLinkPlannedFile>(
                new List<FileLinkPlannedFile>(
                    files
                    ?? Enumerable.Empty<FileLinkPlannedFile>()));
            DirectFileCount = Files.Count(
                file =>
                    file.Transport == FileLinkUploadTransport.Direct);
            ChunkedFileCount = Files.Count(
                file =>
                    file.Transport == FileLinkUploadTransport.Chunked);
            BulkFiles = new ReadOnlyCollection<FileLinkPlannedFile>(
                Files
                    .Where(
                        file =>
                            file.Transport
                            == FileLinkUploadTransport.Bulk)
                    .ToList());
            DirectAndChunkedFiles =
                new ReadOnlyCollection<FileLinkPlannedFile>(
                    Files
                        .Where(
                            file =>
                                file.Transport
                                != FileLinkUploadTransport.Bulk)
                        .ToList());
            DirectoriesToCreate = new ReadOnlyCollection<string>(
                new List<string>(
                    directoriesToCreate
                    ?? Enumerable.Empty<string>()));
            BulkBatches =
                new ReadOnlyCollection<FileLinkBulkUploadBatch>(
                    new List<FileLinkBulkUploadBatch>(
                        bulkBatches
                        ?? Enumerable.Empty<FileLinkBulkUploadBatch>()));
            SelectionBytes =
                new Dictionary<FileLinkSelection, long>(
                    selectionBytes
                    ?? new Dictionary<FileLinkSelection, long>());
            SelectionFileCounts =
                new Dictionary<FileLinkSelection, int>(
                    selectionFileCounts
                    ?? new Dictionary<FileLinkSelection, int>());
            TotalBytes = Math.Max(0, totalBytes);
        }

        internal ReadOnlyCollection<FileLinkPlannedFile> Files
        {
            get;
            private set;
        }

        internal int DirectFileCount { get; private set; }

        internal int ChunkedFileCount { get; private set; }

        internal ReadOnlyCollection<FileLinkPlannedFile> BulkFiles
        {
            get;
            private set;
        }

        internal ReadOnlyCollection<FileLinkPlannedFile>
            DirectAndChunkedFiles
        {
            get;
            private set;
        }

        internal ReadOnlyCollection<string> DirectoriesToCreate
        {
            get;
            private set;
        }

        internal ReadOnlyCollection<FileLinkBulkUploadBatch> BulkBatches
        {
            get;
            private set;
        }

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
}
