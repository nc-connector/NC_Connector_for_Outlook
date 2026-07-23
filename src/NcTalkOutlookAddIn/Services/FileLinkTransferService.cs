// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Coordinates bulk, direct, and chunked upload workers.
    internal sealed class FileLinkTransferService
    {
        private readonly FileLinkBulkUploader _bulkUploader;
        private readonly FileLinkDirectUploader _directUploader;
        private readonly FileLinkChunkUploader _chunkUploader;

        internal FileLinkTransferService(FileLinkDavClient davClient)
        {
            if (davClient == null)
            {
                throw new ArgumentNullException("davClient");
            }

            _bulkUploader = new FileLinkBulkUploader(davClient);
            _directUploader = new FileLinkDirectUploader(davClient);
            _chunkUploader = new FileLinkChunkUploader(davClient);
        }

        internal void PrepareBulkChecksums(
            FileLinkUploadPlan plan,
            CancellationToken cancellationToken)
        {
            _bulkUploader.PrepareChecksums(plan, cancellationToken);
        }

        internal void Upload(
            FileLinkUploadContext context,
            IProgress<FileLinkUploadItemProgress> itemProgress,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress,
            CancellationToken cancellationToken)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            FileLinkUploadPlan plan = context.Plan;
            var coordinator = new FileLinkUploadProgressCoordinator(
                plan,
                itemProgress,
                phaseProgress);
            coordinator.Begin();
            var stopwatch = Stopwatch.StartNew();

            using (DiagnosticsLogger.BeginOperation(
                LogCategories.FileLink,
                "FileLink.UploadTransfer"))
            {
                try
                {
                    UploadBulkBatches(
                        context,
                        plan,
                        coordinator,
                        cancellationToken);
                    UploadRemainingFiles(
                        context,
                        plan.DirectAndChunkedFiles,
                        coordinator,
                        cancellationToken);
                    coordinator.Complete();
                }
                catch
                {
                    coordinator.FailActiveSelections();
                    throw;
                }
            }

            stopwatch.Stop();
            double seconds = Math.Max(
                0.001,
                stopwatch.Elapsed.TotalSeconds);
            long bytesPerSecond = (long)Math.Round(
                plan.TotalBytes / seconds,
                MidpointRounding.AwayFromZero);
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Upload completed (files="
                + plan.Files.Count.ToString(CultureInfo.InvariantCulture)
                + ", bytes="
                + plan.TotalBytes.ToString(CultureInfo.InvariantCulture)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString(
                    CultureInfo.InvariantCulture)
                + ", bytesPerSecond="
                + bytesPerSecond.ToString(CultureInfo.InvariantCulture)
                + ").");
        }

        private void UploadBulkBatches(
            FileLinkUploadContext context,
            FileLinkUploadPlan plan,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            foreach (FileLinkBulkUploadBatch batch in plan.BulkBatches)
            {
                cancellationToken.ThrowIfCancellationRequested();
                _bulkUploader.Upload(
                    context,
                    batch,
                    coordinator,
                    cancellationToken);
                foreach (FileLinkPlannedFile file in batch.Files)
                {
                    coordinator.CompleteFile(file);
                }
            }
        }

        private void UploadRemainingFiles(
            FileLinkUploadContext context,
            IEnumerable<FileLinkPlannedFile> files,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            List<FileLinkPlannedFile> pending = files.ToList();
            if (pending.Count == 0)
            {
                return;
            }

            using (var linkedCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken))
            {
                var options = new ParallelOptions
                {
                    CancellationToken = linkedCancellation.Token,
                    MaxDegreeOfParallelism =
                        FileLinkUploadPolicy.MaxParallelRequests
                };
                try
                {
                    Parallel.ForEach(
                        pending,
                        options,
                        file => UploadFile(
                            context,
                            file,
                            coordinator,
                            linkedCancellation));
                }
                catch (AggregateException ex)
                {
                    ParallelExecution.RethrowFirstFailure(
                        ex,
                        cancellationToken);
                }
            }
        }

        private void UploadFile(
            FileLinkUploadContext context,
            FileLinkPlannedFile file,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationTokenSource linkedCancellation)
        {
            try
            {
                if (file.Transport == FileLinkUploadTransport.Chunked)
                {
                    _chunkUploader.Upload(
                        context,
                        file,
                        coordinator,
                        linkedCancellation.Token);
                }
                else
                {
                    _directUploader.Upload(
                        context,
                        file,
                        coordinator,
                        linkedCancellation.Token);
                }
                coordinator.CompleteFile(file);
            }
            catch
            {
                coordinator.FailFile(file);
                linkedCancellation.Cancel();
                throw;
            }
        }
    }
}
