// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Utilities
{
    // Centralizes upload thresholds, batching, retry status, and worker limits.
    internal static class FileLinkUploadPolicy
    {
        internal const string AutoMkcolHeaderName =
            "X-NC-WebDAV-Auto-Mkcol";
        internal const int CleanupTimeoutMs = 10000;
        internal const int UploadTimeoutMs = 300000;
        internal const long DirectUploadLimitBytes = 20L * 1024L * 1024L;
        internal const long ChunkUploadMinimumChunkSizeBytes =
            5L * 1024L * 1024L;
        internal const long ChunkUploadChunkSizeBytes =
            20L * 1024L * 1024L;
        internal const long ChunkUploadMaximumChunkSizeBytes =
            5L * 1024L * 1024L * 1024L;
        internal const int ChunkUploadMaxChunks = 10000;
        internal const long ChunkUploadMaximumFileSizeBytes =
            ChunkUploadMaximumChunkSizeBytes * ChunkUploadMaxChunks;
        internal const long BulkCandidateLimitBytes = 8L * 1024L * 1024L;
        internal const long BulkBatchLimitBytes = 20L * 1024L * 1024L;
        internal const int BulkBatchFileLimit = 100;
        internal const int BulkMinimumFileCount = 20;
        internal const int MaxParallelRequests = 3;
        internal const int MaxRequestAttempts = 3;

        internal static bool ShouldUseChunkedUpload(long fileSize)
        {
            return fileSize > DirectUploadLimitBytes;
        }

        internal static bool IsBulkCandidate(long fileSize)
        {
            return fileSize >= 0 && fileSize <= BulkCandidateLimitBytes;
        }

        internal static bool IsSupportedFileSize(long fileSize)
        {
            return fileSize >= 0
                   && fileSize <= ChunkUploadMaximumFileSizeBytes;
        }

        internal static long GetChunkUploadChunkSize(long fileSize)
        {
            if (!IsSupportedFileSize(fileSize))
            {
                throw new ArgumentOutOfRangeException("fileSize");
            }

            long minimumForChunkLimit = fileSize == 0
                ? 0
                : ((fileSize - 1) / ChunkUploadMaxChunks) + 1;
            return Math.Max(
                ChunkUploadMinimumChunkSizeBytes,
                Math.Max(
                    ChunkUploadChunkSizeBytes,
                    minimumForChunkLimit));
        }

        internal static int GetTransferRequestCount(long fileSize)
        {
            if (!IsSupportedFileSize(fileSize))
            {
                throw new ArgumentOutOfRangeException("fileSize");
            }
            if (!ShouldUseChunkedUpload(fileSize))
            {
                return 1;
            }

            long chunkSize = GetChunkUploadChunkSize(fileSize);
            long chunkCount = ((fileSize - 1) / chunkSize) + 1;
            return checked((int)chunkCount + 2);
        }

        internal static bool ShouldUseBulkUpload(
            bool capabilityAvailable,
            int candidateFileCount,
            long totalFileRequestCount,
            long directNonFileRequestCount,
            long bulkNonFileRequestCount,
            int batchRequestCount)
        {
            if (!capabilityAvailable
                || candidateFileCount < BulkMinimumFileCount
                || batchRequestCount <= 0)
            {
                return false;
            }

            long normalizedTotalFileRequestCount = Math.Max(
                candidateFileCount,
                totalFileRequestCount);
            long normalizedDirectNonFileRequestCount = Math.Max(
                0,
                directNonFileRequestCount);
            long normalizedBulkNonFileRequestCount = Math.Max(
                0,
                bulkNonFileRequestCount);
            long remainingFileRequests =
                normalizedTotalFileRequestCount - candidateFileCount;
            long directRequests =
                normalizedDirectNonFileRequestCount
                + normalizedTotalFileRequestCount;
            long bulkRequests =
                normalizedBulkNonFileRequestCount
                + remainingFileRequests
                + batchRequestCount;
            return (decimal)bulkRequests * 5m
                   <= (decimal)directRequests * 4m;
        }

        internal static bool IsRetryableStatusCode(int statusCode)
        {
            return statusCode == 408
                   || statusCode == 423
                   || statusCode == 429
                   || statusCode == 502
                   || statusCode == 503
                   || statusCode == 504;
        }

        internal static bool IsIndeterminateStatusCode(int statusCode)
        {
            return statusCode == 408
                   || statusCode == 502
                   || statusCode == 503
                   || statusCode == 504;
        }
    }
}
