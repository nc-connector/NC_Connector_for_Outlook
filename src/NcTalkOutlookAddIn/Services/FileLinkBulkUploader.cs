// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Writes and validates DAV bulk batches, including full-batch retries.
    internal sealed class FileLinkBulkUploader
    {
        private readonly FileLinkDavClient _davClient;

        internal FileLinkBulkUploader(FileLinkDavClient davClient)
        {
            if (davClient == null)
            {
                throw new ArgumentNullException("davClient");
            }

            _davClient = davClient;
        }

        internal void PrepareChecksums(
            FileLinkUploadPlan plan,
            CancellationToken cancellationToken)
        {
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            foreach (FileLinkPlannedFile file in plan.BulkFiles)
            {
                cancellationToken.ThrowIfCancellationRequested();
                file.BulkChecksum =
                    FileLinkSourceFile.ComputeMd5Hex(file);
            }
        }

        internal void Upload(
            FileLinkUploadContext context,
            FileLinkBulkUploadBatch batch,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            if (coordinator == null)
            {
                throw new ArgumentNullException("coordinator");
            }
            if (batch == null || batch.Files.Count == 0)
            {
                return;
            }

            foreach (FileLinkPlannedFile file in batch.Files)
            {
                FileLinkSourceFile.ValidateSnapshot(file);
            }

            string boundary =
                "ncconnector-" + Guid.NewGuid().ToString("N");
            long contentLength = CalculateContentLength(
                batch,
                boundary,
                context.RelativeFolderPath);
            string url = context.NormalizedBaseUrl.TrimEnd('/')
                         + "/remote.php/dav/bulk";
            NcHttpResponse response = _davClient.SendWithRetry(
                () => CreateRequest(
                    context,
                    batch,
                    coordinator,
                    boundary,
                    contentLength,
                    url,
                    cancellationToken),
                "bulk_post",
                cancellationToken,
                () => ResetBatchProgress(batch, coordinator));

            cancellationToken.ThrowIfCancellationRequested();
            if (response == null
                || !response.HasHttpResponse
                || (int)response.StatusCode < 200
                || (int)response.StatusCode >= 300)
            {
                FileLinkDavClient.ThrowFailure(
                    response,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken);
            }
            if (response.ParsedJson == null)
            {
                throw new TalkServiceException(
                    Strings.FileLinkWizardUploadFailed,
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }

            ValidateResponse(
                context.RelativeFolderPath,
                batch,
                response);
        }

        private static NcHttpRequestOptions CreateRequest(
            FileLinkUploadContext context,
            FileLinkBulkUploadBatch batch,
            FileLinkUploadProgressCoordinator coordinator,
            string boundary,
            long contentLength,
            string url,
            CancellationToken cancellationToken)
        {
            return new NcHttpRequestOptions
            {
                Method = "POST",
                Url = url,
                ContentType =
                    "multipart/related; boundary=" + boundary,
                Accept = "application/json",
                TimeoutMs = FileLinkUploadPolicy.UploadTimeoutMs,
                ReadWriteTimeoutMs =
                    FileLinkUploadPolicy.UploadTimeoutMs,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = true,
                ContentLength = contentLength,
                AllowWriteStreamBuffering = false,
                CancellationToken = cancellationToken,
                ConnectionLimit =
                    FileLinkUploadPolicy.MaxParallelRequests,
                BodyWriter = requestStream => WriteBody(
                    requestStream,
                    batch,
                    boundary,
                    context.RelativeFolderPath,
                    coordinator,
                    cancellationToken)
            };
        }

        private static void ValidateResponse(
            string relativeFolderPath,
            FileLinkBulkUploadBatch batch,
            NcHttpResponse response)
        {
            foreach (FileLinkPlannedFile file in batch.Files)
            {
                string responsePath = "/"
                                      + FileLinkPath.Combine(
                                          relativeFolderPath,
                                          file.RemotePath);
                IDictionary<string, object> fileResult =
                    NcJson.GetDictionary(
                        response.ParsedJson,
                        responsePath);
                bool hasError;
                if (fileResult == null
                    || !NcJson.TryGetBoolean(
                        fileResult,
                        "error",
                        out hasError)
                    || hasError)
                {
                    string detail = NcJson.GetTrimmedString(
                        fileResult,
                        "message");
                    DiagnosticsLogger.Log(
                        LogCategories.FileLink,
                        "Bulk upload part failed (path="
                        + file.RemotePath
                        + ", detail="
                        + (string.IsNullOrWhiteSpace(detail)
                            ? "not provided"
                            : detail)
                        + ").");
                    throw new TalkServiceException(
                        Strings.FileLinkWizardUploadFailed,
                        false,
                        response.StatusCode,
                        response.ResponseText);
                }

                FileLinkSourceFile.ValidateSnapshot(file);
            }
        }

        private static void ResetBatchProgress(
            FileLinkBulkUploadBatch batch,
            FileLinkUploadProgressCoordinator coordinator)
        {
            foreach (FileLinkPlannedFile file in batch.Files)
            {
                coordinator.SetFileBytes(file, 0, true);
            }
        }

        private static long CalculateContentLength(
            FileLinkBulkUploadBatch batch,
            string boundary,
            string remoteRootPath)
        {
            long length = 0;
            foreach (FileLinkPlannedFile file in batch.Files)
            {
                string prefix = BuildPartPrefix(
                    file,
                    boundary,
                    remoteRootPath);
                length = checked(
                    length + Encoding.UTF8.GetByteCount(prefix));
                length = checked(length + file.Length);
                length = checked(length + 2);
            }
            length = checked(
                length
                + Encoding.UTF8.GetByteCount(
                    "--" + boundary + "--\r\n"));
            return length;
        }

        private static void WriteBody(
            Stream destination,
            FileLinkBulkUploadBatch batch,
            string boundary,
            string remoteRootPath,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            byte[] buffer = new byte[FileLinkSourceFile.BufferSize];
            foreach (FileLinkPlannedFile file in batch.Files)
            {
                cancellationToken.ThrowIfCancellationRequested();
                WriteUtf8(
                    destination,
                    BuildPartPrefix(
                        file,
                        boundary,
                        remoteRootPath));
                FileLinkSourceFile.WriteRange(
                    file,
                    destination,
                    0,
                    file.Length,
                    buffer,
                    cancellationToken,
                    transferred => coordinator.SetFileBytes(
                        file,
                        transferred,
                        false));
                WriteUtf8(destination, "\r\n");
            }
            WriteUtf8(destination, "--" + boundary + "--\r\n");
        }

        private static string BuildPartPrefix(
            FileLinkPlannedFile file,
            string boundary,
            string remoteRootPath)
        {
            long modified = TimeUtilities
                .ToUnixTimeSeconds(file.LastWriteTimeUtc)
                .GetValueOrDefault();
            var builder = new StringBuilder();
            builder.Append("--").Append(boundary).Append("\r\n");
            builder.Append("Content-Length: ")
                .Append(
                    file.Length.ToString(
                        CultureInfo.InvariantCulture))
                .Append("\r\n");
            builder.Append(
                "Content-Type: application/octet-stream\r\n");
            builder.Append("X-File-MD5: ")
                .Append(file.BulkChecksum ?? string.Empty)
                .Append("\r\n");
            builder.Append("X-File-Mtime: ")
                .Append(
                    modified.ToString(
                        CultureInfo.InvariantCulture))
                .Append("\r\n");
            builder.Append("X-File-Path: /")
                .Append(
                    FileLinkPath.Combine(
                        remoteRootPath,
                        file.RemotePath))
                .Append("\r\n\r\n");
            return builder.ToString();
        }

        private static void WriteUtf8(
            Stream destination,
            string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(
                value ?? string.Empty);
            destination.Write(bytes, 0, bytes.Length);
        }
    }
}
