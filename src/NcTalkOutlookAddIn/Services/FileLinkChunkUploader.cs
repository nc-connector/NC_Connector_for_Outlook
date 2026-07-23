// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Runs Nextcloud chunk upload v2 and its single final MOVE.
    internal sealed class FileLinkChunkUploader
    {
        private readonly FileLinkDavClient _davClient;

        internal FileLinkChunkUploader(FileLinkDavClient davClient)
        {
            if (davClient == null)
            {
                throw new ArgumentNullException("davClient");
            }

            _davClient = davClient;
        }

        internal void Upload(
            FileLinkUploadContext context,
            FileLinkPlannedFile file,
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

            FileLinkSourceFile.ValidateSnapshot(file);
            long totalSize = file.Length;
            long chunkSize =
                FileLinkUploadPolicy.GetChunkUploadChunkSize(totalSize);
            long chunkCount = ((totalSize - 1) / chunkSize) + 1;

            string targetUrl = FileLinkDavClient.BuildFileUrl(
                context.NormalizedBaseUrl,
                context.UserId,
                FileLinkPath.Combine(
                    context.RelativeFolderPath,
                    file.RemotePath));
            string uploadFolderUrl =
                FileLinkDavClient.BuildChunkUploadFolderUrl(
                    context.NormalizedBaseUrl,
                    context.UserId);
            bool cleanupRequired = false;

            try
            {
                CreateUploadFolder(
                    uploadFolderUrl,
                    targetUrl,
                    cancellationToken,
                    () => cleanupRequired = true);
                cleanupRequired = true;
                for (long index = 0; index < chunkCount; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    long offset = index * chunkSize;
                    long length = Math.Min(
                        chunkSize,
                        totalSize - offset);
                    string chunkName = (index + 1).ToString(
                        "00000",
                        CultureInfo.InvariantCulture);
                    UploadChunk(
                        uploadFolderUrl + "/" + chunkName,
                        targetUrl,
                        totalSize,
                        file,
                        offset,
                        length,
                        coordinator,
                        cancellationToken);
                }

                MoveIntoPlace(
                    uploadFolderUrl,
                    targetUrl,
                    totalSize,
                    file.LastWriteTimeUtc,
                    cancellationToken);
                FileLinkSourceFile.ValidateSnapshot(file);
                cleanupRequired = false;
            }
            catch
            {
                if (cleanupRequired)
                {
                    _davClient.DeleteBestEffort(
                        uploadFolderUrl,
                        "Chunk upload cleanup failed");
                }
                throw;
            }
        }

        internal void MoveIntoPlace(
            string uploadFolderUrl,
            string targetUrl,
            long totalSize,
            DateTime lastWriteTimeUtc,
            CancellationToken cancellationToken)
        {
            string sourceUrl = uploadFolderUrl + "/.file";
            long modified = TimeUtilities
                .ToUnixTimeSeconds(lastWriteTimeUtc)
                .GetValueOrDefault();
            NcHttpResponse response = _davClient.Send(
                new NcHttpRequestOptions
                {
                    Method = "MOVE",
                    Url = sourceUrl,
                    TimeoutMs = FileLinkUploadPolicy.UploadTimeoutMs,
                    ReadWriteTimeoutMs =
                        FileLinkUploadPolicy.UploadTimeoutMs,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    CancellationToken = cancellationToken,
                    ConnectionLimit =
                        FileLinkUploadPolicy.MaxParallelRequests,
                    Headers = new Dictionary<string, string>
                    {
                        { "Destination", targetUrl },
                        {
                            "OC-Total-Length",
                            totalSize.ToString(
                                CultureInfo.InvariantCulture)
                        },
                        {
                            "X-OC-Mtime",
                            modified.ToString(
                                CultureInfo.InvariantCulture)
                        }
                    }
                });

            cancellationToken.ThrowIfCancellationRequested();
            if (response != null
                && response.HasHttpResponse
                && (int)response.StatusCode >= 200
                && (int)response.StatusCode < 300)
            {
                return;
            }

            if ((response == null || !response.HasHttpResponse)
                && _davClient.ResourceHasContentLength(
                    targetUrl,
                    totalSize,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken))
            {
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Chunk MOVE recovered after the target length was verified.");
                return;
            }

            FileLinkDavClient.ThrowFailure(
                response,
                Strings.FileLinkWizardUploadFailed,
                cancellationToken);
        }

        private void CreateUploadFolder(
            string uploadFolderUrl,
            string targetUrl,
            CancellationToken cancellationToken,
            Action markFolderMayExist)
        {
            bool createResponseWasIndeterminate = false;
            NcHttpResponse response = _davClient.SendWithRetry(
                () => new NcHttpRequestOptions
                {
                    Method = "MKCOL",
                    Url = uploadFolderUrl,
                    TimeoutMs = 60000,
                    ReadWriteTimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    CancellationToken = cancellationToken,
                    ConnectionLimit =
                        FileLinkUploadPolicy.MaxParallelRequests,
                    Headers = new Dictionary<string, string>
                    {
                        { "Destination", targetUrl }
                    }
                },
                "chunk_folder",
                cancellationToken,
                null,
                retryResponse =>
                {
                    if (retryResponse == null
                        || !retryResponse.HasHttpResponse
                        || FileLinkUploadPolicy
                            .IsIndeterminateStatusCode(
                                (int)retryResponse.StatusCode))
                    {
                        createResponseWasIndeterminate = true;
                        if (markFolderMayExist != null)
                        {
                            markFolderMayExist();
                        }
                    }
                });
            createResponseWasIndeterminate |= response == null
                                              || !response.HasHttpResponse;
            if ((response == null || !response.HasHttpResponse)
                && markFolderMayExist != null)
            {
                markFolderMayExist();
            }
            if (response != null
                && response.HasHttpResponse
                && response.StatusCode == HttpStatusCode.Created)
            {
                return;
            }
            if (createResponseWasIndeterminate
                && (response == null
                    || !response.HasHttpResponse
                    || response.StatusCode
                    == HttpStatusCode.MethodNotAllowed)
                && _davClient.CollectionExists(
                    uploadFolderUrl,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken))
            {
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Chunk upload folder creation recovered after an indeterminate response.");
                return;
            }

            FileLinkDavClient.ThrowFailure(
                response,
                Strings.FileLinkWizardUploadFailed,
                cancellationToken);
        }

        private void UploadChunk(
            string chunkUrl,
            string targetUrl,
            long totalSize,
            FileLinkPlannedFile file,
            long offset,
            long length,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            NcHttpResponse response = _davClient.SendWithRetry(
                () => new NcHttpRequestOptions
                {
                    Method = "PUT",
                    Url = chunkUrl,
                    TimeoutMs = FileLinkUploadPolicy.UploadTimeoutMs,
                    ReadWriteTimeoutMs =
                        FileLinkUploadPolicy.UploadTimeoutMs,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    ContentLength = length,
                    ContentType = "application/octet-stream",
                    AllowWriteStreamBuffering = false,
                    CancellationToken = cancellationToken,
                    ConnectionLimit =
                        FileLinkUploadPolicy.MaxParallelRequests,
                    Headers = new Dictionary<string, string>
                    {
                        { "Destination", targetUrl },
                        {
                            "OC-Total-Length",
                            totalSize.ToString(
                                CultureInfo.InvariantCulture)
                        }
                    },
                    BodyWriter = requestStream =>
                    {
                        byte[] buffer =
                            new byte[FileLinkSourceFile.BufferSize];
                        FileLinkSourceFile.WriteRange(
                            file,
                            requestStream,
                            offset,
                            length,
                            buffer,
                            cancellationToken,
                            transferred =>
                                coordinator.SetFileBytes(
                                    file,
                                    offset + transferred,
                                    false));
                    }
                },
                "chunk_put",
                cancellationToken,
                () => coordinator.SetFileBytes(
                    file,
                    offset,
                    true));

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
        }
    }
}
