// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Uploads one file directly to its final DAV path.
    internal sealed class FileLinkDirectUploader
    {
        private readonly FileLinkDavClient _davClient;

        internal FileLinkDirectUploader(FileLinkDavClient davClient)
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
            string targetUrl = FileLinkDavClient.BuildFileUrl(
                context.NormalizedBaseUrl,
                context.UserId,
                FileLinkPath.Combine(
                    context.RelativeFolderPath,
                    file.RemotePath));
            NcHttpResponse response = _davClient.SendWithRetry(
                () => CreateRequest(
                    targetUrl,
                    file,
                    coordinator,
                    cancellationToken),
                "direct_put",
                cancellationToken,
                () => coordinator.SetFileBytes(file, 0, true));

            if (response == null
                || !response.HasHttpResponse
                || (response.StatusCode != HttpStatusCode.Created
                    && response.StatusCode != HttpStatusCode.OK
                    && response.StatusCode != HttpStatusCode.NoContent))
            {
                FileLinkDavClient.ThrowFailure(
                    response,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken);
            }
            FileLinkSourceFile.ValidateSnapshot(file);
        }

        private static NcHttpRequestOptions CreateRequest(
            string targetUrl,
            FileLinkPlannedFile file,
            FileLinkUploadProgressCoordinator coordinator,
            CancellationToken cancellationToken)
        {
            return new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = targetUrl,
                TimeoutMs = FileLinkUploadPolicy.UploadTimeoutMs,
                ReadWriteTimeoutMs =
                    FileLinkUploadPolicy.UploadTimeoutMs,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                ContentLength = file.Length,
                ContentType = "application/octet-stream",
                AllowWriteStreamBuffering = false,
                CancellationToken = cancellationToken,
                ConnectionLimit =
                    FileLinkUploadPolicy.MaxParallelRequests,
                Headers = new Dictionary<string, string>
                {
                    {
                        FileLinkUploadPolicy.AutoMkcolHeaderName,
                        "1"
                    }
                },
                BodyWriter = requestStream =>
                {
                    byte[] buffer =
                        new byte[FileLinkSourceFile.BufferSize];
                    FileLinkSourceFile.WriteRange(
                        file,
                        requestStream,
                        0,
                        file.Length,
                        buffer,
                        cancellationToken,
                        transferred => coordinator.SetFileBytes(
                            file,
                            transferred,
                            false));
                }
            };
        }
    }
}
