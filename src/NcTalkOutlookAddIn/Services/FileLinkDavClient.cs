// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Owns DAV collection lifecycle operations.
    internal sealed partial class FileLinkDavClient
    {
        private readonly Func<NcHttpRequestOptions, NcHttpResponse>
            _sendRequest;

        internal FileLinkDavClient(NcHttpClient httpClient)
            : this(
                httpClient == null
                    ? null
                    : new Func<NcHttpRequestOptions, NcHttpResponse>(
                        httpClient.Send))
        {
        }

        internal FileLinkDavClient(
            Func<NcHttpRequestOptions, NcHttpResponse> sendRequest)
        {
            if (sendRequest == null)
            {
                throw new ArgumentNullException("sendRequest");
            }

            _sendRequest = sendRequest;
        }

        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            return _sendRequest(options);
        }

        internal void EnsureFolderPath(
            string baseUrl,
            string userId,
            string relativePath,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return;
            }

            string[] segments = relativePath.Split(
                new[] { '/' },
                StringSplitOptions.RemoveEmptyEntries);
            var current = new List<string>();
            foreach (string segment in segments)
            {
                cancellationToken.ThrowIfCancellationRequested();
                current.Add(segment);
                string path = string.Join("/", current.ToArray());
                string url = BuildFileUrl(baseUrl, userId, path);
                NcHttpResponse response = SendWithRetry(
                    () => CreateRequest("MKCOL", url, cancellationToken),
                    "base_folder",
                    cancellationToken,
                    null);
                if (response.HasHttpResponse
                    && response.StatusCode == HttpStatusCode.Created)
                {
                    continue;
                }
                if (response.HasHttpResponse
                    && response.StatusCode == HttpStatusCode.MethodNotAllowed
                    && CollectionExists(
                        url,
                        Strings.FileLinkWizardFolderCheckFailedFormat,
                        cancellationToken))
                {
                    continue;
                }

                ThrowFailure(
                    response,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken);
            }
        }

        internal bool TryCreateShareRoot(
            string baseUrl,
            string userId,
            string relativeFolderPath,
            CancellationToken cancellationToken)
        {
            string url = BuildFileUrl(
                baseUrl,
                userId,
                relativeFolderPath);
            NcHttpResponse response = SendWithRetry(
                () => CreateRequest("MKCOL", url, cancellationToken),
                "share_root",
                cancellationToken,
                null);
            if (response != null
                && response.HasHttpResponse
                && response.StatusCode == HttpStatusCode.Created)
            {
                return true;
            }
            if (response != null
                && response.HasHttpResponse
                && response.StatusCode == HttpStatusCode.MethodNotAllowed)
            {
                return false;
            }

            ThrowFailure(
                response,
                Strings.FileLinkWizardUploadFailed,
                cancellationToken);
            return false;
        }

        internal void CreatePlannedDirectories(
            FileLinkUploadContext context,
            FileLinkUploadPlan plan,
            IProgress<FileLinkUploadPhaseProgress> phaseProgress,
            CancellationToken cancellationToken)
        {
            int totalFolders = plan.DirectoriesToCreate.Count;
            var reporter = new FileLinkFolderProgressReporter(
                phaseProgress,
                totalFolders,
                plan.Files.Count,
                plan.TotalBytes);
            reporter.Report(0, true);
            if (totalFolders == 0)
            {
                return;
            }

            int completed = 0;
            IEnumerable<IGrouping<int, string>> depthGroups =
                plan.DirectoriesToCreate
                    .GroupBy(FileLinkPath.GetDepth)
                    .OrderBy(group => group.Key);
            foreach (IGrouping<int, string> depthGroup in depthGroups)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var options = new ParallelOptions
                {
                    CancellationToken = cancellationToken,
                    MaxDegreeOfParallelism =
                        FileLinkUploadPolicy.MaxParallelRequests
                };
                try
                {
                    Parallel.ForEach(
                        depthGroup,
                        options,
                        remotePath =>
                        {
                            CreatePlannedDirectory(
                                context,
                                remotePath,
                                cancellationToken);
                            int current = Interlocked.Increment(ref completed);
                            reporter.Report(
                                current,
                                current == totalFolders);
                        });
                }
                catch (AggregateException ex)
                {
                    ParallelExecution.RethrowFirstFailure(
                        ex,
                        cancellationToken);
                }
            }
        }

        internal void DeleteShareFolder(
            string baseUrl,
            string userId,
            string relativeFolderPath,
            CancellationToken cancellationToken)
        {
            string url = BuildFileUrl(
                baseUrl,
                userId,
                relativeFolderPath);
            NcHttpResponse response = Send(new NcHttpRequestOptions
            {
                Method = "DELETE",
                Url = url,
                TimeoutMs = 90000,
                ReadWriteTimeoutMs = 90000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                CancellationToken = cancellationToken
            });

            cancellationToken.ThrowIfCancellationRequested();
            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException(
                    transport != null
                        ? transport.Message
                        : Strings.ErrorServerUnavailable,
                    false,
                    0,
                    null);
            }
            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Share folder delete skipped because the folder is already absent.");
                return;
            }
            if (response.StatusCode != HttpStatusCode.NoContent
                && response.StatusCode != HttpStatusCode.OK
                && response.StatusCode != HttpStatusCode.Accepted)
            {
                bool authError =
                    response.StatusCode == HttpStatusCode.Unauthorized
                    || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException(
                    "HTTP "
                    + ((int)response.StatusCode).ToString(
                        CultureInfo.InvariantCulture),
                    authError,
                    response.StatusCode,
                    response.ResponseText);
            }
        }

        internal void DeleteBestEffort(
            string url,
            string failureLogMessage)
        {
            try
            {
                NcHttpResponse response = Send(
                    new NcHttpRequestOptions
                    {
                        Method = "DELETE",
                        Url = url,
                        TimeoutMs =
                            FileLinkUploadPolicy.CleanupTimeoutMs,
                        ReadWriteTimeoutMs =
                            FileLinkUploadPolicy.CleanupTimeoutMs,
                        IncludeAuthHeader = true,
                        IncludeOcsApiHeader = false,
                        ParseJson = false
                    });
                if (response.HasHttpResponse
                    && ((int)response.StatusCode >= 200
                        && (int)response.StatusCode < 300
                        || response.StatusCode == HttpStatusCode.NotFound))
                {
                    return;
                }

                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    failureLogMessage
                    + " (status="
                    + (response.HasHttpResponse
                        ? ((int)response.StatusCode).ToString(
                            CultureInfo.InvariantCulture)
                        : "n/a")
                    + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    failureLogMessage + ".",
                    ex);
            }
        }

        private void CreatePlannedDirectory(
            FileLinkUploadContext context,
            string remotePath,
            CancellationToken cancellationToken)
        {
            string fullPath = FileLinkPath.Combine(
                context.RelativeFolderPath,
                remotePath);
            string url = BuildFileUrl(
                context.NormalizedBaseUrl,
                context.UserId,
                fullPath);
            bool createResponseWasIndeterminate = false;
            NcHttpResponse response = SendWithRetry(
                () => CreateRequest("MKCOL", url, cancellationToken),
                "planned_folder",
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
                    }
                });
            createResponseWasIndeterminate |= response == null
                                              || !response.HasHttpResponse;
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
                && CollectionExists(
                    url,
                    Strings.FileLinkWizardUploadFailed,
                    cancellationToken))
            {
                return;
            }

            ThrowFailure(
                response,
                Strings.FileLinkWizardUploadFailed,
                cancellationToken);
        }
    }
}
