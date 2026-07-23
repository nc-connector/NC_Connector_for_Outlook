// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Owns DAV request retries, failure mapping, and URL construction.
    internal sealed partial class FileLinkDavClient
    {
        internal NcHttpResponse SendWithRetry(
            Func<NcHttpRequestOptions> requestFactory,
            string operation,
            CancellationToken cancellationToken,
            Action resetProgress,
            Action<NcHttpResponse> retryObserver = null)
        {
            NcHttpResponse response = null;
            for (int attempt = 1;
                attempt <= FileLinkUploadPolicy.MaxRequestAttempts;
                attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                response = Send(requestFactory());

                cancellationToken.ThrowIfCancellationRequested();
                bool retryable = response.HasHttpResponse
                    ? FileLinkUploadPolicy.IsRetryableStatusCode(
                        (int)response.StatusCode)
                    : IsRetryableTransport(response.TransportException);
                if (!retryable
                    || attempt >= FileLinkUploadPolicy.MaxRequestAttempts)
                {
                    return response;
                }

                if (retryObserver != null)
                {
                    retryObserver(response);
                }
                if (resetProgress != null)
                {
                    resetProgress();
                }
                TimeSpan delay = ResolveRetryDelay(response, attempt);
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Upload request retry scheduled (operation="
                    + operation
                    + ", attempt="
                    + (attempt + 1).ToString(CultureInfo.InvariantCulture)
                    + ", status="
                    + (response.HasHttpResponse
                        ? ((int)response.StatusCode).ToString(
                            CultureInfo.InvariantCulture)
                        : "transport")
                    + ", delayMs="
                    + ((long)delay.TotalMilliseconds).ToString(
                        CultureInfo.InvariantCulture)
                    + ").");
                WaitForRetry(delay, cancellationToken);
            }
            return response;
        }

        internal static NcHttpRequestOptions CreateRequest(
            string method,
            string url,
            CancellationToken cancellationToken)
        {
            return new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                TimeoutMs = 60000,
                ReadWriteTimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                CancellationToken = cancellationToken,
                ConnectionLimit = FileLinkUploadPolicy.MaxParallelRequests
            };
        }

        internal static void ThrowFailure(
            NcHttpResponse response,
            string fallbackMessage,
            CancellationToken cancellationToken,
            bool appendTechnicalDetail = true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string message = string.IsNullOrWhiteSpace(fallbackMessage)
                ? Strings.FileLinkWizardUploadFailed
                : fallbackMessage;

            if (response == null || !response.HasHttpResponse)
            {
                Exception transport = response != null
                    ? response.TransportException
                    : null;
                throw new TalkServiceException(
                    message
                    + (appendTechnicalDetail && transport != null
                        ? " " + transport.Message
                        : string.Empty),
                    false,
                    0,
                    null);
            }
            if ((int)response.StatusCode == 507)
            {
                throw new TalkServiceException(
                    Strings.FileLinkUploadInsufficientStorage,
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }

            bool authError =
                response.StatusCode == HttpStatusCode.Unauthorized
                || response.StatusCode == HttpStatusCode.Forbidden;
            throw new TalkServiceException(
                message
                + (appendTechnicalDetail
                    ? " (HTTP "
                      + ((int)response.StatusCode).ToString(
                          CultureInfo.InvariantCulture)
                      + ")"
                    : string.Empty),
                authError,
                response.StatusCode,
                response.ResponseText);
        }

        internal static string BuildFileUrl(
            string baseUrl,
            string userId,
            string relativePath)
        {
            string normalizedPath =
                FileLinkPath.NormalizeRelativePath(relativePath);
            string[] segments = normalizedPath.Split(
                new[] { '/' },
                StringSplitOptions.RemoveEmptyEntries);
            string encoded = string.Join(
                "/",
                segments.Select(Uri.EscapeDataString));
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/files/{1}/{2}",
                baseUrl.TrimEnd('/'),
                Uri.EscapeDataString(userId ?? string.Empty),
                encoded);
        }

        internal static string BuildChunkUploadFolderUrl(
            string baseUrl,
            string userId)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/uploads/{1}/{2}",
                baseUrl.TrimEnd('/'),
                Uri.EscapeDataString(userId ?? string.Empty),
                Uri.EscapeDataString(
                    "ncconnector-" + Guid.NewGuid().ToString("N")));
        }

        private static bool IsRetryableTransport(WebException exception)
        {
            if (exception == null)
            {
                return false;
            }

            return exception.Status == WebExceptionStatus.ConnectFailure
                   || exception.Status == WebExceptionStatus.ConnectionClosed
                   || exception.Status == WebExceptionStatus.KeepAliveFailure
                   || exception.Status == WebExceptionStatus.PipelineFailure
                   || exception.Status == WebExceptionStatus.ReceiveFailure
                   || exception.Status == WebExceptionStatus.SendFailure
                   || exception.Status == WebExceptionStatus.Timeout;
        }

        private static TimeSpan ResolveRetryDelay(
            NcHttpResponse response,
            int attempt)
        {
            string retryAfter = null;
            if (response != null && response.Headers != null)
            {
                response.Headers.TryGetValue("Retry-After", out retryAfter);
            }

            int seconds;
            if (!string.IsNullOrWhiteSpace(retryAfter)
                && int.TryParse(
                    retryAfter.Trim(),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out seconds))
            {
                return TimeSpan.FromSeconds(
                    Math.Max(0, Math.Min(30, seconds)));
            }

            DateTimeOffset retryDate;
            if (!string.IsNullOrWhiteSpace(retryAfter)
                && DateTimeOffset.TryParse(
                    retryAfter,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal,
                    out retryDate))
            {
                TimeSpan serverDelay = retryDate - DateTimeOffset.UtcNow;
                if (serverDelay < TimeSpan.Zero)
                {
                    serverDelay = TimeSpan.Zero;
                }
                if (serverDelay > TimeSpan.FromSeconds(30))
                {
                    serverDelay = TimeSpan.FromSeconds(30);
                }
                return serverDelay;
            }

            return TimeSpan.FromMilliseconds(500 * attempt);
        }

        private static void WaitForRetry(
            TimeSpan delay,
            CancellationToken cancellationToken)
        {
            if (delay <= TimeSpan.Zero)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return;
            }
            if (cancellationToken.WaitHandle.WaitOne(delay))
            {
                cancellationToken.ThrowIfCancellationRequested();
            }
        }
    }
}
