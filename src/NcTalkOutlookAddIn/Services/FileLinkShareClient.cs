// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Creates public shares in one OCS request and validates the OCS result.
    internal sealed partial class FileLinkShareClient
    {
        private const int ShareTypePublicLink = 3;

        private readonly Func<NcHttpRequestOptions, NcHttpResponse>
            _sendRequest;
        private readonly object _createSync = new object();
        private readonly IDictionary<string, FileLinkShareData>
            _createdShares =
                new Dictionary<string, FileLinkShareData>(
                    StringComparer.Ordinal);
        private readonly ISet<string> _indeterminateSharePaths =
            new HashSet<string>(StringComparer.Ordinal);

        internal FileLinkShareClient(NcHttpClient httpClient)
            : this(
                httpClient == null
                    ? null
                    : new Func<NcHttpRequestOptions, NcHttpResponse>(
                        httpClient.Send))
        {
        }

        internal FileLinkShareClient(
            Func<NcHttpRequestOptions, NcHttpResponse> sendRequest)
        {
            if (sendRequest == null)
            {
                throw new ArgumentNullException("sendRequest");
            }

            _sendRequest = sendRequest;
        }

        internal FileLinkShareData Create(
            string baseUrl,
            string relativeFolderPath,
            string shareName,
            FileLinkRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }

            string normalizedBaseUrl = NormalizeBaseUrl(baseUrl);
            string normalizedFolderPath =
                FileLinkPath.NormalizeRelativePath(relativeFolderPath);
            if (string.IsNullOrWhiteSpace(normalizedFolderPath))
            {
                throw new ArgumentException(
                    "A relative share folder path is required.",
                    "relativeFolderPath");
            }

            string url = normalizedBaseUrl
                         + "/ocs/v2.php/apps/files_sharing/api/v1/shares";
            string shareKey = normalizedBaseUrl
                              + "\n"
                              + normalizedFolderPath;

            EnterCreateGate(cancellationToken);
            try
            {
                FileLinkShareData knownShare;
                if (_createdShares.TryGetValue(
                        shareKey,
                        out knownShare))
                {
                    return knownShare;
                }

                if (_indeterminateSharePaths.Contains(shareKey))
                {
                    ShareLookupResult existing = LookupExistingShare(
                        url,
                        normalizedFolderPath,
                        cancellationToken);
                    if (existing.State == ShareLookupState.Found)
                    {
                        return RememberCreatedShare(
                            shareKey,
                            existing.Share);
                    }
                    if (existing.State == ShareLookupState.Unknown)
                    {
                        throw existing.Error
                              ?? CreateUnavailableException(null);
                    }
                    _indeterminateSharePaths.Remove(shareKey);
                }

                string payload = BuildCreatePayload(
                    normalizedFolderPath,
                    shareName,
                    request);
                NcHttpResponse response = SendCreateRequest(
                    url,
                    payload,
                    cancellationToken);
                if (!response.HasHttpResponse)
                {
                    TalkServiceException transportError =
                        CreateUnavailableException(
                            response.TransportException);
                    FileLinkShareData recovered;
                    if (TryRecoverAmbiguousCreate(
                            url,
                            normalizedFolderPath,
                            shareKey,
                            cancellationToken,
                            out recovered))
                    {
                        return recovered;
                    }
                    throw transportError;
                }

                bool validated = false;
                try
                {
                    ValidateOcsResponse(response);
                    validated = true;
                    return RememberCreatedShare(
                        shareKey,
                        ParseShareData(
                            response.ParsedJson,
                            response.ResponseText));
                }
                catch (TalkServiceException)
                {
                    bool ambiguousHttpResult =
                        IsAmbiguousHttpCreateResult(response);
                    if (!validated
                        && !ambiguousHttpResult
                        && ((int)response.StatusCode < 200
                            || (int)response.StatusCode >= 300
                            || HasExplicitOcsResult(
                                response.ParsedJson)))
                    {
                        throw;
                    }

                    FileLinkShareData recovered;
                    if (TryRecoverAmbiguousCreate(
                            url,
                            normalizedFolderPath,
                            shareKey,
                            cancellationToken,
                            out recovered))
                    {
                        return recovered;
                    }
                    throw;
                }
            }
            finally
            {
                Monitor.Exit(_createSync);
            }
        }

        private void EnterCreateGate(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            while (!Monitor.TryEnter(_createSync, 100))
            {
                cancellationToken.ThrowIfCancellationRequested();
            }
        }

        private NcHttpResponse SendCreateRequest(
            string url,
            string formPayload,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            NcHttpResponse response = _sendRequest(
                new NcHttpRequestOptions
                {
                    Method = "POST",
                    Url = url,
                    Payload = formPayload ?? string.Empty,
                    Accept = "application/json",
                    ContentType = "application/x-www-form-urlencoded",
                    TimeoutMs = 90000,
                    ReadWriteTimeoutMs = 90000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = true,
                    ParseJson = true,
                    CancellationToken = cancellationToken
                });

            return response;
        }

        private static void ValidateOcsResponse(
            NcHttpResponse response)
        {
            if ((int)response.StatusCode < 200
                || (int)response.StatusCode >= 300)
            {
                string detail = NcJson.ExtractOcsErrorMessage(
                    response.ParsedJson);
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "HTTP "
                             + ((int)response.StatusCode).ToString(
                                 CultureInfo.InvariantCulture);
                }
                bool authError =
                    response.StatusCode == HttpStatusCode.Unauthorized
                    || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException(
                    detail,
                    authError,
                    response.StatusCode,
                    response.ResponseText);
            }

            string ocsDetail = string.Empty;
            if (response.ParsedJson == null
                || !NcJson.IsOcsSuccess(
                    response.ParsedJson,
                    out ocsDetail))
            {
                throw new TalkServiceException(
                    string.IsNullOrWhiteSpace(ocsDetail)
                        ? Strings.ErrorServerUnavailable
                        : ocsDetail,
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }
        }

        private static bool HasExplicitOcsResult(
            IDictionary<string, object> parsedJson)
        {
            IDictionary<string, object> meta =
                NcJson.GetOcsMeta(parsedJson);
            if (meta == null)
            {
                return false;
            }

            int statusCode;
            if (NcJson.TryGetInt(meta, "statuscode", out statusCode))
            {
                return true;
            }
            return !string.IsNullOrWhiteSpace(
                NcJson.GetTrimmedString(meta, "status"));
        }

        private static bool IsAmbiguousHttpCreateResult(
            NcHttpResponse response)
        {
            if (response == null
                || !response.HasHttpResponse
                || HasExplicitOcsResult(response.ParsedJson))
            {
                return false;
            }

            return FileLinkUploadPolicy.IsIndeterminateStatusCode(
                (int)response.StatusCode);
        }

        private static string NormalizeBaseUrl(string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                throw new ArgumentException(
                    "A Nextcloud base URL is required.",
                    "baseUrl");
            }
            return baseUrl.Trim().TrimEnd('/');
        }

        private static string BuildCreatePayload(
            string relativeFolderPath,
            string shareName,
            FileLinkRequest request)
        {
            var builder = new StringBuilder();
            builder.Append("path=")
                .Append(Uri.EscapeDataString("/" + relativeFolderPath));
            builder.Append("&shareType=")
                .Append(
                    ShareTypePublicLink.ToString(
                        CultureInfo.InvariantCulture));
            builder.Append("&permissions=")
                .Append(
                    CalculatePermissionValue(request.Permissions).ToString(
                        CultureInfo.InvariantCulture));

            if (request.PasswordEnabled
                && !string.IsNullOrEmpty(request.Password))
            {
                builder.Append("&password=")
                    .Append(Uri.EscapeDataString(request.Password));
            }
            if (request.ExpireEnabled && request.ExpireDate.HasValue)
            {
                builder.Append("&expireDate=")
                    .Append(
                        Uri.EscapeDataString(
                            request.ExpireDate.Value.ToString(
                                "yyyy-MM-dd",
                                CultureInfo.InvariantCulture)));
            }
            if (!string.IsNullOrWhiteSpace(shareName))
            {
                builder.Append("&label=")
                    .Append(Uri.EscapeDataString(shareName));
            }
            if (request.NoteEnabled)
            {
                builder.Append("&note=")
                    .Append(
                        Uri.EscapeDataString(
                            request.Note ?? string.Empty));
            }
            return builder.ToString();
        }

        private static FileLinkShareData ParseShareData(
            IDictionary<string, object> parsedJson,
            string responseText)
        {
            IDictionary<string, object> data = NcJson.GetOcsData(parsedJson);
            var result = new FileLinkShareData(
                NcJson.GetTrimmedString(data, "id"),
                NcJson.GetTrimmedString(data, "url"),
                NcJson.GetTrimmedString(data, "token"));
            if (string.IsNullOrWhiteSpace(result.Id)
                || string.IsNullOrWhiteSpace(result.Url))
            {
                throw new TalkServiceException(
                    Strings.ErrorServerUnavailable,
                    false,
                    0,
                    responseText);
            }
            return result;
        }

        private static int CalculatePermissionValue(
            FileLinkPermissionFlags permissions)
        {
            int value = 0;
            if ((permissions & FileLinkPermissionFlags.Read)
                == FileLinkPermissionFlags.Read)
            {
                value |= 1;
            }
            if ((permissions & FileLinkPermissionFlags.Write)
                == FileLinkPermissionFlags.Write)
            {
                value |= 2;
            }
            if ((permissions & FileLinkPermissionFlags.Create)
                == FileLinkPermissionFlags.Create)
            {
                value |= 4;
            }
            if ((permissions & FileLinkPermissionFlags.Delete)
                == FileLinkPermissionFlags.Delete)
            {
                value |= 8;
            }
            return value;
        }

    }

    internal sealed class FileLinkShareData
    {
        internal FileLinkShareData(
            string id,
            string url,
            string token)
        {
            Id = id ?? string.Empty;
            Url = url ?? string.Empty;
            Token = token ?? string.Empty;
        }

        internal string Id { get; private set; }

        internal string Url { get; private set; }

        internal string Token { get; private set; }
    }
}
