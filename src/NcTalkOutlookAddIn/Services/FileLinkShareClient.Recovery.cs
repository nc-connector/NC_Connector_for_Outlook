// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Resolves ambiguous create-share results without issuing a duplicate POST.
    internal sealed partial class FileLinkShareClient
    {
        private bool TryRecoverAmbiguousCreate(
            string sharesUrl,
            string relativeFolderPath,
            string shareKey,
            CancellationToken cancellationToken,
            out FileLinkShareData recoveredShare)
        {
            recoveredShare = null;
            _indeterminateSharePaths.Add(shareKey);
            cancellationToken.ThrowIfCancellationRequested();

            DiagnosticsLogger.Log(
                LogCategories.Api,
                "Share creation result is unknown; checking the exact path.");
            ShareLookupResult lookup = LookupExistingShare(
                sharesUrl,
                relativeFolderPath,
                cancellationToken);
            if (lookup.State == ShareLookupState.Found)
            {
                recoveredShare = RememberCreatedShare(
                    shareKey,
                    lookup.Share);
                DiagnosticsLogger.Log(
                    LogCategories.Api,
                    "Recovered the created share from the exact path.");
                return true;
            }
            if (lookup.State == ShareLookupState.Absent)
            {
                _indeterminateSharePaths.Remove(shareKey);
                return false;
            }

            DiagnosticsLogger.LogException(
                LogCategories.Api,
                "Could not verify the share creation result.",
                lookup.Error);
            return false;
        }

        private ShareLookupResult LookupExistingShare(
            string sharesUrl,
            string relativeFolderPath,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string requestedPath = "/" + relativeFolderPath;
            NcHttpResponse response = _sendRequest(
                new NcHttpRequestOptions
                {
                    Method = "GET",
                    Url = sharesUrl
                          + "?path="
                          + Uri.EscapeDataString(requestedPath)
                          + "&reshares=false&subfiles=false",
                    Accept = "application/json",
                    TimeoutMs = 60000,
                    ReadWriteTimeoutMs = 60000,
                    IncludeAuthHeader = true,
                    IncludeOcsApiHeader = true,
                    ParseJson = true,
                    CancellationToken = cancellationToken
                });

            cancellationToken.ThrowIfCancellationRequested();
            if (!response.HasHttpResponse)
            {
                return ShareLookupResult.Unknown(
                    CreateUnavailableException(
                        response.TransportException));
            }
            try
            {
                ValidateOcsResponse(response);
            }
            catch (TalkServiceException ex)
            {
                return ShareLookupResult.Unknown(ex);
            }

            IList<IDictionary<string, object>> shares =
                NcJson.GetOcsDataArray(response.ParsedJson);
            if (shares == null)
            {
                return ShareLookupResult.Unknown(
                    CreateUnavailableException(null));
            }

            foreach (IDictionary<string, object> share in shares)
            {
                int shareType;
                if (!NcJson.TryGetInt(
                        share,
                        "share_type",
                        out shareType)
                    && !NcJson.TryGetInt(
                        share,
                        "shareType",
                        out shareType))
                {
                    continue;
                }
                if (shareType != ShareTypePublicLink)
                {
                    continue;
                }

                string returnedPath =
                    NcJson.GetTrimmedString(share, "path");
                if (!string.IsNullOrWhiteSpace(returnedPath)
                    && !string.Equals(
                        "/" + FileLinkPath.NormalizeRelativePath(
                            returnedPath),
                        requestedPath,
                        StringComparison.Ordinal))
                {
                    continue;
                }

                var result = new FileLinkShareData(
                    NcJson.GetTrimmedString(share, "id"),
                    NcJson.GetTrimmedString(share, "url"),
                    NcJson.GetTrimmedString(share, "token"));
                if (!string.IsNullOrWhiteSpace(result.Id)
                    && !string.IsNullOrWhiteSpace(result.Url))
                {
                    return ShareLookupResult.Found(result);
                }
            }
            return ShareLookupResult.Absent();
        }

        private FileLinkShareData RememberCreatedShare(
            string shareKey,
            FileLinkShareData share)
        {
            _createdShares[shareKey] = share;
            _indeterminateSharePaths.Remove(shareKey);
            return share;
        }

        private static TalkServiceException CreateUnavailableException(
            Exception transportException)
        {
            return new TalkServiceException(
                transportException != null
                    ? transportException.Message
                    : Strings.ErrorServerUnavailable,
                false,
                0,
                null);
        }

        private enum ShareLookupState
        {
            Found,
            Absent,
            Unknown
        }

        private sealed class ShareLookupResult
        {
            private ShareLookupResult(
                ShareLookupState state,
                FileLinkShareData share,
                TalkServiceException error)
            {
                State = state;
                Share = share;
                Error = error;
            }

            internal ShareLookupState State { get; private set; }

            internal FileLinkShareData Share { get; private set; }

            internal TalkServiceException Error { get; private set; }

            internal static ShareLookupResult Found(
                FileLinkShareData share)
            {
                return new ShareLookupResult(
                    ShareLookupState.Found,
                    share,
                    null);
            }

            internal static ShareLookupResult Absent()
            {
                return new ShareLookupResult(
                    ShareLookupState.Absent,
                    null,
                    null);
            }

            internal static ShareLookupResult Unknown(
                TalkServiceException error)
            {
                return new ShareLookupResult(
                    ShareLookupState.Unknown,
                    null,
                    error);
            }
        }
    }
}
