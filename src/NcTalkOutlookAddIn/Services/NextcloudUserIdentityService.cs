// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Net;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Resolves the authenticated account's canonical Nextcloud UID for user-scoped DAV routes.
    internal static class NextcloudUserIdentityService
    {
        private static readonly object CacheSync = new object();
        private static readonly Dictionary<string, string> CurrentUserIdCache =
            new Dictionary<string, string>(StringComparer.Ordinal);

        internal static string ResolveCurrentUserId(TalkServiceConfiguration configuration, bool forceRefresh = false)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            if (!configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud credentials are incomplete.");
            }

            string baseUrl = configuration.GetNormalizedBaseUrl();
            string cacheKey = baseUrl + "\n" + (configuration.Username ?? string.Empty).Trim();
            if (!forceRefresh)
            {
                lock (CacheSync)
                {
                    string cachedUserId;
                    if (CurrentUserIdCache.TryGetValue(cacheKey, out cachedUserId))
                    {
                        return cachedUserId;
                    }
                }
            }

            string url = baseUrl + "/ocs/v2.php/cloud/user?format=json";
            DiagnosticsLogger.LogApi("GET " + url);
            NcHttpResponse response = new NcHttpClient(configuration).Send(new NcHttpRequestOptions
            {
                Method = "GET",
                Url = url,
                Accept = "application/json",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "Current Nextcloud user ID request failed.", transport);
                throw new TalkServiceException(
                    "Current Nextcloud user ID could not be resolved: " + (transport != null ? transport.Message : "no HTTP response"),
                    false,
                    0,
                    null,
                    true);
            }

            DiagnosticsLogger.LogApi("GET " + url + " -> " + response.StatusCode);
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                string detail = NcJson.ExtractOcsErrorMessage(response.ParsedJson);
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "HTTP " + (int)response.StatusCode;
                }
                bool authenticationError = response.StatusCode == HttpStatusCode.Unauthorized
                                           || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException(
                    "Current Nextcloud user ID could not be resolved: " + detail,
                    authenticationError,
                    response.StatusCode,
                    response.ResponseText);
            }

            string userId = ExtractCurrentUserId(response.ParsedJson);
            if (string.IsNullOrWhiteSpace(userId))
            {
                // Authentication aliases such as email addresses are not valid substitutes in DAV paths.
                throw new TalkServiceException(
                    "Nextcloud did not return a canonical user ID.",
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }

            lock (CacheSync)
            {
                CurrentUserIdCache[cacheKey] = userId;
            }
            DiagnosticsLogger.Log(LogCategories.Api, "Canonical Nextcloud user ID resolved.");
            return userId;
        }

        internal static string ExtractCurrentUserId(IDictionary<string, object> payload)
        {
            return NcJson.GetStringOrEmpty(NcJson.GetOcsData(payload), "id");
        }
    }
}
