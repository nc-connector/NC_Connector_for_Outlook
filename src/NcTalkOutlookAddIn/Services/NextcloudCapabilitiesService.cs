// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Fetches and caches the OCS capabilities shared by connection and upload flows.
    internal sealed class NextcloudCapabilitiesService
    {
        private static readonly object CacheSync = new object();
        private static readonly Dictionary<string, CacheSlot> Cache =
            new Dictionary<string, CacheSlot>(StringComparer.Ordinal);
        private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(5);

        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;

        internal NextcloudCapabilitiesService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        internal NextcloudCapabilitiesSnapshot GetSnapshot(bool forceRefresh, bool forceFreshConnection)
        {
            if (!_configuration.IsComplete())
            {
                throw new TalkServiceException(Strings.ErrorMissingCredentials, true, 0, null);
            }

            string cacheKey = BuildCacheKey(_configuration);
            CacheSlot slot;
            lock (CacheSync)
            {
                RemoveExpiredSlots(DateTime.UtcNow);
                if (!Cache.TryGetValue(cacheKey, out slot))
                {
                    slot = new CacheSlot();
                    Cache[cacheKey] = slot;
                }
                slot.ActiveUsers++;
            }

            try
            {
                lock (slot.SyncRoot)
                {
                    if (!forceRefresh
                        && slot.Snapshot != null
                        && DateTime.UtcNow - slot.FetchedAtUtc
                        <= CacheLifetime)
                    {
                        return slot.Snapshot;
                    }

                    NextcloudCapabilitiesSnapshot snapshot = FetchSnapshot(
                        forceFreshConnection);
                    slot.Snapshot = snapshot;
                    slot.FetchedAtUtc = DateTime.UtcNow;
                    return snapshot;
                }
            }
            finally
            {
                lock (CacheSync)
                {
                    slot.ActiveUsers--;
                    RemoveExpiredSlots(DateTime.UtcNow);
                }
            }
        }

        internal NextcloudCapabilitiesSnapshot GetRequiredSnapshot(bool forceRefresh, bool forceFreshConnection)
        {
            NextcloudCapabilitiesSnapshot snapshot = GetSnapshot(forceRefresh, forceFreshConnection);
            RequireSupportedSnapshot(snapshot);
            return snapshot;
        }

        internal static void RequireSupportedSnapshot(
            NextcloudCapabilitiesSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }
            if (!NextcloudVersionHelper.IsSupported(snapshot.ServerVersion))
            {
                string detected = string.IsNullOrWhiteSpace(snapshot.ServerVersionText)
                    ? Strings.TalkVersionUnknown
                    : snapshot.ServerVersionText;
                string message = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.NextcloudMinimumVersionRequiredFormat,
                    detected);
                throw new TalkServiceException(message, false, 0, null);
            }
        }

        private NextcloudCapabilitiesSnapshot FetchSnapshot(bool forceFreshConnection)
        {
            string url = _configuration.GetNormalizedBaseUrl()
                         + "/ocs/v2.php/cloud/capabilities?format=json";
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "GET",
                Url = url,
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true,
                ForceFreshConnection = forceFreshConnection
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                throw new TalkServiceException(
                    transport != null ? transport.Message : Strings.ErrorServerUnavailable,
                    false,
                    0,
                    null);
            }
            if (response.StatusCode == HttpStatusCode.Unauthorized
                || response.StatusCode == HttpStatusCode.Forbidden)
            {
                throw new TalkServiceException(
                    Strings.ErrorCredentialsNotVerified,
                    true,
                    response.StatusCode,
                    response.ResponseText);
            }
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                throw new TalkServiceException(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.ErrorConnectionFailed,
                        "HTTP "
                        + ((int)response.StatusCode).ToString(
                            CultureInfo.InvariantCulture)),
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }
            if (response.ParsedJson == null)
            {
                throw new TalkServiceException(
                    Strings.ErrorCredentialsNotVerified,
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }

            ValidateOcsStatus(
                response.ParsedJson,
                response.StatusCode,
                response.ResponseText);

            Version serverVersion;
            string versionText;
            NextcloudVersionHelper.TryExtractFromCapabilities(
                response.ParsedJson,
                out serverVersion,
                out versionText);

            IDictionary<string, object> data = NcJson.GetOcsData(response.ParsedJson);
            IDictionary<string, object> capabilities = NcJson.GetDictionary(data, "capabilities");
            if (capabilities == null)
            {
                throw new TalkServiceException(
                    Strings.ErrorCredentialsNotVerified,
                    false,
                    response.StatusCode,
                    response.ResponseText);
            }

            IDictionary<string, object> dav = NcJson.GetDictionary(capabilities, "dav");
            bool bulkUploadSupported = string.Equals(
                NcJson.GetTrimmedString(dav, "bulkupload"),
                "1.0",
                StringComparison.Ordinal);

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Nextcloud capabilities loaded (version="
                + (string.IsNullOrWhiteSpace(versionText) ? "unknown" : versionText)
                + ", davBulkUpload="
                + bulkUploadSupported
                + ").");

            return new NextcloudCapabilitiesSnapshot(
                serverVersion,
                versionText,
                capabilities,
                bulkUploadSupported);
        }

        private static void ValidateOcsStatus(
            IDictionary<string, object> root,
            HttpStatusCode statusCode,
            string responseText)
        {
            string detail;
            if (!NcJson.IsOcsSuccess(root, out detail))
            {
                throw new TalkServiceException(
                    string.IsNullOrWhiteSpace(detail)
                        ? Strings.ErrorCredentialsNotVerified
                        : detail,
                    false,
                    statusCode,
                    responseText);
            }
        }

        private static string BuildCacheKey(TalkServiceConfiguration configuration)
        {
            return configuration.GetNormalizedBaseUrl().Trim()
                   + "\n"
                   + (configuration.Username ?? string.Empty).Trim();
        }

        private static void RemoveExpiredSlots(DateTime nowUtc)
        {
            var expiredKeys = new List<string>();
            foreach (KeyValuePair<string, CacheSlot> pair in Cache)
            {
                CacheSlot slot = pair.Value;
                if (slot != null
                    && slot.ActiveUsers == 0
                    && (slot.Snapshot == null
                        || nowUtc - slot.FetchedAtUtc > CacheLifetime))
                {
                    expiredKeys.Add(pair.Key);
                }
            }
            foreach (string key in expiredKeys)
            {
                Cache.Remove(key);
            }
        }

        private sealed class CacheSlot
        {
            internal CacheSlot()
            {
                SyncRoot = new object();
            }

            internal object SyncRoot { get; private set; }

            internal int ActiveUsers { get; set; }

            internal NextcloudCapabilitiesSnapshot Snapshot { get; set; }

            internal DateTime FetchedAtUtc { get; set; }
        }
    }
}
