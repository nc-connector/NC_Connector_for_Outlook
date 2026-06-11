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
    internal sealed class SecretsService
    {
        private const string CreateSecretPath = "/ocs/v2.php/apps/secrets/api/v1/secrets";

        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;

        internal SecretsService(TalkServiceConfiguration configuration)
        {
            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        internal SecretsLinkResult CreateSecretLink(string plainText, string title, int expireDays)
        {
            if (_configuration == null || !_configuration.IsComplete())
            {
                throw new TalkServiceException("Secrets link creation failed: credentials are incomplete.", false, 0, null);
            }
            if (string.IsNullOrWhiteSpace(plainText))
            {
                throw new ArgumentException("Secret content is required.", "plainText");
            }

            string baseUrl = _configuration.GetNormalizedBaseUrl();
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                throw new TalkServiceException("Secrets link creation failed: base URL is invalid.", false, 0, null);
            }

            int normalizedExpireDays = SharePasswordDeliveryPolicy.ClampSecretsExpireDays(expireDays);
            SecretsEncryptedPayload encrypted = SecretsCrypto.EncryptToSecretsPayload(plainText);
            string url = baseUrl.TrimEnd('/') + CreateSecretPath;
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Secrets password link create prepared (expireDays="
                + normalizedExpireDays.ToString(CultureInfo.InvariantCulture)
                + ", hasTitle="
                + (!string.IsNullOrWhiteSpace(title)).ToString(CultureInfo.InvariantCulture)
                + ").");
            var payload = new Dictionary<string, object>
            {
                { "title", string.IsNullOrWhiteSpace(title) ? "NC Connector share password" : title.Trim() },
                { "encrypted", encrypted.Encrypted },
                { "iv", encrypted.Iv },
                { "expires", DateTime.UtcNow.AddDays(normalizedExpireDays).ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture) }
            };

            DiagnosticsLogger.LogApi("POST " + url);
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "POST",
                Url = url,
                Payload = NcJson.Serialize(payload),
                Accept = "application/json",
                ContentType = "application/json",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {
                Exception transport = response.TransportException;
                DiagnosticsLogger.LogException(LogCategories.Api, "Secrets create failed without HTTP response.", transport);
                throw new TalkServiceException(
                    "Secrets link creation failed: " + (transport != null ? transport.Message : "no HTTP response"),
                    false,
                    0,
                    null,
                    true);
            }

            DiagnosticsLogger.LogApi("POST " + url + " -> " + response.StatusCode);
            if (response.StatusCode != HttpStatusCode.Created)
            {
                string detail = NcJson.ExtractOcsErrorMessage(response.ParsedJson);
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "HTTP " + (int)response.StatusCode;
                }
                bool authError = response.StatusCode == HttpStatusCode.Unauthorized
                                 || response.StatusCode == HttpStatusCode.Forbidden;
                throw new TalkServiceException(
                    "Secrets link creation failed: " + detail,
                    authError,
                    response.StatusCode,
                    response.ResponseText);
            }

            IDictionary<string, object> data = NcJson.GetOcsData(response.ParsedJson);
            string uuid = NcJson.GetTrimmedString(data, "uuid");
            if (string.IsNullOrWhiteSpace(uuid))
            {
                uuid = NcJson.GetTrimmedString(data, "id");
            }
            if (string.IsNullOrWhiteSpace(uuid))
            {
                throw new TalkServiceException("Secrets link creation failed: response did not contain a UUID.", false, response.StatusCode, response.ResponseText);
            }

            DateTime expires;
            DateTime? parsedExpires = null;
            string expiresText = NcJson.GetTrimmedString(data, "expires");
            if (!string.IsNullOrWhiteSpace(expiresText)
                && DateTime.TryParse(expiresText, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out expires))
            {
                parsedExpires = expires;
            }

            string shareUrl = baseUrl.TrimEnd('/')
                              + "/index.php/apps/secrets/share/"
                              + Uri.EscapeDataString(uuid)
                              + "#"
                              + encrypted.Key;
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Secrets password link create succeeded (hasUuid=True, hasExpires="
                + parsedExpires.HasValue.ToString(CultureInfo.InvariantCulture)
                + ").");
            return new SecretsLinkResult(uuid, shareUrl, parsedExpires);
        }
    }
}
