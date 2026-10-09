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
        // Loads centralized NC Connector policy status from Nextcloud backend endpoint.
    internal sealed class BackendPolicyService
    {
        private const string StatusEndpointPath = "/apps/ncc_backend_4mc/api/v1/status";
        private readonly TalkServiceConfiguration _configuration;
        private readonly NcHttpClient _httpClient;

        internal BackendPolicyService(TalkServiceConfiguration configuration)
        {
            _configuration = configuration;
            _httpClient = new NcHttpClient(configuration);
        }

        internal BackendPolicyStatus FetchStatus()
        {
            if (_configuration == null || !_configuration.IsComplete())
            {
                return BuildLocalStatus(
                    endpointAvailable: false,
                    fetchSucceeded: false,
                    reason: "credentials_incomplete");
            }
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                return BuildLocalStatus(
                    endpointAvailable: false,
                    fetchSucceeded: false,
                    reason: "base_url_invalid");
            }
            string endpointUrl = baseUrl.TrimEnd('/') + StatusEndpointPath;

            NcHttpResponse response = ExecuteJsonRequest(endpointUrl);
            IDictionary<string, object> payload = response.ParsedJson;
            HttpStatusCode statusCode = response.StatusCode;
            bool httpOk = response.HasHttpResponse && (int)statusCode >= 200 && (int)statusCode < 300;

            // Some installations return this endpoint's valid JSON with HTTP 404.
            IDictionary<string, object> normalized = NormalizePayload(payload);
            if (httpOk || (statusCode == HttpStatusCode.NotFound
                && normalized != null && normalized.ContainsKey("status")))
            {
                return ParseStatus(payload);
            }

            if (statusCode == HttpStatusCode.NotFound)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Policy status endpoint missing: " + endpointUrl, null);
                return BuildLocalStatus(
                    endpointAvailable: false,
                    fetchSucceeded: false,
                    reason: "backend_unavailable");
            }

            DiagnosticsLogger.LogException(LogCategories.Core, "Policy status endpoint unavailable (status=" + (int)statusCode + ").", null);
            string failureReason = !response.HasHttpResponse
                ? "nextcloud_unavailable"
                : (statusCode == HttpStatusCode.Unauthorized
                    ? "authentication_rejected"
                    : ((int)statusCode == 429 ? "rate_limited"
                        : ((int)statusCode == 500 || (int)statusCode == 502
                           || (int)statusCode == 503 || (int)statusCode == 504
                            ? "backend_unavailable" : "check_failed")));
            BackendPolicyStatus failure = BuildLocalStatus(
                endpointAvailable: true,
                fetchSucceeded: false,
                reason: failureReason);
            if (failureReason == "rate_limited")
            {
                failure.RetryAfterUtc = ReadRetryAfterUtc(response);
            }
            return failure;
        }

        internal static BackendPolicyStatus ParseStatus(IDictionary<string, object> payload)
        {
            IDictionary<string, object> normalized = NormalizePayload(payload);
            IDictionary<string, object> status = NcJson.GetDictionary(normalized, "status");
            object rawSeatAssigned;
            object rawIsValid;
            object rawSeatState;
            if (status == null
                || !status.TryGetValue("seat_assigned", out rawSeatAssigned) || !(rawSeatAssigned is bool)
                || !status.TryGetValue("is_valid", out rawIsValid) || !(rawIsValid is bool)
                || !status.TryGetValue("seat_state", out rawSeatState) || !(rawSeatState is string)
                || string.IsNullOrWhiteSpace((string)rawSeatState))
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Policy status response has missing or invalid personal status fields.", null);
                return BuildLocalStatus(
                    endpointAvailable: true,
                    fetchSucceeded: false,
                    reason: "invalid_payload");
            }
            IDictionary<string, object> licenseActivation = NcJson.GetDictionary(status, "license_activation");
            IDictionary<string, object> policy = NcJson.GetDictionary(normalized, "policy");
            IDictionary<string, object> policyEditable = NcJson.GetDictionary(normalized, "policy_editable");
            IDictionary<string, object> sharePolicy = NcJson.GetDictionary(policy, "share");
            IDictionary<string, object> talkPolicy = NcJson.GetDictionary(policy, "talk");
            IDictionary<string, object> emailSignaturePolicy = NcJson.GetDictionary(policy, "email_signature");
            IDictionary<string, object> shareEditable = NcJson.GetDictionary(policyEditable, "share");
            IDictionary<string, object> talkEditable = NcJson.GetDictionary(policyEditable, "talk");
            IDictionary<string, object> emailSignatureEditable = NcJson.GetDictionary(policyEditable, "email_signature");

            object rawExpireDays;
            int expireDays;
            if (sharePolicy != null
                && sharePolicy.TryGetValue("share_expire_days", out rawExpireDays)
                && BackendPolicyStatus.TryConvertInt(rawExpireDays, out expireDays)
                && expireDays == 0)
            {
                // Backends before 1.4.2 could return zero; new shares require at least one day.
                sharePolicy = new Dictionary<string, object>(sharePolicy);
                sharePolicy["share_expire_days"] = 1;
            }

            bool seatAssigned = (bool)rawSeatAssigned;
            bool isValid = (bool)rawIsValid;
            string seatState = ((string)rawSeatState).Trim();
            object rawCanManageLicense;
            bool canManageLicense = status != null
                                    && status.TryGetValue("can_manage_license", out rawCanManageLicense)
                                    && rawCanManageLicense is bool
                                    && (bool)rawCanManageLicense;
            object rawDefaultsSourceEditable;
            bool defaultsSourceEditable = normalized.TryGetValue("defaults_source_editable", out rawDefaultsSourceEditable)
                                          && rawDefaultsSourceEditable is bool
                                          && (bool)rawDefaultsSourceEditable;

            bool seatUsable = seatAssigned
                              && isValid
                              && string.Equals(seatState, "active", StringComparison.OrdinalIgnoreCase);
            bool sharePolicyActive = seatUsable && sharePolicy != null && shareEditable != null;
            bool talkPolicyActive = seatUsable && talkPolicy != null && talkEditable != null;
            bool emailSignaturePolicyActive = seatUsable && emailSignaturePolicy != null && emailSignatureEditable != null;
            bool policyActive = sharePolicyActive || talkPolicyActive || emailSignaturePolicyActive;

            BackendPolicyStatus normalizedStatus = new BackendPolicyStatus(
                endpointAvailable: true,
                fetchSucceeded: true,
                policyActive: policyActive,
                mode: policyActive ? "policy" : "local",
                reason: policyActive ? "policy_active" : (seatUsable ? "policy_domains_unavailable" : "seat_not_usable"),
                seatAssigned: seatAssigned,
                isValid: isValid,
                seatState: seatState,
                sharePolicy: sharePolicy,
                talkPolicy: talkPolicy,
                emailSignaturePolicy: emailSignaturePolicy,
                shareEditable: shareEditable,
                talkEditable: talkEditable,
                emailSignatureEditable: emailSignatureEditable,
                licenseStatus: NcJson.GetStringOrEmpty(status, "license_status"),
                accessStatus: NcJson.GetStringOrEmpty(status, "access_status"),
                canManageLicense: canManageLicense,
                graceUntilIso: NcJson.GetStringOrEmpty(status, "grace_until_iso"),
                licenseActivationState: NcJson.GetStringOrEmpty(licenseActivation, "state"),
                licenseConnectionError: GetBool(status, "license_connection_error"),
                licenseLastSyncAtIso: NcJson.GetStringOrEmpty(status, "license_last_sync_at_iso"),
                licenseOfflineUntilIso: NcJson.GetStringOrEmpty(status, "license_offline_until_iso"),
                defaultsSource: NcJson.GetStringOrEmpty(normalized, "defaults_source"),
                defaultsSourceEditable: defaultsSourceEditable);
            return normalizedStatus;
        }

        private static BackendPolicyStatus BuildLocalStatus(bool endpointAvailable, bool fetchSucceeded, string reason)
        {
            return new BackendPolicyStatus(
                endpointAvailable: endpointAvailable,
                fetchSucceeded: fetchSucceeded,
                policyActive: false,
                mode: "local",
                reason: reason,
                seatAssigned: false,
                isValid: false,
                seatState: string.Empty,
                sharePolicy: null,
                talkPolicy: null,
                emailSignaturePolicy: null,
                shareEditable: null,
                talkEditable: null,
                emailSignatureEditable: null);
        }

        private NcHttpResponse ExecuteJsonRequest(string url)
        {
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = "GET",
                Url = url,
                TimeoutMs = 45000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = true,
                ParseJson = true
            });

            if (!response.HasHttpResponse)
            {
                if (response.TransportException != null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed without HTTP response.", response.TransportException);
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed without HTTP response.", null);
                }
            }
            return response;
        }

        private static DateTime ReadRetryAfterUtc(NcHttpResponse response)
        {
            string value;
            if (response.Headers != null && response.Headers.TryGetValue("Retry-After", out value))
            {
                int seconds;
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out seconds)
                    && seconds >= 0)
                {
                    return DateTime.UtcNow.AddSeconds(seconds);
                }
                DateTimeOffset retryAt;
                if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal, out retryAt))
                {
                    return retryAt.UtcDateTime;
                }
            }
            return DateTime.UtcNow.AddMinutes(1);
        }

        private static IDictionary<string, object> NormalizePayload(IDictionary<string, object> payload)
        {
            if (payload == null)
            {
                return null;
            }

            IDictionary<string, object> ocs = NcJson.GetDictionary(payload, "ocs");
            IDictionary<string, object> data = NcJson.GetDictionary(ocs, "data");
            return data ?? payload;
        }

        private static bool GetBool(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            object raw;
            if (!parent.TryGetValue(key, out raw) || raw == null)
            {
                return false;
            }
            bool value;
            return BackendPolicyStatus.TryConvertBool(raw, out value) && value;
        }
    }
}

