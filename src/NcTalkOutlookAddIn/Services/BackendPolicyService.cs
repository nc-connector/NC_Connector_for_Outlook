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
        {            if (_configuration == null || !_configuration.IsComplete())
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

            IDictionary<string, object> payload;
            HttpStatusCode statusCode;
            bool httpOk = ExecuteJsonRequest(endpointUrl, out statusCode, out payload);

            if (!httpOk)
            {
                if (statusCode == HttpStatusCode.NotFound)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Policy status endpoint missing: " + endpointUrl, null);
                    return BuildLocalStatus(
                        endpointAvailable: false,
                        fetchSucceeded: true,
                        reason: "endpoint_missing");
                }

                DiagnosticsLogger.LogException(LogCategories.Core, "Policy status endpoint unavailable (status=" + (int)statusCode + ").", null);
                return BuildLocalStatus(
                    endpointAvailable: true,
                    fetchSucceeded: false,
                    reason: "endpoint_unavailable");
            }

            IDictionary<string, object> normalized = NormalizePayload(payload);
            IDictionary<string, object> status = NcJson.GetDictionary(normalized, "status");
            IDictionary<string, object> policy = NcJson.GetDictionary(normalized, "policy");
            IDictionary<string, object> policyEditable = NcJson.GetDictionary(normalized, "policy_editable");
            IDictionary<string, object> sharePolicy = NcJson.GetDictionary(policy, "share");
            IDictionary<string, object> talkPolicy = NcJson.GetDictionary(policy, "talk");
            IDictionary<string, object> shareEditable = NcJson.GetDictionary(policyEditable, "share");
            IDictionary<string, object> talkEditable = NcJson.GetDictionary(policyEditable, "talk");

            bool seatAssigned = GetBool(status, "seat_assigned");
            bool isValid = GetBool(status, "is_valid");
            string seatState = NcJson.GetStringOrEmpty(status, "seat_state");

            bool seatUsable = seatAssigned
                              && isValid
                              && string.Equals(seatState, "active", StringComparison.OrdinalIgnoreCase);
            bool policyActive = seatUsable
                                && sharePolicy != null
                                && talkPolicy != null
                                && shareEditable != null
                                && talkEditable != null;
            bool warningVisible = !policyActive && ShouldWarnForSeat(status);
            string warningMessage = warningVisible ? BuildSeatWarningMessage(seatAssigned, isValid, seatState) : string.Empty;

            BackendPolicyStatus normalizedStatus = new BackendPolicyStatus(
                endpointAvailable: true,
                fetchSucceeded: true,
                policyActive: policyActive,
                warningVisible: warningVisible,
                warningMessage: warningMessage,
                mode: policyActive ? "policy" : "local",
                reason: policyActive ? "policy_active" : "seat_not_usable",
                seatAssigned: seatAssigned,
                isValid: isValid,
                seatState: seatState,
                sharePolicy: sharePolicy,
                talkPolicy: talkPolicy,
                shareEditable: shareEditable,
                talkEditable: talkEditable);
            return normalizedStatus;
        }

        private static BackendPolicyStatus BuildLocalStatus(bool endpointAvailable, bool fetchSucceeded, string reason)
        {
            return new BackendPolicyStatus(
                endpointAvailable: endpointAvailable,
                fetchSucceeded: fetchSucceeded,
                policyActive: false,
                warningVisible: false,
                warningMessage: string.Empty,
                mode: "local",
                reason: reason,
                seatAssigned: false,
                isValid: false,
                seatState: string.Empty,
                sharePolicy: null,
                talkPolicy: null,
                shareEditable: null,
                talkEditable: null);
        }

        private bool ExecuteJsonRequest(string url, out HttpStatusCode statusCode, out IDictionary<string, object> parsed)
        {
            statusCode = 0;
            parsed = null;

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
            {                if (response.TransportException != null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed without HTTP response.", response.TransportException);
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed without HTTP response.", null);
                }
                return false;
            }

            statusCode = response.StatusCode;
            if ((int)statusCode < 200 || (int)statusCode >= 300)
            {
                return false;
            }

            parsed = response.ParsedJson;
            return parsed != null;
        }

        private static IDictionary<string, object> NormalizePayload(IDictionary<string, object> payload)
        {            if (payload == null)
            {
                return null;
            }

            IDictionary<string, object> ocs = NcJson.GetDictionary(payload, "ocs");
            IDictionary<string, object> data = NcJson.GetDictionary(ocs, "data");
            return data ?? payload;
        }

        private static string BuildSeatWarningMessage(bool seatAssigned, bool isValid, string seatState)
        {
            if (!seatAssigned)
            {
                return Strings.PolicyWarningNoSeat;
            }
            if (!isValid)
            {
                if (!string.IsNullOrWhiteSpace(seatState))
                {
                    return string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.PolicyWarningSeatStateFormat,
                        seatState);
                }
                return Strings.PolicyWarningLicenseInvalid;
            }
            if (!string.IsNullOrWhiteSpace(seatState))
            {
                return string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.PolicyWarningSeatStateFormat,
                    seatState);
            }
            return Strings.PolicyWarningLicenseInvalid;
        }

        private static bool ShouldWarnForSeat(IDictionary<string, object> status)
        {            if (status == null)
            {
                return false;
            }
            bool seatAssigned = GetBool(status, "seat_assigned");
            bool isValid = GetBool(status, "is_valid");
            string seatState = NcJson.GetStringOrEmpty(status, "seat_state");

            if (seatAssigned
                && (!isValid
                    || !string.Equals(seatState, "active", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
            return false;
        }
        private static bool GetBool(IDictionary<string, object> parent, string key)
        {            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            object raw;            if (!parent.TryGetValue(key, out raw) || raw == null)
            {
                return false;
            }
            bool value;
            return BackendPolicyStatus.TryConvertBool(raw, out value) && value;
        }
    }
}

