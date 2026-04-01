/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Loads centralized NC Connector policy status from Nextcloud backend endpoint.
     */
    internal sealed class BackendPolicyService
    {
        private const string StatusEndpointPath = "/apps/ncc_backend_4mc/api/v1/status";
        private readonly TalkServiceConfiguration _configuration;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        internal BackendPolicyService(TalkServiceConfiguration configuration)
        {
            _configuration = configuration;
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
            IDictionary<string, object> status = GetDictionary(normalized, "status");
            IDictionary<string, object> policy = GetDictionary(normalized, "policy");
            IDictionary<string, object> policyEditable = GetDictionary(normalized, "policy_editable");
            IDictionary<string, object> sharePolicy = GetDictionary(policy, "share");
            IDictionary<string, object> talkPolicy = GetDictionary(policy, "talk");
            IDictionary<string, object> shareEditable = GetDictionary(policyEditable, "share");
            IDictionary<string, object> talkEditable = GetDictionary(policyEditable, "talk");

            bool seatAssigned = GetBool(status, "seat_assigned");
            bool isValid = GetBool(status, "is_valid");
            string seatState = GetString(status, "seat_state");

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
            HttpWebResponse response = null;
            string responseText = null;

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.Accept = "application/json, text/plain, */*";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Headers["OCS-APIRequest"] = "true";
                request.Timeout = 45000;

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException ex)
                {
                    response = ex.Response as HttpWebResponse;
                    if (response == null)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed without HTTP response.", ex);
                        return false;
                    }
                }

                statusCode = response.StatusCode;
                if ((int)statusCode < 200 || (int)statusCode >= 300)
                {
                    return false;
                }

                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream ?? Stream.Null, Encoding.UTF8))
                {
                    responseText = reader.ReadToEnd();
                }

                if (string.IsNullOrWhiteSpace(responseText))
                {
                    return false;
                }

                string prepared = PrepareJsonPayload(responseText);
                if (string.IsNullOrWhiteSpace(prepared))
                {
                    return false;
                }

                parsed = _serializer.DeserializeObject(prepared) as IDictionary<string, object>;
                return parsed != null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Policy status request failed.", ex);
                return false;
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static IDictionary<string, object> NormalizePayload(IDictionary<string, object> payload)
        {
            if (payload == null)
            {
                return null;
            }

            IDictionary<string, object> ocs = GetDictionary(payload, "ocs");
            IDictionary<string, object> data = GetDictionary(ocs, "data");
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
        {
            if (status == null)
            {
                return false;
            }

            bool seatAssigned = GetBool(status, "seat_assigned");
            bool isValid = GetBool(status, "is_valid");
            string seatState = GetString(status, "seat_state");

            if (seatAssigned
                && (!isValid
                    || !string.Equals(seatState, "active", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return false;
        }

        private static string PrepareJsonPayload(string responseText)
        {
            if (string.IsNullOrWhiteSpace(responseText))
            {
                return string.Empty;
            }

            string payload = responseText.Trim().TrimStart('\uFEFF');
            if (payload.StartsWith(")]}',", StringComparison.Ordinal))
            {
                int newlineIndex = payload.IndexOf('\n');
                payload = newlineIndex >= 0 ? payload.Substring(newlineIndex + 1) : string.Empty;
            }

            if (payload.StartsWith("while(1);", StringComparison.Ordinal))
            {
                payload = payload.Substring("while(1);".Length);
            }

            if (payload.StartsWith("for(;;);", StringComparison.Ordinal))
            {
                payload = payload.Substring("for(;;);".Length);
            }

            return payload.Trim();
        }

        private static IDictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            object raw;
            if (!parent.TryGetValue(key, out raw) || raw == null)
            {
                return null;
            }

            return raw as IDictionary<string, object>;
        }

        private static string GetString(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            object raw;
            if (!parent.TryGetValue(key, out raw) || raw == null)
            {
                return string.Empty;
            }

            string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
            return text == null ? string.Empty : text.Trim();
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
