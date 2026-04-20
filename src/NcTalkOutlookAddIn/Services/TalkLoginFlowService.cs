/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Implements Nextcloud Login Flow v2 to automatically obtain app passwords.
     */
    internal sealed class TalkLoginFlowService
    {
        private const string DeviceName = "NC Connector for Outlook";
        private readonly string _baseUrl;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private readonly NcHttpClient _httpClient = new NcHttpClient(string.Empty, string.Empty);

        internal TalkLoginFlowService(string baseUrl)
        {
            _baseUrl = (baseUrl ?? string.Empty).Trim().TrimEnd('/');
        }

        internal LoginFlowStart StartLoginFlow()
        {
            if (string.IsNullOrWhiteSpace(_baseUrl))
            {
                throw new TalkServiceException("Server URL is empty.", false, HttpStatusCode.BadRequest, null);
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "LoginFlow.Start"))
            {
                string url = _baseUrl + "/index.php/login/v2";
                DiagnosticsLogger.LogApi("POST " + url);

            var payloadObject = new Dictionary<string, object>();
            payloadObject["name"] = DeviceName;

            string payload = _serializer.Serialize(payloadObject);
            HttpStatusCode statusCode;
            string responseText;
            IDictionary<string, object> data = ExecuteRequest(url, "POST", payload, out statusCode, out responseText);
            if (statusCode != HttpStatusCode.OK)
            {
                throw new TalkServiceException("HTTP " + (int)statusCode, statusCode == HttpStatusCode.Unauthorized, statusCode, responseText);
            }

            string loginUrl = GetString(data, "login");
            IDictionary<string, object> poll = GetDictionary(data, "poll");
            string pollEndpoint = GetString(poll, "endpoint");
            string pollToken = GetString(poll, "token");

            if (string.IsNullOrEmpty(loginUrl) || string.IsNullOrEmpty(pollEndpoint) || string.IsNullOrEmpty(pollToken))
            {
                throw new TalkServiceException("Login flow response is incomplete.", false, HttpStatusCode.InternalServerError, null);
            }

            if (!pollEndpoint.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                pollEndpoint = _baseUrl + pollEndpoint;
            }

            return new LoginFlowStart(loginUrl, pollEndpoint, pollToken);
            }
        }

        internal LoginFlowCredentials CompleteLoginFlow(LoginFlowStart start, TimeSpan timeout, TimeSpan pollInterval)
        {
            if (start == null)
            {
                throw new ArgumentNullException("start");
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "LoginFlow.Complete"))
            {
                DateTime expire = DateTime.UtcNow + (timeout <= TimeSpan.Zero ? TimeSpan.FromMinutes(2) : timeout);
                TimeSpan interval = pollInterval <= TimeSpan.Zero ? TimeSpan.FromSeconds(2) : pollInterval;

                while (DateTime.UtcNow < expire)
                {
                    var payloadObject = new Dictionary<string, object>();
                    payloadObject["token"] = start.Token;
                    payloadObject["deviceName"] = DeviceName;

                    string payload = _serializer.Serialize(payloadObject);
                    HttpStatusCode statusCode;
                    string responseText;
                    IDictionary<string, object> response = ExecuteRequest(start.PollEndpoint, "POST", payload, out statusCode, out responseText);

                    if (statusCode == HttpStatusCode.NotFound)
                    {
                        Thread.Sleep(interval);
                        continue;
                    }
                    if (statusCode != HttpStatusCode.OK)
                    {
                        throw new TalkServiceException("HTTP " + (int)statusCode, statusCode == HttpStatusCode.Unauthorized, statusCode, responseText);
                    }

                    string appPassword = GetString(response, "appPassword") ?? GetString(GetDictionary(response, "ocs"), "token");
                    string loginName = GetString(response, "loginName");

                    if (string.IsNullOrEmpty(appPassword) || string.IsNullOrEmpty(loginName))
                    {
                        throw new TalkServiceException("Login flow did not return an app password.", false, HttpStatusCode.InternalServerError, null);
                    }

                    return new LoginFlowCredentials(loginName, appPassword);
                }

                throw new TalkServiceException("Login flow did not complete (timeout).", false, HttpStatusCode.RequestTimeout, null);
            }
        }

        private IDictionary<string, object> ExecuteRequest(string url, string method, string payload, out HttpStatusCode statusCode, out string responseText)
        {
            statusCode = 0;
            NcHttpResponse response = _httpClient.Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Payload = payload,
                Accept = "application/json",
                ContentType = "application/json",
                TimeoutMs = 60000,
                IncludeAuthHeader = false,
                IncludeOcsApiHeader = true,
                ParseJson = true,
                ForceFreshConnection = true,
                UserAgent = BuildUserAgent()
            });

            if (!response.HasHttpResponse)
            {
                // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
                if (response.TransportException != null)
                {
                    HttpFailureInfo failure = response.FailureInfo ?? HttpFailureDiagnostics.Analyze(response.TransportException);
                    DiagnosticsLogger.LogException(LogCategories.Api, "Login flow request failed without HTTP response (" + HttpFailureDiagnostics.BuildLogSummary(response.TransportException, failure) + ").", response.TransportException);
                    throw new TalkServiceException(failure.BuildUserMessage(), false, 0, null, true);
                }

                throw new TalkServiceException("Login flow request failed without HTTP response.", false, 0, null, true);
            }

            statusCode = response.StatusCode;
            responseText = response.ResponseText;

            if (string.IsNullOrWhiteSpace(responseText))
            {
                return new Dictionary<string, object>();
            }

            // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
            if (response.JsonParseException != null)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Login flow JSON parsing failed.", response.JsonParseException);
                throw new TalkServiceException("Could not parse JSON: " + response.JsonParseException.Message, false, statusCode, responseText);
            }

            return response.ParsedJson ?? new Dictionary<string, object>();
        }

        private static IDictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            return NcJson.GetDictionary(parent, key);
        }

        private static string GetString(IDictionary<string, object> parent, string key)
        {
            return NcJson.GetString(parent, key);
        }

        private static string BuildUserAgent()
        {
            try
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string versionString = version != null ? version.ToString() : "0.0.0.0";
                string os = Environment.OSVersion.VersionString;
                return DeviceName + "/" + versionString + " (" + os + ")";
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Failed to build user agent string.", ex);
                return DeviceName;
            }
        }
    }

    internal sealed class LoginFlowStart
    {
        private readonly string _loginUrl;
        private readonly string _pollEndpoint;
        private readonly string _token;

        internal LoginFlowStart(string loginUrl, string pollEndpoint, string token)
        {
            _loginUrl = loginUrl;
            _pollEndpoint = pollEndpoint;
            _token = token;
        }

        internal string LoginUrl
        {
            get { return _loginUrl; }
        }

        internal string PollEndpoint
        {
            get { return _pollEndpoint; }
        }

        internal string Token
        {
            get { return _token; }
        }
    }

    internal sealed class LoginFlowCredentials
    {
        private readonly string _loginName;
        private readonly string _appPassword;

        internal LoginFlowCredentials(string loginName, string appPassword)
        {
            _loginName = loginName;
            _appPassword = appPassword;
        }

        internal string LoginName
        {
            get { return _loginName; }
        }

        internal string AppPassword
        {
            get { return _appPassword; }
        }
    }
}


