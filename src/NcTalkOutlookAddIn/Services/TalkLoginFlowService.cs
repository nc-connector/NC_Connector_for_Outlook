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

        private static void LogApi(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Api, message);
        }

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
                LogApi("POST " + url);

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
            HttpWebResponse response = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = method;
                request.Accept = "application/json";
                request.Timeout = 60000;
                request.Headers["OCS-APIRequest"] = "true";
                request.UserAgent = BuildUserAgent();

                if (!string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase) && payload != null)
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(payload);
                    request.ContentType = "application/json";
                    request.ContentLength = bytes.Length;
                    using (Stream stream = request.GetRequestStream())
                    {
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
                else if (!string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase))
                {
                    request.ContentLength = 0;
                }

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException ex)
                {
                    response = ex.Response as HttpWebResponse;
                    if (response == null)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Api, "Login flow request failed without HTTP response.", ex);
                        throw new TalkServiceException("HTTP connection error: " + ex.Message, false, 0, null);
                    }
                }

                statusCode = response.StatusCode;
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream ?? Stream.Null))
                {
                    responseText = reader.ReadToEnd();
                }

                if (string.IsNullOrWhiteSpace(responseText))
                {
                    return new Dictionary<string, object>();
                }

                try
                {
                    return _serializer.DeserializeObject(responseText) as IDictionary<string, object> ?? new Dictionary<string, object>();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Api, "Login flow JSON parsing failed.", ex);
                    throw new TalkServiceException("Could not parse JSON: " + ex.Message, false, response.StatusCode, responseText);
                }
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static IDictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value))
            {
                return value as IDictionary<string, object>;
            }

            return null;
        }

        private static string GetString(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value) && value != null)
            {
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return null;
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


