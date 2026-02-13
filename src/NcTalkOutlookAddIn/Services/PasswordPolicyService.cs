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
     * Fetches password policy information (min_length) and the generator endpoint from the Nextcloud Password Policy app.
     */
    internal sealed class PasswordPolicyService
    {
        private readonly TalkServiceConfiguration _configuration;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        private static void LogApi(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Api, message);
        }

        internal PasswordPolicyService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _configuration = configuration;
        }

        internal PasswordPolicyInfo FetchPolicy()
        {
            if (!_configuration.IsComplete())
            {
                return new PasswordPolicyInfo(false, 0, string.Empty);
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "PasswordPolicy.FetchPolicy"))
            {
            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string url = baseUrl + "/ocs/v2.php/cloud/capabilities?format=json";
            LogApi("GET " + url);

            IDictionary<string, object> root;
            HttpStatusCode statusCode;
            ExecuteJsonRequest("GET", url, null, out statusCode, out root);

            if (statusCode != HttpStatusCode.OK || root == null)
            {
                return new PasswordPolicyInfo(false, 0, string.Empty);
            }

            var ocs = GetDictionary(root, "ocs");
            var data = GetDictionary(ocs, "data");
            var caps = GetDictionary(data, "capabilities");
            var policy = GetDictionary(caps, "password_policy");

            if (policy == null)
            {
                return new PasswordPolicyInfo(false, 0, string.Empty);
            }

            int minLength = GetInt(policy, "min_length");
            string generateUrl = GetString(policy, "api_generate") ?? string.Empty;

            bool hasPolicy = minLength > 0 || !string.IsNullOrEmpty(generateUrl);
            return new PasswordPolicyInfo(hasPolicy, minLength, generateUrl);
            }
        }

        internal string GeneratePassword(PasswordPolicyInfo policy)
        {
            if (policy == null || !policy.HasPolicy || string.IsNullOrEmpty(policy.GenerateUrl) || !_configuration.IsComplete())
            {
                return null;
            }

            using (DiagnosticsLogger.BeginOperation(LogCategories.Api, "PasswordPolicy.GeneratePassword"))
            {
                LogApi("GET " + policy.GenerateUrl);
            IDictionary<string, object> root;
            HttpStatusCode statusCode;
            ExecuteJsonRequest("GET", policy.GenerateUrl, null, out statusCode, out root);

            if (statusCode != HttpStatusCode.OK || root == null)
            {
                return null;
            }

            var data = GetDictionary(GetDictionary(root, "ocs"), "data");
            string password = GetString(data, "password");
            if (string.IsNullOrWhiteSpace(password))
            {
                return null;
            }

            return password.Trim();
            }
        }

        private void ExecuteJsonRequest(string method, string url, string payload, out HttpStatusCode statusCode, out IDictionary<string, object> parsed)
        {
            parsed = null;
            statusCode = 0;

            if (string.IsNullOrEmpty(url))
            {
                return;
            }

            HttpWebResponse response = null;
            string responseText = null;

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = method;
                request.Accept = "application/json";
                request.ContentType = "application/json";
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Headers["OCS-APIRequest"] = "true";
                request.Timeout = 60000;

                if (!string.IsNullOrEmpty(payload))
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(payload);
                    request.ContentLength = bytes.Length;
                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
                else
                {
                    request.ContentLength = 0;
                }

                response = (HttpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                response = ex.Response as HttpWebResponse;
                if (response == null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Api, "Password policy request failed without HTTP response.", ex);
                    return;
                }
            }

            statusCode = response.StatusCode;

            try
            {
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream ?? Stream.Null))
                {
                    responseText = reader.ReadToEnd();
                }

                if (!string.IsNullOrEmpty(responseText))
                {
                    parsed = _serializer.DeserializeObject(responseText) as IDictionary<string, object>;
                }
            }
            catch (Exception ex)
            {
                parsed = null;
                DiagnosticsLogger.LogException(LogCategories.Api, "Password policy response parsing failed.", ex);
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static Dictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            if (parent == null)
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value))
            {
                return value as Dictionary<string, object>;
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
                if (value is string)
                {
                    return (string)value;
                }

                if (value is IDictionary<string, object>)
                {
                    return null;
                }

                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static int GetInt(IDictionary<string, object> parent, string key)
        {
            string raw = GetString(parent, key);
            int value;
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }

            if (parent != null)
            {
                object obj;
                if (parent.TryGetValue(key, out obj) && obj is int)
                {
                    return (int)obj;
                }
            }

            return 0;
        }
    }
}
