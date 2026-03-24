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
     * Fetches password policy information and the generator endpoint from Nextcloud capabilities.
     */
    internal sealed class PasswordPolicyService
    {
        private readonly TalkServiceConfiguration _configuration;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private static readonly string[] MinLengthKeys = { "minLength", "min_length", "minimumLength", "minimum_length" };
        private static readonly string[] GenerateUrlKeys = { "api_generate", "apiGenerateUrl", "generateUrl", "generate_url" };

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

                if (statusCode != HttpStatusCode.OK)
                {
                    LogApi("Password policy fetch returned HTTP " + (int)statusCode + ".");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }

                if (root == null)
                {
                    LogApi("Password policy fetch returned no parsable JSON payload.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }

                var ocs = GetDictionary(root, "ocs");
                var data = GetDictionary(ocs, "data");
                var caps = GetDictionary(data, "capabilities");
                if (caps == null)
                {
                    LogApi("Password policy capabilities block missing.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }

                var policy = GetDictionary(caps, "password_policy") ?? GetDictionary(caps, "passwordPolicy");
                if (policy == null)
                {
                    LogApi("Password policy block missing.");
                    return new PasswordPolicyInfo(false, 0, string.Empty);
                }

                int minLength = GetFirstPositiveInt(policy, MinLengthKeys);
                if (minLength <= 0)
                {
                    var policies = GetDictionary(policy, "policies");
                    var accountPolicy = GetDictionary(policies, "account");
                    minLength = GetFirstPositiveInt(accountPolicy, MinLengthKeys);
                }

                var api = GetDictionary(policy, "api");
                string generateRaw = GetFirstString(api, "generate", "generateUrl", "generate_url");
                if (string.IsNullOrWhiteSpace(generateRaw))
                {
                    generateRaw = GetFirstString(policy, GenerateUrlKeys);
                }

                string generateUrl = ResolvePolicyUrl(generateRaw, baseUrl);
                bool hasPolicy = minLength > 0 || !string.IsNullOrWhiteSpace(generateUrl);

                LogApi("Password policy normalized (hasPolicy=" + hasPolicy + ", minLength=" + minLength + ", hasGenerateUrl=" + (!string.IsNullOrWhiteSpace(generateUrl)) + ").");
                return new PasswordPolicyInfo(hasPolicy, minLength, generateUrl ?? string.Empty);
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
                string apiUrl = ResolvePolicyUrl(policy.GenerateUrl, _configuration.GetNormalizedBaseUrl());
                if (string.IsNullOrWhiteSpace(apiUrl))
                {
                    LogApi("Password generate skipped: no usable generator URL.");
                    return null;
                }

                LogApi("GET " + apiUrl);
                IDictionary<string, object> root;
                HttpStatusCode statusCode;
                ExecuteJsonRequest("GET", apiUrl, null, out statusCode, out root);

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
                request.Accept = "application/json, text/plain, */*";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Headers["OCS-APIRequest"] = "true";
                request.Timeout = 60000;

                bool hasBody = !string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase)
                               && !string.Equals(method, "DELETE", StringComparison.OrdinalIgnoreCase);

                if (hasBody)
                {
                    request.ContentType = "application/json";
                    if (!string.IsNullOrEmpty(payload))
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(payload);
                        request.ContentLength = bytes.Length;
                        using (var stream = request.GetRequestStream())
                        {
                            stream.Write(bytes, 0, bytes.Length);
                        }
                    }
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
            LogApi(method + " " + url + " -> " + statusCode);

            try
            {
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream ?? Stream.Null, Encoding.UTF8))
                {
                    responseText = reader.ReadToEnd();
                }

                if (!string.IsNullOrEmpty(responseText))
                {
                    string payloadText = PrepareJsonPayload(responseText);
                    if (!string.IsNullOrEmpty(payloadText))
                    {
                        parsed = _serializer.DeserializeObject(payloadText) as IDictionary<string, object>;
                    }
                }
            }
            catch (Exception ex)
            {
                parsed = null;
                DiagnosticsLogger.LogException(LogCategories.Api, "Password policy response parsing failed.", ex);
                LogApi("Password policy response sample: " + GetResponseSample(responseText));
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static string ResolvePolicyUrl(string rawUrl, string baseUrl)
        {
            string trimmed = rawUrl == null ? string.Empty : rawUrl.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return null;
            }

            try
            {
                Uri absoluteUri;
                if (Uri.TryCreate(trimmed, UriKind.Absolute, out absoluteUri))
                {
                    return absoluteUri.ToString();
                }

                Uri baseUri;
                if (!string.IsNullOrWhiteSpace(baseUrl) && Uri.TryCreate(baseUrl, UriKind.Absolute, out baseUri))
                {
                    Uri resolved = new Uri(baseUri, trimmed);
                    return resolved.ToString();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Api, "Password policy URL normalization failed.", ex);
            }

            return null;
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

        private static string GetResponseSample(string responseText)
        {
            if (string.IsNullOrEmpty(responseText))
            {
                return "<empty>";
            }

            string normalized = responseText.Replace("\r", " ").Replace("\n", " ").Trim();
            return normalized.Length <= 180 ? normalized : normalized.Substring(0, 180) + "...";
        }

        private static IDictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            object value;
            if (parent.TryGetValue(key, out value) && value != null)
            {
                return value as IDictionary<string, object>;
            }

            return null;
        }

        private static string GetString(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            object value;
            if (!parent.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            string direct = value as string;
            if (direct != null)
            {
                return direct;
            }

            if (value is IDictionary<string, object>)
            {
                return null;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string GetFirstString(IDictionary<string, object> parent, params string[] keys)
        {
            if (parent == null || keys == null)
            {
                return null;
            }

            foreach (string key in keys)
            {
                string value = GetString(parent, key);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return null;
        }

        private static int GetInt(IDictionary<string, object> parent, string key)
        {
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return 0;
            }

            object rawObject;
            if (!parent.TryGetValue(key, out rawObject) || rawObject == null)
            {
                return 0;
            }

            string raw = Convert.ToString(rawObject, CultureInfo.InvariantCulture);
            int value;
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }

            return 0;
        }

        private static int GetFirstPositiveInt(IDictionary<string, object> parent, params string[] keys)
        {
            if (parent == null || keys == null)
            {
                return 0;
            }

            foreach (string key in keys)
            {
                int value = GetInt(parent, key);
                if (value > 0)
                {
                    return value;
                }
            }

            return 0;
        }
    }
}
