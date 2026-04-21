// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class NcHttpRequestOptions
    {
        internal NcHttpRequestOptions()
        {
            Method = "GET";
            Accept = "application/json, text/plain, */*";
            ContentType = "application/json";
            RequestEncoding = Encoding.UTF8;
            ResponseEncoding = Encoding.UTF8;
            TimeoutMs = 60000;
            IncludeAuthHeader = true;
            IncludeOcsApiHeader = true;
            EnableAutomaticDecompression = true;
            ParseJson = true;
        }

        internal string Method { get; set; }
        internal string Url { get; set; }
        internal string Payload { get; set; }
        internal byte[] PayloadBytes { get; set; }
        internal Action<Stream> BodyWriter { get; set; }
        internal string Accept { get; set; }
        internal string ContentType { get; set; }
        internal string UserAgent { get; set; }
        internal IDictionary<string, string> Headers { get; set; }
        internal Encoding RequestEncoding { get; set; }
        internal Encoding ResponseEncoding { get; set; }
        internal int TimeoutMs { get; set; }
        internal bool IncludeAuthHeader { get; set; }
        internal bool IncludeOcsApiHeader { get; set; }
        internal bool EnableAutomaticDecompression { get; set; }
        internal bool ParseJson { get; set; }
        internal bool ForceFreshConnection { get; set; }
        internal bool ReadResponseAsBytes { get; set; }
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse { get; set; }
        internal HttpStatusCode StatusCode { get; set; }
        internal string ContentType { get; set; }
        internal string ResponseText { get; set; }
        internal byte[] ResponseBytes { get; set; }
        internal IDictionary<string, object> ParsedJson { get; set; }
        internal WebException TransportException { get; set; }
        internal HttpFailureInfo FailureInfo { get; set; }
        internal Exception JsonParseException { get; set; }
    }

        // Internal HTTP client wrapper that keeps auth/header/timeout behavior consistent.
    internal sealed class NcHttpClient
    {
        private readonly string _username;
        private readonly string _appPassword;

        internal NcHttpClient(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }

            _username = configuration.Username ?? string.Empty;
            _appPassword = configuration.AppPassword ?? string.Empty;
        }

        internal NcHttpClient(string username, string appPassword)
        {
            _username = username ?? string.Empty;
            _appPassword = appPassword ?? string.Empty;
        }

        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException("options");
            }
            if (string.IsNullOrWhiteSpace(options.Url))
            {
                throw new ArgumentException("URL is required.", "options");
            }
            string method = string.IsNullOrWhiteSpace(options.Method) ? "GET" : options.Method.Trim().ToUpperInvariant();
            var result = new NcHttpResponse();

            HttpWebRequest request = null;
            HttpWebResponse response = null;
            string connectionGroupName = null;

            try
            {
                request = (HttpWebRequest)WebRequest.Create(options.Url);
                request.Method = method;
                request.Accept = string.IsNullOrWhiteSpace(options.Accept)
                    ? "application/json, text/plain, */*"
                    : options.Accept;
                request.Timeout = options.TimeoutMs > 0 ? options.TimeoutMs : 60000;
                if (!string.IsNullOrWhiteSpace(options.UserAgent))
                {
                    request.UserAgent = options.UserAgent;
                }
                if (options.EnableAutomaticDecompression)
                {
                    request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                }
                if (options.IncludeAuthHeader)
                {
                    request.Headers["Authorization"] =
                        HttpAuthUtilities.BuildBasicAuthHeader(_username, _appPassword);
                }
                if (options.IncludeOcsApiHeader)
                {
                    request.Headers["OCS-APIRequest"] = "true";
                }
                if (options.Headers != null)
                {
                    foreach (var kvp in options.Headers)
                    {
                        if (!string.IsNullOrWhiteSpace(kvp.Key))
                        {
                            request.Headers[kvp.Key] = kvp.Value ?? string.Empty;
                        }
                    }
                }
                if (options.ForceFreshConnection)
                {
                    connectionGroupName = "nc-http-" + Guid.NewGuid().ToString("N");
                    request.ConnectionGroupName = connectionGroupName;
                    request.KeepAlive = false;
                    request.Pipelined = false;
                }
                bool hasBody = !string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase)
                               && !string.Equals(method, "DELETE", StringComparison.OrdinalIgnoreCase);

                if (hasBody)
                {
                    request.ContentType = string.IsNullOrWhiteSpace(options.ContentType)
                        ? "application/json"
                        : options.ContentType;                    if (options.BodyWriter != null)
                    {
                        using (Stream stream = request.GetRequestStream())
                        {
                            options.BodyWriter(stream);
                        }
                    }
                    else
                    {
                        byte[] bytes = options.PayloadBytes;                        if (bytes == null)
                        {
                            string payload = options.Payload ?? string.Empty;
                            Encoding requestEncoding = options.RequestEncoding ?? Encoding.UTF8;
                            bytes = requestEncoding.GetBytes(payload);
                        }

                        request.ContentLength = bytes.Length;
                        if (bytes.Length > 0)
                        {
                            using (Stream stream = request.GetRequestStream())
                            {
                                stream.Write(bytes, 0, bytes.Length);
                            }
                        }
                    }
                }
                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException ex)
                {
                    response = ex.Response as HttpWebResponse;                    if (response == null)
                    {
                        result.HasHttpResponse = false;
                        result.TransportException = ex;
                        result.FailureInfo = HttpFailureDiagnostics.Analyze(ex);
                        return result;
                    }
                }

                result.HasHttpResponse = true;
                result.StatusCode = response.StatusCode;
                result.ContentType = response.ContentType ?? string.Empty;

                using (Stream stream = response.GetResponseStream() ?? Stream.Null)
                {
                    if (options.ReadResponseAsBytes)
                    {
                        using (var memory = new MemoryStream())
                        {
                            stream.CopyTo(memory);
                            result.ResponseBytes = memory.ToArray();
                        }
                    }
                    else
                    {
                        using (StreamReader reader = new StreamReader(stream, options.ResponseEncoding ?? Encoding.UTF8))
                        {
                            result.ResponseText = reader.ReadToEnd();
                        }
                    }
                }
                if (result.ResponseText == null && result.ResponseBytes != null && result.ResponseBytes.Length > 0)
                {
                    Encoding responseEncoding = options.ResponseEncoding ?? Encoding.UTF8;
                    result.ResponseText = responseEncoding.GetString(result.ResponseBytes);
                }
                if (options.ParseJson && !string.IsNullOrWhiteSpace(result.ResponseText))
                {
                    try
                    {
                        result.ParsedJson = NcJson.DeserializeObject(result.ResponseText);
                    }
                    catch (Exception ex)
                    {
                        result.ParsedJson = null;
                        result.JsonParseException = ex;
                    }
                }
            }
            finally
            {                if (response != null)
                {
                    response.Close();
                }
                if (options.ForceFreshConnection &&
                    request != null &&
                    !string.IsNullOrEmpty(connectionGroupName))
                {
                    try
                    {
                        request.ServicePoint.CloseConnectionGroup(connectionGroupName);
                    }
                    catch
                    {
                        // Best-effort connection group cleanup.
                    }
                }
            }
            return result;
        }
    }
}

