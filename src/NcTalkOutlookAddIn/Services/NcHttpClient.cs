// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
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
            ContentLength = -1;
            MaximumResponseBytes = 0;
            CancellationToken = CancellationToken.None;
            AllowWriteStreamBuffering = true;
        }

        internal string Method { get; set; }
        internal string Url { get; set; }
        internal string Payload { get; set; }
        internal byte[] PayloadBytes { get; set; }
        internal Action<Stream> BodyWriter { get; set; }
        internal long ContentLength { get; set; }
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
        internal bool VerifyRejectedCredentials { get; set; }
        internal bool ReadResponseAsBytes { get; set; }
        internal long MaximumResponseBytes { get; set; }
        internal CancellationToken CancellationToken { get; set; }
        internal int ReadWriteTimeoutMs { get; set; }
        internal int ConnectionLimit { get; set; }
        internal bool AllowWriteStreamBuffering { get; set; }
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
        internal IDictionary<string, string> Headers { get; set; }
        internal long RequestSequence { get; set; }
    }

    // Internal HTTP client wrapper that keeps auth/header/timeout behavior consistent.
    internal sealed class NcHttpClient
    {
        private static long _requestSequence;
        private readonly TalkServiceConfiguration _configuration;
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
            _configuration = configuration;
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
            options.CancellationToken.ThrowIfCancellationRequested();
            string method = string.IsNullOrWhiteSpace(options.Method) ? "GET" : options.Method.Trim().ToUpperInvariant();
            ConnectionPauseStatus pause = NextcloudConnectionState.GetRequestStatus(
                _configuration, options.Url, options.VerifyRejectedCredentials, options.IncludeAuthHeader);
            if (pause.IsPaused)
            {
                return new NcHttpResponse
                {
                    HasHttpResponse = true,
                    StatusCode = pause.Reason == "rate_limited" ? (HttpStatusCode)429 : HttpStatusCode.Unauthorized,
                    Headers = pause.Reason == "rate_limited"
                        ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            { "Retry-After", pause.RetryAfterUtc.ToString("R", CultureInfo.InvariantCulture) }
                        } : null
                };
            }
            var result = new NcHttpResponse
            {
                RequestSequence = Interlocked.Increment(ref _requestSequence)
            };

            HttpWebRequest request = null;
            HttpWebResponse response = null;
            string connectionGroupName = null;
            CancellationTokenRegistration cancellationRegistration =
                default(CancellationTokenRegistration);

            try
            {
                request = TransportSecurityConfigurator.CreateRequest(options.Url);
                request.Method = method;
                request.Accept = string.IsNullOrWhiteSpace(options.Accept)
                    ? "application/json, text/plain, */*"
                    : options.Accept;
                request.Timeout = options.TimeoutMs > 0 ? options.TimeoutMs : 60000;
                request.ReadWriteTimeout = options.ReadWriteTimeoutMs > 0
                    ? options.ReadWriteTimeoutMs
                    : request.Timeout;
                request.AllowWriteStreamBuffering = options.AllowWriteStreamBuffering;
                if (options.ConnectionLimit > 0
                    && request.ServicePoint.ConnectionLimit < options.ConnectionLimit)
                {
                    request.ServicePoint.ConnectionLimit = options.ConnectionLimit;
                }
                if (options.CancellationToken.CanBeCanceled)
                {
                    cancellationRegistration = options.CancellationToken.Register(() =>
                    {
                        try
                        {
                            request.Abort();
                        }
                        catch
                        {
                            // Request cancellation can race with response disposal.
                        }
                    });
                }
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
                        : options.ContentType;
                    if (options.BodyWriter != null)
                    {
                        if (options.ContentLength >= 0)
                        {
                            request.ContentLength = options.ContentLength;
                        }
                        using (Stream stream = request.GetRequestStream())
                        {
                            options.BodyWriter(stream);
                        }
                    }
                    else
                    {
                        byte[] bytes = options.PayloadBytes;
                        if (bytes == null)
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
                    response = ex.Response as HttpWebResponse;
                    if (response == null)
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
                result.Headers = CopyHeaders(response.Headers);
                NextcloudConnectionState.RecordResponse(_configuration, options.Url, options.IncludeAuthHeader, result);

                using (Stream stream = response.GetResponseStream() ?? Stream.Null)
                {
                    if (options.ReadResponseAsBytes)
                    {
                        if (options.MaximumResponseBytes > 0
                            && response.ContentLength
                            > options.MaximumResponseBytes)
                        {
                            throw new InvalidDataException(
                                "The HTTP response exceeds the configured byte limit.");
                        }
                        result.ResponseBytes = ReadResponseBytes(
                            stream,
                            options.MaximumResponseBytes,
                            options.CancellationToken);
                    }
                    else
                    {
                        using (StreamReader reader = new StreamReader(stream, options.ResponseEncoding ?? Encoding.UTF8))
                        {
                            result.ResponseText = reader.ReadToEnd();
                        }
                    }
                }
                if (options.ParseJson
                    && result.ResponseText == null
                    && result.ResponseBytes != null
                    && result.ResponseBytes.Length > 0)
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
            catch (WebException ex)
            {
                result.HasHttpResponse = false;
                result.TransportException = ex;
                result.FailureInfo = HttpFailureDiagnostics.Analyze(ex);
                return result;
            }
            catch (IOException ex)
            {
                WebException transport = FindWebException(ex);
                if (transport == null)
                {
                    throw;
                }

                result.HasHttpResponse = false;
                result.TransportException = transport;
                result.FailureInfo = HttpFailureDiagnostics.Analyze(transport);
                return result;
            }
            finally
            {
                cancellationRegistration.Dispose();
                if (response != null)
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

        private static byte[] ReadResponseBytes(
            Stream source,
            long maximumBytes,
            CancellationToken cancellationToken)
        {
            const int BufferSize = 81920;
            var buffer = new byte[BufferSize];
            using (var memory = new MemoryStream())
            {
                int read;
                while ((read = source.Read(
                    buffer,
                    0,
                    buffer.Length)) > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (maximumBytes > 0
                        && memory.Length > maximumBytes - read)
                    {
                        throw new InvalidDataException(
                            "The HTTP response exceeds the configured byte limit.");
                    }
                    memory.Write(buffer, 0, read);
                }
                cancellationToken.ThrowIfCancellationRequested();
                return memory.ToArray();
            }
        }

        private static WebException FindWebException(Exception exception)
        {
            Exception current = exception;
            while (current != null)
            {
                WebException webException = current as WebException;
                if (webException != null)
                {
                    return webException;
                }
                current = current.InnerException;
            }
            return null;
        }

        private static IDictionary<string, string> CopyHeaders(WebHeaderCollection source)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (source == null)
            {
                return result;
            }

            foreach (string key in source.AllKeys)
            {
                if (!string.IsNullOrWhiteSpace(key))
                {
                    result[key] = source[key] ?? string.Empty;
                }
            }
            return result;
        }
    }
}

