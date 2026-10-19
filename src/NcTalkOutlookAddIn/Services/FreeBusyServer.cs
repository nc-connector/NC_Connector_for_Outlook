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
using System.Threading.Tasks;
using System.Xml;
using NcTalkOutlookAddIn.Utilities;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Services
{
        // Local HTTP endpoint for Outlook that proxies free/busy data to Nextcloud.
    internal sealed class FreeBusyServer : IDisposable
    {
        private const int DefaultPort = 7777;
        private const int MinPort = 1024;
        private const int MaxPort = 49151;
        private const int MaxConcurrentRequests = 4;
        private const string PrefixPath = "/nc-ifb/";
        private const string LogCategory = LogCategories.Ifb;
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        private readonly object _syncRoot = new object();
        private readonly IfbAddressBookCache _addressBookCache;
        private readonly CalendarResourceDirectory _resourceDirectory;
        private readonly Func<TalkServiceConfiguration, BackendPolicyStatus> _fetchRolloutPolicy;
        private bool _enterpriseRollout;
        private readonly SemaphoreSlim _requestSlots =
            new SemaphoreSlim(
                MaxConcurrentRequests,
                MaxConcurrentRequests);

        private HttpListener _listener;
        private CancellationTokenSource _cancellation;
        private Task _listenerTask;

        private TalkServiceConfiguration _configuration;
        private int _defaultDays = 30;
        private int _cacheHours = 24;
        private int _listenPort = DefaultPort;
        private string _listenPrefix = BuildPrefix(DefaultPort);
        private string _requestSecret = string.Empty;

        internal FreeBusyServer(IfbAddressBookCache addressBookCache,
            Func<TalkServiceConfiguration, BackendPolicyStatus> fetchRolloutPolicy = null,
            CalendarResourceDirectory resourceDirectory = null)
        {
            if (addressBookCache == null)
            {
                throw new ArgumentNullException("addressBookCache");
            }

            _addressBookCache = addressBookCache;
            _fetchRolloutPolicy = fetchRolloutPolicy;
            _resourceDirectory = resourceDirectory ?? new CalendarResourceDirectory(fetchRolloutPolicy);
        }

        internal void UpdateSettings(
            TalkServiceConfiguration configuration,
            int defaultDays,
            int cacheHours,
            string requestSecret,
            bool enterpriseRollout = false)
        {
            lock (_syncRoot)
            {
                if (string.IsNullOrWhiteSpace(requestSecret))
                {
                    throw new ArgumentException(
                        "IFB request secret is required.",
                        "requestSecret");
                }
                _configuration = configuration;
                _enterpriseRollout = enterpriseRollout;
                _defaultDays = defaultDays > 0 ? defaultDays : 30;
                _cacheHours = cacheHours >= 1 ? cacheHours : 24;
                _requestSecret = requestSecret.Trim();
            }
        }

        internal void Start(int port)
        {
            lock (_syncRoot)
            {
                int normalizedPort = NormalizePort(port);
                if (_listener != null && _listener.IsListening)
                {
                    if (_listenPort == normalizedPort)
                    {
                        return;
                    }

                    DiagnosticsLogger.Log(
                        LogCategory,
                        "Restarting listener due to port change (oldPort="
                        + _listenPort.ToString(CultureInfo.InvariantCulture)
                        + ", newPort="
                        + normalizedPort.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    StopListenerLocked();
                }
                if (_configuration == null || !_configuration.IsComplete())
                {
                    throw new InvalidOperationException("IFB cannot be started: credentials are incomplete.");
                }

                _listenPort = normalizedPort;
                _listenPrefix = BuildPrefix(_listenPort);
                _listener = new HttpListener();
                _listener.Prefixes.Add(_listenPrefix);
                _listener.Start();

                DiagnosticsLogger.Log(LogCategory, "Listener started (prefix=" + _listenPrefix + ").");

                _cancellation = new CancellationTokenSource();
                _listenerTask = Task.Run(() => ListenLoop(_cancellation.Token));
            }
        }

        internal void Stop()
        {
            lock (_syncRoot)
            {
                StopListenerLocked();
            }
        }

        private void StopListenerLocked()
        {
            if (_cancellation != null)
            {
                _cancellation.Cancel();
                _cancellation.Dispose();
                _cancellation = null;
            }
            if (_listener != null)
            {
                try
                {
                    DiagnosticsLogger.Log(LogCategory, "Stopping listener.");
                    _listener.Stop();
                    _listener.Close();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategory, "Failed to stop IFB listener.", ex);
                }
                finally
                {
                    _listener = null;
                }
            }
            if (_listenerTask != null)
            {
                try
                {
                    _listenerTask.Wait(1000);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategory, "Failed while waiting for IFB listener task shutdown.", ex);
                }
                finally
                {
                    _listenerTask = null;
                }
            }
        }

        private static string BuildPrefix(int port)
        {
            return "http://127.0.0.1:"
                   + port.ToString(CultureInfo.InvariantCulture)
                   + PrefixPath;
        }

        private static int NormalizePort(int port)
        {
            if (port < MinPort || port > MaxPort)
            {
                return DefaultPort;
            }
            return port;
        }

        private void ListenLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                HttpListenerContext context = null;
                try
                {
                    context = _listener.GetContext();
                    if (context != null)
                    {
                        DiagnosticsLogger.Log(
                            LogCategory,
                            "HTTP " + context.Request.HttpMethod + " request received.");
                    }
                }
                catch (HttpListenerException ex)
                {
                    if (token.IsCancellationRequested)
                    {
                        DiagnosticsLogger.Log(LogCategory, "Listener loop cancelled.");
                        return;
                    }

                    DiagnosticsLogger.LogException(LogCategory, "IFB listener failed while accepting request.", ex);
                }
                catch (ObjectDisposedException ex)
                {
                    DiagnosticsLogger.LogException(LogCategory, "IFB listener disposed while accepting request.", ex);
                    return;
                }
                if (context != null)
                {
                    if (!_requestSlots.Wait(0))
                    {
                        DiagnosticsLogger.Log(
                            LogCategory,
                            "IFB request rejected because the concurrency limit was reached.");
                        WriteError(
                            context,
                            (HttpStatusCode)429,
                            "Too many concurrent requests.");
                        continue;
                    }

                    Task.Run(
                        () =>
                        {
                            try
                            {
                                HandleRequest(context);
                            }
                            finally
                            {
                                _requestSlots.Release();
                            }
                        });
                }
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            try
            {
                if (!string.Equals(context.Request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
                {
                    DiagnosticsLogger.Log(LogCategory, "Method not allowed: " + context.Request.HttpMethod);
                    WriteError(context, HttpStatusCode.MethodNotAllowed, "Only GET is supported.");
                    return;
                }
                string path = context.Request.Url.AbsolutePath ?? string.Empty;
                string namePart;
                if (!TryParseAuthorizedEmailPath(
                        path,
                        out namePart))
                {
                    DiagnosticsLogger.Log(LogCategory, "Unsupported or unauthorized IFB request path.");
                    WriteError(context, HttpStatusCode.NotFound, "Unsupported path.");
                    return;
                }
                string email = CalendarResourceDirectory.NormalizeEmail(Uri.UnescapeDataString(namePart ?? string.Empty));
                if (string.IsNullOrEmpty(email))
                {
                    DiagnosticsLogger.Log(LogCategory, "Request rejected without email.");
                    WriteError(context, HttpStatusCode.BadRequest, "Email address is missing.");
                    return;
                }
                TalkServiceConfiguration configuration;
                int days;
                int cacheHours;
                bool enterpriseRollout;
                lock (_syncRoot)
                {
                    configuration = _configuration;
                    days = _defaultDays;
                    cacheHours = _cacheHours;
                    enterpriseRollout = _enterpriseRollout;
                }
                string daysParam = context.Request.QueryString["days"];
                int parsedDays;
                if (!string.IsNullOrEmpty(daysParam) && int.TryParse(daysParam, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedDays))
                {
                    if (parsedDays > 0 && parsedDays <= 365)
                    {
                        days = parsedDays;
                    }
                }
                string freeBusy = GetFreeBusyPayload(configuration, email, days, cacheHours, enterpriseRollout);
                if (string.IsNullOrEmpty(freeBusy))
                {
                    DiagnosticsLogger.Log(LogCategory, "No free/busy data available for " + email);
                    WriteError(context, HttpStatusCode.NoContent, "No free/busy data available.");
                    return;
                }

                DiagnosticsLogger.Log(LogCategory, "Response ok for " + email + " (len=" + freeBusy.Length + ").");

                context.Response.StatusCode = (int)HttpStatusCode.OK;
                context.Response.ContentType = "text/calendar; charset=utf-8";
                using (var writer = new StreamWriter(context.Response.OutputStream, Utf8NoBom))
                {
                    writer.Write(freeBusy);
                }
            }
            catch (FreeBusyRequestException ex)
            {
                DiagnosticsLogger.Log(LogCategory, "IFB request failed: " + ex.Message);
                WriteError(context, ex.StatusCode ?? HttpStatusCode.BadGateway, ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategory, "IFB request failed.", ex);
                WriteError(context, HttpStatusCode.BadGateway, "Free/busy data could not be retrieved.");
            }
            finally
            {
                try
                {
                    context.Response.OutputStream.Close();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategory, "Failed to close response output stream.", ex);
                }
            }
        }

        internal string GetFreeBusyPayload(TalkServiceConfiguration configuration, string email, int days, int cacheHours, bool enterpriseRollout)
        {
            if (configuration == null || !configuration.IsComplete()) { throw new InvalidOperationException("IFB credentials are incomplete."); }
            email = CalendarResourceDirectory.NormalizeEmail(email);
            if (email.Length == 0 || days < 1 || days > 365) { throw new FreeBusyRequestException("Invalid free/busy query.", HttpStatusCode.BadRequest, false); }
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            if (enterpriseRollout)
            {
                BackendPolicyStatus status = _fetchRolloutPolicy != null ? _fetchRolloutPolicy(configuration) : null;
                if (status == null || !status.FetchSucceeded || !PolicyUiHelper.HasBackendSeatEntitlement(status))
                {
                    throw new FreeBusyRequestException("Managed rollout access unavailable.", HttpStatusCode.Forbidden, false);
                }
            }
            string uid;
            // A failed address-book read throws. Only a confirmed miss enters resource discovery.
            if (!_addressBookCache.TryGetUid(configuration, cacheHours, email, out uid) || string.IsNullOrEmpty(uid))
            {
                CalendarResource resource = _resourceDirectory.FindByEmail(configuration, cacheHours, email);
                if (resource == null) { throw new FreeBusyRequestException("Calendar participant not found.", HttpStatusCode.NotFound, false); }
                string resourceOriginator;
                string resourceOutbox;
                _resourceDirectory.GetSchedulingIdentity(configuration, out resourceOriginator, out resourceOutbox);
                DiagnosticsLogger.Log(LogCategory, "Resource IFB request through the current user's scheduling outbox.");
                return RequestFreeBusyViaScheduling(configuration, resourceOutbox, resourceOriginator, resource.Email, days);
            }

            string freeBusy = null;
            try { freeBusy = RequestFreeBusyFromCalendar(configuration, uid, days); }
            catch (FreeBusyRequestException ex)
            {
                if (!ex.ShouldFallback) { throw; }
                DiagnosticsLogger.Log(LogCategory, "Person calendar REPORT requires scheduling fallback.");
            }
            if (!string.IsNullOrEmpty(freeBusy)) { return freeBusy; }
            string originEmail;
            string currentUserId = NextcloudUserIdentityService.ResolveCurrentUserId(configuration);
            if (!_addressBookCache.TryGetPrimaryEmailForUid(configuration, cacheHours, currentUserId, out originEmail) || string.IsNullOrEmpty(originEmail))
            {
                throw new FreeBusyRequestException("Own user account not found in address book.", HttpStatusCode.PreconditionFailed, false);
            }
            string outboxUrl = configuration.GetNormalizedBaseUrl() + "/remote.php/dav/calendars/" + Uri.EscapeDataString(currentUserId) + "/outbox/";
            return RequestFreeBusyViaScheduling(configuration, outboxUrl, originEmail, email, days);
        }

        internal bool TryParseAuthorizedEmailPath(
            string path,
            out string encodedEmail)
        {
            encodedEmail = string.Empty;
            string[] segments = (path ?? string.Empty).Split(
                new[] { '/' },
                StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length != 4
                || !string.Equals(
                    segments[0],
                    "nc-ifb",
                    StringComparison.OrdinalIgnoreCase)
                || !string.Equals(
                    segments[2],
                    "freebusy",
                    StringComparison.OrdinalIgnoreCase)
                || !segments[3].EndsWith(
                    ".vfb",
                    StringComparison.OrdinalIgnoreCase)
                || !SecretMatches(segments[1]))
            {
                return false;
            }

            encodedEmail = segments[3].Substring(
                0,
                segments[3].Length - ".vfb".Length);
            return true;
        }

        private bool SecretMatches(string candidate)
        {
            string expected;
            lock (_syncRoot)
            {
                expected = _requestSecret ?? string.Empty;
            }
            if (expected.Length == 0 || candidate == null)
            {
                return false;
            }

            int difference = expected.Length ^ candidate.Length;
            int length = Math.Max(
                expected.Length,
                candidate.Length);
            for (int i = 0; i < length; i++)
            {
                difference |=
                    (i < expected.Length
                        ? expected[i]
                        : '\0')
                    ^ (i < candidate.Length
                        ? candidate[i]
                        : '\0');
            }
            return difference == 0;
        }

        private string RequestFreeBusyFromCalendar(TalkServiceConfiguration configuration, string uid, int days)
        {
            DateTime startUtc = DateTime.UtcNow;
            DateTime endUtc = startUtc.AddDays(days);

            string baseUrl = configuration.GetNormalizedBaseUrl();
            string calendarUrl = string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/calendars/{1}/personal",
                baseUrl,
                Uri.EscapeDataString(uid));

            string reportBody = string.Format(
                CultureInfo.InvariantCulture,
                "<?xml version=\"1.0\" encoding=\"utf-8\" ?><c:free-busy-query xmlns:c=\"urn:ietf:params:xml:ns:caldav\" xmlns:d=\"DAV:\"><c:time-range start=\"{0}\" end=\"{1}\"/></c:free-busy-query>",
                startUtc.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture),
                endUtc.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture));

            var httpClient = new NcHttpClient(configuration);
            NcHttpResponse response = httpClient.Send(new NcHttpRequestOptions
            {
                Method = "REPORT",
                Url = calendarUrl,
                Payload = reportBody,
                ContentType = "application/xml; charset=utf-8",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = new Dictionary<string, string> { { "Depth", "1" } }
            });

            if (!response.HasHttpResponse)
            {
                if (response.TransportException != null)
                {
                    HttpFailureInfo failure = response.FailureInfo ?? HttpFailureDiagnostics.Analyze(response.TransportException);
                    DiagnosticsLogger.LogException(LogCategory, "Free/busy REPORT request failed without HTTP response.", response.TransportException);
                    throw new FreeBusyRequestException("Free/busy request failed: " + failure.BuildUserMessage(), null, ShouldFallback(null));
                }

                throw new FreeBusyRequestException("Free/busy request failed: no HTTP response.", null, ShouldFallback(null));
            }
            string payloadText = response.ResponseText ?? string.Empty;
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                if (!string.IsNullOrEmpty(payloadText) &&
                    payloadText.IndexOf("BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return payloadText;
                }

                throw new FreeBusyRequestException(
                    "Free/busy request failed: HTTP " + (int)response.StatusCode,
                    response.StatusCode,
                    ShouldFallback(response.StatusCode),
                    payloadText);
            }
            if (string.IsNullOrEmpty(payloadText))
            {
                return null;
            }
            if (!string.IsNullOrEmpty(response.ContentType) &&
                response.ContentType.IndexOf("text/calendar", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return payloadText;
            }
            string extracted = ExtractCalendarData(payloadText);
            return string.IsNullOrEmpty(extracted) ? payloadText : extracted;
        }

        private string RequestFreeBusyViaScheduling(TalkServiceConfiguration configuration, string outboxUrl, string originatorEmail, string attendeeEmail, int days)
        {
            DateTime startUtc = DateTime.UtcNow;
            return CalDavFreeBusyService.RequestForParticipant(configuration, outboxUrl, originatorEmail, attendeeEmail, startUtc, startUtc.AddDays(days));
        }

        internal static string ExtractFreeBusyFromScheduleResponse(string responseText, string attendeeEmail)
        {
            return CalDavFreeBusyService.ExtractFreeBusyFromScheduleResponse(responseText, attendeeEmail);
        }

        private static string ExtractCalendarData(string payload)
        {
            try
            {
                var startTag = "<c:calendar-data";
                int start = payload.IndexOf(startTag, StringComparison.OrdinalIgnoreCase);
                if (start < 0)
                {
                    startTag = "<calendar-data";
                    start = payload.IndexOf(startTag, StringComparison.OrdinalIgnoreCase);
                }
                if (start < 0)
                {
                    return null;
                }
                int closing = payload.IndexOf('>', start);
                if (closing < 0)
                {
                    return null;
                }
                int end = payload.IndexOf("</", closing, StringComparison.OrdinalIgnoreCase);
                if (end < 0)
                {
                    return null;
                }
                string data = payload.Substring(closing + 1, end - closing - 1);
                return WebUtility.HtmlDecode(data).Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategory, "Failed to extract calendar-data from DAV payload.", ex);
                return null;
            }
        }

        private static bool ShouldFallback(HttpStatusCode? statusCode)
        {
            if (!statusCode.HasValue)
            {
                return true;
            }

            switch (statusCode.Value)
            {
                case HttpStatusCode.NotFound:
                case HttpStatusCode.Forbidden:
                case HttpStatusCode.Unauthorized:
                case HttpStatusCode.MethodNotAllowed:
                case HttpStatusCode.NotImplemented:
                    return true;
                default:
                    return false;
            }
        }

        private static void WriteError(HttpListenerContext context, HttpStatusCode statusCode, string message)
        {
            try
            {
                context.Response.StatusCode = (int)statusCode;
                context.Response.ContentType = "text/plain; charset=utf-8";
                using (var writer = new StreamWriter(context.Response.OutputStream, Utf8NoBom))
                {
                    writer.Write(message ?? string.Empty);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategory, "Failed to write error response (status=" + (int)statusCode + ").", ex);
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }

    internal sealed class FreeBusyRequestException : Exception
    {
        internal FreeBusyRequestException(string message, HttpStatusCode? statusCode, bool shouldFallback)
            : this(message, statusCode, shouldFallback, null)
        {
        }

        internal FreeBusyRequestException(string message, HttpStatusCode? statusCode, bool shouldFallback, string payload)
            : base(message)
        {
            StatusCode = statusCode;
            ShouldFallback = shouldFallback;
            Payload = payload;
        }

        internal HttpStatusCode? StatusCode { get; private set; }

        internal bool ShouldFallback { get; private set; }

        internal string Payload { get; private set; }
    }
}

