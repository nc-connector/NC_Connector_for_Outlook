/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Local HTTP endpoint for Outlook that proxies free/busy data to Nextcloud.
     */
    internal sealed class FreeBusyServer : IDisposable
    {
        private const string Prefix = "http://127.0.0.1:7777/nc-ifb/";
        private const string LogCategory = LogCategories.Ifb;
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        private readonly object _syncRoot = new object();
        private readonly IfbAddressBookCache _addressBookCache;

        private HttpListener _listener;
        private CancellationTokenSource _cancellation;
        private Task _listenerTask;

        private TalkServiceConfiguration _configuration;
        private int _defaultDays = 30;
        private int _cacheHours = 24;

        internal FreeBusyServer(IfbAddressBookCache addressBookCache)
        {
            if (addressBookCache == null)
            {
                throw new ArgumentNullException("addressBookCache");
            }

            _addressBookCache = addressBookCache;
        }

        internal void UpdateSettings(TalkServiceConfiguration configuration, int defaultDays, int cacheHours)
        {
            lock (_syncRoot)
            {
                _configuration = configuration;
                _defaultDays = defaultDays > 0 ? defaultDays : 30;
                _cacheHours = cacheHours >= 1 ? cacheHours : 24;
            }
        }

        internal void Start()
        {
            lock (_syncRoot)
            {
                if (_listener != null && _listener.IsListening)
                {
                    return;
                }

                if (_configuration == null || !_configuration.IsComplete())
                {
                    throw new InvalidOperationException("IFB cannot be started: credentials are incomplete.");
                }

                _listener = new HttpListener();
                _listener.Prefixes.Add(Prefix);
                _listener.Start();

                DiagnosticsLogger.Log(LogCategory, "Listener gestartet (Prefix=" + Prefix + ").");

                _cancellation = new CancellationTokenSource();
                _listenerTask = Task.Run(() => ListenLoop(_cancellation.Token));
            }
        }

        internal void Stop()
        {
            lock (_syncRoot)
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
                        DiagnosticsLogger.Log(LogCategory, "Listener wird gestoppt.");
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
                        DiagnosticsLogger.Log(LogCategory, "HTTP " + context.Request.HttpMethod + " " + context.Request.RawUrl);
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
                    Task.Run(() => HandleRequest(context), token);
                }
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            try
            {
                if (!string.Equals(context.Request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
                {
                    DiagnosticsLogger.Log(LogCategory, "Methode nicht erlaubt: " + context.Request.HttpMethod);
                    WriteError(context, HttpStatusCode.MethodNotAllowed, "Nur GET wird unterstuetzt.");
                    return;
                }

                string path = context.Request.Url.AbsolutePath ?? string.Empty;
                if (!path.StartsWith("/nc-ifb/freebusy/", StringComparison.OrdinalIgnoreCase) ||
                    !path.EndsWith(".vfb", StringComparison.OrdinalIgnoreCase))
                {
                    DiagnosticsLogger.Log(LogCategory, "Pfad nicht unterstuetzt: " + path);
                    WriteError(context, HttpStatusCode.NotFound, "Pfad nicht unterstuetzt.");
                    return;
                }

                string namePart = path.Substring("/nc-ifb/freebusy/".Length);
                namePart = namePart.Substring(0, namePart.Length - ".vfb".Length);
                string email = Uri.UnescapeDataString(namePart ?? string.Empty).Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(email))
                {
                    DiagnosticsLogger.Log(LogCategory, "Request ohne E-Mail verworfen.");
                    WriteError(context, HttpStatusCode.BadRequest, "E-Mail-Adresse fehlt.");
                    return;
                }

                int days = _defaultDays;
                string daysParam = context.Request.QueryString["days"];
                int parsedDays;
                if (!string.IsNullOrEmpty(daysParam) && int.TryParse(daysParam, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedDays))
                {
                    if (parsedDays > 0 && parsedDays <= 365)
                    {
                        days = parsedDays;
                    }
                }

                string resolvedEmail;
                if (!_addressBookCache.TryResolveEmail(_configuration, _cacheHours, email, out resolvedEmail))
                {
                    DiagnosticsLogger.Log(LogCategory, "E-Mail konnte nicht aufgeloest werden: " + email);
                    WriteError(context, HttpStatusCode.NotFound, "Benutzer nicht gefunden.");
                    return;
                }

                DiagnosticsLogger.Log(LogCategory, "IFB-Request: email=" + resolvedEmail + " (original=" + email + ") days=" + days);
                email = resolvedEmail;

                string uid;
                if (!_addressBookCache.TryGetUid(_configuration, _cacheHours, email, out uid) || string.IsNullOrEmpty(uid))
                {
                    DiagnosticsLogger.Log(LogCategory, "E-Mail nicht im Adressbuch gefunden: " + email);
                    WriteError(context, HttpStatusCode.NotFound, "Benutzer nicht gefunden.");
                    return;
                }

                string freeBusy = null;
                bool needsFallback = false;

                try
                {
                    freeBusy = RequestFreeBusyFromCalendar(uid, days);
                }
                catch (FreeBusyRequestException ex)
                {
                    DiagnosticsLogger.Log(LogCategory, "REPORT fehlgeschlagen: " + ex.Message);
                    if (ex.ShouldFallback)
                    {
                        needsFallback = true;
                    }
                    else
                    {
                        WriteError(context, ex.StatusCode ?? HttpStatusCode.InternalServerError, ex.Message);
                        return;
                    }
                }

                if (string.IsNullOrEmpty(freeBusy))
                {
                    needsFallback = true;
                }

                if (needsFallback)
                {
                    DiagnosticsLogger.Log(LogCategory, "Fallback via Scheduling fuer " + email + " gestartet.");
                    string originEmail;
                    if (!_addressBookCache.TryGetPrimaryEmailForUid(_configuration, _cacheHours, _configuration.Username, out originEmail) ||
                        string.IsNullOrEmpty(originEmail))
                    {
                        DiagnosticsLogger.Log(LogCategory, "Eigenes Benutzerkonto im Adressbuch nicht gefunden.");
                        WriteError(context, HttpStatusCode.PreconditionFailed, "Eigenes Benutzerkonto im Adressbuch nicht gefunden.");
                        return;
                    }

                    try
                    {
                        freeBusy = RequestFreeBusyViaScheduling(originEmail, email, days);
                    }
                    catch (FreeBusyRequestException ex)
                    {
                        DiagnosticsLogger.Log(LogCategory, "Scheduling fehlgeschlagen: " + ex.Message);
                        WriteError(context, ex.StatusCode ?? HttpStatusCode.InternalServerError, ex.Message);
                        return;
                    }
                }

                if (string.IsNullOrEmpty(freeBusy))
                {
                    DiagnosticsLogger.Log(LogCategory, "Keine Frei/Gebucht-Daten verfuegbar fuer " + email);
                    WriteError(context, HttpStatusCode.NoContent, "Keine Frei/Gebucht-Daten verfuegbar.");
                    return;
                }

                DiagnosticsLogger.Log(LogCategory, "Antwort ok fuer " + email + " (Len=" + freeBusy.Length + ").");

                context.Response.StatusCode = (int)HttpStatusCode.OK;
                context.Response.ContentType = "text/calendar; charset=utf-8";
                using (var writer = new StreamWriter(context.Response.OutputStream, Utf8NoBom))
                {
                    writer.Write(freeBusy);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log(LogCategory, "Fehler bei HandleRequest: " + ex);
                WriteError(context, HttpStatusCode.InternalServerError, ex.Message);
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

        private string RequestFreeBusyFromCalendar(string uid, int days)
        {
            DateTime startUtc = DateTime.UtcNow;
            DateTime endUtc = startUtc.AddDays(days);

            string baseUrl = _configuration.GetNormalizedBaseUrl();
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

            HttpWebResponse response = null;
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(calendarUrl);
                request.Method = "REPORT";
                request.ContentType = "application/xml; charset=utf-8";
                request.Headers["Depth"] = "1";
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Timeout = 60000;

                using (var stream = request.GetRequestStream())
                using (var writer = new StreamWriter(stream, Utf8NoBom))
                {
                    writer.Write(reportBody);
                }

                response = (HttpWebResponse)request.GetResponse();
                using (var responseStream = response.GetResponseStream())
                using (var reader = new StreamReader(responseStream ?? Stream.Null, Encoding.UTF8))
                {
                    string payload = reader.ReadToEnd();
                    if (string.IsNullOrEmpty(payload))
                    {
                        return null;
                    }

                    if (!string.IsNullOrEmpty(response.ContentType) &&
                        response.ContentType.IndexOf("text/calendar", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return payload;
                    }

                    string extracted = ExtractCalendarData(payload);
                    return string.IsNullOrEmpty(extracted) ? payload : extracted;
                }
            }
            catch (WebException ex)
            {
                HttpStatusCode? statusCode = null;
                HttpWebResponse errorResponse = ex.Response as HttpWebResponse;
                if (errorResponse != null)
                {
                    statusCode = errorResponse.StatusCode;
                    using (var responseStream = errorResponse.GetResponseStream())
                    using (var reader = new StreamReader(responseStream ?? Stream.Null, Encoding.UTF8))
                    {
                        string payload = reader.ReadToEnd();
                        if (!string.IsNullOrEmpty(payload) &&
                            payload.IndexOf("BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return payload;
                        }
                    }
                }

                DiagnosticsLogger.LogException(LogCategory, "Free/busy REPORT request failed (status=" + (statusCode.HasValue ? ((int)statusCode.Value).ToString(CultureInfo.InvariantCulture) : "unknown") + ").", ex);
                throw new FreeBusyRequestException("Free/busy request failed: " + ex.Message, statusCode, ShouldFallback(statusCode));
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private string RequestFreeBusyViaScheduling(string originatorEmail, string attendeeEmail, int days)
        {
            if (string.IsNullOrEmpty(originatorEmail) || string.IsNullOrEmpty(attendeeEmail))
            {
                throw new FreeBusyRequestException("Invalid email address for scheduling.", HttpStatusCode.BadRequest, false);
            }

            DateTime startUtc = DateTime.UtcNow;
            DateTime endUtc = startUtc.AddDays(days);

            string baseUrl = _configuration.GetNormalizedBaseUrl();
            string outboxUrl = string.Format(
                CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/calendars/{1}/outbox/",
                baseUrl,
                Uri.EscapeDataString(_configuration.Username));

            string uid = Guid.NewGuid().ToString();
            string payload = BuildVFreeBusyRequest(uid, originatorEmail, attendeeEmail, startUtc, endUtc);

            HttpWebResponse response = null;
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(outboxUrl);
                request.Method = "POST";
                request.ContentType = "text/calendar; charset=\"UTF-8\"";
                request.Headers["Depth"] = "0";
                request.Headers["Authorization"] = HttpAuthUtilities.BuildBasicAuthHeader(_configuration.Username, _configuration.AppPassword);
                request.Headers["Originator"] = "mailto:" + originatorEmail;
                request.Headers["Recipient"] = "mailto:" + attendeeEmail;
                request.Timeout = 60000;

                using (var stream = request.GetRequestStream())
                using (var writer = new StreamWriter(stream, Utf8NoBom))
                {
                    writer.Write(payload);
                }

                response = (HttpWebResponse)request.GetResponse();
                using (var responseStream = response.GetResponseStream())
                using (var reader = new StreamReader(responseStream ?? Stream.Null, Encoding.UTF8))
                {
                    string responseText = reader.ReadToEnd();
                    return ExtractFreeBusyFromScheduleResponse(responseText, attendeeEmail);
                }
            }
            catch (WebException ex)
            {
                HttpStatusCode? statusCode = null;
                HttpWebResponse errorResponse = ex.Response as HttpWebResponse;
                string responseText = null;
                if (errorResponse != null)
                {
                    statusCode = errorResponse.StatusCode;
                    using (var responseStream = errorResponse.GetResponseStream())
                    using (var reader = new StreamReader(responseStream ?? Stream.Null, Encoding.UTF8))
                    {
                        responseText = reader.ReadToEnd();
                    }
                }

                DiagnosticsLogger.LogException(LogCategory, "CalDAV scheduling request failed (status=" + (statusCode.HasValue ? ((int)statusCode.Value).ToString(CultureInfo.InvariantCulture) : "unknown") + ").", ex);
                throw new FreeBusyRequestException("CalDAV scheduling failed: " + ex.Message, statusCode, false, responseText);
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static string BuildVFreeBusyRequest(string uid, string originator, string attendee, DateTime start, DateTime end)
        {
            var builder = new StringBuilder();
            builder.AppendLine("BEGIN:VCALENDAR");
            builder.AppendLine("PRODID:-//Nextcloud Talk Direkt//IFB//DE");
            builder.AppendLine("CALSCALE:GREGORIAN");
            builder.AppendLine("VERSION:2.0");
            builder.AppendLine("METHOD:REQUEST");
            builder.AppendLine("BEGIN:VFREEBUSY");
            builder.AppendLine("DTSTAMP:" + start.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture));
            builder.AppendLine("UID:" + uid);
            builder.AppendLine("DTSTART:" + start.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture));
            builder.AppendLine("DTEND:" + end.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture));
            builder.AppendLine("ORGANIZER;CN=" + EscapeVCalendarText(originator) + ":mailto:" + originator);
            builder.AppendLine("ATTENDEE;CN=" + EscapeVCalendarText(attendee) + ":mailto:" + attendee);
            builder.AppendLine("END:VFREEBUSY");
            builder.AppendLine("END:VCALENDAR");
            return builder.ToString();
        }

        private static string ExtractFreeBusyFromScheduleResponse(string responseText, string attendeeEmail)
        {
            if (string.IsNullOrEmpty(responseText))
            {
                return null;
            }

            try
            {
                int index = 0;
                while (index >= 0 && index < responseText.Length)
                {
                    int startTag = responseText.IndexOf("<cal:calendar-data", index, StringComparison.OrdinalIgnoreCase);
                    string endToken = "</cal:calendar-data>";
                    if (startTag < 0)
                    {
                        startTag = responseText.IndexOf("<calendar-data", index, StringComparison.OrdinalIgnoreCase);
                        endToken = "</calendar-data>";
                    }

                    if (startTag < 0)
                    {
                        break;
                    }

                    int startContent = responseText.IndexOf('>', startTag);
                    if (startContent < 0)
                    {
                        break;
                    }
                    startContent++;

                    int endTag = responseText.IndexOf(endToken, startContent, StringComparison.OrdinalIgnoreCase);
                    if (endTag < 0)
                    {
                        break;
                    }

                    string raw = responseText.Substring(startContent, endTag - startContent);
                    string decoded = WebUtility.HtmlDecode(raw ?? string.Empty).Trim();

                    if (!string.IsNullOrEmpty(decoded) &&
                        decoded.IndexOf(attendeeEmail ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        decoded = NormalizeFreeBusyPayload(decoded);
                        return decoded;
                    }

                    index = endTag + endToken.Length;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategory, "Failed to extract free/busy data from scheduling response.", ex);
            }

            return null;
        }

        private static string NormalizeFreeBusyPayload(string payload)
        {
            if (string.IsNullOrEmpty(payload))
            {
                return string.Empty;
            }

            string[] lines = payload.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            var builder = new StringBuilder();
            bool methodSet = false;

            foreach (string rawLine in lines)
            {
                string line = rawLine ?? string.Empty;
                if (line.StartsWith("METHOD:", StringComparison.OrdinalIgnoreCase))
                {
                    builder.AppendLine("METHOD:PUBLISH");
                    methodSet = true;
                }
                else if (line.Length > 0)
                {
                    builder.AppendLine(line);
                }
            }

            if (!methodSet)
            {
                builder.Insert(0, "METHOD:PUBLISH" + Environment.NewLine);
            }

            string result = builder.ToString().Trim();
            if (!result.StartsWith("BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase))
            {
                result = "BEGIN:VCALENDAR" + Environment.NewLine + result;
            }
            if (!result.EndsWith("END:VCALENDAR", StringComparison.OrdinalIgnoreCase))
            {
                result = result + Environment.NewLine + "END:VCALENDAR";
            }

            return result;
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

        private static string EscapeVCalendarText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace("\\", "\\\\").Replace(";", "\\;").Replace(",", "\\,").Replace("\n", "\\n");
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
