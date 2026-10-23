// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Shared scheduling transport for IFB and the appointment-scoped resource picker.
    internal sealed class CalDavFreeBusyService
    {
        private readonly CalendarResourceDirectory _directory;
        private readonly TalkServiceConfiguration _configuration;
        private readonly int _cacheHours;
        private readonly SemaphoreSlim _slots = new SemaphoreSlim(2, 2);
        private readonly object _sync = new object();
        private readonly Dictionary<string, CachedAvailability> _cache = new Dictionary<string, CachedAvailability>(StringComparer.OrdinalIgnoreCase);
        private string _originator, _outbox;
        private long _querySequence;

        private sealed class CachedAvailability
        {
            internal CalendarResourceAvailability Value;
            internal DateTime ExpiresUtc;
            internal long Sequence;
        }

        internal CalDavFreeBusyService(CalendarResourceDirectory directory, TalkServiceConfiguration configuration, int cacheHours)
        {
            _directory = directory;
            _configuration = configuration;
            _cacheHours = cacheHours;
        }

        internal async Task<Dictionary<string, CalendarResourceAvailability>> QueryAsync(IEnumerable<CalendarResource> resources,
            DateTime startUtc, DateTime endUtc, CancellationToken token)
        {
            if (startUtc.Kind != DateTimeKind.Utc || endUtc.Kind != DateTimeKind.Utc || endUtc <= startUtc)
            { throw new ArgumentException("A positive UTC interval is required."); }
            List<CalendarResource> requested = resources.Where(r => r.CanSchedule)
                .GroupBy(r => r.Email, StringComparer.OrdinalIgnoreCase).Select(g => g.First()).ToList();
            long sequence = Interlocked.Increment(ref _querySequence);
            await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();
                var allowed = new HashSet<string>(_directory.GetResources(_configuration, _cacheHours).Where(r => r.CanSchedule).Select(r => r.Email), StringComparer.OrdinalIgnoreCase);
                if (requested.Any(r => !allowed.Contains(r.Email))) { throw new InvalidDataException("Resource is not in the visible directory."); }
                if (requested.Count == 0) { return; }
                lock (_sync)
                {
                    if (_outbox == null) { _directory.GetSchedulingIdentity(_configuration, out _originator, out _outbox); }
                }
                token.ThrowIfCancellationRequested();
            }, token).ConfigureAwait(false);
            var result = new Dictionary<string, CalendarResourceAvailability>(StringComparer.OrdinalIgnoreCase);
            var missing = new List<string>();
            lock (_sync)
            {
                foreach (CalendarResource resource in requested)
                {
                    CachedAvailability cached;
                    if (_cache.TryGetValue(resource.Email, out cached) && cached.ExpiresUtc > DateTime.UtcNow
                        && cached.Value.StartUtc <= startUtc && cached.Value.EndUtc >= endUtc)
                    { result[resource.Email] = cached.Value.ForInterval(startUtc, endUtc); }
                    else { missing.Add(resource.Email); }
                }
            }
            DateTime requestStart = new DateTime(startUtc.Ticks - startUtc.Ticks % TimeSpan.TicksPerSecond, DateTimeKind.Utc);
            DateTime requestEnd = new DateTime(endUtc.Ticks - endUtc.Ticks % TimeSpan.TicksPerSecond, DateTimeKind.Utc);
            if (requestEnd < endUtc) { requestEnd = requestEnd.AddSeconds(1); }
            var tasks = new List<Task>();
            for (int offset = 0; offset < missing.Count; offset += 50)
            {
                string[] batch = missing.Skip(offset).Take(50).ToArray();
                tasks.Add(QueryBatchAsync(batch, requestStart, requestEnd, startUtc, endUtc, sequence, result, token));
            }
            await Task.WhenAll(tasks).ConfigureAwait(false);
            await Task.Run(() => _directory.GetResources(_configuration, _cacheHours), token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            DiagnosticsLogger.Log(LogCategories.Ifb, "Picker availability completed (resources=" + requested.Count + ", queried=" + missing.Count + ").");
            return result;
        }

        private async Task QueryBatchAsync(string[] emails, DateTime requestStart, DateTime requestEnd,
            DateTime start, DateTime end, long sequence, Dictionary<string, CalendarResourceAvailability> results, CancellationToken token)
        {
            await _slots.WaitAsync(token).ConfigureAwait(false);
            try
            {
                Dictionary<string, CalendarResourceAvailability> batch = await Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();
                    // Recheck before each queued request; a cached success cannot override a Seat refusal.
                    _directory.GetResources(_configuration, _cacheHours);
                    string response = SendSchedulingRequest(_configuration, _outbox, _originator, emails, requestStart, requestEnd, token);
                    return ParseAvailabilityResponse(response, emails, requestStart, requestEnd);
                }, token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                lock (_sync)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (string email in emails)
                    {
                        CalendarResourceAvailability value = batch[email];
                        CacheResult(value, sequence);
                        lock (results) { results[email] = value.ForInterval(start, end); }
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                token.ThrowIfCancellationRequested();
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource availability batch failed.", ex);
                lock (_sync)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (string email in emails)
                    {
                        var unknown = CalendarResourceAvailability.Unknown(email, start, end);
                        CacheResult(unknown, sequence);
                        lock (results) { results[email] = unknown; }
                    }
                }
            }
            finally { _slots.Release(); }
        }

        // Called under _sync; failed results also retain their sequence so older requests cannot revive them.
        private void CacheResult(CalendarResourceAvailability value, long sequence)
        {
            CachedAvailability current;
            if (_cache.TryGetValue(value.Email, out current) && current.Sequence > sequence) { return; }
            _cache[value.Email] = new CachedAvailability
            {
                Value = value, Sequence = sequence,
                ExpiresUtc = value.IsKnown ? DateTime.UtcNow.AddMinutes(2) : DateTime.MinValue
            };
        }

        internal static string RequestForParticipant(TalkServiceConfiguration configuration, string outbox, string originator,
            string attendee, DateTime startUtc, DateTime endUtc)
        {
            string response = SendSchedulingRequest(configuration, outbox, originator, new[] { attendee }, startUtc, endUtc, CancellationToken.None);
            return ExtractFreeBusyFromScheduleResponse(response, attendee);
        }

        private static string SendSchedulingRequest(TalkServiceConfiguration configuration, string outbox, string originator,
            IEnumerable<string> attendees, DateTime startUtc, DateTime endUtc, CancellationToken token)
        {
            originator = CalendarResourceDirectory.NormalizeEmail(originator);
            string[] addresses = attendees.Select(CalendarResourceDirectory.NormalizeEmail).ToArray();
            if (originator.Length == 0 || addresses.Length == 0 || addresses.Any(a => a.Length == 0))
            { throw new InvalidDataException("Invalid scheduling address."); }
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            var payload = new StringBuilder();
            payload.Append("BEGIN:VCALENDAR\r\nPRODID:-//NC Connector//Availability//EN\r\nVERSION:2.0\r\nMETHOD:REQUEST\r\nBEGIN:VFREEBUSY\r\n");
            payload.Append("UID:").Append(Guid.NewGuid().ToString("N")).Append("\r\n");
            payload.Append("DTSTAMP:").Append(DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
            payload.Append("DTSTART:").Append(startUtc.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
            payload.Append("DTEND:").Append(endUtc.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
            payload.Append("ORGANIZER:mailto:").Append(originator).Append("\r\n");
            foreach (string email in addresses) { payload.Append("ATTENDEE:mailto:").Append(email).Append("\r\n"); }
            payload.Append("END:VFREEBUSY\r\nEND:VCALENDAR\r\n");
            NcHttpResponse response = new NcHttpClient(configuration).Send(new NcHttpRequestOptions
            {
                Method = "POST", Url = outbox, Payload = payload.ToString(), ContentType = "text/calendar; charset=UTF-8",
                IncludeAuthHeader = true, IncludeOcsApiHeader = false, ParseJson = false, MaximumResponseBytes = 4 * 1024 * 1024,
                TimeoutMs = 30000, CancellationToken = token,
                Headers = new Dictionary<string, string> { { "Originator", "mailto:" + originator },
                    { "Recipient", string.Join(", ", addresses.Select(a => "mailto:" + a)) }, { "Depth", "0" } }
            });
            token.ThrowIfCancellationRequested();
            if (!response.HasHttpResponse || (int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            { throw new FreeBusyRequestException("CalDAV scheduling request failed.", response.HasHttpResponse ? (HttpStatusCode?)response.StatusCode : null, false); }
            return response.ResponseText ?? string.Empty;
        }

        internal static Dictionary<string, CalendarResourceAvailability> ParseAvailabilityResponse(string response,
            IEnumerable<string> emails, DateTime start, DateTime end)
        {
            XmlElement root = CalendarResourceDirectory.ReadXml(response).DocumentElement;
            if (root == null || root.LocalName != "schedule-response" || root.NamespaceURI != CalendarResourceDirectory.CalDavNamespace)
            { throw new InvalidDataException("Invalid scheduling response."); }
            var results = new Dictionary<string, CalendarResourceAvailability>(StringComparer.OrdinalIgnoreCase);
            foreach (string email in emails)
            {
                try { results[email] = ParseAvailability(ExtractFreeBusy(root, email), email, start, end); }
                catch (InvalidDataException) { results[email] = CalendarResourceAvailability.Unknown(email, start, end); }
                catch (FreeBusyRequestException) { results[email] = CalendarResourceAvailability.Unknown(email, start, end); }
                catch (FormatException) { results[email] = CalendarResourceAvailability.Unknown(email, start, end); }
                catch (ArgumentOutOfRangeException) { results[email] = CalendarResourceAvailability.Unknown(email, start, end); }
                catch (OverflowException) { results[email] = CalendarResourceAvailability.Unknown(email, start, end); }
            }
            return results;
        }

        private static CalendarResourceAvailability ParseAvailability(string payload, string email, DateTime requestedStart, DateTime requestedEnd)
        {
            DateTime start = DateTime.MinValue, end = DateTime.MinValue;
            var busy = new List<CalendarBusyPeriod>();
            foreach (string line in payload.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                int separator = CalendarPropertySeparator(line);
                if (separator < 0) { continue; }
                string property = line.Substring(0, separator);
                string value = line.Substring(separator + 1);
                string name = property.Split(';')[0];
                if (name.Equals("DTSTART", StringComparison.OrdinalIgnoreCase)) { start = ParseCalendarUtc(value); }
                if (name.Equals("DTEND", StringComparison.OrdinalIgnoreCase)) { end = ParseCalendarUtc(value); }
                if (!name.Equals("FREEBUSY", StringComparison.OrdinalIgnoreCase)) { continue; }
                if (IsFreePeriod(property)) { continue; }
                foreach (string period in value.Split(','))
                {
                    string[] limits = period.Split('/');
                    DateTime from = ParseCalendarUtc(limits[0]);
                    DateTime to = limits[1].StartsWith("P", StringComparison.Ordinal) ? from.Add(XmlConvert.ToTimeSpan(limits[1])) : ParseCalendarUtc(limits[1]);
                    if (from < requestedEnd && to > requestedStart)
                    { busy.Add(new CalendarBusyPeriod(from < requestedStart ? requestedStart : from, to > requestedEnd ? requestedEnd : to)); }
                }
            }
            if (start > requestedStart || end < requestedEnd) { return CalendarResourceAvailability.Unknown(email, requestedStart, requestedEnd); }
            return new CalendarResourceAvailability(email, requestedStart, requestedEnd, true, busy.OrderBy(p => p.StartUtc));
        }

        internal static string ExtractFreeBusyFromScheduleResponse(string responseText, string attendeeEmail)
        {
            return ExtractFreeBusy(CalendarResourceDirectory.ReadXml(responseText).DocumentElement, attendeeEmail);
        }

        private static string ExtractFreeBusy(XmlElement root, string attendeeEmail)
        {
            string expected = CalendarResourceDirectory.NormalizeEmail(attendeeEmail);
            const string cal = CalendarResourceDirectory.CalDavNamespace;
            if (expected.Length == 0 || root == null || root.LocalName != "schedule-response" || root.NamespaceURI != cal)
            {
                throw new InvalidDataException("Invalid scheduling response.");
            }
            string result = null;
            foreach (XmlNode node in root.ChildNodes)
            {
                XmlElement response = node as XmlElement;
                if (response == null) { continue; }
                if (response.LocalName != "response" || response.NamespaceURI != cal) { continue; }
                string href;
                try
                {
                    XmlElement recipient = CalendarResourceDirectory.Child(response, "recipient", cal);
                    href = recipient == null ? string.Empty : CalendarResourceDirectory.ChildText(recipient, "href", CalendarResourceDirectory.DavNamespace);
                }
                catch (InvalidDataException ex)
                {
                    // An ambiguous envelope affects only addresses it claims, not the other batch recipients.
                    foreach (XmlElement candidate in response.GetElementsByTagName("href", CalendarResourceDirectory.DavNamespace))
                    {
                        string address = candidate.InnerText.Trim();
                        if (address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                            && string.Equals(CalendarResourceDirectory.NormalizeEmail(address.Substring(7)), expected, StringComparison.OrdinalIgnoreCase)) { throw; }
                    }
                    DiagnosticsLogger.Log(LogCategories.Ifb, "Skipped invalid scheduling recipient: " + ex.Message);
                    continue;
                }
                string email = href.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? CalendarResourceDirectory.NormalizeEmail(href.Substring(7)) : string.Empty;
                if (!string.Equals(email, expected, StringComparison.OrdinalIgnoreCase)) { continue; }
                if (result != null) { throw new InvalidDataException("Duplicate scheduling recipient."); }
                string status = CalendarResourceDirectory.ChildText(response, "request-status", cal);
                if (!string.Equals(status.Split(';')[0].Trim(), "2.0", StringComparison.Ordinal))
                {
                    throw new FreeBusyRequestException("Calendar participant availability could not be retrieved.", HttpStatusCode.BadGateway, false);
                }
                string payload = CalendarResourceDirectory.ChildText(response, "calendar-data", cal);
                result = NormalizeFreeBusyPayload(payload, expected);
            }
            if (result == null) { throw new InvalidDataException("Scheduling response did not contain the requested participant."); }
            return result;
        }

        private static string NormalizeFreeBusyPayload(string payload, string expectedEmail)
        {
            var lines = new List<string>();
            foreach (string line in (payload ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None))
            {
                if (line.Length == 0) { continue; }
                if (line[0] == ' ' || line[0] == '\t')
                {
                    if (lines.Count == 0) { throw new InvalidDataException("Invalid folded calendar data."); }
                    lines[lines.Count - 1] += line.Substring(1);
                }
                else { lines.Add(line); }
            }
            int state = 0;
            bool freeBusySeen = false;
            bool attendeeSeen = false;
            var required = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            DateTime start = DateTime.MinValue;
            DateTime end = DateTime.MinValue;
            var builder = new StringBuilder();
            foreach (string line in lines)
            {
                if (string.Equals(line, "BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase) && state == 0)
                {
                    state = 1;
                    builder.Append("BEGIN:VCALENDAR\r\nMETHOD:PUBLISH\r\n");
                    continue;
                }
                if (string.Equals(line, "BEGIN:VFREEBUSY", StringComparison.OrdinalIgnoreCase) && state == 1 && !freeBusySeen)
                {
                    state = 2;
                    freeBusySeen = true;
                }
                else if (string.Equals(line, "END:VFREEBUSY", StringComparison.OrdinalIgnoreCase) && state == 2) { state = 1; }
                else if (string.Equals(line, "END:VCALENDAR", StringComparison.OrdinalIgnoreCase) && state == 1 && freeBusySeen) { state = 3; }
                else
                {
                    int separator = CalendarPropertySeparator(line);
                    if (separator <= 0 || state == 0 || state == 3) { throw new InvalidDataException("Invalid free/busy calendar structure."); }
                    string name = line.Substring(0, separator).Split(';')[0];
                    string value = line.Substring(separator + 1);
                    if (name.Equals("BEGIN", StringComparison.OrdinalIgnoreCase) || name.Equals("END", StringComparison.OrdinalIgnoreCase)) { throw new InvalidDataException("Unexpected calendar component."); }
                    if (name.Equals("METHOD", StringComparison.OrdinalIgnoreCase))
                    {
                        if (state != 1) { throw new InvalidDataException("Invalid calendar method position."); }
                        continue;
                    }
                    if (name.Equals("VERSION", StringComparison.OrdinalIgnoreCase))
                    {
                        if (state != 1 || value != "2.0" || !required.Add("VERSION")) { throw new InvalidDataException("Invalid calendar version."); }
                    }
                    else if (name.Equals("ATTENDEE", StringComparison.OrdinalIgnoreCase))
                    {
                        if (state != 2 || !value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                            || !string.Equals(CalendarResourceDirectory.NormalizeEmail(value.Substring(7)), expectedEmail, StringComparison.OrdinalIgnoreCase))
                        { throw new InvalidDataException("Calendar attendee did not match the requested participant."); }
                        attendeeSeen = true;
                    }
                    else if (name.Equals("UID", StringComparison.OrdinalIgnoreCase) || name.Equals("DTSTAMP", StringComparison.OrdinalIgnoreCase)
                        || name.Equals("DTSTART", StringComparison.OrdinalIgnoreCase) || name.Equals("DTEND", StringComparison.OrdinalIgnoreCase))
                    {
                        if (state != 2 || value.Length == 0 || !required.Add(name)) { throw new InvalidDataException("Invalid free/busy identity or time range."); }
                        if (!name.Equals("UID", StringComparison.OrdinalIgnoreCase))
                        {
                            DateTime date = ParseCalendarUtc(value);
                            if (name.Equals("DTSTART", StringComparison.OrdinalIgnoreCase)) { start = date; }
                            if (name.Equals("DTEND", StringComparison.OrdinalIgnoreCase)) { end = date; }
                        }
                    }
                    else if (name.Equals("FREEBUSY", StringComparison.OrdinalIgnoreCase))
                    {
                        if (state != 2 || value.Length == 0) { throw new InvalidDataException("Invalid busy periods."); }
                        IsFreePeriod(line.Substring(0, separator));
                        foreach (string period in value.Split(','))
                        {
                            string[] limits = period.Split('/');
                            if (limits.Length != 2) { throw new InvalidDataException("Invalid busy period."); }
                            DateTime from = ParseCalendarUtc(limits[0]);
                            if (limits[1].StartsWith("P", StringComparison.Ordinal))
                            {
                                if (XmlConvert.ToTimeSpan(limits[1]) <= TimeSpan.Zero) { throw new InvalidDataException("Invalid busy duration."); }
                            }
                            else if (ParseCalendarUtc(limits[1]) <= from) { throw new InvalidDataException("Invalid busy period range."); }
                        }
                    }
                }
                builder.Append(line).Append("\r\n");
            }
            // A successful, complete VFREEBUSY with no periods means free; malformed data never does.
            if (state != 3 || !attendeeSeen || required.Count != 5 || end <= start)
            { throw new InvalidDataException("Incomplete free/busy calendar."); }
            return builder.ToString();
        }

        private static bool IsFreePeriod(string property)
        {
            string type = null;
            bool quoted = false;
            int start = property.IndexOf(';') + 1;
            if (start == 0) { return false; }
            for (int index = start; index <= property.Length; index++)
            {
                if (index < property.Length && property[index] == '"') { quoted = !quoted; }
                if (index < property.Length && (property[index] != ';' || quoted)) { continue; }
                if (quoted) { throw new InvalidDataException("Unclosed free/busy parameter quote."); }
                string parameter = property.Substring(start, index - start);
                int equals = parameter.IndexOf('=');
                if (equals <= 0 || equals == parameter.Length - 1) { throw new InvalidDataException("Invalid free/busy parameter."); }
                if (parameter.Substring(0, equals).Equals("FBTYPE", StringComparison.OrdinalIgnoreCase))
                {
                    if (type != null) { throw new InvalidDataException("Duplicate free/busy type."); }
                    type = parameter.Substring(equals + 1);
                    if (type.Any(c => !(c >= 'A' && c <= 'Z') && !(c >= 'a' && c <= 'z') && !(c >= '0' && c <= '9') && c != '-'))
                    { throw new InvalidDataException("Invalid free/busy type."); }
                }
                start = index + 1;
            }
            // RFC 5545 section 3.2.9: absent or unrecognized types count as busy.
            return string.Equals(type, "FREE", StringComparison.OrdinalIgnoreCase);
        }

        private static int CalendarPropertySeparator(string line)
        {
            bool quoted = false;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"') { quoted = !quoted; }
                if (line[i] == ':' && !quoted) { return i; }
            }
            return -1;
        }

        private static DateTime ParseCalendarUtc(string value)
        {
            DateTime result;
            if (!DateTime.TryParseExact(value, "yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out result))
            { throw new InvalidDataException("Invalid UTC calendar time."); }
            return result;
        }

    }
}

