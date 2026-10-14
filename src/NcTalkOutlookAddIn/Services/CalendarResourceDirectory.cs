// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Keeps the access-filtered resource directory separate from the Talk user address book.
    internal sealed class CalendarResourceDirectory
    {
        internal const string DavNamespace = "DAV:";
        internal const string CalDavNamespace = "urn:ietf:params:xml:ns:caldav";
        private const string NextcloudNamespace = "http://nextcloud.com/ns";
        private const string SabreNamespace = "http://sabredav.org/ns";
        private const string Namespaces = " xmlns:d=\"DAV:\" xmlns:c=\"urn:ietf:params:xml:ns:caldav\" xmlns:n=\"http://nextcloud.com/ns\" xmlns:s=\"http://sabredav.org/ns\"";
        private const string ResourceProperties = "<d:displayname/><c:calendar-user-type/><c:calendar-user-address-set/><s:email-address/><n:resource-type/><n:room-type/><n:room-seating-capacity/><n:room-building-name/><n:room-building-address/><n:room-building-story/><n:room-building-room-number/><n:room-features/>";
        private readonly object _syncRoot = new object();
        private readonly Func<TalkServiceConfiguration, BackendPolicyStatus> _fetchPolicy;
        private string _scope = string.Empty;
        private DateTime _generatedUtc = DateTime.MinValue;
        private bool _refreshFailed;
        private List<CalendarResource> _resources = new List<CalendarResource>();

        internal CalendarResourceDirectory(Func<TalkServiceConfiguration, BackendPolicyStatus> fetchPolicy)
        {
            _fetchPolicy = fetchPolicy;
        }

        internal List<CalendarResource> GetResources(TalkServiceConfiguration configuration, int cacheHours, bool forceRefresh = false)
        {
            RequireAccess(configuration);
            lock (_syncRoot)
            {
                string scope = BuildScope(configuration);
                if (!string.Equals(_scope, scope, StringComparison.Ordinal))
                {
                    _scope = scope;
                    _generatedUtc = DateTime.MinValue;
                    _refreshFailed = false;
                    _resources = new List<CalendarResource>();
                }
                if (forceRefresh || _refreshFailed || _generatedUtc.AddHours(Math.Max(1, Math.Min(cacheHours, 168))) <= DateTime.UtcNow)
                {
                    using (DiagnosticsLogger.BeginOperation(LogCategories.Ifb, "Read resource directory"))
                    {
                        try
                        {
                            var resources = ReadCollection(configuration, "calendar-rooms", "ROOM");
                            resources.AddRange(ReadCollection(configuration, "calendar-resources", "RESOURCE"));
                            var addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            foreach (CalendarResource resource in resources)
                            {
                                if (!addresses.Add(resource.Email)) { throw new InvalidDataException("Ambiguous resource email address."); }
                            }
                            _resources = resources;
                            _generatedUtc = DateTime.UtcNow;
                            _refreshFailed = false;
                            DiagnosticsLogger.Log(LogCategories.Ifb, "Resource directory refreshed (count=" + resources.Count.ToString(CultureInfo.InvariantCulture) + ").");
                        }
                        catch (Exception ex)
                        {
                            _refreshFailed = true;
                            DiagnosticsLogger.LogException(LogCategories.Ifb, "Resource directory refresh failed.", ex);
                            throw;
                        }
                    }
                }
                return _resources.ConvertAll(resource => resource.Clone());
            }
        }

        internal CalendarResource FindByEmail(TalkServiceConfiguration configuration, int cacheHours, string email)
        {
            string normalized = NormalizeEmail(email);
            foreach (CalendarResource resource in GetResources(configuration, cacheHours))
            {
                if (string.Equals(resource.Email, normalized, StringComparison.OrdinalIgnoreCase)) { return resource; }
            }
            return null;
        }

        private void RequireAccess(TalkServiceConfiguration configuration)
        {
            if (configuration == null || !configuration.IsComplete()) { throw new InvalidOperationException("Nextcloud credentials are incomplete."); }
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            BackendPolicyStatus status = _fetchPolicy != null ? _fetchPolicy(configuration) : null;
            if (status == null || !status.FetchSucceeded || !PolicyUiHelper.HasBackendSeatEntitlement(status))
            {
                throw new FreeBusyRequestException("Resource planning requires a valid active assigned Seat.", HttpStatusCode.Forbidden, false);
            }
        }

        private static List<CalendarResource> ReadCollection(TalkServiceConfiguration configuration, string collection, string kind)
        {
            string url = configuration.GetNormalizedBaseUrl() + "/remote.php/dav/principals/" + collection + "/";
            // Email search applies Nextcloud group restrictions; depth-one PROPFIND does not.
            string body = "<d:principal-property-search" + Namespaces + "><d:property-search><d:prop><s:email-address/></d:prop><d:match>@</d:match></d:property-search><d:prop>" + ResourceProperties + "</d:prop></d:principal-property-search>";
            XmlDocument document = SendDav(configuration, "REPORT", url, body);
            var resources = new List<CalendarResource>();
            var paths = new HashSet<string>(StringComparer.Ordinal);
            foreach (XmlElement response in GetResponses(document))
            {
                string href = ChildText(response, "href", DavNamespace);
                Uri principal = ValidateDavHref(configuration, url, href, "principals/" + collection + "/", true);
                if (!paths.Add(principal.AbsolutePath)) { throw new InvalidDataException("Duplicate resource principal."); }
                Dictionary<string, XmlElement> properties = ReadSuccessfulProperties(response);
                string responseKind = PropertyText(properties, "calendar-user-type", CalDavNamespace);
                if (!string.Equals(responseKind, kind, StringComparison.Ordinal)) { throw new InvalidDataException("Unexpected resource principal kind."); }
                string email = ReadEmail(properties);
                if (email.Length == 0) { throw new InvalidDataException("Resource principal has no usable email address."); }
                string capacityText = PropertyText(properties, "room-seating-capacity", NextcloudNamespace);
                int capacity;
                resources.Add(new CalendarResource
                {
                    PrincipalPath = principal.AbsolutePath,
                    Email = email,
                    DisplayName = PropertyText(properties, "displayname", DavNamespace),
                    Kind = kind,
                    ResourceType = PropertyText(properties, "resource-type", NextcloudNamespace),
                    RoomType = PropertyText(properties, "room-type", NextcloudNamespace),
                    BuildingName = PropertyText(properties, "room-building-name", NextcloudNamespace),
                    BuildingAddress = PropertyText(properties, "room-building-address", NextcloudNamespace),
                    Floor = PropertyText(properties, "room-building-story", NextcloudNamespace),
                    RoomNumber = PropertyText(properties, "room-building-room-number", NextcloudNamespace),
                    Features = PropertyText(properties, "room-features", NextcloudNamespace),
                    Capacity = int.TryParse(capacityText, NumberStyles.Integer, CultureInfo.InvariantCulture, out capacity) && capacity >= 0 ? (int?)capacity : null
                });
            }
            return resources;
        }

        internal void GetSchedulingIdentity(TalkServiceConfiguration configuration, out string originatorEmail, out string outboxUrl)
        {
            RequireAccess(configuration);
            string davRoot = configuration.GetNormalizedBaseUrl() + "/remote.php/dav/";
            XmlDocument current = SendDav(configuration, "PROPFIND", davRoot,
                "<d:propfind" + Namespaces + "><d:prop><d:current-user-principal/></d:prop></d:propfind>");
            Dictionary<string, XmlElement> currentProperties = ReadSingleResponse(current, configuration, davRoot);
            XmlElement currentPrincipal;
            if (!currentProperties.TryGetValue(PropertyKey("current-user-principal", DavNamespace), out currentPrincipal))
            {
                throw new InvalidDataException("Current DAV principal is unavailable.");
            }
            Uri principal = ValidateDavHref(configuration, davRoot, ChildText(currentPrincipal, "href", DavNamespace), "principals/users/", true);
            XmlDocument details = SendDav(configuration, "PROPFIND", principal.AbsoluteUri,
                "<d:propfind" + Namespaces + "><d:prop><c:calendar-user-address-set/><c:schedule-outbox-URL/></d:prop></d:propfind>");
            Dictionary<string, XmlElement> properties = ReadSingleResponse(details, configuration, principal.AbsoluteUri);
            originatorEmail = ReadEmail(properties);
            if (originatorEmail.Length == 0) { throw new InvalidDataException("Current DAV principal has no usable scheduling address."); }
            XmlElement outbox;
            if (!properties.TryGetValue(PropertyKey("schedule-outbox-URL", CalDavNamespace), out outbox))
            {
                throw new InvalidDataException("Current DAV principal has no scheduling outbox.");
            }
            Uri target = ValidateDavHref(configuration, principal.AbsoluteUri, ChildText(outbox, "href", DavNamespace), "calendars/", false);
            string relative = target.AbsolutePath.Substring(new Uri(davRoot).AbsolutePath.Length + "calendars/".Length).TrimEnd('/');
            string[] segments = relative.Split('/');
            if (segments.Length != 2 || !string.Equals(segments[1], "outbox", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Invalid scheduling outbox path.");
            }
            outboxUrl = target.AbsoluteUri;
        }

        private static Dictionary<string, XmlElement> ReadSingleResponse(XmlDocument document, TalkServiceConfiguration configuration, string url)
        {
            List<XmlElement> responses = GetResponses(document);
            if (responses.Count != 1) { throw new InvalidDataException("Expected one DAV principal response."); }
            Uri actual = ValidateDavHref(configuration, url, ChildText(responses[0], "href", DavNamespace), string.Empty, false);
            if (!string.Equals(actual.AbsolutePath.TrimEnd('/'), new Uri(url).AbsolutePath.TrimEnd('/'), StringComparison.Ordinal))
            {
                throw new InvalidDataException("DAV principal response did not match the request.");
            }
            return ReadSuccessfulProperties(responses[0]);
        }

        private static XmlDocument SendDav(TalkServiceConfiguration configuration, string method, string url, string payload)
        {
            NcHttpResponse response = new NcHttpClient(configuration).Send(new NcHttpRequestOptions
            {
                Method = method, Url = url, Payload = payload, ContentType = "application/xml; charset=utf-8",
                Accept = "application/xml", TimeoutMs = 60000, IncludeAuthHeader = true, IncludeOcsApiHeader = false,
                ParseJson = false, MaximumResponseBytes = 4 * 1024 * 1024,
                Headers = new Dictionary<string, string> { { "Depth", "0" } }
            });
            // The response may have rejected credentials or paused requests after the entry check.
            NextcloudConnectionState.AssertRequestAllowed(configuration);
            if (!response.HasHttpResponse || (int)response.StatusCode != 207)
            {
                throw new FreeBusyRequestException("DAV resource request failed" + (response.HasHttpResponse ? ": HTTP " + ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture) : ": no HTTP response"), response.HasHttpResponse ? (HttpStatusCode?)response.StatusCode : null, false);
            }
            return ReadXml(response.ResponseText);
        }

        internal static XmlDocument ReadXml(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) { throw new InvalidDataException("Empty DAV response."); }
            var document = new XmlDocument { XmlResolver = null };
            using (var reader = XmlReader.Create(new StringReader(text), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024 }))
            {
                document.Load(reader);
            }
            return document;
        }

        private static List<XmlElement> GetResponses(XmlDocument document)
        {
            XmlElement root = document.DocumentElement;
            if (root == null || root.LocalName != "multistatus" || root.NamespaceURI != DavNamespace) { throw new InvalidDataException("Invalid DAV multistatus response."); }
            var result = new List<XmlElement>();
            foreach (XmlNode node in root.ChildNodes)
            {
                XmlElement element = node as XmlElement;
                if (element != null && element.LocalName == "response" && element.NamespaceURI == DavNamespace) { result.Add(element); }
                else if (element != null) { throw new InvalidDataException("Unexpected DAV multistatus element."); }
            }
            return result;
        }

        private static Dictionary<string, XmlElement> ReadSuccessfulProperties(XmlElement response)
        {
            var result = new Dictionary<string, XmlElement>(StringComparer.Ordinal);
            string status = ChildText(response, "status", DavNamespace);
            if (status.Length > 0 && !IsSuccessfulDavStatus(status)) { throw new InvalidDataException("DAV principal request was rejected."); }
            bool hasProperties = false;
            foreach (XmlNode node in response.ChildNodes)
            {
                XmlElement propstat = node as XmlElement;
                if (propstat == null || propstat.LocalName != "propstat" || propstat.NamespaceURI != DavNamespace) { continue; }
                string propStatus = ChildText(propstat, "status", DavNamespace);
                if (!IsSuccessfulDavStatus(propStatus))
                {
                    if (string.Equals(propStatus, "HTTP/1.1 404 Not Found", StringComparison.Ordinal)) { continue; }
                    throw new InvalidDataException("DAV principal properties were rejected.");
                }
                XmlElement prop = Child(propstat, "prop", DavNamespace);
                if (prop == null) { throw new InvalidDataException("Missing DAV properties."); }
                hasProperties = true;
                foreach (XmlNode propertyNode in prop.ChildNodes)
                {
                    XmlElement property = propertyNode as XmlElement;
                    if (property == null) { continue; }
                    string key = PropertyKey(property.LocalName, property.NamespaceURI);
                    if (result.ContainsKey(key)) { throw new InvalidDataException("Duplicate DAV property."); }
                    result.Add(key, property);
                }
            }
            if (!hasProperties) { throw new InvalidDataException("No successful DAV properties."); }
            return result;
        }

        private static bool IsSuccessfulDavStatus(string status)
        {
            string[] parts = (status ?? string.Empty).Split(' ');
            int code;
            return parts.Length >= 2 && parts[0].StartsWith("HTTP/", StringComparison.Ordinal) && int.TryParse(parts[1], out code) && code >= 200 && code < 300;
        }

        private static string ReadEmail(Dictionary<string, XmlElement> properties)
        {
            XmlElement addresses;
            if (properties.TryGetValue(PropertyKey("calendar-user-address-set", CalDavNamespace), out addresses))
            {
                foreach (XmlNode node in addresses.ChildNodes)
                {
                    XmlElement href = node as XmlElement;
                    if (href != null && href.LocalName == "href" && href.NamespaceURI == DavNamespace && href.InnerText.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                    {
                        string address = NormalizeEmail(href.InnerText.Substring(7));
                        if (address.Length > 0) { return address; }
                    }
                }
            }
            return NormalizeEmail(PropertyText(properties, "email-address", SabreNamespace));
        }

        internal static string NormalizeEmail(string value)
        {
            string address = (value ?? string.Empty).Trim();
            if (address.Length == 0 || address.IndexOf('@') <= 0 || address.IndexOfAny(new[] { '\r', '\n', '\0', ';', ',', '<', '>', '"', ' ' }) >= 0) { return string.Empty; }
            try
            {
                var parsed = new MailAddress(address);
                return string.Equals(parsed.Address, address, StringComparison.Ordinal) ? address.ToLowerInvariant() : string.Empty;
            }
            catch (FormatException) { return string.Empty; }
        }

        internal static Uri ValidateDavHref(TalkServiceConfiguration configuration, string requestUrl, string href, string davSubpath, bool singleChild)
        {
            Uri target;
            Uri davRoot = new Uri(configuration.GetNormalizedBaseUrl() + "/remote.php/dav/");
            if (string.IsNullOrWhiteSpace(href) || href.IndexOf('\\') >= 0 || !Uri.TryCreate(new Uri(requestUrl), href, out target)
                || !string.Equals(target.Scheme, davRoot.Scheme, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(target.Host, davRoot.Host, StringComparison.OrdinalIgnoreCase) || target.Port != davRoot.Port
                || target.UserInfo.Length > 0 || target.Query.Length > 0 || target.Fragment.Length > 0)
            {
                throw new InvalidDataException("Untrusted DAV response URL.");
            }
            string prefix = davRoot.AbsolutePath + davSubpath;
            string path = target.AbsolutePath;
            string decoded = Uri.UnescapeDataString(path);
            if (!path.StartsWith(prefix, StringComparison.Ordinal) || decoded.IndexOf('\\') >= 0
                || !string.Equals(Uri.UnescapeDataString(prefix), decoded.Substring(0, Math.Min(decoded.Length, Uri.UnescapeDataString(prefix).Length)), StringComparison.Ordinal)
                || path.IndexOf("%2f", StringComparison.OrdinalIgnoreCase) >= 0 || path.IndexOf("%2e", StringComparison.OrdinalIgnoreCase) >= 0 || path.IndexOf("%25", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidDataException("DAV response URL escaped its collection.");
            }
            if (singleChild)
            {
                string child = path.Substring(prefix.Length).TrimEnd('/');
                if (child.Length == 0 || child.IndexOf('/') >= 0) { throw new InvalidDataException("Invalid DAV principal path."); }
            }
            return target;
        }

        private static string PropertyKey(string name, string ns) { return "{" + ns + "}" + name; }
        private static string PropertyText(Dictionary<string, XmlElement> properties, string name, string ns)
        {
            XmlElement element;
            return properties.TryGetValue(PropertyKey(name, ns), out element) ? element.InnerText.Trim() : string.Empty;
        }
        internal static XmlElement Child(XmlElement parent, string name, string ns)
        {
            XmlElement result = null;
            foreach (XmlNode node in parent.ChildNodes)
            {
                XmlElement element = node as XmlElement;
                if (element == null || element.LocalName != name || element.NamespaceURI != ns) { continue; }
                if (result != null) { throw new InvalidDataException("Duplicate DAV element."); }
                result = element;
            }
            return result;
        }
        internal static string ChildText(XmlElement parent, string name, string ns)
        {
            XmlElement child = Child(parent, name, ns);
            return child != null ? child.InnerText.Trim() : string.Empty;
        }
        private static string BuildScope(TalkServiceConfiguration configuration)
        {
            string value = configuration.GetNormalizedBaseUrl() + "\n" + configuration.Username.Trim() + "\n" + configuration.AppPassword;
            using (SHA256 hash = SHA256.Create()) { return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(value))); }
        }
    }
}
