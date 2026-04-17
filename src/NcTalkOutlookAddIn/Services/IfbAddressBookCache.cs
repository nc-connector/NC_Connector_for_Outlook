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
     * Loads the Nextcloud system address book (z-server-generated--system) and caches mappings
     * from email addresses to calendar UIDs (.vcf).
     */
    internal sealed class IfbAddressBookCache
    {
        private readonly object _syncRoot = new object();
        private readonly string _cacheFilePath;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        private Dictionary<string, string> _emailToUid = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _uidToEmail = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _localPartToEmail = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private DateTime _generatedUtc = DateTime.MinValue;

        internal IfbAddressBookCache(string dataDirectory)
        {
            if (string.IsNullOrEmpty(dataDirectory))
            {
                dataDirectory = Path.GetTempPath();
            }

            Directory.CreateDirectory(dataDirectory);
            _cacheFilePath = Path.Combine(dataDirectory, "ifb-addressbook-cache.json");
        }

        internal sealed class SystemAddressbookStatus
        {
            internal SystemAddressbookStatus(bool available, int count, string error)
            {
                Available = available;
                Count = count;
                Error = error ?? string.Empty;
            }

            internal bool Available { get; private set; }
            internal int Count { get; private set; }
            internal string Error { get; private set; }
        }

        internal SystemAddressbookStatus GetSystemAddressbookStatus(TalkServiceConfiguration configuration, int cacheHours, bool forceRefresh)
        {
            if (configuration == null || !configuration.IsComplete())
            {
                const string detail = "Talk credentials are incomplete.";
                DiagnosticsLogger.Log(LogCategories.Ifb, "System address book status check failed: " + detail);
                return new SystemAddressbookStatus(false, 0, detail);
            }

            lock (_syncRoot)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Ifb,
                    "System address book status check started (forceRefresh=" + forceRefresh + ").");

                try
                {
                    if (forceRefresh)
                    {
                        RefreshFromServer(configuration);
                    }
                    else
                    {
                        EnsureCache(configuration, cacheHours);
                    }

                    int count = _uidToEmail != null ? _uidToEmail.Count : 0;
                    DiagnosticsLogger.Log(
                        LogCategories.Ifb,
                        "System address book status check completed (available=True, count=" + count + ", forceRefresh=" + forceRefresh + ").");
                    return new SystemAddressbookStatus(true, count, string.Empty);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "System address book status check failed (forceRefresh=" + forceRefresh + ").",
                        ex);
                    return new SystemAddressbookStatus(false, 0, ex.Message ?? "System address book status check failed.");
                }
            }
        }

        internal bool TryGetUid(TalkServiceConfiguration configuration, int cacheHours, string email, out string uid)
        {
            uid = null;
            if (configuration == null || !configuration.IsComplete() || string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(configuration, cacheHours);
                return _emailToUid.TryGetValue(email.Trim().ToLowerInvariant(), out uid);
            }
        }

        internal bool TryGetPrimaryEmailForUid(TalkServiceConfiguration configuration, int cacheHours, string uid, out string email)
        {
            email = null;
            if (configuration == null || !configuration.IsComplete() || string.IsNullOrWhiteSpace(uid))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(configuration, cacheHours);
                return _uidToEmail.TryGetValue(uid.Trim(), out email);
            }
        }

        internal bool TryResolveEmail(TalkServiceConfiguration configuration, int cacheHours, string emailOrLocalPart, out string resolvedEmail)
        {
            resolvedEmail = null;
            if (configuration == null || !configuration.IsComplete())
            {
                return false;
            }

            string candidate = (emailOrLocalPart ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(candidate))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(configuration, cacheHours);

                string lowered = candidate.ToLowerInvariant();
                if (lowered.IndexOf('@') >= 0)
                {
                    resolvedEmail = lowered;
                    return true;
                }

                string mapped;
                if (_localPartToEmail.TryGetValue(lowered, out mapped))
                {
                    resolvedEmail = mapped;
                    return true;
                }

                foreach (var entry in _emailToUid.Keys)
                {
                    int at = entry.IndexOf('@');
                    if (at > 0)
                    {
                        string local = entry.Substring(0, at);
                        if (string.Equals(local, lowered, StringComparison.OrdinalIgnoreCase))
                        {
                            _localPartToEmail[lowered] = entry;
                            resolvedEmail = entry;
                            return true;
                        }
                    }
                }

                return false;
            }
        }

        internal List<NextcloudUser> GetUsers(TalkServiceConfiguration configuration, int cacheHours)
        {
            return GetUsers(configuration, cacheHours, false);
        }

        internal List<NextcloudUser> GetUsers(TalkServiceConfiguration configuration, int cacheHours, bool forceRefresh)
        {
            var users = new List<NextcloudUser>();
            if (configuration == null || !configuration.IsComplete())
            {
                return users;
            }

            lock (_syncRoot)
            {
                if (forceRefresh)
                {
                    RefreshFromServer(configuration);
                }
                else
                {
                    EnsureCache(configuration, cacheHours);
                }

                foreach (var pair in _uidToEmail)
                {
                    if (string.IsNullOrWhiteSpace(pair.Key))
                    {
                        continue;
                    }

                    users.Add(new NextcloudUser(pair.Key.Trim(), pair.Value ?? string.Empty));
                }
            }

            users.Sort((a, b) => string.Compare(a.UserId, b.UserId, StringComparison.OrdinalIgnoreCase));
            return users;
        }

        private void EnsureCache(TalkServiceConfiguration configuration, int cacheHours)
        {
            if (cacheHours < 1)
            {
                cacheHours = 1;
            }

            if (!LoadFromDisk(cacheHours))
            {
                RefreshFromServer(configuration);
            }
        }

        private bool LoadFromDisk(int cacheHours)
        {
            if (!File.Exists(_cacheFilePath))
            {
                return false;
            }

            try
            {
                string json = File.ReadAllText(_cacheFilePath);
                var data = _serializer.Deserialize<CacheContainer>(json);
                if (data == null || data.GeneratedUtc <= DateTime.MinValue || data.Entries == null)
                {
                    return false;
                }

                if (data.GeneratedUtc.AddHours(cacheHours) <= DateTime.UtcNow)
                {
                    return false;
                }

                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var uidMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var localMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in data.Entries)
                {
                    if (entry == null || string.IsNullOrWhiteSpace(entry.Email) || string.IsNullOrWhiteSpace(entry.Uid))
                    {
                        continue;
                    }

                    var key = entry.Email.Trim().ToLowerInvariant();
                    if (!map.ContainsKey(key))
                    {
                        map[key] = entry.Uid.Trim();
                    }

                    var uidKey = entry.Uid.Trim();
                    if (!string.IsNullOrEmpty(uidKey) && !uidMap.ContainsKey(uidKey))
                    {
                        uidMap[uidKey] = key;
                    }

                    int at = key.IndexOf('@');
                    if (at > 0)
                    {
                        string local = key.Substring(0, at);
                        if (!localMap.ContainsKey(local))
                        {
                            localMap[local] = key;
                        }
                    }
                }

                _emailToUid = map;
                _uidToEmail = uidMap;
                _localPartToEmail = localMap;
                _generatedUtc = data.GeneratedUtc;
                return map.Count > 0;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to load IFB address book cache from disk.", ex);
                return false;
            }
        }

        private void RefreshFromServer(TalkServiceConfiguration configuration)
        {
            string baseUrl = configuration.GetNormalizedBaseUrl();
            if (string.IsNullOrEmpty(baseUrl))
            {
                throw new InvalidOperationException("Server URL is not configured.");
            }

            string addressBookUrl = string.Format(CultureInfo.InvariantCulture,
                "{0}/remote.php/dav/addressbooks/users/{1}/z-server-generated--system?export",
                baseUrl,
                Uri.EscapeDataString(configuration.Username ?? string.Empty));

            string responseText = null;
            var httpClient = new NcHttpClient(configuration);
            NcHttpResponse response = httpClient.Send(new NcHttpRequestOptions
            {
                Method = "GET",
                Url = addressBookUrl,
                Accept = "text/vcard,text/x-vcard,text/plain,*/*",
                TimeoutMs = 60000,
                IncludeAuthHeader = true,
                IncludeOcsApiHeader = false,
                ParseJson = false
            });

            if (!response.HasHttpResponse)
            {
                if (response.TransportException != null)
                {
                    DiagnosticsLogger.LogException(LogCategories.Ifb, "Address book could not be loaded from server.", response.TransportException);
                    throw new InvalidOperationException("Address book could not be loaded: " + response.TransportException.Message, response.TransportException);
                }

                throw new InvalidOperationException("Address book could not be loaded: no HTTP response.");
            }

            responseText = response.ResponseText ?? string.Empty;
            if ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                string status = ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture);
                throw new InvalidOperationException("Address book could not be loaded: HTTP " + status + ".");
            }

            if (string.IsNullOrEmpty(responseText))
            {
                throw new InvalidOperationException("Address book response was empty.");
            }

            var entries = ParseAddressBook(responseText);
            var emailMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var uidMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var localMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Uid))
                {
                    continue;
                }

                string uidKey = entry.Uid.Trim();
                if (entry.Emails != null)
                {
                    foreach (var email in entry.Emails)
                    {
                        if (string.IsNullOrWhiteSpace(email))
                        {
                            continue;
                        }

                        string emailKey = email.Trim().ToLowerInvariant();
                        if (!emailMap.ContainsKey(emailKey))
                        {
                            emailMap[emailKey] = uidKey;
                        }

                        if (!uidMap.ContainsKey(uidKey))
                        {
                            uidMap[uidKey] = emailKey;
                        }

                        int at = emailKey.IndexOf('@');
                        if (at > 0)
                        {
                            string local = emailKey.Substring(0, at);
                            if (!localMap.ContainsKey(local))
                            {
                                localMap[local] = emailKey;
                            }
                        }
                    }
                }
            }

            _emailToUid = emailMap;
            _uidToEmail = uidMap;
            _localPartToEmail = localMap;
            _generatedUtc = DateTime.UtcNow;

            try
            {
                var data = new CacheContainer
                {
                    GeneratedUtc = _generatedUtc,
                    Entries = new List<CacheEntry>()
                };

                foreach (var kvp in emailMap)
                {
                    data.Entries.Add(new CacheEntry { Email = kvp.Key, Uid = kvp.Value });
                }

                string json = _serializer.Serialize(data);
                File.WriteAllText(_cacheFilePath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // Cache writes must never fail the request.
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to write IFB address book cache to disk.", ex);
            }
        }

        private static List<AddressBookEntry> ParseAddressBook(string data)
        {
            var result = new List<AddressBookEntry>();

            if (string.IsNullOrEmpty(data))
            {
                return result;
            }

            string normalized = data
                .Replace("\r\n ", string.Empty)
                .Replace("\n ", string.Empty)
                .Replace("\r\n\t", string.Empty)
                .Replace("\n\t", string.Empty);

            using (var reader = new StringReader(normalized))
            {
                string line;
                bool inside = false;
                string uid = null;
                List<string> emails = new List<string>();

                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    if (line.StartsWith("BEGIN:VCARD", StringComparison.OrdinalIgnoreCase))
                    {
                        inside = true;
                        uid = null;
                        emails.Clear();
                        continue;
                    }

                    if (line.StartsWith("END:VCARD", StringComparison.OrdinalIgnoreCase))
                    {
                        if (inside && !string.IsNullOrEmpty(uid))
                        {
                            result.Add(new AddressBookEntry(uid.Trim(), new List<string>(emails)));
                        }

                        inside = false;
                        uid = null;
                        emails.Clear();
                        continue;
                    }

                    if (!inside)
                    {
                        continue;
                    }

                    if (line.StartsWith("UID", StringComparison.OrdinalIgnoreCase))
                    {
                        int colon = line.IndexOf(':', 3);
                        if (colon >= 0 && colon + 1 < line.Length)
                        {
                            uid = line.Substring(colon + 1).Trim();
                        }
                        continue;
                    }

                    if (line.StartsWith("EMAIL", StringComparison.OrdinalIgnoreCase))
                    {
                        int colon = line.IndexOf(':', 5);
                        if (colon >= 0 && colon + 1 < line.Length)
                        {
                            string mail = line.Substring(colon + 1).Trim().ToLowerInvariant();
                            if (!string.IsNullOrEmpty(mail))
                            {
                                emails.Add(mail);
                            }
                        }
                    }
                }
            }

            return result;
        }

        private sealed class CacheContainer
        {
            public DateTime GeneratedUtc { get; set; }
            public List<CacheEntry> Entries { get; set; }
        }

        private sealed class CacheEntry
        {
            public string Email { get; set; }
            public string Uid { get; set; }
        }

        private sealed class AddressBookEntry
        {
            internal AddressBookEntry(string uid, List<string> emails)
            {
                Uid = uid;
                Emails = emails ?? new List<string>();
            }

            internal string Uid { get; private set; }
            internal List<string> Emails { get; private set; }
        }
    }
}
