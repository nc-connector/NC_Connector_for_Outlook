// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Caches system-address-book email and UID mappings for one Outlook profile.
    internal sealed class IfbAddressBookCache
    {
        private readonly object _syncRoot = new object();
        private readonly string _dataDirectory;
        private readonly string _profileScope;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        private Dictionary<string, string> _emailToUid =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _uidToEmail =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
        private DateTime _generatedUtc = DateTime.MinValue;
        private string _activeScopeFingerprint = string.Empty;
        private static readonly ConcurrentDictionary<string, byte> FailedRefreshScopes =
            new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        internal IfbAddressBookCache(string dataDirectory)
            : this(dataDirectory, "default")
        {
        }

        internal IfbAddressBookCache(
            string dataDirectory,
            string profileScope)
        {
            _dataDirectory = string.IsNullOrWhiteSpace(dataDirectory)
                ? AppDataPaths.EnsureLocalRootDirectory()
                : dataDirectory;
            _profileScope = string.IsNullOrWhiteSpace(profileScope)
                ? "default"
                : profileScope.Trim();
            Directory.CreateDirectory(_dataDirectory);
        }

        internal sealed class SystemAddressbookStatus
        {
            internal SystemAddressbookStatus(
                bool available,
                int count,
                string error)
            {
                Available = available;
                Count = count;
                Error = error ?? string.Empty;
            }

            internal bool Available { get; private set; }

            internal int Count { get; private set; }

            internal string Error { get; private set; }
        }

        internal SystemAddressbookStatus GetSystemAddressbookStatus(
            TalkServiceConfiguration configuration,
            int cacheHours,
            bool forceRefresh)
        {
            if (configuration == null
                || !configuration.IsComplete())
            {
                string detail = Strings.ErrorMissingCredentials;
                DiagnosticsLogger.Log(
                    LogCategories.Ifb,
                    "System address book status check failed: "
                    + detail);
                return new SystemAddressbookStatus(
                    false,
                    0,
                    detail);
            }

            lock (_syncRoot)
            {
                try
                {
                    CacheScope scope = CreateScope(configuration);
                    if (forceRefresh)
                    {
                        RefreshFromServer(
                            configuration,
                            scope);
                    }
                    else
                    {
                        EnsureCache(
                            configuration,
                            cacheHours,
                            scope);
                    }
                    return new SystemAddressbookStatus(
                        true,
                        _uidToEmail.Count,
                        string.Empty);
                }
                catch (Exception ex)
                {
                    string error = ex is InvalidDataException
                        ? Strings.TalkSystemAddressbookInvalidResponse
                        : Strings.TalkSystemAddressbookFetchFailed;
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "System address book status check failed.",
                        new InvalidOperationException(error));
                    return new SystemAddressbookStatus(
                        false,
                        0,
                        error);
                }
            }
        }

        internal bool TryGetUid(
            TalkServiceConfiguration configuration,
            int cacheHours,
            string email,
            out string uid)
        {
            uid = null;
            if (configuration == null
                || !configuration.IsComplete()
                || string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(
                    configuration,
                    cacheHours,
                    CreateScope(configuration));
                return _emailToUid.TryGetValue(
                    email.Trim().ToLowerInvariant(),
                    out uid);
            }
        }

        internal bool TryGetPrimaryEmailForUid(
            TalkServiceConfiguration configuration,
            int cacheHours,
            string uid,
            out string email)
        {
            email = null;
            if (configuration == null
                || !configuration.IsComplete()
                || string.IsNullOrWhiteSpace(uid))
            {
                return false;
            }

            lock (_syncRoot)
            {
                EnsureCache(
                    configuration,
                    cacheHours,
                    CreateScope(configuration));
                return _uidToEmail.TryGetValue(
                    uid.Trim(),
                    out email);
            }
        }

        internal List<NextcloudUser> GetUsers(
            TalkServiceConfiguration configuration,
            int cacheHours,
            bool forceRefresh)
        {
            var users = new List<NextcloudUser>();
            if (configuration == null
                || !configuration.IsComplete())
            {
                return users;
            }

            lock (_syncRoot)
            {
                CacheScope scope = CreateScope(configuration);
                if (forceRefresh)
                {
                    RefreshFromServer(
                        configuration,
                        scope);
                }
                else
                {
                    EnsureCache(
                        configuration,
                        cacheHours,
                        scope);
                }
                foreach (KeyValuePair<string, string> pair
                    in _uidToEmail)
                {
                    if (!string.IsNullOrWhiteSpace(pair.Key))
                    {
                        users.Add(
                            new NextcloudUser(
                                pair.Key.Trim(),
                                pair.Value
                                ?? string.Empty));
                    }
                }
            }

            users.Sort(
                (left, right) => string.Compare(
                    left.UserId,
                    right.UserId,
                    StringComparison.OrdinalIgnoreCase));
            return users;
        }

        private void EnsureCache(
            TalkServiceConfiguration configuration,
            int cacheHours,
            CacheScope scope)
        {
            int validHours = Math.Max(1, cacheHours);
            if (FailedRefreshScopes.ContainsKey(BuildCacheFilePath(scope)))
            {
                RefreshFromServer(configuration, scope);
                return;
            }
            if (string.Equals(
                    _activeScopeFingerprint,
                    scope.Fingerprint,
                    StringComparison.Ordinal)
                && _generatedUtc > DateTime.MinValue
                && _generatedUtc.AddHours(validHours)
                   > DateTime.UtcNow)
            {
                return;
            }

            // Scope construction is local, so disk cache is tried before UID resolution.
            if (!LoadFromDisk(scope, validHours))
            {
                RefreshFromServer(configuration, scope);
            }
        }

        private bool LoadFromDisk(
            CacheScope scope,
            int cacheHours)
        {
            string path = BuildCacheFilePath(scope);
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                CacheContainer data =
                    _serializer.Deserialize<CacheContainer>(
                        File.ReadAllText(path, Encoding.UTF8));
                if (data == null
                    || data.GeneratedUtc <= DateTime.MinValue
                    || data.Entries == null
                    || data.GeneratedUtc.AddHours(cacheHours)
                       <= DateTime.UtcNow
                    || !string.Equals(
                        data.ScopeFingerprint,
                        scope.Fingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                ApplyEntries(
                    data.Entries,
                    data.GeneratedUtc,
                    scope.Fingerprint);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to load IFB address book cache from disk ("
                    + ex.GetType().Name + ").",
                    new InvalidDataException("The cached system address book could not be read."));
                return false;
            }
        }

        private void RefreshFromServer(
            TalkServiceConfiguration configuration,
            CacheScope scope)
        {
            // A failed refresh must not become a fresh cache hit on the next lookup.
            string cachePath = BuildCacheFilePath(scope);
            FailedRefreshScopes[cachePath] = 0;
            using (DiagnosticsLogger.BeginOperation(
                LogCategories.Ifb,
                "Refresh system address book"))
            {
                try
                {
                    string currentUserId =
                        NextcloudUserIdentityService.ResolveCurrentUserId(
                            configuration);
                    string addressBookUrl = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}/remote.php/dav/addressbooks/users/{1}/z-server-generated--system?export",
                        scope.ServerBaseUrl,
                        Uri.EscapeDataString(currentUserId));

                    var httpClient = new NcHttpClient(configuration);
                    NcHttpResponse response = httpClient.Send(
                        new NcHttpRequestOptions
                        {
                            Method = "GET",
                            Url = addressBookUrl,
                            Accept =
                                "text/vcard,text/x-vcard,text/plain,*/*",
                            TimeoutMs = 60000,
                            IncludeAuthHeader = true,
                            IncludeOcsApiHeader = false,
                            ParseJson = false
                        });
                    if (response == null || !response.HasHttpResponse)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Ifb,
                            "System address book fetch failed without an HTTP response.");
                        throw new InvalidOperationException(
                            Strings.TalkSystemAddressbookFetchFailed);
                    }
                    int statusCode = (int)response.StatusCode;
                    if ((statusCode < 200 || statusCode >= 300)
                        && statusCode != 404)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Ifb,
                            "System address book fetch failed: HTTP "
                            + statusCode.ToString(
                                CultureInfo.InvariantCulture)
                            + ".");
                        throw new InvalidOperationException(
                            Strings.TalkSystemAddressbookFetchFailed);
                    }
                    string responseText =
                        response.ResponseText ?? string.Empty;
                    string contentType = (response.ContentType ?? string.Empty)
                        .Split(';')[0].Trim().ToLowerInvariant();
                    bool isVcardContentType = contentType == "text/vcard"
                        || contentType == "text/x-vcard"
                        || contentType == "text/directory";
                    if (string.IsNullOrWhiteSpace(responseText)
                        && (statusCode == 404 || !isVcardContentType))
                    {
                        throw new InvalidDataException(
                            Strings.TalkSystemAddressbookInvalidResponse);
                    }

                    List<CacheEntry> entries =
                        ParseAddressBook(responseText);
                    if (statusCode == 404 && entries.Count == 0)
                    {
                        throw new InvalidDataException(
                            Strings.TalkSystemAddressbookInvalidResponse);
                    }
                    if (statusCode == 404)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Ifb,
                            "System address book accepted a valid non-empty HTTP 404 export.");
                    }
                    DateTime generatedUtc = DateTime.UtcNow;
                    ApplyEntries(
                        entries,
                        generatedUtc,
                        scope.Fingerprint);
                    SaveToDisk(scope, generatedUtc);
                    byte ignored;
                    FailedRefreshScopes.TryRemove(cachePath, out ignored);
                    DiagnosticsLogger.Log(
                        LogCategories.Ifb,
                        "System address book refreshed (users="
                        + _uidToEmail.Count.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                catch (Exception ex)
                {
                    // Upstream exceptions may include response bodies; expose only the failure category.
                    Exception failure = ex is InvalidDataException
                        ? (Exception)new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse)
                        : new InvalidOperationException(Strings.TalkSystemAddressbookFetchFailed);
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "System address book refresh failed (" + ex.GetType().Name + ").",
                        failure);
                    throw failure;
                }
            }
        }

        private void ApplyEntries(
            IEnumerable<CacheEntry> entries,
            DateTime generatedUtc,
            string scopeFingerprint)
        {
            var emailMap = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var uidMap = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (CacheEntry entry in entries)
            {
                if (entry == null
                    || string.IsNullOrWhiteSpace(entry.Uid))
                {
                    continue;
                }
                string email =
                    (entry.Email ?? string.Empty).Trim().ToLowerInvariant();
                string uid = entry.Uid.Trim();
                if (email.Length > 0 && !emailMap.ContainsKey(email))
                {
                    emailMap[email] = uid;
                }
                if (!uidMap.ContainsKey(uid)
                    || string.IsNullOrEmpty(uidMap[uid]))
                {
                    uidMap[uid] = email;
                }
            }

            _emailToUid = emailMap;
            _uidToEmail = uidMap;
            _generatedUtc = generatedUtc;
            _activeScopeFingerprint = scopeFingerprint;
        }

        private void SaveToDisk(
            CacheScope scope,
            DateTime generatedUtc)
        {
            var data = new CacheContainer
            {
                GeneratedUtc = generatedUtc,
                ScopeFingerprint = scope.Fingerprint,
                Entries = new List<CacheEntry>()
            };
            foreach (KeyValuePair<string, string> pair
                in _emailToUid)
            {
                data.Entries.Add(
                    new CacheEntry
                    {
                        Email = pair.Key,
                        Uid = pair.Value
                    });
            }
            foreach (KeyValuePair<string, string> pair in _uidToEmail)
            {
                if (string.IsNullOrEmpty(pair.Value))
                {
                    data.Entries.Add(new CacheEntry
                    {
                        Email = string.Empty,
                        Uid = pair.Key
                    });
                }
            }

            try
            {
                File.WriteAllText(
                    BuildCacheFilePath(scope),
                    _serializer.Serialize(data),
                    new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Ifb,
                    "Failed to write IFB address book cache to disk.",
                    ex);
            }
        }

        private CacheScope CreateScope(
            TalkServiceConfiguration configuration)
        {
            string serverBaseUrl =
                configuration.GetNormalizedBaseUrl();
            if (string.IsNullOrWhiteSpace(serverBaseUrl))
            {
                throw new InvalidOperationException(
                    "Server URL is invalid.");
            }
            string username =
                (configuration.Username
                 ?? string.Empty).Trim();
            return new CacheScope(
                serverBaseUrl,
                BuildScopeFingerprint(
                    _profileScope,
                    serverBaseUrl,
                    username));
        }

        private string BuildCacheFilePath(CacheScope scope)
        {
            return Path.Combine(
                _dataDirectory,
                "ifb-addressbook-cache-"
                + scope.Fingerprint
                + ".json");
        }

        internal static string BuildScopeFingerprint(
            string profileScope,
            string serverBaseUrl,
            string username)
        {
            string input =
                (profileScope ?? string.Empty).Trim()
                + "\n"
                + (serverBaseUrl ?? string.Empty).TrimEnd('/')
                + "\n"
                + (username ?? string.Empty).Trim();
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create())
            {
                hash = sha256.ComputeHash(
                    Encoding.UTF8.GetBytes(input));
            }

            var builder = new StringBuilder(32);
            for (int i = 0; i < 16; i++)
            {
                builder.Append(
                    hash[i].ToString(
                        "x2",
                        CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }

        private static List<CacheEntry> ParseAddressBook(
            string data)
        {
            var result = new List<CacheEntry>();
            string normalized = Regex.Replace(
                (data ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n"),
                "\n[ \t]",
                string.Empty);
            using (var reader = new StringReader(normalized))
            {
                string line;
                string uid = null;
                var emails = new List<string>();
                bool inside = false;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    line = line.TrimEnd();
                    if (line.Equals(
                            "BEGIN:VCARD",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        if (inside)
                        {
                            throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                        }
                        inside = true;
                        uid = null;
                        emails.Clear();
                    }
                    else if (line.Equals(
                                 "END:VCARD",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (!inside || string.IsNullOrWhiteSpace(uid))
                        {
                            throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                        }
                        if (emails.Count == 0)
                        {
                            result.Add(new CacheEntry { Email = string.Empty, Uid = uid });
                        }
                        else
                        {
                            foreach (string email in emails)
                            {
                                result.Add(
                                    new CacheEntry
                                    {
                                        Email = email,
                                        Uid = uid.Trim()
                                    });
                            }
                        }
                        inside = false;
                    }
                    else
                    {
                        if (!inside)
                        {
                            throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                        }
                        string propertyName;
                        string value = ReadVCardValue(line, out propertyName);
                        if (propertyName == "BEGIN" || propertyName == "END")
                        {
                            throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                        }
                        if (propertyName == "UID")
                        {
                            if (uid != null || string.IsNullOrWhiteSpace(value)
                                || Regex.IsMatch(value, @"[\x00-\x1f\x7f]"))
                            {
                                throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                            }
                            uid = value;
                        }
                        else if (propertyName == "EMAIL" && value.Length > 0)
                        {
                            emails.Add(value.ToLowerInvariant());
                        }
                    }
                }
                if (inside)
                {
                    throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
                }
            }
            return result;
        }

        private static string ReadVCardValue(
            string line,
            out string propertyName)
        {
            Match header = Regex.Match(line,
                @"^(?:[A-Za-z0-9-]+\.)?([A-Za-z0-9-]+)(?:;[A-Za-z0-9-]+=(?:""[^""\r\n]*""|[^;:""\r\n]+)(?:,(?:""[^""\r\n]*""|[^;:""\r\n]+))*)*:");
            if (!header.Success)
            {
                throw new InvalidDataException(Strings.TalkSystemAddressbookInvalidResponse);
            }
            propertyName = header.Groups[1].Value.ToUpperInvariant();
            string value = line.Substring(header.Length).Trim();
            return Regex.Replace(value, @"\\([\\,;nN])", match =>
                match.Groups[1].Value.Equals("n", StringComparison.OrdinalIgnoreCase)
                    ? "\n"
                    : match.Groups[1].Value);
        }

        private sealed class CacheScope
        {
            internal CacheScope(
                string serverBaseUrl,
                string fingerprint)
            {
                ServerBaseUrl = serverBaseUrl;
                Fingerprint = fingerprint;
            }

            internal string ServerBaseUrl { get; private set; }

            internal string Fingerprint { get; private set; }
        }

        private sealed class CacheContainer
        {
            public DateTime GeneratedUtc { get; set; }

            public string ScopeFingerprint { get; set; }

            public List<CacheEntry> Entries { get; set; }
        }

        private sealed class CacheEntry
        {
            public string Email { get; set; }

            public string Uid { get; set; }
        }
    }
}
