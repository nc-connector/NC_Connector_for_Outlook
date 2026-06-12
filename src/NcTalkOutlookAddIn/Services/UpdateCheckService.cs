// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class UpdateCheckService
    {
        private const string EndpointUrl = "https://nc-connector.de/wp-json/ncc/v1/update-check";
        private const string ProductKey = "outlook";
        private const string Channel = "stable";
        private const int TimeoutMs = 5000;

        internal async Task<UpdateCheckResult> CheckAsync(AddinSettings settings, bool force)
        {
            if (settings == null)
            {
                settings = new AddinSettings();
            }
            EnsureInstallId(settings);

            if (!force && WasCheckedToday(settings.UpdateLastCheckedAtUtc))
            {
                return BuildCachedResult(settings);
            }

            string currentVersion = AddinVersionInfo.GetVersion();
            string clientDayId = BuildClientDayId(settings.UpdateInstallId, DateTime.UtcNow.Date);
            string requestUrl = BuildRequestUrl(currentVersion, clientDayId);

            DiagnosticsLogger.Log(LogCategories.Core, "Update check started (product=outlook, version=" + currentVersion + ").");
            using (DiagnosticsLogger.BeginOperation(LogCategories.Core, "UpdateCheck.Fetch"))
            {
                string responseText = await FetchStringAsync(requestUrl).ConfigureAwait(false);
                UpdateCheckResult result = ParseResponse(responseText);
                result.CurrentVersion = currentVersion;
                result.CheckedAtUtc = DateTime.UtcNow;
                ApplyResult(settings, result);
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Update check completed (latest=" + result.LatestVersion
                    + ", updateAvailable=" + result.UpdateAvailable
                    + ", counted=" + result.Counted
                    + ").");
                return result;
            }
        }

        internal static bool ShouldNotify(AddinSettings settings, UpdateCheckResult result)
        {
            if (settings == null || result == null || !settings.UpdateNotifyEnabled || !result.UpdateAvailable)
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(result.LatestVersion))
            {
                return false;
            }

            string today = DateTime.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return !string.Equals(settings.UpdateLastNotifiedVersion, result.LatestVersion, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(settings.UpdateLastNotifiedDateUtc, today, StringComparison.OrdinalIgnoreCase);
        }

        internal static void MarkNotified(AddinSettings settings, UpdateCheckResult result)
        {
            if (settings == null || result == null)
            {
                return;
            }

            settings.UpdateLastNotifiedVersion = result.LatestVersion ?? string.Empty;
            settings.UpdateLastNotifiedDateUtc = DateTime.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        internal static string GetPreferredOpenUrl(UpdateCheckResult result)
        {
            if (result == null)
            {
                return string.Empty;
            }
            if (!string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                return result.DownloadUrl.Trim();
            }
            return result.ReleaseUrl == null ? string.Empty : result.ReleaseUrl.Trim();
        }

        internal static string BuildNotificationMessage(UpdateCheckResult result)
        {
            if (result == null)
            {
                return string.Empty;
            }

            string version = string.IsNullOrWhiteSpace(result.LatestVersion) ? Strings.TalkVersionUnknown : result.LatestVersion.Trim();
            string details = string.IsNullOrWhiteSpace(result.ChangelogText)
                ? Strings.UpdateChangelogEmpty
                : result.ChangelogText.Trim();
            return string.Format(CultureInfo.CurrentCulture, Strings.UpdateAvailableMessageFormat, version, details);
        }

        internal static UpdateCheckResult BuildCachedResult(AddinSettings settings)
        {
            var result = new UpdateCheckResult();
            if (settings == null)
            {
                result.FromCache = true;
                return result;
            }

            result.CurrentVersion = AddinVersionInfo.GetVersion();
            result.LatestVersion = settings.UpdateLatestVersion ?? string.Empty;
            result.ReleaseUrl = settings.UpdateReleaseUrl ?? string.Empty;
            result.DownloadUrl = settings.UpdateDownloadUrl ?? string.Empty;
            result.PublishedAt = settings.UpdatePublishedAt ?? string.Empty;
            result.ChangelogTitle = settings.UpdateChangelogTitle ?? string.Empty;
            result.ChangelogText = settings.UpdateChangelogText ?? string.Empty;
            result.CheckedAtUtc = ParseUtc(settings.UpdateLastCheckedAtUtc) ?? DateTime.MinValue;
            result.UpdateAvailable = IsNewerVersion(result.LatestVersion, result.CurrentVersion);
            result.FromCache = true;
            return result;
        }

        private static string BuildRequestUrl(string currentVersion, string clientDayId)
        {
            var query = new StringBuilder();
            query.Append("?product=").Append(Uri.EscapeDataString(ProductKey));
            query.Append("&version=").Append(Uri.EscapeDataString(currentVersion ?? string.Empty));
            query.Append("&channel=").Append(Uri.EscapeDataString(Channel));
            query.Append("&client_day_id=").Append(Uri.EscapeDataString(clientDayId ?? string.Empty));
            return EndpointUrl + query;
        }

        private static async Task<string> FetchStringAsync(string requestUrl)
        {
            // Public homepage check: keep this async and separate from the authenticated Nextcloud OCS client.
            var request = (HttpWebRequest)WebRequest.Create(requestUrl);
            request.Method = "GET";
            request.Accept = "application/json";
            request.UserAgent = "NC-Connector-Outlook/" + (AddinVersionInfo.GetVersion() ?? string.Empty);
            request.Timeout = TimeoutMs;
            request.ReadWriteTimeout = TimeoutMs;
            request.KeepAlive = false;

            Task<WebResponse> responseTask = request.GetResponseAsync();
            Task completed = await Task.WhenAny(responseTask, Task.Delay(TimeoutMs)).ConfigureAwait(false);
            if (!ReferenceEquals(completed, responseTask))
            {
                request.Abort();
                throw new TimeoutException("Update check timed out.");
            }

            using (var response = (HttpWebResponse)await responseTask.ConfigureAwait(false))
            using (Stream stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream ?? Stream.Null, Encoding.UTF8))
            {
                return await reader.ReadToEndAsync().ConfigureAwait(false);
            }
        }

        private static UpdateCheckResult ParseResponse(string responseText)
        {
            IDictionary<string, object> payload = NcJson.DeserializeObject(responseText);
            if (payload == null)
            {
                throw new InvalidDataException("Update check returned invalid JSON.");
            }

            var result = new UpdateCheckResult
            {
                LatestVersion = NcJson.GetStringOrEmpty(payload, "latest_version"),
                ReleaseUrl = NcJson.GetStringOrEmpty(payload, "release_url"),
                DownloadUrl = NcJson.GetStringOrEmpty(payload, "download_url"),
                PublishedAt = NcJson.GetStringOrEmpty(payload, "published_at"),
                Message = NcJson.GetStringOrEmpty(payload, "message"),
                UpdateAvailable = GetBool(payload, "update_available"),
                Counted = GetBool(payload, "counted")
            };

            IDictionary<string, object> changelog = NcJson.GetDictionary(payload, "changelog");
            result.ChangelogTitle = NcJson.GetStringOrEmpty(changelog, "title");
            result.ChangelogText = BuildChangelogText(changelog);
            return result;
        }

        private static string BuildChangelogText(IDictionary<string, object> changelog)
        {
            IDictionary<string, object> sections = NcJson.GetDictionary(changelog, "sections");
            if (sections == null)
            {
                return string.Empty;
            }

            var parts = new List<string>();
            AppendChangelogSection(parts, sections, "added", Strings.UpdateChangelogAdded);
            AppendChangelogSection(parts, sections, "changed", Strings.UpdateChangelogChanged);
            AppendChangelogSection(parts, sections, "fixed", Strings.UpdateChangelogFixed);
            return string.Join(Environment.NewLine + Environment.NewLine, parts.ToArray());
        }

        private static void AppendChangelogSection(List<string> parts, IDictionary<string, object> sections, string key, string label)
        {
            object raw;
            if (parts == null || sections == null || !sections.TryGetValue(key, out raw))
            {
                return;
            }

            IEnumerable items = raw as IEnumerable;
            if (items == null || raw is string)
            {
                return;
            }

            var lines = new List<string>();
            foreach (object item in items)
            {
                string text = Convert.ToString(item, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    lines.Add("- " + text.Trim());
                }
            }
            if (lines.Count == 0)
            {
                return;
            }

            parts.Add((label ?? key) + Environment.NewLine + string.Join(Environment.NewLine, lines.ToArray()));
        }

        private static bool GetBool(IDictionary<string, object> payload, string key)
        {
            if (payload == null || string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            object raw;
            if (!payload.TryGetValue(key, out raw) || raw == null)
            {
                return false;
            }
            if (raw is bool)
            {
                return (bool)raw;
            }
            string value = Convert.ToString(raw, CultureInfo.InvariantCulture);
            bool parsed;
            return bool.TryParse(value, out parsed) && parsed;
        }

        private static void ApplyResult(AddinSettings settings, UpdateCheckResult result)
        {
            if (settings == null || result == null)
            {
                return;
            }

            settings.UpdateLastCheckedAtUtc = result.CheckedAtUtc.ToString("o", CultureInfo.InvariantCulture);
            settings.UpdateLatestVersion = result.LatestVersion ?? string.Empty;
            settings.UpdateReleaseUrl = result.ReleaseUrl ?? string.Empty;
            settings.UpdateDownloadUrl = result.DownloadUrl ?? string.Empty;
            settings.UpdatePublishedAt = result.PublishedAt ?? string.Empty;
            settings.UpdateChangelogTitle = result.ChangelogTitle ?? string.Empty;
            settings.UpdateChangelogText = result.ChangelogText ?? string.Empty;
        }

        internal static void EnsureInstallId(AddinSettings settings)
        {
            if (settings == null || !string.IsNullOrWhiteSpace(settings.UpdateInstallId))
            {
                return;
            }

            settings.UpdateInstallId = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
        }

        private static string BuildClientDayId(string installId, DateTime dayUtc)
        {
            string raw = (installId ?? string.Empty) + "|" + dayUtc.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "|" + ProductKey;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                var builder = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }
                return builder.ToString();
            }
        }

        private static bool WasCheckedToday(string lastCheckedAtUtc)
        {
            DateTime? checkedAt = ParseUtc(lastCheckedAtUtc);
            return checkedAt.HasValue && checkedAt.Value.Date == DateTime.UtcNow.Date;
        }

        private static DateTime? ParseUtc(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            DateTime parsed;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out parsed))
            {
                return parsed.ToUniversalTime();
            }
            return null;
        }

        private static bool IsNewerVersion(string latestVersion, string currentVersion)
        {
            Version latest;
            Version current;
            if (!Version.TryParse(NormalizeVersion(latestVersion), out latest) || !Version.TryParse(NormalizeVersion(currentVersion), out current))
            {
                return false;
            }
            return latest > current;
        }

        private static string NormalizeVersion(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
            {
                return string.Empty;
            }
            string value = version.Trim().TrimStart('v', 'V');
            int dashIndex = value.IndexOf('-');
            if (dashIndex >= 0)
            {
                value = value.Substring(0, dashIndex);
            }
            return value;
        }
    }
}
