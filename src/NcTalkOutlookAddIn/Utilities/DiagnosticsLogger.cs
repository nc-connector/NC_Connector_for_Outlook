/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Writes internal diagnostics to a log file under %LOCALAPPDATA%.
     */
    internal static class DiagnosticsLogger
    {
        private static readonly object SyncRoot = new object();
        private static readonly string LogDirectory = AppDataPaths.GetLocalRootDirectory();
        private const string LegacyLogFileName = "addin-runtime.log";
        private const string DailyLogFilePrefix = "addin-runtime.log_";
        private const string DailyLogDateFormat = "yyyyMMdd";
        private const int KeepDailyLogCount = 7;
        private const int CleanupOlderThanDays = 30;
        private static readonly string LegacyLogFilePath = Path.Combine(LogDirectory, LegacyLogFileName);
        private static readonly Regex AuthorizationHeaderRegex = new Regex(
            "(Authorization\\s*:\\s*(?:Bearer|Basic)\\s+)[^\\s,;]+",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex UrlCredentialRegex = new Regex(
            "(https?://)[^\\s/@]+:[^\\s/@]+@",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SecretQueryRegex = new Regex(
            "([?&](?:access_token|refresh_token|token|password|pass|secret|code|auth|apikey|app_password)=)[^&\\s]+",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SecretJsonRegex = new Regex(
            "(\"(?:access_token|refresh_token|token|password|pass|secret|code|auth|apikey|app_password)\"\\s*:\\s*)\"(?:[^\"\\\\]|\\\\.)*\"",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ShareCallTokenRegex = new Regex(
            "(\\/(?:s|call)\\/)([A-Za-z0-9_-]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex DavUserPathRegex = new Regex(
            "(\\/dav\\/files\\/)([^\\/\\s]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex UserFieldRegex = new Regex(
            "(\\bUser(?:name)?\\s*=\\s*)([^,\\)\\s]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex JsonUserFieldRegex = new Regex(
            "(\"(?:user|username)\"\\s*:\\s*\")([^\"]*)(\")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex EmailRegex = new Regex(
            "\\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}\\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex WindowsUserPathRegex = new Regex(
            "([A-Za-z]:\\\\Users\\\\)([^\\\\]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ProviderUserPathRegex = new Regex(
            "(\\\\\\\\[^\\\\\\s]+\\\\[^\\\\\\s]+\\\\Users\\\\)([^\\\\]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static DateTime _lastCleanupDateLocal = DateTime.MinValue.Date;
        private static bool _enabled;
        private static bool _anonymizationEnabled = true;
        private static string[] _serverUrlTokens = new string[0];

        private struct DatedLogFile
        {
            internal DateTime DateLocal;
            internal string Path;
        }

        internal struct OperationScope : IDisposable
        {
            private readonly string _category;
            private readonly string _operation;
            private readonly DateTime _startedUtc;

            internal OperationScope(string category, string operation)
            {
                _category = category;
                _operation = operation ?? string.Empty;
                _startedUtc = DateTime.UtcNow;

                if (!string.IsNullOrWhiteSpace(_operation))
                {
                    Log(_category, "BEGIN " + _operation);
                }
            }

            public void Dispose()
            {
                if (string.IsNullOrWhiteSpace(_operation))
                {
                    return;
                }

                double elapsedMs = (DateTime.UtcNow - _startedUtc).TotalMilliseconds;
                Log(_category, string.Format(CultureInfo.InvariantCulture, "END {0} ({1:0} ms)", _operation, elapsedMs));
            }
        }

        internal static string LogFileFullPath
        {
            get { return GetDailyLogFilePath(DateTime.Now); }
        }

        internal static void SetEnabled(bool enabled)
        {
            lock (SyncRoot)
            {
                _enabled = enabled;
            }
        }

        internal static void SetAnonymization(bool enabled, string serverUrl)
        {
            lock (SyncRoot)
            {
                _anonymizationEnabled = enabled;
                _serverUrlTokens = BuildServerUrlTokens(serverUrl);
            }
        }

        internal static bool IsEnabled
        {
            get
            {
                lock (SyncRoot)
                {
                    return _enabled;
                }
            }
        }

        internal static void Log(string category, string message)
        {
            if (!IsEnabled)
            {
                return;
            }

            WriteLogLine(category, message);
        }

        internal static void LogApi(string message)
        {
            Log(LogCategories.Api, message);
        }

        private static void WriteLogLine(string category, string message)
        {
            try
            {
                lock (SyncRoot)
                {
                    DateTime nowLocal = DateTime.Now;
                    Directory.CreateDirectory(LogDirectory);
                    CleanupLogsIfNeeded(nowLocal);

                    string logFilePath = GetDailyLogFilePath(nowLocal);
                    string sanitizedMessage = SanitizeMessage(message ?? string.Empty);
                    string line = string.Format(
                        CultureInfo.InvariantCulture,
                        "[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1}] {2}",
                        DateTime.UtcNow,
                        string.IsNullOrWhiteSpace(category) ? "GENERAL" : category.Trim().ToUpperInvariant(),
                        sanitizedMessage);
                    File.AppendAllText(logFilePath, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                // Logging must never throw. Fall back to Trace to make failures discoverable.
                try
                {
                    Trace.WriteLine("DiagnosticsLogger failed: " + ex);
                }
                catch (Exception traceEx)
                {
                    Debug.WriteLine("DiagnosticsLogger Trace fallback failed: " + traceEx);
                }
            }
        }

        internal static void Log(string message)
        {
            Log(null, message);
        }

        internal static void LogException(string category, string context, Exception ex)
        {
            // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
            if (ex == null)
            {
                WriteLogLine(category, context);
                return;
            }

            string message = string.IsNullOrWhiteSpace(context) ? "Exception" : context.Trim();
            WriteLogLine(category, message + ": " + ex);
        }

        internal static OperationScope BeginOperation(string category, string operation)
        {
            return new OperationScope(category, operation);
        }

        private static string GetDailyLogFilePath(DateTime nowLocal)
        {
            return Path.Combine(
                LogDirectory,
                DailyLogFilePrefix + nowLocal.ToString(DailyLogDateFormat, CultureInfo.InvariantCulture));
        }

        private static string SanitizeMessage(string message)
        {
            if (string.IsNullOrEmpty(message) || !_anonymizationEnabled)
            {
                return message ?? string.Empty;
            }

            try
            {
                string value = message;
                value = AuthorizationHeaderRegex.Replace(value, "$1<REDACTED>");
                value = UrlCredentialRegex.Replace(value, "$1<CRED>@");
                value = SecretQueryRegex.Replace(value, "$1<REDACTED>");
                value = SecretJsonRegex.Replace(value, "$1\"<REDACTED>\"");
                value = ShareCallTokenRegex.Replace(value, "$1<TOKEN>");
                value = DavUserPathRegex.Replace(value, "$1<USER>");
                value = UserFieldRegex.Replace(value, "$1<REDACTED_USER>");
                value = JsonUserFieldRegex.Replace(value, "$1<REDACTED_USER>$3");
                value = WindowsUserPathRegex.Replace(value, "$1<USER>");
                value = ProviderUserPathRegex.Replace(value, "$1<USER>");
                value = ReplaceEmails(value);
                value = ReplaceServerUrls(value);
                return value;
            }
            catch
            {
                return "<LOG_REDACTION_FAILED>";
            }
        }

        private static string ReplaceServerUrls(string value)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (string.IsNullOrEmpty(value) || _serverUrlTokens == null || _serverUrlTokens.Length == 0)
            {
                return value;
            }

            string redacted = value;
            for (int i = 0; i < _serverUrlTokens.Length; i++)
            {
                string token = _serverUrlTokens[i];
                if (string.IsNullOrEmpty(token))
                {
                    continue;
                }

                redacted = Regex.Replace(redacted, Regex.Escape(token), "<NC_URL>", RegexOptions.IgnoreCase);
            }

            return redacted;
        }

        private static string ReplaceEmails(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }

            return EmailRegex.Replace(value, match =>
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (match == null || string.IsNullOrEmpty(match.Value))
                {
                    return "<EMAIL>";
                }

                string hash = ComputeShortHash(match.Value);
                return "<EMAIL#" + hash + ">";
            });
        }

        private static string ComputeShortHash(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "00000000";
            }

            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(value.ToLowerInvariant());
                byte[] hash = sha.ComputeHash(bytes);
                return BitConverter.ToString(hash, 0, 4).Replace("-", string.Empty);
            }
        }

        private static string[] BuildServerUrlTokens(string serverUrl)
        {
            if (string.IsNullOrWhiteSpace(serverUrl))
            {
                return new string[0];
            }

            var tokens = new List<string>();
            string trimmed = serverUrl.Trim();
            if (trimmed.Length == 0)
            {
                return new string[0];
            }

            string raw = trimmed.TrimEnd('/');
            if (!string.IsNullOrEmpty(raw))
            {
                tokens.Add(raw);
            }

            Uri uri;
            if (Uri.TryCreate(trimmed, UriKind.Absolute, out uri))
            {
                string authority = uri.GetLeftPart(UriPartial.Authority).TrimEnd('/');
                if (!string.IsNullOrEmpty(authority))
                {
                    tokens.Add(authority);
                }

                if (!string.IsNullOrWhiteSpace(uri.Host))
                {
                    tokens.Add(uri.Host);
                }
            }

            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string token in tokens)
            {
                if (!string.IsNullOrWhiteSpace(token))
                {
                    unique.Add(token);
                }
            }

            var ordered = new List<string>(unique);
            ordered.Sort(delegate (string left, string right)
            {
                return right.Length.CompareTo(left.Length);
            });

            return ordered.ToArray();
        }

        private static void CleanupLogsIfNeeded(DateTime nowLocal)
        {
            DateTime todayLocal = nowLocal.Date;
            if (_lastCleanupDateLocal == todayLocal)
            {
                return;
            }

            _lastCleanupDateLocal = todayLocal;
            CleanupLegacyLogFile(nowLocal);
            CleanupDailyLogFiles(nowLocal);
        }

        private static void CleanupLegacyLogFile(DateTime nowLocal)
        {
            try
            {
                if (!File.Exists(LegacyLogFilePath))
                {
                    return;
                }

                DateTime cutoffLocal = nowLocal.AddDays(-CleanupOlderThanDays);
                DateTime lastWriteLocal = File.GetLastWriteTime(LegacyLogFilePath);
                if (lastWriteLocal < cutoffLocal)
                {
                    File.Delete(LegacyLogFilePath);
                }
            }
            catch
            {
                // Cleanup failures must never affect regular logging.
            }
        }

        private static void CleanupDailyLogFiles(DateTime nowLocal)
        {
            try
            {
                string[] files = Directory.GetFiles(LogDirectory, DailyLogFilePrefix + "*");
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (files == null || files.Length == 0)
                {
                    return;
                }

                List<DatedLogFile> datedFiles = new List<DatedLogFile>(files.Length);
                DateTime ageCutoffLocal = nowLocal.Date.AddDays(-CleanupOlderThanDays);

                foreach (string path in files)
                {
                    string fileName = Path.GetFileName(path);
                    if (string.IsNullOrEmpty(fileName) || fileName.Length <= DailyLogFilePrefix.Length)
                    {
                        continue;
                    }

                    string datePart = fileName.Substring(DailyLogFilePrefix.Length);
                    DateTime parsedDateLocal;
                    if (DateTime.TryParseExact(
                            datePart,
                            DailyLogDateFormat,
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out parsedDateLocal))
                    {
                        datedFiles.Add(new DatedLogFile
                        {
                            DateLocal = parsedDateLocal.Date,
                            Path = path
                        });
                        continue;
                    }

                    DateTime lastWriteLocal = File.GetLastWriteTime(path);
                    if (lastWriteLocal < ageCutoffLocal)
                    {
                        File.Delete(path);
                    }
                }

                if (datedFiles.Count == 0)
                {
                    return;
                }

                datedFiles.Sort(delegate (DatedLogFile a, DatedLogFile b)
                {
                    return b.DateLocal.CompareTo(a.DateLocal);
                });

                for (int i = 0; i < datedFiles.Count; i++)
                {
                    bool beyondRetentionWindow = i >= KeepDailyLogCount;
                    bool tooOld = datedFiles[i].DateLocal < ageCutoffLocal;
                    if (beyondRetentionWindow || tooOld)
                    {
                        File.Delete(datedFiles[i].Path);
                    }
                }
            }
            catch
            {
                // Cleanup failures must never affect regular logging.
            }
        }
    }
}
