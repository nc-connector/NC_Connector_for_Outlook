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
using System.Text;

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
        private static DateTime _lastCleanupDateLocal = DateTime.MinValue.Date;
        private static bool _enabled;

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
                    string line = string.Format(
                        CultureInfo.InvariantCulture,
                        "[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1}] {2}",
                        DateTime.UtcNow,
                        string.IsNullOrWhiteSpace(category) ? "GENERAL" : category.Trim().ToUpperInvariant(),
                        message ?? string.Empty);
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
