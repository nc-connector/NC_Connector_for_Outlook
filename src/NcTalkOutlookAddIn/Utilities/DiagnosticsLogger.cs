/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
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
        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NextcloudTalkOutlookAddInData");
        private static readonly string LogFilePath = Path.Combine(LogDirectory, "addin-runtime.log");
        private static bool _enabled;

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
            get { return LogFilePath; }
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

            try
            {
                lock (SyncRoot)
                {
                    Directory.CreateDirectory(LogDirectory);
                    string line = string.Format(
                        CultureInfo.InvariantCulture,
                        "[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1}] {2}",
                        DateTime.UtcNow,
                        string.IsNullOrWhiteSpace(category) ? "GENERAL" : category.Trim().ToUpperInvariant(),
                        message ?? string.Empty);
                    File.AppendAllText(LogFilePath, line + Environment.NewLine, Encoding.UTF8);
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
                Log(category, context);
                return;
            }

            string message = string.IsNullOrWhiteSpace(context) ? "Exception" : context.Trim();
            Log(category, message + ": " + ex);
        }

        internal static OperationScope BeginOperation(string category, string operation)
        {
            return new OperationScope(category, operation);
        }
    }
}
