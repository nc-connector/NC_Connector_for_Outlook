/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn
{
    /**
     * Logging helpers and throttled log guards used across the add-in runtime.
     */
    public sealed partial class NextcloudTalkAddIn
    {
        private static void LogCore(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Core, message);
        }

        private static void LogSettings(string message)
        {
            LogCore(message);
        }

        private static void LogTalk(string message)
        {
            DiagnosticsLogger.Log(LogCategories.Talk, message);
        }

        private static void LogFileLink(string message)
        {
            DiagnosticsLogger.Log(LogCategories.FileLink, message);
        }

        internal static void LogTalkMessage(string message)
        {
            LogTalk(message);
        }

        internal static void LogFileLinkMessage(string message)
        {
            LogFileLink(message);
        }

        private void LogDeferredAppointmentEnsureRestriction(string message)
        {
            string throttledMessage;
            if (!_deferredAppointmentEnsureState.ShouldLogRestriction(message, DateTime.UtcNow, out throttledMessage))
            {
                return;
            }

            DiagnosticsLogger.Log(LogCategories.Core, throttledMessage);
        }
    }
}
