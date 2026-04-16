/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Diagnostics;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Central utility for opening browser URLs via shell execution.
     */
    internal static class BrowserLauncher
    {
        internal static bool OpenUrl(string url, string logCategory, string failureContext)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(logCategory, failureContext ?? "Failed to open URL.", ex);
                return false;
            }
        }
    }
}

