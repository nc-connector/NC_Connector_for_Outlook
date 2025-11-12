/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Settings
{
    /**
     * Persistente Einstellungen des Add-ins (Anmeldedaten, Filelink/IFB-Optionen etc.).
     */
    internal class AddinSettings
    {
        public AddinSettings()
        {
            ServerUrl = string.Empty;
            Username = string.Empty;
            AppPassword = string.Empty;
            AuthMode = AuthenticationMode.LoginFlow;
            OutlookMuzzleEnabled = true;
            IfbEnabled = false;
            IfbDays = 30;
            IfbCacheHours = 24;
            IfbPreviousFreeBusyPath = string.Empty;
            DebugLoggingEnabled = false;
            LastKnownServerVersion = string.Empty;
            FileLinkBasePath = "90 Freigaben - extern";
            OutlookMuzzleAccounts = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        }

        public string ServerUrl { get; set; }

        public string Username { get; set; }

        public string AppPassword { get; set; }

        public AuthenticationMode AuthMode { get; set; }

        public bool OutlookMuzzleEnabled { get; set; }

        public Dictionary<string, bool> OutlookMuzzleAccounts { get; set; }

        public bool IfbEnabled { get; set; }

        public int IfbDays { get; set; }

        public int IfbCacheHours { get; set; }

        public string IfbPreviousFreeBusyPath { get; set; }

        public bool DebugLoggingEnabled { get; set; }

        public string LastKnownServerVersion { get; set; }

        public string FileLinkBasePath { get; set; }

        public AddinSettings Clone()
        {
            var copy = (AddinSettings)MemberwiseClone();
            if (OutlookMuzzleAccounts != null)
            {
                copy.OutlookMuzzleAccounts = new Dictionary<string, bool>(OutlookMuzzleAccounts, StringComparer.OrdinalIgnoreCase);
            }
            return copy;
        }

        public bool IsMuzzleEnabledForAccount(string accountKey)
        {
            if (!OutlookMuzzleEnabled)
            {
                return false;
            }

            bool hasPerAccount = OutlookMuzzleAccounts != null && OutlookMuzzleAccounts.Count > 0;
            if (!hasPerAccount)
            {
                return OutlookMuzzleEnabled;
            }

            if (string.IsNullOrWhiteSpace(accountKey))
            {
                return false;
            }

            bool enabled;
            if (OutlookMuzzleAccounts.TryGetValue(accountKey, out enabled))
            {
                return enabled;
            }

            return false;
        }

        public void SetMuzzleStateForAccount(string accountKey, bool enabled)
        {
            if (string.IsNullOrWhiteSpace(accountKey))
            {
                return;
            }

            if (OutlookMuzzleAccounts == null)
            {
                OutlookMuzzleAccounts = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            }

            OutlookMuzzleAccounts[accountKey] = enabled;
        }

        public bool HasAnyMuzzleAccountEnabled()
        {
            if (!OutlookMuzzleEnabled)
            {
                return false;
            }

            if (OutlookMuzzleAccounts == null || OutlookMuzzleAccounts.Count == 0)
            {
                return OutlookMuzzleEnabled;
            }

            foreach (var pair in OutlookMuzzleAccounts)
            {
                if (pair.Value)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
