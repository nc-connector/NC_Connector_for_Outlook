// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class ManagedSetupPolicy
    {
        private const string PolicyKeyPath = @"Software\Policies\NC Connector";
        private const string NextcloudUrlValueName = "NextcloudUrl";
        private const string NextcloudUrlLockedValueName = "NextcloudUrlLocked";

        private ManagedSetupPolicy(string nextcloudUrl, bool nextcloudUrlLocked, string source)
        {
            NextcloudUrl = nextcloudUrl ?? string.Empty;
            NextcloudUrlLocked = nextcloudUrlLocked;
            Source = source ?? string.Empty;
        }

        internal string NextcloudUrl { get; private set; }

        internal bool NextcloudUrlLocked { get; private set; }

        internal string Source { get; private set; }

        internal bool HasNextcloudUrl
        {
            get { return !string.IsNullOrWhiteSpace(NextcloudUrl); }
        }

        internal static ManagedSetupPolicy Load()
        {
            ManagedSetupPolicy machinePolicy = ReadFirstPolicy(RegistryHive.LocalMachine, "HKLM");
            if (machinePolicy != null && machinePolicy.HasNextcloudUrl)
            {
                return machinePolicy;
            }

            ManagedSetupPolicy userPolicy = ReadFirstPolicy(RegistryHive.CurrentUser, "HKCU");
            if (userPolicy != null && userPolicy.HasNextcloudUrl)
            {
                return userPolicy;
            }

            return new ManagedSetupPolicy(string.Empty, false, string.Empty);
        }

        private static ManagedSetupPolicy ReadFirstPolicy(RegistryHive hive, string hiveName)
        {
            foreach (RegistryView view in GetRegistryViews())
            {
                ManagedSetupPolicy policy = ReadPolicy(hive, view, hiveName);
                if (policy != null && policy.HasNextcloudUrl)
                {
                    return policy;
                }
            }

            return null;
        }

        private static IEnumerable<RegistryView> GetRegistryViews()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                yield return RegistryView.Registry64;
            }
            yield return RegistryView.Registry32;
        }

        private static ManagedSetupPolicy ReadPolicy(RegistryHive hive, RegistryView view, string hiveName)
        {
            string source = hiveName + "\\" + PolicyKeyPath + " (" + view + ")";
            try
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (RegistryKey policyKey = baseKey.OpenSubKey(PolicyKeyPath, false))
                {
                    if (policyKey == null)
                    {
                        return null;
                    }

                    string nextcloudUrl = NormalizeNextcloudUrl(policyKey.GetValue(NextcloudUrlValueName));
                    bool locked = ReadBoolean(policyKey.GetValue(NextcloudUrlLockedValueName));

                    if (string.IsNullOrWhiteSpace(nextcloudUrl))
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "Managed Nextcloud URL policy ignored because no valid URL is configured (source=" + source + ").");
                        return null;
                    }

                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Managed Nextcloud URL policy loaded (source=" + source + ", locked=" + locked + ").");
                    return new ManagedSetupPolicy(nextcloudUrl, locked, source);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to read managed Nextcloud URL policy (source=" + source + ").",
                    ex);
                return null;
            }
        }

        private static string NormalizeNextcloudUrl(object rawValue)
        {
            string raw = rawValue as string;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            string trimmed = raw.Trim().TrimEnd('/');
            Uri uri;
            if (!Uri.TryCreate(trimmed, UriKind.Absolute, out uri))
            {
                return string.Empty;
            }
            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return trimmed;
        }

        private static bool ReadBoolean(object rawValue)
        {
            if (rawValue == null)
            {
                return false;
            }
            if (rawValue is int)
            {
                return (int)rawValue != 0;
            }
            if (rawValue is long)
            {
                return (long)rawValue != 0L;
            }
            if (rawValue is bool)
            {
                return (bool)rawValue;
            }

            string converted = Convert.ToString(rawValue);
            string text = converted == null ? string.Empty : converted.Trim();
            return string.Equals(text, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "on", StringComparison.OrdinalIgnoreCase);
        }
    }
}
