/**
 * Copyright (c) 2026 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Live probe for host-side large attachment upload options.
     * Reads environment overrides and registry values on every call.
     */
    internal sealed class OutlookAttachmentAutomationGuardService
    {
        internal sealed class GuardState
        {
            internal GuardState(bool lockActive, int thresholdMb, string source)
            {
                LockActive = lockActive;
                ThresholdMb = NormalizeThresholdMb(thresholdMb);
                Source = source ?? string.Empty;
            }

            internal bool LockActive { get; private set; }

            internal int ThresholdMb { get; private set; }

            internal string Source { get; private set; }
        }

        private const int DefaultThresholdMb = 5;
        private static readonly string[] OfficeVersions = { "16.0", "15.0", "14.0" };
        private static readonly string[] RegistryRelativePaths =
        {
            @"Software\Microsoft\Office\{0}\Outlook\Options\Mail",
            @"Software\Microsoft\Office\{0}\Outlook\Preferences",
            @"Software\Policies\Microsoft\Office\{0}\Outlook\Options\Mail",
            @"Software\Policies\Microsoft\Office\{0}\Outlook\Preferences"
        };

        private static readonly string[] LockValueNames =
        {
            "UploadForLargeAttachmentsEnabled",
            "UploadLargeAttachmentsEnabled",
            "LargeAttachmentUploadEnabled",
            "AutomaticLargeAttachmentUpload"
        };

        private static readonly string[] ThresholdValueNames =
        {
            "UploadForLargeAttachmentsThresholdMb",
            "UploadLargeAttachmentsThresholdMb",
            "LargeAttachmentThresholdMb",
            "LargeAttachmentThreshold"
        };

        internal GuardState ReadLiveState()
        {
            bool envLock;
            int envThreshold;
            if (TryReadEnvironmentOverride(out envLock, out envThreshold))
            {
                return new GuardState(envLock, envThreshold, "env");
            }

            bool lockActive = false;
            int thresholdMb = DefaultThresholdMb;
            string source = "default";

            foreach (var root in new[] { Registry.CurrentUser, Registry.LocalMachine })
            {
                foreach (string version in OfficeVersions)
                {
                    foreach (string pattern in RegistryRelativePaths)
                    {
                        string path = string.Format(pattern, version);
                        int detectedThreshold;
                        bool hasThreshold = TryReadRegistryInteger(root, path, ThresholdValueNames, out detectedThreshold);
                        bool? detectedLock = TryReadRegistryBoolean(root, path, LockValueNames);

                        if (hasThreshold)
                        {
                            thresholdMb = NormalizeThresholdMb(detectedThreshold);
                            source = "registry:" + root.Name + "\\" + path;
                        }

                        if (detectedLock.HasValue)
                        {
                            if (detectedLock.Value)
                            {
                                lockActive = true;
                                source = "registry:" + root.Name + "\\" + path;
                            }
                            else if (!lockActive && hasThreshold)
                            {
                                lockActive = false;
                            }
                        }
                        else if (!lockActive && hasThreshold)
                        {
                            // Some hosts expose only threshold values; treat them as active lock.
                            lockActive = true;
                            source = "registry:" + root.Name + "\\" + path;
                        }
                    }
                }
            }

            return new GuardState(lockActive, thresholdMb, source);
        }

        private static bool TryReadEnvironmentOverride(out bool lockActive, out int thresholdMb)
        {
            lockActive = false;
            thresholdMb = DefaultThresholdMb;

            string envLock = Environment.GetEnvironmentVariable("NC4OL_HOST_ATTACHMENT_LOCK_ACTIVE");
            if (string.IsNullOrWhiteSpace(envLock))
            {
                return false;
            }

            bool parsedLock;
            if (!TryParseBoolean(envLock, out parsedLock))
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Invalid environment override for NC4OL_HOST_ATTACHMENT_LOCK_ACTIVE.");
                return false;
            }

            lockActive = parsedLock;
            string envThreshold = Environment.GetEnvironmentVariable("NC4OL_HOST_ATTACHMENT_THRESHOLD_MB");
            int parsedThreshold;
            if (!string.IsNullOrWhiteSpace(envThreshold) && int.TryParse(envThreshold, out parsedThreshold))
            {
                thresholdMb = NormalizeThresholdMb(parsedThreshold);
            }

            return true;
        }

        private static bool TryReadRegistryInteger(RegistryKey root, string relativePath, string[] valueNames, out int value)
        {
            value = 0;
            // Guard remains tolerant because registry probing is best-effort diagnostics.
            if (root == null || string.IsNullOrWhiteSpace(relativePath) || valueNames == null)
            {
                return false;
            }

            foreach (string valueName in valueNames)
            {
                if (string.IsNullOrWhiteSpace(valueName))
                {
                    continue;
                }

                try
                {
                    object raw = Registry.GetValue(root.Name + "\\" + relativePath, valueName, null);
                    int parsed;                    if (raw != null && int.TryParse(Convert.ToString(raw), out parsed))
                    {
                        value = parsed;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to read integer registry value '" + valueName + "' from '" + root.Name + "\\" + relativePath + "'.",
                        ex);
                }
            }

            return false;
        }

        private static bool? TryReadRegistryBoolean(RegistryKey root, string relativePath, string[] valueNames)
        {
            // Guard remains tolerant because registry probing is best-effort diagnostics.
            if (root == null || string.IsNullOrWhiteSpace(relativePath) || valueNames == null)
            {
                return null;
            }

            foreach (string valueName in valueNames)
            {
                if (string.IsNullOrWhiteSpace(valueName))
                {
                    continue;
                }

                try
                {
                    object raw = Registry.GetValue(root.Name + "\\" + relativePath, valueName, null);                    if (raw == null)
                    {
                        continue;
                    }

                    bool parsed;
                    if (TryParseBoolean(Convert.ToString(raw), out parsed))
                    {
                        return parsed;
                    }

                    int intValue;
                    if (int.TryParse(Convert.ToString(raw), out intValue))
                    {
                        return intValue != 0;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to read boolean registry value '" + valueName + "' from '" + root.Name + "\\" + relativePath + "'.",
                        ex);
                }
            }

            return null;
        }

        private static bool TryParseBoolean(string value, out bool parsed)
        {
            parsed = false;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            if (string.Equals(trimmed, "1", StringComparison.OrdinalIgnoreCase))
            {
                parsed = true;
                return true;
            }
            if (string.Equals(trimmed, "0", StringComparison.OrdinalIgnoreCase))
            {
                parsed = false;
                return true;
            }
            return bool.TryParse(trimmed, out parsed);
        }

        internal static int NormalizeThresholdMb(int thresholdMb)
        {
            return Math.Max(1, thresholdMb <= 0 ? DefaultThresholdMb : thresholdMb);
        }
    }
}


