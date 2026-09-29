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
        private const string ShowMainRibbonTabValueName = "ShowMainRibbonTab";

        private ManagedSetupPolicy(
            object nextcloudUrlValue,
            object nextcloudUrlLockedValue,
            object showMainRibbonTabValue,
            string source)
        {
            HasNextcloudUrlValue = nextcloudUrlValue != null;
            HasNextcloudUrlLockedValue = nextcloudUrlLockedValue != null;
            HasShowMainRibbonTabValue = showMainRibbonTabValue != null;
            NextcloudUrl = NormalizeNextcloudUrl(nextcloudUrlValue);
            NextcloudUrlLocked = ReadBoolean(nextcloudUrlLockedValue);
            bool showMainRibbonTab;
            ShowMainRibbonTab = !TryReadBoolean(showMainRibbonTabValue, out showMainRibbonTab)
                || showMainRibbonTab;
            IsEnterpriseRollout = HasNextcloudUrlValue || HasNextcloudUrlLockedValue || HasShowMainRibbonTabValue;
            Source = source ?? string.Empty;
        }

        internal string NextcloudUrl { get; private set; }

        internal bool NextcloudUrlLocked { get; private set; }

        internal string Source { get; private set; }

        internal bool IsEnterpriseRollout { get; private set; }

        internal bool ShowMainRibbonTab { get; private set; }

        private bool HasNextcloudUrlValue { get; set; }

        private bool HasNextcloudUrlLockedValue { get; set; }

        private bool HasShowMainRibbonTabValue { get; set; }

        internal bool HasNextcloudUrl
        {
            get { return !string.IsNullOrWhiteSpace(NextcloudUrl); }
        }

        internal static ManagedSetupPolicy Load()
        {
            var policies = new List<ManagedSetupPolicy>();
            foreach (RegistryHive hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                foreach (RegistryView view in GetRegistryViews())
                {
                    policies.Add(ReadPolicy(hive, view, hive == RegistryHive.LocalMachine ? "HKLM" : "HKCU"));
                }
            }
            return Resolve(policies);
        }

        internal static ManagedSetupPolicy Resolve(IEnumerable<ManagedSetupPolicy> policies)
        {
            var result = new ManagedSetupPolicy(null, null, null, string.Empty);
            if (policies == null)
            {
                return result;
            }

            foreach (ManagedSetupPolicy policy in policies)
            {
                if (policy == null)
                {
                    continue;
                }
                result.IsEnterpriseRollout |= policy.IsEnterpriseRollout;
                result.HasNextcloudUrlLockedValue |= policy.HasNextcloudUrlLockedValue;
                if (!result.HasNextcloudUrlValue && policy.HasNextcloudUrlValue)
                {
                    result.HasNextcloudUrlValue = true;
                    result.NextcloudUrl = policy.NextcloudUrl;
                    result.NextcloudUrlLocked = policy.NextcloudUrlLocked;
                    result.Source = policy.Source;
                }
                if (!result.HasShowMainRibbonTabValue && policy.HasShowMainRibbonTabValue)
                {
                    result.HasShowMainRibbonTabValue = true;
                    result.ShowMainRibbonTab = policy.ShowMainRibbonTab;
                }
            }
            return result;
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

                    var policy = new ManagedSetupPolicy(
                        policyKey.GetValue(NextcloudUrlValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(NextcloudUrlLockedValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(ShowMainRibbonTabValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        source);

                    if (!policy.IsEnterpriseRollout)
                    {
                        return null;
                    }

                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Managed setup policy loaded (source=" + source
                        + ", urlPresent=" + policy.HasNextcloudUrlValue
                        + ", urlValid=" + policy.HasNextcloudUrl
                        + ", urlLockValuePresent=" + policy.HasNextcloudUrlLockedValue
                        + ", locked=" + policy.NextcloudUrlLocked
                        + ", ribbonValuePresent=" + policy.HasShowMainRibbonTabValue
                        + ", showMainRibbonTab=" + policy.ShowMainRibbonTab + ").");
                    return policy;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to read managed setup policy (source=" + source + ").",
                    ex);
                return null;
            }
        }

        private static string NormalizeNextcloudUrl(object rawValue)
        {
            string raw = rawValue as string;
            string normalized;
            return NextcloudUriValidator.TryNormalizeBaseUrl(raw, out normalized)
                ? normalized
                : string.Empty;
        }

        private static bool ReadBoolean(object rawValue)
        {
            bool value;
            return TryReadBoolean(rawValue, out value) && value;
        }

        private static bool TryReadBoolean(object rawValue, out bool value)
        {
            value = false;
            if (rawValue == null)
            {
                return false;
            }
            if (rawValue is int)
            {
                value = (int)rawValue != 0;
                return true;
            }
            if (rawValue is long)
            {
                value = (long)rawValue != 0L;
                return true;
            }
            if (rawValue is bool)
            {
                value = (bool)rawValue;
                return true;
            }

            string converted = Convert.ToString(rawValue);
            string text = converted == null ? string.Empty : converted.Trim();
            value = string.Equals(text, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "on", StringComparison.OrdinalIgnoreCase);
            return value
                || string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "false", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "no", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "off", StringComparison.OrdinalIgnoreCase);
        }
    }
}
