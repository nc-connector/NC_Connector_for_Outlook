// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class ManagedSetupPolicy
    {
        private const string PolicyKeyPath = @"Software\Policies\NC Connector";
        private const string NextcloudUrlValueName = "NextcloudUrl";
        private const string NextcloudUrlLockedValueName = "NextcloudUrlLocked";
        private const string ShowMainRibbonTabValueName = "ShowMainRibbonTab";
        private const string TransportTlsUseSystemDefaultValueName = "TransportTlsUseSystemDefault";
        private const string TransportTlsEnable12ValueName = "TransportTlsEnable12";
        private const string TransportTlsEnable13ValueName = "TransportTlsEnable13";
        private const string DebugLoggingEnabledValueName = "DebugLoggingEnabled";
        private const string LogAnonymizationEnabledValueName = "LogAnonymizationEnabled";
        private const string UpdateNotifyEnabledValueName = "UpdateNotifyEnabled";
        private const string IfbEnabledValueName = "IfbEnabled";
        private const string IfbDaysValueName = "IfbDays";
        private const string IfbCacheHoursValueName = "IfbCacheHours";
        private const string IfbPortValueName = "IfbPort";
        private const string DefaultsSourceValueName = "DefaultsSource";
        private static readonly object InvalidRegistryValue = new object();
        private bool _tlsSystemValueValid;
        private bool _tls12ValueValid;
        private bool _tls13ValueValid;
        private bool _debugLoggingValueValid;
        private bool _logAnonymizationValueValid;
        private bool _updateNotifyValueValid;
        private bool _ifbEnabledValueValid;
        private bool _ifbDaysValueValid;
        private bool _ifbCacheHoursValueValid;
        private bool _ifbPortValueValid;
        private bool _defaultsSourceValueValid;

        private ManagedSetupPolicy(
            object nextcloudUrlValue,
            object nextcloudUrlLockedValue,
            object showMainRibbonTabValue,
            object transportTlsUseSystemDefaultValue,
            object transportTlsEnable12Value,
            object transportTlsEnable13Value,
            object debugLoggingEnabledValue,
            object logAnonymizationEnabledValue,
            object updateNotifyEnabledValue,
            object ifbEnabledValue,
            object ifbDaysValue,
            object ifbCacheHoursValue,
            object ifbPortValue,
            object defaultsSourceValue,
            string source)
        {
            HasNextcloudUrlValue = nextcloudUrlValue != null;
            HasNextcloudUrlLockedValue = nextcloudUrlLockedValue != null;
            HasShowMainRibbonTabValue = showMainRibbonTabValue != null;
            HasTransportTlsUseSystemDefaultValue = transportTlsUseSystemDefaultValue != null;
            HasTransportTlsEnable12Value = transportTlsEnable12Value != null;
            HasTransportTlsEnable13Value = transportTlsEnable13Value != null;
            HasDebugLoggingEnabledValue = debugLoggingEnabledValue != null;
            HasLogAnonymizationEnabledValue = logAnonymizationEnabledValue != null;
            HasUpdateNotifyPolicy = updateNotifyEnabledValue != null;
            HasIfbEnabledValue = ifbEnabledValue != null;
            HasIfbDaysValue = ifbDaysValue != null;
            HasIfbCacheHoursValue = ifbCacheHoursValue != null;
            HasIfbPortValue = ifbPortValue != null;
            HasDefaultsSourcePolicy = defaultsSourceValue != null;
            string defaultsSource = BackendPolicyStatus.NormalizeDefaultsSource(defaultsSourceValue as string);
            _defaultsSourceValueValid = defaultsSourceValue == null || defaultsSource != null;
            DefaultsSource = defaultsSource ?? "local";
            IfbEnabled = ReadPolicyBoolean(ifbEnabledValue, false, out _ifbEnabledValueValid);
            IfbDays = ReadPolicyDword(ifbDaysValue, AddinSettings.DefaultIfbDays,
                value => value == 10 || value == 30 || value == 60 || value == 90, out _ifbDaysValueValid);
            IfbCacheHours = ReadPolicyDword(ifbCacheHoursValue, AddinSettings.DefaultIfbCacheHours,
                value => value >= 1 && value <= 24, out _ifbCacheHoursValueValid);
            IfbPort = ReadPolicyDword(ifbPortValue, AddinSettings.DefaultIfbPort,
                value => value >= AddinSettings.MinIfbPort && value <= AddinSettings.MaxIfbPort, out _ifbPortValueValid);
            UpdateNotifyEnabled = ReadPolicyBoolean(updateNotifyEnabledValue, false, out _updateNotifyValueValid);
            TransportTlsUseSystemDefault = ReadPolicyBoolean(transportTlsUseSystemDefaultValue, false, out _tlsSystemValueValid);
            TransportTlsEnable12 = ReadPolicyBoolean(transportTlsEnable12Value, true, out _tls12ValueValid);
            TransportTlsEnable13 = ReadPolicyBoolean(transportTlsEnable13Value, false, out _tls13ValueValid);
            DebugLoggingEnabled = ReadPolicyBoolean(debugLoggingEnabledValue, false, out _debugLoggingValueValid);
            LogAnonymizationEnabled = ReadPolicyBoolean(logAnonymizationEnabledValue, true, out _logAnonymizationValueValid);
            if (!_logAnonymizationValueValid)
            {
                LogAnonymizationEnabled = true;
            }
            NextcloudUrl = NormalizeNextcloudUrl(nextcloudUrlValue);
            NextcloudUrlLocked = ReadBoolean(nextcloudUrlLockedValue);
            bool showMainRibbonTab;
            ShowMainRibbonTab = !TryReadBoolean(showMainRibbonTabValue, out showMainRibbonTab)
                || showMainRibbonTab;
            IsEnterpriseRollout = HasNextcloudUrlValue || HasNextcloudUrlLockedValue || HasShowMainRibbonTabValue || HasTransportTlsPolicy || HasLoggingPolicy || HasUpdateNotifyPolicy || HasIfbPolicy || HasDefaultsSourcePolicy;
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

        private bool HasTransportTlsUseSystemDefaultValue { get; set; }
        private bool HasTransportTlsEnable12Value { get; set; }
        private bool HasTransportTlsEnable13Value { get; set; }
        private bool HasDebugLoggingEnabledValue { get; set; }
        private bool HasLogAnonymizationEnabledValue { get; set; }
        private bool HasIfbEnabledValue { get; set; }
        private bool HasIfbDaysValue { get; set; }
        private bool HasIfbCacheHoursValue { get; set; }
        private bool HasIfbPortValue { get; set; }

        internal bool DebugLoggingEnabled { get; private set; }
        internal bool LogAnonymizationEnabled { get; private set; }
        internal bool HasLoggingPolicy { get { return HasDebugLoggingEnabledValue || HasLogAnonymizationEnabledValue; } }
        internal bool IsLoggingPolicyValid { get { return _debugLoggingValueValid && _logAnonymizationValueValid; } }
        internal bool UpdateNotifyEnabled { get; private set; }
        internal bool HasUpdateNotifyPolicy { get; private set; }
        internal bool IsUpdateNotifyPolicyValid { get { return _updateNotifyValueValid; } }

        internal string DefaultsSource { get; private set; }
        internal bool HasDefaultsSourcePolicy { get; private set; }
        internal bool IsDefaultsSourcePolicyValid { get { return _defaultsSourceValueValid; } }

        internal bool IfbEnabled { get; private set; }
        internal int IfbDays { get; private set; }
        internal int IfbCacheHours { get; private set; }
        internal int IfbPort { get; private set; }
        internal bool HasIfbPolicy
        {
            get { return HasIfbEnabledValue || HasIfbDaysValue || HasIfbCacheHoursValue || HasIfbPortValue; }
        }
        internal bool IsIfbPolicyValid
        {
            get { return _ifbEnabledValueValid && _ifbDaysValueValid && _ifbCacheHoursValueValid && _ifbPortValueValid; }
        }

        internal bool TransportTlsUseSystemDefault { get; private set; }
        internal bool TransportTlsEnable12 { get; private set; }
        internal bool TransportTlsEnable13 { get; private set; }

        internal bool HasTransportTlsPolicy
        {
            get { return HasTransportTlsUseSystemDefaultValue || HasTransportTlsEnable12Value || HasTransportTlsEnable13Value; }
        }

        internal bool IsTransportTlsPolicyValid
        {
            get
            {
                return _tlsSystemValueValid && _tls12ValueValid && _tls13ValueValid
                    && (TransportTlsUseSystemDefault || TransportTlsEnable12 || TransportTlsEnable13);
            }
        }

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
            var result = new ManagedSetupPolicy(null, null, null, null, null, null, null, null, null, null, null, null, null, null, string.Empty);
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
                if (!result.HasTransportTlsUseSystemDefaultValue && policy.HasTransportTlsUseSystemDefaultValue)
                {
                    result.HasTransportTlsUseSystemDefaultValue = true;
                    result.TransportTlsUseSystemDefault = policy.TransportTlsUseSystemDefault;
                    result._tlsSystemValueValid = policy._tlsSystemValueValid;
                }
                if (!result.HasTransportTlsEnable12Value && policy.HasTransportTlsEnable12Value)
                {
                    result.HasTransportTlsEnable12Value = true;
                    result.TransportTlsEnable12 = policy.TransportTlsEnable12;
                    result._tls12ValueValid = policy._tls12ValueValid;
                }
                if (!result.HasTransportTlsEnable13Value && policy.HasTransportTlsEnable13Value)
                {
                    result.HasTransportTlsEnable13Value = true;
                    result.TransportTlsEnable13 = policy.TransportTlsEnable13;
                    result._tls13ValueValid = policy._tls13ValueValid;
                }
                if (!result.HasDebugLoggingEnabledValue && policy.HasDebugLoggingEnabledValue)
                {
                    result.HasDebugLoggingEnabledValue = true;
                    result.DebugLoggingEnabled = policy.DebugLoggingEnabled;
                    result._debugLoggingValueValid = policy._debugLoggingValueValid;
                }
                if (!result.HasLogAnonymizationEnabledValue && policy.HasLogAnonymizationEnabledValue)
                {
                    result.HasLogAnonymizationEnabledValue = true;
                    result.LogAnonymizationEnabled = policy.LogAnonymizationEnabled;
                    result._logAnonymizationValueValid = policy._logAnonymizationValueValid;
                }
                if (!result.HasUpdateNotifyPolicy && policy.HasUpdateNotifyPolicy)
                {
                    result.HasUpdateNotifyPolicy = true;
                    result.UpdateNotifyEnabled = policy.UpdateNotifyEnabled;
                    result._updateNotifyValueValid = policy._updateNotifyValueValid;
                }
                if (!result.HasIfbEnabledValue && policy.HasIfbEnabledValue)
                {
                    result.HasIfbEnabledValue = true;
                    result.IfbEnabled = policy.IfbEnabled;
                    result._ifbEnabledValueValid = policy._ifbEnabledValueValid;
                }
                if (!result.HasIfbDaysValue && policy.HasIfbDaysValue)
                {
                    result.HasIfbDaysValue = true;
                    result.IfbDays = policy.IfbDays;
                    result._ifbDaysValueValid = policy._ifbDaysValueValid;
                }
                if (!result.HasIfbCacheHoursValue && policy.HasIfbCacheHoursValue)
                {
                    result.HasIfbCacheHoursValue = true;
                    result.IfbCacheHours = policy.IfbCacheHours;
                    result._ifbCacheHoursValueValid = policy._ifbCacheHoursValueValid;
                }
                if (!result.HasIfbPortValue && policy.HasIfbPortValue)
                {
                    result.HasIfbPortValue = true;
                    result.IfbPort = policy.IfbPort;
                    result._ifbPortValueValid = policy._ifbPortValueValid;
                }
                if (!result.HasDefaultsSourcePolicy && policy.HasDefaultsSourcePolicy)
                {
                    result.HasDefaultsSourcePolicy = true;
                    result.DefaultsSource = policy.DefaultsSource;
                    result._defaultsSourceValueValid = policy._defaultsSourceValueValid;
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
                        policyKey.GetValue(TransportTlsUseSystemDefaultValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(TransportTlsEnable12ValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(TransportTlsEnable13ValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(DebugLoggingEnabledValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(LogAnonymizationEnabledValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        policyKey.GetValue(UpdateNotifyEnabledValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        ReadIfbPolicyValue(policyKey, IfbEnabledValueName, false),
                        ReadIfbPolicyValue(policyKey, IfbDaysValueName, true),
                        ReadIfbPolicyValue(policyKey, IfbCacheHoursValueName, true),
                        ReadIfbPolicyValue(policyKey, IfbPortValueName, true),
                        ReadDefaultsSourcePolicyValue(policyKey),
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
                        + ", showMainRibbonTab=" + policy.ShowMainRibbonTab
                        + ", tlsPolicyPresent=" + policy.HasTransportTlsPolicy
                        + ", tlsPolicyValid=" + policy.IsTransportTlsPolicyValid
                        + ", loggingPolicyPresent=" + policy.HasLoggingPolicy
                        + ", loggingPolicyValid=" + policy.IsLoggingPolicyValid
                        + ", updateNotifyPolicyPresent=" + policy.HasUpdateNotifyPolicy
                        + ", updateNotifyPolicyValid=" + policy.IsUpdateNotifyPolicyValid
                        + ", ifbPolicyPresent=" + policy.HasIfbPolicy
                        + ", ifbPolicyValid=" + policy.IsIfbPolicyValid
                        + ", defaultsSourcePolicyPresent=" + policy.HasDefaultsSourcePolicy
                        + ", defaultsSourcePolicyValid=" + policy.IsDefaultsSourcePolicyValid + ").");
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

        private static object ReadIfbPolicyValue(RegistryKey policyKey, string valueName, bool requireDword)
        {
            if (!Array.Exists(policyKey.GetValueNames(), name => string.Equals(name, valueName, StringComparison.OrdinalIgnoreCase)))
            {
                return null;
            }

            RegistryValueKind kind = policyKey.GetValueKind(valueName);
            if (requireDword ? kind != RegistryValueKind.DWord
                : kind != RegistryValueKind.DWord && kind != RegistryValueKind.QWord
                    && kind != RegistryValueKind.String && kind != RegistryValueKind.ExpandString)
            {
                return InvalidRegistryValue;
            }
            return policyKey.GetValue(valueName, InvalidRegistryValue, RegistryValueOptions.DoNotExpandEnvironmentNames)
                ?? InvalidRegistryValue;
        }

        private static object ReadDefaultsSourcePolicyValue(RegistryKey policyKey)
        {
            if (!Array.Exists(policyKey.GetValueNames(), name => string.Equals(name, DefaultsSourceValueName, StringComparison.OrdinalIgnoreCase)))
            {
                return null;
            }
            if (policyKey.GetValueKind(DefaultsSourceValueName) != RegistryValueKind.String)
            {
                return InvalidRegistryValue;
            }
            return policyKey.GetValue(DefaultsSourceValueName, InvalidRegistryValue, RegistryValueOptions.DoNotExpandEnvironmentNames)
                ?? InvalidRegistryValue;
        }

        private static int ReadPolicyDword(object rawValue, int defaultValue, Predicate<int> allowedValue, out bool valid)
        {
            valid = rawValue == null || (rawValue is int && allowedValue((int)rawValue));
            return rawValue != null && valid ? (int)rawValue : defaultValue;
        }

        private static bool ReadBoolean(object rawValue)
        {
            bool value;
            return TryReadBoolean(rawValue, out value) && value;
        }

        private static bool ReadPolicyBoolean(object rawValue, bool defaultValue, out bool valid)
        {
            if (rawValue == null)
            {
                valid = true;
                return defaultValue;
            }
            bool value;
            valid = TryReadBoolean(rawValue, out value);
            return value;
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
