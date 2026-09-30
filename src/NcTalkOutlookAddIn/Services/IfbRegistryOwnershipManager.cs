// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Owns the two per-user Outlook IFB values and restores only unchanged writes.
    internal sealed class IfbRegistryOwnershipManager
    {
        private const string CalendarValueName = "FreeBusySearchPath";
        private const string InternetValueName = "Read URL";

        private readonly IfbRegistryStateStore _stateStore;
        private IfbRegistryState _state;

        internal IfbRegistryOwnershipManager(string dataDirectory, string profileScope)
        {
            _stateStore = new IfbRegistryStateStore(dataDirectory, profileScope);
            _state = _stateStore.Load() ?? new IfbRegistryState();
            if (_state.Ownership == null)
            {
                _state.Ownership = new List<IfbRegistryOwnership>();
            }
        }

        internal void Apply(string outlookVersion, string desired, AddinSettings settings)
        {
            RegistryTarget[] targets = BuildTargets(outlookVersion);
            foreach (RegistryTarget target in targets)
            {
                ThrowIfPolicyConflicts(target, desired);
            }

            AdoptLegacyOwnership(targets, settings);
            foreach (RegistryTarget target in targets)
            {
                ApplyTarget(target, desired);
            }
        }

        internal void Restore()
        {
            if (_state.Ownership.Count == 0)
            {
                return;
            }

            var remaining = new List<IfbRegistryOwnership>();
            foreach (IfbRegistryOwnership ownership in _state.Ownership)
            {
                try
                {
                    RegistryValueSnapshot current = ReadValue(
                        ownership.RegistryPath,
                        ownership.ValueName);
                    if (current.Exists && !MatchesWritten(current, ownership))
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Ifb,
                            "IFB restore skipped after an external registry change.");
                        continue;
                    }

                    using (RegistryKey key = OpenOrCreateUserKey(ownership.RegistryPath))
                    {
                        if (ownership.OriginalExists)
                        {
                            key.SetValue(
                                ownership.ValueName,
                                ownership.OriginalValue ?? string.Empty,
                                ParseStringKind(ownership.OriginalKind));
                        }
                        else
                        {
                            key.DeleteValue(ownership.ValueName, false);
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "Failed to restore an owned Outlook IFB value.",
                        ex);
                    remaining.Add(ownership);
                }
            }
            _state.Ownership = remaining;
            _stateStore.Save(_state);
        }

        private void AdoptLegacyOwnership(
            RegistryTarget[] targets,
            AddinSettings settings)
        {
            string previous = settings == null
                ? string.Empty
                : (settings.IfbPreviousFreeBusyPath ?? string.Empty).Trim();
            bool hasRecoverablePrevious =
                !string.IsNullOrEmpty(previous)
                && !IsNcConnectorIfbUrl(previous);

            bool changed = false;
            foreach (RegistryTarget target in targets)
            {
                if (FindOwnership(target) != null)
                {
                    continue;
                }
                RegistryValueSnapshot current = ReadValue(
                    target.UserPath,
                    target.ValueName);
                if (!current.Exists
                    || !IsLegacyNcConnectorIfbUrl(current.Value))
                {
                    continue;
                }

                // A pre-3.3.1 endpoint proves an upgrade. Without a saved external
                // value, removing it on disable is safer than restoring stale Connector data.
                _state.Ownership.Add(new IfbRegistryOwnership
                {
                    RegistryPath = target.UserPath,
                    ValueName = target.ValueName,
                    OriginalExists = hasRecoverablePrevious,
                    OriginalKind = (int)RegistryValueKind.String,
                    OriginalValue = hasRecoverablePrevious
                        ? previous
                        : string.Empty,
                    WrittenValue = current.Value
                });
                changed = true;
            }

            // Both user values are inspected before the legacy field is retired.
            if (changed)
            {
                _stateStore.Save(_state);
                DiagnosticsLogger.Log(
                    LogCategories.Ifb,
                    "Migrated legacy NC Connector IFB registry values to the ownership journal.");
                if (settings != null)
                {
                    settings.IfbPreviousFreeBusyPath = string.Empty;
                }
            }
        }

        private void ApplyTarget(RegistryTarget target, string desired)
        {
            RegistryValueSnapshot current = ReadValue(
                target.UserPath,
                target.ValueName);
            IfbRegistryOwnership ownership = FindOwnership(target);
            if (ownership != null)
            {
                // MSI maintenance removes our endpoint but retains its saved predecessor.
                if (current.Exists && !MatchesWritten(current, ownership))
                {
                    _state.Ownership.Remove(ownership);
                    _stateStore.Save(_state);
                    throw new InvalidOperationException(
                        "An Outlook IFB value was changed outside NC Connector and was not overwritten.");
                }
                if (string.Equals(current.Value, desired, StringComparison.Ordinal))
                {
                    return;
                }

                string previousWritten = ownership.WrittenValue;
                ownership.WrittenValue = desired;
                _stateStore.Save(_state);
                try
                {
                    WriteStringValue(target, desired);
                }
                catch
                {
                    ownership.WrittenValue = previousWritten;
                    _stateStore.Save(_state);
                    throw;
                }
                return;
            }

            if (current.Exists && IsNcConnectorIfbUrl(current.Value))
            {
                throw new InvalidOperationException(
                    "An unowned NC Connector IFB value already exists and was not overwritten.");
            }
            if (current.Exists
                && current.Kind != RegistryValueKind.String
                && current.Kind != RegistryValueKind.ExpandString)
            {
                throw new InvalidOperationException(
                    "The existing Outlook IFB value has an unsupported registry type.");
            }

            ownership = new IfbRegistryOwnership
            {
                RegistryPath = target.UserPath,
                ValueName = target.ValueName,
                OriginalExists = current.Exists,
                OriginalKind = current.Exists
                    ? (int)current.Kind
                    : (int)RegistryValueKind.String,
                OriginalValue = current.Exists ? current.Value : string.Empty,
                WrittenValue = desired
            };
            _state.Ownership.Add(ownership);
            _stateStore.Save(_state);
            try
            {
                WriteStringValue(target, desired);
            }
            catch
            {
                _state.Ownership.Remove(ownership);
                _stateStore.Save(_state);
                throw;
            }
        }

        private static void ThrowIfPolicyConflicts(
            RegistryTarget target,
            string desired)
        {
            RegistryValueSnapshot policy = ReadValue(
                target.PolicyPath,
                target.ValueName);
            if (policy.Exists
                && !string.IsNullOrWhiteSpace(policy.Value)
                && !string.Equals(policy.Value, desired, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Outlook Free/Busy is managed by policy. NC Connector did not change the policy value.");
            }
        }

        private IfbRegistryOwnership FindOwnership(RegistryTarget target)
        {
            foreach (IfbRegistryOwnership ownership in _state.Ownership)
            {
                if (string.Equals(
                        target.UserPath,
                        ownership.RegistryPath,
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(
                        target.ValueName,
                        ownership.ValueName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return ownership;
                }
            }
            return null;
        }

        private static bool MatchesWritten(
            RegistryValueSnapshot current,
            IfbRegistryOwnership ownership)
        {
            return current.Exists
                   && string.Equals(
                       current.Value,
                       ownership.WrittenValue ?? string.Empty,
                       StringComparison.Ordinal);
        }

        private static RegistryValueSnapshot ReadValue(
            string path,
            string valueName)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, false))
            {
                if (key == null)
                {
                    return new RegistryValueSnapshot();
                }
                object raw = key.GetValue(
                    valueName,
                    null,
                    RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (raw == null)
                {
                    return new RegistryValueSnapshot();
                }
                return new RegistryValueSnapshot
                {
                    Exists = true,
                    Kind = key.GetValueKind(valueName),
                    Value = Convert.ToString(
                        raw,
                        CultureInfo.InvariantCulture) ?? string.Empty
                };
            }
        }

        private static void WriteStringValue(RegistryTarget target, string value)
        {
            using (RegistryKey key = OpenOrCreateUserKey(target.UserPath))
            {
                key.SetValue(
                    target.ValueName,
                    value ?? string.Empty,
                    RegistryValueKind.String);
            }
        }

        private static RegistryKey OpenOrCreateUserKey(string path)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true);
            return key ?? Registry.CurrentUser.CreateSubKey(path);
        }

        private static RegistryValueKind ParseStringKind(int rawKind)
        {
            return (RegistryValueKind)rawKind == RegistryValueKind.ExpandString
                ? RegistryValueKind.ExpandString
                : RegistryValueKind.String;
        }

        private static RegistryTarget[] BuildTargets(string outlookVersion)
        {
            string version = string.IsNullOrWhiteSpace(outlookVersion)
                ? "16.0"
                : outlookVersion.Trim();
            string userCalendar =
                @"Software\Microsoft\Office\" + version
                + @"\Outlook\Options\Calendar";
            string policyCalendar =
                @"Software\Policies\Microsoft\Office\" + version
                + @"\Outlook\Options\Calendar";
            return new[]
            {
                new RegistryTarget(
                    userCalendar,
                    policyCalendar,
                    CalendarValueName),
                new RegistryTarget(
                    userCalendar + @"\Internet Free/Busy",
                    policyCalendar + @"\Internet Free/Busy",
                    InternetValueName)
            };
        }

        internal static bool IsNcConnectorIfbUrl(string value)
        {
            Uri uri;
            return IfbRegistryEndpoint.TryGetConnectorUri(value, out uri);
        }

        internal static bool IsLegacyNcConnectorIfbUrl(string value)
        {
            Uri uri;
            return IfbRegistryEndpoint.TryGetConnectorUri(value, out uri)
                   && uri.AbsolutePath.StartsWith(
                       "/nc-ifb/freebusy/",
                       StringComparison.OrdinalIgnoreCase);
        }

        private sealed class RegistryTarget
        {
            internal RegistryTarget(
                string userPath,
                string policyPath,
                string valueName)
            {
                UserPath = userPath;
                PolicyPath = policyPath;
                ValueName = valueName;
            }

            internal string UserPath { get; private set; }

            internal string PolicyPath { get; private set; }

            internal string ValueName { get; private set; }
        }

        private sealed class RegistryValueSnapshot
        {
            internal bool Exists { get; set; }

            internal RegistryValueKind Kind { get; set; }

            internal string Value { get; set; }
        }
    }
}
