/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Text;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Coordinates the local IFB server and Outlook-specific registry settings.
     */
    internal sealed class FreeBusyManager : IDisposable
    {
        private readonly IfbAddressBookCache _addressBookCache;
        private readonly FreeBusyServer _server;

        private Outlook.Application _application;

        internal FreeBusyManager(string dataDirectory)
        {
            _addressBookCache = new IfbAddressBookCache(dataDirectory);
            _server = new FreeBusyServer(_addressBookCache);
        }

        internal void Initialize(Outlook.Application application)
        {
            _application = application;
        }

        internal void ApplySettings(AddinSettings settings)
        {
            if (_application == null || settings == null)
            {
                return;
            }

            bool credentialsComplete = !string.IsNullOrWhiteSpace(settings.ServerUrl)
                                       && !string.IsNullOrWhiteSpace(settings.Username)
                                       && !string.IsNullOrEmpty(settings.AppPassword);

            if (!settings.IfbEnabled || !credentialsComplete)
            {
                StopServer();
                RestoreFreeBusyPath(settings);
                return;
            }

            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            if (!configuration.IsComplete())
            {
                StopServer();
                RestoreFreeBusyPath(settings);
                return;
            }

            _server.UpdateSettings(configuration, settings.IfbDays, settings.IfbCacheHours);
            try
            {
                _server.Start();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("IFB server could not be started: " + ex.Message, ex);
            }

            EnsureFreeBusyPath(settings);
        }

        private void EnsureFreeBusyPath(AddinSettings settings)
        {
            string desired = BuildIfbUrl(settings);

            bool primaryUpdated = TryUpdateKey(BuildCalendarOptionsKey(), "FreeBusySearchPath", desired, settings, true);
            if (!primaryUpdated)
            {
                throw new InvalidOperationException("IFB path could not be set: access denied or key not available.");
            }

            TryUpdateKey(BuildPolicyCalendarOptionsKey(), "FreeBusySearchPath", desired, settings, false);

            TryUpdateKey(BuildInternetFreeBusyKey(), "Read URL", desired, settings, true);
            TryUpdateKey(BuildPolicyInternetFreeBusyKey(), "Read URL", desired, settings, false);
        }

        private void RestoreFreeBusyPath(AddinSettings settings)
        {
            string previous = settings.IfbPreviousFreeBusyPath ?? string.Empty;
            TryRestoreKey(BuildCalendarOptionsKey(), "FreeBusySearchPath", previous);
            TryRestoreKey(BuildPolicyCalendarOptionsKey(), "FreeBusySearchPath", previous);
            TryRestoreKey(BuildInternetFreeBusyKey(), "Read URL", previous);
            TryRestoreKey(BuildPolicyInternetFreeBusyKey(), "Read URL", previous);
        }

        private string BuildCalendarOptionsKey()
        {
            return @"Software\Microsoft\Office\" + GetOutlookVersionSegment() + @"\Outlook\Options\Calendar";
        }

        private string BuildInternetFreeBusyKey()
        {
            return BuildCalendarOptionsKey() + @"\Internet Free/Busy";
        }

        private string BuildPolicyCalendarOptionsKey()
        {
            return @"Software\Policies\Microsoft\Office\" + GetOutlookVersionSegment() + @"\Outlook\Options\Calendar";
        }

        private string BuildPolicyInternetFreeBusyKey()
        {
            return BuildPolicyCalendarOptionsKey() + @"\Internet Free/Busy";
        }

        private string GetOutlookVersionSegment()
        {
            string version = "16.0";
            try
            {
                string versionString = _application != null ? _application.Version : null;
                if (!string.IsNullOrEmpty(versionString))
                {
                    string[] parts = versionString.Split('.');
                    if (parts.Length >= 2)
                    {
                        version = parts[0] + "." + parts[1];
                    }
                    else if (parts.Length == 1)
                    {
                        version = parts[0] + ".0";
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to read Outlook version. Falling back to default version segment.", ex);
                version = "16.0";
            }

            return version;
        }

        private static string BuildIfbUrl(AddinSettings settings)
        {
            string domain = GuessDefaultDomain(settings);
            var builder = new StringBuilder("http://127.0.0.1:7777/nc-ifb/freebusy/%NAME%");
            if (!string.IsNullOrEmpty(domain))
            {
                builder.Append("@").Append(domain);
            }
            builder.Append(".vfb");
            return builder.ToString();
        }

        private static string GuessDefaultDomain(AddinSettings settings)
        {
            if (!string.IsNullOrWhiteSpace(settings.Username) && settings.Username.Contains("@"))
            {
                var parts = settings.Username.Split('@');
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    return parts[1].Trim();
                }
            }

            try
            {
                var uri = new Uri(settings.ServerUrl);
                string host = uri.Host ?? string.Empty;
                var segments = host.Split('.');
                if (segments.Length >= 2)
                {
                    return segments[segments.Length - 2] + "." + segments[segments.Length - 1];
                }
                return host;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to derive IFB default domain from server URL.", ex);
                return string.Empty;
            }
        }

        private bool TryUpdateKey(string path, string valueName, string desired, AddinSettings settings, bool critical)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            string current;
            bool hasCurrent = TryReadValue(path, valueName, out current);
            if (hasCurrent && string.IsNullOrEmpty(settings.IfbPreviousFreeBusyPath) && !string.IsNullOrEmpty(current))
            {
                settings.IfbPreviousFreeBusyPath = current;
            }
            if (hasCurrent && string.Equals(current, desired, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!hasCurrent && !critical)
            {
                return true;
            }

            try
            {
                using (var key = OpenOrCreateSubKey(path))
                {
                    if (key == null)
                    {
                        DiagnosticsLogger.Log(LogCategories.Ifb, "Kein Zugriff auf Registry '" + path + "' (Value '" + valueName + "').");
                        if (critical)
                        {
                            throw new InvalidOperationException("IFB path could not be set: access to '" + path + "' denied.");
                        }
                        return false;
                    }

                    string writableCurrent = key.GetValue(valueName, string.Empty) as string ?? string.Empty;
                    if (string.IsNullOrEmpty(settings.IfbPreviousFreeBusyPath) && !string.IsNullOrEmpty(writableCurrent))
                    {
                        settings.IfbPreviousFreeBusyPath = writableCurrent;
                    }

                    if (!string.Equals(writableCurrent, desired, StringComparison.OrdinalIgnoreCase))
                    {
                        key.SetValue(valueName, desired, RegistryValueKind.String);
                    }

                    return true;
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Kein Zugriff auf Registry '" + path + "': " + ex.Message);
                if (critical)
                {
                    throw new InvalidOperationException("IFB path could not be set: access denied.", ex);
                }
                return false;
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Sicherheitsausnahme bei Registry '" + path + "': " + ex.Message);
                if (critical)
                {
                    throw new InvalidOperationException("IFB path could not be set: a security policy prevented access.", ex);
                }
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Registry '" + path + "' konnte nicht aktualisiert werden: " + ex.Message);
                if (critical)
                {
                    throw new InvalidOperationException("IFB path could not be set: " + ex.Message, ex);
                }
                return false;
            }
        }

        private void TryRestoreKey(string path, string valueName, string previous)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            string existing;
            if (TryReadValue(path, valueName, out existing))
            {
                if (!string.IsNullOrEmpty(previous))
                {
                    if (string.Equals(existing, previous, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
                else if (string.IsNullOrEmpty(existing))
                {
                    return;
                }
            }

            try
            {
                using (var key = OpenOrCreateSubKey(path))
                {
                    if (key == null)
                    {
                        DiagnosticsLogger.Log(LogCategories.Ifb, "Kein Zugriff auf Registry '" + path + "' beim Wiederherstellen.");
                        return;
                    }

                    if (!string.IsNullOrEmpty(previous))
                    {
                        key.SetValue(valueName, previous, RegistryValueKind.String);
                    }
                    else
                    {
                        try
                        {
                            key.DeleteValue(valueName, false);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to delete registry value '" + valueName + "' at '" + path + "'. Falling back to setting empty string.", ex);
                            key.SetValue(valueName, string.Empty, RegistryValueKind.String);
                        }
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Kein Zugriff auf Registry '" + path + "' beim Wiederherstellen: " + ex.Message);
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.Log(LogCategories.Ifb, "Sicherheitsausnahme bei Registry '" + path + "' beim Wiederherstellen: " + ex.Message);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to restore registry value '" + valueName + "' at '" + path + "'.", ex);
            }
        }

        private bool TryReadValue(string path, string valueName, out string value)
        {
            value = string.Empty;
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(path, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object raw = key.GetValue(valueName, string.Empty);
                    value = raw as string ?? string.Empty;
                    return true;
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "No access to registry '" + path + "' while reading '" + valueName + "'.", ex);
                return false;
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Security exception while reading registry '" + path + "' value '" + valueName + "'.", ex);
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Failed to read registry '" + path + "' value '" + valueName + "'.", ex);
                return false;
            }
        }

        private static RegistryKey OpenOrCreateSubKey(string path)
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(path, true);
                if (key != null)
                {
                    return key;
                }

                return Registry.CurrentUser.CreateSubKey(path);
            }
            catch (UnauthorizedAccessException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "No access to registry '" + path + "' (open/create).", ex);
                return null;
            }
            catch (System.Security.SecurityException ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Security exception for registry '" + path + "' (open/create).", ex);
                return null;
            }
        }

        internal void StopServer()
        {
            _server.Stop();
        }

        public void Dispose()
        {
            StopServer();
        }
    }
}
