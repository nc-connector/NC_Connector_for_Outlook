/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    /// <summary>
    /// Einfache Persistenz fuer Add-in Einstellungen. Die Daten werden aktuell
    /// unverschluesselt in einer INI-aehnlichen Struktur gespeichert.
    /// TODO: App-Passwort sicher ablegen (Credential Locker oder DPAPI).
    /// </summary>
    internal sealed class SettingsStorage
    {
        private const string FileName = "settings.ini";
        private const string LegacyFolderName = "NextcloudTalkOutlookAddIn";
        private const string DataFolderName = "NextcloudTalkOutlookAddInData";

        private readonly string _filePath;
        private readonly string _legacyFilePath;

        internal string DataDirectory
        {
            get { return Path.GetDirectoryName(_filePath); }
        }

        internal SettingsStorage()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var dataPath = Path.Combine(localAppData, DataFolderName);
            Directory.CreateDirectory(dataPath);

            _filePath = Path.Combine(dataPath, FileName);
            _legacyFilePath = Path.Combine(localAppData, LegacyFolderName, FileName);

            if (!File.Exists(_filePath) && File.Exists(_legacyFilePath))
            {
                try
                {
                    File.Copy(_legacyFilePath, _filePath, true);
                }
                catch
                {
                    // Falls Kopie misslingt, lesen wir spaeter direkt aus dem Legacy-Pfad.
                }
            }
        }

        internal AddinSettings Load()
        {
            var settings = new AddinSettings();

            var pathToRead = _filePath;
            if (!File.Exists(pathToRead) && File.Exists(_legacyFilePath))
            {
                pathToRead = _legacyFilePath;
            }

            if (!File.Exists(pathToRead))
            {
                return settings;
            }

            foreach (var line in File.ReadAllLines(pathToRead))
            {
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                var parts = line.Split(new[] { '=' }, 2);
                if (parts.Length != 2)
                {
                    continue;
                }

                var key = parts[0].Trim();
                var value = parts[1].Trim();

                switch (key)
                {
                    case "ServerUrl":
                        settings.ServerUrl = value;
                        break;
                    case "Username":
                        settings.Username = value;
                        break;
                    case "AppPassword":
                        settings.AppPassword = value;
                        break;
                    case "AuthMode":
                        AuthenticationMode mode;
                        if (Enum.TryParse<AuthenticationMode>(value, out mode))
                        {
                            settings.AuthMode = mode;
                        }
                        break;
                    case "OutlookMuzzleEnabled":
                        bool muzzle;
                        if (bool.TryParse(value, out muzzle))
                        {
                            settings.OutlookMuzzleEnabled = muzzle;
                        }
                        break;
                    case "IfbEnabled":
                        bool ifbEnabled;
                        if (bool.TryParse(value, out ifbEnabled))
                        {
                            settings.IfbEnabled = ifbEnabled;
                        }
                        break;
                    case "IfbDays":
                        int ifbDays;
                        if (int.TryParse(value, out ifbDays))
                        {
                            settings.IfbDays = ifbDays;
                        }
                        break;
                    case "IfbCacheHours":
                        int cacheHours;
                        if (int.TryParse(value, out cacheHours))
                        {
                            settings.IfbCacheHours = cacheHours;
                        }
                        break;
                    case "IfbPreviousFreeBusyPath":
                        settings.IfbPreviousFreeBusyPath = value;
                        break;
                    case "DebugLoggingEnabled":
                        bool debug;
                        if (bool.TryParse(value, out debug))
                        {
                            settings.DebugLoggingEnabled = debug;
                        }
                        break;
                    case "LastKnownServerVersion":
                        settings.LastKnownServerVersion = value;
                        break;
                    case "FileLinkBasePath":
                        settings.FileLinkBasePath = value;
                        break;
                    case "OutlookMuzzleAccounts":
                        settings.OutlookMuzzleAccounts = ParseMuzzleAccounts(value);
                        break;
                    default:
                        break;
                }
            }

            return settings;
        }

        internal void Save(AddinSettings settings)
        {
            var entries = new List<string>();
            entries.Add("# Nextcloud Talk Direkt Outlook Add-in Einstellungen");
            entries.Add("# TODO: App-Passwort sicher speichern (Credential Locker / DPAPI).");
            entries.Add("ServerUrl=" + Safe(settings.ServerUrl));
            entries.Add("Username=" + Safe(settings.Username));
            entries.Add("AppPassword=" + Safe(settings.AppPassword));
            entries.Add("AuthMode=" + settings.AuthMode);
            entries.Add("OutlookMuzzleEnabled=" + settings.OutlookMuzzleEnabled);
            entries.Add("OutlookMuzzleAccounts=" + SerializeMuzzleAccounts(settings.OutlookMuzzleAccounts));
            entries.Add("IfbEnabled=" + settings.IfbEnabled);
            entries.Add("IfbDays=" + settings.IfbDays);
            entries.Add("IfbCacheHours=" + settings.IfbCacheHours);
            entries.Add("IfbPreviousFreeBusyPath=" + Safe(settings.IfbPreviousFreeBusyPath));
            entries.Add("DebugLoggingEnabled=" + settings.DebugLoggingEnabled);
            entries.Add("LastKnownServerVersion=" + Safe(settings.LastKnownServerVersion));
            entries.Add("FileLinkBasePath=" + Safe(settings.FileLinkBasePath));

            var dir = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllLines(_filePath, entries);
        }

        private static string Safe(string value)
        {
            return value ?? string.Empty;
        }

        private static Dictionary<string, bool> ParseMuzzleAccounts(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            }

            var result = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var entries = value.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var entry in entries)
            {
                var pair = entry.Split(new[] { ':' }, 2);
                var key = pair[0].Trim();
                bool enabled = true;
                if (pair.Length == 2)
                {
                    bool.TryParse(pair[1], out enabled);
                }

                if (!string.IsNullOrWhiteSpace(key))
                {
                    result[OutlookAccountHelper.NormalizeAccountIdentifier(key)] = enabled;
                }
            }

            return result;
        }

        private static string SerializeMuzzleAccounts(Dictionary<string, bool> accounts)
        {
            if (accounts == null || accounts.Count == 0)
            {
                return string.Empty;
            }

            var parts = accounts
                .Where(kv => !string.IsNullOrWhiteSpace(kv.Key))
                .Select(kv => kv.Key + ":" + kv.Value);
            return string.Join(";", parts);
        }
    }
}


