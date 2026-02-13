/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.IO;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    /// <summary>
    /// Simple persistence for add-in settings. The data is currently stored
    /// unencrypted in an INI-like structure.
    /// TODO: Store the app password securely (Credential Locker or DPAPI).
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
                catch (Exception ex)
                {
                    // If the copy fails, we will read directly from the legacy path later.
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to copy legacy settings file to new location.", ex);
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
                    case "SharingDefaultShareName":
                        settings.SharingDefaultShareName = value;
                        break;
                    case "SharingDefaultPermCreate":
                        bool permCreate;
                        if (bool.TryParse(value, out permCreate))
                        {
                            settings.SharingDefaultPermCreate = permCreate;
                        }
                        break;
                    case "SharingDefaultPermWrite":
                        bool permWrite;
                        if (bool.TryParse(value, out permWrite))
                        {
                            settings.SharingDefaultPermWrite = permWrite;
                        }
                        break;
                    case "SharingDefaultPermDelete":
                        bool permDelete;
                        if (bool.TryParse(value, out permDelete))
                        {
                            settings.SharingDefaultPermDelete = permDelete;
                        }
                        break;
                    case "SharingDefaultPasswordEnabled":
                        bool sharingPassword;
                        if (bool.TryParse(value, out sharingPassword))
                        {
                            settings.SharingDefaultPasswordEnabled = sharingPassword;
                        }
                        break;
                    case "SharingDefaultExpireDays":
                        int expireDays;
                        if (int.TryParse(value, out expireDays))
                        {
                            settings.SharingDefaultExpireDays = expireDays;
                        }
                        break;
                    case "ShareBlockLang":
                        settings.ShareBlockLang = string.IsNullOrWhiteSpace(value) ? "default" : value.Trim();
                        break;
                    case "EventDescriptionLang":
                        settings.EventDescriptionLang = string.IsNullOrWhiteSpace(value) ? "default" : value.Trim();
                        break;
                    case "TalkDefaultLobbyEnabled":
                        bool talkLobby;
                        if (bool.TryParse(value, out talkLobby))
                        {
                            settings.TalkDefaultLobbyEnabled = talkLobby;
                        }
                        break;
                    case "TalkDefaultSearchVisible":
                        bool talkSearch;
                        if (bool.TryParse(value, out talkSearch))
                        {
                            settings.TalkDefaultSearchVisible = talkSearch;
                        }
                        break;
                    case "TalkDefaultRoomType":
                        TalkRoomType roomType;
                        if (Enum.TryParse<TalkRoomType>(value, out roomType))
                        {
                            settings.TalkDefaultRoomType = roomType;
                        }
                        break;
                    case "TalkDefaultPasswordEnabled":
                        bool talkPasswordEnabled;
                        if (bool.TryParse(value, out talkPasswordEnabled))
                        {
                            settings.TalkDefaultPasswordEnabled = talkPasswordEnabled;
                        }
                        break;
                    case "TalkDefaultAddUsers":
                        bool talkAddUsers;
                        if (bool.TryParse(value, out talkAddUsers))
                        {
                            settings.TalkDefaultAddUsers = talkAddUsers;
                        }
                        break;
                    case "TalkDefaultAddGuests":
                        bool talkAddGuests;
                        if (bool.TryParse(value, out talkAddGuests))
                        {
                            settings.TalkDefaultAddGuests = talkAddGuests;
                        }
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
            entries.Add("# NC Connector for Outlook settings");
            entries.Add("# TODO: Store the app password securely (Credential Locker / DPAPI).");
            entries.Add("ServerUrl=" + Safe(settings.ServerUrl));
            entries.Add("Username=" + Safe(settings.Username));
            entries.Add("AppPassword=" + Safe(settings.AppPassword));
            entries.Add("AuthMode=" + settings.AuthMode);
            entries.Add("IfbEnabled=" + settings.IfbEnabled);
            entries.Add("IfbDays=" + settings.IfbDays);
            entries.Add("IfbCacheHours=" + settings.IfbCacheHours);
            entries.Add("IfbPreviousFreeBusyPath=" + Safe(settings.IfbPreviousFreeBusyPath));
            entries.Add("DebugLoggingEnabled=" + settings.DebugLoggingEnabled);
            entries.Add("LastKnownServerVersion=" + Safe(settings.LastKnownServerVersion));
            entries.Add("FileLinkBasePath=" + Safe(settings.FileLinkBasePath));
            entries.Add("SharingDefaultShareName=" + Safe(settings.SharingDefaultShareName));
            entries.Add("SharingDefaultPermCreate=" + settings.SharingDefaultPermCreate);
            entries.Add("SharingDefaultPermWrite=" + settings.SharingDefaultPermWrite);
            entries.Add("SharingDefaultPermDelete=" + settings.SharingDefaultPermDelete);
            entries.Add("SharingDefaultPasswordEnabled=" + settings.SharingDefaultPasswordEnabled);
            entries.Add("SharingDefaultExpireDays=" + settings.SharingDefaultExpireDays);
            entries.Add("ShareBlockLang=" + Safe(settings.ShareBlockLang));
            entries.Add("EventDescriptionLang=" + Safe(settings.EventDescriptionLang));
            entries.Add("TalkDefaultLobbyEnabled=" + settings.TalkDefaultLobbyEnabled);
            entries.Add("TalkDefaultSearchVisible=" + settings.TalkDefaultSearchVisible);
            entries.Add("TalkDefaultRoomType=" + settings.TalkDefaultRoomType);
            entries.Add("TalkDefaultPasswordEnabled=" + settings.TalkDefaultPasswordEnabled);
            entries.Add("TalkDefaultAddUsers=" + settings.TalkDefaultAddUsers);
            entries.Add("TalkDefaultAddGuests=" + settings.TalkDefaultAddGuests);

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
    }
}


