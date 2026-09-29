// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Settings
{
    /// <summary>
    /// Profile-aware settings storage.
    /// Settings are persisted in XML under %LOCALAPPDATA%\NC4OL.
    /// Password values are protected with DPAPI (CurrentUser scope).
    /// </summary>
    internal sealed class SettingsStorage
    {
        private const string LegacyDataFolderName = "NextcloudTalkOutlookAddInData";
        private const string LegacyFolderName = "NextcloudTalkOutlookAddIn";
        private const string LegacyIniFileName = "settings.ini";
        private const string RuntimeLogFileName = "addin-runtime.log";
        private const string IfbCacheFileName = "ifb-addressbook-cache.json";
        private const string ProfileFilePrefix = "settings_";
        private const string ProfileFileSuffix = ".xml";
        private const string DefaultProfileName = "default";
        private const int SchemaVersion = 1;

        private static readonly byte[] PasswordEntropy = Encoding.UTF8.GetBytes("NC4OL::Settings::AppPassword::v1");

        private readonly string _profileName;
        private readonly string _dataDirectory;
        private readonly string _filePath;
        private readonly string _legacyDataDirectory;
        private readonly string _legacyFolderDirectory;
        private readonly string _legacyDataIniPath;
        private readonly string _legacyFolderIniPath;
        private readonly SettingsFileTransaction _settingsFileTransaction;
        private bool _automaticSaveBlocked;

        internal string DataDirectory
        {
            get { return _dataDirectory; }
        }

        internal SettingsStorage(string outlookProfileName)
        {
            _profileName = NormalizeProfileName(outlookProfileName);
            _dataDirectory = AppDataPaths.EnsureLocalRootDirectory();
            _filePath = Path.Combine(_dataDirectory, BuildProfileFileName(_profileName));
            _settingsFileTransaction = new SettingsFileTransaction(_filePath);

            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _legacyDataDirectory = Path.Combine(localAppData, LegacyDataFolderName);
            _legacyFolderDirectory = Path.Combine(localAppData, LegacyFolderName);
            _legacyDataIniPath = Path.Combine(_legacyDataDirectory, LegacyIniFileName);
            _legacyFolderIniPath = Path.Combine(_legacyFolderDirectory, LegacyIniFileName);

            using (_settingsFileTransaction.AcquireLock())
            {
                MigrateLegacyRuntimeArtifacts();
                EnsureLegacyIniMigrated();
            }
        }

        internal AddinSettings Load()
        {
            try
            {
                using (_settingsFileTransaction.AcquireLock())
                {
                    return LoadUnderLock();
                }
            }
            catch (Exception ex)
            {
                _automaticSaveBlocked = true;
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to acquire or read profile settings.", ex);
                return ApplyManagedSetupPolicy(new AddinSettings());
            }
        }

        internal void Save(AddinSettings settings)
        {
            SaveCore(settings, false);
        }

        internal void SaveUserInitiated(AddinSettings settings)
        {
            SaveCore(settings, true);
        }

        private AddinSettings LoadUnderLock()
        {
            _automaticSaveBlocked = false;
            bool primaryExists = File.Exists(_filePath);
            bool backupExists = File.Exists(_settingsFileTransaction.BackupPath);

            if (primaryExists)
            {
                try
                {
                    bool passwordReadFailed;
                    AddinSettings settings = LoadFromXmlFile(_filePath, out passwordReadFailed);
                    if (passwordReadFailed)
                    {
                        _automaticSaveBlocked = true;
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "Profile settings loaded without the app password because DPAPI decryption failed; automatic saves are blocked.");
                    }
                    return ApplyManagedSetupPolicy(settings);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load profile settings XML.", ex);
                }
            }

            if (backupExists)
            {
                try
                {
                    bool passwordReadFailed;
                    AddinSettings backupSettings = LoadFromXmlFile(
                        _settingsFileTransaction.BackupPath,
                        out passwordReadFailed);
                    _automaticSaveBlocked = passwordReadFailed;
                    if (passwordReadFailed)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "Profile settings backup loaded without the app password because DPAPI decryption failed; automatic saves are blocked.");
                    }
                    else
                    {
                        TryRestorePrimarySettingsFromBackup();
                    }

                    DiagnosticsLogger.Log(LogCategories.Core, "Recovered profile settings from the last valid backup.");
                    return ApplyManagedSetupPolicy(backupSettings);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load profile settings backup.", ex);
                }
            }

            string legacyPath = ResolveLegacyIniPath();
            if (!string.IsNullOrEmpty(legacyPath))
            {
                try
                {
                    _automaticSaveBlocked = primaryExists || backupExists;
                    return ApplyManagedSetupPolicy(LoadFromIniFile(legacyPath));
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load legacy settings INI.", ex);
                }
            }

            _automaticSaveBlocked = primaryExists || backupExists;
            if (_automaticSaveBlocked)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "No valid settings file could be recovered; automatic saves are blocked until the user saves settings explicitly.");
            }
            return ApplyManagedSetupPolicy(new AddinSettings());
        }

        private void SaveCore(AddinSettings settings, bool userInitiated)
        {
            if (settings == null)
            {
                settings = new AddinSettings();
            }
            if (_automaticSaveBlocked && !userInitiated)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Automatic settings save skipped because the last settings load was not fully recoverable.");
                return;
            }

            try
            {
                AddinSettings persistedSettings = ApplyManagedSetupPolicy(settings.Clone());
                using (_settingsFileTransaction.AcquireLock())
                {
                    _settingsFileTransaction.Commit(
                        stream => SaveToXmlStream(stream, persistedSettings, _profileName),
                        IsSettingsFileHealthy);
                    _automaticSaveBlocked = false;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save profile settings XML.", ex);
                throw;
            }
        }

        private void TryRestorePrimarySettingsFromBackup()
        {
            try
            {
                if (!_settingsFileTransaction.TryRestorePrimaryFromBackup(IsSettingsFileHealthy))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Profile settings backup was readable but could not be restored to the primary file.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to restore profile settings backup to the primary file.",
                    ex);
            }
        }

        private void EnsureLegacyIniMigrated()
        {
            string sourcePath = ResolveLegacyIniPath();
            if (string.IsNullOrEmpty(sourcePath))
            {
                TryDeleteDirectoryIfEmpty(_legacyDataDirectory);
                TryDeleteDirectoryIfEmpty(_legacyFolderDirectory);
                return;
            }
            bool migrated = TryMigrateLegacyIniToProfiles(sourcePath);
            if (migrated)
            {
                CleanupLegacyIniArtifacts();
                return;
            }

            DiagnosticsLogger.Log(LogCategories.Core, "Legacy settings migration was not fully successful. Legacy INI file is kept.");
        }

        private bool TryMigrateLegacyIniToProfiles(string sourcePath)
        {
            AddinSettings legacySettings;
            try
            {
                legacySettings = LoadFromIniFile(sourcePath);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to parse legacy INI for profile migration.", ex);
                return false;
            }
            var targetProfiles = ResolveMigrationProfileNames();
            bool allSucceeded = true;

            foreach (string profile in targetProfiles)
            {
                string targetFilePath = Path.Combine(_dataDirectory, BuildProfileFileName(profile));
                if (File.Exists(targetFilePath))
                {
                    continue;
                }
                try
                {
                    SaveToXmlFile(targetFilePath, legacySettings, profile);
                    DiagnosticsLogger.Log(LogCategories.Core, "Migrated legacy settings to profile XML (profile=" + profile + ").");
                }
                catch (Exception ex)
                {
                    allSucceeded = false;
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to migrate legacy settings to profile XML (profile=" + profile + ").", ex);
                }
            }
            if (!allSucceeded)
            {
                return false;
            }
            return true;
        }

        private void MigrateLegacyRuntimeArtifacts()
        {
            TryMoveLegacyDataFile(RuntimeLogFileName);
            TryMoveLegacyDataFile(IfbCacheFileName);
            TryDeleteDirectoryIfEmpty(_legacyDataDirectory);
            TryDeleteDirectoryIfEmpty(_legacyFolderDirectory);
        }

        private void TryMoveLegacyDataFile(string fileName)
        {
            string sourcePath = Path.Combine(_legacyDataDirectory, fileName);
            if (!File.Exists(sourcePath))
            {
                return;
            }
            string targetPath = Path.Combine(_dataDirectory, fileName);
            try
            {
                if (!File.Exists(targetPath))
                {
                    File.Move(sourcePath, targetPath);
                    DiagnosticsLogger.Log(LogCategories.Core, "Migrated legacy runtime file to NC4OL path (" + fileName + ").");
                    return;
                }

                File.Delete(sourcePath);
                DiagnosticsLogger.Log(LogCategories.Core, "Removed duplicate legacy runtime file (" + fileName + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to migrate legacy runtime file '" + fileName + "'.", ex);
            }
        }

        private void CleanupLegacyIniArtifacts()
        {
            TryDeleteLegacyFile(_legacyDataIniPath);
            TryDeleteLegacyFile(_legacyFolderIniPath);
            TryDeleteDirectoryIfEmpty(_legacyDataDirectory);
            TryDeleteDirectoryIfEmpty(_legacyFolderDirectory);
            DiagnosticsLogger.Log(LogCategories.Core, "Legacy settings INI cleanup completed after successful migration.");
        }

        private void TryDeleteLegacyFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }
            try
            {
                File.Delete(filePath);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to delete legacy file '" + filePath + "'.", ex);
            }
        }

        private void TryDeleteDirectoryIfEmpty(string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            {
                return;
            }
            try
            {
                if (Directory.GetFileSystemEntries(directoryPath).Length == 0)
                {
                    Directory.Delete(directoryPath, false);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to delete legacy directory '" + directoryPath + "'.", ex);
            }
        }

        private List<string> ResolveMigrationProfileNames()
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            names.Add(_profileName);

            foreach (string profile in EnumerateOutlookProfileNames())
            {
                string normalized = NormalizeProfileName(profile);
                names.Add(normalized);
            }
            if (names.Count == 0)
            {
                names.Add(DefaultProfileName);
            }
            var ordered = new List<string>(names);
            ordered.Sort(StringComparer.OrdinalIgnoreCase);
            return ordered;
        }

        private static IEnumerable<string> EnumerateOutlookProfileNames()
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            paths.Add(@"Software\Microsoft\Office\Outlook\Profiles");
            paths.Add(@"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles");

            try
            {
                using (RegistryKey officeRoot = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office", false))
                {
                    if (officeRoot != null)
                    {
                        foreach (string version in officeRoot.GetSubKeyNames())
                        {
                            if (!string.IsNullOrWhiteSpace(version))
                            {
                                paths.Add(@"Software\Microsoft\Office\" + version + @"\Outlook\Profiles");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to enumerate Office registry versions for Outlook profile discovery.", ex);
            }
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in paths)
            {
                RegistryKey profileRoot = null;
                try
                {
                    profileRoot = Registry.CurrentUser.OpenSubKey(path, false);
                    if (profileRoot == null)
                    {
                        continue;
                    }
                    foreach (string profileName in profileRoot.GetSubKeyNames())
                    {
                        if (string.IsNullOrWhiteSpace(profileName))
                        {
                            continue;
                        }

                        names.Add(profileName.Trim());
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook profile list from registry path '" + path + "'.", ex);
                }
                finally
                {
                    if (profileRoot != null)
                    {
                        profileRoot.Dispose();
                    }
                }
            }
            return names;
        }

        private string ResolveLegacyIniPath()
        {
            if (File.Exists(_legacyDataIniPath))
            {
                return _legacyDataIniPath;
            }
            if (File.Exists(_legacyFolderIniPath))
            {
                return _legacyFolderIniPath;
            }
            return null;
        }

        private static string NormalizeProfileName(string profileName)
        {
            if (string.IsNullOrWhiteSpace(profileName))
            {
                return DefaultProfileName;
            }
            string trimmed = profileName.Trim();
            return trimmed.Length == 0 ? DefaultProfileName : trimmed;
        }

        private static string BuildProfileFileName(string profileName)
        {
            string safeProfile = SanitizeFileNameSegment(NormalizeProfileName(profileName));
            return ProfileFilePrefix + safeProfile + ProfileFileSuffix;
        }

        private static string SanitizeFileNameSegment(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DefaultProfileName;
            }
            var builder = new StringBuilder(value.Trim());
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                builder.Replace(invalid, '_');
            }
            string sanitized = builder.ToString().Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? DefaultProfileName : sanitized;
        }

        private static AddinSettings LoadFromIniFile(string path)
        {
            var settings = new AddinSettings();

            foreach (string line in File.ReadAllLines(path))
            {
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }
                string[] parts = line.Split(new[] { '=' }, 2);
                if (parts.Length != 2)
                {
                    continue;
                }
                string key = parts[0].Trim();
                string value = parts[1].Trim();
                ApplySettingValue(settings, key, value);
            }
            return settings;
        }

        private static AddinSettings LoadFromXmlFile(string path)
        {
            bool passwordReadFailed;
            return LoadFromXmlFile(path, out passwordReadFailed);
        }

        private static AddinSettings LoadFromXmlFile(string path, out bool passwordReadFailed)
        {
            passwordReadFailed = false;
            var settings = new AddinSettings();
            var document = new XmlDocument();
            document.Load(path);

            XmlElement root = document.DocumentElement;
            if (root == null || !string.Equals(root.Name, "Settings", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Profile settings XML root element is missing.");
            }
            foreach (XmlNode child in root.ChildNodes)
            {
                XmlElement element = child as XmlElement;
                if (element == null)
                {
                    continue;
                }
                string key = element.Name;
                string value = element.InnerText ?? string.Empty;

                if (string.Equals(key, "AppPasswordProtected", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        settings.AppPassword = UnprotectPassword(value);
                    }
                    catch (FormatException ex)
                    {
                        passwordReadFailed = true;
                        settings.AppPassword = string.Empty;
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Stored app password is not valid DPAPI data.",
                            ex);
                    }
                    catch (CryptographicException ex)
                    {
                        passwordReadFailed = true;
                        settings.AppPassword = string.Empty;
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Stored app password could not be decrypted for the current Windows user.",
                            ex);
                    }
                    continue;
                }
                if (string.Equals(key, "AppPassword", StringComparison.OrdinalIgnoreCase))
                {
                    settings.AppPassword = value;
                    continue;
                }

                ApplySettingValue(settings, key, value);
            }
            return settings;
        }

        private static AddinSettings ApplyManagedSetupPolicy(AddinSettings settings)
        {
            if (settings == null)
            {
                settings = new AddinSettings();
            }

            settings.ApplyManagedSetupPolicy(ManagedSetupPolicy.Load());
            return settings;
        }

        private static void SaveToXmlFile(string path, AddinSettings settings, string profileName)
        {
            var transaction = new SettingsFileTransaction(path);
            using (transaction.AcquireLock())
            {
                transaction.Commit(
                    stream => SaveToXmlStream(stream, settings, profileName),
                    IsSettingsFileHealthy);
            }
        }

        private static void SaveToXmlStream(Stream stream, AddinSettings settings, string profileName)
        {
            var document = new XmlDocument();
            XmlDeclaration declaration = document.CreateXmlDeclaration("1.0", "utf-8", null);
            document.AppendChild(declaration);

            XmlElement root = document.CreateElement("Settings");
            root.SetAttribute("SchemaVersion", SchemaVersion.ToString(CultureInfo.InvariantCulture));
            root.SetAttribute("Profile", NormalizeProfileName(profileName));
            document.AppendChild(root);

            AppendElement(document, root, "ServerUrl", Safe(settings.ServerUrl));
            AppendElement(document, root, "Username", Safe(settings.Username));
            AppendElement(document, root, "AppPasswordProtected", ProtectPassword(settings.AppPassword));
            AppendElement(document, root, "AuthMode", settings.AuthMode.ToString());
            AppendElement(document, root, "IfbEnabled", settings.IfbEnabled.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "IfbDays", settings.IfbDays.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "IfbCacheHours", settings.IfbCacheHours.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "IfbPort", AddinSettings.NormalizeIfbPort(settings.IfbPort).ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "IfbPreviousFreeBusyPath", Safe(settings.IfbPreviousFreeBusyPath));
            AppendElement(document, root, "IfbUserDecisionRecorded", settings.IfbUserDecisionRecorded.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "DebugLoggingEnabled", settings.LocalDebugLoggingEnabled.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "LogAnonymizationEnabled", settings.LocalLogAnonymizationEnabled.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "TransportTlsUseSystemDefault", settings.LocalTransportTlsUseSystemDefault.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "TransportTlsEnable12", settings.LocalTransportTlsEnable12.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "TransportTlsEnable13", settings.LocalTransportTlsEnable13.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "UpdateNotifyEnabled", settings.LocalUpdateNotifyEnabled.ToString(CultureInfo.InvariantCulture));
            AppendElement(document, root, "UpdateInstallId", Safe(settings.UpdateInstallId));
            AppendElement(document, root, "UpdateLastCheckedAtUtc", Safe(settings.UpdateLastCheckedAtUtc));
            AppendElement(document, root, "UpdateLatestVersion", Safe(settings.UpdateLatestVersion));
            AppendElement(document, root, "UpdateReleaseUrl", Safe(settings.UpdateReleaseUrl));
            AppendElement(document, root, "UpdateDownloadUrl", Safe(settings.UpdateDownloadUrl));
            AppendElement(document, root, "UpdatePublishedAt", Safe(settings.UpdatePublishedAt));
            AppendElement(document, root, "UpdateChangelogTitle", Safe(settings.UpdateChangelogTitle));
            AppendElement(document, root, "UpdateChangelogText", Safe(settings.UpdateChangelogText));
            AppendElement(document, root, "UpdateLastNotifiedVersion", Safe(settings.UpdateLastNotifiedVersion));
            AppendElement(document, root, "UpdateLastNotifiedDateUtc", Safe(settings.UpdateLastNotifiedDateUtc));
            AppendLocalChoiceElement(document, root, settings, "FileLinkBasePath", Safe(settings.FileLinkBasePath));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultShareName", Safe(settings.SharingDefaultShareName));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPermCreate", settings.SharingDefaultPermCreate.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPermWrite", settings.SharingDefaultPermWrite.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPermDelete", settings.SharingDefaultPermDelete.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPasswordEnabled", settings.SharingDefaultPasswordEnabled.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPasswordSeparateEnabled", settings.SharingDefaultPasswordSeparateEnabled.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultPasswordDeliveryMode", SharePasswordDeliveryPolicy.ToStorageValue(settings.SharingDefaultPasswordDeliveryMode));
            AppendLocalChoiceElement(document, root, settings, "SharingDefaultExpireDays", settings.SharingDefaultExpireDays.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingAttachmentsAlwaysConnector", settings.SharingAttachmentsAlwaysConnector.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingAttachmentsOfferAboveEnabled", settings.SharingAttachmentsOfferAboveEnabled.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "SharingAttachmentsOfferAboveMb", Math.Max(1, settings.SharingAttachmentsOfferAboveMb).ToString(CultureInfo.InvariantCulture));
            if (settings.SharingAttachmentLinkTarget.HasValue)
            {
                AppendElement(
                    document,
                    root,
                    "SharingAttachmentLinkTarget",
                    AttachmentLinkTargetPolicy.ToStorageValue(settings.SharingAttachmentLinkTarget.Value));
            }
            AppendLocalChoiceElement(document, root, settings, "ShareBlockLang", Safe(settings.ShareBlockLang));
            AppendLocalChoiceElement(document, root, settings, "EventDescriptionLang", Safe(settings.EventDescriptionLang));
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultLobbyEnabled", settings.TalkDefaultLobbyEnabled.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultSearchVisible", settings.TalkDefaultSearchVisible.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultRoomType", settings.TalkDefaultRoomType.ToString());
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultPasswordEnabled", settings.TalkDefaultPasswordEnabled.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultAddUsers", settings.TalkDefaultAddUsers.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "TalkDefaultAddGuests", settings.TalkDefaultAddGuests.ToString(CultureInfo.InvariantCulture));
            AppendLocalChoiceElement(document, root, settings, "TalkDeleteRoomOnEventDelete", settings.TalkDeleteRoomOnEventDelete.ToString(CultureInfo.InvariantCulture));
            AppendOptionalBoolElement(document, root, "EmailSignatureOnCompose", settings.EmailSignatureOnCompose);
            AppendOptionalBoolElement(document, root, "EmailSignatureOnReply", settings.EmailSignatureOnReply);
            AppendOptionalBoolElement(document, root, "EmailSignatureOnForward", settings.EmailSignatureOnForward);

            var writerSettings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = new UTF8Encoding(false)
            };

            using (var writer = XmlWriter.Create(stream, writerSettings))
            {
                document.Save(writer);
            }
        }

        private static bool IsSettingsFileHealthy(string path)
        {
            try
            {
                bool passwordReadFailed;
                LoadFromXmlFile(path, out passwordReadFailed);
                return !passwordReadFailed;
            }
            catch
            {
                return false;
            }
        }

        private static void ApplySettingValue(AddinSettings settings, string key, string value)
        {
            if (settings == null || string.IsNullOrWhiteSpace(key))
            {
                return;
            }

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
                case "IfbPort":
                    int ifbPort;
                    if (int.TryParse(value, out ifbPort))
                    {
                        settings.IfbPort = AddinSettings.NormalizeIfbPort(ifbPort);
                    }
                    break;
                case "IfbPreviousFreeBusyPath":
                    settings.IfbPreviousFreeBusyPath = value;
                    break;
                case "IfbUserDecisionRecorded":
                    bool ifbUserDecisionRecorded;
                    if (bool.TryParse(value, out ifbUserDecisionRecorded))
                    {
                        settings.IfbUserDecisionRecorded = ifbUserDecisionRecorded;
                    }
                    break;
                case "DebugLoggingEnabled":
                    bool debug;
                    if (bool.TryParse(value, out debug))
                    {
                        settings.DebugLoggingEnabled = debug;
                    }
                    break;
                case "LogAnonymizationEnabled":
                    bool logAnonymizationEnabled;
                    if (bool.TryParse(value, out logAnonymizationEnabled))
                    {
                        settings.LogAnonymizationEnabled = logAnonymizationEnabled;
                    }
                    break;
                case "TransportTlsUseSystemDefault":
                    bool transportTlsUseSystemDefault;
                    if (bool.TryParse(value, out transportTlsUseSystemDefault))
                    {
                        settings.TransportTlsUseSystemDefault = transportTlsUseSystemDefault;
                    }
                    break;
                case "TransportTlsEnable12":
                    bool transportTlsEnable12;
                    if (bool.TryParse(value, out transportTlsEnable12))
                    {
                        settings.TransportTlsEnable12 = transportTlsEnable12;
                    }
                    break;
                case "TransportTlsEnable13":
                    bool transportTlsEnable13;
                    if (bool.TryParse(value, out transportTlsEnable13))
                    {
                        settings.TransportTlsEnable13 = transportTlsEnable13;
                    }
                    break;
                case "UpdateNotifyEnabled":
                    bool updateNotifyEnabled;
                    if (bool.TryParse(value, out updateNotifyEnabled))
                    {
                        settings.UpdateNotifyEnabled = updateNotifyEnabled;
                    }
                    break;
                case "UpdateInstallId":
                    settings.UpdateInstallId = value;
                    break;
                case "UpdateLastCheckedAtUtc":
                    settings.UpdateLastCheckedAtUtc = value;
                    break;
                case "UpdateLatestVersion":
                    settings.UpdateLatestVersion = value;
                    break;
                case "UpdateReleaseUrl":
                    settings.UpdateReleaseUrl = value;
                    break;
                case "UpdateDownloadUrl":
                    settings.UpdateDownloadUrl = value;
                    break;
                case "UpdatePublishedAt":
                    settings.UpdatePublishedAt = value;
                    break;
                case "UpdateChangelogTitle":
                    settings.UpdateChangelogTitle = value;
                    break;
                case "UpdateChangelogText":
                    settings.UpdateChangelogText = value;
                    break;
                case "UpdateLastNotifiedVersion":
                    settings.UpdateLastNotifiedVersion = value;
                    break;
                case "UpdateLastNotifiedDateUtc":
                    settings.UpdateLastNotifiedDateUtc = value;
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
                case "SharingDefaultPasswordSeparateEnabled":
                    bool sharingPasswordSeparate;
                    if (bool.TryParse(value, out sharingPasswordSeparate))
                    {
                        settings.SharingDefaultPasswordSeparateEnabled = sharingPasswordSeparate;
                    }
                    break;
                case "SharingDefaultPasswordDeliveryMode":
                    settings.SharingDefaultPasswordDeliveryMode = SharePasswordDeliveryPolicy.ParseMode(value);
                    break;
                case "SharingDefaultExpireDays":
                    int expireDays;
                    if (int.TryParse(value, out expireDays))
                    {
                        settings.SharingDefaultExpireDays = expireDays;
                    }
                    break;
                case "SharingAttachmentsAlwaysConnector":
                    bool attachmentsAlwaysConnector;
                    if (bool.TryParse(value, out attachmentsAlwaysConnector))
                    {
                        settings.SharingAttachmentsAlwaysConnector = attachmentsAlwaysConnector;
                    }
                    break;
                case "SharingAttachmentsOfferAboveEnabled":
                    bool attachmentsOfferAboveEnabled;
                    if (bool.TryParse(value, out attachmentsOfferAboveEnabled))
                    {
                        settings.SharingAttachmentsOfferAboveEnabled = attachmentsOfferAboveEnabled;
                    }
                    break;
                case "SharingAttachmentsOfferAboveMb":
                    int attachmentsOfferAboveMb;
                    if (int.TryParse(value, out attachmentsOfferAboveMb))
                    {
                        settings.SharingAttachmentsOfferAboveMb = Math.Max(1, attachmentsOfferAboveMb);
                    }
                    break;
                case "SharingAttachmentLinkTarget":
                    AttachmentLinkTarget attachmentLinkTarget;
                    settings.SharingAttachmentLinkTarget = AttachmentLinkTargetPolicy.TryParse(value, out attachmentLinkTarget)
                        ? attachmentLinkTarget
                        : (AttachmentLinkTarget?)null;
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
                case "TalkDeleteRoomOnEventDelete":
                    bool talkDeleteRoomOnEventDelete;
                    if (bool.TryParse(value, out talkDeleteRoomOnEventDelete))
                    {
                        settings.TalkDeleteRoomOnEventDelete = talkDeleteRoomOnEventDelete;
                    }
                    break;
                case "EmailSignatureOnCompose":
                    bool emailSignatureOnCompose;
                    if (bool.TryParse(value, out emailSignatureOnCompose))
                    {
                        settings.EmailSignatureOnCompose = emailSignatureOnCompose;
                    }
                    break;
                case "EmailSignatureOnReply":
                    bool emailSignatureOnReply;
                    if (bool.TryParse(value, out emailSignatureOnReply))
                    {
                        settings.EmailSignatureOnReply = emailSignatureOnReply;
                    }
                    break;
                case "EmailSignatureOnForward":
                    bool emailSignatureOnForward;
                    if (bool.TryParse(value, out emailSignatureOnForward))
                    {
                        settings.EmailSignatureOnForward = emailSignatureOnForward;
                    }
                    break;
                default:
                    break;
            }
        }

        private static void AppendLocalChoiceElement(XmlDocument document, XmlElement root, AddinSettings settings, string name, string value)
        {
            if (settings.HasLocalValue(name))
            {
                AppendElement(document, root, name, value);
            }
        }

        private static void AppendOptionalBoolElement(XmlDocument document, XmlElement root, string name, bool? value)
        {
            if (!value.HasValue)
            {
                return;
            }
            AppendElement(document, root, name, value.Value.ToString(CultureInfo.InvariantCulture));
        }

        private static void AppendElement(XmlDocument document, XmlElement root, string name, string value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value ?? string.Empty;
            root.AppendChild(element);
        }

        private static string ProtectPassword(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                return string.Empty;
            }

            byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);
            byte[] protectedBytes = ProtectedData.Protect(plainBytes, PasswordEntropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        private static string UnprotectPassword(string protectedValue)
        {
            if (string.IsNullOrWhiteSpace(protectedValue))
            {
                return string.Empty;
            }

            byte[] protectedBytes = Convert.FromBase64String(protectedValue);
            byte[] plainBytes = ProtectedData.Unprotect(protectedBytes, PasswordEntropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plainBytes);
        }

        private static string Safe(string value)
        {
            return value ?? string.Empty;
        }
    }
}
