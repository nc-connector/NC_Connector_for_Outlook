// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class ProtectedJsonStateStoreMessages
    {
        internal ProtectedJsonStateStoreMessages(
            string recoveredFromBackup,
            string restoreBackupFailed,
            string recoveryFailed,
            string writeBlocked,
            string invalidStructure,
            string loadFailedPrefix)
        {
            RecoveredFromBackup = recoveredFromBackup;
            RestoreBackupFailed = restoreBackupFailed;
            RecoveryFailed = recoveryFailed;
            WriteBlocked = writeBlocked;
            InvalidStructure = invalidStructure;
            LoadFailedPrefix = loadFailedPrefix;
        }

        internal string RecoveredFromBackup { get; private set; }

        internal string RestoreBackupFailed { get; private set; }

        internal string RecoveryFailed { get; private set; }

        internal string WriteBlocked { get; private set; }

        internal string InvalidStructure { get; private set; }

        internal string LoadFailedPrefix { get; private set; }
    }

    internal sealed class ProtectedJsonStateStoreDefinition<TState>
        where TState : class
    {
        internal ProtectedJsonStateStoreDefinition(
            string fileNamePrefix,
            string protectionEntropy,
            Func<TState, bool> isValid,
            Func<TState> createEmptyState,
            string logCategory,
            ProtectedJsonStateStoreMessages messages)
        {
            FileNamePrefix = fileNamePrefix;
            ProtectionEntropy = protectionEntropy;
            IsValid = isValid;
            CreateEmptyState = createEmptyState;
            LogCategory = logCategory;
            Messages = messages;
        }

        internal string FileNamePrefix { get; private set; }

        internal string ProtectionEntropy { get; private set; }

        internal Func<TState, bool> IsValid { get; private set; }

        internal Func<TState> CreateEmptyState { get; private set; }

        internal string LogCategory { get; private set; }

        internal ProtectedJsonStateStoreMessages Messages
        {
            get;
            private set;
        }
    }

    internal sealed class ProtectedJsonStateStore<TState>
        where TState : class
    {
        private readonly object _syncRoot = new object();
        private readonly string _filePath;
        private readonly string _backupFilePath;
        private readonly byte[] _protectionEntropy;
        private readonly ProtectedJsonStateStoreDefinition<TState> _definition;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private bool _writeAllowed = true;

        internal bool IsWriteAllowed
        {
            get { lock (_syncRoot) { return _writeAllowed; } }
        }

        internal ProtectedJsonStateStore(
            string dataDirectory,
            string profileScope,
            ProtectedJsonStateStoreDefinition<TState> definition)
        {
            if (definition == null)
            {
                throw new ArgumentNullException("definition");
            }

            _definition = definition;
            _protectionEntropy = Encoding.UTF8.GetBytes(
                definition.ProtectionEntropy);
            string directory = string.IsNullOrWhiteSpace(dataDirectory)
                ? AppDataPaths.EnsureLocalRootDirectory()
                : dataDirectory;
            Directory.CreateDirectory(directory);
            _filePath = Path.Combine(
                directory,
                definition.FileNamePrefix
                + "-"
                + BuildScopeHash(profileScope)
                + ".dat");
            _backupFilePath = _filePath + ".bak";
        }

        internal TState Load()
        {
            lock (_syncRoot)
            {
                TState state;
                if (TryLoadFile(_filePath, out state))
                {
                    return state;
                }
                if (TryLoadFile(_backupFilePath, out state))
                {
                    DiagnosticsLogger.Log(
                        _definition.LogCategory,
                        _definition.Messages.RecoveredFromBackup);
                    try
                    {
                        File.Copy(
                            _backupFilePath,
                            _filePath,
                            true);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            _definition.LogCategory,
                            _definition.Messages.RestoreBackupFailed,
                            ex);
                    }
                    return state;
                }

                if (File.Exists(_filePath)
                    || File.Exists(_backupFilePath))
                {
                    _writeAllowed = false;
                    DiagnosticsLogger.Log(
                        _definition.LogCategory,
                        _definition.Messages.RecoveryFailed);
                }
                return _definition.CreateEmptyState();
            }
        }

        internal void Save(TState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException("state");
            }

            lock (_syncRoot)
            {
                if (!_writeAllowed)
                {
                    throw new InvalidOperationException(
                        _definition.Messages.WriteBlocked);
                }

                byte[] clearBytes = Encoding.UTF8.GetBytes(
                    _serializer.Serialize(state));
                byte[] protectedBytes = ProtectedData.Protect(
                    clearBytes,
                    _protectionEntropy,
                    DataProtectionScope.CurrentUser);
                string temporaryPath =
                    _filePath
                    + "."
                    + Guid.NewGuid().ToString("N")
                    + ".tmp";
                try
                {
                    File.WriteAllText(
                        temporaryPath,
                        Convert.ToBase64String(protectedBytes),
                        new UTF8Encoding(false));
                    if (DurableFileReplace.CommitPreparedFile(
                        temporaryPath,
                        _filePath,
                        _backupFilePath))
                    {
                        temporaryPath = null;
                    }
                }
                finally
                {
                    if (!string.IsNullOrEmpty(temporaryPath)
                        && File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }
            }
        }

        private bool TryLoadFile(
            string path,
            out TState state)
        {
            state = null;
            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
                byte[] protectedBytes = Convert.FromBase64String(
                    File.ReadAllText(path, Encoding.UTF8));
                byte[] clearBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    _protectionEntropy,
                    DataProtectionScope.CurrentUser);
                TState candidate = _serializer.Deserialize<TState>(
                    Encoding.UTF8.GetString(clearBytes));
                if (!_definition.IsValid(candidate))
                {
                    throw new InvalidDataException(
                        _definition.Messages.InvalidStructure);
                }
                state = candidate;
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    _definition.LogCategory,
                    _definition.Messages.LoadFailedPrefix
                    + Path.GetFileName(path)
                    + "'.",
                    ex);
                return false;
            }
        }

        private static string BuildScopeHash(string profileScope)
        {
            byte[] input = Encoding.UTF8.GetBytes(
                string.IsNullOrWhiteSpace(profileScope)
                    ? "default"
                    : profileScope.Trim().ToUpperInvariant());
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create())
            {
                hash = sha256.ComputeHash(input);
            }

            var builder = new StringBuilder(24);
            for (int i = 0; i < 12; i++)
            {
                builder.Append(
                    hash[i].ToString(
                        "x2",
                        CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }
    }
}
