// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace NcTalkOutlookAddIn.Settings
{
    /// <summary>
    /// Serializes profile settings access across Outlook processes and commits files atomically.
    /// </summary>
    internal sealed class SettingsFileTransaction
    {
        private const int LockTimeoutMilliseconds = 15000;
        private readonly string _primaryPath;
        private readonly string _backupPath;
        private readonly string _mutexName;

        internal SettingsFileTransaction(string primaryPath)
        {
            if (string.IsNullOrWhiteSpace(primaryPath))
            {
                throw new ArgumentException("A settings file path is required.", "primaryPath");
            }

            _primaryPath = Path.GetFullPath(primaryPath);
            _backupPath = _primaryPath + ".bak";
            _mutexName = BuildMutexName(_primaryPath);
        }

        internal string PrimaryPath
        {
            get { return _primaryPath; }
        }

        internal string BackupPath
        {
            get { return _backupPath; }
        }

        internal IDisposable AcquireLock()
        {
            var mutex = new Mutex(false, _mutexName);
            bool acquired = false;
            try
            {
                try
                {
                    acquired = mutex.WaitOne(LockTimeoutMilliseconds);
                }
                catch (AbandonedMutexException)
                {
                    acquired = true;
                }

                if (!acquired)
                {
                    throw new TimeoutException("Timed out while waiting for exclusive settings file access.");
                }

                return new MutexLease(mutex);
            }
            catch
            {
                if (acquired)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Dispose();
                throw;
            }
        }

        internal void Commit(
            Action<Stream> writeContent,
            Func<string, bool> isHealthySettingsFile,
            bool replaceBackupWithNewContent = false)
        {
            if (writeContent == null)
            {
                throw new ArgumentNullException("writeContent");
            }
            if (isHealthySettingsFile == null)
            {
                throw new ArgumentNullException("isHealthySettingsFile");
            }

            string directory = Path.GetDirectoryName(_primaryPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("The settings file directory is unavailable.");
            }

            Directory.CreateDirectory(directory);
            string pendingPath = BuildPendingPath(_primaryPath);
            try
            {
                WritePendingFile(pendingPath, writeContent);
                if (!isHealthySettingsFile(pendingPath))
                {
                    throw new InvalidDataException("The newly serialized settings file did not pass validation.");
                }

                bool primaryHealthy = File.Exists(_primaryPath) && isHealthySettingsFile(_primaryPath);
                bool backupHealthy = File.Exists(_backupPath) && isHealthySettingsFile(_backupPath);
                if (replaceBackupWithNewContent)
                {
                    // Explicit credential removal must not leave a recoverable copy in the backup.
                    ReplaceWithCopy(pendingPath, _backupPath);
                }
                else if (primaryHealthy)
                {
                    ReplaceWithCopy(_primaryPath, _backupPath);
                }
                else if (!backupHealthy)
                {
                    ReplaceWithCopy(pendingPath, _backupPath);
                }

                ReplacePendingFile(pendingPath, _primaryPath);
            }
            finally
            {
                TryDeletePendingFile(pendingPath);
            }
        }

        internal bool TryRestorePrimaryFromBackup(Func<string, bool> isHealthySettingsFile)
        {
            if (isHealthySettingsFile == null
                || !File.Exists(_backupPath)
                || !isHealthySettingsFile(_backupPath))
            {
                return false;
            }

            ReplaceWithCopy(_backupPath, _primaryPath);
            return isHealthySettingsFile(_primaryPath);
        }

        private static void WritePendingFile(string pendingPath, Action<Stream> writeContent)
        {
            using (var stream = new FileStream(
                pendingPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                4096,
                FileOptions.WriteThrough))
            {
                writeContent(stream);
                stream.Flush(true);
            }
        }

        private static void ReplaceWithCopy(string sourcePath, string destinationPath)
        {
            string pendingPath = BuildPendingPath(destinationPath);
            try
            {
                using (var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var destination = new FileStream(
                    pendingPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    4096,
                    FileOptions.WriteThrough))
                {
                    source.CopyTo(destination);
                    destination.Flush(true);
                }

                ReplacePendingFile(pendingPath, destinationPath);
            }
            finally
            {
                TryDeletePendingFile(pendingPath);
            }
        }

        private static void ReplacePendingFile(string pendingPath, string destinationPath)
        {
            if (File.Exists(destinationPath))
            {
                File.Replace(pendingPath, destinationPath, null, true);
                return;
            }

            File.Move(pendingPath, destinationPath);
        }

        private static string BuildPendingPath(string targetPath)
        {
            return targetPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        }

        private static string BuildMutexName(string path)
        {
            string canonicalPath = Path.GetFullPath(path).ToUpperInvariant();
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(canonicalPath));
                return @"Local\NC4OL.Settings." + BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }

        private static void TryDeletePendingFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // The primary operation already determines success; stale temp cleanup is best effort.
            }
        }

        private sealed class MutexLease : IDisposable
        {
            private Mutex _mutex;

            internal MutexLease(Mutex mutex)
            {
                _mutex = mutex;
            }

            public void Dispose()
            {
                Mutex mutex = Interlocked.Exchange(ref _mutex, null);
                if (mutex == null)
                {
                    return;
                }

                try
                {
                    mutex.ReleaseMutex();
                }
                finally
                {
                    mutex.Dispose();
                }
            }
        }
    }
}
