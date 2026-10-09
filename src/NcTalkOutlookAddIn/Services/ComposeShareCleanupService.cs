// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Persists accepted cleanup jobs, never ownership of an open or saved draft.
    internal sealed class ComposeShareCleanupService : IDisposable
    {
        private static readonly ProtectedJsonStateStoreDefinition<ComposeShareCleanupState> StoreDefinition =
            new ProtectedJsonStateStoreDefinition<ComposeShareCleanupState>(
                "compose-share-cleanup", "NC4OL::ComposeShareCleanup::v1", IsValid,
                () => new ComposeShareCleanupState(), LogCategories.FileLink,
                new ProtectedJsonStateStoreMessages(
                    "Recovered pending share cleanup from its backup.",
                    "Could not restore the share cleanup backup.",
                    "Share cleanup recovery failed; existing files are preserved.",
                    "Share cleanup storage is unreadable; existing files are preserved.",
                    "Share cleanup storage is invalid.",
                    "Failed to load share cleanup file '"));

        private readonly object _syncRoot = new object();
        private ProtectedJsonStateStore<ComposeShareCleanupState> _store;
        private ComposeShareCleanupState _state;
        private Timer _timer;
        private int _workerRunning;
        private bool _disposed;

        internal void Initialize(string dataDirectory, string profileScope)
        {
            lock (_syncRoot)
            {
                if (_store != null) { return; }
                _store = new ProtectedJsonStateStore<ComposeShareCleanupState>(dataDirectory, profileScope, StoreDefinition);
                _state = _store.Load();
                _timer = new Timer(OnTimer, null, Timeout.Infinite, Timeout.Infinite);
                NextcloudConnectionState.ConnectionRestored += OnConnectionRestored;
            }
            VerifiedNextcloudIdentity identity = NextcloudConnectionState.GetVerifiedIdentity();
            if (identity != null) { OnConnectionRestored(identity); }
            else if (_state.Records.Exists(record => record.ConnectionPaused && !string.IsNullOrWhiteSpace(record.AccountId))
                && !NextcloudConnectionState.GetStatus().IsPaused)
            {
                ThreadPool.QueueUserWorkItem(ignored =>
                {
                    try
                    {
                        VerifiedNextcloudIdentity restored = NextcloudConnectionState.ResolveVerifiedIdentity();
                        if (restored != null) { OnConnectionRestored(restored); }
                    }
                    catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.FileLink, "Share cleanup recovery awaits a verified connection.", ex); }
                });
            }
            ScheduleNext();
        }

        internal bool QueueCleanup(IList<ComposeShareCleanupRecord> records, string reason)
        {
            if (records == null) { return true; }
            try
            {
                lock (_syncRoot)
                {
                    if (_store == null || _disposed || !_store.IsWriteAllowed) { return false; }
                    var next = new ComposeShareCleanupState { Records = new List<ComposeShareCleanupJob>(_state.Records) };
                    foreach (ComposeShareCleanupRecord entry in records)
                    {
                        if (entry == null || entry.Origin == null || string.IsNullOrWhiteSpace(entry.RelativeFolder)
                            || string.IsNullOrWhiteSpace(entry.Origin.ServerUrl)) { continue; }
                        string server = entry.Origin.ToConfiguration().GetNormalizedBaseUrl();
                        if (string.IsNullOrWhiteSpace(server)) { continue; }
                        string userId = entry.Origin.AccountId ?? string.Empty;
                        bool exists = next.Records.Exists(record => record.ServerBaseUrl == server
                            && record.AccountId == userId && record.RelativeFolder == entry.RelativeFolder);
                        if (exists) { continue; }
                        next.Records.Add(new ComposeShareCleanupJob
                        {
                            ServerBaseUrl = server,
                            AccountId = userId,
                            RelativeFolder = entry.RelativeFolder,
                            ConnectionPaused = string.IsNullOrWhiteSpace(userId),
                            NextAttemptUtc = DateTime.UtcNow
                        });
                    }
                    // The persisted descriptor deliberately contains no captured app password.
                    _store.Save(next);
                    _state = next;
                }
                DiagnosticsLogger.Log(LogCategories.FileLink, "Share cleanup retained (reason=" + (reason ?? string.Empty) + ").");
                ScheduleNext();
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Share cleanup could not be persisted.", ex);
                return false;
            }
        }

        private void OnTimer(object state)
        {
            if (Interlocked.CompareExchange(ref _workerRunning, 1, 0) != 0) { return; }
            try { ProcessDueRecords(); }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Pending share cleanup failed.", ex);
            }
            finally
            {
                Interlocked.Exchange(ref _workerRunning, 0);
                ScheduleNext();
            }
        }

        private void ProcessDueRecords()
        {
            List<ComposeShareCleanupJob> due;
            lock (_syncRoot)
            {
                if (_disposed) { return; }
                due = _state.Records.FindAll(record => !record.ConnectionPaused && record.NextAttemptUtc <= DateTime.UtcNow);
            }
            if (due.Count == 0) { return; }
            ConnectionPauseStatus pause = NextcloudConnectionState.GetStatus();
            if (pause.IsPaused)
            {
                foreach (ComposeShareCleanupJob record in due) { Defer(record.Id, pause, false); }
                return;
            }

            VerifiedNextcloudIdentity identity;
            try { identity = NextcloudConnectionState.GetVerifiedIdentity() ?? NextcloudConnectionState.ResolveVerifiedIdentity(); }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Share cleanup retained until its account can be verified.", ex);
                pause = NextcloudConnectionState.GetStatus();
                foreach (ComposeShareCleanupJob record in due) { Defer(record.Id, pause, true); }
                return;
            }
            foreach (ComposeShareCleanupJob record in due)
            {
                if (!MatchesAccount(record, identity))
                {
                    Defer(record.Id, null, false);
                    continue;
                }
                try
                {
                    new FileLinkService(identity.Configuration).DeleteShareFolder(record.RelativeFolder, CancellationToken.None);
                    lock (_syncRoot)
                    {
                        _state.Records.RemoveAll(item => item.Id == record.Id);
                        _store.Save(_state);
                    }
                    DiagnosticsLogger.Log(LogCategories.FileLink, "Pending share cleanup completed.");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Share cleanup retained after a failed deletion.", ex);
                    Defer(record.Id, NextcloudConnectionState.GetStatus(identity.Configuration), true);
                }
            }
        }

        private void Defer(string id, ConnectionPauseStatus pause, bool failed)
        {
            lock (_syncRoot)
            {
                ComposeShareCleanupJob record = _state.Records.Find(item => item.Id == id);
                if (record == null) { return; }
                record.ConnectionPaused = pause == null || pause.Reason == "auth_required";
                if (pause != null && pause.Reason == "rate_limited")
                {
                    record.NextAttemptUtc = pause.RetryAfterUtc;
                }
                else if (!record.ConnectionPaused)
                {
                    if (failed) { record.AttemptCount++; }
                    record.NextAttemptUtc = DateTime.UtcNow.AddMinutes(record.AttemptCount <= 1 ? 1 : record.AttemptCount == 2 ? 5 : 60);
                }
                _store.Save(_state);
            }
            if (pause == null || pause.Reason == "auth_required")
            {
                VerifiedNextcloudIdentity restored = NextcloudConnectionState.GetVerifiedIdentity();
                if (restored != null) { OnConnectionRestored(restored); }
            }
        }

        private void OnConnectionRestored(VerifiedNextcloudIdentity identity)
        {
            try
            {
                lock (_syncRoot)
                {
                    if (_disposed || _store == null) { return; }
                    bool changed = false;
                    foreach (ComposeShareCleanupJob record in _state.Records)
                    {
                        if (!MatchesAccount(record, identity)) { continue; }
                        record.ConnectionPaused = false;
                        record.NextAttemptUtc = DateTime.UtcNow;
                        changed = true;
                    }
                    if (changed) { _store.Save(_state); }
                }
                ScheduleNext();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Pending share cleanup could not be resumed.", ex);
            }
        }

        private static bool MatchesAccount(ComposeShareCleanupJob record, VerifiedNextcloudIdentity identity)
        {
            return identity != null && identity.Configuration != null
                && !string.IsNullOrWhiteSpace(record.AccountId)
                && string.Equals(record.ServerBaseUrl, identity.BaseUrl, StringComparison.Ordinal)
                && string.Equals(record.AccountId, identity.UserId, StringComparison.Ordinal);
        }

        private void ScheduleNext()
        {
            lock (_syncRoot)
            {
                if (_disposed || _timer == null) { return; }
                DateTime? next = null;
                foreach (ComposeShareCleanupJob record in _state.Records)
                {
                    if (!record.ConnectionPaused && (!next.HasValue || record.NextAttemptUtc < next.Value))
                    { next = record.NextAttemptUtc; }
                }
                int delay = !next.HasValue ? Timeout.Infinite
                    : (int)Math.Min(int.MaxValue, Math.Max(0, Math.Ceiling((next.Value - DateTime.UtcNow).TotalMilliseconds)));
                _timer.Change(delay, Timeout.Infinite);
            }
        }

        private static bool IsValid(ComposeShareCleanupState state)
        {
            if (state == null || state.Records == null) { return false; }
            var ids = new HashSet<string>(StringComparer.Ordinal);
            return state.Records.TrueForAll(record => record != null && !string.IsNullOrWhiteSpace(record.Id)
                && ids.Add(record.Id) && !string.IsNullOrWhiteSpace(record.ServerBaseUrl)
                && !string.IsNullOrWhiteSpace(record.RelativeFolder));
        }

        public void Dispose()
        {
            lock (_syncRoot)
            {
                _disposed = true;
                NextcloudConnectionState.ConnectionRestored -= OnConnectionRestored;
                if (_timer != null) { _timer.Dispose(); }
            }
        }
    }

    internal sealed class ComposeShareCleanupState
    {
        public ComposeShareCleanupState() { Records = new List<ComposeShareCleanupJob>(); }
        public List<ComposeShareCleanupJob> Records { get; set; }
    }

    internal sealed class ComposeShareCleanupJob
    {
        public ComposeShareCleanupJob() { Id = Guid.NewGuid().ToString("N"); }
        public string Id { get; set; }
        public string ServerBaseUrl { get; set; }
        public string AccountId { get; set; }
        public string RelativeFolder { get; set; }
        public bool ConnectionPaused { get; set; }
        public int AttemptCount { get; set; }
        public DateTime NextAttemptUtc { get; set; }
    }
}
