// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Persists requested Talk room deletions and retries failed requests.
    internal sealed class TalkRoomLifecycleCoordinator : IDisposable
    {
        private readonly object _syncRoot = new object();
        private readonly TalkRoomLifecycleStore _store;
        private readonly TalkRoomLifecycleState _state;
        private readonly Func<AddinSettings> _getCurrentSettings;
        private readonly Func<bool> _isSavedEventDeletionEnabled;
        private readonly Func<TalkServiceConfiguration, string> _resolveAccountId;
        private readonly Action<TalkServiceConfiguration, string, bool> _deleteRoom;
        private readonly Timer _timer;
        private int _workerRunning;
        private bool _disposed;

        internal TalkRoomLifecycleCoordinator(
            string dataDirectory,
            string profileScope,
            Func<AddinSettings> getCurrentSettings,
            Func<bool> isSavedEventDeletionEnabled)
            : this(
                dataDirectory,
                profileScope,
                getCurrentSettings,
                isSavedEventDeletionEnabled,
                configuration =>
                {
                    VerifiedNextcloudIdentity identity = NextcloudConnectionState.ResolveVerifiedIdentity();
                    return identity != null && string.Equals(identity.BaseUrl, configuration.GetNormalizedBaseUrl(), StringComparison.Ordinal)
                        ? identity.UserId : string.Empty;
                },
                (configuration, roomToken, isEventConversation) =>
                    new TalkService(configuration).DeleteRoom(roomToken, isEventConversation))
        {
        }

        internal TalkRoomLifecycleCoordinator(
            string dataDirectory,
            string profileScope,
            Func<AddinSettings> getCurrentSettings,
            Func<bool> isSavedEventDeletionEnabled,
            Func<TalkServiceConfiguration, string> resolveAccountId,
            Action<TalkServiceConfiguration, string, bool> deleteRoom)
        {
            if (getCurrentSettings == null)
            {
                throw new ArgumentNullException("getCurrentSettings");
            }
            if (isSavedEventDeletionEnabled == null)
            {
                throw new ArgumentNullException("isSavedEventDeletionEnabled");
            }
            if (resolveAccountId == null)
            {
                throw new ArgumentNullException("resolveAccountId");
            }
            if (deleteRoom == null)
            {
                throw new ArgumentNullException("deleteRoom");
            }

            _store = new TalkRoomLifecycleStore(dataDirectory, profileScope);
            _state = _store.Load() ?? new TalkRoomLifecycleState();
            if (_state.Records == null)
            {
                _state.Records = new List<TalkRoomLifecycleRecord>();
            }
            DiscardTrackingRecords();
            _getCurrentSettings = getCurrentSettings;
            _isSavedEventDeletionEnabled = isSavedEventDeletionEnabled;
            _resolveAccountId = resolveAccountId;
            _deleteRoom = deleteRoom;
            _timer = new Timer(OnTimer, null, Timeout.Infinite, Timeout.Infinite);
            NextcloudConnectionState.ConnectionRestored += OnConnectionRestored;
        }

        internal void StartPendingProcessing()
        {
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
                    catch (Exception ex) { DiagnosticsLogger.LogException(LogCategories.Talk, "Talk cleanup recovery awaits a verified connection.", ex); }
                });
            }
            ScheduleNext();
        }

        internal bool QueueDeletion(
            string roomToken,
            bool isEventConversation,
            TalkServiceConfiguration configuration,
            bool policyRequired)
        {
            if (string.IsNullOrWhiteSpace(roomToken)
                || configuration == null
                || !configuration.IsComplete())
            {
                return false;
            }

            try
            {
                string normalizedToken = roomToken.Trim();
                string baseUrl = NormalizeBaseUrl(configuration);
                string login = NormalizeLogin(configuration);
                VerifiedNextcloudIdentity identity = NextcloudConnectionState.GetKnownIdentity(configuration);
                lock (_syncRoot)
                {
                    TalkRoomLifecycleRecord record =
                        FindPendingDeletionLocked(normalizedToken, baseUrl, login);
                    if (record == null)
                    {
                        record = new TalkRoomLifecycleRecord
                        {
                            RoomToken = normalizedToken,
                            IsEventConversation = isEventConversation,
                            ServerBaseUrl = baseUrl,
                            AccountLogin = login,
                            AccountId = identity != null ? identity.UserId : string.Empty,
                            PendingDeletion = true,
                            PolicyRequired = policyRequired
                        };
                        _state.Records.Add(record);
                    }
                    else
                    {
                        record.IsEventConversation = isEventConversation;
                        record.PendingDeletion = true;
                        record.PolicyRequired = record.PolicyRequired && policyRequired;
                    }

                    record.AttemptCount = 0;
                    record.NextAttemptUtc = DateTime.UtcNow;
                    // Persist before scheduling so restart retries observed deletions
                    // without scanning Outlook calendars.
                    SaveLocked();
                }

                DiagnosticsLogger.Log(
                    LogCategories.Talk,
                    "Talk room deletion queued (policyRequired=" + policyRequired + ").");
                Schedule(TimeSpan.Zero);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room deletion could not be queued.",
                    ex);
                return false;
            }
        }

        internal static bool MatchesAccount(
            TalkRoomLifecycleRecord record,
            string baseUrl,
            string login,
            string accountId)
        {
            if (record == null
                || !string.Equals(record.ServerBaseUrl, baseUrl, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            return !string.IsNullOrWhiteSpace(record.AccountId)
                ? string.Equals(record.AccountId, accountId, StringComparison.Ordinal)
                : string.Equals(record.AccountLogin, login, StringComparison.OrdinalIgnoreCase);
        }

        private void DiscardTrackingRecords()
        {
            int removed = _state.Records.RemoveAll(
                record => record == null || !record.PendingDeletion);
            if (removed <= 0)
            {
                return;
            }

            try
            {
                _store.Save(_state);
                DiagnosticsLogger.Log(
                    LogCategories.Talk,
                    "Removed " + removed + " obsolete Talk room tracking records.");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Obsolete Talk room tracking records could not be removed from the persisted queue.",
                    ex);
            }
        }

        private TalkRoomLifecycleRecord FindPendingDeletionLocked(
            string roomToken,
            string baseUrl,
            string login)
        {
            for (int i = 0; i < _state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = _state.Records[i];
                if (record != null
                    && record.PendingDeletion
                    && string.Equals(record.RoomToken, roomToken, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(record.ServerBaseUrl, baseUrl, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(record.AccountLogin, login, StringComparison.OrdinalIgnoreCase))
                {
                    return record;
                }
            }
            return null;
        }

        private void OnTimer(object state)
        {
            if (Interlocked.CompareExchange(ref _workerRunning, 1, 0) != 0)
            {
                return;
            }
            try
            {
                ProcessDueRecords();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room deletion worker failed.",
                    ex);
            }
            finally
            {
                Interlocked.Exchange(ref _workerRunning, 0);
                ScheduleNext();
            }
        }

        private void ProcessDueRecords()
        {
            List<TalkRoomLifecycleRecord> due = GetDueRecords(DateTime.UtcNow);
            if (due.Count == 0)
            {
                return;
            }

            TalkServiceConfiguration configuration = BuildConfiguration(_getCurrentSettings());
            ConnectionPauseStatus pause = NextcloudConnectionState.GetStatus(configuration);
            if (pause.IsPaused)
            {
                PauseAll(due, pause);
                return;
            }
            if (configuration == null)
            {
                PauseAll(due, new ConnectionPauseStatus("auth_required", DateTime.MinValue));
                return;
            }

            string accountId;
            try
            {
                accountId = _resolveAccountId(configuration);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Talk room deletions retained because the account could not be verified.",
                    ex);
                pause = NextcloudConnectionState.GetStatus(configuration);
                if (pause.IsPaused) { PauseAll(due, pause); }
                else { RescheduleAll(due, TimeSpan.FromHours(1)); }
                return;
            }
            if (string.IsNullOrWhiteSpace(accountId))
            {
                pause = NextcloudConnectionState.GetStatus(configuration);
                if (pause.IsPaused) { PauseAll(due, pause); }
                else { RescheduleAll(due, TimeSpan.FromHours(1)); }
                return;
            }

            string baseUrl = NormalizeBaseUrl(configuration);
            string login = NormalizeLogin(configuration);
            bool policyChecked = false;
            bool policyAvailable = true;
            bool policyEnabled = false;
            for (int i = 0; i < due.Count; i++)
            {
                TalkRoomLifecycleRecord record = due[i];
                pause = NextcloudConnectionState.GetStatus(configuration);
                if (pause.IsPaused)
                {
                    Pause(record.Id, pause);
                    continue;
                }
                if (!MatchesAccount(record, baseUrl, login, accountId))
                {
                    Reschedule(record.Id, TimeSpan.FromHours(1));
                    continue;
                }
                BindAccountId(record.Id, accountId);

                // Saved appointment deletion stays opt-in until execution because the setting
                // may change after Outlook recorded the deletion event.
                if (record.PolicyRequired)
                {
                    if (!policyChecked)
                    {
                        policyChecked = true;
                        try
                        {
                            policyEnabled = _isSavedEventDeletionEnabled();
                        }
                        catch (Exception ex)
                        {
                            policyAvailable = false;
                            DiagnosticsLogger.LogException(
                                LogCategories.Talk,
                                "Saved-event room deletion policy could not be evaluated.",
                                ex);
                        }
                    }
                    // An authentication rejection during the policy fetch must not look like
                    // a disabled cleanup policy and discard an already accepted deletion.
                    pause = NextcloudConnectionState.GetStatus(configuration);
                    if (pause.IsPaused)
                    {
                        Pause(record.Id, pause);
                        continue;
                    }
                    if (!policyAvailable)
                    {
                        Reschedule(record.Id, TimeSpan.FromHours(1));
                        continue;
                    }
                    if (!policyEnabled)
                    {
                        Remove(record.Id);
                        DiagnosticsLogger.Log(
                            LogCategories.Talk,
                            "Saved-event room deletion skipped because the setting is disabled.");
                        continue;
                    }
                }

                try
                {
                    _deleteRoom(
                        configuration,
                        record.RoomToken,
                        record.IsEventConversation);
                    Remove(record.Id);
                    DiagnosticsLogger.Log(
                        LogCategories.Talk,
                        "Queued Talk room deletion completed.");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Queued Talk room deletion failed.",
                        ex);
                    pause = NextcloudConnectionState.GetStatus(configuration);
                    if (pause.IsPaused) { Pause(record.Id, pause); }
                    else { Reschedule(record.Id, GetRetryDelay(record.AttemptCount + 1)); }
                }
            }
        }

        private List<TalkRoomLifecycleRecord> GetDueRecords(DateTime nowUtc)
        {
            var result = new List<TalkRoomLifecycleRecord>();
            lock (_syncRoot)
            {
                for (int i = 0; i < _state.Records.Count; i++)
                {
                    TalkRoomLifecycleRecord record = _state.Records[i];
                    if (record != null
                        && record.PendingDeletion
                        && !record.ConnectionPaused
                        && (record.NextAttemptUtc <= DateTime.MinValue
                            || record.NextAttemptUtc <= nowUtc))
                    {
                        result.Add(CloneForDeletion(record));
                    }
                }
            }
            return result;
        }

        private void RescheduleAll(
            IList<TalkRoomLifecycleRecord> records,
            TimeSpan delay)
        {
            for (int i = 0; i < records.Count; i++)
            {
                Reschedule(records[i].Id, delay);
            }
        }

        private void PauseAll(IList<TalkRoomLifecycleRecord> records, ConnectionPauseStatus pause)
        {
            foreach (TalkRoomLifecycleRecord record in records) { Pause(record.Id, pause); }
        }

        private void Pause(string id, ConnectionPauseStatus pause)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(id);
                if (record == null) { return; }
                record.ConnectionPaused = pause.Reason == "auth_required";
                if (pause.Reason == "rate_limited") { record.NextAttemptUtc = pause.RetryAfterUtc; }
                SaveLocked();
            }
            // A verified save may have completed concurrently with the failed request.
            VerifiedNextcloudIdentity identity = NextcloudConnectionState.GetVerifiedIdentity();
            if (identity != null) { OnConnectionRestored(identity); }
        }

        private void OnConnectionRestored(VerifiedNextcloudIdentity identity)
        {
            if (identity == null) { return; }
            try
            {
                lock (_syncRoot)
                {
                    if (_disposed) { return; }
                    bool changed = false;
                    foreach (TalkRoomLifecycleRecord record in _state.Records)
                    {
                        if (!record.PendingDeletion || string.IsNullOrWhiteSpace(record.AccountId)
                            || !string.Equals(record.ServerBaseUrl, identity.BaseUrl, StringComparison.Ordinal)
                            || !string.Equals(record.AccountId, identity.UserId, StringComparison.Ordinal)) { continue; }
                        record.ConnectionPaused = false;
                        record.NextAttemptUtc = DateTime.UtcNow;
                        changed = true;
                    }
                    if (changed) { SaveLocked(); }
                }
                ScheduleNext();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Pending Talk cleanup could not be resumed.", ex);
            }
        }

        private void BindAccountId(string recordId, string accountId)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                if (record == null
                    || !record.PendingDeletion
                    || string.Equals(record.AccountId, accountId, StringComparison.Ordinal))
                {
                    return;
                }
                record.AccountId = accountId;
                SaveLocked();
            }
        }

        private void Reschedule(string recordId, TimeSpan delay)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                if (record == null || !record.PendingDeletion)
                {
                    return;
                }
                record.AttemptCount++;
                record.NextAttemptUtc = DateTime.UtcNow.Add(delay);
                SaveLocked();
            }
        }

        private void Remove(string recordId)
        {
            lock (_syncRoot)
            {
                TalkRoomLifecycleRecord record = FindByIdLocked(recordId);
                if (record == null)
                {
                    return;
                }
                _state.Records.Remove(record);
                SaveLocked();
            }
        }

        private TalkRoomLifecycleRecord FindByIdLocked(string recordId)
        {
            for (int i = 0; i < _state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record = _state.Records[i];
                if (record != null
                    && string.Equals(record.Id, recordId, StringComparison.Ordinal))
                {
                    return record;
                }
            }
            return null;
        }

        private void ScheduleNext()
        {
            DateTime? earliest = null;
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }
                for (int i = 0; i < _state.Records.Count; i++)
                {
                    TalkRoomLifecycleRecord record = _state.Records[i];
                    if (record == null || !record.PendingDeletion || record.ConnectionPaused)
                    {
                        continue;
                    }
                    DateTime due = record.NextAttemptUtc <= DateTime.MinValue
                        ? DateTime.UtcNow
                        : record.NextAttemptUtc;
                    if (!earliest.HasValue || due < earliest.Value)
                    {
                        earliest = due;
                    }
                }
            }
            if (earliest.HasValue)
            {
                TimeSpan delay = earliest.Value - DateTime.UtcNow;
                Schedule(delay > TimeSpan.Zero ? delay : TimeSpan.Zero);
            }
        }

        private void Schedule(TimeSpan delay)
        {
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }
                int milliseconds = delay <= TimeSpan.Zero
                    ? 0
                    : (int)Math.Min(int.MaxValue, Math.Ceiling(delay.TotalMilliseconds));
                _timer.Change(milliseconds, Timeout.Infinite);
            }
        }

        private static TalkServiceConfiguration BuildConfiguration(AddinSettings settings)
        {
            if (settings == null)
            {
                return null;
            }
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            return configuration.IsComplete() ? configuration : null;
        }

        private static string NormalizeBaseUrl(TalkServiceConfiguration configuration)
        {
            return configuration == null
                ? string.Empty
                : configuration.GetNormalizedBaseUrl().TrimEnd('/');
        }

        private static string NormalizeLogin(TalkServiceConfiguration configuration)
        {
            return configuration == null
                ? string.Empty
                : (configuration.Username ?? string.Empty).Trim();
        }

        private static TimeSpan GetRetryDelay(int attempt)
        {
            if (attempt <= 1)
            {
                return TimeSpan.FromMinutes(1);
            }
            if (attempt == 2)
            {
                return TimeSpan.FromMinutes(5);
            }
            if (attempt == 3)
            {
                return TimeSpan.FromMinutes(15);
            }
            return TimeSpan.FromHours(1);
        }

        private static TalkRoomLifecycleRecord CloneForDeletion(
            TalkRoomLifecycleRecord source)
        {
            return new TalkRoomLifecycleRecord
            {
                Id = source.Id,
                RoomToken = source.RoomToken,
                IsEventConversation = source.IsEventConversation,
                ServerBaseUrl = source.ServerBaseUrl,
                AccountLogin = source.AccountLogin,
                AccountId = source.AccountId,
                PendingDeletion = source.PendingDeletion,
                PolicyRequired = source.PolicyRequired,
                ConnectionPaused = source.ConnectionPaused,
                AttemptCount = source.AttemptCount,
                NextAttemptUtc = source.NextAttemptUtc
            };
        }

        private void SaveLocked()
        {
            _store.Save(_state);
        }

        public void Dispose()
        {
            lock (_syncRoot)
            {
                _disposed = true;
                NextcloudConnectionState.ConnectionRestored -= OnConnectionRestored;
                _timer.Change(Timeout.Infinite, Timeout.Infinite);
            }
            _timer.Dispose();
        }
    }
}
