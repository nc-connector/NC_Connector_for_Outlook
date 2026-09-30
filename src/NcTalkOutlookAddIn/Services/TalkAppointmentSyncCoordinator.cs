// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class TalkAppointmentSyncCoordinator : IDisposable
    {
        private readonly object _syncRoot = new object();
        private readonly Dictionary<string, SyncSlot> _slots =
            new Dictionary<string, SyncSlot>(StringComparer.OrdinalIgnoreCase);
        private readonly Func<TalkAppointmentSyncSnapshot, TalkAppointmentSyncResult> _execute;
        private readonly Func<TalkAppointmentSyncResult, Task> _completed;

        private bool _disposed;

        internal TalkAppointmentSyncCoordinator(
            Func<TalkAppointmentSyncSnapshot, TalkAppointmentSyncResult> execute,
            Func<TalkAppointmentSyncResult, Task> completed)
        {
            if (execute == null)
            {
                throw new ArgumentNullException("execute");
            }
            if (completed == null)
            {
                throw new ArgumentNullException("completed");
            }

            _execute = execute;
            _completed = completed;
        }

        internal void Queue(TalkAppointmentSyncSnapshot snapshot)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.RoomToken))
            {
                return;
            }

            string key = snapshot.RoomToken.Trim();
            bool startWorker = false;
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }

                SyncSlot slot;
                if (!_slots.TryGetValue(key, out slot))
                {
                    slot = new SyncSlot();
                    _slots[key] = slot;
                }
                // A deferred lobby check must not discard a pending full edit.
                slot.Pending = snapshot.LobbyOnly
                               && slot.Pending != null
                               && !slot.Pending.LobbyOnly
                    ? slot.Pending.WithLobbyUpdate(snapshot)
                    : snapshot;
                if (!slot.Running)
                {
                    slot.Running = true;
                    startWorker = true;
                }
            }

            if (startWorker)
            {
                // Snapshots contain no Outlook COM objects, so remote updates can run
                // outside the Outlook STA.
                Task.Run(() => ProcessAsync(key));
            }
        }

        private async Task ProcessAsync(string key)
        {
            while (true)
            {
                TalkAppointmentSyncSnapshot snapshot;
                lock (_syncRoot)
                {
                    SyncSlot slot;
                    if (_disposed || !_slots.TryGetValue(key, out slot))
                    {
                        return;
                    }

                    snapshot = slot.Pending;
                    slot.Pending = null;
                    if (snapshot == null)
                    {
                        _slots.Remove(key);
                        return;
                    }
                }

                TalkAppointmentSyncResult result;
                try
                {
                    result = _execute(snapshot);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Appointment background synchronization failed.",
                        ex);
                    result = new TalkAppointmentSyncResult(snapshot.RoomToken);
                }

                try
                {
                    lock (_syncRoot)
                    {
                        if (_disposed)
                        {
                            return;
                        }
                    }
                    Task completion = _completed(result);
                    if (completion != null)
                    {
                        await completion.ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Appointment synchronization completion dispatch failed.",
                        ex);
                }
            }
        }

        public void Dispose()
        {
            lock (_syncRoot)
            {
                _disposed = true;
                _slots.Clear();
            }
        }

        private sealed class SyncSlot
        {
            internal bool Running { get; set; }

            internal TalkAppointmentSyncSnapshot Pending { get; set; }
        }
    }
}
