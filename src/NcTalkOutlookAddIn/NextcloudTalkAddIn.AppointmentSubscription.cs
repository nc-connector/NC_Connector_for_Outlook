/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        private sealed class AppointmentSubscription : IDisposable
        {
            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.AppointmentItem _appointment;
            private readonly string _key;
            private readonly string _roomToken;
            private readonly bool _lobbyEnabled;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly bool _isEventConversation;
            private long? _lastLobbyTimer;
            private bool _roomDeleted;
            private bool _disposed;
            private bool _unsavedCloseCleanupPending;
            private int _unsavedCloseCleanupAttempts;
            private System.Windows.Forms.Timer _unsavedCloseCleanupTimer;
            private System.Windows.Forms.Timer _deferredWriteLobbyTimer;
            private int _deferredWriteLobbyAttempts;
            private string _entryId;
            private const int UnsavedCloseCleanupMaxAttempts = 90;
            private const int DeferredWriteLobbyMaxAttempts = 4;

            internal AppointmentSubscription(
                NextcloudTalkAddIn owner,
                Outlook.AppointmentItem appointment,
                string key,
                string roomToken,
                bool lobbyEnabled,
                bool isEventConversation,
                string entryId)
            {
                _owner = owner;
                _appointment = appointment;
                _key = key;
                _roomToken = roomToken;
                _lobbyEnabled = lobbyEnabled;
                _isEventConversation = isEventConversation;
                _lastLobbyTimer = GetIcalStartEpochOrNull(appointment);
                _entryId = entryId;
                _events = appointment as Outlook.ItemEvents_10_Event;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_events != null)
                {
                    _events.BeforeDelete += OnBeforeDelete;
                    _events.Write += OnWrite;
                    _events.Close += OnClose;
                }

                LogTalk("Subscription registered (token=" + _roomToken + ", lobby=" + _lobbyEnabled + ", event=" + _isEventConversation + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }

            private void OnWrite(ref bool cancel)
            {
                if (_unsavedCloseCleanupPending)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("OnWrite canceled pending unsaved cleanup (token=" + _roomToken + ").");
                }

                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("OnWrite ignored (no token).");
                    return;
                }

                long currentStartEpoch = 0;
                bool hasStampedStartEpoch = _owner.TryStampIcalStartEpoch(_appointment, _roomToken, out currentStartEpoch);

                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnWrite ignored (not organizer, token=" + _roomToken + ", stampedStart=" + hasStampedStartEpoch + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("OnWrite skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                LogTalk("OnWrite for appointment (token=" + _roomToken + ").");

                bool effectiveLobbyKnown;
                bool effectiveLobbyEnabled;
                bool effectiveIsEventConversation;
                _owner._talkAppointmentController.ResolveRuntimeRoomTraits(_appointment, _roomToken, _lobbyEnabled, _isEventConversation, out effectiveLobbyKnown, out effectiveLobbyEnabled, out effectiveIsEventConversation);
                LogTalk("OnWrite traits resolved (token=" + _roomToken + ", lobbyKnown=" + effectiveLobbyKnown + ", lobby=" + effectiveLobbyEnabled + ", event=" + effectiveIsEventConversation + ").");

                string pendingDelegateId;
                bool delegationPending = _owner._talkAppointmentController.IsDelegationPending(_appointment, out pendingDelegateId);
                if (delegationPending)
                {
                    LogTalk("OnWrite delegation-pending path (token=" + _roomToken + ", delegate=" + pendingDelegateId + ").");
                }

                bool roomNameSynced = false;
                bool lobbySynced = true;
                bool descriptionSynced = false;
                bool participantsSynced;

                LogTalk("Updating room name during OnWrite (token=" + _roomToken + ").");
                roomNameSynced = _owner._talkAppointmentController.TryUpdateRoomName(_appointment, _roomToken);

                bool shouldAttemptLobbyUpdate = effectiveLobbyEnabled || !effectiveLobbyKnown;
                if (shouldAttemptLobbyUpdate)
                {
                    if (!hasStampedStartEpoch)
                    {
                        LogTalk("Lobby update skipped: X-NCTALK-START is unavailable after stamping attempt (token=" + _roomToken + ").");
                        lobbySynced = false;
                    }
                    else if (!_lastLobbyTimer.HasValue || currentStartEpoch != _lastLobbyTimer.Value)
                    {
                        LogTalk("Attempting lobby update during OnWrite (token=" + _roomToken + ", startEpoch=" + currentStartEpoch.ToString(CultureInfo.InvariantCulture) + ", lobbyKnown=" + effectiveLobbyKnown + ").");
                        if (_owner._talkAppointmentController.TryUpdateLobby(_appointment, _roomToken, effectiveIsEventConversation))
                        {
                            _lastLobbyTimer = currentStartEpoch;
                            if (!effectiveLobbyKnown)
                            {
                                _owner._talkAppointmentController.PersistLobbyTraits(_appointment, _roomToken, true);
                            }

                            LogTalk("Lobby update successful (token=" + _roomToken + ").");
                        }
                        else
                        {
                            LogTalk("Lobby update failed (token=" + _roomToken + ").");
                            lobbySynced = false;
                        }
                    }

                    ScheduleDeferredWriteLobbyVerification();
                }

                LogTalk("Updating room description during OnWrite (token=" + _roomToken + ").");
                descriptionSynced = _owner._talkAppointmentController.TryUpdateRoomDescription(_appointment, _roomToken, effectiveIsEventConversation);

                participantsSynced = _owner._talkAppointmentController.TrySyncRoomParticipants(_appointment, _roomToken, effectiveIsEventConversation);
                LogTalk(
                    "OnWrite pre-delegation sync result (token="
                    + _roomToken
                    + ", roomName="
                    + roomNameSynced
                    + ", lobby="
                    + lobbySynced
                    + ", description="
                    + descriptionSynced
                    + ", participants="
                    + participantsSynced
                    + ").");
                if (delegationPending)
                {
                    // Delegation is always executed when pending; pre-step failures are logged with
                    // explicit per-step status to keep runtime behavior transparent.
                    if (!roomNameSynced || !lobbySynced || !descriptionSynced || !participantsSynced)
                    {
                        LogTalk(
                            "Delegation continues despite pre-delegation sync failures (token="
                            + _roomToken
                            + ", roomName="
                            + roomNameSynced
                            + ", lobby="
                            + lobbySynced
                            + ", description="
                            + descriptionSynced
                            + ", participants="
                            + participantsSynced
                            + ").");
                    }

                    _owner._talkAppointmentController.TryApplyDelegation(_appointment, _roomToken);
                }
                _owner.RefreshEntryBinding(this);
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("BeforeDelete ignored (not organizer, token=" + _roomToken + ").");
                    return;
                }

                LogTalk("BeforeDelete -> EnsureRoomDeleted (Token=" + _roomToken + ").");
                EnsureRoomDeleted();
            }

            private void OnClose(ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnClose without organizer (token=" + _roomToken + ").");
                    Dispose();
                    return;
                }

                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (!_roomDeleted && _appointment != null && !_appointment.Saved)
                {
                    LogTalk("OnClose unsaved state observed (token=" + _roomToken + ", cancel=" + cancel + ").");
                    if (!cancel)
                    {
                        ScheduleUnsavedCloseCleanup();
                    }

                    return;
                }

                LogTalk("OnClose completed (token=" + _roomToken + ", deleted=" + _roomDeleted + ").");
            }

            private void ScheduleUnsavedCloseCleanup()
            {
                if (_unsavedCloseCleanupPending)
                {
                    return;
                }

                _unsavedCloseCleanupPending = true;
                _unsavedCloseCleanupAttempts = 0;
                _unsavedCloseCleanupTimer = new System.Windows.Forms.Timer();
                _unsavedCloseCleanupTimer.Interval = 1000;
                _unsavedCloseCleanupTimer.Tick += OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Start();
                LogTalk("OnClose unsaved cleanup scheduled (token=" + _roomToken + ").");
            }

            private void ScheduleDeferredWriteLobbyVerification()
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    return;
                }

                _deferredWriteLobbyAttempts = 0;
                // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
                if (_deferredWriteLobbyTimer == null)
                {
                    _deferredWriteLobbyTimer = new System.Windows.Forms.Timer();
                    _deferredWriteLobbyTimer.Interval = 750;
                    _deferredWriteLobbyTimer.Tick += OnDeferredWriteLobbyTick;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Start();
                LogTalk("Deferred post-write lobby verification scheduled (token=" + _roomToken + ").");
            }

            private void OnDeferredWriteLobbyTick(object sender, EventArgs e)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }

                _deferredWriteLobbyAttempts++;

                if (!_owner.IsOrganizer(_appointment))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (not organizer, token=" + _roomToken + ").");
                    return;
                }

                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    return;
                }

                bool effectiveLobbyKnown;
                bool effectiveLobbyEnabled;
                bool effectiveIsEventConversation;
                _owner._talkAppointmentController.ResolveRuntimeRoomTraits(_appointment, _roomToken, _lobbyEnabled, _isEventConversation, out effectiveLobbyKnown, out effectiveLobbyEnabled, out effectiveIsEventConversation);
                bool shouldAttemptLobbyUpdate = effectiveLobbyEnabled || !effectiveLobbyKnown;
                if (!shouldAttemptLobbyUpdate)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }

                long currentStartEpoch;
                if (!_owner.TryStampIcalStartEpoch(_appointment, _roomToken, out currentStartEpoch))
                {
                    if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                    {
                        _deferredWriteLobbyAttempts = 0;
                        StopDeferredWriteLobbyTimer();
                        LogTalk("Deferred post-write lobby verification stopped after missing X-NCTALK-START (token=" + _roomToken + ").");
                    }

                    return;
                }

                if (_lastLobbyTimer.HasValue && currentStartEpoch == _lastLobbyTimer.Value)
                {
                    if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                    {
                        _deferredWriteLobbyAttempts = 0;
                        StopDeferredWriteLobbyTimer();
                        LogTalk("Deferred post-write lobby verification completed without detected start change (token=" + _roomToken + ", startEpoch=" + currentStartEpoch.ToString(CultureInfo.InvariantCulture) + ").");
                    }

                    return;
                }

                LogTalk("Deferred post-write lobby verification applying update (token=" + _roomToken + ", startEpoch=" + currentStartEpoch.ToString(CultureInfo.InvariantCulture) + ").");
                if (_owner._talkAppointmentController.TryUpdateLobby(_appointment, _roomToken, effectiveIsEventConversation))
                {
                    _lastLobbyTimer = currentStartEpoch;
                    if (!effectiveLobbyKnown)
                    {
                        _owner._talkAppointmentController.PersistLobbyTraits(_appointment, _roomToken, true);
                    }

                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification successful (token=" + _roomToken + ").");
                    return;
                }

                if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification failed after retries (token=" + _roomToken + ").");
                }
            }

            private void OnUnsavedCloseCleanupTick(object sender, EventArgs e)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    return;
                }

                _unsavedCloseCleanupAttempts++;
                bool saved;
                bool hasSavedState = TryGetAppointmentSaved(_appointment, out saved);
                bool inspectorOpen = IsAppointmentOpenInAnyInspector(_appointment);
                if (saved)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True).");
                    return;
                }

                if (inspectorOpen)
                {
                    if (_unsavedCloseCleanupAttempts >= UnsavedCloseCleanupMaxAttempts)
                    {
                        _unsavedCloseCleanupPending = false;
                        _unsavedCloseCleanupAttempts = 0;
                        StopUnsavedCloseCleanupTimer();
                        LogTalk("Deferred OnClose cleanup skipped after timeout (token=" + _roomToken + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ", inspectorOpen=True).");
                        return;
                    }

                    if (_unsavedCloseCleanupAttempts == 1 || (_unsavedCloseCleanupAttempts % 5) == 0)
                    {
                        LogTalk("Deferred OnClose cleanup waiting for inspector close (token=" + _roomToken + ", attempt=" + _unsavedCloseCleanupAttempts.ToString(CultureInfo.InvariantCulture) + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ").");
                    }

                    return;
                }

                _unsavedCloseCleanupPending = false;
                _unsavedCloseCleanupAttempts = 0;
                StopUnsavedCloseCleanupTimer();

                if (!saved)
                {
                    LogTalk("Deferred OnClose cleanup deleting room (token=" + _roomToken + ").");
                    EnsureRoomDeleted();
                    return;
                }

                LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True, hasSavedState=" + hasSavedState + ", inspectorOpen=False).");
            }

            private static bool TryGetAppointmentSaved(Outlook.AppointmentItem appointment, out bool saved)
            {
                saved = false;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (appointment == null)
                {
                    return false;
                }

                try
                {
                    saved = appointment.Saved;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private void StopUnsavedCloseCleanupTimer()
            {
                // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
                if (_unsavedCloseCleanupTimer == null)
                {
                    return;
                }

                _unsavedCloseCleanupTimer.Stop();
                _unsavedCloseCleanupTimer.Tick -= OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Dispose();
                _unsavedCloseCleanupTimer = null;
            }

            private void StopDeferredWriteLobbyTimer()
            {
                // Feld wird lazy initialisiert bzw. beim Shutdown geleert; null ist hier ein erwartbarer Zustand.
                if (_deferredWriteLobbyTimer == null)
                {
                    return;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Tick -= OnDeferredWriteLobbyTick;
                _deferredWriteLobbyTimer.Dispose();
                _deferredWriteLobbyTimer = null;
            }

            private bool IsAppointmentOpenInAnyInspector(Outlook.AppointmentItem appointment)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (appointment == null || _owner == null || _owner._inspectors == null)
                {
                    return false;
                }

                string appointmentEntryId = null;
                try
                {
                    appointmentEntryId = appointment.EntryID;
                }
                catch
                {
                }

                int inspectorCount = 0;
                try
                {
                    inspectorCount = _owner._inspectors.Count;
                }
                catch
                {
                    return false;
                }

                for (int i = 1; i <= inspectorCount; i++)
                {
                    Outlook.Inspector inspector = null;
                    object currentItem = null;
                    Outlook.AppointmentItem currentAppointment = null;
                    try
                    {
                        inspector = _owner._inspectors[i];
                        // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                        if (inspector == null)
                        {
                            continue;
                        }

                        currentItem = inspector.CurrentItem;
                        currentAppointment = currentItem as Outlook.AppointmentItem;
                        // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                        if (currentAppointment == null)
                        {
                            continue;
                        }

                        if (currentAppointment == appointment)
                        {
                            return true;
                        }

                        if (!string.IsNullOrWhiteSpace(appointmentEntryId))
                        {
                            string currentEntryId = null;
                            try
                            {
                                currentEntryId = currentAppointment.EntryID;
                            }
                            catch
                            {
                            }

                            if (!string.IsNullOrWhiteSpace(currentEntryId)
                                && string.Equals(currentEntryId, appointmentEntryId, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                        if (currentAppointment != null && !ReferenceEquals(currentAppointment, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentAppointment,
                                LogCategories.Talk,
                                "Failed to release current appointment COM object during unsaved-close lookup.");
                        }

                        if (currentItem != null
                            && !ReferenceEquals(currentItem, currentAppointment)
                            && !ReferenceEquals(currentItem, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentItem,
                                LogCategories.Talk,
                                "Failed to release current item COM object during unsaved-close lookup.");
                        }

                        // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                        if (inspector != null)
                        {
                            ComInteropScope.TryRelease(
                                inspector,
                                LogCategories.Talk,
                                "Failed to release inspector COM object during unsaved-close lookup.");
                        }
                    }
                }

                return false;
            }

            private void EnsureRoomDeleted()
            {
                if (_roomDeleted || !_owner.IsOrganizer(_appointment))
                {
                    if (_roomDeleted)
                    {
                        LogTalk("EnsureRoomDeleted: room already deleted (token=" + _roomToken + ").");
                    }

                    return;
                }

                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("EnsureRoomDeleted skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner._talkAppointmentController.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    Dispose();
                    return;
                }

                if (_owner.TryDeleteRoom(_roomToken, _isEventConversation))
                {
                    _owner._talkAppointmentController.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    LogTalk("EnsureRoomDeleted successful (token=" + _roomToken + ").");
                    Dispose();
                }
                else
                {
                    LogTalk("EnsureRoomDeleted failed (token=" + _roomToken + ").");
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    LogTalk("Subscription.Dispose called again (token=" + _roomToken + ").");
                    return;
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_events != null)
                {
                    _events.BeforeDelete -= OnBeforeDelete;
                    _events.Write -= OnWrite;
                    _events.Close -= OnClose;
                }

                StopUnsavedCloseCleanupTimer();
                StopDeferredWriteLobbyTimer();

                _owner.UnregisterSubscription(_key, _roomToken, _entryId);
                _disposed = true;
                LogTalk("Subscription.Dispose completed (token=" + _roomToken + ").");
            }

            internal Outlook.AppointmentItem Appointment
            {
                get { return _appointment; }
            }

            internal string EntryId
            {
                get { return _entryId; }
            }

            internal bool IsFor(Outlook.AppointmentItem appointment)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (appointment == null)
                {
                    return false;
                }

                if (appointment == _appointment)
                {
                    return true;
                }

                string thisEntryId = _entryId;
                if (string.IsNullOrWhiteSpace(thisEntryId))
                {
                    thisEntryId = GetEntryId(_appointment);
                }

                string otherEntryId = GetEntryId(appointment);
                if (!string.IsNullOrWhiteSpace(thisEntryId)
                    && !string.IsNullOrWhiteSpace(otherEntryId)
                    && string.Equals(thisEntryId, otherEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                return false;
            }

            internal bool MatchesToken(string token)
            {
                return string.Equals(_roomToken, token, StringComparison.OrdinalIgnoreCase);
            }

            internal void UpdateEntryId(string entryId)
            {
                _entryId = entryId;
                LogTalk("Subscription EntryId updated (token=" + _roomToken + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }
        }

    }
}
