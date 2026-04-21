// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Deferred appointment subscription ensure flow.
    // Isolated from the add-in root file to keep event-restriction handling maintainable.
    public sealed partial class NextcloudTalkAddIn
    {
        private static string TryGetEntryIdForDeferredKey(Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment != null ? appointment.EntryID : null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetGlobalAppointmentIdForDeferredKey(Outlook.AppointmentItem appointment)
        {
            // Deferred ensure path can run after appointment release; null means "no stable key".
            if (appointment == null)
            {
                return null;
            }
            try
            {
                return appointment.GlobalAppointmentID;
            }
            catch
            {
                return null;
            }
        }

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment)
        {
            EnsureSubscriptionForAppointment(appointment, true);
        }

        private void EnsureSubscriptionForAppointment(Outlook.AppointmentItem appointment, bool allowDeferredRetry)
        {            if (appointment == null)
            {
                return;
            }
            try
            {
                string roomToken = ResolveRoomTokenForAppointment(appointment);
                if (string.IsNullOrWhiteSpace(roomToken))
                {
                    return;
                }
                string entryId = GetEntryId(appointment);
                if (!string.IsNullOrEmpty(entryId))
                {
                    AppointmentSubscription existingByEntry;
                    if (_subscriptionByEntryId.TryGetValue(entryId, out existingByEntry))
                    {
                        if (existingByEntry.IsFor(appointment) && existingByEntry.MatchesToken(roomToken))
                        {
                            return;
                        }

                        existingByEntry.Dispose();
                    }
                }

                AppointmentSubscription existingByToken;
                if (_subscriptionByToken.TryGetValue(roomToken, out existingByToken))
                {
                    if (existingByToken.IsFor(appointment))
                    {
                        return;
                    }

                    existingByToken.Dispose();
                }
                bool lobbyKnown;
                bool lobbyEnabled;
                bool isEventConversation;
                _talkAppointmentController.ResolveRuntimeRoomTraits(
                    appointment,
                    roomToken,
                    false,
                    false,
                    out lobbyKnown,
                    out lobbyEnabled,
                    out isEventConversation);

                RegisterSubscription(appointment, roomToken, lobbyEnabled, isEventConversation);
            }
            catch (COMException ex)
            {
                if (IsOutlookEventProcedureRestriction(ex))
                {
                    if (allowDeferredRetry)
                    {
                        QueueDeferredAppointmentSubscriptionEnsure(appointment, ex);
                        return;
                    }

                    LogDeferredAppointmentEnsureRestriction(
                        "Deferred appointment subscription ensure skipped: Outlook still blocks property access in the current event context (hresult=0x" +
                        ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                        ").");
                    return;
                }

                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure subscription for appointment.", ex);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to ensure subscription for appointment.", ex);
            }
        }

        private bool QueueDeferredAppointmentSubscriptionEnsure(Outlook.AppointmentItem appointment, COMException triggerException)
        {            if (appointment == null)
            {
                return false;
            }

            SynchronizationContext context = _uiSynchronizationContext;
            // UI context can be unavailable on shutdown/background paths; guard avoids follow-up failures.
            if (context == null)
            {
                LogDeferredAppointmentEnsureRestriction(
                    "Deferred appointment subscription ensure unavailable: UI synchronization context is missing (hresult=0x" +
                    triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                    ").");
                return false;
            }
            string entryId = TryGetEntryIdForDeferredKey(appointment);
            string globalAppointmentId = TryGetGlobalAppointmentIdForDeferredKey(appointment);
            string ensureKey = !string.IsNullOrWhiteSpace(entryId)
                ? ("entry:" + entryId.Trim())
                : (!string.IsNullOrWhiteSpace(globalAppointmentId)
                    ? ("gid:" + globalAppointmentId.Trim())
                    : ("obj:" + RuntimeHelpers.GetHashCode(appointment).ToString("X8", CultureInfo.InvariantCulture)));

            if (ensureKey.StartsWith("obj:", StringComparison.Ordinal))
            {
                if (_deferredAppointmentEnsureState.ShouldLogUnstableIdentityRestriction(DateTime.UtcNow))
                {
                    LogDeferredAppointmentEnsureRestriction(
                        "Deferred appointment subscription ensure suppressed: unstable appointment identity during Outlook event restriction (hresult=0x" +
                        triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                        ").");
                }
                return false;
            }
            if (!_deferredAppointmentEnsureState.TryQueuePendingKey(ensureKey))
            {
                return true;
            }

            LogDeferredAppointmentEnsureRestriction(
                "Deferred appointment subscription ensure queued (key=" + ensureKey +
                ", hresult=0x" + triggerException.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) +
                ").");

            context.Post(
                _ =>
                {
                    try
                    {
                        EnsureSubscriptionForAppointment(appointment, false);
                    }
                    finally
                    {
                        _deferredAppointmentEnsureState.DequeuePendingKey(ensureKey);
                    }
                },
                null);

            return true;
        }

        private static bool IsOutlookEventProcedureRestriction(COMException ex)
        {
            // Null means there is no valid error context; keep this check intentionally defensive.
            if (ex == null)
            {
                return false;
            }
            return (ex.ErrorCode & 0xFFFF) == 0x0108;
        }
    }
}
