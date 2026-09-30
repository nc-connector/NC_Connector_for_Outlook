// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Captures Outlook appointment state and executes remote Talk synchronization.
    internal sealed partial class TalkAppointmentController
    {
        internal TalkAppointmentSyncSnapshot CaptureRemoteSyncSnapshot(
            Outlook.AppointmentItem appointment,
            string roomToken,
            string roomUrl,
            bool fallbackLobbyEnabled,
            bool fallbackIsEventConversation,
            long? lastLobbyEpoch,
            bool lobbyOnly)
        {
            if (appointment == null
                || string.IsNullOrWhiteSpace(roomToken)
                || _owner.CurrentSettings == null)
            {
                return null;
            }

            AddinSettings settings = _owner.CurrentSettings.Clone();
            var configuration = new TalkServiceConfiguration(
                settings.ServerUrl,
                settings.Username,
                settings.AppPassword);
            if (!configuration.IsComplete())
            {
                NextcloudTalkAddIn.LogTalkMessage(
                    "Appointment synchronization skipped: Talk credentials are incomplete.");
                return null;
            }

            bool lobbyKnown;
            bool lobbyEnabled;
            bool isEventConversation;
            ResolveRuntimeRoomTraits(
                appointment,
                roomToken,
                fallbackLobbyEnabled,
                fallbackIsEventConversation,
                out lobbyKnown,
                out lobbyEnabled,
                out isEventConversation);

            long startEpoch;
            bool hasStartEpoch = PersistCoreIcalProperties(
                appointment,
                roomToken,
                roomUrl,
                lobbyEnabled,
                isEventConversation,
                out startEpoch);

            string delegateId = string.Empty;
            bool delegationPending = !lobbyOnly
                                     && IsDelegationPending(
                                         appointment,
                                         out delegateId);
            if (!delegationPending)
            {
                delegateId = string.Empty;
            }

            var snapshot = new TalkAppointmentSyncSnapshot
            {
                RoomToken = roomToken.Trim(),
                RoomUrl = roomUrl ?? string.Empty,
                LobbyOnly = lobbyOnly,
                IsEventConversation = isEventConversation,
                LobbyKnown = lobbyKnown,
                LobbyEnabled = lobbyEnabled,
                UpdateLobby = hasStartEpoch
                              && (lobbyEnabled || !lobbyKnown)
                              && (!lastLobbyEpoch.HasValue
                                  || lastLobbyEpoch.Value != startEpoch),
                StartEpoch = startEpoch,
                End = ReadAppointmentEnd(appointment),
                Configuration = configuration,
                CacheHours = settings.IfbCacheHours,
                ProfileScope = _owner.OutlookProfileScope,
                DataDirectory = _owner.SettingsStorage != null
                    ? _owner.SettingsStorage.DataDirectory
                    : string.Empty,
                DelegationPending = delegationPending,
                DelegateUserId = delegateId
            };

            if (lobbyOnly)
            {
                return snapshot;
            }

            snapshot.RoomName = isEventConversation
                ? string.Empty
                : GetNormalizedRoomName(appointment);
            snapshot.Description = isEventConversation
                ? string.Empty
                : BuildDescriptionPayload(appointment);
            snapshot.AddUsers =
                HasUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers)
                && GetUserPropertyBool(
                    appointment,
                    NextcloudTalkAddIn.IcalAddUsers);
            snapshot.AddGuests =
                HasUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests)
                && GetUserPropertyBool(
                    appointment,
                    NextcloudTalkAddIn.IcalAddGuests);
            if (snapshot.AddUsers || snapshot.AddGuests)
            {
                snapshot.AttendeeEmails =
                    NextcloudTalkAddIn.GetAppointmentAttendeeEmails(appointment);
            }
            return snapshot;
        }

        internal TalkAppointmentSyncResult ExecuteRemoteSync(
            TalkAppointmentSyncSnapshot snapshot)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.RoomToken))
            {
                return new TalkAppointmentSyncResult(string.Empty);
            }

            var result = new TalkAppointmentSyncResult(snapshot.RoomToken);
            if (snapshot.Configuration == null
                || !snapshot.Configuration.IsComplete())
            {
                return result;
            }

            AddinSettings settings = _owner.CurrentSettings;
            if (settings != null && settings.IsEnterpriseRollout)
            {
                BackendPolicyStatus status = _owner.FetchEnterpriseRolloutPolicyStatus(
                    snapshot.Configuration, "enterprise_talk_sync");
                if (!string.IsNullOrEmpty(PolicyUiHelper.GetEnterpriseRolloutNotice(settings, status)))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Managed rollout access unavailable; Talk update skipped.");
                    return result;
                }
            }

            var service = new TalkService(snapshot.Configuration);
            if (snapshot.LobbyOnly)
            {
                ExecuteLobbySync(service, snapshot, result);
                return result;
            }

            ExecuteRoomNameSync(service, snapshot, result);
            ExecuteLobbySync(service, snapshot, result);
            ExecuteDescriptionSync(service, snapshot, result);
            ExecuteParticipantSync(service, snapshot);
            ExecuteDelegationSync(service, snapshot, result);
            return result;
        }

        internal bool ApplyRemoteSyncResult(
            Outlook.AppointmentItem appointment,
            TalkAppointmentSyncResult result)
        {
            if (appointment == null || result == null)
            {
                return false;
            }

            bool handoffPersisted = false;
            if (result.MarkEventConversation)
            {
                PersistEventConversationTraits(
                    appointment,
                    result.RoomToken);
            }
            if (result.ClearDelegation)
            {
                try
                {
                    ClearDelegationProperties(appointment);
                    appointment.Save();
                }
                catch (Exception ex)
                {
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Delegation,
                            ex.Message));
                }
            }
            else if (result.DelegationApplied)
            {
                try
                {
                    SetUserProperty(
                        appointment,
                        NextcloudTalkAddIn.IcalDelegated,
                        Outlook.OlUserPropertyType.olText,
                        "TRUE");
                    RemoveUserProperty(
                        appointment,
                        NextcloudTalkAddIn.IcalDelegateReady);
                    if (!GetUserPropertyBool(
                            appointment,
                            NextcloudTalkAddIn.IcalDelegated)
                        || HasUserProperty(
                            appointment,
                            NextcloudTalkAddIn.IcalDelegateReady))
                    {
                        throw new InvalidOperationException(
                            "Outlook did not accept the local delegation state.");
                    }

                    appointment.Save();
                    handoffPersisted = true;
                }
                catch (Exception ex)
                {
                    SetUserProperty(
                        appointment,
                        NextcloudTalkAddIn.IcalDelegated,
                        Outlook.OlUserPropertyType.olText,
                        "FALSE");
                    SetUserProperty(
                        appointment,
                        NextcloudTalkAddIn.IcalDelegateReady,
                        Outlook.OlUserPropertyType.olText,
                        "TRUE");
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Delegation,
                            "The moderator was promoted, but Outlook could not save the handoff: "
                            + ex.Message));
                }
            }

            for (int i = 0; i < result.Warnings.Count; i++)
            {
                TalkAppointmentSyncWarning warning = result.Warnings[i];
                if (warning == null)
                {
                    continue;
                }

                string message;
                switch (warning.Kind)
                {
                    case TalkAppointmentSyncWarningKind.Lobby:
                        message = string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.WarningLobbyUpdateFailed,
                            warning.Detail);
                        break;
                    case TalkAppointmentSyncWarningKind.Description:
                        message = string.Format(
                            CultureInfo.CurrentCulture,
                            Strings.WarningDescriptionUpdateFailed,
                            warning.Detail);
                        break;
                    case TalkAppointmentSyncWarningKind.Delegation:
                        message = string.IsNullOrWhiteSpace(warning.Detail)
                            ? Strings.WarningModeratorTransferFailed
                            : string.Format(
                                CultureInfo.CurrentCulture,
                                Strings.WarningModeratorTransferFailedWithReasonFormat,
                                warning.Detail);
                        break;
                    default:
                        continue;
                }
                NextcloudTalkAddIn.ShowWarningDialog(message);
            }
            return handoffPersisted;
        }

        private static DateTime ReadAppointmentEnd(
            Outlook.AppointmentItem appointment)
        {
            try
            {
                return appointment != null
                    ? appointment.End
                    : DateTime.MinValue;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to read Appointment.End for background synchronization.",
                    ex);
                return DateTime.MinValue;
            }
        }

        private static void ExecuteRoomNameSync(
            TalkService service,
            TalkAppointmentSyncSnapshot snapshot,
            TalkAppointmentSyncResult result)
        {
            if (snapshot.IsEventConversation
                || string.IsNullOrWhiteSpace(snapshot.RoomName))
            {
                return;
            }

            try
            {
                service.UpdateRoomName(
                    snapshot.RoomToken,
                    snapshot.RoomName);
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    result.MarkEventConversation = true;
                    return;
                }
                if (!IsMissingOrForbiddenRoomMutationError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage(
                        "Room name background synchronization failed: "
                        + ex.Message);
                }
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage(
                    "Room name background synchronization failed: "
                    + ex.Message);
            }
        }

        private static void ExecuteLobbySync(
            TalkService service,
            TalkAppointmentSyncSnapshot snapshot,
            TalkAppointmentSyncResult result)
        {
            if (!snapshot.UpdateLobby)
            {
                return;
            }

            try
            {
                DateTime startUtc =
                    DateTimeOffset.FromUnixTimeSeconds(snapshot.StartEpoch)
                        .UtcDateTime;
                DateTime end = snapshot.End == DateTime.MinValue
                    ? startUtc
                    : snapshot.End;
                service.UpdateLobby(
                    snapshot.RoomToken,
                    startUtc,
                    end,
                    snapshot.IsEventConversation);
                result.AppliedLobbyEpoch = snapshot.StartEpoch;
            }
            catch (TalkServiceException ex)
            {
                if (!IsMissingOrForbiddenRoomMutationError(ex))
                {
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Lobby,
                            ex.Message));
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    new TalkAppointmentSyncWarning(
                        TalkAppointmentSyncWarningKind.Lobby,
                        ex.Message));
            }
        }

        private static void ExecuteDescriptionSync(
            TalkService service,
            TalkAppointmentSyncSnapshot snapshot,
            TalkAppointmentSyncResult result)
        {
            if (snapshot.IsEventConversation)
            {
                return;
            }

            try
            {
                service.UpdateDescription(
                    snapshot.RoomToken,
                    snapshot.Description ?? string.Empty);
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    result.MarkEventConversation = true;
                    return;
                }
                if (!IsMissingOrForbiddenRoomMutationError(ex))
                {
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Description,
                            ex.Message));
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    new TalkAppointmentSyncWarning(
                        TalkAppointmentSyncWarningKind.Description,
                        ex.Message));
            }
        }

        private static void ExecuteParticipantSync(
            TalkService service,
            TalkAppointmentSyncSnapshot snapshot)
        {
            if ((!snapshot.AddUsers && !snapshot.AddGuests)
                || snapshot.AttendeeEmails == null
                || snapshot.AttendeeEmails.Count == 0)
            {
                return;
            }

            try
            {
                var cache = new IfbAddressBookCache(
                    snapshot.DataDirectory,
                    snapshot.ProfileScope);
                string currentUserId =
                    NextcloudUserIdentityService.ResolveCurrentUserId(
                        snapshot.Configuration);
                string selfEmail;
                cache.TryGetPrimaryEmailForUid(
                    snapshot.Configuration,
                    snapshot.CacheHours,
                    currentUserId,
                    out selfEmail);

                for (int i = 0; i < snapshot.AttendeeEmails.Count; i++)
                {
                    string email = snapshot.AttendeeEmails[i];
                    if (string.IsNullOrWhiteSpace(email)
                        || (!string.IsNullOrWhiteSpace(selfEmail)
                            && string.Equals(
                                email,
                                selfEmail,
                                StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    string uid;
                    if (cache.TryGetUid(
                            snapshot.Configuration,
                            snapshot.CacheHours,
                            email,
                            out uid)
                        && !string.IsNullOrWhiteSpace(uid))
                    {
                        if (snapshot.AddUsers)
                        {
                            service.AddUserParticipant(
                                snapshot.RoomToken,
                                uid);
                        }
                    }
                    else if (snapshot.AddGuests)
                    {
                        service.AddGuestParticipant(
                            snapshot.RoomToken,
                            email);
                    }
                }
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage(
                    "Participant background synchronization failed: "
                    + ex.Message);
            }
        }

        private static void ExecuteDelegationSync(
            TalkService service,
            TalkAppointmentSyncSnapshot snapshot,
            TalkAppointmentSyncResult result)
        {
            if (!snapshot.DelegationPending
                || string.IsNullOrWhiteSpace(snapshot.DelegateUserId))
            {
                return;
            }

            try
            {
                string currentUserId =
                    NextcloudUserIdentityService.ResolveCurrentUserId(
                        snapshot.Configuration);
                string delegateUserId = ResolveDelegateUserId(
                    snapshot,
                    currentUserId);
                if (string.Equals(
                        delegateUserId,
                        currentUserId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    result.ClearDelegation = true;
                    NextcloudTalkAddIn.LogTalkMessage(
                        "Delegation ignored because the selected moderator is the current user.");
                    return;
                }

                service.AddUserParticipant(
                    snapshot.RoomToken,
                    delegateUserId);
                List<TalkParticipant> participants =
                    service.GetParticipants(snapshot.RoomToken);
                int attendeeId = FindParticipantAttendeeId(
                    participants,
                    delegateUserId);
                if (attendeeId <= 0)
                {
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Delegation,
                            "The selected user was not found in the room."));
                    return;
                }

                string promoteError;
                if (!service.PromoteModerator(
                        snapshot.RoomToken,
                        attendeeId,
                        out promoteError))
                {
                    result.Warnings.Add(
                        new TalkAppointmentSyncWarning(
                            TalkAppointmentSyncWarningKind.Delegation,
                            promoteError));
                    return;
                }

                result.DelegationApplied = true;
                result.HandoffConfiguration = snapshot.Configuration;
                NextcloudTalkAddIn.LogTalkMessage(
                    "Moderator promoted; waiting for Outlook to persist the handoff.");
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    new TalkAppointmentSyncWarning(
                        TalkAppointmentSyncWarningKind.Delegation,
                        ex.Message));
            }
        }

        private static string ResolveDelegateUserId(
            TalkAppointmentSyncSnapshot snapshot,
            string currentUserId)
        {
            string candidate = snapshot.DelegateUserId.Trim();
            if (string.Equals(
                    candidate,
                    snapshot.Configuration.Username,
                    StringComparison.OrdinalIgnoreCase))
            {
                return currentUserId;
            }

            if (candidate.IndexOf('@') >= 0)
            {
                var cache = new IfbAddressBookCache(
                    snapshot.DataDirectory,
                    snapshot.ProfileScope);
                string mappedUserId;
                if (cache.TryGetUid(
                        snapshot.Configuration,
                        snapshot.CacheHours,
                        candidate,
                        out mappedUserId)
                    && !string.IsNullOrWhiteSpace(mappedUserId))
                {
                    return mappedUserId.Trim();
                }
            }
            return candidate;
        }

        private static int FindParticipantAttendeeId(
            IList<TalkParticipant> participants,
            string userId)
        {
            if (participants == null)
            {
                return 0;
            }

            for (int i = 0; i < participants.Count; i++)
            {
                TalkParticipant participant = participants[i];
                if (participant != null
                    && string.Equals(
                        participant.ActorType,
                        "users",
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(
                        participant.ActorId,
                        userId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return participant.AttendeeId;
                }
            }
            return 0;
        }

        private static void ClearDelegationProperties(
            Outlook.AppointmentItem appointment)
        {
            RemoveUserProperty(
                appointment,
                NextcloudTalkAddIn.IcalDelegate);
            RemoveUserProperty(
                appointment,
                NextcloudTalkAddIn.IcalDelegateName);
            RemoveUserProperty(
                appointment,
                NextcloudTalkAddIn.IcalDelegated);
            RemoveUserProperty(
                appointment,
                NextcloudTalkAddIn.IcalDelegateReady);
        }
    }
}
