/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed class TalkAppointmentController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal TalkAppointmentController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal void ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            // COM lifecycle guard: appointment/result may be unavailable during event teardown.
            if (appointment == null || result == null)
            {
                return;
            }

            NextcloudTalkAddIn.LogTalkMessage(
                "Writing talk data to appointment (token="
                + result.RoomToken
                + ", type="
                + request.RoomType
                + ", lobby="
                + request.LobbyEnabled
                + ", search="
                + request.SearchVisible
                + ", passwordSet="
                + (!string.IsNullOrEmpty(request.Password))
                + ", addUsers="
                + request.AddUsers
                + ", addGuests="
                + request.AddGuests
                + ", delegate="
                + (string.IsNullOrEmpty(request.DelegateModeratorId) ? "n/a" : request.DelegateModeratorId)
                + ").");

            if (!string.IsNullOrWhiteSpace(request.Title))
            {
                appointment.Subject = request.Title.Trim();
            }

            appointment.Location = result.RoomUrl;
            string normalizedDescriptionType = NextcloudTalkAddIn.NormalizeTalkEventDescriptionType(request.DescriptionType);
            if (string.Equals(normalizedDescriptionType, "html", StringComparison.OrdinalIgnoreCase))
            {
                // AppointmentItem HTML body read is intentionally disabled for compatibility.
                string updatedHtmlBody = NextcloudTalkAddIn.UpdateHtmlBodyWithTalkBlock(
                    string.Empty,
                    appointment.Body,
                    result.RoomUrl,
                    request.Password,
                    request.DescriptionLanguage,
                    request.InvitationTemplate);
                if (!NextcloudTalkAddIn.TryWriteAppointmentHtmlBody(appointment, updatedHtmlBody))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Appointment HTML editor unavailable, falling back to plain-text Talk block insertion.");
                    appointment.Body = NextcloudTalkAddIn.UpdateBodyWithTalkBlock(
                        appointment.Body,
                        result.RoomUrl,
                        request.Password,
                        request.DescriptionLanguage,
                        request.InvitationTemplate);
                }
            }
            else
            {
                appointment.Body = NextcloudTalkAddIn.UpdateBodyWithTalkBlock(
                    appointment.Body,
                    result.RoomUrl,
                    request.Password,
                    request.DescriptionLanguage,
                    request.InvitationTemplate);
            }

            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyToken, Outlook.OlUserPropertyType.olText, result.RoomToken);
            var storedRoomType = result.CreatedAsEventConversation ? TalkRoomType.EventConversation : TalkRoomType.StandardRoom;
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyRoomType, Outlook.OlUserPropertyType.olText, storedRoomType.ToString());
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyLobby, Outlook.OlUserPropertyType.olYesNo, request.LobbyEnabled);
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertySearchVisible, Outlook.OlUserPropertyType.olYesNo, request.SearchVisible);
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyPasswordSet, Outlook.OlUserPropertyType.olYesNo, !string.IsNullOrEmpty(request.Password));
            NextcloudTalkAddIn.LogTalkMessage("User fields updated (token set, lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ").");

            long? startEpoch = TimeUtilities.ToUnixTimeSeconds(appointment.Start);
            long? endEpoch = TimeUtilities.ToUnixTimeSeconds(appointment.End);
            if (startEpoch.HasValue)
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyStartEpoch, Outlook.OlUserPropertyType.olText, startEpoch.Value.ToString(CultureInfo.InvariantCulture));
            }

            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyDataVersion, Outlook.OlUserPropertyType.olNumber, NextcloudTalkAddIn.PropertyVersionValue);
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyAddUsers, Outlook.OlUserPropertyType.olYesNo, request.AddUsers);
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyAddGuests, Outlook.OlUserPropertyType.olYesNo, request.AddGuests);

            if (!string.IsNullOrWhiteSpace(request.DelegateModeratorId))
            {
                string delegateId = request.DelegateModeratorId.Trim();
                string delegateName = !string.IsNullOrWhiteSpace(request.DelegateModeratorName) ? request.DelegateModeratorName.Trim() : delegateId;

                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegateId, Outlook.OlUserPropertyType.olText, delegateId);
                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegated, Outlook.OlUserPropertyType.olYesNo, false);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate, Outlook.OlUserPropertyType.olText, delegateId);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName, Outlook.OlUserPropertyType.olText, delegateName);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated, Outlook.OlUserPropertyType.olText, "FALSE");
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady, Outlook.OlUserPropertyType.olText, "TRUE");
            }
            else
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegateId);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegated);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
            }

            SetUserProperty(appointment, NextcloudTalkAddIn.IcalToken, Outlook.OlUserPropertyType.olText, result.RoomToken);
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalUrl, Outlook.OlUserPropertyType.olText, result.RoomUrl);
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalLobby, Outlook.OlUserPropertyType.olText, request.LobbyEnabled ? "TRUE" : "FALSE");
            if (startEpoch.HasValue)
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalStart, Outlook.OlUserPropertyType.olText, startEpoch.Value.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
            }

            SetUserProperty(appointment, NextcloudTalkAddIn.IcalEvent, Outlook.OlUserPropertyType.olText, result.CreatedAsEventConversation ? "event" : "standard");
            string objectId = (startEpoch.HasValue && endEpoch.HasValue)
                ? (startEpoch.Value.ToString(CultureInfo.InvariantCulture) + "#" + endEpoch.Value.ToString(CultureInfo.InvariantCulture))
                : null;
            if (!string.IsNullOrWhiteSpace(objectId))
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId, Outlook.OlUserPropertyType.olText, objectId);
            }
            else
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
            }

            SetUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers, Outlook.OlUserPropertyType.olText, request.AddUsers ? "TRUE" : "FALSE");
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests, Outlook.OlUserPropertyType.olText, request.AddGuests ? "TRUE" : "FALSE");
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalAddParticipants, Outlook.OlUserPropertyType.olText, (request.AddUsers || request.AddGuests) ? "TRUE" : "FALSE");

            _owner.RegisterSubscription(appointment, result);
            TryUpdateRoomDescription(appointment, result.RoomToken, result.CreatedAsEventConversation);
        }

        internal bool TryStampIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            startEpoch = 0;
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                NextcloudTalkAddIn.LogTalkMessage("Failed to stamp X-NCTALK-START: appointment is null (token=" + (roomToken ?? "n/a") + ").");
                return false;
            }

            DateTime start;
            try
            {
                start = appointment.Start;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Appointment.Start while stamping X-NCTALK-START (token=" + (roomToken ?? "n/a") + ").", ex);
                return false;
            }

            long? epoch = TimeUtilities.ToUnixTimeSeconds(start);
            if (!epoch.HasValue || epoch.Value <= 0)
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyStartEpoch);
                NextcloudTalkAddIn.LogTalkMessage("Failed to stamp X-NCTALK-START: Appointment.Start is missing/invalid (token=" + (roomToken ?? "n/a") + ", start=" + start.ToString("o", CultureInfo.InvariantCulture) + ").");
                return false;
            }

            startEpoch = epoch.Value;
            string value = startEpoch.ToString(CultureInfo.InvariantCulture);
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalStart, Outlook.OlUserPropertyType.olText, value);
            SetUserProperty(appointment, NextcloudTalkAddIn.PropertyStartEpoch, Outlook.OlUserPropertyType.olText, value);
            NextcloudTalkAddIn.LogTalkMessage("X-NCTALK-START stamped (token=" + (roomToken ?? "n/a") + ", startEpoch=" + value + ").");
            return true;
        }

        internal bool TryReadRequiredIcalStartEpoch(Outlook.AppointmentItem appointment, string roomToken, out long startEpoch)
        {
            startEpoch = 0;
            string rawValue = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalStart);
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                NextcloudTalkAddIn.LogTalkMessage("Lobby update blocked: X-NCTALK-START is missing (token=" + (roomToken ?? "n/a") + ").");
                return false;
            }

            long parsed;
            if (!long.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
            {
                NextcloudTalkAddIn.LogTalkMessage("Lobby update blocked: X-NCTALK-START is invalid (token=" + (roomToken ?? "n/a") + ", value='" + rawValue + "').");
                return false;
            }

            startEpoch = parsed;
            return true;
        }

        internal static long? GetIcalStartEpochOrNull(Outlook.AppointmentItem appointment)
        {
            string rawValue = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalStart);
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            long parsed;
            if (!long.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
            {
                return null;
            }

            return parsed;
        }

        internal void ResolveRuntimeRoomTraits(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool fallbackLobbyEnabled,
            bool fallbackIsEventConversation,
            out bool lobbyKnown,
            out bool lobbyEnabled,
            out bool isEventConversation)
        {
            bool? resolvedLobby = null;
            bool? resolvedEventConversation = null;

            try
            {
                bool hasLobbyFlag = HasUserProperty(appointment, NextcloudTalkAddIn.IcalLobby)
                    || HasUserProperty(appointment, NextcloudTalkAddIn.PropertyLobby);
                if (hasLobbyFlag)
                {
                    resolvedLobby = GetUserPropertyBoolPrefer(appointment, NextcloudTalkAddIn.IcalLobby, NextcloudTalkAddIn.PropertyLobby);
                }

                TalkRoomType? roomType = GetRoomType(appointment);
                if (roomType.HasValue)
                {
                    resolvedEventConversation = roomType.Value == TalkRoomType.EventConversation;
                }

                if (!resolvedEventConversation.HasValue && TryIsEventConversationFromObjectId(appointment))
                {
                    resolvedEventConversation = true;
                    PersistEventConversationTraits(appointment, roomToken);
                    NextcloudTalkAddIn.LogTalkMessage("Conversation type inferred from X-NCTALK-OBJECTID (token=" + roomToken + ").");
                }

                TryHydrateMissingRoomTraitsFromServer(appointment, roomToken, ref resolvedLobby, ref resolvedEventConversation);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve runtime room traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }

            lobbyKnown = resolvedLobby.HasValue;
            lobbyEnabled = resolvedLobby.HasValue ? resolvedLobby.Value : fallbackLobbyEnabled;
            isEventConversation = resolvedEventConversation.HasValue ? resolvedEventConversation.Value : fallbackIsEventConversation;
        }

        internal void PersistLobbyTraits(Outlook.AppointmentItem appointment, string roomToken, bool lobbyEnabled)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return;
            }

            try
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalLobby, Outlook.OlUserPropertyType.olText, lobbyEnabled ? "TRUE" : "FALSE");
                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyLobby, Outlook.OlUserPropertyType.olYesNo, lobbyEnabled);
                NextcloudTalkAddIn.LogTalkMessage("Lobby flag persisted from runtime update (token=" + (roomToken ?? "n/a") + ", lobby=" + lobbyEnabled + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist lobby traits from runtime update (token=" + (roomToken ?? "n/a") + ").", ex);
            }
        }

        internal bool IsDelegatedToOtherUser(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyTextPrefer(appointment, NextcloudTalkAddIn.IcalDelegate, NextcloudTalkAddIn.PropertyDelegateId) ?? string.Empty;
            bool delegated = GetUserPropertyBoolPrefer(appointment, NextcloudTalkAddIn.IcalDelegated, NextcloudTalkAddIn.PropertyDelegated);

            if (!delegated || string.IsNullOrWhiteSpace(delegateId))
            {
                delegateId = string.Empty;
                return false;
            }

            string currentUser = _owner.CurrentSettings != null ? (_owner.CurrentSettings.Username ?? string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(currentUser))
            {
                return true;
            }

            return !string.Equals(delegateId.Trim(), currentUser.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        internal bool IsDelegationPending(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyTextPrefer(appointment, NextcloudTalkAddIn.IcalDelegate, NextcloudTalkAddIn.PropertyDelegateId) ?? string.Empty;
            bool delegated = GetUserPropertyBoolPrefer(appointment, NextcloudTalkAddIn.IcalDelegated, NextcloudTalkAddIn.PropertyDelegated);

            if (delegated)
            {
                delegateId = string.Empty;
                return false;
            }

            return !string.IsNullOrWhiteSpace(delegateId);
        }

        internal bool TrySyncRoomParticipants(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken) || _owner.CurrentSettings == null)
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                NextcloudTalkAddIn.LogTalkMessage("Participant sync skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            bool addParticipantsLegacy = GetUserPropertyBool(appointment, NextcloudTalkAddIn.IcalAddParticipants);
            bool hasAddUsers = HasUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers) || HasUserProperty(appointment, NextcloudTalkAddIn.PropertyAddUsers);
            bool hasAddGuests = HasUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests) || HasUserProperty(appointment, NextcloudTalkAddIn.PropertyAddGuests);

            bool addUsers = hasAddUsers ? GetUserPropertyBoolPrefer(appointment, NextcloudTalkAddIn.IcalAddUsers, NextcloudTalkAddIn.PropertyAddUsers) : addParticipantsLegacy;
            bool addGuests = hasAddGuests ? GetUserPropertyBoolPrefer(appointment, NextcloudTalkAddIn.IcalAddGuests, NextcloudTalkAddIn.PropertyAddGuests) : addParticipantsLegacy;
            if (!addUsers && !addGuests)
            {
                return true;
            }

            var configuration = new TalkServiceConfiguration(_owner.CurrentSettings.ServerUrl, _owner.CurrentSettings.Username, _owner.CurrentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                NextcloudTalkAddIn.LogTalkMessage("Participant sync failed: talk service configuration incomplete.");
                return false;
            }

            List<string> attendeeEmails = NextcloudTalkAddIn.GetAppointmentAttendeeEmails(appointment);
            if (attendeeEmails.Count == 0)
            {
                return true;
            }

            var cache = new IfbAddressBookCache(_owner.SettingsStorage != null ? _owner.SettingsStorage.DataDirectory : null);
            string selfEmail;
            cache.TryGetPrimaryEmailForUid(configuration, _owner.CurrentSettings.IfbCacheHours, _owner.CurrentSettings.Username, out selfEmail);

            int userAdds = 0;
            int guestAdds = 0;
            int skipped = 0;
            bool hadFailures = false;

            try
            {
                var service = _owner.CreateTalkService();
                for (int i = 0; i < attendeeEmails.Count; i++)
                {
                    string email = attendeeEmails[i];
                    if (string.IsNullOrWhiteSpace(email))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(selfEmail)
                        && string.Equals(email, selfEmail, StringComparison.OrdinalIgnoreCase))
                    {
                        skipped++;
                        continue;
                    }

                    string uid;
                    if (cache.TryGetUid(configuration, _owner.CurrentSettings.IfbCacheHours, email, out uid)
                        && !string.IsNullOrWhiteSpace(uid))
                    {
                        if (!addUsers)
                        {
                            skipped++;
                            continue;
                        }

                        if (service.AddUserParticipant(roomToken, uid))
                        {
                            userAdds++;
                        }
                        else
                        {
                            hadFailures = true;
                            NextcloudTalkAddIn.LogTalkMessage("Participant sync failed while adding Nextcloud user (uid=" + uid + ", token=" + roomToken + ").");
                        }

                        continue;
                    }

                    if (!addGuests)
                    {
                        skipped++;
                        continue;
                    }

                    if (service.AddGuestParticipant(roomToken, email))
                    {
                        guestAdds++;
                    }
                    else
                    {
                        hadFailures = true;
                        NextcloudTalkAddIn.LogTalkMessage("Participant sync failed while adding guest (email=" + email + ", token=" + roomToken + ").");
                    }
                }
            }
            catch (TalkServiceException ex)
            {
                hadFailures = true;
                NextcloudTalkAddIn.LogTalkMessage("Participant sync failed: " + ex.Message);
            }
            catch (Exception ex)
            {
                hadFailures = true;
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error during participant sync: " + ex.Message);
            }

            NextcloudTalkAddIn.LogTalkMessage("Participant sync completed (users=" + userAdds + ", guests=" + guestAdds + ", skipped=" + skipped + ", failed=" + hadFailures + ", token=" + roomToken + ").");
            return !hadFailures;
        }

        internal void TryApplyDelegation(Outlook.AppointmentItem appointment, string roomToken)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken) || _owner.CurrentSettings == null)
            {
                return;
            }

            string delegateId;
            if (!IsDelegationPending(appointment, out delegateId))
            {
                return;
            }

            string currentUser = _owner.CurrentSettings.Username ?? string.Empty;
            if (!string.IsNullOrEmpty(currentUser)
                && string.Equals(delegateId, currentUser, StringComparison.OrdinalIgnoreCase))
            {
                NextcloudTalkAddIn.LogTalkMessage("Delegation ignored (delegate == current user).");
                RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegateId);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegated);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
                return;
            }

            try
            {
                var service = _owner.CreateTalkService();
                NextcloudTalkAddIn.LogTalkMessage("Starting delegation (token=" + roomToken + ", delegate=" + delegateId + ").");

                service.AddUserParticipant(roomToken, delegateId);

                List<TalkParticipant> participants = service.GetParticipants(roomToken);
                int attendeeId = 0;
                for (int i = 0; i < participants.Count; i++)
                {
                    TalkParticipant participant = participants[i];
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (participant == null)
                    {
                        continue;
                    }

                    if (string.Equals(participant.ActorType, "users", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(participant.ActorId, delegateId, StringComparison.OrdinalIgnoreCase))
                    {
                        attendeeId = participant.AttendeeId;
                        break;
                    }
                }

                if (attendeeId <= 0)
                {
                    NextcloudTalkAddIn.LogTalkMessage("Delegation failed: attendeeId not found (delegate=" + delegateId + ").");
                    return;
                }

                string promoteError;
                if (!service.PromoteModerator(roomToken, attendeeId, out promoteError))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Delegation failed: moderator could not be assigned (" + promoteError + ").");
                    string warning = string.IsNullOrWhiteSpace(promoteError)
                        ? Strings.WarningModeratorTransferFailed
                        : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, promoteError);
                    NextcloudTalkAddIn.ShowWarningDialog(warning);
                    return;
                }

                bool left = service.LeaveRoom(roomToken);
                NextcloudTalkAddIn.LogTalkMessage("Delegation completed (token=" + roomToken + ", delegate=" + delegateId + ", leftSelf=" + left + ").");

                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegated, Outlook.OlUserPropertyType.olYesNo, true);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated, Outlook.OlUserPropertyType.olText, "TRUE");
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
            }
            catch (TalkServiceException ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Delegation failed: " + ex.Message);
                string message = ex.Message;
                NextcloudTalkAddIn.ShowWarningDialog(string.IsNullOrWhiteSpace(message)
                    ? Strings.WarningModeratorTransferFailed
                    : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, message));
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error during delegation: " + ex.Message);
                string message = ex.Message;
                NextcloudTalkAddIn.ShowWarningDialog(string.IsNullOrWhiteSpace(message)
                    ? Strings.WarningModeratorTransferFailed
                    : string.Format(Strings.WarningModeratorTransferFailedWithReasonFormat, message));
            }
        }

        internal bool TryUpdateLobby(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                NextcloudTalkAddIn.LogTalkMessage("Lobby update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            try
            {
                var service = _owner.CreateTalkService();
                long startEpoch;
                if (!TryReadRequiredIcalStartEpoch(appointment, roomToken, out startEpoch))
                {
                    return false;
                }

                DateTime startUtc;
                try
                {
                    startUtc = DateTimeOffset.FromUnixTimeSeconds(startEpoch).UtcDateTime;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Lobby update blocked: X-NCTALK-START out of range (token=" + roomToken + ", startEpoch=" + startEpoch.ToString(CultureInfo.InvariantCulture) + ").", ex);
                    return false;
                }

                DateTime end;
                try
                {
                    end = appointment.End;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Appointment.End during lobby update (token=" + roomToken + ").", ex);
                    return false;
                }

                if (end == DateTime.MinValue)
                {
                    end = startUtc;
                }

                NextcloudTalkAddIn.LogTalkMessage("Updating lobby (token=" + roomToken + ", startEpoch=" + startEpoch.ToString(CultureInfo.InvariantCulture) + ", startUtc=" + startUtc.ToString("o") + ", end=" + end.ToString("o") + ", event=" + isEventConversation + ").");
                service.UpdateLobby(roomToken, startUtc, end, isEventConversation);
                NextcloudTalkAddIn.LogTalkMessage("Lobby updated successfully (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Lobby update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                NextcloudTalkAddIn.LogTalkMessage("Lobby could not be updated: " + ex.Message);
                NextcloudTalkAddIn.ShowWarningDialog(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error while updating lobby: " + ex.Message);
                NextcloudTalkAddIn.ShowWarningDialog(string.Format(Strings.WarningLobbyUpdateFailed, ex.Message));
            }

            return false;
        }

        internal bool TryUpdateRoomName(Outlook.AppointmentItem appointment, string roomToken)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                NextcloudTalkAddIn.LogTalkMessage("Room name update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            string roomName = GetNormalizedRoomName(appointment);
            if (string.IsNullOrWhiteSpace(roomName))
            {
                NextcloudTalkAddIn.LogTalkMessage("Room name update skipped (empty subject, token=" + roomToken + ").");
                return true;
            }

            try
            {
                var service = _owner.CreateTalkService();
                NextcloudTalkAddIn.LogTalkMessage("Updating room name (token=" + roomToken + ", length=" + roomName.Length + ").");
                service.UpdateRoomName(roomToken, roomName);
                NextcloudTalkAddIn.LogTalkMessage("Room name updated (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Room name update skipped for event conversation (token=" + roomToken + ").");
                    PersistEventConversationTraits(appointment, roomToken);
                    return true;
                }

                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Room name update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                NextcloudTalkAddIn.LogTalkMessage("Room name could not be updated: " + ex.Message);
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error while updating room name: " + ex.Message);
            }

            return false;
        }

        internal bool TryUpdateRoomDescription(Outlook.AppointmentItem appointment, string roomToken, bool isEventConversation)
        {
            if (string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string delegateId;
            if (IsDelegatedToOtherUser(appointment, out delegateId))
            {
                NextcloudTalkAddIn.LogTalkMessage("Description update skipped (delegation=" + delegateId + ", token=" + roomToken + ").");
                return true;
            }

            string description = BuildDescriptionPayload(appointment);
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (description == null)
            {
                description = string.Empty;
            }

            try
            {
                var service = _owner.CreateTalkService();
                NextcloudTalkAddIn.LogTalkMessage("Updating room description (token=" + roomToken + ", event=" + isEventConversation + ", textLength=" + description.Length + ").");
                service.UpdateDescription(roomToken, description, isEventConversation);
                NextcloudTalkAddIn.LogTalkMessage("Room description updated (token=" + roomToken + ").");
                return true;
            }
            catch (TalkServiceException ex)
            {
                if (IsEventConversationDescriptionError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Room description update skipped after event-conversation response (token=" + roomToken + ").");
                    PersistEventConversationTraits(appointment, roomToken);
                    return true;
                }

                if (IsMissingOrForbiddenRoomMutationError(ex))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Room description update skipped after access loss (token=" + roomToken + ", status=" + (int)ex.StatusCode + ").");
                    return true;
                }

                NextcloudTalkAddIn.LogTalkMessage("Room description could not be updated: " + ex.Message);
                NextcloudTalkAddIn.ShowWarningDialog(string.Format(Strings.WarningDescriptionUpdateFailed, ex.Message));
            }
            catch (Exception ex)
            {
                NextcloudTalkAddIn.LogTalkMessage("Unexpected error while updating room description: " + ex.Message);
                NextcloudTalkAddIn.ShowWarningDialog(string.Format(Strings.WarningDescriptionUpdateFailed, ex.Message));
            }

            return false;
        }

        private static string BuildDescriptionPayload(Outlook.AppointmentItem appointment)
        {
            // Defensive fallback: upstream callers are internal, but COM callbacks can still race.
            if (appointment == null)
            {
                return null;
            }

            string body = appointment.Body ?? string.Empty;
            return body.Trim();
        }

        internal static TalkRoomType? GetRoomType(Outlook.AppointmentItem appointment)
        {
            string eventRaw = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalEvent);
            if (!string.IsNullOrWhiteSpace(eventRaw))
            {
                string normalized = eventRaw.Trim();
                if (string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.EventConversation;
                }

                if (string.Equals(normalized, "standard", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.StandardRoom;
                }
            }

            string value = GetUserPropertyText(appointment, NextcloudTalkAddIn.PropertyRoomType);
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            TalkRoomType parsed;
            if (Enum.TryParse<TalkRoomType>(value, true, out parsed))
            {
                return parsed;
            }

            return null;
        }

        private static bool TryIsEventConversationFromObjectId(Outlook.AppointmentItem appointment)
        {
            try
            {
                string objectId = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalObjectId);
                return !string.IsNullOrWhiteSpace(objectId);
            }
            catch
            {
                return false;
            }
        }

        private void PersistEventConversationTraits(Outlook.AppointmentItem appointment, string roomToken)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return;
            }

            try
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalEvent, Outlook.OlUserPropertyType.olText, "event");
                SetUserProperty(appointment, NextcloudTalkAddIn.PropertyRoomType, Outlook.OlUserPropertyType.olText, TalkRoomType.EventConversation.ToString());
                NextcloudTalkAddIn.LogTalkMessage("Conversation type persisted as event (token=" + (roomToken ?? "n/a") + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist event conversation traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }
        }

        private void TryHydrateMissingRoomTraitsFromServer(
            Outlook.AppointmentItem appointment,
            string roomToken,
            ref bool? lobbyEnabled,
            ref bool? isEventConversation)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return;
            }

            bool needLobby = !lobbyEnabled.HasValue;
            bool needEvent = !isEventConversation.HasValue;
            if (!needLobby && !needEvent)
            {
                return;
            }

            try
            {
                bool? resolvedLobby;
                bool? resolvedEvent;
                var service = _owner.CreateTalkService();
                if (!service.TryReadRoomTraits(roomToken, out resolvedLobby, out resolvedEvent))
                {
                    return;
                }

                if (needLobby && resolvedLobby.HasValue)
                {
                    lobbyEnabled = resolvedLobby.Value;
                    try
                    {
                        SetUserProperty(appointment, NextcloudTalkAddIn.IcalLobby, Outlook.OlUserPropertyType.olText, resolvedLobby.Value ? "TRUE" : "FALSE");
                        SetUserProperty(appointment, NextcloudTalkAddIn.PropertyLobby, Outlook.OlUserPropertyType.olYesNo, resolvedLobby.Value);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist lobby flag after server bootstrap.", ex);
                    }

                    NextcloudTalkAddIn.LogTalkMessage("Lobby flag bootstrapped from server (token=" + roomToken + ", lobby=" + resolvedLobby.Value + ").");
                }

                if (needEvent && resolvedEvent.HasValue)
                {
                    isEventConversation = resolvedEvent.Value;
                    try
                    {
                        SetUserProperty(appointment, NextcloudTalkAddIn.IcalEvent, Outlook.OlUserPropertyType.olText, resolvedEvent.Value ? "event" : "standard");
                        SetUserProperty(
                            appointment,
                            NextcloudTalkAddIn.PropertyRoomType,
                            Outlook.OlUserPropertyType.olText,
                            resolvedEvent.Value ? TalkRoomType.EventConversation.ToString() : TalkRoomType.StandardRoom.ToString());
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist conversation type after server bootstrap.", ex);
                    }

                    NextcloudTalkAddIn.LogTalkMessage("Conversation type bootstrapped from server (token=" + roomToken + ", event=" + resolvedEvent.Value + ").");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to bootstrap room traits from server (token=" + roomToken + ").", ex);
            }
        }

        internal void ClearTalkProperties(Outlook.AppointmentItem appointment)
        {
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyToken);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyRoomType);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyLobby);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertySearchVisible);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyPasswordSet);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyStartEpoch);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDataVersion);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyAddUsers);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyAddGuests);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegateId);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.PropertyDelegated);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalToken);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalUrl);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalLobby);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalEvent);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalAddParticipants);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
        }

        internal static void SetUserProperty(Outlook.AppointmentItem appointment, string name, Outlook.OlUserPropertyType type, object value)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return;
            }

            try
            {
                var properties = appointment.UserProperties;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (properties == null)
                {
                    return;
                }

                var property = properties[name] ?? properties.Add(name, type, Type.Missing, Type.Missing);
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (property == null)
                {
                    return;
                }

                property.Value = value;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to set user property '" + name + "'.", ex);
            }
        }

        internal static string GetUserPropertyText(Outlook.AppointmentItem appointment, string name)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return null;
            }

            try
            {
                var property = appointment.UserProperties[name];
                return property != null ? property.Value as string : null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read user property '" + name + "'.", ex);
                return null;
            }
        }

        internal static bool HasUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return false;
            }

            try
            {
                return appointment.UserProperties[name] != null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to check user property '" + name + "'.", ex);
                return false;
            }
        }

        internal static string GetUserPropertyTextPrefer(Outlook.AppointmentItem appointment, string primaryName, string legacyName)
        {
            string primaryValue = GetUserPropertyText(appointment, primaryName);
            if (!string.IsNullOrWhiteSpace(primaryValue))
            {
                return primaryValue;
            }

            return GetUserPropertyText(appointment, legacyName);
        }

        internal static bool GetUserPropertyBoolPrefer(Outlook.AppointmentItem appointment, string primaryName, string legacyName)
        {
            if (HasUserProperty(appointment, primaryName))
            {
                return GetUserPropertyBool(appointment, primaryName);
            }

            return GetUserPropertyBool(appointment, legacyName);
        }

        internal static bool GetUserPropertyBool(Outlook.AppointmentItem appointment, string name)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return false;
            }

            try
            {
                var property = appointment.UserProperties[name];
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (property == null)
                {
                    return false;
                }

                object value = property.Value;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (value == null)
                {
                    return false;
                }

                if (value is bool)
                {
                    return (bool)value;
                }

                if (value is int)
                {
                    return (int)value != 0;
                }

                string text = value as string;
                if (!string.IsNullOrEmpty(text))
                {
                    bool boolParsed;
                    if (bool.TryParse(text, out boolParsed))
                    {
                        return boolParsed;
                    }

                    int numericParsed;
                    if (int.TryParse(text, out numericParsed))
                    {
                        return numericParsed != 0;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read boolean user property '" + name + "'.", ex);
                return false;
            }
        }

        internal static void RemoveUserProperty(Outlook.AppointmentItem appointment, string name)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return;
            }

            try
            {
                var property = appointment.UserProperties[name];
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (property != null)
                {
                    property.Delete();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to remove user property '" + name + "'.", ex);
            }
        }

        private static bool IsEventConversationDescriptionError(TalkServiceException ex)
        {
            // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
            if (ex == null || string.IsNullOrWhiteSpace(ex.Message))
            {
                return false;
            }

            string normalized = ex.Message.Trim();
            return string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase)
                   || normalized.IndexOf("event conversation", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsMissingOrForbiddenRoomMutationError(TalkServiceException ex)
        {
            // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
            if (ex == null)
            {
                return false;
            }

            return ex.StatusCode == HttpStatusCode.NotFound || ex.StatusCode == HttpStatusCode.Forbidden;
        }

        private static string GetNormalizedRoomName(Outlook.AppointmentItem appointment)
        {
            // Keep null-tolerant: caller paths run inside Outlook event handlers with COM races.
            if (appointment == null)
            {
                return string.Empty;
            }

            try
            {
                string subject = appointment.Subject;
                return string.IsNullOrWhiteSpace(subject) ? string.Empty : subject.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment subject for room name sync.", ex);
                return string.Empty;
            }
        }
    }
}
