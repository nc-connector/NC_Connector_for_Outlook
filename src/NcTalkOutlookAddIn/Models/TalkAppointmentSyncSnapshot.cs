// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Services;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class TalkAppointmentSyncSnapshot
    {
        internal TalkAppointmentSyncSnapshot()
        {
            RoomToken = string.Empty;
            RoomUrl = string.Empty;
            RoomName = string.Empty;
            Description = string.Empty;
            DelegateUserId = string.Empty;
            AttendeeEmails = new List<string>();
            ProfileScope = string.Empty;
            DataDirectory = string.Empty;
        }

        internal string RoomToken { get; set; }

        internal string RoomUrl { get; set; }

        internal bool LobbyOnly { get; set; }

        internal bool IsEventConversation { get; set; }

        internal bool LobbyKnown { get; set; }

        internal bool LobbyEnabled { get; set; }

        internal bool UpdateLobby { get; set; }

        internal long StartEpoch { get; set; }

        internal DateTime End { get; set; }

        internal string RoomName { get; set; }

        internal string Description { get; set; }

        internal bool AddUsers { get; set; }

        internal bool AddGuests { get; set; }

        internal List<string> AttendeeEmails { get; set; }

        internal bool DelegationPending { get; set; }

        internal string DelegateUserId { get; set; }

        internal int CacheHours { get; set; }

        internal string ProfileScope { get; set; }

        internal string DataDirectory { get; set; }

        internal TalkServiceConfiguration Configuration { get; set; }

        internal TalkAppointmentSyncSnapshot WithLobbyUpdate(
            TalkAppointmentSyncSnapshot lobby)
        {
            var updated = (TalkAppointmentSyncSnapshot)MemberwiseClone();
            updated.IsEventConversation = lobby.IsEventConversation;
            updated.LobbyKnown = lobby.LobbyKnown;
            updated.LobbyEnabled = lobby.LobbyEnabled;
            updated.UpdateLobby = lobby.UpdateLobby;
            updated.StartEpoch = lobby.StartEpoch;
            updated.End = lobby.End;
            return updated;
        }
    }

    internal sealed class TalkAppointmentSyncResult
    {
        internal TalkAppointmentSyncResult(string roomToken)
        {
            RoomToken = roomToken ?? string.Empty;
            Warnings = new List<TalkAppointmentSyncWarning>();
        }

        internal string RoomToken { get; private set; }

        internal long? AppliedLobbyEpoch { get; set; }

        internal bool MarkEventConversation { get; set; }

        internal bool DelegationApplied { get; set; }

        internal bool ClearDelegation { get; set; }

        internal TalkServiceConfiguration HandoffConfiguration { get; set; }

        internal List<TalkAppointmentSyncWarning> Warnings { get; private set; }
    }

    internal sealed class TalkAppointmentSyncWarning
    {
        internal TalkAppointmentSyncWarning(
            TalkAppointmentSyncWarningKind kind,
            string detail)
        {
            Kind = kind;
            Detail = detail ?? string.Empty;
        }

        internal TalkAppointmentSyncWarningKind Kind { get; private set; }

        internal string Detail { get; private set; }
    }

    internal enum TalkAppointmentSyncWarningKind
    {
        Lobby,
        Description,
        Delegation
    }
}
