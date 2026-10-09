// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class TalkRoomLifecycleRecord
    {
        public TalkRoomLifecycleRecord()
        {
            Id = Guid.NewGuid().ToString("N");
            RoomToken = string.Empty;
            ServerBaseUrl = string.Empty;
            AccountLogin = string.Empty;
            AccountId = string.Empty;
            NextAttemptUtc = DateTime.MinValue;
        }

        public string Id { get; set; }

        public string RoomToken { get; set; }

        public bool IsEventConversation { get; set; }

        public string ServerBaseUrl { get; set; }

        public string AccountLogin { get; set; }

        public string AccountId { get; set; }

        public bool PendingDeletion { get; set; }

        public bool PolicyRequired { get; set; }

        public bool ConnectionPaused { get; set; }

        public int AttemptCount { get; set; }

        public DateTime NextAttemptUtc { get; set; }
    }
}
