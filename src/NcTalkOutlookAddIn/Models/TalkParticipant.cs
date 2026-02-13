/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Minimal Talk participant (actorType/actorId + attendeeId) from the Talk API.
     */
    internal sealed class TalkParticipant
    {
        internal TalkParticipant(string actorType, string actorId, int attendeeId)
        {
            ActorType = actorType ?? string.Empty;
            ActorId = actorId ?? string.Empty;
            AttendeeId = attendeeId;
        }

        internal string ActorType { get; private set; }

        internal string ActorId { get; private set; }

        internal int AttendeeId { get; private set; }
    }
}
