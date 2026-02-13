/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Result data of a successfully created Talk room.
     */
    internal sealed class TalkRoomCreationResult
    {
        internal string RoomToken { get; private set; }
        internal string RoomUrl { get; private set; }
        internal bool CreatedAsEventConversation { get; private set; }
        internal bool LobbyEnabled { get; private set; }
        internal bool SearchVisible { get; private set; }

        internal TalkRoomCreationResult(string token, string url, bool createdAsEvent, bool lobbyEnabled, bool searchVisible)
        {
            RoomToken = token;
            RoomUrl = url;
            CreatedAsEventConversation = createdAsEvent;
            LobbyEnabled = lobbyEnabled;
            SearchVisible = searchVisible;
        }
    }
}
