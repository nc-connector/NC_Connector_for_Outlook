/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Simple representation of a Nextcloud user from the system address book.
     * Contains the UID (actorId) and a primary email address.
     */
    internal sealed class NextcloudUser
    {
        internal NextcloudUser(string userId, string email)
        {
            UserId = userId ?? string.Empty;
            Email = email ?? string.Empty;
        }

        internal string UserId { get; private set; }

        internal string Email { get; private set; }

        internal string DisplayLabel
        {
            get
            {
                if (string.IsNullOrEmpty(Email))
                {
                    return UserId;
                }

                return UserId + " <" + Email + ">";
            }
        }

        public override string ToString()
        {
            return DisplayLabel;
        }
    }
}
