/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Aggregates all inputs from the sharing wizard (base path, permissions, expiration,
     * optional note, and the selected files). Used as a parameter object for the service layer.
     */
    internal sealed class FileLinkRequest
    {
        internal FileLinkRequest()
        {
            Items = new List<FileLinkSelection>();
        }

        internal string BasePath { get; set; }

        internal string ShareName { get; set; }

        internal FileLinkPermissionFlags Permissions { get; set; }

        internal bool PasswordEnabled { get; set; }

        internal string Password { get; set; }

        internal bool ExpireEnabled { get; set; }

        internal DateTime? ExpireDate { get; set; }

        internal bool NoteEnabled { get; set; }

        internal string Note { get; set; }

        internal IList<FileLinkSelection> Items { get; private set; }
    }
}
