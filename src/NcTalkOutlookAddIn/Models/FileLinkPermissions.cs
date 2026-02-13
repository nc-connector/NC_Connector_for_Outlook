/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Bitmask of Nextcloud share permissions so UI and services use the same flags.
     */
    [Flags]
    internal enum FileLinkPermissionFlags
    {
        None = 0,
        Read = 1,
        Write = 2,
        Create = 4,
        Delete = 8
    }
}
