/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Distinguishes between a single file and a directory (including recursive upload).
     */
    internal enum FileLinkSelectionType
    {
        File,
        Directory
    }

    /**
     * Wraps the path and type of a user-selected resource.
     */
    internal sealed class FileLinkSelection
    {
        internal FileLinkSelection(FileLinkSelectionType type, string localPath)
        {
            SelectionType = type;
            LocalPath = localPath;
        }

        internal FileLinkSelectionType SelectionType { get; private set; }

        internal string LocalPath { get; private set; }
    }
}
