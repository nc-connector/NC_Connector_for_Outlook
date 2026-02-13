/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Result returned by the sharing service with all data required for the HTML block and follow-up actions.
     */
    internal sealed class FileLinkResult
    {
        internal FileLinkResult(
            string shareUrl,
            string shareToken,
            string password,
            DateTime? expireDate,
            FileLinkPermissionFlags permissions,
            string folderName,
            string relativePath)
        {
            ShareUrl = shareUrl;
            ShareToken = shareToken;
            Password = password;
            ExpireDate = expireDate;
            Permissions = permissions;
            FolderName = folderName;
            RelativePath = relativePath;
        }

        internal string ShareUrl { get; private set; }

        internal string ShareToken { get; private set; }

        internal string Password { get; private set; }

        internal DateTime? ExpireDate { get; private set; }

        internal FileLinkPermissionFlags Permissions { get; private set; }

        internal string FolderName { get; private set; }

        internal string RelativePath { get; private set; }
    }
}
