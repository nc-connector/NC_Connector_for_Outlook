/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Services
{
    /**
     * Carries precomputed paths/metadata from the upload setup into the actual upload process.
     * Avoids duplicate directory creation and keeps user/path parameters centralized.
     */
    internal sealed class FileLinkUploadContext
    {
        internal FileLinkUploadContext(
            string normalizedBaseUrl,
            string username,
            string sanitizedShareName,
            string folderName,
            string relativeFolderPath)
        {
            if (string.IsNullOrWhiteSpace(normalizedBaseUrl))
            {
                throw new ArgumentNullException("normalizedBaseUrl");
            }

            NormalizedBaseUrl = normalizedBaseUrl;
            Username = username ?? string.Empty;
            SanitizedShareName = sanitizedShareName ?? string.Empty;
            FolderName = folderName ?? string.Empty;
            RelativeFolderPath = relativeFolderPath ?? string.Empty;
            KnownFilePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            KnownFolderPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        internal string NormalizedBaseUrl { get; private set; }

        internal string Username { get; private set; }

        internal string SanitizedShareName { get; private set; }

        internal string FolderName { get; private set; }

        internal string RelativeFolderPath { get; private set; }

        internal HashSet<string> KnownFilePaths { get; private set; }

        internal HashSet<string> KnownFolderPaths { get; private set; }
    }
}
