// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
namespace NcTalkOutlookAddIn.Services
{
    // Carries precomputed paths/metadata from upload setup into transfer.
    // Avoids duplicate directory creation and keeps user/path parameters centralized.
    internal sealed class FileLinkUploadContext
    {
        internal FileLinkUploadContext(
            string normalizedBaseUrl,
            string userId,
            string sanitizedShareName,
            string folderName,
            string relativeFolderPath,
            FileLinkUploadPlan plan,
            FileLinkUploadContext previousContext = null)
        {
            if (string.IsNullOrWhiteSpace(normalizedBaseUrl))
            {
                throw new ArgumentNullException("normalizedBaseUrl");
            }
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            NormalizedBaseUrl = normalizedBaseUrl;
            UserId = userId ?? string.Empty;
            SanitizedShareName = sanitizedShareName ?? string.Empty;
            FolderName = folderName ?? string.Empty;
            RelativeFolderPath = relativeFolderPath ?? string.Empty;
            Plan = plan;
            ReservedFiles = new HashSet<string>(
                previousContext == null
                    ? Enumerable.Empty<string>()
                    : previousContext.ReservedFiles,
                StringComparer.OrdinalIgnoreCase);
            ReservedDirectories = new HashSet<string>(
                previousContext == null
                    ? Enumerable.Empty<string>()
                    : previousContext.ReservedDirectories,
                StringComparer.OrdinalIgnoreCase);
            ReservedDirectories.UnionWith(plan.DirectoriesToCreate);
            foreach (FileLinkPlannedFile file in plan.Files)
            {
                ReservedFiles.Add(file.RemotePath);
                string path = file.RemotePath;
                int separator;
                while ((separator = path.LastIndexOf('/')) >= 0)
                {
                    path = path.Substring(0, separator);
                    ReservedDirectories.Add(path);
                }
            }
            IsAppend = previousContext != null;
        }

        internal string NormalizedBaseUrl { get; private set; }

        internal string UserId { get; private set; }

        internal string SanitizedShareName { get; private set; }

        internal string FolderName { get; private set; }

        internal string RelativeFolderPath { get; private set; }

        internal FileLinkUploadPlan Plan { get; private set; }

        internal ISet<string> ReservedFiles { get; private set; }

        internal ISet<string> ReservedDirectories { get; private set; }

        internal bool IsAppend { get; private set; }
    }
}
