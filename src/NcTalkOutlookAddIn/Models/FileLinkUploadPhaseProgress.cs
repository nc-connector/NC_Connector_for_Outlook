// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    // Reports absolute job totals so throttled UI updates do not lose byte deltas.
    internal sealed class FileLinkUploadPhaseProgress
    {
        internal FileLinkUploadPhaseProgress(
            FileLinkUploadPhase phase,
            int completedFolders,
            int totalFolders,
            int completedFiles,
            int totalFiles,
            long uploadedBytes,
            long totalBytes)
        {
            Phase = phase;
            CompletedFolders = completedFolders;
            TotalFolders = totalFolders;
            CompletedFiles = completedFiles;
            TotalFiles = totalFiles;
            UploadedBytes = uploadedBytes;
            TotalBytes = totalBytes;
        }

        internal FileLinkUploadPhase Phase { get; private set; }

        internal int CompletedFolders { get; private set; }

        internal int TotalFolders { get; private set; }

        internal int CompletedFiles { get; private set; }

        internal int TotalFiles { get; private set; }

        internal long UploadedBytes { get; private set; }

        internal long TotalBytes { get; private set; }
    }
}
