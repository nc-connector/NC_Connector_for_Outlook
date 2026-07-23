// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    // Reports the current state of one top-level wizard selection.
    internal sealed class FileLinkUploadItemProgress
    {
        internal FileLinkUploadItemProgress(
            FileLinkSelection selection,
            long uploadedBytes,
            long totalBytes,
            FileLinkUploadStatus status)
        {
            Selection = selection;
            UploadedBytes = uploadedBytes;
            TotalBytes = totalBytes;
            Status = status;
        }

        internal FileLinkSelection Selection { get; private set; }

        internal long UploadedBytes { get; private set; }

        internal long TotalBytes { get; private set; }

        internal FileLinkUploadStatus Status { get; private set; }
    }
}
