/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Status values exchanged between the upload worker and the UI thread.
     */
    internal enum FileLinkUploadStatus
    {
        Pending,
        Uploading,
        Completed,
        Failed
    }
}
