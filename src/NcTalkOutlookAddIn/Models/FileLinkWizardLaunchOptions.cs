// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    // Attachment automation opens the shared wizard directly at the file step.
    internal sealed class FileLinkWizardLaunchOptions
    {
        internal bool AttachmentMode { get; set; }

        internal string AttachmentTrigger { get; set; }

        internal long AttachmentTotalBytes { get; set; }

        internal int AttachmentThresholdMb { get; set; }

        internal string AttachmentLastName { get; set; }

        internal long AttachmentLastSizeBytes { get; set; }

    }
}

