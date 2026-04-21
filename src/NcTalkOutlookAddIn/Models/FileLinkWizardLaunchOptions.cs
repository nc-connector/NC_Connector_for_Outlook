// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
        // Optional launch context for the sharing wizard.
    // Used by compose attachment automation to open directly in attachment mode.
    internal sealed class FileLinkWizardLaunchOptions
    {
        internal FileLinkWizardLaunchOptions()
        {
            InitialSelections = new List<FileLinkSelection>();
        }

        internal bool AttachmentMode { get; set; }

        internal string AttachmentTrigger { get; set; }

        internal long AttachmentTotalBytes { get; set; }

        internal int AttachmentThresholdMb { get; set; }

        internal string AttachmentLastName { get; set; }

        internal long AttachmentLastSizeBytes { get; set; }

        internal IList<FileLinkSelection> InitialSelections { get; private set; }
    }
}

