// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    // Carries the share settings selected in the wizard.
    internal sealed class FileLinkRequest
    {
        internal string BasePath { get; set; }

        internal string ShareName { get; set; }

        internal FileLinkPermissionFlags Permissions { get; set; }

        internal bool PasswordEnabled { get; set; }

        internal string Password { get; set; }

        internal bool PasswordSeparateEnabled { get; set; }

        internal SharePasswordDeliveryMode PasswordDeliveryMode { get; set; }

        internal bool ExpireEnabled { get; set; }

        internal DateTime? ExpireDate { get; set; }

        internal bool NoteEnabled { get; set; }

        internal string Note { get; set; }

        internal bool AttachmentMode { get; set; }

        internal AttachmentLinkTarget AttachmentLinkTarget { get; set; }

        internal DateTime? ShareDate { get; set; }
    }
}
