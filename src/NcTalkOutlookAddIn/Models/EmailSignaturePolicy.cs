// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class EmailSignaturePolicy
    {
        internal bool Active { get; set; }

        internal string Reason { get; set; }

        internal string UserEmail { get; set; }

        internal string TemplateHtml { get; set; }

        internal bool OnCompose { get; set; }

        internal bool OnReply { get; set; }

        internal bool OnForward { get; set; }
    }
}
