/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class SeparatePasswordDispatchEntry
    {
        internal string ShareLabel { get; set; }

        internal string ShareUrl { get; set; }

        internal string Password { get; set; }

        internal string Html { get; set; }

        internal string To { get; set; }

        internal string Cc { get; set; }

        internal string Bcc { get; set; }
    }
}
