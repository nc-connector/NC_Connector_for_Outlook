// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class SecretsLinkResult
    {
        internal SecretsLinkResult(string uuid, string shareUrl, DateTime? expires)
        {
            Uuid = uuid ?? string.Empty;
            ShareUrl = shareUrl ?? string.Empty;
            Expires = expires;
        }

        internal string Uuid { get; private set; }

        internal string ShareUrl { get; private set; }

        internal DateTime? Expires { get; private set; }
    }
}
