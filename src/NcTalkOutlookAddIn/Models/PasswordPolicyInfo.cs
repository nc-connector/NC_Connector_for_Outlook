/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Representation of Nextcloud password policy capabilities (min_length + generator endpoint).
     */
    internal sealed class PasswordPolicyInfo
    {
        internal PasswordPolicyInfo(bool hasPolicy, int minLength, string generateUrl)
        {
            HasPolicy = hasPolicy;
            MinLength = minLength;
            GenerateUrl = generateUrl ?? string.Empty;
        }

        internal bool HasPolicy { get; private set; }

        internal int MinLength { get; private set; }

        internal string GenerateUrl { get; private set; }
    }
}
