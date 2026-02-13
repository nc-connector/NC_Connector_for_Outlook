/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Helper functions for parsing Nextcloud version strings.
     */
    internal static class NextcloudVersionHelper
    {
        internal static bool TryParse(string value, out Version version)
        {
            version = null;

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string candidate = value.Trim();

            int spaceIndex = candidate.IndexOf(' ');
            if (spaceIndex > 0)
            {
                candidate = candidate.Substring(0, spaceIndex);
            }

            int dashIndex = candidate.IndexOf('-');
            if (dashIndex > 0)
            {
                candidate = candidate.Substring(0, dashIndex);
            }

            candidate = candidate.Trim();

            Version parsed;
            if (Version.TryParse(candidate, out parsed))
            {
                version = parsed;
                return true;
            }

            return false;
        }
    }
}
