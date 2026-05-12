// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class PlainTextUtilities
    {
        internal static string NormalizeCrLf(string value)
        {
            return (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
        }

        internal static string NormalizeCrLfAndTrim(string value)
        {
            return NormalizeCrLf(value).Trim();
        }
    }
}
