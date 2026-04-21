// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;

namespace NcTalkOutlookAddIn.Utilities
{
        // Central utility for formatting byte values as megabytes for UI text.
    internal static class SizeFormatting
    {
        internal static string FormatMegabytes(long bytes, CultureInfo culture = null)
        {
            CultureInfo effectiveCulture = culture ?? CultureInfo.CurrentCulture;
            decimal value = Math.Max(0, bytes) / (1024m * 1024m);
            return string.Format(effectiveCulture, "{0:0.0} MB", value);
        }
    }
}

