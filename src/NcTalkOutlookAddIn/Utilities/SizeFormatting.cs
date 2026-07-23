// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;

namespace NcTalkOutlookAddIn.Utilities
{
    // Central utility for formatting byte values for UI text.
    internal static class SizeFormatting
    {
        internal static string FormatMegabytes(long bytes, CultureInfo culture = null)
        {
            CultureInfo effectiveCulture = culture ?? CultureInfo.CurrentCulture;
            decimal value = Math.Max(0, bytes) / (1024m * 1024m);
            return string.Format(effectiveCulture, "{0:0.0} MB", value);
        }

        internal static string FormatBytes(long bytes, CultureInfo culture = null)
        {
            CultureInfo effectiveCulture = culture ?? CultureInfo.CurrentCulture;
            decimal value = Math.Max(0, bytes);
            string unit = "B";

            if (value >= 1024m * 1024m * 1024m)
            {
                value /= 1024m * 1024m * 1024m;
                unit = "GB";
            }
            else if (value >= 1024m * 1024m)
            {
                value /= 1024m * 1024m;
                unit = "MB";
            }
            else if (value >= 1024m)
            {
                value /= 1024m;
                unit = "KB";
            }

            string format = unit == "B" ? "{0:0} {1}" : "{0:0.0} {1}";
            return string.Format(effectiveCulture, format, value, unit);
        }

        internal static string FormatBytesPerSecond(
            long bytesPerSecond,
            CultureInfo culture = null)
        {
            return FormatBytes(bytesPerSecond, culture) + "/s";
        }
    }
}

