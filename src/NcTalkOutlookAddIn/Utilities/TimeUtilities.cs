/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Time conversion helpers used across services and Outlook integrations.
     */
    internal static class TimeUtilities
    {
        internal static long? ToUnixTimeSeconds(DateTime? value)
        {
            if (!value.HasValue)
            {
                return null;
            }

            DateTime date = value.Value;
            if (date == DateTime.MinValue)
            {
                return null;
            }
            if (date.Kind == DateTimeKind.Unspecified)
            {
                date = DateTime.SpecifyKind(date, DateTimeKind.Local);
            }

            DateTimeOffset offset = new DateTimeOffset(date);
            return offset.ToUnixTimeSeconds();
        }
    }
}
