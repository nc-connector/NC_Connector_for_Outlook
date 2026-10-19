// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;

namespace NcTalkOutlookAddIn.Models
{
    internal enum CalendarResourceAvailabilityState { Unknown, Checking, Free, Busy }

    internal sealed class CalendarBusyPeriod
    {
        internal CalendarBusyPeriod(DateTime startUtc, DateTime endUtc)
        {
            StartUtc = startUtc;
            EndUtc = endUtc;
        }

        internal DateTime StartUtc { get; private set; }
        internal DateTime EndUtc { get; private set; }
    }

    internal sealed class CalendarResourceAvailability
    {
        internal CalendarResourceAvailability(string email, DateTime startUtc, DateTime endUtc,
            bool isKnown, IEnumerable<CalendarBusyPeriod> busyPeriods)
        {
            Email = email;
            StartUtc = startUtc;
            EndUtc = endUtc;
            IsKnown = isKnown;
            BusyPeriods = new List<CalendarBusyPeriod>(busyPeriods ?? Enumerable.Empty<CalendarBusyPeriod>()).AsReadOnly();
        }

        internal string Email { get; private set; }
        internal DateTime StartUtc { get; private set; }
        internal DateTime EndUtc { get; private set; }
        internal bool IsKnown { get; private set; }
        internal IReadOnlyList<CalendarBusyPeriod> BusyPeriods { get; private set; }
        internal CalendarResourceAvailabilityState State
        {
            get { return !IsKnown ? CalendarResourceAvailabilityState.Unknown : BusyPeriods.Count > 0 ? CalendarResourceAvailabilityState.Busy : CalendarResourceAvailabilityState.Free; }
        }

        internal CalendarResourceAvailability ForInterval(DateTime startUtc, DateTime endUtc)
        {
            if (!IsKnown || startUtc < StartUtc || endUtc > EndUtc || endUtc <= startUtc)
            {
                return Unknown(Email, startUtc, endUtc);
            }
            return new CalendarResourceAvailability(Email, startUtc, endUtc, true,
                BusyPeriods.Where(period => period.StartUtc < endUtc && period.EndUtc > startUtc)
                    .Select(period => new CalendarBusyPeriod(period.StartUtc < startUtc ? startUtc : period.StartUtc,
                        period.EndUtc > endUtc ? endUtc : period.EndUtc)));
        }

        internal static CalendarResourceAvailability Unknown(string email, DateTime startUtc, DateTime endUtc)
        {
            return new CalendarResourceAvailability(email, startUtc, endUtc, false, null);
        }
    }
}
