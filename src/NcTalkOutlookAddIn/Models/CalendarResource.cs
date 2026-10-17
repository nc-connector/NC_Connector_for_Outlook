// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class CalendarResource
    {
        public string PrincipalPath { get; set; }
        public string Email { get; set; }
        public string DisplayName { get; set; }
        public string Kind { get; set; }
        public string ResourceType { get; set; }
        public string RoomType { get; set; }
        public string BuildingName { get; set; }
        public string BuildingAddress { get; set; }
        public string Floor { get; set; }
        public string RoomNumber { get; set; }
        public string Features { get; set; }
        public int? Capacity { get; set; }

        public bool IsRoom { get { return string.Equals(Kind, "ROOM", StringComparison.Ordinal); } }

        public string Location
        {
            get
            {
                var parts = new List<string>();
                foreach (string value in new[] { RoomNumber, Floor, BuildingAddress })
                {
                    if (!string.IsNullOrWhiteSpace(value)) { parts.Add(value.Trim()); }
                }
                return string.Join(", ", parts);
            }
        }

        internal CalendarResource Clone()
        {
            return (CalendarResource)MemberwiseClone();
        }
    }
}
