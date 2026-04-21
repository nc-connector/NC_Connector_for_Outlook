// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared backend-policy helpers for UI forms.
    internal static class PolicyUiHelper
    {
        internal static bool IsPolicyActive(BackendPolicyStatus status)
        {
            return status != null && status.PolicyActive;
        }

        internal static bool HasBackendSeatEntitlement(BackendPolicyStatus status)
        {
            return status != null
                   && status.EndpointAvailable
                   && status.SeatAssigned
                   && status.IsValid
                   && string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase);
        }

        internal static string GetSeparatePasswordUnavailableTooltip(BackendPolicyStatus status)
        {
            if (status == null || !status.EndpointAvailable)
            {
                return Strings.SharingPasswordSeparateBackendRequiredTooltip;
            }

            if (!status.SeatAssigned)
            {
                return Strings.SharingPasswordSeparateNoSeatTooltip;
            }

            if (!status.IsValid || !string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.SharingPasswordSeparatePausedTooltip;
            }

            return string.Empty;
        }
    }
}
