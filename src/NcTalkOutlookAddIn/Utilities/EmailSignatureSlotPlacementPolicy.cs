// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Utilities
{
    internal enum EmailSignatureSlotPlacementDecision
    {
        KeepExistingSlot,
        MoveToSafeInsertionPoint,
        UnsafeQuoteBoundaryOverlap
    }

    internal static class EmailSignatureSlotPlacementPolicy
    {
        internal static EmailSignatureSlotPlacementDecision Resolve(
            bool isReplyOrForward,
            int slotStart,
            int slotEnd,
            int safeInsertionPosition,
            int quoteBoundaryPosition,
            bool hasMeaningfulTextBetween)
        {
            if (isReplyOrForward)
            {
                if (slotStart <= quoteBoundaryPosition
                    && slotEnd > quoteBoundaryPosition)
                {
                    return EmailSignatureSlotPlacementDecision.UnsafeQuoteBoundaryOverlap;
                }

                if (slotStart > quoteBoundaryPosition)
                {
                    return EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint;
                }
            }

            if (safeInsertionPosition > slotEnd
                && hasMeaningfulTextBetween)
            {
                return EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint;
            }

            return EmailSignatureSlotPlacementDecision.KeepExistingSlot;
        }
    }
}
