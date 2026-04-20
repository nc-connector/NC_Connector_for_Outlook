/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Runtime.InteropServices;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Centralizes COM release/final-release behavior and exception-safe logging.
     * This keeps cleanup paths consistent and reduces duplicated Marshal handling.
     */
    internal static class ComInteropScope
    {
        internal static void TryRelease(object comObject, string category, string failureMessage)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(category, failureMessage, ex);
            }
        }

        internal static void TryFinalRelease(object comObject, string category, string failureMessage)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(category, failureMessage, ex);
            }
        }
    }

}
