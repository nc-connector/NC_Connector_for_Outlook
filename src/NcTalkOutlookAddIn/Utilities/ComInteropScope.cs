/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Globalization;
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
        {            if (comObject == null || !Marshal.IsComObject(comObject))
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
        {            if (comObject == null || !Marshal.IsComObject(comObject))
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

        internal static string ResolveIdentityKey(object comObject, string category, string objectName)
        {            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return string.Empty;
            }

            IntPtr unk = IntPtr.Zero;
            try
            {
                unk = Marshal.GetIUnknownForObject(comObject);
                if (unk == IntPtr.Zero)
                {
                    return string.Empty;
                }
                return unchecked((ulong)unk.ToInt64()).ToString("X16", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    category,
                    "Failed to resolve COM identity key for " + (objectName ?? "object") + ".",
                    ex);
                return string.Empty;
            }
            finally
            {
                if (unk != IntPtr.Zero)
                {
                    try
                    {
                        Marshal.Release(unk);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            category,
                            "Failed to release COM identity pointer for " + (objectName ?? "object") + ".",
                            ex);
                    }
                }
            }
        }

        internal static bool AreSameObject(
            object first,
            object second,
            string category,
            string firstName,
            string secondName)
        {            if (first == null || second == null)
            {
                return false;
            }
            if (!Marshal.IsComObject(first) || !Marshal.IsComObject(second))
            {
                return ReferenceEquals(first, second);
            }
            string firstKey = ResolveIdentityKey(first, category, firstName);
            if (string.IsNullOrWhiteSpace(firstKey))
            {
                return false;
            }
            string secondKey = ResolveIdentityKey(second, category, secondName);
            if (string.IsNullOrWhiteSpace(secondKey))
            {
                return false;
            }
            return string.Equals(firstKey, secondKey, StringComparison.Ordinal);
        }
    }

}

