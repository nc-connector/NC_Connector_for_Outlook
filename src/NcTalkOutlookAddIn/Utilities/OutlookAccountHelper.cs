/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Gemeinsame Hilfsfunktionen fuer die Arbeit mit Outlook-Accounts (Identifier normalisieren).
     */
    internal static class OutlookAccountHelper
    {
        internal static string NormalizeAccountIdentifier(Outlook.Account account)
        {
            if (account == null)
            {
                return string.Empty;
            }

            string candidate = TryGetAccountProperty(() => account.SmtpAddress);
            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = TryGetAccountProperty(() => account.UserName);
            }

            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = TryGetAccountProperty(() => account.DisplayName);
            }

            return NormalizeAccountIdentifier(candidate);
        }

        internal static string NormalizeAccountIdentifier(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value.Trim().ToLowerInvariant();
        }

        private static string TryGetAccountProperty(Func<string> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
