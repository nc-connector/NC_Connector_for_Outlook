/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Text;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Small helper for HTTP Basic authentication header construction.
     */
    internal static class HttpAuthUtilities
    {
        internal static string BuildBasicAuthHeader(string username, string password)
        {
            string raw = (username ?? string.Empty) + ":" + (password ?? string.Empty);
            return "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(raw));
        }
    }
}

