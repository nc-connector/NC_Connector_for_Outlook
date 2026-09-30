// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Text.RegularExpressions;

namespace NcTalkOutlookAddIn.Services
{
    // Shared recognition of the Outlook search paths written by NC Connector.
    internal static class IfbRegistryEndpoint
    {
        private static readonly Regex SearchPathPattern = new Regex(
            @"\Ahttp://(?:127\.0\.0\.1|localhost|\[::1\]):[0-9]{1,5}/nc-ifb/(?:[0-9a-f]{64}/)?freebusy/%NAME%@(?:%SERVER%|[a-z0-9][a-z0-9.-]*)\.vfb\z",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        internal static bool IsSearchPath(string value)
        {
            Uri uri;
            return !string.IsNullOrWhiteSpace(value)
                   && SearchPathPattern.IsMatch(value.Trim())
                   && Uri.TryCreate(value.Trim(), UriKind.Absolute, out uri)
                   && uri.Port > 0
                   && uri.Port <= 65535;
        }

        internal static bool TryGetConnectorUri(string value, out Uri uri)
        {
            uri = null;
            return !string.IsNullOrWhiteSpace(value)
                   && Uri.TryCreate(value.Trim(), UriKind.Absolute, out uri)
                   && uri.IsLoopback
                   && uri.AbsolutePath.StartsWith(
                       "/nc-ifb/",
                       StringComparison.OrdinalIgnoreCase);
        }
    }
}
