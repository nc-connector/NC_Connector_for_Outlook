// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;

namespace NcTalkOutlookAddIn.Utilities
{
    // Parses display strings and structured OCS capability versions.
    internal static class NextcloudVersionHelper
    {
        internal const int MinimumSupportedMajorVersion = 32;

        internal static bool TryParse(string value, out Version version)
        {
            version = null;

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            string candidate = value.Trim();
            for (int start = 0; start < candidate.Length; start++)
            {
                if (!char.IsDigit(candidate[start]))
                {
                    continue;
                }

                int end = start;
                while (end < candidate.Length
                       && (char.IsDigit(candidate[end])
                           || candidate[end] == '.'))
                {
                    end++;
                }

                string versionCandidate = candidate
                    .Substring(start, end - start)
                    .TrimEnd('.');
                if (versionCandidate.IndexOf('.') >= 0)
                {
                    Version parsed;
                    if (Version.TryParse(versionCandidate, out parsed))
                    {
                        version = parsed;
                        return true;
                    }
                }

                start = end - 1;
            }

            int major;
            if (int.TryParse(
                candidate,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out major)
                && major >= 0)
            {
                version = new Version(major, 0);
                return true;
            }
            return false;
        }

        internal static bool TryExtractFromCapabilities(
            IDictionary<string, object> root,
            out Version version,
            out string versionText)
        {
            version = null;
            versionText = string.Empty;

            IDictionary<string, object> data = NcJson.GetOcsData(root);
            if (data == null)
            {
                return false;
            }

            IDictionary<string, object> versionData = NcJson.GetDictionary(data, "version");
            string candidate = NcJson.GetTrimmedString(versionData, "string");
            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = BuildVersionFromParts(versionData);
            }
            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = NcJson.GetTrimmedString(data, "versionstring");
            }
            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = NcJson.GetTrimmedString(data, "version");
            }
            if (!TryParse(candidate, out version))
            {
                return false;
            }

            versionText = version.ToString();
            return true;
        }

        internal static bool IsSupported(Version version)
        {
            return version != null && version.Major >= MinimumSupportedMajorVersion;
        }

        private static string BuildVersionFromParts(IDictionary<string, object> versionData)
        {
            if (versionData == null)
            {
                return null;
            }

            int major;
            if (!NcJson.TryGetInt(versionData, "major", out major) || major < 0)
            {
                return null;
            }

            int minor;
            int micro;
            bool hasMinor = NcJson.TryGetInt(versionData, "minor", out minor) && minor >= 0;
            bool hasMicro = NcJson.TryGetInt(versionData, "micro", out micro) && micro >= 0;

            string result = major.ToString(CultureInfo.InvariantCulture);
            if (hasMinor)
            {
                result += "." + minor.ToString(CultureInfo.InvariantCulture);
                if (hasMicro)
                {
                    result += "." + micro.ToString(CultureInfo.InvariantCulture);
                }
            }
            return result;
        }
    }
}
