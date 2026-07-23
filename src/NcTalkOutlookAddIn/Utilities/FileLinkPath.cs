// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace NcTalkOutlookAddIn.Utilities
{
    // Defines the local-to-DAV naming rules shared by the wizard and upload pipeline.
    internal static class FileLinkPath
    {
        private const string ShareFolderDateFormat = "yyyyMMdd";

        internal static string BuildShareFolderName(
            DateTime shareDate,
            string sanitizedShareName)
        {
            return shareDate.ToString(
                       ShareFolderDateFormat,
                       CultureInfo.InvariantCulture)
                   + "_"
                   + (sanitizedShareName ?? string.Empty);
        }

        internal static string NormalizeRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            return string.Join(
                "/",
                path.Split(
                        new[] { '/', '\\' },
                        StringSplitOptions.RemoveEmptyEntries)
                    .Select(SanitizeComponent)
                    .Where(
                        component =>
                            !string.IsNullOrWhiteSpace(component)));
        }

        internal static string Combine(
            string basePath,
            string component)
        {
            if (string.IsNullOrEmpty(basePath))
            {
                return component ?? string.Empty;
            }
            if (string.IsNullOrEmpty(component))
            {
                return basePath;
            }

            return basePath.TrimEnd('/')
                   + "/"
                   + component.TrimStart('/');
        }

        internal static string SanitizeComponent(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            char[] invalid = Path.GetInvalidFileNameChars();
            var builder = new StringBuilder(value.Length);
            foreach (char character in value)
            {
                builder.Append(
                    invalid.Contains(character)
                        ? '_'
                        : character);
            }
            string sanitized = builder.ToString().Trim();
            if (string.Equals(
                    sanitized,
                    ".",
                    StringComparison.Ordinal)
                || string.Equals(
                    sanitized,
                    "..",
                    StringComparison.Ordinal))
            {
                return sanitized.Replace('.', '_');
            }
            return sanitized;
        }

        internal static int GetDepth(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return 0;
            }

            return path.Count(character => character == '/') + 1;
        }

        internal static string GetParent(string path)
        {
            string normalized = NormalizeRelativePath(path);
            int separator = normalized.LastIndexOf('/');
            return separator < 0
                ? string.Empty
                : normalized.Substring(0, separator);
        }
    }
}
