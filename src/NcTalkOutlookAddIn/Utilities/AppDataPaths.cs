/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.IO;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Centralized local app-data paths used by NC4OL runtime components.
     */
    internal static class AppDataPaths
    {
        private const string Nc4olFolderName = "NC4OL";

        internal static string GetLocalRootDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                Nc4olFolderName);
        }

        internal static string EnsureLocalRootDirectory()
        {
            string directory = GetLocalRootDirectory();
            Directory.CreateDirectory(directory);
            return directory;
        }
    }
}
