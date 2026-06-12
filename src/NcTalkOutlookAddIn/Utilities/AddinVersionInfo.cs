// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Diagnostics;
using System.Reflection;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class AddinVersionInfo
    {
        internal static string GetVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                if (version != null)
                {
                    return version.ToString();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read assembly version.", ex);
            }
            try
            {
                var info = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(info.ProductVersion))
                {
                    return info.ProductVersion;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read file version info.", ex);
            }
            return string.Empty;
        }
    }
}
