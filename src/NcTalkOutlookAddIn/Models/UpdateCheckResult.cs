// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class UpdateCheckResult
    {
        internal string CurrentVersion { get; set; }

        internal string LatestVersion { get; set; }

        internal bool UpdateAvailable { get; set; }

        internal string ReleaseUrl { get; set; }

        internal string DownloadUrl { get; set; }

        internal string PublishedAt { get; set; }

        internal string Message { get; set; }

        internal bool Counted { get; set; }

        internal string ChangelogTitle { get; set; }

        internal string ChangelogText { get; set; }

        internal DateTime CheckedAtUtc { get; set; }

        internal bool FromCache { get; set; }

        internal UpdateCheckResult()
        {
            CurrentVersion = string.Empty;
            LatestVersion = string.Empty;
            ReleaseUrl = string.Empty;
            DownloadUrl = string.Empty;
            PublishedAt = string.Empty;
            Message = string.Empty;
            ChangelogTitle = string.Empty;
            ChangelogText = string.Empty;
            CheckedAtUtc = DateTime.UtcNow;
        }
    }
}
