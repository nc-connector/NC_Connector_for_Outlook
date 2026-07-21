// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal enum AttachmentLinkTarget
    {
        ZipDownload = 0,
        SharePage = 1
    }

    internal static class AttachmentLinkTargetPolicy
    {
        internal const string Domain = "share";
        internal const string Key = "attachment_link_target";

        internal static AttachmentLinkTarget Resolve(
            AttachmentLinkTarget? localValue,
            BackendPolicyStatus status)
        {
            AttachmentLinkTarget backendValue = AttachmentLinkTarget.ZipDownload;
            bool hasBackendValue = status != null
                && status.IsDomainActive(Domain)
                && status.HasPolicyKey(Domain, Key)
                && TryParse(status.GetPolicyString(Domain, Key), out backendValue);

            if (status != null && status.IsLocked(Domain, Key))
            {
                return hasBackendValue ? backendValue : AttachmentLinkTarget.ZipDownload;
            }
            if (localValue.HasValue)
            {
                return localValue.Value;
            }
            return hasBackendValue ? backendValue : AttachmentLinkTarget.ZipDownload;
        }

        internal static AttachmentLinkTarget Parse(string value)
        {
            AttachmentLinkTarget parsed;
            return TryParse(value, out parsed) ? parsed : AttachmentLinkTarget.ZipDownload;
        }

        internal static bool TryParse(string value, out AttachmentLinkTarget target)
        {
            target = AttachmentLinkTarget.ZipDownload;
            string normalized = (value ?? string.Empty).Trim();
            if (string.Equals(normalized, "zip_download", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(normalized, "share_page", StringComparison.OrdinalIgnoreCase))
            {
                target = AttachmentLinkTarget.SharePage;
                return true;
            }
            return false;
        }

        internal static string ToStorageValue(AttachmentLinkTarget target)
        {
            return target == AttachmentLinkTarget.SharePage ? "share_page" : "zip_download";
        }
    }
}
