// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Models
{
    // Holds the parsed server version and capability blocks used by feature services.
    internal sealed class NextcloudCapabilitiesSnapshot
    {
        internal NextcloudCapabilitiesSnapshot(
            Version serverVersion,
            string serverVersionText,
            IDictionary<string, object> capabilities,
            bool bulkUploadSupported)
        {
            ServerVersion = serverVersion;
            ServerVersionText = serverVersionText ?? string.Empty;
            Capabilities = capabilities;
            BulkUploadSupported = bulkUploadSupported;
        }

        internal Version ServerVersion { get; private set; }

        internal string ServerVersionText { get; private set; }

        internal IDictionary<string, object> Capabilities { get; private set; }

        internal bool BulkUploadSupported { get; private set; }
    }
}
