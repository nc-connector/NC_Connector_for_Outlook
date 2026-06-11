// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class SeparatePasswordDispatchEntry
    {
        internal string ShareLabel { get; set; }

        internal string ShareUrl { get; set; }

        internal string ShareId { get; set; }

        internal string ShareToken { get; set; }

        internal string RelativePath { get; set; }

        internal DateTime? ExpireDate { get; set; }

        internal FileLinkPermissionFlags Permissions { get; set; }

        internal string Password { get; set; }

        internal string Html { get; set; }

        internal string PlainText { get; set; }

        internal bool IsPlainText { get; set; }

        internal SharePasswordDeliveryMode DeliveryMode { get; set; }

        internal int SecretsExpireDays { get; set; }

        internal string LanguageOverride { get; set; }

        internal BackendPolicyStatus BackendPolicyStatus { get; set; }

        internal string To { get; set; }

        internal string Cc { get; set; }

        internal string Bcc { get; set; }

        internal string SenderEmail { get; set; }

        internal string SendUsingAccountSmtpAddress { get; set; }

        internal string SentOnBehalfOfName { get; set; }
    }
}
