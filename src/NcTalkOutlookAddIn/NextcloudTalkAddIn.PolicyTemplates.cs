// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn
{
        // Backend policy retrieval and Talk template/language normalization helpers.
    public sealed partial class NextcloudTalkAddIn
    {
        internal BackendPolicyStatus FetchBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            try
            {
                var service = new BackendPolicyService(configuration);
                BackendPolicyStatus status = service.FetchStatus();
                LogCore(
                    "Backend policy status fetched (trigger=" + (trigger ?? "n/a")
                    + ", active=" + (status != null && status.PolicyActive)
                    + ", warningVisible=" + (status != null && status.WarningVisible)
                    + ", mode=" + (status != null ? status.Mode : "local")
                    + ", reason=" + (status != null ? status.Reason : "n/a")
                    + ").");
                return status;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Backend policy status fetch failed (trigger=" + (trigger ?? "n/a") + ").", ex);
                return null;
            }
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForTalkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogTalk("Password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForFileLinkWizard(TalkServiceConfiguration configuration)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                LogFileLink("Sharing password policy could not be loaded: " + ex.Message);
                return null;
            }
        }

        internal static string ResolveTalkDescriptionLanguage(BackendPolicyStatus policyStatus, string fallbackLanguageOverride)
        {
            if (policyStatus != null
                && policyStatus.PolicyActive
                && policyStatus.IsLocked("talk", "language_talk_description"))
            {
                string policyLanguageRaw = policyStatus.GetPolicyString("talk", "language_talk_description");
                if (!string.IsNullOrWhiteSpace(policyLanguageRaw))
                {
                    return TalkDescriptionTemplateController.NormalizeTalkDescriptionLanguage(policyLanguageRaw);
                }
            }
            return TalkDescriptionTemplateController.NormalizeTalkDescriptionLanguage(fallbackLanguageOverride);
        }

        internal static string ResolveTalkInvitationTemplate(BackendPolicyStatus policyStatus)
        {
            // Guard against null/inactive backend policy state.
            if (policyStatus == null || !policyStatus.PolicyActive)
            {
                return string.Empty;
            }
            return policyStatus.GetPolicyString("talk", "talk_invitation_template");
        }

        internal static string ResolveTalkEventDescriptionType(BackendPolicyStatus policyStatus)
        {
            if (policyStatus != null && policyStatus.PolicyActive)
            {
                string policyTypeRaw = policyStatus.GetPolicyString("talk", "event_description_type");
                if (!string.IsNullOrWhiteSpace(policyTypeRaw))
                {
                    return NormalizeTalkEventDescriptionType(policyTypeRaw);
                }
            }
            return "plain_text";
        }

        internal static string NormalizeTalkEventDescriptionType(string descriptionType)
        {
            return string.Equals(descriptionType, "html", StringComparison.OrdinalIgnoreCase)
                ? "html"
                : "plain_text";
        }

    }
}
