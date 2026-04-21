// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared password-policy generation helper used by Talk and Filelink dialogs.
    internal static class PasswordGenerationHelper
    {
        internal static int ResolveMinLength(PasswordPolicyInfo passwordPolicy, int defaultMinLength)
        {
            if (passwordPolicy != null && passwordPolicy.MinLength > 0)
            {
                return passwordPolicy.MinLength;
            }

            return defaultMinLength;
        }

        internal static string GenerateWithServerPolicyFallback(
            TalkServiceConfiguration configuration,
            PasswordPolicyInfo passwordPolicy,
            int minLength,
            string logCategory)
        {
            string generated = null;

            try
            {
                if (configuration != null && passwordPolicy != null && passwordPolicy.HasPolicy)
                {
                    var policyService = new PasswordPolicyService(configuration);
                    generated = policyService.GeneratePassword(passwordPolicy);
                }
            }
            catch (Exception ex)
            {
                generated = null;
                DiagnosticsLogger.LogException(logCategory, "Password generation via server policy failed; falling back to local generator.", ex);
            }

            if (string.IsNullOrWhiteSpace(generated) || generated.Trim().Length < minLength)
            {
                generated = PasswordGenerator.GenerateLocalPassword(minLength);
            }

            return generated;
        }
    }
}
