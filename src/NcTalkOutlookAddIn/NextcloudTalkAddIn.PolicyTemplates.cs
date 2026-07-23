// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn
{
        // Backend policy retrieval and Talk template/language normalization helpers.
    public sealed partial class NextcloudTalkAddIn
    {
        private static readonly TimeSpan EmailSignaturePolicyCacheLifetime = TimeSpan.FromMinutes(5);
        private readonly object _emailSignaturePolicyCacheSync = new object();
        private BackendPolicyStatus _emailSignaturePolicyCache;
        private DateTime _emailSignaturePolicyCacheFetchedAtUtc;
        private string _emailSignaturePolicyCacheKey = string.Empty;
        private Task<BackendPolicyStatus> _emailSignaturePolicyFetchTask;
        private string _emailSignaturePolicyFetchKey = string.Empty;

        internal BackendPolicyStatus FetchBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            try
            {
                var service = new BackendPolicyService(configuration);
                BackendPolicyStatus status = service.FetchStatus();
                LogCore(
                    "Backend policy status fetched (trigger=" + (trigger ?? "n/a")
                    + ", active=" + (status != null && status.PolicyActive)
                    + ", share=" + (status != null && status.IsDomainActive("share"))
                    + ", talk=" + (status != null && status.IsDomainActive("talk"))
                    + ", emailSignature=" + (status != null && status.IsDomainActive("email_signature"))
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

        internal Task<BackendPolicyStatus> GetEmailSignaturePolicyStatusAsync(
            TalkServiceConfiguration configuration,
            string trigger)
        {
            string cacheKey = BuildEmailSignaturePolicyCacheKey(configuration);
            lock (_emailSignaturePolicyCacheSync)
            {
                if (_emailSignaturePolicyCache != null
                    && string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal)
                    && DateTime.UtcNow - _emailSignaturePolicyCacheFetchedAtUtc <= EmailSignaturePolicyCacheLifetime)
                {
                    LogCore("Email signature policy cache hit (trigger=" + (trigger ?? "n/a") + ").");
                    return Task.FromResult(_emailSignaturePolicyCache);
                }

                if (_emailSignaturePolicyFetchTask != null
                    && string.Equals(_emailSignaturePolicyFetchKey, cacheKey, StringComparison.Ordinal))
                {
                    LogCore("Email signature policy fetch joined (trigger=" + (trigger ?? "n/a") + ").");
                    return _emailSignaturePolicyFetchTask;
                }

                _emailSignaturePolicyFetchKey = cacheKey;
                _emailSignaturePolicyFetchTask = FetchAndCacheEmailSignaturePolicyStatusAsync(
                    configuration,
                    cacheKey,
                    trigger);
                return _emailSignaturePolicyFetchTask;
            }
        }

        internal bool TryGetCachedEmailSignaturePolicyStatus(
            TalkServiceConfiguration configuration,
            out BackendPolicyStatus status)
        {
            string cacheKey = BuildEmailSignaturePolicyCacheKey(configuration);
            lock (_emailSignaturePolicyCacheSync)
            {
                status = string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal)
                    ? _emailSignaturePolicyCache
                    : null;
                return status != null;
            }
        }

        private async Task<BackendPolicyStatus> FetchAndCacheEmailSignaturePolicyStatusAsync(
            TalkServiceConfiguration configuration,
            string cacheKey,
            string trigger)
        {
            BackendPolicyStatus fetched = await Task.Run(
                () => FetchBackendPolicyStatus(configuration, trigger)).ConfigureAwait(false);
            BackendPolicyStatus effective = fetched;
            lock (_emailSignaturePolicyCacheSync)
            {
                if (fetched != null && fetched.FetchSucceeded)
                {
                    _emailSignaturePolicyCache = fetched;
                    _emailSignaturePolicyCacheFetchedAtUtc = DateTime.UtcNow;
                    _emailSignaturePolicyCacheKey = cacheKey;
                }
                else if (_emailSignaturePolicyCache != null
                         && string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal))
                {
                    effective = _emailSignaturePolicyCache;
                    LogCore("Email signature policy fetch failed; using last successful snapshot (trigger=" + (trigger ?? "n/a") + ").");
                }

                if (string.Equals(_emailSignaturePolicyFetchKey, cacheKey, StringComparison.Ordinal))
                {
                    _emailSignaturePolicyFetchTask = null;
                    _emailSignaturePolicyFetchKey = string.Empty;
                }
            }
            return effective;
        }

        private static string BuildEmailSignaturePolicyCacheKey(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                return string.Empty;
            }
            return configuration.GetNormalizedBaseUrl()
                   + "\n"
                   + (configuration.Username ?? string.Empty).Trim()
                   + "\n"
                   + (configuration.AppPassword ?? string.Empty);
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForTalkWizard(TalkServiceConfiguration configuration)
        {
            return FetchPasswordPolicy(
                configuration,
                LogTalk,
                "Password policy could not be loaded: ");
        }

        internal PasswordPolicyInfo FetchPasswordPolicyForFileLinkWizard(TalkServiceConfiguration configuration)
        {
            return FetchPasswordPolicy(
                configuration,
                LogFileLink,
                "Sharing password policy could not be loaded: ");
        }

        private static PasswordPolicyInfo FetchPasswordPolicy(
            TalkServiceConfiguration configuration,
            Action<string> logFailure,
            string failurePrefix)
        {
            try
            {
                return new PasswordPolicyService(configuration).FetchPolicy();
            }
            catch (Exception ex)
            {
                logFailure(failurePrefix + ex.Message);
                return null;
            }
        }

        internal static string ResolveTalkDescriptionLanguage(BackendPolicyStatus policyStatus, string fallbackLanguageOverride)
        {
            if (policyStatus != null
                && policyStatus.IsDomainActive("talk")
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
            if (policyStatus == null || !policyStatus.IsDomainActive("talk"))
            {
                return string.Empty;
            }
            return policyStatus.GetPolicyString("talk", "talk_invitation_template");
        }

        internal static string ResolveTalkEventDescriptionType(BackendPolicyStatus policyStatus)
        {
            if (policyStatus != null && policyStatus.IsDomainActive("talk"))
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
