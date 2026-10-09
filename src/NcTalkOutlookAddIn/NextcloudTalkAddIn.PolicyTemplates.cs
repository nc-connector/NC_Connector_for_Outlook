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
        private BackendPolicyStatus _backendPolicyLastCheck;
        private string _backendPolicyLastCheckKey = string.Empty;
        private DateTime _backendPolicyLastCheckedAtUtc;
        private long _backendPolicyRequestSequence;
        private long _backendPolicyStoredSequence;

        internal BackendPolicyStatus FetchBackendPolicyStatus(TalkServiceConfiguration configuration, string trigger)
        {
            long requestSequence;
            lock (_emailSignaturePolicyCacheSync)
            {
                requestSequence = ++_backendPolicyRequestSequence;
            }
            try
            {
                var service = new BackendPolicyService(configuration);
                BackendPolicyStatus status = service.FetchStatus();
                // Availability is separate from the last confirmed policy, including refusals.
                StoreBackendPolicySnapshotIfCurrent(configuration, status, trigger, requestSequence);
                lock (_emailSignaturePolicyCacheSync)
                {
                    if (requestSequence < _backendPolicyStoredSequence
                        && string.Equals(_backendPolicyLastCheckKey, BuildEmailSignaturePolicyCacheKey(configuration), StringComparison.Ordinal))
                    {
                        status = _backendPolicyLastCheck;
                    }
                }
                LogCore(
                    "Backend policy status fetched (trigger=" + (trigger ?? "n/a")
                    + ", active=" + (status != null && status.PolicyActive)
                    + ", share=" + (status != null && status.IsDomainActive("share"))
                    + ", talk=" + (status != null && status.IsDomainActive("talk"))
                    + ", emailSignature=" + (status != null && status.IsDomainActive("email_signature"))
                    + ", accessStatus=" + (status != null ? status.AccessStatus : "n/a")
                    + ", mode=" + (status != null ? status.Mode : "local")
                    + ", reason=" + (status != null ? status.Reason : "n/a")
                    + ").");
                return status;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Backend policy status fetch failed (trigger=" + (trigger ?? "n/a") + ").", ex);
                StoreBackendPolicySnapshotIfCurrent(configuration, null, trigger, requestSequence);
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
                if (_backendPolicyLastCheck != null
                    && string.Equals(_backendPolicyLastCheckKey, cacheKey, StringComparison.Ordinal)
                    && IsBackendPolicyCheckCurrent())
                {
                    BackendPolicyStatus effective = string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal)
                        ? _emailSignaturePolicyCache : null;
                    return Task.FromResult(effective ?? _backendPolicyLastCheck);
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

        internal BackendPolicyStatus FetchEnterpriseRolloutPolicyStatus(
            TalkServiceConfiguration configuration, string trigger)
        {
            // Called by background workers; never wait for HTTP from a ribbon or Send callback.
            BackendPolicyStatus fetched = FetchBackendPolicyStatus(configuration, trigger);
            BackendPolicyStatus known;
            return TryGetCachedEmailSignaturePolicyStatus(configuration, out known) ? known : fetched;
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

        internal bool TryGetCurrentBackendPolicyCheck(
            TalkServiceConfiguration configuration, out BackendPolicyStatus check)
        {
            lock (_emailSignaturePolicyCacheSync)
            {
                check = string.Equals(_backendPolicyLastCheckKey, BuildEmailSignaturePolicyCacheKey(configuration), StringComparison.Ordinal)
                    ? _backendPolicyLastCheck : null;
                return check != null && IsBackendPolicyCheckCurrent();
            }
        }

        internal void InvalidateCurrentBackendPolicyCheck(TalkServiceConfiguration configuration)
        {
            lock (_emailSignaturePolicyCacheSync)
            {
                if (string.Equals(_backendPolicyLastCheckKey, BuildEmailSignaturePolicyCacheKey(configuration), StringComparison.Ordinal))
                {
                    _backendPolicyLastCheckedAtUtc = DateTime.MinValue;
                }
            }
        }

        private bool IsBackendPolicyCheckCurrent()
        {
            TimeSpan lifetime = _backendPolicyLastCheck.FetchSucceeded
                ? EmailSignaturePolicyCacheLifetime : TimeSpan.FromSeconds(15);
            return DateTime.UtcNow - _backendPolicyLastCheckedAtUtc <= lifetime
                   || (!_backendPolicyLastCheck.FetchSucceeded
                       && _backendPolicyLastCheck.RetryAfterUtc > DateTime.UtcNow);
        }

        private async Task<BackendPolicyStatus> FetchAndCacheEmailSignaturePolicyStatusAsync(
            TalkServiceConfiguration configuration,
            string cacheKey,
            string trigger)
        {
            BackendPolicyStatus fetched = await Task.Run(
                () => FetchBackendPolicyStatus(configuration, trigger)).ConfigureAwait(false);
            BackendPolicyStatus effective;
            lock (_emailSignaturePolicyCacheSync)
            {
                effective = string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal)
                    ? _emailSignaturePolicyCache : fetched;
                if (string.Equals(_emailSignaturePolicyFetchKey, cacheKey, StringComparison.Ordinal))
                {
                    _emailSignaturePolicyFetchTask = null;
                    _emailSignaturePolicyFetchKey = string.Empty;
                }
            }
            return effective;
        }

        private BackendPolicyStatus StoreBackendPolicySnapshotIfCurrent(
            TalkServiceConfiguration configuration, BackendPolicyStatus fetched, string trigger, long requestSequence)
        {
            string cacheKey = BuildEmailSignaturePolicyCacheKey(configuration);
            lock (_emailSignaturePolicyCacheSync)
            {
                if (requestSequence < _backendPolicyStoredSequence)
                {
                    return string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal)
                        ? _emailSignaturePolicyCache : fetched;
                }
                _backendPolicyStoredSequence = requestSequence;
                _backendPolicyLastCheck = fetched ?? new BackendPolicyStatus(
                    true, false, false, "local", "check_failed", false, false,
                    string.Empty, null, null, null, null, null, null);
                _backendPolicyLastCheckKey = cacheKey;
                _backendPolicyLastCheckedAtUtc = DateTime.UtcNow;
                if (fetched != null && fetched.FetchSucceeded)
                {
                    _emailSignaturePolicyCache = fetched;
                    _emailSignaturePolicyCacheFetchedAtUtc = DateTime.UtcNow;
                    _emailSignaturePolicyCacheKey = cacheKey;
                }
                else if (_emailSignaturePolicyCache != null
                         && string.Equals(_emailSignaturePolicyCacheKey, cacheKey, StringComparison.Ordinal))
                {
                    LogCore("Email signature policy fetch failed; using last successful snapshot (trigger=" + (trigger ?? "n/a") + ").");
                    return _emailSignaturePolicyCache;
                }
            }
            return fetched;
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
