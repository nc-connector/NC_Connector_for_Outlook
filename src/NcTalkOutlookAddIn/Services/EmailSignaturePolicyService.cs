// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class EmailSignaturePolicyService
    {
        private const string Domain = "email_signature";
        private const string KeyOnCompose = "email_signature_on_compose";
        private const string KeyOnReply = "email_signature_on_reply";
        private const string KeyOnForward = "email_signature_on_forward";
        private const string KeyTemplate = "email_signature_template";
        private const string KeyUserEmail = "user_email";

        private readonly BackendPolicyStatus _status;
        private readonly AddinSettings _settings;

        internal EmailSignaturePolicyService(BackendPolicyStatus status, AddinSettings settings)
        {
            _status = status;
            _settings = settings ?? new AddinSettings();
        }

        internal EmailSignaturePolicy Resolve()
        {
            if (_status == null || !_status.IsDomainActive(Domain))
            {
                if (_status != null
                    && _status.EndpointAvailable
                    && _status.PolicyActive
                    && !_status.IsDomainAvailable(Domain))
                {
                    return Inactive("signature_backend_unsupported");
                }
                return Inactive("policy_inactive");
            }

            bool backendOnCompose;
            if (!_status.TryGetPolicyBool(Domain, KeyOnCompose, out backendOnCompose))
            {
                return Inactive("signature_disabled_by_backend");
            }

            string templateHtml = _status.GetPolicyString(Domain, KeyTemplate);
            string userEmail = NormalizeEmail(_status.GetPolicyString(Domain, KeyUserEmail));

            AddinSettings effective = _settings.ResolvePolicyDefaults(_status);
            bool onCompose = effective.EmailSignatureOnCompose.GetValueOrDefault();
            bool onReply = effective.EmailSignatureOnReply.GetValueOrDefault();
            bool onForward = effective.EmailSignatureOnForward.GetValueOrDefault();

            return new EmailSignaturePolicy
            {
                Active = onCompose,
                Reason = onCompose ? "active" : ResolveComposeInactiveReason(backendOnCompose),
                UserEmail = userEmail,
                TemplateHtml = templateHtml.Trim(),
                OnCompose = onCompose,
                OnReply = onReply,
                OnForward = onForward
            };
        }

        internal static bool IsAvailableForConfiguration(BackendPolicyStatus status)
        {
            bool backendOnCompose;
            return status != null
                   && status.IsDomainActive(Domain)
                   && status.TryGetPolicyBool(Domain, KeyOnCompose, out backendOnCompose)
                   && !string.IsNullOrWhiteSpace(status.GetPolicyString(Domain, KeyTemplate))
                   && !string.IsNullOrWhiteSpace(status.GetPolicyString(Domain, KeyUserEmail));
        }

        internal static string NormalizeEmail(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().ToLowerInvariant();
        }

        internal static bool ResolveFlag(
            BackendPolicyStatus status,
            string key,
            bool? localValue,
            bool preferBackendDefaults = false)
        {
            bool backendValue = false;
            bool hasBackendValue = status != null && status.TryGetPolicyBool(Domain, key, out backendValue);
            if (status != null && status.IsLocked(Domain, key))
            {
                return backendValue;
            }
            if (preferBackendDefaults && hasBackendValue && status.IsDomainActive(Domain))
            {
                return backendValue;
            }
            return localValue.HasValue ? localValue.Value : backendValue;
        }

        private string ResolveComposeInactiveReason(bool backendValue)
        {
            if (_status.IsLocked(Domain, KeyOnCompose)
                || _settings.ResolveDefaultsSource(_status) == "backend"
                || (!backendValue && !_settings.EmailSignatureOnCompose.HasValue))
            {
                return "signature_disabled_by_backend";
            }
            return "signature_disabled_locally";
        }

        private static EmailSignaturePolicy Inactive(string reason)
        {
            return new EmailSignaturePolicy
            {
                Active = false,
                Reason = reason ?? "inactive",
                UserEmail = string.Empty,
                TemplateHtml = string.Empty,
                OnCompose = false,
                OnReply = false,
                OnForward = false
            };
        }
    }
}
