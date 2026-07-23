// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
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
            if (string.IsNullOrWhiteSpace(templateHtml))
            {
                return Inactive("signature_template_missing");
            }

            string userEmail = NormalizeEmail(_status.GetPolicyString(Domain, KeyUserEmail));
            if (string.IsNullOrWhiteSpace(userEmail))
            {
                return Inactive("signature_user_email_missing");
            }

            bool backendOnReply;
            bool backendOnForward;
            _status.TryGetPolicyBool(Domain, KeyOnReply, out backendOnReply);
            _status.TryGetPolicyBool(Domain, KeyOnForward, out backendOnForward);

            bool onCompose = ResolveFlag(_status, KeyOnCompose, _settings.EmailSignatureOnCompose, backendOnCompose);
            bool onReply = ResolveFlag(_status, KeyOnReply, _settings.EmailSignatureOnReply, backendOnReply);
            bool onForward = ResolveFlag(_status, KeyOnForward, _settings.EmailSignatureOnForward, backendOnForward);

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

        internal static bool ResolveFlag(BackendPolicyStatus status, string key, bool? localValue)
        {
            bool backendValue;
            if (status == null || !status.TryGetPolicyBool(Domain, key, out backendValue))
            {
                backendValue = false;
            }
            return ResolveFlag(status, key, localValue, backendValue);
        }

        private static bool ResolveFlag(
            BackendPolicyStatus status,
            string key,
            bool? localValue,
            bool backendValue)
        {
            if (status != null && status.IsLocked(Domain, key))
            {
                return backendValue;
            }
            return localValue.HasValue ? localValue.Value : backendValue;
        }

        private string ResolveComposeInactiveReason(bool backendValue)
        {
            if (_status.IsLocked(Domain, KeyOnCompose)
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
