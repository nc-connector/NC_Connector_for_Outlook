// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared backend-policy helpers for UI forms.
    internal static class PolicyUiHelper
    {
        internal static string GetEnterpriseRolloutNotice(AddinSettings settings, BackendPolicyStatus status)
        {
            if (settings == null || !settings.IsEnterpriseRollout)
            {
                return string.Empty;
            }
            if (status == null || !status.FetchSucceeded)
            {
                return Strings.EnterpriseRolloutStatusUnavailable;
            }
            if (!status.EndpointAvailable)
            {
                return Strings.EnterpriseRolloutBackendRequired;
            }
            return HasBackendSeatEntitlement(status)
                ? string.Empty
                : Strings.EnterpriseRolloutSeatRequired;
        }

        internal static bool ApplyPolicyWarningState(
            BackendPolicyStatus status,
            Panel panel,
            Label textLabel,
            Label titleLabel,
            LinkLabel linkLabel,
            string baseUrl)
        {
            string message = GetPolicyWarningMessage(status);
            bool visible = !string.IsNullOrWhiteSpace(message);
            panel.Visible = visible;
            textLabel.Text = message;
            Color noticeColor = status != null && status.IsValid
                ? Color.FromArgb(156, 108, 0)
                : Color.FromArgb(176, 0, 32);
            titleLabel.ForeColor = noticeColor;
            panel.BackColor = Color.FromArgb(20, noticeColor);
            string adminUrl = GetLicenseAdminUrl(status, baseUrl);
            linkLabel.Tag = adminUrl;
            linkLabel.Visible = !string.IsNullOrEmpty(adminUrl);
            panel.Invalidate();
            return visible;
        }

        internal static string GetPolicyWarningMessage(BackendPolicyStatus status)
        {
            if (status == null || !status.EndpointAvailable)
            {
                return string.Empty;
            }
            if (!status.FetchSucceeded)
            {
                return Strings.PolicyWarningBackendUnavailable;
            }
            if (!status.SeatAssigned && !status.CanManageLicense)
            {
                return Strings.PolicyWarningNoSeat;
            }

            string licenseNotice = GetLicenseNotice(status);
            if (!string.IsNullOrEmpty(licenseNotice))
            {
                return status.SeatAssigned ? licenseNotice : AppendNotice(licenseNotice, Strings.PolicyWarningNoSeat);
            }
            if (!status.SeatAssigned)
            {
                return Strings.PolicyWarningNoSeat;
            }
            return GetSeatNotice(status);
        }

        private static string GetSeatNotice(BackendPolicyStatus status)
        {
            return string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : WithRoleHint(status,
                    string.Equals(status.SeatState, "suspended_overlimit", StringComparison.OrdinalIgnoreCase)
                        ? Strings.PolicyWarningSeatSuspended
                        : Strings.PolicyWarningSeatUnavailable);
        }

        private static string GetLicenseNotice(BackendPolicyStatus status)
        {
            if (status == null || !status.EndpointAvailable || !status.FetchSucceeded
                || (!status.SeatAssigned && !status.CanManageLicense))
            {
                return string.Empty;
            }
            // Grace preserves license validity, not a suspended user's seat access.
            if (status.IsValid && status.SeatAssigned
                && !string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            // The backend decides access. These fields explain its decision only.
            string accessStatus = string.IsNullOrWhiteSpace(status.AccessStatus)
                ? status.LicenseStatus
                : status.AccessStatus;
            string message = string.Empty;
            if (!status.IsValid)
            {
                switch (accessStatus.ToUpperInvariant())
                {
                    case "EXPIRED": message = Strings.PolicyLicenseExpired; break;
                    case "INACTIVE": message = Strings.PolicyLicenseInactive; break;
                    case "INVALID": message = Strings.PolicyLicenseInvalid; break;
                    case "OFFLINE_EXPIRED": message = Strings.PolicyLicenseOfflineExpired; break;
                    case "ACTIVATION_REQUIRED":
                        message = string.Equals(status.LicenseActivationState, "conflict", StringComparison.OrdinalIgnoreCase)
                            ? Strings.PolicyLicenseActivationConflict
                            : Strings.PolicyLicenseActivationRequired;
                        break;
                    default: message = Strings.PolicyWarningLicenseInvalid; break;
                }
            }
            else if (string.Equals(accessStatus, "GRACE", StringComparison.OrdinalIgnoreCase))
            {
                message = FormatStatusDate(Strings.PolicyLicenseGraceFormat, status.GraceUntilIso);
                if (string.IsNullOrEmpty(message))
                {
                    message = Strings.PolicyLicenseGrace;
                }
            }

            if (status.LicenseConnectionError)
            {
                message = AppendNotice(message, Strings.PolicyLicenseConnectionError);
            }
            if (status.LicenseConnectionError
                || (!status.IsValid && string.Equals(accessStatus, "OFFLINE_EXPIRED", StringComparison.OrdinalIgnoreCase)))
            {
                message = AppendNotice(message, FormatStatusDate(Strings.PolicyLicenseLastSyncFormat, status.LicenseLastSyncAtIso));
                message = AppendNotice(message, FormatStatusDate(Strings.PolicyLicenseOfflineUntilFormat, status.LicenseOfflineUntilIso));
            }
            return string.IsNullOrEmpty(message) ? string.Empty : WithRoleHint(status, message);
        }

        private static string WithRoleHint(BackendPolicyStatus status, string message)
        {
            return AppendNotice(message, status.CanManageLicense
                ? Strings.PolicyLicenseAdminHint
                : Strings.PolicyLicenseUserHint);
        }

        private static string AppendNotice(string message, string extra)
        {
            if (string.IsNullOrEmpty(extra))
            {
                return message;
            }
            return string.IsNullOrEmpty(message) ? extra : message + Environment.NewLine + extra;
        }

        private static string FormatStatusDate(string format, string isoDate)
        {
            DateTimeOffset value;
            return DateTimeOffset.TryParse(isoDate, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out value)
                ? string.Format(CultureInfo.CurrentCulture, format, value.LocalDateTime.ToString("g", CultureInfo.CurrentCulture))
                : string.Empty;
        }

        internal static string GetLicenseAdminUrl(BackendPolicyStatus status, string baseUrl)
        {
            string normalizedUrl;
            if (status == null || !status.CanManageLicense || string.IsNullOrEmpty(GetLicenseNotice(status))
                || !NextcloudUriValidator.TryNormalizeBaseUrl(baseUrl, out normalizedUrl))
            {
                return string.Empty;
            }
            return normalizedUrl + "/index.php/settings/admin/ncc_backend_4mc";
        }

        internal static void OpenLicenseAdministration(LinkLabel linkLabel, string logCategory)
        {
            string url = linkLabel.Tag as string;
            if (!string.IsNullOrEmpty(url))
            {
                BrowserLauncher.OpenUrl(url, logCategory, "Failed to open backend license settings.");
            }
        }

        internal static bool IsPolicyActive(BackendPolicyStatus status)
        {
            return status != null && status.PolicyActive;
        }

        internal static bool IsPolicyDomainAvailable(BackendPolicyStatus status, string domain)
        {
            return status != null && status.IsDomainAvailable(domain);
        }

        internal static bool IsPolicyDomainActive(BackendPolicyStatus status, string domain)
        {
            return status != null && status.IsDomainActive(domain);
        }

        internal static bool HasBackendSeatEntitlement(BackendPolicyStatus status)
        {
            return status != null
                   && status.EndpointAvailable
                   && status.SeatAssigned
                   && status.IsValid
                   && string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase);
        }

        internal static string GetSeparatePasswordUnavailableTooltip(BackendPolicyStatus status)
        {
            if (status == null || !status.EndpointAvailable)
            {
                return Strings.SharingPasswordSeparateBackendRequiredTooltip;
            }

            if (!status.FetchSucceeded)
            {
                return Strings.PolicyWarningBackendUnavailable;
            }
            if (!status.SeatAssigned)
            {
                return Strings.SharingPasswordSeparateNoSeatTooltip;
            }
            if (!status.IsValid)
            {
                return GetPolicyWarningMessage(status);
            }

            return GetSeatNotice(status);
        }

        internal static bool HasPasswordDeliveryMode(BackendPolicyStatus status)
        {
            return HasBackendSeatEntitlement(status)
                   && IsPolicyDomainActive(status, "share")
                   && status.HasPolicyKey("share", "share_send_password_mode")
                   && !string.IsNullOrWhiteSpace(status.GetPolicyString("share", "share_send_password_mode"));
        }

        internal static string GetPasswordDeliveryModeUnavailableTooltip(BackendPolicyStatus status)
        {
            string entitlementTooltip = GetSeparatePasswordUnavailableTooltip(status);
            if (!string.IsNullOrWhiteSpace(entitlementTooltip))
            {
                return entitlementTooltip;
            }

            return HasPasswordDeliveryMode(status)
                ? string.Empty
                : Strings.SharingPasswordDeliveryUnavailableTooltip;
        }
    }
}
