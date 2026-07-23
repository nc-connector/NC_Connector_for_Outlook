// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared backend-policy helpers for UI forms.
    internal static class PolicyUiHelper
    {
        internal static void InitializePolicyWarningPanel(
            Panel panel,
            Label titleLabel,
            Label textLabel,
            LinkLabel linkLabel)
        {
            panel.Visible = false;
            panel.BackColor = Color.FromArgb(20, 176, 0, 32);
            panel.Paint += (sender, args) =>
            {
                ControlPaint.DrawBorder(
                    args.Graphics,
                    panel.ClientRectangle,
                    Color.FromArgb(176, 0, 32),
                    ButtonBorderStyle.Solid);
            };

            titleLabel.AutoSize = true;
            titleLabel.ForeColor = Color.FromArgb(176, 0, 32);
            titleLabel.Font = new Font(titleLabel.Font, FontStyle.Bold);
            titleLabel.Text = "\u26a0 " + Strings.PolicyWarningTitle;
            panel.Controls.Add(titleLabel);

            textLabel.AutoSize = true;
            textLabel.Text = string.Empty;
            panel.Controls.Add(textLabel);

            linkLabel.AutoSize = true;
            linkLabel.Text = Strings.PolicyWarningAdminLinkLabel;
            linkLabel.LinkColor = Color.FromArgb(0, 130, 201);
            linkLabel.ActiveLinkColor = Color.FromArgb(0, 102, 153);
            linkLabel.VisitedLinkColor = Color.FromArgb(0, 130, 201);
            panel.Controls.Add(linkLabel);
        }

        internal static bool ApplyPolicyWarningState(
            BackendPolicyStatus status,
            Panel panel,
            Label textLabel)
        {
            bool visible = status != null
                           && status.WarningVisible
                           && !string.IsNullOrWhiteSpace(status.WarningMessage);
            panel.Visible = visible;
            textLabel.Text = visible ? status.WarningMessage : string.Empty;
            return visible;
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

            if (!status.SeatAssigned)
            {
                return Strings.SharingPasswordSeparateNoSeatTooltip;
            }

            if (!status.IsValid || !string.Equals(status.SeatState, "active", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.SharingPasswordSeparatePausedTooltip;
            }

            return string.Empty;
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
