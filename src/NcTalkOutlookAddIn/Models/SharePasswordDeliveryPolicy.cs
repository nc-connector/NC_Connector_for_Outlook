// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class SharePasswordDeliveryPolicy
    {
        internal const int DefaultSecretsExpireDays = 7;
        internal const int MinSecretsExpireDays = 1;
        internal const int MaxSecretsExpireDays = 365;

        private const string Domain = "share";
        private const string ModeKey = "share_send_password_mode";
        private const string SecretsExpireDaysKey = "share_secrets_expire_days";

        private SharePasswordDeliveryPolicy(SharePasswordDeliveryMode mode, int secretsExpireDays)
        {
            Mode = mode;
            SecretsExpireDays = secretsExpireDays;
        }

        internal SharePasswordDeliveryMode Mode { get; private set; }

        internal int SecretsExpireDays { get; private set; }

        internal bool UseSecrets
        {
            get { return Mode == SharePasswordDeliveryMode.Secrets; }
        }

        internal static SharePasswordDeliveryPolicy Resolve(BackendPolicyStatus status)
        {
            return Resolve(status, SharePasswordDeliveryMode.Plain);
        }

        internal static SharePasswordDeliveryPolicy Resolve(BackendPolicyStatus status, SharePasswordDeliveryMode localMode)
        {
            int expireDays = DefaultSecretsExpireDays;
            if (status != null && status.TryGetPolicyInt(Domain, SecretsExpireDaysKey, out expireDays))
            {
                expireDays = ClampSecretsExpireDays(expireDays);
            }
            else
            {
                expireDays = DefaultSecretsExpireDays;
            }

            SharePasswordDeliveryMode mode = localMode;
            if (status != null && status.IsLocked(Domain, ModeKey))
            {
                mode = ParseMode(status.GetPolicyString(Domain, ModeKey));
            }

            return new SharePasswordDeliveryPolicy(mode, expireDays);
        }

        internal static SharePasswordDeliveryMode ParseMode(string value)
        {
            return string.Equals((value ?? string.Empty).Trim(), "secrets", StringComparison.OrdinalIgnoreCase)
                ? SharePasswordDeliveryMode.Secrets
                : SharePasswordDeliveryMode.Plain;
        }

        internal static string ToStorageValue(SharePasswordDeliveryMode mode)
        {
            return mode == SharePasswordDeliveryMode.Secrets ? "secrets" : "plain";
        }

        internal static int ClampSecretsExpireDays(int value)
        {
            if (value < MinSecretsExpireDays)
            {
                return MinSecretsExpireDays;
            }
            if (value > MaxSecretsExpireDays)
            {
                return MaxSecretsExpireDays;
            }
            return value;
        }
    }
}
