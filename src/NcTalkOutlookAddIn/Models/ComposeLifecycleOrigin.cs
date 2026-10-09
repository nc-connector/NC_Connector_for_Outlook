// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using NcTalkOutlookAddIn.Services;

namespace NcTalkOutlookAddIn.Models
{
    [ComVisible(false)]
    public sealed class ComposeLifecycleOrigin
    {
        public string ServerUrl { get; set; }

        public string Username { get; set; }

        public string AppPassword { get; set; }

        public string AccountFingerprint { get; set; }

        public string AccountId { get; set; }

        internal static ComposeLifecycleOrigin Create(
            TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                return new ComposeLifecycleOrigin();
            }

            string serverUrl = configuration.GetNormalizedBaseUrl();
            string username = (configuration.Username ?? string.Empty).Trim();
            VerifiedNextcloudIdentity identity = NextcloudConnectionState.GetKnownIdentity(configuration);
            return new ComposeLifecycleOrigin
            {
                ServerUrl = serverUrl,
                Username = username,
                AppPassword = configuration.AppPassword ?? string.Empty,
                AccountId = identity != null ? identity.UserId : string.Empty,
                AccountFingerprint = BuildFingerprint(serverUrl, username)
            };
        }

        internal TalkServiceConfiguration ToConfiguration()
        {
            return new TalkServiceConfiguration(
                ServerUrl ?? string.Empty,
                Username ?? string.Empty,
                AppPassword ?? string.Empty);
        }

        internal bool IsComplete()
        {
            return ToConfiguration().IsComplete()
                   && !string.IsNullOrWhiteSpace(AccountFingerprint);
        }

        internal ComposeLifecycleOrigin Clone()
        {
            return new ComposeLifecycleOrigin
            {
                ServerUrl = ServerUrl,
                Username = Username,
                AppPassword = AppPassword,
                AccountId = AccountId,
                AccountFingerprint = AccountFingerprint
            };
        }

        private static string BuildFingerprint(
            string serverUrl,
            string username)
        {
            // Keep the former empty user-id segment so existing account
            // fingerprints remain stable across upgrades.
            string value = (serverUrl ?? string.Empty).Trim().ToLowerInvariant()
                           + "\n\n"
                           + (username ?? string.Empty).Trim().ToLowerInvariant();
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(value));
                var builder = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }
    }
}
