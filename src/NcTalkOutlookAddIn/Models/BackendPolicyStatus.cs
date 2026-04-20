/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;

namespace NcTalkOutlookAddIn.Models
{
    /**
     * Normalized backend policy runtime snapshot.
     *
     * The status is loaded from NC Connector backend endpoint:
     * /apps/ncc_backend_4mc/api/v1/status
     */
    internal sealed class BackendPolicyStatus
    {
        internal BackendPolicyStatus(
            bool endpointAvailable,
            bool fetchSucceeded,
            bool policyActive,
            bool warningVisible,
            string warningMessage,
            string mode,
            string reason,
            bool seatAssigned,
            bool isValid,
            string seatState,
            IDictionary<string, object> sharePolicy,
            IDictionary<string, object> talkPolicy,
            IDictionary<string, object> shareEditable,
            IDictionary<string, object> talkEditable)
        {
            EndpointAvailable = endpointAvailable;
            FetchSucceeded = fetchSucceeded;
            PolicyActive = policyActive;
            WarningVisible = warningVisible;
            WarningMessage = warningMessage ?? string.Empty;
            Mode = mode ?? "local";
            Reason = reason ?? string.Empty;
            SeatAssigned = seatAssigned;
            IsValid = isValid;
            SeatState = seatState ?? string.Empty;
            SharePolicy = sharePolicy;
            TalkPolicy = talkPolicy;
            ShareEditable = shareEditable;
            TalkEditable = talkEditable;
        }

        internal bool EndpointAvailable { get; private set; }

        internal bool FetchSucceeded { get; private set; }

        internal bool PolicyActive { get; private set; }

        internal bool WarningVisible { get; private set; }

        internal string WarningMessage { get; private set; }

        internal string Mode { get; private set; }

        internal string Reason { get; private set; }

        internal bool SeatAssigned { get; private set; }

        internal bool IsValid { get; private set; }

        internal string SeatState { get; private set; }

        internal IDictionary<string, object> SharePolicy { get; private set; }

        internal IDictionary<string, object> TalkPolicy { get; private set; }

        internal IDictionary<string, object> ShareEditable { get; private set; }

        internal IDictionary<string, object> TalkEditable { get; private set; }

        /**
         * Read one policy value for a domain/key pair.
         */
        internal object GetPolicyValue(string domain, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            IDictionary<string, object> policy = GetDomainDictionary(domain, true);            if (policy == null)
            {
                return null;
            }

            object value;
            if (policy.TryGetValue(key, out value))
            {
                return value;
            }

            return null;
        }

        /**
         * Return true when the policy payload explicitly contains the key,
         * even if the value is `null`.
         */
        internal bool HasPolicyKey(string domain, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            IDictionary<string, object> policy = GetDomainDictionary(domain, true);            if (policy == null)
            {
                return false;
            }

            return policy.ContainsKey(key);
        }

        /**
         * Return true when a setting is backend-locked (policy_editable == false).
         */
        internal bool IsLocked(string domain, string key)
        {
            if (!PolicyActive || string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            IDictionary<string, object> editable = GetDomainDictionary(domain, false);            if (editable == null)
            {
                return false;
            }

            object rawEditable;
            if (!editable.TryGetValue(key, out rawEditable))
            {
                return false;
            }

            bool isEditable;
            if (!TryConvertBool(rawEditable, out isEditable))
            {
                return false;
            }

            return !isEditable;
        }

        /**
         * Read one boolean policy value.
         */
        internal bool TryGetPolicyBool(string domain, string key, out bool value)
        {
            value = false;
            object raw = GetPolicyValue(domain, key);            if (raw == null)
            {
                return false;
            }

            return TryConvertBool(raw, out value);
        }

        /**
         * Read one integer policy value.
         */
        internal bool TryGetPolicyInt(string domain, string key, out int value)
        {
            value = 0;
            object raw = GetPolicyValue(domain, key);            if (raw == null)
            {
                return false;
            }

            return TryConvertInt(raw, out value);
        }

        /**
         * Read one string policy value.
         */
        internal string GetPolicyString(string domain, string key)
        {
            object raw = GetPolicyValue(domain, key);            if (raw == null)
            {
                return string.Empty;
            }

            string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
            return text == null ? string.Empty : text.Trim();
        }

        private IDictionary<string, object> GetDomainDictionary(string domain, bool readPolicy)
        {
            string normalized = string.Equals(domain, "talk", StringComparison.OrdinalIgnoreCase) ? "talk" : "share";
            if (readPolicy)
            {
                return normalized == "talk" ? TalkPolicy : SharePolicy;
            }

            return normalized == "talk" ? TalkEditable : ShareEditable;
        }

        internal static bool TryConvertBool(object raw, out bool value)
        {
            value = false;            if (raw == null)
            {
                return false;
            }

            if (raw is bool)
            {
                value = (bool)raw;
                return true;
            }

            if (raw is int)
            {
                value = (int)raw != 0;
                return true;
            }

            if (raw is long)
            {
                value = (long)raw != 0L;
                return true;
            }

            if (raw is double)
            {
                value = Math.Abs((double)raw) > 0.0001d;
                return true;
            }

            string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            text = text.Trim();
            if (string.Equals(text, "1", StringComparison.OrdinalIgnoreCase))
            {
                value = true;
                return true;
            }
            if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase))
            {
                value = false;
                return true;
            }
            if (string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase))
            {
                value = true;
                return true;
            }
            if (string.Equals(text, "no", StringComparison.OrdinalIgnoreCase))
            {
                value = false;
                return true;
            }

            bool parsed;
            if (bool.TryParse(text, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        internal static bool TryConvertInt(object raw, out int value)
        {
            value = 0;            if (raw == null)
            {
                return false;
            }

            if (raw is int)
            {
                value = (int)raw;
                return true;
            }

            if (raw is long)
            {
                long asLong = (long)raw;
                if (asLong < int.MinValue || asLong > int.MaxValue)
                {
                    return false;
                }

                value = (int)asLong;
                return true;
            }

            string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            int parsed;
            if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }
    }
}

