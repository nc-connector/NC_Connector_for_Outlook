/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Shared JSON helpers for OCS payload normalization and dictionary access.
     */
    internal static class NcJson
    {
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        internal static string PrepareJsonPayload(string responseText)
        {
            if (string.IsNullOrWhiteSpace(responseText))
            {
                return string.Empty;
            }

            string payload = responseText.Trim().TrimStart('\uFEFF');
            if (payload.StartsWith(")]}',", StringComparison.Ordinal))
            {
                int newlineIndex = payload.IndexOf('\n');
                payload = newlineIndex >= 0 ? payload.Substring(newlineIndex + 1) : string.Empty;
            }

            if (payload.StartsWith("while(1);", StringComparison.Ordinal))
            {
                payload = payload.Substring("while(1);".Length);
            }

            if (payload.StartsWith("for(;;);", StringComparison.Ordinal))
            {
                payload = payload.Substring("for(;;);".Length);
            }

            return payload.Trim();
        }

        internal static IDictionary<string, object> DeserializeObject(string responseText)
        {
            string payload = PrepareJsonPayload(responseText);
            if (string.IsNullOrWhiteSpace(payload))
            {
                return null;
            }

            return Serializer.DeserializeObject(payload) as IDictionary<string, object>;
        }

        internal static string Serialize(object value)
        {
            return Serializer.Serialize(value);
        }

        internal static IDictionary<string, object> GetDictionary(IDictionary<string, object> parent, string key)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            object value;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (!parent.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            return value as IDictionary<string, object>;
        }

        internal static string GetString(IDictionary<string, object> parent, string key)
        {
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            object value;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (!parent.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            if (value is IDictionary<string, object>)
            {
                return null;
            }

            return value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        internal static string GetTrimmedString(IDictionary<string, object> parent, string key)
        {
            string value = GetString(parent, key);
            return value == null ? null : value.Trim();
        }

        internal static string GetStringOrEmpty(IDictionary<string, object> parent, string key)
        {
            string value = GetTrimmedString(parent, key);
            return value ?? string.Empty;
        }

        internal static bool TryGetInt(IDictionary<string, object> parent, string key, out int value)
        {
            value = 0;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (parent == null || string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            object raw;
            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (!parent.TryGetValue(key, out raw) || raw == null)
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
                long longValue = (long)raw;
                if (longValue < int.MinValue || longValue > int.MaxValue)
                {
                    return false;
                }

                value = (int)longValue;
                return true;
            }

            if (raw is bool)
            {
                value = (bool)raw ? 1 : 0;
                return true;
            }

            int parsed;
            if (int.TryParse(Convert.ToString(raw, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        internal static int GetInt(IDictionary<string, object> parent, string key)
        {
            int value;
            return TryGetInt(parent, key, out value) ? value : 0;
        }

        internal static IDictionary<string, object> GetOcsData(IDictionary<string, object> payload)
        {
            return GetDictionary(GetDictionary(payload, "ocs"), "data");
        }

        internal static IDictionary<string, object> GetOcsMeta(IDictionary<string, object> payload)
        {
            return GetDictionary(GetDictionary(payload, "ocs"), "meta");
        }

        internal static string ExtractOcsErrorMessage(IDictionary<string, object> payload)
        {
            IDictionary<string, object> meta = GetOcsMeta(payload);
            IDictionary<string, object> data = GetOcsData(payload);
            string message = GetTrimmedString(meta, "message");
            string detail = GetTrimmedString(data, "error");

            if (string.IsNullOrWhiteSpace(message))
            {
                return detail ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(detail) || string.Equals(message, detail, StringComparison.OrdinalIgnoreCase))
            {
                return message;
            }

            return message + " / " + detail;
        }
    }
}
