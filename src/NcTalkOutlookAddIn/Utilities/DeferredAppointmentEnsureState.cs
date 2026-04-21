// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;

namespace NcTalkOutlookAddIn.Utilities
{
        // Encapsulates runtime state for deferred appointment subscription ensure.
    // Keeps pending-key tracking and log throttling in one cohesive location.
    internal sealed class DeferredAppointmentEnsureState
    {
        private readonly object _syncRoot = new object();
        private readonly HashSet<string> _pendingKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private DateTime _lastRestrictionLogUtc = DateTime.MinValue;
        private DateTime _lastUnstableIdentityRestrictionLogUtc = DateTime.MinValue;
        private int _suppressedRestrictionCount;

        internal bool ShouldLogRestriction(string message, DateTime nowUtc, out string throttledMessage)
        {
            throttledMessage = null;
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            lock (_syncRoot)
            {
                bool shouldEmit = _lastRestrictionLogUtc == DateTime.MinValue ||
                                  (nowUtc - _lastRestrictionLogUtc).TotalSeconds >= 5;
                if (!shouldEmit)
                {
                    _suppressedRestrictionCount++;
                    return false;
                }
                if (_suppressedRestrictionCount > 0)
                {
                    message = message + " (suppressed " + _suppressedRestrictionCount.ToString(CultureInfo.InvariantCulture) + " similar entries)";
                }

                _lastRestrictionLogUtc = nowUtc;
                _suppressedRestrictionCount = 0;
                throttledMessage = message;
                return true;
            }
        }

        internal bool ShouldLogUnstableIdentityRestriction(DateTime nowUtc)
        {
            lock (_syncRoot)
            {
                bool shouldEmit = _lastUnstableIdentityRestrictionLogUtc == DateTime.MinValue ||
                    (nowUtc - _lastUnstableIdentityRestrictionLogUtc).TotalSeconds >= 60;
                if (shouldEmit)
                {
                    _lastUnstableIdentityRestrictionLogUtc = nowUtc;
                }
                return shouldEmit;
            }
        }

        internal bool TryQueuePendingKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            lock (_syncRoot)
            {
                if (_pendingKeys.Contains(key))
                {
                    return false;
                }

                _pendingKeys.Add(key);
                return true;
            }
        }

        internal void DequeuePendingKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return;
            }

            lock (_syncRoot)
            {
                _pendingKeys.Remove(key);
            }
        }

        internal void ClearPendingKeys()
        {
            lock (_syncRoot)
            {
                _pendingKeys.Clear();
            }
        }
    }
}
