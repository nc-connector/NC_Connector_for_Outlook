/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Runtime.InteropServices;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Centralizes COM release/final-release behavior and exception-safe logging.
     * This keeps cleanup paths consistent and reduces duplicated Marshal handling.
     */
    internal static class ComInteropScope
    {
        internal static void TryRelease(object comObject, string category, string failureMessage)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(category, failureMessage, ex);
            }
        }

        internal static void TryFinalRelease(object comObject, string category, string failureMessage)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(category, failureMessage, ex);
            }
        }
    }

    /**
     * Disposable COM wrapper used when a local COM object should always be released
     * at scope end (for example in method-level finally-style flows).
     */
    internal sealed class ComObjectScope<T> : IDisposable where T : class
    {
        private object _instance;
        private readonly string _category;
        private readonly string _failureMessage;
        private readonly bool _finalRelease;

        internal ComObjectScope(T instance, string category, string failureMessage, bool finalRelease = false)
        {
            _instance = instance;
            _category = category ?? string.Empty;
            _failureMessage = failureMessage ?? string.Empty;
            _finalRelease = finalRelease;
        }

        internal T Value
        {
            get { return _instance as T; }
        }

        public void Dispose()
        {
            object instance = _instance;
            _instance = null;
            if (_finalRelease)
            {
                ComInteropScope.TryFinalRelease(instance, _category, _failureMessage);
                return;
            }

            ComInteropScope.TryRelease(instance, _category, _failureMessage);
        }
    }
}
