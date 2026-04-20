/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    /**
     * Thread-safe registry for compose subscriptions.
     * Keeps creation/removal/disposal policy in one place.
     */
    internal sealed class MailComposeSubscriptionRegistryController
    {
        private readonly object _syncRoot = new object();
        private readonly List<NextcloudTalkAddIn.MailComposeSubscription> _subscriptions = new List<NextcloudTalkAddIn.MailComposeSubscription>();

        internal NextcloudTalkAddIn.MailComposeSubscription GetOrCreate(
            Outlook.MailItem mail,
            string mailIdentityKey,
            string inspectorIdentityKey,
            Func<NextcloudTalkAddIn.MailComposeSubscription> factory)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (mail == null || factory == null)
            {
                return null;
            }

            lock (_syncRoot)
            {
                for (int i = 0; i < _subscriptions.Count; i++)
                {
                    NextcloudTalkAddIn.MailComposeSubscription existing = _subscriptions[i];
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (existing != null && existing.IsFor(mail, mailIdentityKey, inspectorIdentityKey))
                    {
                        return existing;
                    }
                }

                NextcloudTalkAddIn.MailComposeSubscription created = factory();
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (created != null)
                {
                    _subscriptions.Add(created);
                }

                return created;
            }
        }

        internal void Remove(NextcloudTalkAddIn.MailComposeSubscription subscription)
        {
            // Idempotent unsubscribe path: disposal races can call Remove with null.
            if (subscription == null)
            {
                return;
            }

            lock (_syncRoot)
            {
                _subscriptions.Remove(subscription);
            }
        }

        internal void DisposeAll()
        {
            NextcloudTalkAddIn.MailComposeSubscription[] current;
            lock (_syncRoot)
            {
                current = _subscriptions.ToArray();
                _subscriptions.Clear();
            }

            for (int i = 0; i < current.Length; i++)
            {
                try
                {
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (current[i] != null)
                    {
                        current[i].Dispose();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose mail compose subscription.", ex);
                }
            }
        }
    }
}
