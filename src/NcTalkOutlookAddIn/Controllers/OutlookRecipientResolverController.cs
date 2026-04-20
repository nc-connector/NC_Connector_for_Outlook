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
     * Centralized Outlook recipient resolution for attendee extraction and SMTP mapping.
     */
    internal static class OutlookRecipientResolverController
    {
        internal static List<string> CollectAppointmentAttendeeEmails(Outlook.AppointmentItem appointment)
        {
            var emails = new List<string>();
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (appointment == null)
            {
                return emails;
            }

            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment recipients.", ex);
                recipients = null;
            }

            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (recipients == null)
            {
                return emails;
            }

            try
            {
                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                        if (recipient == null)
                        {
                            continue;
                        }

                        int type = 0;
                        try
                        {
                            type = recipient.Type;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read recipient.Type.", ex);
                            type = 0;
                        }

                        // 1=Required, 2=Optional, 3=Resource
                        if (type == 3)
                        {
                            continue;
                        }

                        string email = TryResolveRecipientSmtpAddress(recipient);
                        if (string.IsNullOrWhiteSpace(email))
                        {
                            continue;
                        }

                        email = email.Trim().ToLowerInvariant();
                        if (!emails.Contains(email))
                        {
                            emails.Add(email);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(recipient, LogCategories.Talk, "Failed to release Recipient COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to enumerate appointment recipients.", ex);
            }
            finally
            {
                ComInteropScope.TryRelease(recipients, LogCategories.Talk, "Failed to release Recipients COM object.");
            }

            return emails;
        }

        internal static string TryResolveRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
            if (recipient == null)
            {
                return null;
            }

            string address = null;
            try
            {
                address = recipient.Address;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.Address.", ex);
                address = null;
            }

            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            Outlook.AddressEntry entry = null;
            try
            {
                entry = recipient.AddressEntry;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read Recipient.AddressEntry.", ex);
                entry = null;
            }

            // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
            if (entry != null)
            {
                try
                {
                    Outlook.ExchangeUser exUser = null;
                    try
                    {
                        exUser = entry.GetExchangeUser();
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve Exchange user from address entry.", ex);
                        exUser = null;
                    }

                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (exUser != null)
                    {
                        try
                        {
                            address = exUser.PrimarySmtpAddress;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read ExchangeUser.PrimarySmtpAddress.", ex);
                            address = null;
                        }

                        ComInteropScope.TryRelease(exUser, LogCategories.Talk, "Failed to release ExchangeUser COM object.");
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(entry, LogCategories.Talk, "Failed to release AddressEntry COM object.");
                }
            }

            if (!string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0)
            {
                return address;
            }

            try
            {
                Outlook.PropertyAccessor accessor = recipient.PropertyAccessor;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (accessor != null)
                {
                    try
                    {
                        const string SmtpSchema = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        address = accessor.GetProperty(SmtpSchema) as string;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(accessor, LogCategories.Talk, "Failed to release PropertyAccessor COM object.");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve SMTP address via PropertyAccessor.", ex);
            }

            return !string.IsNullOrWhiteSpace(address) && address.IndexOf('@') >= 0 ? address : null;
        }
    }
}
