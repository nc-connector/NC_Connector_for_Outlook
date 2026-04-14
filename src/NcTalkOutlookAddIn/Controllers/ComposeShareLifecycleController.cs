/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    /**
     * Encapsulates compose share cleanup and separate password mail dispatch.
     */
    internal sealed class ComposeShareLifecycleController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal ComposeShareLifecycleController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal bool TryDeleteComposeShareFolder(string relativeFolder, string reason, string shareId, string shareLabel)
        {
            if (string.IsNullOrWhiteSpace(relativeFolder))
            {
                return true;
            }

            _owner.EnsureSettingsLoaded();
            if (_owner.CurrentSettings == null || !_owner.SettingsAreComplete())
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Compose share cleanup skipped (settings incomplete): relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty));
                return false;
            }

            var configuration = new TalkServiceConfiguration(
                _owner.CurrentSettings.ServerUrl,
                _owner.CurrentSettings.Username,
                _owner.CurrentSettings.AppPassword);
            var service = new FileLinkService(configuration);
            try
            {
                service.DeleteShareFolder(relativeFolder, CancellationToken.None);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Compose share cleanup delete success (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (shareId ?? string.Empty)
                    + ", shareLabel="
                    + (shareLabel ?? string.Empty)
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Compose share cleanup delete failure (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (shareId ?? string.Empty)
                    + ", shareLabel="
                    + (shareLabel ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
        }

        internal void DispatchSeparatePasswordMailQueue(string composeKey, List<NextcloudTalkAddIn.SeparatePasswordDispatchEntry> queue)
        {
            if (queue == null || queue.Count == 0 || _owner.OutlookApplication == null)
            {
                return;
            }

            int attemptedDispatches = 0;
            int successfulDispatches = 0;
            int autoSendFailures = 0;
            int fallbackOpenedCount = 0;
            int fallbackOpenFailures = 0;
            string lastFailureMessage = string.Empty;
            var sentRecipients = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var dispatch in queue)
            {
                if (dispatch == null || string.IsNullOrWhiteSpace(dispatch.Password) || string.IsNullOrWhiteSpace(dispatch.Html))
                {
                    continue;
                }

                attemptedDispatches++;
                Outlook.MailItem passwordMail = null;
                try
                {
                    passwordMail = _owner.OutlookApplication.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    if (passwordMail == null)
                    {
                        throw new InvalidOperationException("Password mail draft could not be created.");
                    }

                    passwordMail.Subject = BuildSeparatePasswordMailSubject(dispatch);
                    passwordMail.HTMLBody = dispatch.Html;
                    List<string> resolvedRecipients = ApplySeparatePasswordRecipientsForSend(passwordMail, dispatch, composeKey);
                    int resolvedRecipientCount = resolvedRecipients.Count;

                    NextcloudTalkAddIn.LogFileLinkMessage(
                        "Separate password mail send start (composeKey="
                        + (composeKey ?? string.Empty)
                        + ", to="
                        + BuildNormalizedRecipientCsv(dispatch.To)
                        + ", cc="
                        + BuildNormalizedRecipientCsv(dispatch.Cc)
                        + ", bcc="
                        + BuildNormalizedRecipientCsv(dispatch.Bcc)
                        + ", resolvedRecipients="
                        + resolvedRecipientCount.ToString(CultureInfo.InvariantCulture)
                        + ").");

                    ((Outlook._MailItem)passwordMail).Send();
                    successfulDispatches++;
                    AddRecipientAddresses(sentRecipients, resolvedRecipients);
                    NextcloudTalkAddIn.LogFileLinkMessage("Separate password mail send done (composeKey=" + (composeKey ?? string.Empty) + ").");
                }
                catch (Exception ex)
                {
                    autoSendFailures++;
                    lastFailureMessage = ex.Message ?? string.Empty;
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Separate password mail auto-send failed (composeKey=" + (composeKey ?? string.Empty) + ").",
                        ex);
                    // Auto-send is preferred. If that fails we keep the flow recoverable
                    // by opening a prefilled manual draft for the user.
                    bool fallbackOpened = TryOpenSeparatePasswordFallback(dispatch, composeKey);
                    if (fallbackOpened)
                    {
                        fallbackOpenedCount++;
                    }
                    else
                    {
                        fallbackOpenFailures++;
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        passwordMail,
                        LogCategories.FileLink,
                        "Failed to release password MailItem COM object.");
                }
            }

            int recipientCount = sentRecipients.Count;
            if (attemptedDispatches > 0 && successfulDispatches == attemptedDispatches && autoSendFailures == 0 && recipientCount > 0)
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password mail sent (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", attempted="
                    + attemptedDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", successful="
                    + successfulDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", recipients="
                    + recipientCount.ToString(CultureInfo.InvariantCulture)
                    + ").");
                _owner.ShowPasswordMailSuccessNotification(recipientCount);
            }
            else
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password mail partially sent (manual fallback required) (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", attempted="
                    + attemptedDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", successful="
                    + successfulDispatches.ToString(CultureInfo.InvariantCulture)
                    + ", recipients="
                    + recipientCount.ToString(CultureInfo.InvariantCulture)
                    + ", fallbackOpened="
                    + fallbackOpenedCount.ToString(CultureInfo.InvariantCulture)
                    + ", fallbackOpenFailures="
                    + fallbackOpenFailures.ToString(CultureInfo.InvariantCulture)
                    + ", autoSendFailures="
                    + autoSendFailures.ToString(CultureInfo.InvariantCulture)
                    + ").");

                if (autoSendFailures > 0 && fallbackOpenedCount == 0 && !string.IsNullOrWhiteSpace(lastFailureMessage))
                {
                    ShowPasswordMailFailureDialog(lastFailureMessage);
                }
            }
        }

        internal static void AddRecipientAddresses(HashSet<string> recipients, List<string> addresses)
        {
            if (recipients == null || addresses == null)
            {
                return;
            }

            for (int i = 0; i < addresses.Count; i++)
            {
                string normalized = NormalizeRecipientAddress(addresses[i]);
                if (!string.IsNullOrWhiteSpace(normalized))
                {
                    recipients.Add(normalized);
                }
            }
        }

        internal static void AddUniqueRecipient(List<string> recipients, string address)
        {
            if (recipients == null)
            {
                return;
            }

            string normalized = NormalizeRecipientAddress(address);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            for (int i = 0; i < recipients.Count; i++)
            {
                if (string.Equals(recipients[i], normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            recipients.Add(normalized);
        }

        internal static List<string> ExtractRecipientAddresses(string csv)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(csv))
            {
                return list;
            }

            string[] parts = csv.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                AddUniqueRecipient(list, parts[i]);
            }

            return list;
        }

        internal static string BuildNormalizedRecipientCsv(string csv)
        {
            List<string> recipients = ExtractRecipientAddresses(csv);
            return recipients.Count == 0 ? string.Empty : string.Join("; ", recipients.ToArray());
        }

        internal static int CountRecipientsInCsv(string csv)
        {
            return ExtractRecipientAddresses(csv).Count;
        }

        internal static string NormalizeRecipientAddress(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            string value = raw.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            int lt = value.LastIndexOf('<');
            int gt = value.LastIndexOf('>');
            if (lt >= 0 && gt > lt)
            {
                value = value.Substring(lt + 1, gt - lt - 1).Trim();
            }

            value = value.Trim().Trim('\'', '"');
            return value.Trim();
        }

        private bool TryOpenSeparatePasswordFallback(NextcloudTalkAddIn.SeparatePasswordDispatchEntry dispatch, string composeKey)
        {
            if (dispatch == null || _owner.OutlookApplication == null)
            {
                return false;
            }

            Outlook.MailItem fallback = null;
            try
            {
                fallback = _owner.OutlookApplication.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (fallback == null)
                {
                    return false;
                }

                string toRecipients = BuildNormalizedRecipientCsv(dispatch.To);
                string ccRecipients = BuildNormalizedRecipientCsv(dispatch.Cc);
                string bccRecipients = BuildNormalizedRecipientCsv(dispatch.Bcc);
                if (CountRecipientsInCsv(toRecipients) + CountRecipientsInCsv(ccRecipients) + CountRecipientsInCsv(bccRecipients) <= 0)
                {
                    throw new InvalidOperationException("Separate password fallback draft has no valid recipients.");
                }

                fallback.To = toRecipients;
                fallback.CC = ccRecipients;
                fallback.BCC = bccRecipients;
                fallback.Subject = BuildSeparatePasswordMailSubject(dispatch);
                fallback.HTMLBody = dispatch.Html ?? string.Empty;
                fallback.Display(false);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password mail manual fallback opened (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", to="
                    + toRecipients
                    + ", cc="
                    + ccRecipients
                    + ", bcc="
                    + bccRecipients
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password mail manual fallback failed (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    fallback,
                    LogCategories.FileLink,
                    "Failed to release password fallback MailItem COM object.");
            }
        }

        private List<string> ApplySeparatePasswordRecipientsForSend(
            Outlook.MailItem mail,
            NextcloudTalkAddIn.SeparatePasswordDispatchEntry dispatch,
            string composeKey)
        {
            if (mail == null)
            {
                throw new InvalidOperationException("Password mail is not available.");
            }

            List<string> toRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.To : string.Empty);
            List<string> ccRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.Cc : string.Empty);
            List<string> bccRecipients = ExtractRecipientAddresses(dispatch != null ? dispatch.Bcc : string.Empty);
            int totalRecipients = toRecipients.Count + ccRecipients.Count + bccRecipients.Count;
            if (totalRecipients <= 0)
            {
                throw new InvalidOperationException("Separate password mail has no valid recipients.");
            }

            var resolvedRecipients = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                if (recipients == null)
                {
                    throw new InvalidOperationException("Password mail recipients collection is not available.");
                }

                AddResolvedRecipients(recipients, toRecipients, Outlook.OlMailRecipientType.olTo, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, ccRecipients, Outlook.OlMailRecipientType.olCC, composeKey, resolvedRecipients);
                AddResolvedRecipients(recipients, bccRecipients, Outlook.OlMailRecipientType.olBCC, composeKey, resolvedRecipients);

                bool resolvedAll = recipients.ResolveAll();
                if (!resolvedAll)
                {
                    throw new InvalidOperationException("Separate password mail recipients could not be resolved.");
                }

                if (resolvedRecipients.Count <= 0)
                {
                    throw new InvalidOperationException("Separate password mail has no resolvable recipients.");
                }

                return resolvedRecipients;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    recipients,
                    LogCategories.FileLink,
                    "Failed to release password Recipients COM object.");
            }
        }

        private static void AddResolvedRecipients(
            Outlook.Recipients recipients,
            List<string> addresses,
            Outlook.OlMailRecipientType type,
            string composeKey,
            List<string> resolvedRecipients)
        {
            if (recipients == null || addresses == null || addresses.Count == 0)
            {
                return;
            }

            for (int i = 0; i < addresses.Count; i++)
            {
                string address = addresses[i] ?? string.Empty;
                Outlook.Recipient recipient = null;
                try
                {
                    recipient = recipients.Add(address);
                    if (recipient == null)
                    {
                        throw new InvalidOperationException("Recipient could not be added.");
                    }

                    recipient.Type = (int)type;
                    bool resolved = recipient.Resolve();
                    if (!resolved)
                    {
                        throw new InvalidOperationException(
                            "Recipient could not be resolved (composeKey="
                            + (composeKey ?? string.Empty)
                            + ", address="
                            + address
                            + ", type="
                            + type.ToString()
                            + ").");
                    }

                    AddUniqueRecipient(resolvedRecipients, address);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        recipient,
                        LogCategories.FileLink,
                        "Failed to release password Recipient COM object.");
                }
            }
        }

        private static string BuildSeparatePasswordMailSubject(NextcloudTalkAddIn.SeparatePasswordDispatchEntry dispatch)
        {
            string baseSubject = Strings.SharingPasswordMailSubject;
            string shareLabel = dispatch != null ? (dispatch.ShareLabel ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(shareLabel))
            {
                return baseSubject;
            }

            return string.Format(CultureInfo.CurrentCulture, Strings.SharingPasswordMailSubjectWithLabel, shareLabel);
        }

        private static void ShowPasswordMailFailureDialog(string detailMessage)
        {
            if (string.IsNullOrWhiteSpace(detailMessage))
            {
                return;
            }

            try
            {
                MessageBox.Show(
                    detailMessage.Trim(),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to show separate password failure dialog.", ex);
            }
        }
    }
}
