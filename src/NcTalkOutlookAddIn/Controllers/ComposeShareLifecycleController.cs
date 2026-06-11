// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Encapsulates compose share cleanup and separate password mail dispatch.
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

        internal void DispatchSeparatePasswordMailQueue(string composeKey, List<SeparatePasswordDispatchEntry> queue)
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
            int secretsFallbackCount = 0;
            string lastFailureMessage = string.Empty;
            var sentRecipients = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<SeparatePasswordDispatchEntry> dispatchQueue = ExpandSeparatePasswordDispatchEntries(queue);
            int queuedSecrets = CountSecretsDispatches(queue);
            if (queuedSecrets > 0)
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password dispatch queue prepared (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", queued="
                    + queue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", expanded="
                    + dispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", secretsQueued="
                    + queuedSecrets.ToString(CultureInfo.InvariantCulture)
                    + ", secretsExpanded="
                    + CountSecretsDispatches(dispatchQueue).ToString(CultureInfo.InvariantCulture)
                    + ").");
            }
            foreach (var queuedDispatch in dispatchQueue)
            {
                SeparatePasswordDispatchEntry dispatch = PrepareSeparatePasswordDispatch(
                    queuedDispatch,
                    composeKey,
                    ref secretsFallbackCount);
                if (!IsDispatchUsable(dispatch))
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
                    ApplySeparatePasswordSender(passwordMail, dispatch, composeKey);
                    ApplySeparatePasswordBody(passwordMail, dispatch);
                    ApplySeparatePasswordBackendSignature(passwordMail, dispatch, composeKey);
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
            if (secretsFallbackCount > 0)
            {
                ShowPasswordSecretsFallbackDialog();
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

        private static List<SeparatePasswordDispatchEntry> ExpandSeparatePasswordDispatchEntries(List<SeparatePasswordDispatchEntry> queue)
        {
            var expanded = new List<SeparatePasswordDispatchEntry>();
            if (queue == null)
            {
                return expanded;
            }

            for (int i = 0; i < queue.Count; i++)
            {
                SeparatePasswordDispatchEntry entry = queue[i];
                if (entry == null || entry.DeliveryMode != SharePasswordDeliveryMode.Secrets)
                {
                    expanded.Add(entry);
                    continue;
                }

                int added = 0;
                added += AddPerRecipientEntries(expanded, entry, ExtractRecipientAddresses(entry.To), "to");
                added += AddPerRecipientEntries(expanded, entry, ExtractRecipientAddresses(entry.Cc), "cc");
                added += AddPerRecipientEntries(expanded, entry, ExtractRecipientAddresses(entry.Bcc), "bcc");
                if (added == 0)
                {
                    expanded.Add(entry);
                }
            }

            return expanded;
        }

        private static int CountSecretsDispatches(List<SeparatePasswordDispatchEntry> queue)
        {
            if (queue == null)
            {
                return 0;
            }

            int count = 0;
            for (int i = 0; i < queue.Count; i++)
            {
                SeparatePasswordDispatchEntry entry = queue[i];
                if (entry != null && entry.DeliveryMode == SharePasswordDeliveryMode.Secrets)
                {
                    count++;
                }
            }
            return count;
        }

        private static int AddPerRecipientEntries(
            List<SeparatePasswordDispatchEntry> target,
            SeparatePasswordDispatchEntry source,
            List<string> recipients,
            string field)
        {
            if (target == null || source == null || recipients == null || recipients.Count == 0)
            {
                return 0;
            }

            int added = 0;
            for (int i = 0; i < recipients.Count; i++)
            {
                string address = NormalizeRecipientAddress(recipients[i]);
                if (string.IsNullOrWhiteSpace(address))
                {
                    continue;
                }
                SeparatePasswordDispatchEntry clone = ClonePasswordDispatch(source);
                clone.To = string.Equals(field, "to", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                clone.Cc = string.Equals(field, "cc", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                clone.Bcc = string.Equals(field, "bcc", StringComparison.OrdinalIgnoreCase) ? address : string.Empty;
                target.Add(clone);
                added++;
            }
            return added;
        }

        private SeparatePasswordDispatchEntry PrepareSeparatePasswordDispatch(
            SeparatePasswordDispatchEntry dispatch,
            string composeKey,
            ref int secretsFallbackCount)
        {
            if (dispatch == null || dispatch.DeliveryMode != SharePasswordDeliveryMode.Secrets)
            {
                return dispatch;
            }

            try
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password Secrets link create start (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", to="
                    + CountRecipientsInCsv(dispatch.To).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + CountRecipientsInCsv(dispatch.Cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + CountRecipientsInCsv(dispatch.Bcc).ToString(CultureInfo.InvariantCulture)
                    + ", expireDays="
                    + dispatch.SecretsExpireDays.ToString(CultureInfo.InvariantCulture)
                    + ", mailPlainText="
                    + dispatch.IsPlainText.ToString(CultureInfo.InvariantCulture)
                    + ").");
                SecretsLinkResult secret = CreatePasswordSecret(dispatch);
                SeparatePasswordDispatchEntry prepared = BuildPasswordDispatchBody(
                    dispatch,
                    secret.ShareUrl,
                    SharePasswordDeliveryMode.Secrets,
                    secretLink: true);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password Secrets link created (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", hasUuid="
                    + (!string.IsNullOrWhiteSpace(secret.Uuid)).ToString(CultureInfo.InvariantCulture)
                    + ", hasExpires="
                    + secret.Expires.HasValue.ToString(CultureInfo.InvariantCulture)
                    + ", hasHtml="
                    + (!string.IsNullOrWhiteSpace(prepared.Html)).ToString(CultureInfo.InvariantCulture)
                    + ", hasPlainText="
                    + (!string.IsNullOrWhiteSpace(prepared.PlainText)).ToString(CultureInfo.InvariantCulture)
                    + ").");
                return prepared;
            }
            catch (Exception ex)
            {
                secretsFallbackCount++;
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Separate password Secrets link creation failed, falling back to plain mail (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                SeparatePasswordDispatchEntry fallback = ClonePasswordDispatch(dispatch);
                fallback.DeliveryMode = SharePasswordDeliveryMode.Plain;
                return fallback;
            }
        }

        private SecretsLinkResult CreatePasswordSecret(SeparatePasswordDispatchEntry dispatch)
        {
            _owner.EnsureSettingsLoaded();
            if (_owner.CurrentSettings == null || !_owner.SettingsAreComplete())
            {
                throw new InvalidOperationException("Settings are incomplete.");
            }

            AddinSettings settings = _owner.CurrentSettings;
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            var service = new SecretsService(configuration);
            return service.CreateSecretLink(
                dispatch.Password,
                BuildSecretsTitle(dispatch),
                dispatch.SecretsExpireDays);
        }

        private static string BuildSecretsTitle(SeparatePasswordDispatchEntry dispatch)
        {
            string shareLabel = dispatch != null ? (dispatch.ShareLabel ?? string.Empty).Trim() : string.Empty;
            if (string.IsNullOrWhiteSpace(shareLabel))
            {
                return "NCC share password";
            }
            return "NCC " + shareLabel;
        }

        private static SeparatePasswordDispatchEntry BuildPasswordDispatchBody(
            SeparatePasswordDispatchEntry source,
            string deliveryValue,
            SharePasswordDeliveryMode mode,
            bool secretLink)
        {
            FileLinkResult result = BuildFileLinkResultForDelivery(source, deliveryValue);
            string html = source.IsPlainText
                ? string.Empty
                : FileLinkHtmlBuilder.BuildPasswordOnly(result, source.LanguageOverride, source.BackendPolicyStatus, secretLink);
            string plainText = source.IsPlainText
                ? FileLinkHtmlBuilder.BuildPasswordOnlyPlainText(result, source.LanguageOverride, source.BackendPolicyStatus, secretLink)
                : string.Empty;

            SeparatePasswordDispatchEntry prepared = ClonePasswordDispatch(source);
            prepared.Password = deliveryValue ?? string.Empty;
            prepared.Html = html;
            prepared.PlainText = plainText;
            prepared.DeliveryMode = mode;
            return prepared;
        }

        private static FileLinkResult BuildFileLinkResultForDelivery(SeparatePasswordDispatchEntry dispatch, string deliveryValue)
        {
            if (dispatch == null)
            {
                throw new ArgumentNullException("dispatch");
            }

            return new FileLinkResult(
                dispatch.ShareUrl,
                dispatch.ShareId,
                dispatch.ShareToken,
                deliveryValue ?? string.Empty,
                dispatch.ExpireDate,
                dispatch.Permissions,
                dispatch.ShareLabel,
                dispatch.RelativePath);
        }

        private static SeparatePasswordDispatchEntry ClonePasswordDispatch(SeparatePasswordDispatchEntry source)
        {
            if (source == null)
            {
                return null;
            }

            return new SeparatePasswordDispatchEntry
            {
                ShareLabel = source.ShareLabel,
                ShareUrl = source.ShareUrl,
                ShareId = source.ShareId,
                ShareToken = source.ShareToken,
                RelativePath = source.RelativePath,
                ExpireDate = source.ExpireDate,
                Permissions = source.Permissions,
                Password = source.Password,
                Html = source.Html,
                PlainText = source.PlainText,
                IsPlainText = source.IsPlainText,
                DeliveryMode = source.DeliveryMode,
                SecretsExpireDays = source.SecretsExpireDays,
                LanguageOverride = source.LanguageOverride,
                BackendPolicyStatus = source.BackendPolicyStatus,
                To = source.To,
                Cc = source.Cc,
                Bcc = source.Bcc,
                SenderEmail = source.SenderEmail,
                SendUsingAccountSmtpAddress = source.SendUsingAccountSmtpAddress,
                SentOnBehalfOfName = source.SentOnBehalfOfName
            };
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

        private bool TryOpenSeparatePasswordFallback(SeparatePasswordDispatchEntry dispatch, string composeKey)
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
                ApplySeparatePasswordSender(fallback, dispatch, composeKey);
                ApplySeparatePasswordBody(fallback, dispatch);
                ApplySeparatePasswordBackendSignature(fallback, dispatch, composeKey);
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
            SeparatePasswordDispatchEntry dispatch,
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

        private static bool IsDispatchUsable(SeparatePasswordDispatchEntry dispatch)
        {
            if (dispatch == null || string.IsNullOrWhiteSpace(dispatch.Password))
            {
                return false;
            }

            return dispatch.IsPlainText
                ? !string.IsNullOrWhiteSpace(dispatch.PlainText)
                : !string.IsNullOrWhiteSpace(dispatch.Html);
        }

        private static void ApplySeparatePasswordBody(Outlook.MailItem mail, SeparatePasswordDispatchEntry dispatch)
        {
            if (mail == null || dispatch == null)
            {
                return;
            }

            if (dispatch.IsPlainText)
            {
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                mail.Body = PlainTextUtilities.NormalizeCrLfAndTrim(dispatch.PlainText);
                return;
            }

            mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mail.HTMLBody = dispatch.Html ?? string.Empty;
        }

        private void ApplySeparatePasswordSender(Outlook.MailItem mail, SeparatePasswordDispatchEntry dispatch, string composeKey)
        {
            if (mail == null || dispatch == null)
            {
                return;
            }

            string accountSmtp = EmailSignaturePolicyService.NormalizeEmail(dispatch.SendUsingAccountSmtpAddress);
            if (!string.IsNullOrWhiteSpace(accountSmtp))
            {
                TrySetSeparatePasswordSendUsingAccount(mail, accountSmtp, composeKey);
            }

            string sentOnBehalfOfName = (dispatch.SentOnBehalfOfName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(sentOnBehalfOfName))
            {
                return;
            }

            try
            {
                mail.SentOnBehalfOfName = sentOnBehalfOfName;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to set separate password sent-on-behalf identity (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
            }
        }

        private bool TrySetSeparatePasswordSendUsingAccount(Outlook.MailItem mail, string smtpAddress, string composeKey)
        {
            if (mail == null || _owner.OutlookApplication == null || string.IsNullOrWhiteSpace(smtpAddress))
            {
                return false;
            }

            Outlook.NameSpace session = null;
            Outlook.Accounts accounts = null;
            try
            {
                session = _owner.OutlookApplication.Session;
                if (session == null)
                {
                    return false;
                }

                accounts = session.Accounts;
                if (accounts == null)
                {
                    return false;
                }

                int count = accounts.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Account account = null;
                    try
                    {
                        account = accounts[i];
                        string accountSmtp = EmailSignaturePolicyService.NormalizeEmail(account != null ? account.SmtpAddress : string.Empty);
                        if (!string.Equals(accountSmtp, smtpAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        mail.SendUsingAccount = account;
                        NextcloudTalkAddIn.LogFileLinkMessage(
                            "Separate password send account applied (composeKey="
                            + (composeKey ?? string.Empty)
                            + ", hasAccount=True).");
                        return true;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            account,
                            LogCategories.FileLink,
                            "Failed to release separate password Account COM object.");
                    }
                }

                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Separate password send account not found (composeKey="
                    + (composeKey ?? string.Empty)
                    + ", hasRequestedAccount=True).");
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to apply separate password send account (composeKey=" + (composeKey ?? string.Empty) + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(accounts, LogCategories.FileLink, "Failed to release separate password Accounts COM object.");
                ComInteropScope.TryRelease(session, LogCategories.FileLink, "Failed to release separate password Session COM object.");
            }
        }

        private void ApplySeparatePasswordBackendSignature(Outlook.MailItem mail, SeparatePasswordDispatchEntry dispatch, string composeKey)
        {
            if (mail == null || dispatch == null)
            {
                return;
            }
            _owner.EnsureSettingsLoaded();
            if (_owner.CurrentSettings == null || !_owner.SettingsAreComplete())
            {
                LogSeparatePasswordSignatureSkipped(composeKey, "settings_incomplete");
                return;
            }

            AddinSettings settings = _owner.CurrentSettings ?? new AddinSettings();
            var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            BackendPolicyStatus policyStatus = _owner.FetchBackendPolicyStatus(configuration, "separate_password_email_signature");
            var policy = new EmailSignaturePolicyService(policyStatus, settings).Resolve();
            if (!policy.Active)
            {
                LogSeparatePasswordSignatureSkipped(composeKey, policy.Reason);
                return;
            }

            string senderEmail = EmailSignaturePolicyService.NormalizeEmail(dispatch.SenderEmail);
            if (!string.Equals(senderEmail, policy.UserEmail, StringComparison.OrdinalIgnoreCase))
            {
                LogSeparatePasswordSignatureSkipped(composeKey, "identity_mismatch");
                return;
            }

            string sanitized = HtmlTemplateSanitizer.SanitizeEmailSignatureTemplateHtml(policy.TemplateHtml);
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                LogSeparatePasswordSignatureSkipped(composeKey, "sanitized_empty");
                return;
            }

            if (dispatch.IsPlainText)
            {
                string plainText = HtmlToPlainTextConverter.Convert(sanitized);
                if (string.IsNullOrWhiteSpace(plainText))
                {
                    LogSeparatePasswordSignatureSkipped(composeKey, "plain_text_empty");
                    return;
                }

                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                mail.Body = CombinePlainTextSegments(mail.Body, plainText);
            }
            else
            {
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                mail.HTMLBody = AppendHtmlSignature(mail.HTMLBody, sanitized);
            }

            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password backend signature applied (composeKey="
                + (composeKey ?? string.Empty)
                + ", plainText="
                + dispatch.IsPlainText.ToString(CultureInfo.InvariantCulture)
                + ").");
        }

        private static string CombinePlainTextSegments(string body, string signature)
        {
            string normalizedBody = PlainTextUtilities.NormalizeCrLfAndTrim(body);
            string normalizedSignature = PlainTextUtilities.NormalizeCrLfAndTrim(signature);
            if (string.IsNullOrWhiteSpace(normalizedBody))
            {
                return normalizedSignature;
            }
            if (string.IsNullOrWhiteSpace(normalizedSignature))
            {
                return normalizedBody;
            }
            return normalizedBody + "\r\n\r\n" + normalizedSignature;
        }

        private static string AppendHtmlSignature(string html, string sanitizedSignature)
        {
            string existing = html ?? string.Empty;
            string signatureBlock = "<br><br><div data-nc-connector-signature=\"true\">" + (sanitizedSignature ?? string.Empty) + "</div>";
            int bodyEnd = existing.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyEnd >= 0)
            {
                return existing.Insert(bodyEnd, signatureBlock);
            }
            return existing + signatureBlock;
        }

        private static void LogSeparatePasswordSignatureSkipped(string composeKey, string reason)
        {
            NextcloudTalkAddIn.LogFileLinkMessage(
                "Separate password backend signature skipped (composeKey="
                + (composeKey ?? string.Empty)
                + ", reason="
                + (reason ?? "n/a")
                + ").");
        }

        private static string BuildSeparatePasswordMailSubject(SeparatePasswordDispatchEntry dispatch)
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

        private static void ShowPasswordSecretsFallbackDialog()
        {
            try
            {
                MessageBox.Show(
                    Strings.SharingPasswordSecretsFallbackWarning,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to show separate password Secrets fallback warning.", ex);
            }
        }
    }
}

