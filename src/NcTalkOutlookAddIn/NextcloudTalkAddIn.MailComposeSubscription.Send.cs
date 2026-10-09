// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        // Handles send validation and separate-password preparation.
        internal sealed partial class MailComposeSubscription
        {
            private bool _sendPolicySignatureWarning;
            private bool _sendPolicyAttachmentWarning;
            private BackendPolicyStatus _sendPolicyWarningCheck;
            private string _lastSendPolicyWarning = string.Empty;
            private bool _sendAccepted;

            private void OnSend(ref bool cancel)
            {
                if (_disposed || cancel)
                {
                    return;
                }
                _sendPolicySignatureWarning = false;
                _sendAccepted = false;
                _sendPolicyAttachmentWarning = false;
                _sendPolicyWarningCheck = null;
                if (!TryValidateKnownSendPolicyBeforeSend(ref cancel)
                    || !TryValidateAttachmentPolicyBeforeSend(ref cancel)
                    || !TryFinalizeEmailSignatureBeforeSend(ref cancel))
                {
                    return;
                }
                ShowSendPolicyWarning();
                _sendAccepted = true;
                _emailSignatureTimer.Stop();
                _emailSignatureRequestGeneration++;

                // Send provides the final account and recipients. The primary may then stay
                // in Outbox, so password delivery follows this action, not later transport.
                CapturePasswordDispatchRecipients();
                CapturePasswordDispatchSender();

                int passwordDispatchCount =
                    _passwordDispatchQueue.Count;
                if (passwordDispatchCount > 0)
                {
                    var queue =
                        new List<SeparatePasswordDispatchEntry>(
                            _passwordDispatchQueue);
                    // Clear before submission: an ambiguous Outlook result must not create
                    // a second password mail.
                    _passwordDispatchQueue.Clear();
                    LogFileLink(
                        "Separate password direct dispatch triggered by primary send (composeKey="
                        + _composeKey
                        + ", queued="
                        + queue.Count.ToString(
                            CultureInfo.InvariantCulture)
                        + ").");
                    try
                    {
                        _owner.DispatchSeparatePasswordMails(
                            _composeKey,
                            queue);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Separate password direct dispatch failed unexpectedly (composeKey="
                            + _composeKey
                            + "). The primary send continues.",
                            ex);
                    }
                }

                LogFileLink(
                    "Compose send handling completed (composeKey="
                    + _composeKey
                    + ", passwordDispatched="
                    + passwordDispatchCount.ToString(
                        CultureInfo.InvariantCulture)
                    + ").");
            }

            private bool TryValidateKnownSendPolicyBeforeSend(ref bool cancel)
            {
                _owner.EnsureSettingsLoaded();
                var settings = _owner._currentSettings;
                var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
                BackendPolicyStatus known;
                if (_owner.TryGetCachedEmailSignaturePolicyStatus(configuration, out known)
                    && known.FetchSucceeded)
                {
                    return true;
                }
                ScheduleEmailSignatureApplication("send_policy_initial_check");
                if (!settings.SendPolicyFailClosed)
                {
                    return true;
                }
                BackendPolicyStatus check;
                _owner.TryGetCurrentBackendPolicyCheck(configuration, out check);
                return BlockSendPolicyFailure(ref cancel, check, true);
            }

            private void RecordSendPolicyWarning(bool signature, bool attachments, BackendPolicyStatus check)
            {
                _sendPolicySignatureWarning |= signature;
                _sendPolicyAttachmentWarning |= attachments;
                _sendPolicyWarningCheck = check;
            }

            private void ShowSendPolicyWarning()
            {
                if (!_sendPolicySignatureWarning && !_sendPolicyAttachmentWarning)
                {
                    _lastSendPolicyWarning = string.Empty;
                    return;
                }
                string body = _sendPolicySignatureWarning && _sendPolicyAttachmentWarning
                    ? Strings.SendPolicyCombinedWarning
                    : (_sendPolicySignatureWarning ? Strings.SendPolicySignatureWarning : Strings.SendPolicyAttachmentWarning);
                string message = GetSendPolicyFailureMessage(_sendPolicyWarningCheck) + " " + body;
                if (!string.Equals(_lastSendPolicyWarning, message, StringComparison.Ordinal))
                {
                    _lastSendPolicyWarning = message;
                    _owner.ShowComposeWarning(message);
                }
            }

            private bool BlockSendPolicyFailure(ref bool cancel, BackendPolicyStatus check, bool unknown)
            {
                cancel = true;
                string message = unknown
                    ? Strings.EmailSignaturePolicyUnavailable
                    : (check == null || !check.IsServiceUnavailable
                        ? Strings.SendPolicyCheckFailed
                        : GetSendPolicyFailureMessage(check) + " " + Strings.SendPolicyUnavailableBlocked);
                if (check != null && check.Reason == "authentication_rejected")
                {
                    message = Strings.SendPolicyAuthenticationRejected + " "
                              + (unknown ? Strings.EmailSignaturePolicyUnavailable : Strings.SendPolicyCheckFailed);
                }
                else if (unknown && check != null && check.Reason == "rate_limited")
                {
                    message = GetSendPolicyFailureMessage(check) + " " + message;
                }
                DiagnosticsLogger.Log(LogCategories.Core,
                    "Send policy blocked (unknown=" + unknown + ", reason=" + (check != null ? check.Reason : "pending") + ").");
                MessageBox.Show(message, Strings.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            private static string GetSendPolicyFailureMessage(BackendPolicyStatus check)
            {
                if (check != null && check.Reason == "rate_limited")
                {
                    return Strings.SendPolicyRateLimited
                           + (check.RetryAfterUtc > DateTime.UtcNow
                               ? " " + check.RetryAfterUtc.ToLocalTime().ToString("T", CultureInfo.CurrentCulture) : string.Empty);
                }
                return check != null && (check.Reason == "backend_unavailable" || check.IsEndpointMissing)
                    ? Strings.SendPolicyBackendUnavailable : Strings.SendPolicyNextcloudUnavailable;
            }

            private void CapturePasswordDispatchRecipients()
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }
                string to;
                string cc;
                string bcc;
                bool capturedFromRecipients =
                    TryCaptureRecipientListsFromRecipientsCollection(
                        out to,
                        out cc,
                        out bcc);
                if (!capturedFromRecipients)
                {
                    to = RecipientAddressList.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("To"));
                    cc = RecipientAddressList.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("CC"));
                    bcc = RecipientAddressList.BuildNormalizedRecipientCsv(
                        ReadMailRecipientList("BCC"));
                }
                for (int i = 0; i < _passwordDispatchQueue.Count; i++)
                {
                    _passwordDispatchQueue[i].To = to;
                    _passwordDispatchQueue[i].Cc = cc;
                    _passwordDispatchQueue[i].Bcc = bcc;
                }

                LogFileLink(
                    "Separate password recipients captured (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", to="
                    + RecipientAddressList.CountRecipientsInCsv(to).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + RecipientAddressList.CountRecipientsInCsv(cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + RecipientAddressList.CountRecipientsInCsv(bcc).ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (capturedFromRecipients
                        ? "recipients_collection"
                        : "mail_fields")
                    + ").");
            }

            private void CapturePasswordDispatchSender()
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                string senderEmail =
                    EmailSignaturePolicyService.NormalizeEmail(
                        ResolveCurrentSenderEmail());
                string accountSmtp =
                    EmailSignaturePolicyService.NormalizeEmail(
                        OutlookRecipientResolverController
                            .ResolveSendUsingAccountSmtpAddress(
                                _mail,
                                LogCategories.Core,
                                "compose",
                                string.Empty));
                string sentOnBehalfOfName =
                    ReadCurrentSentOnBehalfOfName();
                for (int i = 0; i < _passwordDispatchQueue.Count; i++)
                {
                    _passwordDispatchQueue[i].SenderEmail = senderEmail;
                    _passwordDispatchQueue[i]
                        .SendUsingAccountSmtpAddress = accountSmtp;
                    _passwordDispatchQueue[i]
                        .SentOnBehalfOfName = sentOnBehalfOfName;
                }

                LogFileLink(
                    "Separate password sender captured (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", hasSender="
                    + (!string.IsNullOrWhiteSpace(senderEmail))
                        .ToString(CultureInfo.InvariantCulture)
                    + ", hasAccount="
                    + (!string.IsNullOrWhiteSpace(accountSmtp))
                        .ToString(CultureInfo.InvariantCulture)
                    + ", sentOnBehalf="
                    + (!string.IsNullOrWhiteSpace(sentOnBehalfOfName))
                        .ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private string ReadCurrentSentOnBehalfOfName()
            {
                try
                {
                    return _mail != null
                        ? (_mail.SentOnBehalfOfName ?? string.Empty).Trim()
                        : string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose sent-on-behalf name for separate password mail (composeKey=" + _composeKey + ").",
                        ex);
                    return string.Empty;
                }
            }

            private bool TryCaptureRecipientListsFromRecipientsCollection(
                out string to,
                out string cc,
                out string bcc)
            {
                to = string.Empty;
                cc = string.Empty;
                bcc = string.Empty;
                if (_mail == null)
                {
                    return false;
                }

                var toRecipients = new List<string>();
                var ccRecipients = new List<string>();
                var bccRecipients = new List<string>();
                Outlook.Recipients recipients = null;
                try
                {
                    recipients = _mail.Recipients;
                    if (recipients == null)
                    {
                        return false;
                    }

                    int count;
                    try
                    {
                        count = recipients.Count;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to read compose Recipients.Count (composeKey=" + _composeKey + ").",
                            ex);
                        count = 0;
                    }

                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Recipient recipient = null;
                        try
                        {
                            recipient = recipients[i];
                            if (recipient == null)
                            {
                                continue;
                            }

                            string address =
                                TryGetRecipientSmtpAddress(recipient);
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                try
                                {
                                    address =
                                        RecipientAddressList
                                            .NormalizeRecipientAddress(
                                                recipient.Address);
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(
                                        LogCategories.FileLink,
                                        "Failed to read compose recipient.Address (composeKey=" + _composeKey + ").",
                                        ex);
                                    address = string.Empty;
                                }
                            }
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                continue;
                            }

                            int recipientType;
                            try
                            {
                                recipientType = recipient.Type;
                            }
                            catch (Exception ex)
                            {
                                DiagnosticsLogger.LogException(
                                    LogCategories.FileLink,
                                    "Failed to read compose recipient.Type (composeKey=" + _composeKey + ").",
                                    ex);
                                recipientType =
                                    (int)Outlook.OlMailRecipientType.olTo;
                            }

                            if (recipientType
                                == (int)Outlook.OlMailRecipientType.olCC)
                            {
                                RecipientAddressList
                                    .AddUniqueRecipient(
                                        ccRecipients,
                                        address);
                            }
                            else if (recipientType
                                     == (int)Outlook.OlMailRecipientType
                                         .olBCC)
                            {
                                RecipientAddressList
                                    .AddUniqueRecipient(
                                        bccRecipients,
                                        address);
                            }
                            else
                            {
                                RecipientAddressList
                                    .AddUniqueRecipient(
                                        toRecipients,
                                        address);
                            }
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                recipient,
                                LogCategories.FileLink,
                                "Failed to release compose Recipient COM object.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to capture compose recipients from Recipients collection (composeKey=" + _composeKey + ").",
                        ex);
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        recipients,
                        LogCategories.FileLink,
                        "Failed to release compose Recipients COM object.");
                }

                to = toRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", toRecipients.ToArray());
                cc = ccRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", ccRecipients.ToArray());
                bcc = bccRecipients.Count == 0
                    ? string.Empty
                    : string.Join("; ", bccRecipients.ToArray());
                return toRecipients.Count
                       + ccRecipients.Count
                       + bccRecipients.Count > 0;
            }

            private string ReadMailRecipientList(string fieldName)
            {
                try
                {
                    if (_mail == null)
                    {
                        return string.Empty;
                    }

                    switch ((fieldName ?? string.Empty)
                        .Trim()
                        .ToUpperInvariant())
                    {
                        case "TO":
                            return _mail.To ?? string.Empty;
                        case "CC":
                            return _mail.CC ?? string.Empty;
                        case "BCC":
                            return _mail.BCC ?? string.Empty;
                        default:
                            return string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose recipient field '" + (fieldName ?? string.Empty) + "' (composeKey=" + _composeKey + ").",
                        ex);
                    return string.Empty;
                }
            }

        }
    }
}
