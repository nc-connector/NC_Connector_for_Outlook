// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private void OnSend(ref bool cancel)
            {
                if (_disposed)
                {
                    return;
                }
                if (cancel)
                {
                    LogFileLink("Compose send cancelled before dispatch handling (composeKey=" + _composeKey + ").");
                    return;
                }

                _sendPending = true;
                _sendPendingAtUtc = DateTime.UtcNow;
                _cleanupGraceTimer.Stop();
                _awaitingGraceCloseResolution = false;

                CapturePasswordDispatchRecipients();
                LogFileLink(
                    "Compose send state updated (composeKey="
                    + _composeKey
                    + ", sendPending="
                    + _sendPending.ToString(CultureInfo.InvariantCulture)
                    + ", cleanupArmedCount="
                    + _cleanupEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void OnClose(ref bool cancel)
            {
                if (_disposed)
                {
                    return;
                }

                ComposeSendState sendState = EvaluateMailSendState();
                bool hasPendingPostSendWork = _cleanupEntries.Count > 0 || _passwordDispatchQueue.Count > 0;
                int delayMs = 0;
                if (_sendPending)
                {
                    double elapsedMs = (DateTime.UtcNow - _sendPendingAtUtc).TotalMilliseconds;
                    delayMs = (int)Math.Max(0, ComposeShareCleanupSendGraceMs - elapsedMs);
                }
                if (sendState == ComposeSendState.Sent)
                {
                    ClearShareCleanupEntries("after_send_success");
                    DispatchSeparatePasswordQueue("after_send_success");
                    Dispose();
                    return;
                }
                if (sendState == ComposeSendState.UnavailableAfterSend)
                {
                    if (hasPendingPostSendWork && delayMs > 0)
                    {
                        _awaitingGraceCloseResolution = true;
                        _cleanupGraceTimer.Interval = Math.Max(250, delayMs);
                        _cleanupGraceTimer.Start();

                        LogFileLink(
                            "Compose share cleanup delayed (composeKey="
                            + _composeKey
                            + ", delayMs="
                            + delayMs.ToString(CultureInfo.InvariantCulture)
                            + ", reason=close_send_state_unavailable).");
                        return;
                    }

                    ClearShareCleanupEntries("after_send_state_unavailable");
                    DispatchSeparatePasswordQueue("after_send_state_unavailable");
                    Dispose();
                    return;
                }
                if (hasPendingPostSendWork && _sendPending && delayMs > 0)
                {
                    _awaitingGraceCloseResolution = true;
                    _cleanupGraceTimer.Interval = Math.Max(250, delayMs);
                    _cleanupGraceTimer.Start();

                    LogFileLink(
                        "Compose share cleanup delayed (composeKey="
                        + _composeKey
                        + ", delayMs="
                        + delayMs.ToString(CultureInfo.InvariantCulture)
                        + ", reason=close_send_pending).");
                    return;
                }
                if (hasPendingPostSendWork && _sendPending)
                {
                    LogFileLink(
                        "Compose send state not confirmed after grace; applying unsent cleanup path (composeKey="
                        + _composeKey
                        + ", reason=close_send_pending_timeout).");
                    DeleteShareCleanupEntries("close_send_pending_timeout_without_successful_send");
                    ClearSeparatePasswordDispatchQueue("close_send_pending_timeout_without_successful_send");
                    Dispose();
                    return;
                }
                if (_cleanupEntries.Count > 0)
                {
                    DeleteShareCleanupEntries("close_without_successful_send");
                }
                if (_passwordDispatchQueue.Count > 0)
                {
                    ClearSeparatePasswordDispatchQueue("close_without_successful_send");
                }

                Dispose();
            }

            private void OnCleanupGraceTimerTick(object sender, EventArgs e)
            {
                _cleanupGraceTimer.Stop();

                ComposeSendState sendState = EvaluateMailSendState();
                if (sendState == ComposeSendState.Sent || sendState == ComposeSendState.UnavailableAfterSend)
                {
                    string reason = sendState == ComposeSendState.Sent
                        ? "delayed_after_send_success"
                        : "delayed_after_send_state_unavailable";
                    ClearShareCleanupEntries(reason);
                    DispatchSeparatePasswordQueue(reason);
                }
                else
                {
                    if (_sendPending)
                    {
                        LogFileLink(
                            "Compose send state not confirmed after delayed grace; applying unsent cleanup path (composeKey="
                            + _composeKey
                            + ", reason=delayed_send_pending_timeout).");
                        DeleteShareCleanupEntries("delayed_send_pending_timeout_without_successful_send");
                        ClearSeparatePasswordDispatchQueue("delayed_send_pending_timeout_without_successful_send");
                    }
                    else
                    {
                        DeleteShareCleanupEntries("delayed_close_without_successful_send");
                        ClearSeparatePasswordDispatchQueue("delayed_close_without_successful_send");
                    }
                }

                Dispose();
            }

            private enum ComposeSendState
            {
                NotSent,
                Sent,
                UnavailableAfterSend
            }

            private ComposeSendState EvaluateMailSendState()
            {
                try
                {
                    return _mail != null && _mail.Sent ? ComposeSendState.Sent : ComposeSendState.NotSent;
                }
                catch (Exception ex)
                {
                    if (_sendPending && IsMailSentUnavailableAfterSend(ex))
                    {
                        LogFileLink(
                            "Compose send state unavailable after send (composeKey="
                            + _composeKey
                            + ", hresult="
                            + ToHResultHex(ex)
                            + ").");
                        return ComposeSendState.UnavailableAfterSend;
                    }

                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read MailItem.Sent (composeKey=" + _composeKey + ").",
                        ex);
                    return ComposeSendState.NotSent;
                }
            }

            private static bool IsMailSentUnavailableAfterSend(Exception ex)
            {
                var comException = ex as COMException;                if (comException == null)
                {
                    return false;
                }

                uint errorCode = unchecked((uint)comException.ErrorCode);
                return (errorCode & 0xFFFFu) == 0x010Au;
            }

            private static string ToHResultHex(Exception ex)
            {                if (ex == null)
                {
                    return "0x00000000";
                }
                return "0x" + unchecked((uint)ex.HResult).ToString("X8", CultureInfo.InvariantCulture);
            }

            private void ClearShareCleanupEntries(string reason)
            {
                if (_cleanupEntries.Count == 0)
                {
                    return;
                }
                for (int i = 0; i < _cleanupEntries.Count; i++)
                {
                    ComposeShareCleanupEntry entry = _cleanupEntries[i];
                    LogFileLink(
                        "Compose share cleanup cleared (composeKey="
                        + _composeKey
                        + ", reason="
                        + (reason ?? string.Empty)
                        + ", relativeFolder="
                        + (entry.RelativeFolder ?? string.Empty)
                        + ", shareId="
                        + (entry.ShareId ?? string.Empty)
                        + ", shareLabel="
                        + (entry.ShareLabel ?? string.Empty)
                        + ").");
                }

                _cleanupEntries.Clear();
                _sendPending = false;
                _awaitingGraceCloseResolution = false;
            }

            private void DeleteShareCleanupEntries(string reason)
            {
                if (_cleanupEntries.Count == 0)
                {
                    return;
                }

                List<ComposeShareCleanupEntry> entries = new List<ComposeShareCleanupEntry>(_cleanupEntries);
                _cleanupEntries.Clear();
                _sendPending = false;
                _awaitingGraceCloseResolution = false;

                for (int i = 0; i < entries.Count; i++)
                {
                    ComposeShareCleanupEntry entry = entries[i];
                    _owner._composeShareLifecycleController.TryDeleteComposeShareFolder(
                        entry.RelativeFolder,
                        reason,
                        entry.ShareId,
                        entry.ShareLabel);
                }
            }

            private void ClearSeparatePasswordDispatchQueue(string reason)
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }

                _passwordDispatchQueue.Clear();
                LogFileLink(
                    "Separate password dispatch cleared (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ").");
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
                bool capturedFromRecipients = TryCaptureRecipientListsFromRecipientsCollection(out to, out cc, out bcc);
                if (!capturedFromRecipients)
                {
                    to = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("To"));
                    cc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("CC"));
                    bcc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("BCC"));
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
                    + CountRecipients(to).ToString(CultureInfo.InvariantCulture)
                    + ", cc="
                    + CountRecipients(cc).ToString(CultureInfo.InvariantCulture)
                    + ", bcc="
                    + CountRecipients(bcc).ToString(CultureInfo.InvariantCulture)
                    + ", source="
                    + (capturedFromRecipients ? "recipients_collection" : "mail_fields")
                    + ").");
            }

            private bool TryCaptureRecipientListsFromRecipientsCollection(out string to, out string cc, out string bcc)
            {
                to = string.Empty;
                cc = string.Empty;
                bcc = string.Empty;                if (_mail == null)
                {
                    return false;
                }
                var toRecipients = new List<string>();
                var ccRecipients = new List<string>();
                var bccRecipients = new List<string>();
                Outlook.Recipients recipients = null;
                try
                {
                    recipients = _mail.Recipients;                    if (recipients == null)
                    {
                        return false;
                    }
                    int count = 0;
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
                            recipient = recipients[i];                            if (recipient == null)
                            {
                                continue;
                            }
                            string address = TryGetRecipientSmtpAddress(recipient);
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                try
                                {
                                    address = ComposeShareLifecycleController.NormalizeRecipientAddress(recipient.Address);
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
                            int recipientType = 1;
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
                                recipientType = 1;
                            }
                            if (recipientType == (int)Outlook.OlMailRecipientType.olCC)
                            {
                                ComposeShareLifecycleController.AddUniqueRecipient(ccRecipients, address);
                            }
                            else if (recipientType == (int)Outlook.OlMailRecipientType.olBCC)
                            {
                                ComposeShareLifecycleController.AddUniqueRecipient(bccRecipients, address);
                            }
                            else
                            {
                                ComposeShareLifecycleController.AddUniqueRecipient(toRecipients, address);
                            }
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(recipient, LogCategories.FileLink, "Failed to release compose Recipient COM object.");
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
                    ComInteropScope.TryRelease(recipients, LogCategories.FileLink, "Failed to release compose Recipients COM object.");
                }

                to = toRecipients.Count == 0 ? string.Empty : string.Join("; ", toRecipients.ToArray());
                cc = ccRecipients.Count == 0 ? string.Empty : string.Join("; ", ccRecipients.ToArray());
                bcc = bccRecipients.Count == 0 ? string.Empty : string.Join("; ", bccRecipients.ToArray());
                return toRecipients.Count + ccRecipients.Count + bccRecipients.Count > 0;
            }

            private void DispatchSeparatePasswordQueue(string reason)
            {
                if (_passwordDispatchQueue.Count == 0)
                {
                    return;
                }
                var queue = new List<SeparatePasswordDispatchEntry>(_passwordDispatchQueue);
                _passwordDispatchQueue.Clear();
                LogFileLink(
                    "Separate password dispatch taken (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", queued="
                    + queue.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                _owner._composeShareLifecycleController.DispatchSeparatePasswordMailQueue(_composeKey, queue);
            }

            private string ReadMailRecipientList(string fieldName)
            {
                try
                {                    if (_mail == null)
                    {
                        return string.Empty;
                    }

                    switch ((fieldName ?? string.Empty).Trim().ToUpperInvariant())
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

            private static int CountRecipients(string csv)
            {
                return ComposeShareLifecycleController.CountRecipientsInCsv(csv);
            }

        }
    }
}

