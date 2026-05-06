// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private const string ManagedSignatureStartPrefix = "<!-- nc-connector-signature:start hash=";
            private const string ManagedSignatureEnd = "<!-- nc-connector-signature:end -->";

            private static readonly Regex ManagedSignatureBlockRegex = new Regex(
                @"<!--\s*nc-connector-signature:start\b[^>]*-->.*?<!--\s*nc-connector-signature:end\s*-->",
                RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

            private enum EmailSignatureComposeKind
            {
                Unknown,
                New,
                Reply,
                Forward
            }

            private void ScheduleEmailSignatureApplication(string reason)
            {
                if (_disposed)
                {
                    return;
                }
                try
                {
                    _pendingEmailSignatureReason = string.IsNullOrWhiteSpace(reason) ? "scheduled" : reason.Trim();
                    _emailSignatureTimer.Stop();
                    _emailSignatureTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to schedule email signature processing (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void OnEmailSignatureTimerTick(object sender, EventArgs e)
            {
                _emailSignatureTimer.Stop();
                if (_disposed || _emailSignatureApplying)
                {
                    return;
                }

                _emailSignatureApplying = true;
                try
                {
                    ApplyEmailSignaturePolicy(_pendingEmailSignatureReason);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Email signature processing failed (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    _emailSignatureApplying = false;
                }
            }

            private void ApplyEmailSignaturePolicy(string reason)
            {
                if (_owner == null || _mail == null)
                {
                    return;
                }
                if (!IsMailComposeCandidate(_mail, "email_signature_" + (reason ?? "scheduled")))
                {
                    LogEmailSignature("Email signature skipped for non-compose mail (trigger=" + (reason ?? "n/a") + ").");
                    return;
                }

                _owner.EnsureSettingsLoaded();
                AddinSettings settings = _owner._currentSettings ?? new AddinSettings();
                var configuration = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
                BackendPolicyStatus policyStatus = _owner.FetchBackendPolicyStatus(configuration, "compose_email_signature");
                var policy = new EmailSignaturePolicyService(policyStatus, settings).Resolve();

                if (!policy.Active)
                {
                    ClearManagedEmailSignatureIfNeeded(policy.Reason);
                    LogEmailSignature("Email signature inactive (reason=" + (policy.Reason ?? "n/a") + ", trigger=" + (reason ?? "n/a") + ").");
                    return;
                }

                string senderEmail = EmailSignaturePolicyService.NormalizeEmail(ResolveCurrentSenderEmail());
                if (!string.Equals(senderEmail, policy.UserEmail, StringComparison.OrdinalIgnoreCase))
                {
                    ClearManagedEmailSignatureIfNeeded("identity_mismatch");
                    LogEmailSignature(
                        "Email signature skipped for non-seat sender (trigger="
                        + (reason ?? "n/a")
                        + ", hasSender="
                        + (!string.IsNullOrWhiteSpace(senderEmail)).ToString(CultureInfo.InvariantCulture)
                        + ", hasPolicyEmail="
                        + (!string.IsNullOrWhiteSpace(policy.UserEmail)).ToString(CultureInfo.InvariantCulture)
                        + ").");
                    return;
                }

                EmailSignatureComposeKind composeKind = ResolveEmailSignatureComposeKind();
                bool shouldInsert = ShouldInsertEmailSignature(policy, composeKind);
                bool shouldClearForeign = ShouldClearForeignEmailSignature(policy, composeKind);
                if (!shouldInsert)
                {
                    if (shouldClearForeign)
                    {
                        ClearEmailSignatureSlotWithoutInsert("compose_type_disabled", composeKind);
                    }
                    else
                    {
                        ClearManagedEmailSignatureIfNeeded("compose_type_disabled");
                    }
                    LogEmailSignature(
                        "Email signature skipped for compose kind (trigger="
                        + (reason ?? "n/a")
                        + ", kind="
                        + composeKind
                        + ", clearForeign="
                        + shouldClearForeign.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    return;
                }

                string sanitized = HtmlTemplateSanitizer.SanitizeEmailSignatureTemplateHtml(policy.TemplateHtml);
                if (string.IsNullOrWhiteSpace(sanitized))
                {
                    ClearManagedEmailSignatureIfNeeded("sanitized_empty");
                    LogEmailSignature("Email signature sanitized to empty output (trigger=" + (reason ?? "n/a") + ").");
                    return;
                }

                ApplyManagedEmailSignature(sanitized, composeKind, reason);
            }

            private void ApplyManagedEmailSignature(string sanitizedHtml, EmailSignatureComposeKind composeKind, string reason)
            {
                if (!EnsureHtmlBodyForEmailSignature("apply"))
                {
                    return;
                }

                string existing = _mail.HTMLBody ?? string.Empty;
                int slotStart;
                int slotEnd;
                bool hasQuoteBoundary;
                TryGetComposeSignatureSlotBounds(existing, out slotStart, out slotEnd, out hasQuoteBoundary);
                string managedBlock = BuildManagedSignatureBlock(sanitizedHtml, hasQuoteBoundary);
                string inlineManagedBlock = _isInlineResponse
                    ? BuildManagedSignatureBlock(sanitizedHtml, false)
                    : managedBlock;
                bool replacedManaged;
                string updated = ReplaceManagedSignatureBlocks(existing, managedBlock, out replacedManaged);
                bool removedInitialSlot = false;
                if (!replacedManaged)
                {
                    bool removedManaged;
                    string bodyWithoutManaged = RemoveManagedSignatureBlocks(updated, out removedManaged);
                    updated = ReplaceComposeSignatureSlot(
                        bodyWithoutManaged,
                        managedBlock,
                        !_emailSignatureInitialSlotHandled,
                        composeKind,
                        out removedInitialSlot);
                    replacedManaged = removedManaged;
                }
                if (!TryWriteEmailSignatureHtmlBody(updated, "apply", inlineManagedBlock))
                {
                    return;
                }
                _emailSignatureInitialSlotHandled = true;
                _emailSignatureManaged = true;

                LogEmailSignature(
                    "Email signature applied (trigger="
                    + (reason ?? "n/a")
                    + ", kind="
                    + composeKind
                    + ", htmlLength="
                    + sanitizedHtml.Length.ToString(CultureInfo.InvariantCulture)
                    + ", removedManaged="
                    + replacedManaged.ToString(CultureInfo.InvariantCulture)
                    + ", removedInitialSlot="
                    + removedInitialSlot.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void ClearEmailSignatureSlotWithoutInsert(string reason, EmailSignatureComposeKind composeKind)
            {
                if (_mail == null)
                {
                    return;
                }

                try
                {
                    Outlook.OlBodyFormat bodyFormat = ReadMailBodyFormatForEmailSignature("clear");
                    if (bodyFormat == Outlook.OlBodyFormat.olFormatPlain)
                    {
                        ClearManagedEmailSignatureIfNeeded(reason);
                        LogEmailSignature(
                            "Email signature slot cleanup skipped in plain-text compose (reason="
                            + (reason ?? "n/a")
                            + ", kind="
                            + composeKind
                            + ").");
                        return;
                    }

                    if (!EnsureHtmlBodyForEmailSignature("clear"))
                    {
                        return;
                    }

                    string existing = _mail.HTMLBody ?? string.Empty;
                    bool removedManaged;
                    string cleaned = RemoveManagedSignatureBlocks(existing, out removedManaged);
                    bool removedInitialSlot = false;
                    if (!removedManaged && !_emailSignatureInitialSlotHandled)
                    {
                        cleaned = ReplaceComposeSignatureSlot(
                            cleaned,
                            string.Empty,
                            true,
                            composeKind,
                            out removedInitialSlot);
                    }
                    if (removedManaged || removedInitialSlot)
                    {
                        if (!TryWriteEmailSignatureHtmlBody(cleaned, "clear_slot", string.Empty))
                        {
                            return;
                        }
                    }
                    _emailSignatureInitialSlotHandled = true;
                    _emailSignatureManaged = false;

                    LogEmailSignature(
                        "Email signature slot cleared without backend insert (reason="
                        + (reason ?? "n/a")
                        + ", kind="
                        + composeKind
                        + ", removedManaged="
                        + removedManaged.ToString(CultureInfo.InvariantCulture)
                        + ", removedInitialSlot="
                        + removedInitialSlot.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to clear email signature slot.", ex);
                }
            }

            private void ClearManagedEmailSignatureIfNeeded(string reason)
            {
                if (_mail == null)
                {
                    return;
                }
                try
                {
                    if (_mail.BodyFormat == Outlook.OlBodyFormat.olFormatPlain)
                    {
                        return;
                    }

                    string existing = _mail.HTMLBody ?? string.Empty;
                    if (!_emailSignatureManaged && !ManagedSignatureBlockRegex.IsMatch(existing))
                    {
                        return;
                    }

                    bool removed;
                    string cleaned = RemoveManagedSignatureBlocks(existing, out removed);
                    if (removed)
                    {
                        if (!TryWriteEmailSignatureHtmlBody(cleaned, "clear_managed", string.Empty))
                        {
                            return;
                        }
                    }
                    _emailSignatureManaged = false;
                    LogEmailSignature(
                        "Managed email signature cleared (reason="
                        + (reason ?? "n/a")
                        + ", changed="
                        + removed.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to clear managed email signature.", ex);
                }
            }

            private Outlook.OlBodyFormat ReadMailBodyFormatForEmailSignature(string operation)
            {
                try
                {
                    return _mail.BodyFormat;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read mail body format for email signature (" + (operation ?? "n/a") + ").", ex);
                    return Outlook.OlBodyFormat.olFormatUnspecified;
                }
            }

            private bool EnsureHtmlBodyForEmailSignature(string operation)
            {
                Outlook.OlBodyFormat bodyFormat = ReadMailBodyFormatForEmailSignature(operation);
                if (bodyFormat != Outlook.OlBodyFormat.olFormatPlain
                    && bodyFormat != Outlook.OlBodyFormat.olFormatRichText)
                {
                    return true;
                }

                try
                {
                    _mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                    return true;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to switch mail to HTML for email signature (" + (operation ?? "n/a") + ").", ex);
                    return false;
                }
            }

            private static string RemoveManagedSignatureBlocks(string html, out bool removed)
            {
                removed = false;
                if (string.IsNullOrEmpty(html))
                {
                    return html ?? string.Empty;
                }
                removed = ManagedSignatureBlockRegex.IsMatch(html);
                return removed ? ManagedSignatureBlockRegex.Replace(html, string.Empty) : html;
            }

            private bool TryWriteEmailSignatureHtmlBody(string html, string operation, string inlineSignatureSlotHtml = null)
            {
                if (!_isInlineResponse)
                {
                    _mail.HTMLBody = html ?? string.Empty;
                    return true;
                }

                if (_owner == null || _owner._mailInteropController == null)
                {
                    LogEmailSignature("Inline email signature write skipped (operation=" + (operation ?? "n/a") + ", reason=interop_unavailable).");
                    return false;
                }

                bool written = _owner._mailInteropController.TryReplaceActiveInlineResponseSignatureSlot(
                    _mail,
                    inlineSignatureSlotHtml ?? html ?? string.Empty,
                    _composeKey,
                    operation);
                if (!written)
                {
                    LogEmailSignature("Inline email signature write skipped (operation=" + (operation ?? "n/a") + ", reason=inline_editor_unavailable).");
                }
                return written;
            }

            private static string BuildManagedSignatureBlock(string sanitizedHtml, bool addTrailingQuoteGap)
            {
                string html = sanitizedHtml ?? string.Empty;
                string hash = ComputeSignatureHash(html);
                return ManagedSignatureStartPrefix
                       + hash
                       + " -->"
                       + "<p style=\"margin:0;\"><br /></p>"
                       + "<p style=\"margin:0;\"><br /></p>"
                       + "<div data-nc-connector-signature=\"true\">"
                       + html
                       + "</div>"
                       + (addTrailingQuoteGap ? "<p style=\"margin:0;\"><br /></p>" : string.Empty)
                       + ManagedSignatureEnd;
            }

            private static string ReplaceManagedSignatureBlocks(string html, string replacement, out bool replaced)
            {
                replaced = false;
                if (string.IsNullOrEmpty(html))
                {
                    return html ?? string.Empty;
                }
                replaced = ManagedSignatureBlockRegex.IsMatch(html);
                if (!replaced)
                {
                    return html;
                }
                return ManagedSignatureBlockRegex.Replace(html, replacement ?? string.Empty);
            }

            private static string ReplaceComposeSignatureSlot(
                string html,
                string replacement,
                bool mayReplaceForeignSlot,
                EmailSignatureComposeKind composeKind,
                out bool removedForeignSlot)
            {
                removedForeignSlot = false;
                string existing = html ?? string.Empty;
                int slotStart;
                int slotEnd;
                bool hasQuoteBoundary;
                if (!TryGetComposeSignatureSlotBounds(existing, out slotStart, out slotEnd, out hasQuoteBoundary))
                {
                    return (replacement ?? string.Empty) + existing;
                }

                string slot = slotEnd > slotStart ? existing.Substring(slotStart, slotEnd - slotStart) : string.Empty;
                bool slotHasText = !string.IsNullOrWhiteSpace(StripHtmlToText(slot));
                bool canReplaceForeignSlot = mayReplaceForeignSlot
                    && (composeKind == EmailSignatureComposeKind.New || hasQuoteBoundary);
                if (slotHasText && !canReplaceForeignSlot)
                {
                    return existing.Insert(slotStart, replacement ?? string.Empty);
                }

                removedForeignSlot = slotHasText && canReplaceForeignSlot;
                return existing.Substring(0, slotStart)
                       + (replacement ?? string.Empty)
                       + existing.Substring(slotEnd);
            }

            private static bool TryGetComposeSignatureSlotBounds(string html, out int slotStart, out int slotEnd, out bool hasQuoteBoundary)
            {
                slotStart = 0;
                slotEnd = 0;
                hasQuoteBoundary = false;
                if (string.IsNullOrEmpty(html))
                {
                    return false;
                }

                int bodyStart = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyStart >= 0)
                {
                    int bodyTagEnd = html.IndexOf(">", bodyStart);
                    if (bodyTagEnd >= 0)
                    {
                        slotStart = bodyTagEnd + 1;
                    }
                }

                int bodyEnd = html.IndexOf("</body>", slotStart, StringComparison.OrdinalIgnoreCase);
                if (bodyEnd < 0)
                {
                    bodyEnd = html.Length;
                }

                int quoteStart = FindQuoteBoundary(html, slotStart, bodyEnd);
                slotEnd = quoteStart >= 0 ? quoteStart : bodyEnd;
                hasQuoteBoundary = quoteStart >= 0;
                if (slotEnd < slotStart)
                {
                    slotEnd = slotStart;
                }
                return true;
            }

            private static int FindQuoteBoundary(string html, int start, int end)
            {
                string[] structuralMarkers = new[]
                {
                    "id=\"divRplyFwdMsg\"",
                    "id=divRplyFwdMsg",
                    "id=\"x_divRplyFwdMsg\"",
                    "id=x_divRplyFwdMsg",
                    "border:none;border-top",
                    "border: none; border-top",
                    "mso-border-top-alt",
                    "<blockquote",
                    "<hr"
                };
                int structuralBoundary = FindEarliestBoundaryMarker(html, start, end, structuralMarkers);
                if (structuralBoundary >= 0)
                {
                    return NormalizeBoundaryToElementStart(html, start, structuralBoundary);
                }

                string[] textMarkers = new[]
                {
                    "-----Original Message-----",
                    "Von:",
                    "From:"
                };
                int textBoundary = FindEarliestBoundaryMarker(html, start, end, textMarkers);
                if (textBoundary < 0)
                {
                    return -1;
                }

                int headerElementStart = FindReplyHeaderElementStart(html, start, textBoundary);
                if (headerElementStart >= 0)
                {
                    return headerElementStart;
                }

                // Keep visual separation intact when Outlook places a divider
                // immediately before the localized header text marker.
                int hrBeforeText = html.LastIndexOf("<hr", textBoundary, StringComparison.OrdinalIgnoreCase);
                if (hrBeforeText >= start && hrBeforeText < textBoundary && (textBoundary - hrBeforeText) <= 512)
                {
                    return hrBeforeText;
                }
                return -1;
            }

            private static int NormalizeBoundaryToElementStart(string html, int start, int markerIndex)
            {
                if (string.IsNullOrEmpty(html) || markerIndex < start)
                {
                    return markerIndex;
                }

                if (markerIndex + 1 < html.Length && html[markerIndex] == '<')
                {
                    return markerIndex;
                }

                int tagStart = html.LastIndexOf("<", markerIndex, StringComparison.OrdinalIgnoreCase);
                int tagEnd = html.IndexOf(">", tagStart >= 0 ? tagStart : markerIndex, StringComparison.OrdinalIgnoreCase);
                if (tagStart >= start && tagEnd >= markerIndex)
                {
                    return tagStart;
                }
                return markerIndex;
            }

            private static int FindReplyHeaderElementStart(string html, int start, int textBoundary)
            {
                int openBlock = FindOpenAncestorTagStart(html, start, textBoundary, "div");
                if (openBlock >= 0)
                {
                    return openBlock;
                }

                openBlock = FindOpenAncestorTagStart(html, start, textBoundary, "table");
                if (openBlock >= 0)
                {
                    return openBlock;
                }

                int paragraphStart = html.LastIndexOf("<p", textBoundary, StringComparison.OrdinalIgnoreCase);
                if (paragraphStart >= start)
                {
                    int paragraphEnd = html.IndexOf(">", paragraphStart, StringComparison.OrdinalIgnoreCase);
                    if (paragraphEnd >= textBoundary)
                    {
                        return paragraphStart;
                    }
                    int previousParagraphClose = html.LastIndexOf("</p>", textBoundary, StringComparison.OrdinalIgnoreCase);
                    if (previousParagraphClose < paragraphStart)
                    {
                        return paragraphStart;
                    }
                }

                return -1;
            }

            private static int FindOpenAncestorTagStart(string html, int start, int textBoundary, string tagName)
            {
                if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(tagName) || textBoundary <= start)
                {
                    return -1;
                }

                string openMarker = "<" + tagName;
                string closeMarker = "</" + tagName + ">";
                int open = html.LastIndexOf(openMarker, textBoundary, StringComparison.OrdinalIgnoreCase);
                if (open < start)
                {
                    return -1;
                }

                int close = html.LastIndexOf(closeMarker, textBoundary, StringComparison.OrdinalIgnoreCase);
                if (close > open)
                {
                    return -1;
                }

                int tagEnd = html.IndexOf(">", open, StringComparison.OrdinalIgnoreCase);
                if (tagEnd < 0 || tagEnd > textBoundary)
                {
                    return -1;
                }

                return open;
            }

            private static int FindEarliestBoundaryMarker(string html, int start, int end, string[] markers)
            {
                if (string.IsNullOrEmpty(html) || markers == null || markers.Length == 0)
                {
                    return -1;
                }

                int result = -1;
                for (int i = 0; i < markers.Length; i++)
                {
                    string marker = markers[i];
                    if (string.IsNullOrEmpty(marker))
                    {
                        continue;
                    }
                    int index = html.IndexOf(marker, start, StringComparison.OrdinalIgnoreCase);
                    if (index < 0 || index >= end)
                    {
                        continue;
                    }
                    if (result < 0 || index < result)
                    {
                        result = index;
                    }
                }
                return result;
            }

            private static string StripHtmlToText(string html)
            {
                string text = Regex.Replace(html ?? string.Empty, @"<\s*br\s*/?\s*>", "\n", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, "<[^>]+>", string.Empty, RegexOptions.IgnoreCase);
                text = HttpUtility.HtmlDecode(text) ?? string.Empty;
                return text.Replace('\u00A0', ' ').Trim();
            }

            private EmailSignatureComposeKind ResolveEmailSignatureComposeKind()
            {
                int verb;
                if (!TryReadLastVerbExecuted(out verb))
                {
                    return EmailSignatureComposeKind.Unknown;
                }
                if (verb == 102 || verb == 103)
                {
                    return EmailSignatureComposeKind.Reply;
                }
                if (verb == 104)
                {
                    return EmailSignatureComposeKind.Forward;
                }
                return EmailSignatureComposeKind.New;
            }

            private bool TryReadLastVerbExecuted(out int verb)
            {
                verb = 0;
                Outlook.PropertyAccessor accessor = null;
                try
                {
                    accessor = _mail.PropertyAccessor;
                    object value = accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003");
                    if (value == null)
                    {
                        return true;
                    }
                    int parsed;
                    if (!int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                    {
                        return true;
                    }
                    verb = parsed;
                    return true;
                }
                catch (COMException ex)
                {
                    uint errorCode = unchecked((uint)ex.ErrorCode);
                    if (errorCode == 0x8004010Fu)
                    {
                        if (DiagnosticsLogger.IsEnabled)
                        {
                            LogEmailSignature("Last verb is not set; treating unsent compose as new mail.");
                        }
                        return true;
                    }
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read mail last verb for email signature compose kind.", ex);
                    return false;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read mail last verb for email signature compose kind.", ex);
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release mail PropertyAccessor COM object.");
                }
            }

            private static bool ShouldInsertEmailSignature(EmailSignaturePolicy policy, EmailSignatureComposeKind composeKind)
            {
                if (policy == null || !policy.OnCompose)
                {
                    return false;
                }
                if (composeKind == EmailSignatureComposeKind.New)
                {
                    return true;
                }
                if (composeKind == EmailSignatureComposeKind.Reply)
                {
                    return policy.OnReply;
                }
                if (composeKind == EmailSignatureComposeKind.Forward)
                {
                    return policy.OnForward;
                }
                return false;
            }

            private static bool ShouldClearForeignEmailSignature(EmailSignaturePolicy policy, EmailSignatureComposeKind composeKind)
            {
                if (policy == null || !policy.OnCompose)
                {
                    return false;
                }
                return composeKind == EmailSignatureComposeKind.New
                       || composeKind == EmailSignatureComposeKind.Reply
                       || composeKind == EmailSignatureComposeKind.Forward;
            }

            private string ResolveCurrentSenderEmail()
            {
                string sentOnBehalfOf;
                bool hasSentOnBehalfOf;
                if (TryResolveSentOnBehalfOfSmtpAddress(out sentOnBehalfOf, out hasSentOnBehalfOf))
                {
                    return sentOnBehalfOf;
                }
                if (hasSentOnBehalfOf)
                {
                    LogEmailSignature("Sender identity unresolved because SentOnBehalfOfName is set but not resolvable.");
                    return string.Empty;
                }

                return ResolveSendUsingAccountSmtpAddress();
            }

            private string ResolveSendUsingAccountSmtpAddress()
            {
                Outlook.Account account = null;
                try
                {
                    account = _mail.SendUsingAccount;
                    if (account != null)
                    {
                        string smtpAddress = account.SmtpAddress;
                        if (IsSmtpEmailCandidate(smtpAddress))
                        {
                            return smtpAddress.Trim();
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve compose sender account SMTP address.", ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(account, LogCategories.Core, "Failed to release compose sender account COM object.");
                }

                return string.Empty;
            }

            private bool TryResolveSentOnBehalfOfSmtpAddress(out string smtpAddress, out bool hasSentOnBehalfOf)
            {
                smtpAddress = string.Empty;
                hasSentOnBehalfOf = false;

                string sentOnBehalfOfName;
                try
                {
                    sentOnBehalfOfName = _mail.SentOnBehalfOfName;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read compose sent-on-behalf identity.", ex);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(sentOnBehalfOfName))
                {
                    return false;
                }

                hasSentOnBehalfOf = true;
                sentOnBehalfOfName = sentOnBehalfOfName.Trim();
                if (IsSmtpEmailCandidate(sentOnBehalfOfName))
                {
                    smtpAddress = sentOnBehalfOfName;
                    return true;
                }

                Outlook.NameSpace session = null;
                Outlook.Recipient recipient = null;
                try
                {
                    Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
                    if (application == null)
                    {
                        return false;
                    }

                    session = application.Session;
                    if (session == null)
                    {
                        return false;
                    }

                    recipient = session.CreateRecipient(sentOnBehalfOfName);
                    if (recipient == null || !recipient.Resolve())
                    {
                        return false;
                    }

                    string resolved = OutlookRecipientResolverController.TryResolveRecipientSmtpAddress(recipient);
                    if (!IsSmtpEmailCandidate(resolved))
                    {
                        return false;
                    }

                    smtpAddress = resolved.Trim();
                    return true;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve compose sent-on-behalf SMTP address.", ex);
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(recipient, LogCategories.Core, "Failed to release sent-on-behalf Recipient COM object.");
                    ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release sent-on-behalf Session COM object.");
                }
            }

            private static bool IsSmtpEmailCandidate(string value)
            {
                return !string.IsNullOrWhiteSpace(value)
                       && value.IndexOf("@", StringComparison.Ordinal) > 0;
            }

            private static bool IsEmailSignaturePropertyChange(string propertyName)
            {
                return string.Equals(propertyName, "SendUsingAccount", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(propertyName, "SenderEmailAddress", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(propertyName, "SentOnBehalfOfName", StringComparison.OrdinalIgnoreCase);
            }

            private static string ComputeSignatureHash(string value)
            {
                unchecked
                {
                    uint hash = 2166136261;
                    string text = value ?? string.Empty;
                    for (int i = 0; i < text.Length; i++)
                    {
                        hash ^= text[i];
                        hash *= 16777619;
                    }
                    return hash.ToString("x8", CultureInfo.InvariantCulture);
                }
            }

            private void LogEmailSignature(string message)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Email signature: "
                    + (message ?? string.Empty)
                    + " (composeKey="
                    + (_composeKey ?? string.Empty)
                    + ").");
            }
        }
    }
}
