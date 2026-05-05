// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Web;
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
                bool removedManaged;
                string bodyWithoutManaged = RemoveManagedSignatureBlocks(existing, out removedManaged);
                bool removedInitialSlot = false;
                if (!_emailSignatureInitialSlotHandled)
                {
                    bodyWithoutManaged = RemoveInitialSignatureSlot(bodyWithoutManaged, out removedInitialSlot);
                    _emailSignatureInitialSlotHandled = true;
                }

                string managedBlock = BuildManagedSignatureBlock(sanitizedHtml);
                _mail.HTMLBody = InsertAfterBodyStart(bodyWithoutManaged, managedBlock);
                _emailSignatureManaged = true;

                LogEmailSignature(
                    "Email signature applied (trigger="
                    + (reason ?? "n/a")
                    + ", kind="
                    + composeKind
                    + ", htmlLength="
                    + sanitizedHtml.Length.ToString(CultureInfo.InvariantCulture)
                    + ", removedManaged="
                    + removedManaged.ToString(CultureInfo.InvariantCulture)
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
                    if (!_emailSignatureInitialSlotHandled)
                    {
                        cleaned = RemoveInitialSignatureSlot(cleaned, out removedInitialSlot);
                        _emailSignatureInitialSlotHandled = true;
                    }

                    if (removedManaged || removedInitialSlot)
                    {
                        _mail.HTMLBody = cleaned;
                    }
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
                        _mail.HTMLBody = cleaned;
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

            private static string BuildManagedSignatureBlock(string sanitizedHtml)
            {
                string html = sanitizedHtml ?? string.Empty;
                string hash = ComputeSignatureHash(html);
                return ManagedSignatureStartPrefix
                       + hash
                       + " -->"
                       + "<div data-nc-connector-signature=\"true\">"
                       + html
                       + "</div>"
                       + ManagedSignatureEnd;
            }

            private static string InsertAfterBodyStart(string html, string managedBlock)
            {
                string existing = html ?? string.Empty;
                string insert = (managedBlock ?? string.Empty) + "<br />";
                int bodyTagIndex = existing.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyTagIndex >= 0)
                {
                    int bodyTagEnd = existing.IndexOf(">", bodyTagIndex);
                    if (bodyTagEnd >= 0)
                    {
                        return existing.Insert(bodyTagEnd + 1, insert);
                    }
                }
                return insert + existing;
            }

            private string RemoveInitialSignatureSlot(string html, out bool removed)
            {
                removed = false;
                if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(_initialEmailSignatureSlotHtml))
                {
                    return html ?? string.Empty;
                }

                int slotStart = html.IndexOf(_initialEmailSignatureSlotHtml, StringComparison.Ordinal);
                if (slotStart < 0)
                {
                    return html;
                }

                removed = true;
                return html.Substring(0, slotStart)
                       + html.Substring(slotStart + _initialEmailSignatureSlotHtml.Length);
            }

            private static string CaptureInitialSignatureSlotHtml(Outlook.MailItem mail)
            {
                if (mail == null)
                {
                    return string.Empty;
                }

                try
                {
                    Outlook.OlBodyFormat bodyFormat = mail.BodyFormat;
                    if (bodyFormat == Outlook.OlBodyFormat.olFormatPlain)
                    {
                        return string.Empty;
                    }
                    string html = mail.HTMLBody ?? string.Empty;
                    string slot = ExtractInitialSignatureSlot(html);
                    return string.IsNullOrWhiteSpace(StripHtmlToText(slot)) ? string.Empty : slot;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to capture initial Outlook signature slot.", ex);
                    return string.Empty;
                }
            }

            private static string ExtractInitialSignatureSlot(string html)
            {
                if (string.IsNullOrWhiteSpace(html))
                {
                    return string.Empty;
                }

                int bodyStart = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                int contentStart = 0;
                if (bodyStart >= 0)
                {
                    int bodyTagEnd = html.IndexOf(">", bodyStart);
                    if (bodyTagEnd >= 0)
                    {
                        contentStart = bodyTagEnd + 1;
                    }
                }

                int bodyEnd = html.IndexOf("</body>", contentStart, StringComparison.OrdinalIgnoreCase);
                if (bodyEnd < 0)
                {
                    bodyEnd = html.Length;
                }

                int quoteStart = FindQuoteBoundary(html, contentStart, bodyEnd);
                int contentEnd = quoteStart >= 0 ? quoteStart : bodyEnd;
                return contentEnd <= contentStart
                    ? string.Empty
                    : html.Substring(contentStart, contentEnd - contentStart);
            }

            private static int FindQuoteBoundary(string html, int start, int end)
            {
                string[] markers = new[]
                {
                    "id=\"divRplyFwdMsg\"",
                    "id=divRplyFwdMsg",
                    "<blockquote",
                    "<hr",
                    "-----Original Message-----",
                    "Von:",
                    "From:"
                };
                int result = -1;
                for (int i = 0; i < markers.Length; i++)
                {
                    int index = html.IndexOf(markers[i], start, StringComparison.OrdinalIgnoreCase);
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
                int verb = ReadLastVerbExecuted();
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

            private int ReadLastVerbExecuted()
            {
                Outlook.PropertyAccessor accessor = null;
                try
                {
                    accessor = _mail.PropertyAccessor;
                    object value = accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003");
                    if (value == null)
                    {
                        return 0;
                    }
                    int parsed;
                    return int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                        ? parsed
                        : 0;
                }
                catch
                {
                    return 0;
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
                Outlook.Account account = null;
                try
                {
                    account = _mail.SendUsingAccount;
                    if (account != null)
                    {
                        string smtpAddress = account.SmtpAddress;
                        if (!string.IsNullOrWhiteSpace(smtpAddress))
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

                try
                {
                    string sender = _mail.SenderEmailAddress;
                    if (!string.IsNullOrWhiteSpace(sender) && sender.IndexOf("@", StringComparison.Ordinal) >= 0)
                    {
                        return sender.Trim();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve compose sender email address.", ex);
                }

                try
                {
                    string sentOnBehalfOf = _mail.SentOnBehalfOfName;
                    return string.IsNullOrWhiteSpace(sentOnBehalfOf) || sentOnBehalfOf.IndexOf("@", StringComparison.Ordinal) < 0
                        ? string.Empty
                        : sentOnBehalfOf.Trim();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve compose sent-on-behalf address.", ex);
                    return string.Empty;
                }
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
