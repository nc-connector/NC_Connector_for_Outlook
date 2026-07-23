// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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
            private const int EmailSignatureReadyRetryDelayMs = 750;
            private const int EmailSignatureReadyRetryLimit = 4;

            private int _emailSignatureRequestGeneration;
            private int _emailSignatureReadyRetryCount;
            private bool _emailSignatureStateStable;

            private sealed class EmailSignatureApplicationResult
            {
                internal bool Success { get; set; }

                internal string Source { get; set; }
            }

            private enum EmailSignatureComposeKind
            {
                Unknown,
                New,
                Reply,
                Forward,
                Response
            }

            private void ScheduleEmailSignatureApplication(string reason)
            {
                if (_disposed)
                {
                    return;
                }

                string normalizedReason = string.IsNullOrWhiteSpace(reason)
                    ? "scheduled"
                    : reason.Trim();
                if (!normalizedReason.StartsWith("retry_", StringComparison.OrdinalIgnoreCase))
                {
                    _emailSignatureReadyRetryCount = 0;
                }

                _pendingEmailSignatureReason = normalizedReason;
                _emailSignatureStateStable = false;
                _emailSignatureRequestGeneration++;
                try
                {
                    _emailSignatureTimer.Stop();
                    if (_composeSurfaceState == ComposeSurfaceState.Detached)
                    {
                        DeferEmailSignatureApplication(normalizedReason, "schedule_detached");
                        return;
                    }
                    if (_emailSignatureApplying)
                    {
                        LogEmailSignature(
                            "reconcile queued while policy fetch is active (reason="
                            + normalizedReason
                            + ", generation="
                            + _emailSignatureRequestGeneration.ToString(CultureInfo.InvariantCulture)
                            + ").");
                        return;
                    }

                    ConfigureEmailSignatureTimer(normalizedReason);
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

            private void ConfigureEmailSignatureTimer(string reason)
            {
                _emailSignatureTimer.Interval = !string.IsNullOrWhiteSpace(reason)
                                                && reason.StartsWith("retry_", StringComparison.OrdinalIgnoreCase)
                    ? EmailSignatureReadyRetryDelayMs
                    : (_composeSurfaceState == ComposeSurfaceState.InlineResponse
                        ? EmailSignatureInlineApplyDebounceMs
                        : EmailSignatureApplyDebounceMs);
            }

            private async void OnEmailSignatureTimerTick(object sender, EventArgs e)
            {
                _emailSignatureTimer.Stop();
                if (_disposed)
                {
                    return;
                }
                if (_emailSignatureApplying)
                {
                    DeferEmailSignatureApplication(_pendingEmailSignatureReason, "timer_busy");
                    return;
                }

                _emailSignatureApplying = true;
                int generation = _emailSignatureRequestGeneration;
                string reason = string.IsNullOrWhiteSpace(_pendingEmailSignatureReason)
                    ? "scheduled"
                    : _pendingEmailSignatureReason;
                Exception processingException = null;
                try
                {
                    _owner.EnsureSettingsLoaded();
                    AddinSettings settings = _owner._currentSettings ?? new AddinSettings();
                    var configuration = new TalkServiceConfiguration(
                        settings.ServerUrl,
                        settings.Username,
                        settings.AppPassword);

                    BackendPolicyStatus policyStatus = null;
                    if (configuration.IsComplete())
                    {
                        policyStatus = await _owner.GetEmailSignaturePolicyStatusAsync(
                            configuration,
                            "compose_email_signature_" + reason).ConfigureAwait(false);
                    }

                    await _owner.RunOnOutlookUiThreadAsync(
                        () => CompleteEmailSignaturePolicyFetch(
                            generation,
                            reason,
                            settings,
                            configuration,
                            policyStatus)).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    processingException = ex;
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Email signature policy processing failed (composeKey=" + _composeKey + ").",
                        ex);
                }

                if (processingException != null)
                {
                    try
                    {
                        await _owner.RunOnOutlookUiThreadAsync(
                            () =>
                            {
                                if (!_disposed && generation == _emailSignatureRequestGeneration)
                                {
                                    ScheduleEmailSignatureRetry(reason, "policy_exception");
                                }
                            }).ConfigureAwait(false);
                    }
                    catch (Exception dispatchException)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to dispatch email signature retry to the Outlook UI thread.",
                            dispatchException);
                    }
                }

                try
                {
                    await _owner.RunOnOutlookUiThreadAsync(
                        () => CompleteEmailSignatureAsyncRun(generation)).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    _emailSignatureApplying = false;
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to complete email signature processing on the Outlook UI thread.",
                        ex);
                }
            }

            private void CompleteEmailSignaturePolicyFetch(
                int generation,
                string reason,
                AddinSettings settings,
                TalkServiceConfiguration configuration,
                BackendPolicyStatus policyStatus)
            {
                if (_disposed || generation != _emailSignatureRequestGeneration)
                {
                    LogEmailSignature(
                        "stale policy result ignored (generation="
                        + generation.ToString(CultureInfo.InvariantCulture)
                        + ", currentGeneration="
                        + _emailSignatureRequestGeneration.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    return;
                }

                EmailSignatureApplicationResult result = configuration != null && configuration.IsComplete()
                    ? ApplyEmailSignaturePolicy(policyStatus, settings, reason)
                    : ReconcileEmailSignatureWithoutBackendConfiguration(reason);
                if (result.Success)
                {
                    _emailSignatureStateStable = true;
                    _emailSignatureReadyRetryCount = 0;
                    _pendingEmailSignatureReason = string.Empty;
                    LogEmailSignature(
                        "reconcile completed (trigger="
                        + (reason ?? "n/a")
                        + ", source="
                        + (result.Source ?? "n/a")
                        + ", surface="
                        + _composeSurfaceState.ToString()
                        + ").");
                    return;
                }

                _emailSignatureStateStable = false;
                ScheduleEmailSignatureRetry(reason, result.Source);
            }

            private void CompleteEmailSignatureAsyncRun(int generation)
            {
                _emailSignatureApplying = false;
                if (_disposed)
                {
                    return;
                }

                if (ResumeDeferredEmailSignatureApplication("policy_complete"))
                {
                    return;
                }

                if (generation != _emailSignatureRequestGeneration
                    && _composeSurfaceState != ComposeSurfaceState.Detached
                    && !string.IsNullOrWhiteSpace(_pendingEmailSignatureReason)
                    && !_emailSignatureTimer.Enabled)
                {
                    ConfigureEmailSignatureTimer(_pendingEmailSignatureReason);
                    _emailSignatureTimer.Start();
                }
            }

            private void ScheduleEmailSignatureRetry(string originalReason, string source)
            {
                if (_disposed)
                {
                    return;
                }

                _emailSignatureReadyRetryCount++;
                if (_emailSignatureReadyRetryCount > EmailSignatureReadyRetryLimit)
                {
                    _pendingEmailSignatureReason = "retry_exhausted_"
                                                   + (string.IsNullOrWhiteSpace(originalReason)
                                                       ? "scheduled"
                                                       : originalReason);
                    LogEmailSignature(
                        "reconcile retry limit reached (trigger="
                        + (originalReason ?? "n/a")
                        + ", source="
                        + (source ?? "n/a")
                        + ", retries="
                        + EmailSignatureReadyRetryLimit.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    return;
                }

                string retryReason = "retry_"
                                     + _emailSignatureReadyRetryCount.ToString(CultureInfo.InvariantCulture)
                                     + "_"
                                     + (string.IsNullOrWhiteSpace(originalReason)
                                         ? "scheduled"
                                         : originalReason);
                LogEmailSignature(
                    "reconcile retry scheduled (trigger="
                    + (originalReason ?? "n/a")
                    + ", source="
                    + (source ?? "n/a")
                    + ", retry="
                    + _emailSignatureReadyRetryCount.ToString(CultureInfo.InvariantCulture)
                    + ").");
                ScheduleEmailSignatureApplication(retryReason);
            }

            private EmailSignatureApplicationResult ReconcileEmailSignatureWithoutBackendConfiguration(string reason)
            {
                return ClearManagedEmailSignature(
                    "configuration_incomplete:" + (reason ?? "n/a"));
            }

            private EmailSignatureApplicationResult ApplyEmailSignaturePolicy(
                BackendPolicyStatus policyStatus,
                AddinSettings settings,
                string reason)
            {
                if (_owner == null || _mail == null)
                {
                    return EmailSignatureFailure("compose_unavailable");
                }
                if (!IsMailComposeCandidate(_mail, "email_signature_" + (reason ?? "scheduled")))
                {
                    return EmailSignatureSuccess("not_compose");
                }
                if (policyStatus == null || !policyStatus.FetchSucceeded)
                {
                    LogEmailSignature(
                        "policy result unavailable; message body left unchanged (trigger="
                        + (reason ?? "n/a")
                        + ").");
                    return EmailSignatureFailure("policy_unavailable");
                }

                var policy = new EmailSignaturePolicyService(
                    policyStatus,
                    settings ?? new AddinSettings()).Resolve();
                if (!policy.Active)
                {
                    return ClearManagedEmailSignature(
                        "policy_inactive:" + (policy.Reason ?? "n/a"));
                }

                string senderEmail = EmailSignaturePolicyService.NormalizeEmail(
                    ResolveCurrentSenderEmail());
                if (!string.Equals(
                        senderEmail,
                        policy.UserEmail,
                        StringComparison.OrdinalIgnoreCase))
                {
                    LogEmailSignature(
                        "sender does not match backend seat (trigger="
                        + (reason ?? "n/a")
                        + ", hasSender="
                        + (!string.IsNullOrWhiteSpace(senderEmail)).ToString(CultureInfo.InvariantCulture)
                        + ", hasPolicyEmail="
                        + (!string.IsNullOrWhiteSpace(policy.UserEmail)).ToString(CultureInfo.InvariantCulture)
                        + ").");
                    return ClearManagedEmailSignature("identity_mismatch");
                }

                EmailSignatureComposeKind composeKind = ResolveEmailSignatureComposeKind();
                if (composeKind == EmailSignatureComposeKind.Unknown)
                {
                    LogEmailSignature("compose kind is not yet reliable; body left unchanged.");
                    return EmailSignatureFailure("compose_kind_unknown");
                }
                if (composeKind == EmailSignatureComposeKind.Response
                    && policy.OnReply != policy.OnForward)
                {
                    LogEmailSignature(
                        "ambiguous response not changed because reply/forward policy differs.");
                    return EmailSignatureFailure("response_kind_ambiguous");
                }

                bool shouldInsert = ShouldInsertEmailSignature(policy, composeKind);
                LogEmailSignature(
                    "decision (trigger="
                    + (reason ?? "n/a")
                    + ", kind="
                    + composeKind.ToString()
                    + ", onCompose="
                    + policy.OnCompose.ToString(CultureInfo.InvariantCulture)
                    + ", onReply="
                    + policy.OnReply.ToString(CultureInfo.InvariantCulture)
                    + ", onForward="
                    + policy.OnForward.ToString(CultureInfo.InvariantCulture)
                    + ", insert="
                    + shouldInsert.ToString(CultureInfo.InvariantCulture)
                    + ", surface="
                    + _composeSurfaceState.ToString()
                    + ").");
                if (!shouldInsert)
                {
                    return ClearInitialEmailSignatureSlot(
                        "compose_type_disabled:" + (reason ?? "n/a"));
                }

                string sanitized = HtmlTemplateSanitizer.SanitizeEmailSignatureTemplateHtml(
                    policy.TemplateHtml);
                if (string.IsNullOrWhiteSpace(sanitized))
                {
                    EmailSignatureApplicationResult clearResult =
                        ClearManagedEmailSignature("sanitized_empty");
                    return EmailSignatureFailure(
                        clearResult.Success
                            ? "sanitized_empty"
                            : "sanitized_empty_clear_failed:" + (clearResult.Source ?? "n/a"));
                }

                Outlook.OlBodyFormat bodyFormat;
                try
                {
                    bodyFormat = _mail.BodyFormat;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to read compose body format for email signature.",
                        ex);
                    return EmailSignatureFailure("body_format_unavailable");
                }

                bool isPlainText = bodyFormat == Outlook.OlBodyFormat.olFormatPlain;
                LogEmailSignature(
                    "format resolved (trigger="
                    + (reason ?? "n/a")
                    + ", bodyFormat="
                    + bodyFormat.ToString()
                    + ", plainText="
                    + isPlainText.ToString(CultureInfo.InvariantCulture)
                    + ").");
                string signatureContent = isPlainText
                    ? EmailSignatureContentBuilder.BuildPlainText(sanitized)
                    : EmailSignatureContentBuilder.BuildManagedHtml(sanitized);
                if (string.IsNullOrWhiteSpace(signatureContent))
                {
                    return EmailSignatureFailure("signature_content_empty");
                }

                MailInteropController.EmailSignatureReconcileResult reconcile =
                    _owner._mailInteropController.ApplyManagedEmailSignature(
                        _mail,
                        _isInlineResponse,
                        isPlainText,
                        signatureContent,
                        IsReplyOrForwardComposeKind(composeKind),
                        _composeKey,
                        "apply:" + (reason ?? "n/a"),
                        _inlineExplorerIdentityKey);
                return ApplyEmailSignatureReconcileResult(reconcile);
            }

            private EmailSignatureApplicationResult ClearManagedEmailSignature(string reason)
            {
                if (_owner == null || _owner._mailInteropController == null || _mail == null)
                {
                    return EmailSignatureFailure("interop_unavailable");
                }

                MailInteropController.EmailSignatureReconcileResult reconcile =
                    _owner._mailInteropController.ClearManagedEmailSignature(
                        _mail,
                        _isInlineResponse,
                        _composeKey,
                        "clear_managed:" + (reason ?? "n/a"),
                        _inlineExplorerIdentityKey);
                return ApplyEmailSignatureReconcileResult(reconcile);
            }

            private EmailSignatureApplicationResult ClearInitialEmailSignatureSlot(string reason)
            {
                if (_owner == null || _owner._mailInteropController == null || _mail == null)
                {
                    return EmailSignatureFailure("interop_unavailable");
                }

                MailInteropController.EmailSignatureReconcileResult reconcile =
                    _owner._mailInteropController.ClearInitialEmailSignatureSlot(
                        _mail,
                        _isInlineResponse,
                        _composeKey,
                        "clear_initial:" + (reason ?? "n/a"),
                        _inlineExplorerIdentityKey);
                return ApplyEmailSignatureReconcileResult(reconcile);
            }

            private EmailSignatureApplicationResult ApplyEmailSignatureReconcileResult(
                MailInteropController.EmailSignatureReconcileResult reconcile)
            {
                if (reconcile == null || !reconcile.Success)
                {
                    return EmailSignatureFailure(
                        reconcile != null ? reconcile.Source : "reconcile_unavailable");
                }

                return EmailSignatureSuccess(reconcile.Source);
            }

            private static EmailSignatureApplicationResult EmailSignatureSuccess(string source)
            {
                return new EmailSignatureApplicationResult
                {
                    Success = true,
                    Source = source ?? "success"
                };
            }

            private static EmailSignatureApplicationResult EmailSignatureFailure(string source)
            {
                return new EmailSignatureApplicationResult
                {
                    Success = false,
                    Source = source ?? "failed"
                };
            }

            private bool TryFinalizeEmailSignatureBeforeSend(ref bool cancel)
            {
                if (_disposed || cancel || _owner == null || _mail == null)
                {
                    return !cancel;
                }

                _emailSignatureTimer.Stop();
                _emailSignatureRequestGeneration++;
                _owner.EnsureSettingsLoaded();
                AddinSettings settings = _owner._currentSettings ?? new AddinSettings();
                var configuration = new TalkServiceConfiguration(
                    settings.ServerUrl,
                    settings.Username,
                    settings.AppPassword);

                if (!configuration.IsComplete())
                {
                    EmailSignatureApplicationResult noConfiguration =
                        ReconcileEmailSignatureWithoutBackendConfiguration("send");
                    if (noConfiguration.Success)
                    {
                        _emailSignatureStateStable = true;
                        _pendingEmailSignatureReason = string.Empty;
                        _deferredEmailSignatureReason = string.Empty;
                        return true;
                    }

                    LogEmailSignature(
                        "best-effort managed cleanup could not run before send with incomplete configuration (source="
                        + (noConfiguration.Source ?? "n/a")
                        + ").");
                    return true;
                }

                BackendPolicyStatus policyStatus;
                if (!_owner.TryGetCachedEmailSignaturePolicyStatus(
                        configuration,
                        out policyStatus)
                    || policyStatus == null
                    || !policyStatus.FetchSucceeded)
                {
                    ScheduleEmailSignatureApplication("send_policy_required");
                    cancel = true;
                    return BlockEmailSignatureSend(
                        ref cancel,
                        "policy_unavailable",
                        true);
                }

                if (_composeSurfaceState == ComposeSurfaceState.Detached)
                {
                    if (_emailSignatureStateStable
                        && string.IsNullOrWhiteSpace(_pendingEmailSignatureReason)
                        && string.IsNullOrWhiteSpace(_deferredEmailSignatureReason))
                    {
                        LogEmailSignature(
                            "send accepted from detached inline transition using the last stable reconcile.");
                        return true;
                    }

                    DeferEmailSignatureApplication("send_reconcile_required", "send_detached");
                    cancel = true;
                    return BlockEmailSignatureSend(
                        ref cancel,
                        "compose_surface_detached",
                        false);
                }

                EmailSignatureApplicationResult finalResult = ApplyEmailSignaturePolicy(
                    policyStatus,
                    settings,
                    "send");
                if (finalResult.Success)
                {
                    _emailSignatureStateStable = true;
                    _emailSignatureReadyRetryCount = 0;
                    _pendingEmailSignatureReason = string.Empty;
                    _deferredEmailSignatureReason = string.Empty;
                    return true;
                }

                ScheduleEmailSignatureApplication("send_reconcile_failed");
                cancel = true;
                return BlockEmailSignatureSend(
                    ref cancel,
                    finalResult.Source,
                    string.Equals(
                        finalResult.Source,
                        "policy_unavailable",
                        StringComparison.OrdinalIgnoreCase));
            }

            private bool BlockEmailSignatureSend(
                ref bool cancel,
                string source,
                bool policyUnavailable)
            {
                cancel = true;
                _emailSignatureStateStable = false;
                LogEmailSignature(
                    "send blocked because final reconcile is not reliable (source="
                    + (source ?? "n/a")
                    + ").");

                string message = policyUnavailable
                    ? Strings.PolicyWarningTitle
                    : string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.ErrorInsertHtmlFailed,
                        "email signature (" + (source ?? "unknown") + ")");
                MessageBox.Show(
                    message,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            private EmailSignatureComposeKind ResolveEmailSignatureComposeKind()
            {
                int verb;
                bool verbRead = TryReadLastVerbExecuted(out verb);
                EmailSignatureComposeKind kind;
                if (verbRead && TryMapLastVerbToComposeKind(verb, out kind))
                {
                    return kind;
                }

                if (_isInlineResponse)
                {
                    return EmailSignatureComposeKind.Response;
                }

                if (TryResolveConversationComposeKind(out kind))
                {
                    return kind;
                }

                if (verbRead && verb == 0)
                {
                    return EmailSignatureComposeKind.New;
                }

                return EmailSignatureComposeKind.Unknown;
            }

            private static bool TryMapLastVerbToComposeKind(int verb, out EmailSignatureComposeKind kind)
            {
                if (verb == 102 || verb == 103)
                {
                    kind = EmailSignatureComposeKind.Reply;
                    return true;
                }
                if (verb == 104)
                {
                    kind = EmailSignatureComposeKind.Forward;
                    return true;
                }
                kind = EmailSignatureComposeKind.Unknown;
                return false;
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
                            LogEmailSignature("Last verb is not set; checking conversation index for email signature compose kind.");
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

            private bool TryResolveConversationComposeKind(out EmailSignatureComposeKind kind)
            {
                kind = EmailSignatureComposeKind.Unknown;
                int byteCount;
                if (!TryReadConversationIndexByteCount(out byteCount))
                {
                    return false;
                }
                if (byteCount > 22)
                {
                    kind = EmailSignatureComposeKind.Response;
                    return true;
                }
                if (byteCount == 22)
                {
                    kind = EmailSignatureComposeKind.New;
                    return true;
                }
                return false;
            }

            private bool TryReadConversationIndexByteCount(out int byteCount)
            {
                byteCount = 0;
                Outlook.PropertyAccessor accessor = null;
                try
                {
                    accessor = _mail.PropertyAccessor;
                    object value = accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x00710102");
                    byteCount = ResolveConversationIndexByteCount(value);
                    return byteCount > 0;
                }
                catch (COMException ex)
                {
                    uint errorCode = unchecked((uint)ex.ErrorCode);
                    if (errorCode == 0x8004010Fu)
                    {
                        if (DiagnosticsLogger.IsEnabled)
                        {
                            LogEmailSignature("Conversation index is not set for email signature compose kind.");
                        }
                        return false;
                    }
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read conversation index for email signature compose kind.", ex);
                    return false;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read conversation index for email signature compose kind.", ex);
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release mail PropertyAccessor COM object.");
                }
            }

            private static int ResolveConversationIndexByteCount(object value)
            {
                if (value == null)
                {
                    return 0;
                }

                byte[] bytes = value as byte[];
                if (bytes != null)
                {
                    return bytes.Length;
                }

                Array array = value as Array;
                if (array != null)
                {
                    return array.Length;
                }

                string text = Convert.ToString(value, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(text))
                {
                    return 0;
                }

                string normalized = text.Trim();
                if (normalized.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    normalized = normalized.Substring(2);
                }
                if (normalized.Length % 2 == 0 && Regex.IsMatch(normalized, "^[0-9A-Fa-f]+$"))
                {
                    return normalized.Length / 2;
                }
                return 0;
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
                if (composeKind == EmailSignatureComposeKind.Response)
                {
                    return policy.OnReply && policy.OnForward;
                }
                return false;
            }

            private static bool IsReplyOrForwardComposeKind(EmailSignatureComposeKind composeKind)
            {
                return composeKind == EmailSignatureComposeKind.Reply
                       || composeKind == EmailSignatureComposeKind.Forward
                       || composeKind == EmailSignatureComposeKind.Response;
            }

            private string ResolveCurrentSenderEmail()
            {
                Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
                return OutlookRecipientResolverController.ResolveEffectiveSenderSmtpAddress(
                    _mail,
                    application,
                    LogCategories.Core,
                    "compose",
                    string.Empty,
                    true);
            }

            private static bool IsEmailSignaturePropertyChange(string propertyName)
            {
                return string.Equals(propertyName, "SendUsingAccount", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(propertyName, "SenderEmailAddress", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(propertyName, "SentOnBehalfOfName", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(propertyName, "BodyFormat", StringComparison.OrdinalIgnoreCase);
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
