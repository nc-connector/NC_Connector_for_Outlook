// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Controllers;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription : IDisposable
        {
            private enum ComposeSurfaceState
            {
                Detached,
                Inspector,
                InlineResponse
            }

            private readonly NextcloudTalkAddIn _owner;
            private Outlook.MailItem _mail;
            private Outlook.ItemEvents_10_Event _events;
            private Outlook.Inspector _inspector;
            private Outlook.InspectorEvents_10_Event _inspectorEvents;
            private string _boundInspectorIdentityKey = string.Empty;
            private readonly string _mailIdentityKey;
            private readonly string _inspectorIdentityKey;
            private readonly string _composeKey;
            private readonly System.Windows.Forms.Timer _attachmentEvalTimer = new System.Windows.Forms.Timer();
            private readonly System.Windows.Forms.Timer _emailSignatureTimer = new System.Windows.Forms.Timer();
            private readonly List<AttachmentBatchEntry> _pendingAddedBatch = new List<AttachmentBatchEntry>();
            // Final recipients and sender are captured at Send time. A reopened draft can
            // change both, so password follow-ups stay with the live compose subscription.
            private readonly List<SeparatePasswordDispatchEntry> _passwordDispatchQueue = new List<SeparatePasswordDispatchEntry>();
            private bool _attachmentSuppressed;
            private bool _attachmentPromptOpen;
            private bool _attachmentEvaluationPending;
            private bool _attachmentEvaluationRunning;
            private bool _emailSignatureApplying;
            private ComposeSurfaceState _composeSurfaceState;
            private string _activeInspectorIdentityKey = string.Empty;
            private string _inlineExplorerIdentityKey = string.Empty;
            private string _deferredEmailSignatureReason = string.Empty;
            private string _pendingEmailSignatureReason = string.Empty;
            private bool _disposed;
            private const int EmailSignatureApplyDebounceMs = 900;
            private const int EmailSignatureInlineApplyDebounceMs = 250;

            internal MailComposeSubscription(
                NextcloudTalkAddIn owner,
                Outlook.MailItem mail,
                string mailIdentityKey,
                string inspectorIdentityKey,
                bool isInlineResponse,
                string inlineExplorerIdentityKey)
            {
                _owner = owner;
                _mail = mail;
                _mailIdentityKey = string.IsNullOrWhiteSpace(mailIdentityKey)
                    ? ComInteropScope.ResolveIdentityKey(mail, LogCategories.FileLink, "MailItem")
                    : mailIdentityKey.Trim();
                _inspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityKey)
                    ? (isInlineResponse ? string.Empty : MailInteropController.ResolveMailInspectorIdentityKey(mail))
                    : inspectorIdentityKey.Trim();
                _composeSurfaceState = isInlineResponse ? ComposeSurfaceState.InlineResponse : ComposeSurfaceState.Inspector;
                _activeInspectorIdentityKey = isInlineResponse ? string.Empty : _inspectorIdentityKey;
                _inlineExplorerIdentityKey = isInlineResponse && !string.IsNullOrWhiteSpace(inlineExplorerIdentityKey)
                    ? inlineExplorerIdentityKey.Trim()
                    : string.Empty;
                _composeKey = BuildComposeKey(mail, _mailIdentityKey, _inspectorIdentityKey);

                _attachmentEvalTimer.Interval = ComposeAttachmentEvalDebounceMs;
                _attachmentEvalTimer.Tick += OnAttachmentEvalTimerTick;

                _emailSignatureTimer.Interval = isInlineResponse ? EmailSignatureInlineApplyDebounceMs : EmailSignatureApplyDebounceMs;
                _emailSignatureTimer.Tick += OnEmailSignatureTimerTick;

                _events = mail as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.AttachmentAdd += OnAttachmentAdd;
                    _events.PropertyChange += OnPropertyChange;
                    _events.Send += OnSend;
                    _events.AfterWrite += OnAfterWrite;
                    _events.Unload += OnUnload;
                }

                LogFileLink(
                    "Compose subscription registered (composeKey="
                    + _composeKey
                    + ", mailIdentity="
                    + (_mailIdentityKey ?? string.Empty)
                    + ", inspectorIdentity="
                    + (_inspectorIdentityKey ?? string.Empty)
                    + ", surface="
                    + _composeSurfaceState.ToString()
                    + ", inlineExplorerIdentity="
                    + (_inlineExplorerIdentityKey ?? string.Empty)
                    + ").");

                BeginAttachmentAutomationSettingsRefresh();
                // Forwarded and reopened drafts can already contain files before attachment events are hooked.
                ScheduleAttachmentEvaluation();
                ScheduleEmailSignatureApplication("compose_open");
            }

            internal void BindInspectorLifecycle(
                Outlook.Inspector inspector)
            {
                if (_disposed
                    || (_inspectorEvents != null && inspector == null))
                {
                    return;
                }

                Outlook.Inspector resolvedInspector = inspector;
                if (resolvedInspector == null)
                {
                    try
                    {
                        resolvedInspector = _mail != null
                            ? _mail.GetInspector
                            : null;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to resolve compose Inspector for lifecycle binding (composeKey="
                            + _composeKey
                            + ").",
                            ex);
                        return;
                    }
                }

                Outlook.InspectorEvents_10_Event inspectorEvents =
                    resolvedInspector as Outlook.InspectorEvents_10_Event;
                if (inspectorEvents == null)
                {
                    LogFileLink(
                        "Compose Inspector lifecycle binding unavailable (composeKey="
                        + _composeKey
                        + ").");
                    return;
                }

                string inspectorIdentityKey =
                    ComInteropScope.ResolveIdentityKey(
                        resolvedInspector,
                        LogCategories.FileLink,
                        "Compose Inspector");
                if (_inspectorEvents != null
                    && (ComInteropScope.AreSameObject(
                            _inspector,
                            resolvedInspector,
                            LogCategories.FileLink,
                            "Bound compose Inspector",
                            "New compose Inspector")
                        || (!string.IsNullOrWhiteSpace(
                                _boundInspectorIdentityKey)
                            && string.Equals(
                                _boundInspectorIdentityKey,
                                inspectorIdentityKey,
                                StringComparison.Ordinal))))
                {
                    return;
                }
                if (_inspectorEvents != null)
                {
                    UnbindInspectorLifecycle();
                }

                try
                {
                    inspectorEvents.Close += OnInspectorClosed;
                    _inspector = resolvedInspector;
                    _inspectorEvents = inspectorEvents;
                    _boundInspectorIdentityKey =
                        inspectorIdentityKey;
                    LogFileLink(
                        "Compose Inspector.Close lifecycle bound (composeKey="
                        + _composeKey
                        + ", inspectorKey="
                        + (_boundInspectorIdentityKey ?? string.Empty)
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to bind compose Inspector.Close lifecycle (composeKey="
                        + _composeKey
                        + ").",
                        ex);
                }
            }

            internal void MarkInlineResponse(string explorerIdentityKey)
            {
                if (_disposed)
                {
                    return;
                }

                string normalizedExplorerIdentityKey = string.IsNullOrWhiteSpace(explorerIdentityKey)
                    ? _inlineExplorerIdentityKey
                    : explorerIdentityKey.Trim();
                if (_composeSurfaceState == ComposeSurfaceState.InlineResponse
                    && string.Equals(
                        _inlineExplorerIdentityKey,
                        normalizedExplorerIdentityKey,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                ComposeSurfaceState previousSurface = _composeSurfaceState;
                _composeSurfaceState = ComposeSurfaceState.InlineResponse;
                _activeInspectorIdentityKey = string.Empty;
                _inlineExplorerIdentityKey = normalizedExplorerIdentityKey;
                _emailSignatureTimer.Interval = EmailSignatureInlineApplyDebounceMs;
                if (_attachmentEvaluationPending) { ScheduleAttachmentEvaluation(); }
                LogEmailSignature(
                    "compose surface changed (from="
                    + previousSurface.ToString()
                    + ", to=InlineResponse, explorerKey="
                    + (_inlineExplorerIdentityKey ?? string.Empty)
                    + ").");
                if (!ResumeDeferredEmailSignatureApplication("inline_surface"))
                {
                    ScheduleEmailSignatureApplication("inline_response");
                }
            }

            internal void MarkInspector(string inspectorIdentityKey)
            {
                if (_disposed)
                {
                    return;
                }

                string normalizedInspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityKey)
                    ? _activeInspectorIdentityKey
                    : inspectorIdentityKey.Trim();
                if (_composeSurfaceState == ComposeSurfaceState.Inspector
                    && string.Equals(
                        _activeInspectorIdentityKey,
                        normalizedInspectorIdentityKey,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                ComposeSurfaceState previousSurface = _composeSurfaceState;
                _composeSurfaceState = ComposeSurfaceState.Inspector;
                _activeInspectorIdentityKey = normalizedInspectorIdentityKey;
                _inlineExplorerIdentityKey = string.Empty;
                _emailSignatureTimer.Interval = EmailSignatureApplyDebounceMs;
                if (_attachmentEvaluationPending) { ScheduleAttachmentEvaluation(); }
                LogEmailSignature(
                    "compose surface changed (from="
                    + previousSurface.ToString()
                    + ", to=Inspector, inspectorKey="
                    + (_activeInspectorIdentityKey ?? string.Empty)
                    + ").");
                if (!ResumeDeferredEmailSignatureApplication("inspector_surface"))
                {
                    ScheduleEmailSignatureApplication("inspector_surface");
                }
            }

            internal void MarkInlineResponseClosed(string explorerIdentityKey)
            {
                if (_disposed || _composeSurfaceState != ComposeSurfaceState.InlineResponse)
                {
                    return;
                }

                string normalizedExplorerIdentityKey = string.IsNullOrWhiteSpace(explorerIdentityKey)
                    ? string.Empty
                    : explorerIdentityKey.Trim();
                if (!string.IsNullOrWhiteSpace(_inlineExplorerIdentityKey)
                    && !string.Equals(
                        _inlineExplorerIdentityKey,
                        normalizedExplorerIdentityKey,
                        StringComparison.OrdinalIgnoreCase))
                {
                    LogEmailSignature(
                        "inline close ignored for non-owning Explorer (expected="
                        + _inlineExplorerIdentityKey
                        + ", actual="
                        + normalizedExplorerIdentityKey
                        + ").");
                    return;
                }

                _emailSignatureTimer.Stop();
                _attachmentEvalTimer.Stop();
                _composeSurfaceState = ComposeSurfaceState.Detached;
                _activeInspectorIdentityKey = string.Empty;
                _inlineExplorerIdentityKey = string.Empty;
                LogEmailSignature(
                    "compose surface changed (from=InlineResponse, to=Detached, explorerKey="
                    + normalizedExplorerIdentityKey
                    + ").");
            }

            private void DeferEmailSignatureApplication(string reason, string source)
            {
                if (_disposed)
                {
                    return;
                }

                _deferredEmailSignatureReason = string.IsNullOrWhiteSpace(reason)
                    ? "property_unknown"
                    : reason.Trim();
                LogEmailSignature(
                    "reconcile deferred (reason="
                    + _deferredEmailSignatureReason
                    + ", source="
                    + (source ?? string.Empty)
                    + ", surface="
                    + _composeSurfaceState.ToString()
                    + ", attachmentSuppressed="
                    + _attachmentSuppressed.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private bool ResumeDeferredEmailSignatureApplication(string trigger)
            {
                if (_disposed
                    || _attachmentSuppressed
                    || _composeSurfaceState == ComposeSurfaceState.Detached
                    || string.IsNullOrWhiteSpace(_deferredEmailSignatureReason))
                {
                    return false;
                }

                string reason = _deferredEmailSignatureReason;
                _deferredEmailSignatureReason = string.Empty;
                LogEmailSignature(
                    "deferred reconcile resumed (reason="
                    + reason
                    + ", trigger="
                    + (trigger ?? string.Empty)
                    + ", surface="
                    + _composeSurfaceState.ToString()
                    + ").");
                ScheduleEmailSignatureApplication(reason);
                return true;
            }

            internal bool IsFor(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                if (mail == null)
                {
                    return false;
                }
                if (ReferenceEquals(mail, _mail) || mail == _mail)
                {
                    return true;
                }
                if (string.IsNullOrWhiteSpace(_mailIdentityKey))
                {
                    if (string.IsNullOrWhiteSpace(_inspectorIdentityKey))
                    {
                        return false;
                    }
                }
                if (!string.IsNullOrWhiteSpace(_mailIdentityKey)
                    && string.Equals(_mailIdentityKey, mailIdentityKey ?? string.Empty, StringComparison.Ordinal))
                {
                    return true;
                }
                return !string.IsNullOrWhiteSpace(_inspectorIdentityKey)
                    && string.Equals(_inspectorIdentityKey, inspectorIdentityKey ?? string.Empty, StringComparison.Ordinal);
            }

            internal void RegisterSeparatePasswordDispatch(
                FileLinkResult result,
                FileLinkRequest request,
                string passwordOnlyHtml,
                string passwordOnlyPlainText,
                string secretsHtmlTemplate,
                string secretsPlainTextTemplate,
                bool isPlainText,
                string languageOverride,
                BackendPolicyStatus policyStatus,
                ComposeLifecycleOrigin origin)
            {
                if (result == null || request == null)
                {
                    return;
                }
                if (isPlainText && string.IsNullOrWhiteSpace(passwordOnlyPlainText))
                {
                    return;
                }
                if (!isPlainText && string.IsNullOrWhiteSpace(passwordOnlyHtml))
                {
                    return;
                }
                string password = result.Password ?? string.Empty;
                if (string.IsNullOrWhiteSpace(password))
                {
                    return;
                }
                SharePasswordDeliveryPolicy deliveryPolicy = SharePasswordDeliveryPolicy.Resolve(policyStatus, request.PasswordDeliveryMode);
                var entry = new SeparatePasswordDispatchEntry
                {
                    ShareLabel = result.FolderName ?? string.Empty,
                    ShareUrl = result.ShareUrl ?? string.Empty,
                    ShareId = result.ShareId ?? string.Empty,
                    ShareToken = result.ShareToken ?? string.Empty,
                    RelativePath = result.RelativePath ?? string.Empty,
                    ExpireDate = result.ExpireDate,
                    Permissions = result.Permissions,
                    Password = password.Trim(),
                    Html = passwordOnlyHtml,
                    PlainText = passwordOnlyPlainText,
                    SecretsHtmlTemplate = secretsHtmlTemplate,
                    SecretsPlainTextTemplate = secretsPlainTextTemplate,
                    IsPlainText = isPlainText,
                    DeliveryMode = deliveryPolicy.Mode,
                    SecretsExpireDays = deliveryPolicy.SecretsExpireDays,
                    LanguageOverride = string.IsNullOrWhiteSpace(languageOverride) ? "default" : languageOverride,
                    BackendPolicyStatus = policyStatus,
                    To = RecipientAddressList.BuildNormalizedRecipientCsv(ReadMailRecipientList("To")),
                    Cc = RecipientAddressList.BuildNormalizedRecipientCsv(ReadMailRecipientList("CC")),
                    Bcc = RecipientAddressList.BuildNormalizedRecipientCsv(ReadMailRecipientList("BCC")),
                    Origin = origin != null ? origin.Clone() : null
                };

                _owner.CaptureSeparatePasswordSignatureSnapshot(
                    new List<SeparatePasswordDispatchEntry> { entry },
                    policyStatus,
                    _owner.CurrentSettings != null
                        ? _owner.CurrentSettings.Clone()
                        : new AddinSettings(),
                    _composeKey);
                _passwordDispatchQueue.Add(entry);
                LogFileLink(
                    "Separate password dispatch registered (composeKey="
                    + _composeKey
                    + ", queued="
                    + _passwordDispatchQueue.Count.ToString(CultureInfo.InvariantCulture)
                    + ", hasShareUrl="
                    + (!string.IsNullOrWhiteSpace(entry.ShareUrl)).ToString(CultureInfo.InvariantCulture)
                    + ", plainText="
                    + entry.IsPlainText.ToString(CultureInfo.InvariantCulture)
                    + ", deliveryMode="
                    + entry.DeliveryMode.ToString()
                    + ").");
            }

            private static string BuildComposeKey(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                if (mail != null)
                {
                    try
                    {
                        string entryId = mail.EntryID;
                        if (!string.IsNullOrWhiteSpace(entryId))
                        {
                            return entryId.Trim();
                        }
                    }
                    catch (COMException ex)
                    {
                        uint errorCode = unchecked((uint)ex.ErrorCode);
                        if ((errorCode & 0xFFFFu) == 0x0108u)
                        {
                            LogFileLink(
                                "MailItem.EntryID unavailable while building compose key (hresult=0x"
                                + errorCode.ToString("X8", CultureInfo.InvariantCulture)
                                + ").");
                        }
                        else
                        {
                            DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.EntryID for compose key.", ex);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.EntryID for compose key.", ex);
                    }
                }
                if (!string.IsNullOrWhiteSpace(mailIdentityKey))
                {
                    return mailIdentityKey.Trim();
                }
                if (!string.IsNullOrWhiteSpace(inspectorIdentityKey))
                {
                    return inspectorIdentityKey.Trim();
                }
                return Guid.NewGuid().ToString("N");
            }

            private static string ReadAttachmentName(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return string.Empty;
                }
                try
                {
                    string fileName = attachment.FileName;
                    if (!string.IsNullOrWhiteSpace(fileName))
                    {
                        return fileName.Trim();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.FileName.", ex);
                }
                try
                {
                    string displayName = attachment.DisplayName;
                    if (!string.IsNullOrWhiteSpace(displayName))
                    {
                        return displayName.Trim();
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.DisplayName.", ex);
                }
                return string.Empty;
            }

            private static long ReadAttachmentSizeBytes(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return 0;
                }
                try
                {
                    return Math.Max(0, attachment.Size);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.Size.", ex);
                    return 0;
                }
            }


            public void Dispose()
            {
                Dispose(true);
            }

            private void UnbindInspectorLifecycle()
            {
                Outlook.InspectorEvents_10_Event inspectorEvents =
                    _inspectorEvents;
                _inspectorEvents = null;
                _inspector = null;
                _boundInspectorIdentityKey = string.Empty;
                if (inspectorEvents == null)
                {
                    return;
                }

                try
                {
                    inspectorEvents.Close -= OnInspectorClosed;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to detach compose Inspector.Close lifecycle (composeKey="
                        + _composeKey
                        + ").",
                        ex);
                }
            }

            private void Dispose(bool detachItemEvents)
            {
                if (_disposed)
                {
                    return;
                }

                int pendingShareCleanup =
                    _shareCleanupTracker.ReleaseAll();
                _attachmentEvalTimer.Stop();
                _emailSignatureTimer.Stop();

                if (_sharingWizard != null && !_sharingWizard.IsDisposed) { _sharingWizard.Close(); }
                try
                {
                    _attachmentEvalTimer.Tick -= OnAttachmentEvalTimerTick;
                    _emailSignatureTimer.Tick -= OnEmailSignatureTimerTick;
                    _attachmentEvalTimer.Dispose();
                    _emailSignatureTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose compose timers.", ex);
                }
                UnbindInspectorLifecycle();
                if (detachItemEvents && _events != null)
                {
                    try
                    {
                        _events.AttachmentAdd -= OnAttachmentAdd;
                        _events.PropertyChange -= OnPropertyChange;
                        _events.Send -= OnSend;
                        _events.AfterWrite -= OnAfterWrite;
                        _events.Unload -= OnUnload;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to detach compose event handlers.", ex);
                    }
                }

                _owner.RemoveMailComposeSubscription(this);
                _events = null;
                _mail = null;
                _disposed = true;
                LogFileLink(
                    "Compose subscription disposed (composeKey="
                    + _composeKey
                    + ", hadPasswordDispatch="
                    + (_passwordDispatchQueue.Count > 0).ToString(CultureInfo.InvariantCulture)
                    + ", releasedShareCleanup="
                    + pendingShareCleanup.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }
        }

    }
}

