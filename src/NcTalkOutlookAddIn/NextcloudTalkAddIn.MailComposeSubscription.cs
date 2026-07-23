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
        internal sealed partial class MailComposeSubscription : IDisposable
        {
            private sealed class AttachmentBatchEntry
            {
                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentBatchInfo
            {
                internal int Count { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentSnapshot
            {
                internal int Index { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentAutomationSettings
            {
                internal bool AlwaysConnector { get; set; }

                internal bool OfferAboveEnabled { get; set; }

                internal int ThresholdMb { get; set; }

                internal long ThresholdBytes { get; set; }
            }

            private sealed class ComposeShareCleanupEntry
            {
                internal string RelativeFolder { get; set; }

                internal string ShareId { get; set; }

                internal string ShareLabel { get; set; }

            }

            private sealed class BeforeAddShareEntry
            {
                internal AttachmentBatchEntry Candidate { get; set; }

                internal string LocalPath { get; set; }

                internal int ThresholdMb { get; set; }

                internal bool CleanupLocalPathAfterFlow { get; set; }

                internal string Trigger { get; set; }

                internal int BaselineAttachmentCount { get; set; }
            }

            private enum ComposeSurfaceState
            {
                Detached,
                Inspector,
                InlineResponse
            }

            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.MailItem _mail;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly string _mailIdentityKey;
            private readonly string _inspectorIdentityKey;
            private readonly string _composeKey;
            private readonly System.Windows.Forms.Timer _attachmentEvalTimer = new System.Windows.Forms.Timer();
            private readonly System.Windows.Forms.Timer _beforeAddShareTimer = new System.Windows.Forms.Timer();
            private readonly System.Windows.Forms.Timer _cleanupGraceTimer = new System.Windows.Forms.Timer();
            private readonly System.Windows.Forms.Timer _emailSignatureTimer = new System.Windows.Forms.Timer();
            private readonly List<AttachmentBatchEntry> _pendingAddedBatch = new List<AttachmentBatchEntry>();
            private readonly List<BeforeAddShareEntry> _pendingBeforeAddShareEntries = new List<BeforeAddShareEntry>();
            private readonly List<ComposeShareCleanupEntry> _cleanupEntries = new List<ComposeShareCleanupEntry>();
            private readonly List<SeparatePasswordDispatchEntry> _passwordDispatchQueue = new List<SeparatePasswordDispatchEntry>();
            private bool _attachmentSuppressed;
            private bool _attachmentPromptOpen;
            private bool _beforeAddShareFlowRunning;
            private bool _sendPending;
            private DateTime _sendPendingAtUtc;
            private bool _awaitingGraceCloseResolution;
            private bool _emailSignatureApplying;
            private bool _isInlineResponse;
            private ComposeSurfaceState _composeSurfaceState;
            private string _activeInspectorIdentityKey = string.Empty;
            private string _inlineExplorerIdentityKey = string.Empty;
            private string _deferredEmailSignatureReason = string.Empty;
            private string _pendingEmailSignatureReason = string.Empty;
            private bool _disposed;
            private const int BeforeAddShareBatchDebounceMs = 3000;
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
                _isInlineResponse = isInlineResponse;
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

                // Debounce long enough so Outlook can raise all BeforeAttachmentAdd callbacks
                // for a multi-select operation before we open the wizard.
                _beforeAddShareTimer.Interval = BeforeAddShareBatchDebounceMs;
                _beforeAddShareTimer.Tick += OnBeforeAddShareTimerTick;

                _cleanupGraceTimer.Interval = ComposeShareCleanupSendGraceMs;
                _cleanupGraceTimer.Tick += OnCleanupGraceTimerTick;

                _emailSignatureTimer.Interval = isInlineResponse ? EmailSignatureInlineApplyDebounceMs : EmailSignatureApplyDebounceMs;
                _emailSignatureTimer.Tick += OnEmailSignatureTimerTick;

                _events = mail as Outlook.ItemEvents_10_Event;
                if (_events != null)
                {
                    _events.BeforeAttachmentAdd += OnBeforeAttachmentAdd;
                    _events.AttachmentAdd += OnAttachmentAdd;
                    _events.PropertyChange += OnPropertyChange;
                    _events.Send += OnSend;
                    _events.Close += OnClose;
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

                ScheduleEmailSignatureApplication("compose_open");
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
                _isInlineResponse = true;
                _activeInspectorIdentityKey = string.Empty;
                _inlineExplorerIdentityKey = normalizedExplorerIdentityKey;
                _emailSignatureTimer.Interval = EmailSignatureInlineApplyDebounceMs;
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
                _isInlineResponse = false;
                _activeInspectorIdentityKey = normalizedInspectorIdentityKey;
                _inlineExplorerIdentityKey = string.Empty;
                _emailSignatureTimer.Interval = EmailSignatureApplyDebounceMs;
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
                _composeSurfaceState = ComposeSurfaceState.Detached;
                _isInlineResponse = false;
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

            internal void ArmShareCleanup(FileLinkResult result)
            {
                if (result == null)
                {
                    return;
                }
                string relativeFolder = string.IsNullOrWhiteSpace(result.RelativePath)
                    ? string.Empty
                    : result.RelativePath.Trim();
                if (string.IsNullOrWhiteSpace(relativeFolder))
                {
                    LogFileLink("Compose share cleanup arm skipped (composeKey=" + _composeKey + ", reason=missing_relative_folder).");
                    return;
                }

                _cleanupGraceTimer.Stop();
                _awaitingGraceCloseResolution = false;
                _sendPending = false;

                _cleanupEntries.Add(new ComposeShareCleanupEntry
                {
                    RelativeFolder = relativeFolder,
                    ShareId = result.ShareId ?? string.Empty,
                    ShareLabel = result.FolderName ?? string.Empty
                });

                LogFileLink(
                    "Compose share cleanup armed (composeKey="
                    + _composeKey
                    + ", relativeFolder="
                    + relativeFolder
                    + ", shareId="
                    + (result.ShareId ?? string.Empty)
                    + ", shareLabel="
                    + (result.FolderName ?? string.Empty)
                    + ", armedCount="
                    + _cleanupEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            internal void RegisterSeparatePasswordDispatch(
                FileLinkResult result,
                FileLinkRequest request,
                string passwordOnlyHtml,
                string passwordOnlyPlainText,
                bool isPlainText,
                string languageOverride,
                BackendPolicyStatus policyStatus)
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
                    IsPlainText = isPlainText,
                    DeliveryMode = deliveryPolicy.Mode,
                    SecretsExpireDays = deliveryPolicy.SecretsExpireDays,
                    LanguageOverride = string.IsNullOrWhiteSpace(languageOverride) ? "default" : languageOverride,
                    BackendPolicyStatus = policyStatus,
                    To = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("To")),
                    Cc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("CC")),
                    Bcc = ComposeShareLifecycleController.BuildNormalizedRecipientCsv(ReadMailRecipientList("BCC"))
                };

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

            private static string ReadAttachmentPathName(Outlook.Attachment attachment)
            {
                if (attachment == null)
                {
                    return string.Empty;
                }
                try
                {
                    return attachment.PathName ?? string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read Attachment.PathName.", ex);
                    return string.Empty;
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _attachmentEvalTimer.Stop();
                _beforeAddShareTimer.Stop();
                _cleanupGraceTimer.Stop();
                _emailSignatureTimer.Stop();

                if (_pendingBeforeAddShareEntries.Count > 0)
                {
                    var pendingTemporaryFiles = new List<string>();
                    for (int i = 0; i < _pendingBeforeAddShareEntries.Count; i++)
                    {
                        BeforeAddShareEntry pendingEntry = _pendingBeforeAddShareEntries[i];
                        if (pendingEntry == null || !pendingEntry.CleanupLocalPathAfterFlow || string.IsNullOrWhiteSpace(pendingEntry.LocalPath))
                        {
                            continue;
                        }

                        pendingTemporaryFiles.Add(pendingEntry.LocalPath);
                    }

                    CleanupTemporaryFiles(pendingTemporaryFiles);
                    _pendingBeforeAddShareEntries.Clear();
                }
                try
                {
                    _attachmentEvalTimer.Tick -= OnAttachmentEvalTimerTick;
                    _beforeAddShareTimer.Tick -= OnBeforeAddShareTimerTick;
                    _cleanupGraceTimer.Tick -= OnCleanupGraceTimerTick;
                    _emailSignatureTimer.Tick -= OnEmailSignatureTimerTick;
                    _attachmentEvalTimer.Dispose();
                    _beforeAddShareTimer.Dispose();
                    _cleanupGraceTimer.Dispose();
                    _emailSignatureTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose compose timers.", ex);
                }
                if (_events != null)
                {
                    try
                    {
                        _events.BeforeAttachmentAdd -= OnBeforeAttachmentAdd;
                        _events.AttachmentAdd -= OnAttachmentAdd;
                        _events.PropertyChange -= OnPropertyChange;
                        _events.Send -= OnSend;
                        _events.Close -= OnClose;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to detach compose event handlers.", ex);
                    }
                }

                _owner.RemoveMailComposeSubscription(this);
                _disposed = true;
                LogFileLink(
                    "Compose subscription disposed (composeKey="
                    + _composeKey
                    + ", hadCleanup="
                    + (_cleanupEntries.Count > 0).ToString(CultureInfo.InvariantCulture)
                    + ", hadPasswordDispatch="
                    + (_passwordDispatchQueue.Count > 0).ToString(CultureInfo.InvariantCulture)
                    + ", delayed="
                    + _awaitingGraceCloseResolution.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }
        }

    }
}

