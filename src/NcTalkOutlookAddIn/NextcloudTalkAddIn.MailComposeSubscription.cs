/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

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
        internal sealed class MailComposeSubscription : IDisposable
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

                internal DateTime CreatedUtc { get; set; }
            }

            private sealed class BeforeAddShareEntry
            {
                internal AttachmentBatchEntry Candidate { get; set; }

                internal string LocalPath { get; set; }

                internal int ThresholdMb { get; set; }

                internal bool CleanupLocalPathAfterFlow { get; set; }

                internal string Trigger { get; set; }
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
            private bool _disposed;
            private const int BeforeAddShareBatchDebounceMs = 3000;

            internal MailComposeSubscription(NextcloudTalkAddIn owner, Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                _owner = owner;
                _mail = mail;
                _mailIdentityKey = string.IsNullOrWhiteSpace(mailIdentityKey)
                    ? ComInteropScope.ResolveIdentityKey(mail, LogCategories.FileLink, "MailItem")
                    : mailIdentityKey.Trim();
                _inspectorIdentityKey = string.IsNullOrWhiteSpace(inspectorIdentityKey)
                    ? ResolveMailInspectorIdentityKey(mail)
                    : inspectorIdentityKey.Trim();
                _composeKey = BuildComposeKey(mail, _mailIdentityKey, _inspectorIdentityKey);

                _attachmentEvalTimer.Interval = ComposeAttachmentEvalDebounceMs;
                _attachmentEvalTimer.Tick += OnAttachmentEvalTimerTick;

                // Debounce long enough so Outlook can raise all BeforeAttachmentAdd callbacks
                // for a multi-select operation before we open the wizard.
                _beforeAddShareTimer.Interval = BeforeAddShareBatchDebounceMs;
                _beforeAddShareTimer.Tick += OnBeforeAddShareTimerTick;

                _cleanupGraceTimer.Interval = ComposeShareCleanupSendGraceMs;
                _cleanupGraceTimer.Tick += OnCleanupGraceTimerTick;

                _events = mail as Outlook.ItemEvents_10_Event;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                    + ").");
            }

            internal bool IsFor(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                    ShareLabel = result.FolderName ?? string.Empty,
                    CreatedUtc = DateTime.UtcNow
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

            internal void RegisterSeparatePasswordDispatch(FileLinkResult result, FileLinkRequest request, string passwordOnlyHtml)
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (result == null || request == null || string.IsNullOrWhiteSpace(passwordOnlyHtml))
                {
                    return;
                }

                string password = result.Password ?? string.Empty;
                if (string.IsNullOrWhiteSpace(password))
                {
                    return;
                }

                var entry = new SeparatePasswordDispatchEntry
                {
                    ShareLabel = result.FolderName ?? string.Empty,
                    ShareUrl = result.ShareUrl ?? string.Empty,
                    Password = password.Trim(),
                    Html = passwordOnlyHtml,
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
                    + ").");
            }

            private static string BuildComposeKey(Outlook.MailItem mail, string mailIdentityKey, string inspectorIdentityKey)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

            private void OnAttachmentAdd(Outlook.Attachment attachment)
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                _pendingAddedBatch.Add(new AttachmentBatchEntry
                {
                    Name = ReadAttachmentName(attachment),
                    SizeBytes = ReadAttachmentSizeBytes(attachment)
                });

                LogFileLink(
                    "Compose attachment added (composeKey="
                    + _composeKey
                    + ", pendingBatchCount="
                    + _pendingAddedBatch.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                ScheduleAttachmentEvaluation();
            }

            private void OnBeforeAttachmentAdd(Outlook.Attachment attachment, ref bool cancel)
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                try
                {
                    OutlookAttachmentAutomationGuardService.GuardState guardState;
                    if (_owner.TryGetAttachmentAutomationGuardState("before_add", _composeKey, out guardState))
                    {
                        return;
                    }

                    AttachmentAutomationSettings settings = ReadAttachmentAutomationSettings();
                    if (!settings.AlwaysConnector && !settings.OfferAboveEnabled)
                    {
                        LogFileLink(
                            "Compose before-attachment-add skipped (composeKey="
                            + _composeKey
                            + ", reason=automation_disabled).");
                        return;
                    }

                    AttachmentBatchEntry candidate;
                    string candidatePath;
                    bool candidatePathIsTemporary;
                    if (!TryBuildBeforeAddAttachmentCandidate(attachment, out candidate, out candidatePath, out candidatePathIsTemporary))
                    {
                        LogFileLink(
                            "Compose before-attachment-add preflight skipped (composeKey="
                            + _composeKey
                            + ", reason=candidate_unavailable).");
                        return;
                    }

                    LogFileLink(
                        "Compose before-attachment-add candidate resolved (composeKey="
                        + _composeKey
                        + ", name="
                        + (candidate.Name ?? string.Empty)
                        + ", sizeBytes="
                        + Math.Max(0, candidate.SizeBytes).ToString(CultureInfo.InvariantCulture)
                        + ", thresholdBytes="
                        + Math.Max(0, settings.ThresholdBytes).ToString(CultureInfo.InvariantCulture)
                        + ", mode="
                        + (settings.AlwaysConnector ? "always" : "threshold")
                        + ").");

                    bool shouldIntercept = settings.AlwaysConnector
                                           || (settings.OfferAboveEnabled && candidate.SizeBytes > settings.ThresholdBytes);
                    if (!shouldIntercept)
                    {
                        LogFileLink(
                            "Compose before-attachment-add allow host add (composeKey="
                            + _composeKey
                            + ", reason=below_threshold, sizeBytes="
                            + Math.Max(0, candidate.SizeBytes).ToString(CultureInfo.InvariantCulture)
                            + ", thresholdBytes="
                            + Math.Max(0, settings.ThresholdBytes).ToString(CultureInfo.InvariantCulture)
                            + ").");
                        return;
                    }

                    if (settings.AlwaysConnector)
                    {
                        cancel = true;
                        QueueBeforeAddAttachmentShareFlow("always_preadd", candidate, candidatePath, settings.ThresholdMb, candidatePathIsTemporary);
                        return;
                    }

                    if (_attachmentPromptOpen)
                    {
                        LogFileLink("Compose before-attachment-add prompt skipped (composeKey=" + _composeKey + ", reason=prompt_already_open).");
                        return;
                    }

                    string reasonText = string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.AttachmentPromptReason,
                        SizeFormatting.FormatMegabytes(Math.Max(0, candidate.SizeBytes)),
                        SizeFormatting.FormatMegabytes((long)settings.ThresholdMb * 1024L * 1024L),
                        string.IsNullOrWhiteSpace(candidate.Name) ? Strings.AttachmentPromptLastUnknown : candidate.Name,
                        SizeFormatting.FormatMegabytes(Math.Max(0, candidate.SizeBytes)));

                    _attachmentPromptOpen = true;
                    ComposeAttachmentPromptDecision decision;
                    try
                    {
                        decision = ComposeAttachmentPromptForm.ShowPrompt(
                            _owner.TryCreateMailInspectorDialogOwner(_mail),
                            reasonText);
                    }
                    finally
                    {
                        _attachmentPromptOpen = false;
                    }

                    if (decision == ComposeAttachmentPromptDecision.Share)
                    {
                        cancel = true;
                        LogFileLink(
                            "Compose before-attachment-add threshold decision (composeKey="
                            + _composeKey
                            + ", decision=share, attachment="
                            + (candidate.Name ?? string.Empty)
                            + ").");
                        StartBeforeAddAttachmentShareFlow("threshold_preadd", candidate, candidatePath, settings.ThresholdMb, candidatePathIsTemporary);
                        return;
                    }

                    if (decision == ComposeAttachmentPromptDecision.RemoveLast)
                    {
                        cancel = true;
                        LogFileLink(
                            "Compose before-attachment-add threshold decision (composeKey="
                            + _composeKey
                            + ", decision=cancel_add, attachment="
                            + (candidate.Name ?? string.Empty)
                            + ").");
                    }
                    else
                    {
                        LogFileLink(
                            "Compose before-attachment-add threshold decision (composeKey="
                            + _composeKey
                            + ", decision=keep_host_add, attachment="
                            + (candidate.Name ?? string.Empty)
                            + ").");
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Compose before-attachment-add preflight failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void OnPropertyChange(string name)
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                string propertyName = string.IsNullOrWhiteSpace(name) ? string.Empty : name.Trim();
                if (propertyName.IndexOf("Attach", StringComparison.OrdinalIgnoreCase) < 0
                    && !string.Equals(propertyName, "HasAttachment", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                LogFileLink(
                    "Compose property changed (composeKey="
                    + _composeKey
                    + ", property="
                    + propertyName
                    + ").");
                ScheduleAttachmentEvaluation();
            }

            private void ScheduleAttachmentEvaluation()
            {
                if (_disposed)
                {
                    return;
                }

                try
                {
                    _attachmentEvalTimer.Stop();
                    _attachmentEvalTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to schedule compose attachment evaluation (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void OnAttachmentEvalTimerTick(object sender, EventArgs e)
            {
                _attachmentEvalTimer.Stop();

                try
                {
                    EvaluateAttachmentAutomation();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Compose attachment evaluation failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void OnBeforeAddShareTimerTick(object sender, EventArgs e)
            {
                _beforeAddShareTimer.Stop();

                try
                {
                    RunQueuedBeforeAddAttachmentShareFlow();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Queued before-attachment-add share flow failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void EvaluateAttachmentAutomation()
            {
                if (_disposed || _attachmentSuppressed)
                {
                    return;
                }

                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("evaluate", _composeKey, out guardState))
                {
                    _pendingAddedBatch.Clear();
                    return;
                }

                AttachmentAutomationSettings settings = ReadAttachmentAutomationSettings();
                if (!settings.AlwaysConnector && !settings.OfferAboveEnabled)
                {
                    _pendingAddedBatch.Clear();
                    LogFileLink("Compose attachment evaluation skipped (composeKey=" + _composeKey + ", reason=automation_disabled).");
                    return;
                }

                List<AttachmentSnapshot> attachments = SnapshotAttachments();
                if (attachments.Count == 0)
                {
                    _pendingAddedBatch.Clear();
                    LogFileLink("Compose attachment evaluation skipped (composeKey=" + _composeKey + ", reason=no_attachments).");
                    return;
                }

                long totalBytes = SumAttachmentBytes(attachments);
                AttachmentBatchInfo lastAdded = BuildLastAddedBatchInfo(attachments);

                LogFileLink(
                    "Compose attachment evaluation (composeKey="
                    + _composeKey
                    + ", attachmentCount="
                    + attachments.Count.ToString(CultureInfo.InvariantCulture)
                    + ", totalBytes="
                    + totalBytes.ToString(CultureInfo.InvariantCulture)
                    + ", lastAddedCount="
                    + lastAdded.Count.ToString(CultureInfo.InvariantCulture)
                    + ", alwaysConnector="
                    + settings.AlwaysConnector.ToString(CultureInfo.InvariantCulture)
                    + ", offerAboveEnabled="
                    + settings.OfferAboveEnabled.ToString(CultureInfo.InvariantCulture)
                    + ", thresholdBytes="
                    + settings.ThresholdBytes.ToString(CultureInfo.InvariantCulture)
                    + ").");

                if (settings.AlwaysConnector)
                {
                    StartComposeAttachmentShareFlow("always", totalBytes, settings.ThresholdMb, lastAdded);
                    return;
                }

                if (!settings.OfferAboveEnabled || totalBytes <= settings.ThresholdBytes)
                {
                    return;
                }

                string reasonText = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.AttachmentPromptReason,
                    SizeFormatting.FormatMegabytes(totalBytes),
                    SizeFormatting.FormatMegabytes((long)settings.ThresholdMb * 1024L * 1024L),
                    string.IsNullOrWhiteSpace(lastAdded.Name) ? Strings.AttachmentPromptLastUnknown : lastAdded.Name,
                    SizeFormatting.FormatMegabytes(lastAdded.SizeBytes));
                if (_attachmentPromptOpen)
                {
                    LogFileLink("Compose attachment prompt skipped (composeKey=" + _composeKey + ", reason=prompt_already_open).");
                    return;
                }

                _attachmentPromptOpen = true;
                ComposeAttachmentPromptDecision decision;
                try
                {
                    decision = ComposeAttachmentPromptForm.ShowPrompt(
                        _owner.TryCreateMailInspectorDialogOwner(_mail),
                        reasonText);
                }
                finally
                {
                    _attachmentPromptOpen = false;
                }

                if (_owner.TryGetAttachmentAutomationGuardState("prompt_action", _composeKey, out guardState))
                {
                    return;
                }

                if (decision == ComposeAttachmentPromptDecision.Share)
                {
                    LogFileLink(
                        "Compose attachment threshold decision (composeKey="
                        + _composeKey
                        + ", decision=share, totalBytes="
                        + totalBytes.ToString(CultureInfo.InvariantCulture)
                        + ", thresholdBytes="
                        + settings.ThresholdBytes.ToString(CultureInfo.InvariantCulture)
                        + ").");
                    StartComposeAttachmentShareFlow("threshold", totalBytes, settings.ThresholdMb, lastAdded);
                    return;
                }

                LogFileLink(
                    "Compose attachment threshold decision (composeKey="
                    + _composeKey
                    + ", decision=remove_last, removeCount="
                    + lastAdded.Count.ToString(CultureInfo.InvariantCulture)
                    + ").");
                RemoveLastAddedAttachmentBatch(lastAdded);
            }

            private AttachmentAutomationSettings ReadAttachmentAutomationSettings()
            {
                _owner.EnsureSettingsLoaded();

                var settings = _owner._currentSettings ?? new AddinSettings();
                int thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(settings.SharingAttachmentsOfferAboveMb);
                bool alwaysConnector = settings.SharingAttachmentsAlwaysConnector;
                bool offerAboveEnabled = settings.SharingAttachmentsOfferAboveEnabled && !alwaysConnector;

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (_owner._currentSettings != null)
                {
                    var configuration = new TalkServiceConfiguration(
                        _owner._currentSettings.ServerUrl,
                        _owner._currentSettings.Username,
                        _owner._currentSettings.AppPassword);
                    BackendPolicyStatus policyStatus = _owner.FetchBackendPolicyStatus(configuration, "compose_attachment_evaluate");
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (policyStatus != null && policyStatus.PolicyActive)
                    {
                        bool policyBool;
                        int policyInt;

                        if (policyStatus.IsLocked("share", "attachments_always_via_ncconnector")
                            && policyStatus.TryGetPolicyBool("share", "attachments_always_via_ncconnector", out policyBool))
                        {
                            alwaysConnector = policyBool;
                        }

                        if (policyStatus.IsLocked("share", "attachments_min_size_mb"))
                        {
                            if (policyStatus.TryGetPolicyInt("share", "attachments_min_size_mb", out policyInt))
                            {
                                thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(policyInt);
                                offerAboveEnabled = true;
                            }
                            else if (policyStatus.HasPolicyKey("share", "attachments_min_size_mb"))
                            {
                                offerAboveEnabled = false;
                            }
                        }
                    }
                }

                return new AttachmentAutomationSettings
                {
                    AlwaysConnector = alwaysConnector,
                    OfferAboveEnabled = offerAboveEnabled,
                    ThresholdMb = thresholdMb,
                    ThresholdBytes = (long)thresholdMb * 1024L * 1024L
                };
            }

            private List<AttachmentSnapshot> SnapshotAttachments()
            {
                var snapshots = new List<AttachmentSnapshot>();
                Outlook.Attachments attachments = null;

                try
                {
                    attachments = _mail.Attachments;
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (attachments == null)
                    {
                        return snapshots;
                    }

                    int count = attachments.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                            if (attachment == null)
                            {
                                continue;
                            }

                            snapshots.Add(new AttachmentSnapshot
                            {
                                Index = index,
                                Name = ReadAttachmentName(attachment),
                                SizeBytes = ReadAttachmentSizeBytes(attachment)
                            });
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.FileLink,
                                "Failed to snapshot compose attachment (composeKey="
                                + _composeKey
                                + ", index="
                                + index.ToString(CultureInfo.InvariantCulture)
                                + ").",
                                ex);
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                attachment,
                                LogCategories.FileLink,
                                "Failed to release COM object (compose attachment snapshot).");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read compose attachments (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments,
                        LogCategories.FileLink,
                        "Failed to release COM object (compose attachments collection snapshot).");
                }

                return snapshots;
            }

            private static long SumAttachmentBytes(List<AttachmentSnapshot> snapshots)
            {
                long total = 0;
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (snapshots == null)
                {
                    return 0;
                }

                for (int i = 0; i < snapshots.Count; i++)
                {
                    total += Math.Max(0, snapshots[i].SizeBytes);
                }

                return total;
            }

            private AttachmentBatchInfo BuildLastAddedBatchInfo(List<AttachmentSnapshot> snapshots)
            {
                if (_pendingAddedBatch.Count > 0)
                {
                    long total = 0;
                    for (int i = 0; i < _pendingAddedBatch.Count; i++)
                    {
                        total += Math.Max(0, _pendingAddedBatch[i].SizeBytes);
                    }

                    var info = new AttachmentBatchInfo
                    {
                        Count = _pendingAddedBatch.Count,
                        Name = _pendingAddedBatch[_pendingAddedBatch.Count - 1].Name ?? string.Empty,
                        SizeBytes = total
                    };
                    _pendingAddedBatch.Clear();
                    return info;
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (snapshots == null || snapshots.Count == 0)
                {
                    return new AttachmentBatchInfo
                    {
                        Count = 0,
                        Name = string.Empty,
                        SizeBytes = 0
                    };
                }

                AttachmentSnapshot latest = snapshots[snapshots.Count - 1];
                return new AttachmentBatchInfo
                {
                    Count = 1,
                    Name = latest.Name ?? string.Empty,
                    SizeBytes = Math.Max(0, latest.SizeBytes)
                };
            }

            private void StartComposeAttachmentShareFlow(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            {
                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("start_flow", _composeKey, out guardState))
                {
                    return;
                }

                var selections = new List<FileLinkSelection>();
                var removeIndices = new List<int>();
                var tempFiles = new List<string>();

                CollectAttachmentSelectionsForShare(selections, removeIndices, tempFiles);
                if (selections.Count == 0)
                {
                    CleanupTemporaryFiles(tempFiles);
                    LogFileLink("Compose attachment flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                    return;
                }

                RemoveAttachmentsByIndices(removeIndices, "share_flow");

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "always" : trigger,
                    AttachmentTotalBytes = Math.Max(0, totalBytes),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = lastAdded != null ? (lastAdded.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = lastAdded != null ? Math.Max(0, lastAdded.SizeBytes) : 0
                };
                for (int i = 0; i < selections.Count; i++)
                {
                    launchOptions.InitialSelections.Add(new FileLinkSelection(selections[i].SelectionType, selections[i].LocalPath));
                }

                try
                {
                    bool wizardAccepted = _owner.RunFileLinkWizardForMail(_mail, launchOptions);
                    LogFileLink(
                        "Compose attachment flow completed (composeKey="
                        + _composeKey
                        + ", trigger="
                        + launchOptions.AttachmentTrigger
                        + ", queued="
                        + selections.Count.ToString(CultureInfo.InvariantCulture)
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                finally
                {
                    CleanupTemporaryFiles(tempFiles);
                }
            }

            private bool TryBuildBeforeAddAttachmentCandidate(
                Outlook.Attachment attachment,
                out AttachmentBatchEntry entry,
                out string path,
                out bool pathIsTemporary)
            {
                entry = new AttachmentBatchEntry
                {
                    Name = string.Empty,
                    SizeBytes = 0
                };
                path = string.Empty;
                pathIsTemporary = false;

                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (attachment == null)
                {
                    LogFileLink("Compose before-attachment-add candidate build skipped (composeKey=" + _composeKey + ", reason=attachment_null).");
                    return false;
                }

                path = ReadAttachmentPathName(attachment);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    path = TryResolveBeforeAddPathFromAttachmentMetadata(attachment);
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        string fallbackName = ReadAttachmentName(attachment);
                        string materializedPath;
                        // Some add paths expose no stable local source path pre-add.
                        // Materialize the candidate into temp so threshold logic can still evaluate it.
                        if (TryMaterializeBeforeAddAttachmentToTemp(attachment, fallbackName, out materializedPath))
                        {
                            path = materializedPath;
                            pathIsTemporary = true;
                            LogFileLink(
                                "Compose before-attachment-add candidate path materialized (composeKey="
                                + _composeKey
                                + ", source=save_as_file).");
                        }
                        else
                        {
                            long fallbackSize = ReadAttachmentSizeBytes(attachment);
                            LogFileLink(
                                "Compose before-attachment-add candidate path missing (composeKey="
                                + _composeKey
                                + ", hasPath="
                                + (!string.IsNullOrWhiteSpace(path)).ToString(CultureInfo.InvariantCulture)
                                + ", attachment="
                                + (fallbackName ?? string.Empty)
                                + ", sizeBytes="
                                + Math.Max(0, fallbackSize).ToString(CultureInfo.InvariantCulture)
                                + ").");
                            path = string.Empty;
                            return false;
                        }
                    }
                }

                entry.Name = ReadAttachmentName(attachment);
                if (string.IsNullOrWhiteSpace(entry.Name))
                {
                    entry.Name = Path.GetFileName(path) ?? string.Empty;
                }

                long attachmentSize = ReadAttachmentSizeBytes(attachment);
                long measuredPathSize = 0;
                try
                {
                    measuredPathSize = new FileInfo(path).Length;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read file size from before-attachment-add path (composeKey=" + _composeKey + ").",
                        ex);
                }

                if (measuredPathSize > 0)
                {
                    if (attachmentSize > 0
                        && measuredPathSize != attachmentSize
                        && DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink(
                            "Compose before-attachment-add size mismatch (composeKey="
                            + _composeKey
                            + ", attachmentSizeBytes="
                            + Math.Max(0, attachmentSize).ToString(CultureInfo.InvariantCulture)
                            + ", pathSizeBytes="
                            + Math.Max(0, measuredPathSize).ToString(CultureInfo.InvariantCulture)
                            + ").");
                    }

                    // Prefer measured file size when we have a materialized path.
                    entry.SizeBytes = measuredPathSize;
                }
                else
                {
                    entry.SizeBytes = attachmentSize;
                }

                return true;
            }

            private bool TryMaterializeBeforeAddAttachmentToTemp(Outlook.Attachment attachment, string attachmentName, out string path)
            {
                path = string.Empty;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (attachment == null)
                {
                    return false;
                }

                string safeName = FileLinkService.SanitizeComponent(attachmentName);
                if (string.IsNullOrWhiteSpace(safeName))
                {
                    safeName = "attachment.bin";
                }

                string tempRoot = Path.Combine(Path.GetTempPath(), "NCConnectorOutlook", "BeforeAdd", _composeKey);
                try
                {
                    Directory.CreateDirectory(tempRoot);
                    string targetPath = BuildUniqueFilePath(tempRoot, safeName);
                    attachment.SaveAsFile(targetPath);
                    if (!File.Exists(targetPath))
                    {
                        return false;
                    }

                    path = targetPath;
                    return true;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to materialize before-attachment-add file (composeKey="
                        + _composeKey
                        + ", attachment="
                        + safeName
                        + ").",
                        ex);
                    return false;
                }
            }

            private string TryResolveBeforeAddPathFromAttachmentMetadata(Outlook.Attachment attachment)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (attachment == null)
                {
                    return string.Empty;
                }

                string fileName = string.Empty;
                string displayName = string.Empty;

                try
                {
                    fileName = attachment.FileName ?? string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read before-attachment-add Attachment.FileName (composeKey=" + _composeKey + ").",
                        ex);
                }

                try
                {
                    displayName = attachment.DisplayName ?? string.Empty;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read before-attachment-add Attachment.DisplayName (composeKey=" + _composeKey + ").",
                        ex);
                }

                string[] candidates = { fileName, displayName };
                for (int i = 0; i < candidates.Length; i++)
                {
                    string raw = candidates[i];
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        continue;
                    }

                    string normalized = raw.Trim().Trim('"');
                    if (string.IsNullOrWhiteSpace(normalized))
                    {
                        continue;
                    }

                    try
                    {
                        if (Path.IsPathRooted(normalized) && File.Exists(normalized))
                        {
                            LogFileLink(
                                "Compose before-attachment-add candidate path resolved from metadata (composeKey="
                                + _composeKey
                                + ", source="
                                + (i == 0 ? "FileName" : "DisplayName")
                                + ").");
                            return normalized;
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to probe before-attachment-add metadata path (composeKey="
                            + _composeKey
                            + ", source="
                            + (i == 0 ? "FileName" : "DisplayName")
                            + ").",
                            ex);
                    }
                }

                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink(
                        "Compose before-attachment-add metadata path probe failed (composeKey="
                        + _composeKey
                        + ", fileName="
                        + (fileName ?? string.Empty)
                        + ", displayName="
                        + (displayName ?? string.Empty)
                        + ").");
                }

                return string.Empty;
            }

            private void StartBeforeAddAttachmentShareFlow(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow)
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    LogFileLink("Compose before-attachment-add share flow skipped (composeKey=" + _composeKey + ", reason=missing_local_path).");
                    return;
                }

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "threshold_preadd" : trigger,
                    AttachmentTotalBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = candidate != null ? (candidate.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0)
                };
                launchOptions.InitialSelections.Add(new FileLinkSelection(FileLinkSelectionType.File, localPath));

                try
                {
                    bool wizardAccepted = _owner.RunFileLinkWizardForMail(_mail, launchOptions);
                    LogFileLink(
                        "Compose before-attachment-add share flow completed (composeKey="
                        + _composeKey
                        + ", trigger="
                        + launchOptions.AttachmentTrigger
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ", attachment="
                        + (candidate != null ? (candidate.Name ?? string.Empty) : string.Empty)
                        + ").");
                }
                finally
                {
                    if (cleanupLocalPathAfterFlow)
                    {
                        CleanupTemporaryFiles(new List<string> { localPath });
                    }
                }
            }

            private void QueueBeforeAddAttachmentShareFlow(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow)
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    LogFileLink("Compose before-attachment-add share queue skipped (composeKey=" + _composeKey + ", reason=missing_local_path).");
                    return;
                }

                _pendingBeforeAddShareEntries.Add(new BeforeAddShareEntry
                {
                    Candidate = new AttachmentBatchEntry
                    {
                        Name = candidate != null ? (candidate.Name ?? string.Empty) : string.Empty,
                        SizeBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0)
                    },
                    LocalPath = localPath,
                    ThresholdMb = Math.Max(1, thresholdMb),
                    CleanupLocalPathAfterFlow = cleanupLocalPathAfterFlow,
                    Trigger = string.IsNullOrWhiteSpace(trigger) ? "always_preadd" : trigger
                });

                LogFileLink(
                    "Compose before-attachment-add candidate queued (composeKey="
                    + _composeKey
                    + ", queued="
                    + _pendingBeforeAddShareEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ", attachment="
                    + (candidate != null ? (candidate.Name ?? string.Empty) : string.Empty)
                    + ").");

                if (_beforeAddShareFlowRunning)
                {
                    return;
                }

                try
                {
                    _beforeAddShareTimer.Stop();
                    _beforeAddShareTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to schedule queued before-attachment-add share flow (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void RunQueuedBeforeAddAttachmentShareFlow()
            {
                if (_disposed || _beforeAddShareFlowRunning || _pendingBeforeAddShareEntries.Count == 0)
                {
                    return;
                }

                _beforeAddShareFlowRunning = true;
                var batch = new List<BeforeAddShareEntry>(_pendingBeforeAddShareEntries);
                _pendingBeforeAddShareEntries.Clear();

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = "always_preadd_batch"
                };
                var temporaryFiles = new List<string>();
                long totalBytes = 0;
                int thresholdMb = 1;
                string lastName = string.Empty;
                long lastSize = 0;

                try
                {
                    for (int i = 0; i < batch.Count; i++)
                    {
                        BeforeAddShareEntry entry = batch[i];
                        // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                        if (entry == null || string.IsNullOrWhiteSpace(entry.LocalPath) || !File.Exists(entry.LocalPath))
                        {
                            continue;
                        }

                        launchOptions.InitialSelections.Add(new FileLinkSelection(FileLinkSelectionType.File, entry.LocalPath));
                        totalBytes += Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                        thresholdMb = Math.Max(thresholdMb, Math.Max(1, entry.ThresholdMb));
                        lastName = entry.Candidate != null ? (entry.Candidate.Name ?? string.Empty) : string.Empty;
                        lastSize = Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                        if (entry.CleanupLocalPathAfterFlow)
                        {
                            temporaryFiles.Add(entry.LocalPath);
                        }
                    }

                    if (launchOptions.InitialSelections.Count == 0)
                    {
                        LogFileLink("Compose before-attachment-add queued share flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                        return;
                    }

                    launchOptions.AttachmentTotalBytes = Math.Max(0, totalBytes);
                    launchOptions.AttachmentThresholdMb = thresholdMb;
                    launchOptions.AttachmentLastName = lastName;
                    launchOptions.AttachmentLastSizeBytes = lastSize;

                    bool wizardAccepted = _owner.RunFileLinkWizardForMail(_mail, launchOptions);
                    LogFileLink(
                        "Compose before-attachment-add queued share flow completed (composeKey="
                        + _composeKey
                        + ", queued="
                        + batch.Count.ToString(CultureInfo.InvariantCulture)
                        + ", selected="
                        + launchOptions.InitialSelections.Count.ToString(CultureInfo.InvariantCulture)
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                finally
                {
                    CleanupTemporaryFiles(temporaryFiles);
                    _beforeAddShareFlowRunning = false;
                    if (_pendingBeforeAddShareEntries.Count > 0 && !_disposed)
                    {
                        _beforeAddShareTimer.Stop();
                        _beforeAddShareTimer.Start();
                    }
                }
            }

            private void CollectAttachmentSelectionsForShare(List<FileLinkSelection> selections, List<int> removeIndices, List<string> temporaryFiles)
            {
                Outlook.Attachments attachments = null;
                try
                {
                    attachments = _mail.Attachments;
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (attachments == null)
                    {
                        return;
                    }

                    int count = attachments.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                            if (attachment == null)
                            {
                                continue;
                            }

                            string attachmentName = ReadAttachmentName(attachment);
                            string localPath;
                            if (!TryResolveAttachmentLocalPath(attachment, attachmentName, temporaryFiles, out localPath))
                            {
                                continue;
                            }

                            selections.Add(new FileLinkSelection(FileLinkSelectionType.File, localPath));
                            removeIndices.Add(index);
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.FileLink,
                                "Failed to collect compose attachment for sharing (composeKey="
                                + _composeKey
                                + ", index="
                                + index.ToString(CultureInfo.InvariantCulture)
                                + ").",
                                ex);
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                attachment,
                                LogCategories.FileLink,
                                "Failed to release COM object (compose attachment collect).");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to collect compose attachments for sharing (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments,
                        LogCategories.FileLink,
                        "Failed to release COM object (compose attachments collection collect).");
                }
            }

            private bool TryResolveAttachmentLocalPath(Outlook.Attachment attachment, string attachmentName, List<string> temporaryFiles, out string localPath)
            {
                localPath = string.Empty;
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                if (attachment == null)
                {
                    return false;
                }

                string pathName = ReadAttachmentPathName(attachment);
                if (!string.IsNullOrWhiteSpace(pathName) && File.Exists(pathName))
                {
                    localPath = pathName;
                    return true;
                }

                string safeName = FileLinkService.SanitizeComponent(attachmentName);
                if (string.IsNullOrWhiteSpace(safeName))
                {
                    safeName = "attachment.bin";
                }

                string tempRoot = Path.Combine(Path.GetTempPath(), "NCConnectorOutlook", "Attachments", _composeKey);
                try
                {
                    Directory.CreateDirectory(tempRoot);
                    string targetPath = BuildUniqueFilePath(tempRoot, safeName);
                    attachment.SaveAsFile(targetPath);
                    if (!File.Exists(targetPath))
                    {
                        return false;
                    }

                    localPath = targetPath;
                    // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                    if (temporaryFiles != null)
                    {
                        temporaryFiles.Add(targetPath);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to materialize compose attachment file (composeKey="
                        + _composeKey
                        + ", attachment="
                        + safeName
                        + ").",
                        ex);
                    return false;
                }
            }

            private static string BuildUniqueFilePath(string directory, string fileName)
            {
                string candidate = Path.Combine(directory, fileName);
                if (!File.Exists(candidate))
                {
                    return candidate;
                }

                for (int suffix = 1; suffix < 1000; suffix++)
                {
                    string slotDirectory = Path.Combine(
                        directory,
                        "dup_" + suffix.ToString(CultureInfo.InvariantCulture));
                    candidate = Path.Combine(slotDirectory, fileName);
                    if (!File.Exists(candidate))
                    {
                        Directory.CreateDirectory(slotDirectory);
                        return candidate;
                    }
                }

                string fallbackDirectory = Path.Combine(directory, Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(fallbackDirectory);
                return Path.Combine(fallbackDirectory, fileName);
            }

            private void RemoveAttachmentsByIndices(List<int> indices, string reason)
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (indices == null || indices.Count == 0)
                {
                    return;
                }

                indices.Sort();

                Outlook.Attachments attachments = null;
                int removed = 0;
                _attachmentSuppressed = true;
                _pendingAddedBatch.Clear();

                try
                {
                    attachments = _mail.Attachments;
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (attachments == null)
                    {
                        return;
                    }

                    for (int index = indices.Count - 1; index >= 0; index--)
                    {
                        int attachmentIndex = indices[index];
                        if (attachmentIndex <= 0 || attachmentIndex > attachments.Count)
                        {
                            continue;
                        }

                        attachments.Remove(attachmentIndex);
                        removed++;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to remove compose attachments (composeKey="
                        + _composeKey
                        + ", reason="
                        + (reason ?? string.Empty)
                        + ").",
                        ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments,
                        LogCategories.FileLink,
                        "Failed to release COM object (compose attachments collection remove).");
                    _attachmentSuppressed = false;
                }

                LogFileLink(
                    "Compose attachments removed (composeKey="
                    + _composeKey
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", requested="
                    + indices.Count.ToString(CultureInfo.InvariantCulture)
                    + ", removed="
                    + removed.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void RemoveLastAddedAttachmentBatch(AttachmentBatchInfo lastAdded)
            {
                int removeCount = lastAdded != null ? Math.Max(1, lastAdded.Count) : 1;

                Outlook.Attachments attachments = null;
                int removed = 0;
                _attachmentSuppressed = true;
                _pendingAddedBatch.Clear();

                try
                {
                    attachments = _mail.Attachments;
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (attachments == null || attachments.Count <= 0)
                    {
                        return;
                    }

                    int totalCount = attachments.Count;
                    int effectiveCount = Math.Min(totalCount, removeCount);
                    for (int i = 0; i < effectiveCount; i++)
                    {
                        attachments.Remove(attachments.Count);
                        removed++;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to remove last attachment batch (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments,
                        LogCategories.FileLink,
                        "Failed to release COM object (compose attachments collection remove_last).");
                    _attachmentSuppressed = false;
                }

                LogFileLink(
                    "Compose attachment batch removed (composeKey="
                    + _composeKey
                    + ", requested="
                    + removeCount.ToString(CultureInfo.InvariantCulture)
                    + ", removed="
                    + removed.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void CleanupTemporaryFiles(List<string> temporaryFiles)
            {
                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
                if (temporaryFiles == null || temporaryFiles.Count == 0)
                {
                    return;
                }

                for (int i = 0; i < temporaryFiles.Count; i++)
                {
                    string path = temporaryFiles[i];
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    try
                    {
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Failed to delete temporary compose attachment file '" + path + "'.",
                            ex);
                    }
                }
            }

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
                var comException = ex as COMException;
                // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
                if (comException == null)
                {
                    return false;
                }

                uint errorCode = unchecked((uint)comException.ErrorCode);
                return (errorCode & 0xFFFFu) == 0x010Au;
            }

            private static string ToHResultHex(Exception ex)
            {
                // Null bedeutet hier "kein passender Fehlerkontext"; Auswertung bleibt absichtlich defensiv.
                if (ex == null)
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
                    _owner.TryDeleteComposeShareFolder(
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
                bcc = string.Empty;

                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (recipients == null)
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
                            recipient = recipients[i];
                            // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                            if (recipient == null)
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
                _owner.DispatchSeparatePasswordMailQueue(_composeKey, queue);
            }

            private string ReadMailRecipientList(string fieldName)
            {
                try
                {
                    // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
                    if (_mail == null)
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

            private static string ReadAttachmentName(Outlook.Attachment attachment)
            {
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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
                // Outlook/COM kann hier null liefern (Lifecycle/Interop-Randfall); fail-soft behalten.
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

                if (_pendingBeforeAddShareEntries.Count > 0)
                {
                    var pendingTemporaryFiles = new List<string>();
                    for (int i = 0; i < _pendingBeforeAddShareEntries.Count; i++)
                    {
                        BeforeAddShareEntry pendingEntry = _pendingBeforeAddShareEntries[i];
                        // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
                    _attachmentEvalTimer.Dispose();
                    _beforeAddShareTimer.Dispose();
                    _cleanupGraceTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to dispose compose timers.", ex);
                }

                // Defensiver Null-Guard: dieser Pfad soll bei unvollständigem Runtime-Zustand kontrolliert abbrechen.
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
