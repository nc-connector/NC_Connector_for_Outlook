// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            private void OnAttachmentAdd(Outlook.Attachment attachment)
            {
                if (!CanApplyComposeChanges
                    || _attachmentSuppressed
                    || IsHiddenAttachment(attachment))
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
                if (!CanApplyComposeChanges
                    || cancel
                    || _attachmentSuppressed
                    || IsHiddenAttachment(attachment))
                {
                    return;
                }
                AttachmentAutomationSettings settings =
                    ReadAttachmentAutomationSettings();
                string candidatePath = string.Empty;
                bool candidatePathIsTemporary = false;
                bool candidateQueued = false;
                try
                {
                    OutlookAttachmentAutomationGuardService.GuardState guardState;
                    if (_owner.TryGetAttachmentAutomationGuardState("before_add", _composeKey, out guardState))
                    {
                        return;
                    }

                    if (!settings.AlwaysConnector && !settings.OfferAboveEnabled)
                    {
                        LogFileLink(
                            "Compose before-attachment-add skipped (composeKey="
                            + _composeKey
                            + ", reason=automation_disabled).");
                        return;
                    }

                    AttachmentBatchEntry candidate;
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

                    // The event argument is borrowed; acquire an owned reference after the host adds it.
                    candidate.OriginalAttachment = attachment;
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
                    if (PauseUnavailableAttachmentAutomation(settings, candidate.SizeBytes)
                        || !CanApplyComposeChanges)
                    {
                        return;
                    }
                    if (settings.AlwaysConnector)
                    {
                        QueueBeforeAddAttachmentShareFlow("always_preadd", candidate, candidatePath, settings.ThresholdMb, candidatePathIsTemporary);
                        candidateQueued = true;
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
                        if (!CanApplyComposeChanges)
                        {
                            return;
                        }
                        decision = ComposeAttachmentPromptForm.ShowPrompt(
                            _owner._mailInteropController.TryCreateMailInspectorDialogOwner(_mail),
                            reasonText);
                    }
                    finally
                    {
                        _attachmentPromptOpen = false;
                    }
                    if (!CanApplyComposeChanges)
                    {
                        return;
                    }
                    if (decision == ComposeAttachmentPromptDecision.Share)
                    {
                        LogFileLink(
                            "Compose before-attachment-add threshold decision (composeKey="
                            + _composeKey
                            + ", decision=share, attachment="
                            + (candidate.Name ?? string.Empty)
                            + ").");
                        QueueBeforeAddAttachmentShareFlow("threshold_preadd", candidate, candidatePath, settings.ThresholdMb, candidatePathIsTemporary);
                        candidateQueued = true;
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
                        "Compose before-attachment-add preflight failed; retaining the native attachment (composeKey=" + _composeKey + ").",
                        ex);
                }
                finally
                {
                    if (candidatePathIsTemporary && !candidateQueued)
                    {
                        CleanupTemporaryFiles(new List<string> { candidatePath });
                    }
                }
            }

            private void OnPropertyChange(string name)
            {
                if (!CanApplyComposeChanges)
                {
                    return;
                }
                string propertyName = string.IsNullOrWhiteSpace(name) ? string.Empty : name.Trim();
                if (IsEmailSignaturePropertyChange(propertyName))
                {
                    string signatureReason = "property_" + propertyName;
                    if (_attachmentSuppressed)
                    {
                        DeferEmailSignatureApplication(signatureReason, "attachment_suppression");
                    }
                    else if (_composeSurfaceState == ComposeSurfaceState.Detached)
                    {
                        DeferEmailSignatureApplication(signatureReason, "detached_surface");
                    }
                    else
                    {
                        ScheduleEmailSignatureApplication(signatureReason);
                    }
                }
                if (_attachmentSuppressed)
                {
                    return;
                }
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
                if (!CanApplyComposeChanges)
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

            private async void OnAttachmentEvalTimerTick(object sender, EventArgs e)
            {
                _attachmentEvalTimer.Stop();

                try
                {
                    await EvaluateAttachmentAutomationAsync();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Compose attachment evaluation failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private async void OnBeforeAddShareTimerTick(object sender, EventArgs e)
            {
                _beforeAddShareTimer.Stop();
                if (!CanApplyComposeChanges)
                {
                    return;
                }

                try
                {
                    await RunQueuedBeforeAddAttachmentShareFlowAsync();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Queued before-attachment-add share flow failed (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private async Task EvaluateAttachmentAutomationAsync()
            {
                if (!CanApplyComposeChanges || _attachmentSuppressed)
                {
                    return;
                }
                if (_beforeAddShareFlowRunning || _pendingBeforeAddShareEntries.Count > 0)
                {
                    _pendingAddedBatch.Clear();
                    LogFileLink("Compose attachment evaluation skipped (composeKey=" + _composeKey + ", reason=before_add_flow_pending).");
                    return;
                }

                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("evaluate", _composeKey, out guardState))
                {
                    _pendingAddedBatch.Clear();
                    return;
                }

                AttachmentAutomationSettings settings = await ReadAttachmentAutomationSettingsAsync();
                if (!CanApplyComposeChanges || _attachmentSuppressed)
                {
                    return;
                }
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
                    if (PauseUnavailableAttachmentAutomation(settings, totalBytes) || !CanApplyComposeChanges) { return; }
                    await StartComposeAttachmentShareFlowAsync("always", totalBytes, settings.ThresholdMb, lastAdded);
                    return;
                }
                if (!settings.OfferAboveEnabled || totalBytes <= settings.ThresholdBytes)
                {
                    return;
                }
                if (PauseUnavailableAttachmentAutomation(settings, totalBytes) || !CanApplyComposeChanges) { return; }
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
                    if (!CanApplyComposeChanges)
                    {
                        return;
                    }
                    decision = ComposeAttachmentPromptForm.ShowPrompt(
                        _owner._mailInteropController.TryCreateMailInspectorDialogOwner(_mail),
                        reasonText);
                }
                finally
                {
                    _attachmentPromptOpen = false;
                }
                if (!CanApplyComposeChanges
                    || _owner.TryGetAttachmentAutomationGuardState("prompt_action", _composeKey, out guardState))
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
                    await StartComposeAttachmentShareFlowAsync("threshold", totalBytes, settings.ThresholdMb, lastAdded);
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

        }
    }
}

