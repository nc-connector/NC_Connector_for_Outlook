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
            private sealed class AttachmentSnapshot
            {
                internal int Index { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentBatchInfo
            {
                internal int Count { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentBatchEntry
            {
                internal Outlook.Attachment OriginalAttachment { get; set; }

                internal string Name { get; set; }

                internal long SizeBytes { get; set; }
            }

            private sealed class AttachmentShareOriginal
            {
                internal Outlook.Attachment Attachment { get; set; }

                internal string LocalPath { get; set; }
            }

            private int CountPolicyRelevantAttachments(out long totalBytes)
            {
                Outlook.Attachments attachments = null;
                int relevant = 0;
                totalBytes = 0;
                try
                {
                    attachments = _mail != null ? _mail.Attachments : null;
                    int count = attachments != null ? attachments.Count : 0;
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[i];
                            if (attachment != null
                                && !IsHiddenAttachment(attachment))
                            {
                                relevant++;
                                totalBytes += Math.Max(0, attachment.Size);
                            }
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                attachment,
                                LogCategories.FileLink,
                                "Failed to release send-gate attachment.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to inspect attachments for required routing (composeKey=" + _composeKey + ").",
                        ex);
                    totalBytes = long.MaxValue;
                    return int.MaxValue;
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments,
                        LogCategories.FileLink,
                        "Failed to release send-gate Attachments collection.");
                }
                return relevant;
            }

            private static bool IsHiddenAttachment(Outlook.Attachment attachment)
            {
                Outlook.PropertyAccessor accessor = null;
                try
                {
                    accessor = attachment != null
                        ? attachment.PropertyAccessor
                        : null;
                    object value = accessor != null
                        ? accessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B")
                        : null;
                    return value is bool && (bool)value;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        accessor,
                        LogCategories.FileLink,
                        "Failed to release attachment PropertyAccessor.");
                }
            }

            private List<AttachmentSnapshot> SnapshotAttachments()
            {
                var snapshots = new List<AttachmentSnapshot>();
                Outlook.Attachments attachments = null;

                try
                {
                    attachments = _mail.Attachments;
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
                            if (attachment == null)
                            {
                                continue;
                            }
                            if (IsHiddenAttachment(attachment))
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
                    AttachmentBatchEntry latestBatchEntry =
                        _pendingAddedBatch[
                            _pendingAddedBatch.Count - 1];
                    var info = new AttachmentBatchInfo
                    {
                        Count = _pendingAddedBatch.Count,
                        Name = latestBatchEntry.Name ?? string.Empty,
                        SizeBytes = Math.Max(
                            0,
                            latestBatchEntry.SizeBytes)
                    };
                    _pendingAddedBatch.Clear();
                    return info;
                }
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
                if (attachment == null)
                {
                    return false;
                }
                string safeName = FileLinkPath.SanitizeComponent(attachmentName);
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

            private void CollectAttachmentSelectionsForShare(List<FileLinkSelection> selections, List<AttachmentShareOriginal> originals, List<string> temporaryFiles)
            {
                Outlook.Attachments attachments = null;
                try
                {
                    attachments = _mail.Attachments;
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
                            if (attachment == null)
                            {
                                continue;
                            }
                            if (IsHiddenAttachment(attachment))
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
                            originals.Add(new AttachmentShareOriginal { Attachment = attachment, LocalPath = localPath });
                            attachment = null;
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
                if (attachment == null)
                {
                    return false;
                }
                string safeName = FileLinkPath.SanitizeComponent(attachmentName);
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

            private void CaptureBeforeAddAttachmentOriginal(
                AttachmentBatchEntry candidate,
                List<FileLinkSelection> selections,
                List<AttachmentShareOriginal> originals,
                List<string> temporaryFiles)
            {
                if (!CanApplyComposeChanges || candidate == null || candidate.OriginalAttachment == null)
                {
                    return;
                }

                Outlook.Attachments attachments = null;
                try
                {
                    attachments = _mail != null ? _mail.Attachments : null;
                    int count = attachments != null ? attachments.Count : 0;
                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Attachment attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            if (attachment == null || IsHiddenAttachment(attachment)
                                || !ComInteropScope.AreSameObject(
                                    attachment, candidate.OriginalAttachment,
                                    LogCategories.FileLink, "Attachment", "BeforeAttachmentAdd"))
                            {
                                continue;
                            }
                            foreach (AttachmentShareOriginal existing in originals)
                            {
                                if (ComInteropScope.AreSameObject(
                                    attachment, existing.Attachment,
                                    LogCategories.FileLink, "Attachment", "CapturedAttachment"))
                                {
                                    return;
                                }
                            }
                            string localPath;
                            if (!TryResolveAttachmentLocalPath(attachment, ReadAttachmentName(attachment), temporaryFiles, out localPath))
                            {
                                return;
                            }
                            selections.Add(new FileLinkSelection(FileLinkSelectionType.File, localPath));
                            originals.Add(new AttachmentShareOriginal { Attachment = attachment, LocalPath = localPath });
                            attachment = null;
                            return;
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(
                                attachment, LogCategories.FileLink,
                                "Failed to release before-add attachment capture.");
                        }
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(
                        attachments, LogCategories.FileLink,
                        "Failed to release before-add attachment collection.");
                }
            }

            private void RemoveAttachmentOriginals(
                List<AttachmentShareOriginal> originals,
                IList<string> sharedLocalPaths,
                bool cancelledByUser)
            {
                if (!CanApplyComposeChanges
                    || (!cancelledByUser && (sharedLocalPaths == null || sharedLocalPaths.Count == 0)))
                {
                    return;
                }
                var shared = new HashSet<FileLinkSelection>(FileLinkSelection.IdentityComparer);
                if (!cancelledByUser)
                {
                    foreach (string path in sharedLocalPaths)
                    {
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            shared.Add(new FileLinkSelection(FileLinkSelectionType.File, path));
                        }
                    }
                }
                string reason = cancelledByUser ? "cancelled_originals" : "shared_originals";
                int removedCount = 0;
                _attachmentSuppressed = true;
                try
                {
                    foreach (AttachmentShareOriginal original in originals)
                    {
                        if (original.Attachment == null
                            || (!cancelledByUser
                                && !shared.Contains(new FileLinkSelection(FileLinkSelectionType.File, original.LocalPath)))
                            || IsHiddenAttachment(original.Attachment))
                        {
                            continue;
                        }
                        try
                        {
                            if (!CanApplyComposeChanges)
                            {
                                return;
                            }
                            original.Attachment.Delete();
                            removedCount++;
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.FileLink,
                                "Failed to remove an original attachment (composeKey=" + _composeKey
                                    + ", reason=" + reason + ").",
                                ex);
                        }
                    }
                }
                finally
                {
                    LogFileLink("Compose attachment originals removed (composeKey=" + _composeKey
                        + ", reason=" + reason
                        + ", removed=" + removedCount.ToString(CultureInfo.InvariantCulture) + ").");
                    EndAttachmentSuppression(reason);
                }
            }

            private static void ReleaseAttachmentShareOriginals(List<AttachmentShareOriginal> originals)
            {
                foreach (AttachmentShareOriginal original in originals)
                {
                    try
                    {
                        ComInteropScope.TryRelease(
                            original.Attachment, LogCategories.FileLink,
                            "Failed to release an original shared attachment.");
                    }
                    finally
                    {
                        original.Attachment = null;
                    }
                }
                originals.Clear();
            }

            private void RemoveAttachmentsByIndices(List<int> indices, string reason)
            {
                if (!CanApplyComposeChanges || indices == null || indices.Count == 0)
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

                        if (!CanApplyComposeChanges)
                        {
                            return;
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
                    EndAttachmentSuppression("remove_by_index");
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
                if (!CanApplyComposeChanges)
                {
                    return;
                }
                int removeCount = lastAdded != null ? Math.Max(1, lastAdded.Count) : 1;
                List<AttachmentSnapshot> attachments = SnapshotAttachments();
                var removeIndices = new List<int>();
                for (int index = attachments.Count - 1;
                     index >= 0 && removeIndices.Count < removeCount;
                     index--)
                {
                    removeIndices.Add(attachments[index].Index);
                }
                RemoveAttachmentsByIndices(removeIndices, "remove_last_batch");
            }

            private void CleanupTemporaryFiles(List<string> temporaryFiles)
            {
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

        }
    }
}
