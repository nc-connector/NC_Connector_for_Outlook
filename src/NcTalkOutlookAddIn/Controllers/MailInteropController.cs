// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    // Encapsulates Outlook Mail/Inspector interop and HTML body insertion bridges.
    // Keeps COM/editor-specific behavior out of the add-in orchestration root.
    internal sealed class MailInteropController
    {
        private const string OutlookAutoSignatureBookmarkName = "_MailAutoSig";
        private const string OutlookOriginalMessageBookmarkName = "_MailOriginal";
        private const string ManagedEmailSignatureBookmarkName = "NcConnectorSignature";
        private const int OutlookOriginalMessageProtectedGap = 2;
        private readonly NextcloudTalkAddIn _owner;

        internal sealed class EmailSignatureReconcileResult
        {
            internal bool Success { get; set; }

            internal bool Changed { get; set; }

            internal bool Managed { get; set; }

            internal string Source { get; set; }
        }

        private enum EmailSignatureReconcileMode
        {
            Apply,
            ClearManaged,
            ClearInitial
        }

        internal MailInteropController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal static string ResolveMailInspectorIdentityKey(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                return ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                uint errorCode = unchecked((uint)ex.ErrorCode);
                if ((errorCode & 0xFFFFu) == 0x0108u)
                {
                    NextcloudTalkAddIn.LogFileLinkMessage(
                        "MailItem.GetInspector unavailable while resolving compose inspector identity (hresult=0x"
                        + errorCode.ToString("X8", CultureInfo.InvariantCulture)
                        + ").");
                }
                else
                {
                    DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.GetInspector for compose identity.", ex);
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to read MailItem.GetInspector for compose identity.", ex);
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release compose Inspector COM object.");
            }
        }

        internal IWin32Window TryCreateMailInspectorDialogOwner(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;
                if (inspector == null)
                {
                    return null;
                }
                int hwnd = ReadInspectorWindowHandle(inspector);
                return hwnd > 0 ? new NativeWindowOwner(new IntPtr(hwnd)) : null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve compose prompt owner inspector.", ex);
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release compose prompt owner Inspector COM object.");
            }
        }

        private static int ReadInspectorWindowHandle(Outlook.Inspector inspector)
        {
            if (inspector == null)
            {
                return 0;
            }
            foreach (string propertyName in new[] { "HWND", "Hwnd" })
            {
                try
                {
                    PropertyInfo property = inspector.GetType().GetProperty(propertyName);
                    if (property == null)
                    {
                        continue;
                    }

                    object value = property.GetValue(inspector, null);
                    if (value == null)
                    {
                        continue;
                    }
                    int hwnd;
                    if (int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out hwnd) && hwnd > 0)
                    {
                        return hwnd;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to read inspector window handle property '" + propertyName + "'.",
                        ex);
                }
            }
            return 0;
        }

        internal Outlook.MailItem GetActiveMailItem()
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            if (application == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = application.ActiveInspector();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook ActiveInspector.", ex);
                inspector = null;
            }
            if (inspector != null)
            {
                try
                {
                    return inspector.CurrentItem as Outlook.MailItem;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read CurrentItem from ActiveInspector.", ex);
                }
            }

            Outlook.Explorer explorer = null;
            try
            {
                explorer = application.ActiveExplorer();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook ActiveExplorer.", ex);
                explorer = null;
            }
            if (explorer != null)
            {
                object inlineResponse = null;
                try
                {
                    inlineResponse = explorer.ActiveInlineResponse;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read ActiveInlineResponse from Explorer.", ex);
                    inlineResponse = null;
                }
                var mailItem = inlineResponse as Outlook.MailItem;
                if (mailItem != null)
                {
                    return mailItem;
                }
            }
            return null;
        }

        internal string ResolveActiveInspectorIdentityKey()
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            if (application == null)
            {
                return string.Empty;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = application.ActiveInspector();
                return ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "ActiveInspector");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to resolve active inspector identity key.", ex);
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(inspector, LogCategories.FileLink, "Failed to release active Inspector COM object.");
            }
        }

        internal bool IsActiveInlineResponse(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return false;
            }

            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            Outlook.Explorer explorer = null;
            Outlook.MailItem activeInlineMail = null;
            try
            {
                explorer = application != null ? application.ActiveExplorer() : null;
                activeInlineMail = explorer != null ? explorer.ActiveInlineResponse as Outlook.MailItem : null;
                return ComInteropScope.AreSameObject(mail, activeInlineMail, LogCategories.FileLink, "MailItem", "ActiveInlineResponse");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to check active inline response.", ex);
                return false;
            }
            finally
            {
                if (!ReferenceEquals(activeInlineMail, mail))
                {
                    ComInteropScope.TryRelease(activeInlineMail, LogCategories.FileLink, "Failed to release ActiveInlineResponse MailItem COM object.");
                }
                ComInteropScope.TryRelease(explorer, LogCategories.FileLink, "Failed to release active Explorer COM object.");
            }
        }

        internal void InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {
            if (mail == null || string.IsNullOrWhiteSpace(html))
            {
                return;
            }
            if (IsActiveInlineResponse(mail))
            {
                if (TryInsertHtmlIntoActiveInlineResponseWordEditor(mail, html))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into inline response (ActiveInlineResponseWordEditor).");
                    return;
                }

                DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML into inline response: inline WordEditor insertion failed.");
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, "inline WordEditor insertion failed"),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            if (TryInsertHtmlIntoInspectorWordEditor(mail, html))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (WordEditor InsertFile primary).");
                return;
            }

            if (TryInsertHtmlIntoMailBody(mail, html))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (HTMLBody compatibility fallback).");
                return;
            }

            DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML into mail: all insertion paths exhausted.");
            MessageBox.Show(
                string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, "all insertion paths exhausted"),
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        internal void InsertPlainTextIntoMail(Outlook.MailItem mail, string plainText)
        {
            if (mail == null || string.IsNullOrWhiteSpace(plainText))
            {
                return;
            }

            try
            {
                string insertText = PlainTextUtilities.NormalizeCrLfAndTrim(plainText);
                if (!string.IsNullOrEmpty(insertText))
                {
                    insertText += "\r\n\r\n";
                }

                string source;
                if (TryInsertPlainTextViaWordEditor(mail, insertText, out source))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Inserted plain-text share block into mail (source=" + source + ").");
                    return;
                }

                throw new InvalidOperationException("Outlook WordEditor insertion point unavailable.");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain-text share block into mail.", ex);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private bool TryInsertPlainTextViaWordEditor(Outlook.MailItem mail, string text, out string source)
        {
            source = string.Empty;
            if (mail == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            bool activeInlineMatches = IsActiveInlineResponse(mail);
            OutlookWordEditorContext context;

            try
            {
                if (activeInlineMatches)
                {
                    if (!OutlookWordEditorContext.TryOpenInline(application, mail, "plain-text share insert", null, out context))
                    {
                        return false;
                    }
                }
                else if (!OutlookWordEditorContext.TryOpenInspector(mail, "plain-text share insert", out context))
                {
                    return false;
                }

                using (context)
                {
                    source = context.Source ?? string.Empty;
                    if (activeInlineMatches)
                    {
                        int replyCursorStart = context.GetDocumentStart(0);
                        context.SetSelectionRange(replyCursorStart, replyCursorStart);
                        context.TypeParagraph();
                        context.TypeParagraph();
                        context.TypeText(text);
                        context.SetSelectionRange(replyCursorStart, replyCursorStart);
                    }
                    else
                    {
                        context.TypeText(text);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain text via Outlook WordEditor.", ex);
                return false;
            }
        }

        internal static bool IsPlainTextMail(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return false;
            }

            try
            {
                return mail.BodyFormat == Outlook.OlBodyFormat.olFormatPlain;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read MailItem.BodyFormat.", ex);
                return false;
            }
        }

        private bool TryInsertHtmlIntoActiveInlineResponseWordEditor(Outlook.MailItem mail, string html)
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            OutlookWordEditorContext context;
            string tempHtmlPath = null;

            try
            {
                if (!OutlookWordEditorContext.TryOpenInline(application, mail, "inline share insert", null, out context))
                {
                    return false;
                }

                using (context)
                {
                    InsertHtmlIntoWordEditor(context, html, "nc4ol-inline-share-", ref tempHtmlPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert HTML into inline response Word editor.", ex);
                return false;
            }
            finally
            {
                TryDeleteTemporaryHtmlFile(tempHtmlPath, "Failed to delete temporary inline share HTML file.");
            }
        }

        private static bool TryInsertHtmlIntoInspectorWordEditor(Outlook.MailItem mail, string html)
        {
            OutlookWordEditorContext context;
            string tempHtmlPath = null;

            try
            {
                if (!OutlookWordEditorContext.TryOpenInspector(mail, "HTML share insert", out context))
                {
                    return false;
                }

                using (context)
                {
                    InsertHtmlIntoWordEditor(context, html, "nc4ol-share-", ref tempHtmlPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert HTML via inspector WordEditor.", ex);
                return false;
            }
            finally
            {
                TryDeleteTemporaryHtmlFile(tempHtmlPath, "Failed to delete temporary share HTML file.");
            }
        }

        private static void InsertHtmlIntoWordEditor(
            OutlookWordEditorContext context,
            string html,
            string tempFilePrefix,
            ref string tempHtmlPath)
        {
            int replyCursorStart = context.GetDocumentStart(0);
            context.SetSelectionRange(replyCursorStart, replyCursorStart);
            context.TypeParagraph();
            context.TypeParagraph();

            tempHtmlPath = Path.Combine(
                Path.GetTempPath(),
                tempFilePrefix + Guid.NewGuid().ToString("N") + ".html");
            File.WriteAllText(tempHtmlPath, EnsureHtmlDocumentForWordInsert(html), new UTF8Encoding(true));
            context.InsertFile(tempHtmlPath);
            context.SetSelectionRange(replyCursorStart, replyCursorStart);
        }

        private static void TryDeleteTemporaryHtmlFile(string tempHtmlPath, string failureMessage)
        {
            if (string.IsNullOrWhiteSpace(tempHtmlPath))
            {
                return;
            }

            try
            {
                File.Delete(tempHtmlPath);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, failureMessage, ex);
            }
        }

        private static bool TryInsertHtmlIntoMailBody(Outlook.MailItem mail, string html)
        {
            try
            {
                string existing = mail.HTMLBody ?? string.Empty;
                string insertHtml = "<br><br>" + html;
                int bodyTagIndex = existing.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyTagIndex >= 0)
                {
                    int bodyTagEnd = existing.IndexOf(">", bodyTagIndex);
                    if (bodyTagEnd >= 0)
                    {
                        mail.HTMLBody = existing.Insert(bodyTagEnd + 1, insertHtml);
                    }
                    else
                    {
                        mail.HTMLBody = insertHtml + existing;
                    }
                }
                else
                {
                    mail.HTMLBody = insertHtml + existing;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML via HTMLBody path: " + ex.Message);
                return false;
            }
        }

        private static int ClampWordPosition(int value, int documentEnd)
        {
            if (value < 0)
            {
                return 0;
            }

            if (documentEnd >= 0 && value > documentEnd)
            {
                return documentEnd;
            }

            return value;
        }

        internal EmailSignatureReconcileResult ApplyManagedEmailSignature(
            Outlook.MailItem mail,
            bool isInlineResponse,
            bool isPlainText,
            string signatureContent,
            bool isReplyOrForward,
            string composeKey,
            string operation,
            string inlineExplorerIdentityKey = null)
        {
            return ReconcileEmailSignatureWordSlot(
                mail,
                isInlineResponse,
                isPlainText,
                signatureContent,
                isReplyOrForward,
                composeKey,
                operation,
                inlineExplorerIdentityKey,
                EmailSignatureReconcileMode.Apply);
        }

        internal EmailSignatureReconcileResult ClearManagedEmailSignature(
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            string operation,
            string inlineExplorerIdentityKey = null)
        {
            return ReconcileEmailSignatureWordSlot(
                mail,
                isInlineResponse,
                false,
                string.Empty,
                false,
                composeKey,
                operation,
                inlineExplorerIdentityKey,
                EmailSignatureReconcileMode.ClearManaged);
        }

        internal EmailSignatureReconcileResult ClearInitialEmailSignatureSlot(
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            string operation,
            string inlineExplorerIdentityKey = null)
        {
            return ReconcileEmailSignatureWordSlot(
                mail,
                isInlineResponse,
                false,
                string.Empty,
                false,
                composeKey,
                operation,
                inlineExplorerIdentityKey,
                EmailSignatureReconcileMode.ClearInitial);
        }

        private EmailSignatureReconcileResult ReconcileEmailSignatureWordSlot(
            Outlook.MailItem mail,
            bool isInlineResponse,
            bool isPlainText,
            string signatureContent,
            bool isReplyOrForward,
            string composeKey,
            string operation,
            string inlineExplorerIdentityKey,
            EmailSignatureReconcileMode mode)
        {
            var result = new EmailSignatureReconcileResult
            {
                Source = "word_editor_unavailable"
            };
            if (mail == null)
            {
                return result;
            }

            OutlookWordEditorContext context;
            if (!TryOpenEmailSignatureWordEditor(
                    mail,
                    isInlineResponse,
                    composeKey,
                    operation,
                    inlineExplorerIdentityKey,
                    out context))
            {
                return result;
            }

            using (context)
            {
                object bookmarks = null;
                string selectionBookmarkName = string.Empty;
                string previousSlotBookmarkName = string.Empty;
                int originalSelectionStart = context.Selection != null
                    ? OutlookWordEditorContext.GetIntProperty(context.Selection, "Start")
                    : context.GetDocumentStart(0);
                int originalSelectionEnd = context.Selection != null
                    ? OutlookWordEditorContext.GetIntProperty(context.Selection, "End")
                    : originalSelectionStart;
                try
                {
                    bookmarks = OutlookWordEditorContext.GetProperty(context.Document, "Bookmarks");
                    if (bookmarks == null)
                    {
                        result.Source = "bookmarks_unavailable";
                        return result;
                    }
                    TryShowHiddenBookmarks(bookmarks);
                    selectionBookmarkName = CaptureEmailSignatureSelectionBookmark(
                        context.Document,
                        bookmarks,
                        originalSelectionStart,
                        originalSelectionEnd);

                    int slotStart;
                    int slotEnd;
                    bool slotIsTable;
                    string slotSource;
                    bool hasManagedSlot = TryGetEmailSignatureBookmarkSlot(
                        context.Document,
                        bookmarks,
                        ManagedEmailSignatureBookmarkName,
                        out slotStart,
                        out slotEnd,
                        out slotIsTable,
                        out slotSource);

                    if (mode == EmailSignatureReconcileMode.ClearManaged)
                    {
                        if (!hasManagedSlot)
                        {
                            result.Success = true;
                            result.Source = "managed_not_found";
                            return result;
                        }
                        result.Success = TryDeleteEmailSignatureSlot(
                            context.Document,
                            slotStart,
                            slotEnd,
                            slotIsTable,
                            slotSource);
                        if (result.Success)
                        {
                            TryDeleteEmailSignatureBookmark(bookmarks, ManagedEmailSignatureBookmarkName);
                        }
                        result.Changed = result.Success;
                        result.Managed = false;
                        result.Source = slotSource;
                        return result;
                    }

                    bool hasInitialSlot = hasManagedSlot;
                    if (!hasInitialSlot)
                    {
                        hasInitialSlot = TryGetEmailSignatureBookmarkSlot(
                            context.Document,
                            bookmarks,
                            OutlookAutoSignatureBookmarkName,
                            out slotStart,
                            out slotEnd,
                            out slotIsTable,
                            out slotSource);
                    }

                    if (mode == EmailSignatureReconcileMode.ClearInitial)
                    {
                        if (!hasInitialSlot)
                        {
                            result.Success = true;
                            result.Source = "initial_not_found";
                            return result;
                        }
                        result.Success = TryDeleteEmailSignatureSlot(
                            context.Document,
                            slotStart,
                            slotEnd,
                            slotIsTable,
                            slotSource);
                        if (result.Success)
                        {
                            TryDeleteEmailSignatureBookmark(
                                bookmarks,
                                hasManagedSlot
                                    ? ManagedEmailSignatureBookmarkName
                                    : OutlookAutoSignatureBookmarkName);
                        }
                        result.Changed = result.Success;
                        result.Managed = false;
                        result.Source = slotSource;
                        return result;
                    }

                    if (string.IsNullOrWhiteSpace(signatureContent))
                    {
                        result.Source = "signature_content_empty";
                        return result;
                    }

                    bool safeFallback = false;
                    bool addQuoteGap = false;
                    if (!hasInitialSlot)
                    {
                        int unusedQuoteBoundaryPosition;
                        if (!TryResolveSafeEmailSignatureInsertionPoint(
                                context,
                                bookmarks,
                                isReplyOrForward,
                                false,
                                0,
                                0,
                                out slotStart,
                                out unusedQuoteBoundaryPosition,
                                out addQuoteGap,
                                out slotSource))
                        {
                            result.Source = slotSource;
                            return result;
                        }
                        slotEnd = slotStart;
                        slotIsTable = false;
                        safeFallback = true;
                    }

                    string resolvedSlotSource = slotSource;
                    bool initialSlotManaged = hasInitialSlot
                                              && string.Equals(
                                                  slotSource,
                                                  "managed",
                                                  StringComparison.Ordinal);
                    int insertionPosition = hasInitialSlot ? slotEnd : slotStart;
                    if (hasInitialSlot)
                    {
                        int safeInsertionPosition;
                        int quoteBoundaryPosition;
                        bool safeAddQuoteGap;
                        string safeSource;
                        bool hasSafeInsertionPoint = TryResolveSafeEmailSignatureInsertionPoint(
                                context,
                                bookmarks,
                                isReplyOrForward,
                                true,
                                slotStart,
                                slotEnd,
                                out safeInsertionPosition,
                                out quoteBoundaryPosition,
                                out safeAddQuoteGap,
                                out safeSource);
                        if (isReplyOrForward && !hasSafeInsertionPoint)
                        {
                            result.Source = safeSource;
                            return result;
                        }

                        bool hasMeaningfulTextBetween = false;
                        if (hasSafeInsertionPoint
                            && safeInsertionPosition > slotEnd
                            && !TryHasMeaningfulEmailSignatureText(
                                    context.Document,
                                    slotEnd,
                                    safeInsertionPosition,
                                    out hasMeaningfulTextBetween))
                        {
                            result.Source = "authored_content_probe_failed";
                            return result;
                        }

                        EmailSignatureSlotPlacementDecision placementDecision = hasSafeInsertionPoint
                            ? EmailSignatureSlotPlacementPolicy.Resolve(
                                isReplyOrForward,
                                slotStart,
                                slotEnd,
                                safeInsertionPosition,
                                quoteBoundaryPosition,
                                hasMeaningfulTextBetween)
                            : EmailSignatureSlotPlacementDecision.KeepExistingSlot;
                        if (placementDecision == EmailSignatureSlotPlacementDecision.UnsafeQuoteBoundaryOverlap)
                        {
                            result.Source = "initial_slot_crosses_" + safeSource;
                            DiagnosticsLogger.Log(
                                LogCategories.Core,
                                "Email signature slot overlaps the reply quote boundary; reconciliation skipped (source="
                                + result.Source
                                + ", slot="
                                + (slotSource ?? "n/a")
                                + ", safe="
                                + safeInsertionPosition.ToString(CultureInfo.InvariantCulture)
                                + ", boundary="
                                + quoteBoundaryPosition.ToString(CultureInfo.InvariantCulture)
                                + ", composeKey="
                                + (composeKey ?? string.Empty)
                                + ").");
                            return result;
                        }

                        if (placementDecision == EmailSignatureSlotPlacementDecision.MoveToSafeInsertionPoint)
                        {
                            bool movedAboveQuoteBoundary = isReplyOrForward
                                                           && safeInsertionPosition < slotEnd;
                            insertionPosition = safeInsertionPosition;
                            addQuoteGap = safeAddQuoteGap;
                            safeFallback = true;
                            resolvedSlotSource += movedAboveQuoteBoundary
                                ? "_moved_above_" + safeSource
                                : "_moved_to_" + safeSource;
                            DiagnosticsLogger.Log(
                                LogCategories.Core,
                                (movedAboveQuoteBoundary
                                    ? "Email signature slot moved above quoted content (source="
                                    : "Email signature slot moved below authored content (source=")
                                + resolvedSlotSource
                                + ", composeKey="
                                + (composeKey ?? string.Empty)
                                + ").");
                        }
                    }

                    if (hasInitialSlot && safeFallback && insertionPosition < slotStart)
                    {
                        previousSlotBookmarkName = "NcSigPrevious" + Guid.NewGuid().ToString("N").Substring(0, 12);
                        if (!TryAddEmailSignatureBookmark(
                                context.Document,
                                bookmarks,
                                previousSlotBookmarkName,
                                slotStart,
                                slotEnd))
                        {
                            result.Source = "initial_slot_track_failed";
                            return result;
                        }
                    }

                    int stagedMutationStart = insertionPosition;
                    if (safeFallback
                        && !TryAddMissingEmailSignatureLeadingParagraphs(
                            context.Document,
                            ref insertionPosition,
                            out slotSource))
                    {
                        TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            string.Empty,
                            stagedMutationStart,
                            insertionPosition,
                            "leading_gap_rollback");
                        result.Source = slotSource;
                        return result;
                    }

                    int insertedStart;
                    int insertedEnd;
                    if (!TryInsertEmailSignatureContent(
                            context,
                            insertionPosition,
                            isPlainText,
                            signatureContent,
                            out insertedStart,
                            out insertedEnd,
                            out slotSource))
                    {
                        TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            string.Empty,
                            stagedMutationStart,
                            insertionPosition,
                            "content_insert_rollback");
                        result.Source = slotSource;
                        return result;
                    }

                    bool includeManagedQuoteGap = isReplyOrForward
                                                  && ((safeFallback && addQuoteGap)
                                                      || (hasInitialSlot
                                                          && string.Equals(
                                                              resolvedSlotSource,
                                                              "managed",
                                                              StringComparison.Ordinal)));
                    if (includeManagedQuoteGap)
                    {
                        int gapEnd;
                        if (!TryInsertEmailSignatureParagraphs(
                                context.Document,
                                insertedEnd,
                                1,
                                out slotSource,
                                out gapEnd))
                        {
                            TryRollbackStagedEmailSignatureMutation(
                                context.Document,
                                bookmarks,
                                string.Empty,
                                stagedMutationStart,
                                insertedEnd,
                                "quote_gap_rollback");
                            result.Source = slotSource;
                            return result;
                        }
                        insertedEnd = gapEnd;
                    }

                    string stagedBookmarkName = "NcSigStage" + Guid.NewGuid().ToString("N").Substring(0, 12);
                    if (!TryAddEmailSignatureBookmark(
                            context.Document,
                            bookmarks,
                            stagedBookmarkName,
                            stagedMutationStart,
                            insertedEnd))
                    {
                        TryDeleteEmailSignatureBookmark(bookmarks, stagedBookmarkName);
                        TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            string.Empty,
                            stagedMutationStart,
                            insertedEnd,
                            "staged_insert");
                        result.Source = "managed_bookmark_stage_failed";
                        return result;
                    }

                    int currentSlotStart = slotStart;
                    int currentSlotEnd = slotEnd;
                    if (!string.IsNullOrWhiteSpace(previousSlotBookmarkName)
                        && !TryGetBookmarkRange(
                            bookmarks,
                            previousSlotBookmarkName,
                            out currentSlotStart,
                            out currentSlotEnd))
                    {
                        TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            stagedBookmarkName,
                            stagedMutationStart,
                            insertedEnd,
                            "staged_insert");
                        result.Source = "initial_slot_track_lost";
                        return result;
                    }

                    bool replacingManaged = initialSlotManaged;
                    if (replacingManaged)
                    {
                        TryDeleteEmailSignatureBookmark(bookmarks, ManagedEmailSignatureBookmarkName);
                    }

                    if (!TryAddEmailSignatureBookmark(
                            context.Document,
                            bookmarks,
                            ManagedEmailSignatureBookmarkName,
                            insertedStart,
                            insertedEnd))
                    {
                        TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            stagedBookmarkName,
                            stagedMutationStart,
                            insertedEnd,
                            "staged_insert");
                        TryDeleteEmailSignatureBookmark(bookmarks, ManagedEmailSignatureBookmarkName);
                        if (replacingManaged)
                        {
                            if (TryResolveTrackedEmailSignatureSlot(
                                    bookmarks,
                                    previousSlotBookmarkName,
                                    slotStart,
                                    slotEnd,
                                    out currentSlotStart,
                                    out currentSlotEnd))
                            {
                                TryAddEmailSignatureBookmark(
                                    context.Document,
                                    bookmarks,
                                    ManagedEmailSignatureBookmarkName,
                                    currentSlotStart,
                                    currentSlotEnd);
                            }
                        }
                        result.Source = "managed_bookmark_add_failed";
                        return result;
                    }

                    if (hasInitialSlot
                        && !TryDeleteEmailSignatureSlot(
                            context.Document,
                            currentSlotStart,
                            currentSlotEnd,
                            slotIsTable,
                            resolvedSlotSource))
                    {
                        bool stagedMutationRolledBack = TryRollbackStagedEmailSignatureMutation(
                            context.Document,
                            bookmarks,
                            stagedBookmarkName,
                            stagedMutationStart,
                            insertedEnd,
                            "managed_rollback");
                        if (stagedMutationRolledBack)
                        {
                            TryDeleteEmailSignatureBookmark(bookmarks, ManagedEmailSignatureBookmarkName);
                            if (replacingManaged
                                && TryResolveTrackedEmailSignatureSlot(
                                    bookmarks,
                                    previousSlotBookmarkName,
                                    slotStart,
                                    slotEnd,
                                    out currentSlotStart,
                                    out currentSlotEnd))
                            {
                                TryAddEmailSignatureBookmark(
                                    context.Document,
                                    bookmarks,
                                    ManagedEmailSignatureBookmarkName,
                                    currentSlotStart,
                                    currentSlotEnd);
                            }
                        }
                        result.Source = "initial_slot_delete_failed";
                        return result;
                    }

                    TryDeleteEmailSignatureBookmark(bookmarks, stagedBookmarkName);
                    int managedStart;
                    int managedEnd;
                    if (!TryGetBookmarkRange(
                            bookmarks,
                            ManagedEmailSignatureBookmarkName,
                            out managedStart,
                            out managedEnd))
                    {
                        result.Source = "managed_bookmark_lost";
                        return result;
                    }

                    result.Success = true;
                    result.Changed = true;
                    result.Managed = true;
                    result.Source = hasInitialSlot ? resolvedSlotSource : "safe_" + resolvedSlotSource;
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Email signature Word slot reconciled (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", editor="
                        + (context.Source ?? "n/a")
                        + ", body="
                        + (isPlainText ? "plain" : "formatted")
                        + ", source="
                        + (result.Source ?? "n/a")
                        + ").");
                    return result;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to reconcile email signature Word slot (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ").",
                        ex);
                    result.Source = "exception";
                    return result;
                }
                finally
                {
                    if (!string.IsNullOrWhiteSpace(previousSlotBookmarkName))
                    {
                        TryDeleteEmailSignatureBookmark(bookmarks, previousSlotBookmarkName);
                    }
                    RestoreEmailSignatureSelection(
                        context,
                        bookmarks,
                        selectionBookmarkName,
                        originalSelectionStart,
                        originalSelectionEnd);
                    ComInteropScope.TryRelease(bookmarks, LogCategories.Core, "Failed to release email signature Word bookmarks COM object.");
                }
            }
        }

        private bool TryOpenEmailSignatureWordEditor(
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            string operation,
            string inlineExplorerIdentityKey,
            out OutlookWordEditorContext context)
        {
            context = null;
            bool activeInline = _owner != null && _owner.IsActiveInlineResponse(mail);
            if (isInlineResponse || activeInline)
            {
                if (OutlookWordEditorContext.TryOpenInline(
                        _owner != null ? _owner.OutlookApplication : null,
                        mail,
                        operation,
                        composeKey,
                        inlineExplorerIdentityKey,
                        out context))
                {
                    return true;
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Email signature Word editor is not ready for the tracked inline response (composeKey="
                    + (composeKey ?? string.Empty)
                    + ").");
                return false;
            }

            if (OutlookWordEditorContext.TryOpenInspector(mail, operation, out context))
            {
                return true;
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Email signature Word editor is unavailable (composeKey="
                + (composeKey ?? string.Empty)
                + ", inlineState="
                + isInlineResponse.ToString(CultureInfo.InvariantCulture)
                + ").");
            return false;
        }

        private static bool TryGetEmailSignatureBookmarkSlot(
            object wordEditor,
            object bookmarks,
            string bookmarkName,
            out int start,
            out int end,
            out bool isTable,
            out string source)
        {
            start = 0;
            end = 0;
            isTable = false;
            source = string.Equals(bookmarkName, ManagedEmailSignatureBookmarkName, StringComparison.Ordinal)
                ? "managed"
                : "mail_auto_sig";
            if (!TryGetBookmarkRange(bookmarks, bookmarkName, out start, out end))
            {
                source += "_not_found";
                return false;
            }

            if (string.Equals(bookmarkName, OutlookAutoSignatureBookmarkName, StringComparison.Ordinal))
            {
                int tableStart;
                int tableEnd;
                if (TryExpandRangeToContainingTable(wordEditor, start, end, out tableStart, out tableEnd))
                {
                    start = tableStart;
                    end = tableEnd;
                    isTable = true;
                    source = "mail_auto_sig_table";
                }
            }
            return true;
        }

        private static bool TryResolveSafeEmailSignatureInsertionPoint(
            OutlookWordEditorContext context,
            object bookmarks,
            bool isReplyOrForward,
            bool hasExcludedRange,
            int excludedStart,
            int excludedEnd,
            out int insertionPosition,
            out int quoteBoundaryPosition,
            out bool addQuoteGap,
            out string source)
        {
            insertionPosition = 0;
            quoteBoundaryPosition = 0;
            addQuoteGap = false;
            source = "safe_slot_unavailable";
            if (context == null || context.Document == null)
            {
                return false;
            }

            int documentStart = context.GetDocumentStart(0);
            if (!isReplyOrForward)
            {
                int documentEnd = context.GetDocumentEnd(documentStart);
                insertionPosition = Math.Max(documentStart, documentEnd - 1);
                quoteBoundaryPosition = insertionPosition;
                source = "document_end";
                return true;
            }

            int originalStart;
            int originalEnd;
            if (TryGetBookmarkRange(
                    bookmarks,
                    OutlookOriginalMessageBookmarkName,
                    out originalStart,
                    out originalEnd))
            {
                quoteBoundaryPosition = originalStart;
                insertionPosition = Math.Max(
                    documentStart,
                    originalStart - OutlookOriginalMessageProtectedGap);
                addQuoteGap = true;
                source = "mail_original";
                return true;
            }

            int separatorStart;
            if (TryFindInlineQuoteSeparatorStart(
                    context.Document,
                    documentStart,
                    hasExcludedRange,
                    excludedStart,
                    excludedEnd,
                    out separatorStart))
            {
                quoteBoundaryPosition = separatorStart;
                insertionPosition = Math.Max(documentStart, separatorStart);
                addQuoteGap = true;
                source = "quote_separator";
                return true;
            }

            source = "quote_boundary_unavailable";
            return false;
        }

        private static bool TryAddMissingEmailSignatureLeadingParagraphs(
            object wordEditor,
            ref int insertionPosition,
            out string source)
        {
            source = "leading_gap";
            if (wordEditor == null)
            {
                source = "word_editor_unavailable";
                return false;
            }

            int existingParagraphMarks;
            if (!TryCountTrailingParagraphMarks(
                    wordEditor,
                    insertionPosition,
                    out existingParagraphMarks))
            {
                source = "leading_gap_read_failed";
                return false;
            }

            int missing = Math.Max(0, 2 - existingParagraphMarks);
            if (missing == 0)
            {
                source = "leading_gap_preserved";
                return true;
            }

            int insertedEnd;
            if (!TryInsertEmailSignatureParagraphs(
                    wordEditor,
                    insertionPosition,
                    missing,
                    out source,
                    out insertedEnd))
            {
                return false;
            }
            insertionPosition = insertedEnd;
            source = "leading_gap_added_" + missing.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryHasMeaningfulEmailSignatureText(
            object wordEditor,
            int start,
            int end,
            out bool hasMeaningfulText)
        {
            hasMeaningfulText = false;
            if (wordEditor == null || end <= start)
            {
                return true;
            }

            object range = null;
            try
            {
                range = OutlookWordEditorContext.InvokeMethod(
                    wordEditor,
                    "Range",
                    new object[] { start, end });
                if (range == null)
                {
                    return false;
                }

                string text = Convert.ToString(
                    OutlookWordEditorContext.GetProperty(range, "Text"),
                    CultureInfo.InvariantCulture) ?? string.Empty;
                hasMeaningfulText = text.Trim(
                    '\r',
                    '\n',
                    '\t',
                    ' ',
                    '\u00A0').Length > 0;
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to inspect authored content after the email signature slot.",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    range,
                    LogCategories.Core,
                    "Failed to release authored-content inspection range COM object.");
            }
        }

        private static bool TryCountTrailingParagraphMarks(
            object wordEditor,
            int insertionPosition,
            out int paragraphMarks)
        {
            paragraphMarks = 0;
            object content = null;
            object range = null;
            try
            {
                content = OutlookWordEditorContext.GetProperty(wordEditor, "Content");
                int documentStart = content != null
                    ? OutlookWordEditorContext.GetIntProperty(content, "Start")
                    : 0;
                int readStart = Math.Max(documentStart, insertionPosition - 256);
                range = OutlookWordEditorContext.InvokeMethod(
                    wordEditor,
                    "Range",
                    new object[] { readStart, Math.Max(readStart, insertionPosition) });
                string text = range != null
                    ? Convert.ToString(OutlookWordEditorContext.GetProperty(range, "Text"), CultureInfo.InvariantCulture)
                    : string.Empty;
                if (string.IsNullOrEmpty(text))
                {
                    return true;
                }

                for (int i = text.Length - 1; i >= 0; i--)
                {
                    char current = text[i];
                    if (current == '\r')
                    {
                        paragraphMarks++;
                        continue;
                    }
                    if (current == '\n'
                        || current == '\t'
                        || current == ' '
                        || current == '\u00A0')
                    {
                        continue;
                    }
                    break;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to inspect paragraphs before the email signature slot.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release email signature gap range COM object.");
                ComInteropScope.TryRelease(content, LogCategories.Core, "Failed to release email signature gap content COM object.");
            }
        }

        private static bool TryInsertEmailSignatureParagraphs(
            object wordEditor,
            int position,
            int count,
            out string source)
        {
            int insertedEnd;
            return TryInsertEmailSignatureParagraphs(
                wordEditor,
                position,
                count,
                out source,
                out insertedEnd);
        }

        private static bool TryInsertEmailSignatureParagraphs(
            object wordEditor,
            int position,
            int count,
            out string source,
            out int insertedEnd)
        {
            source = "paragraph_gap";
            insertedEnd = position;
            if (wordEditor == null || count <= 0)
            {
                return true;
            }

            object range = null;
            try
            {
                range = OutlookWordEditorContext.InvokeMethod(
                    wordEditor,
                    "Range",
                    new object[] { position, position });
                if (range == null)
                {
                    source = "paragraph_range_unavailable";
                    return false;
                }
                OutlookWordEditorContext.SetProperty(range, "Text", new string('\r', count));
                insertedEnd = OutlookWordEditorContext.GetIntProperty(range, "End");
                if (insertedEnd <= position)
                {
                    insertedEnd = position + count;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert email signature paragraph gap.", ex);
                source = "paragraph_insert_failed";
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release email signature paragraph range COM object.");
            }
        }

        private static bool TryInsertEmailSignatureContent(
            OutlookWordEditorContext context,
            int position,
            bool isPlainText,
            string content,
            out int insertedStart,
            out int insertedEnd,
            out string source)
        {
            insertedStart = position;
            insertedEnd = position;
            source = isPlainText ? "plain_text" : "formatted_html";
            if (context == null || context.Document == null || string.IsNullOrWhiteSpace(content))
            {
                source = "content_unavailable";
                return false;
            }

            if (isPlainText)
            {
                object range = null;
                try
                {
                    range = OutlookWordEditorContext.InvokeMethod(
                        context.Document,
                        "Range",
                        new object[] { position, position });
                    if (range == null)
                    {
                        source = "plain_range_unavailable";
                        return false;
                    }
                    OutlookWordEditorContext.SetProperty(range, "Text", content);
                    insertedEnd = OutlookWordEditorContext.GetIntProperty(range, "End");
                    if (insertedEnd <= insertedStart)
                    {
                        insertedEnd = insertedStart + content.Length;
                    }
                    return insertedEnd > insertedStart;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain-text email signature through Word range.", ex);
                    source = "plain_insert_failed";
                    return false;
                }
                finally
                {
                    ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release plain-text email signature range COM object.");
                }
            }

            string tempHtmlPath = null;
            try
            {
                int documentEndBefore = context.GetDocumentEnd(position);
                if (!context.SetSelectionRange(position, position))
                {
                    source = "selection_unavailable";
                    return false;
                }

                tempHtmlPath = Path.Combine(
                    Path.GetTempPath(),
                    "nc4ol-signature-" + Guid.NewGuid().ToString("N") + ".html");
                File.WriteAllText(
                    tempHtmlPath,
                    EnsureHtmlDocumentForWordInsert(content),
                    new UTF8Encoding(true));
                context.InsertFile(tempHtmlPath);

                int selectionStart = OutlookWordEditorContext.GetIntProperty(context.Selection, "Start");
                int selectionEnd = OutlookWordEditorContext.GetIntProperty(context.Selection, "End");
                int documentEndAfter = context.GetDocumentEnd(documentEndBefore);
                insertedEnd = Math.Max(selectionStart, selectionEnd);
                if (insertedEnd <= insertedStart)
                {
                    insertedEnd = insertedStart + Math.Max(0, documentEndAfter - documentEndBefore);
                }
                if (insertedEnd <= insertedStart)
                {
                    source = "formatted_insert_empty";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert formatted email signature through Word editor.", ex);
                source = "formatted_insert_failed";
                return false;
            }
            finally
            {
                TryDeleteTemporaryHtmlFile(tempHtmlPath, "Failed to delete temporary email signature HTML file.");
            }
        }

        private static bool TryAddEmailSignatureBookmark(
            object wordEditor,
            object bookmarks,
            string bookmarkName,
            int start,
            int end)
        {
            if (wordEditor == null
                || bookmarks == null
                || string.IsNullOrWhiteSpace(bookmarkName)
                || end < start)
            {
                return false;
            }

            object range = null;
            object bookmark = null;
            try
            {
                range = OutlookWordEditorContext.InvokeMethod(
                    wordEditor,
                    "Range",
                    new object[] { start, end });
                if (range == null)
                {
                    return false;
                }
                bookmark = OutlookWordEditorContext.InvokeMethod(
                    bookmarks,
                    "Add",
                    new object[] { bookmarkName, range });
                int verifiedStart;
                int verifiedEnd;
                return TryGetBookmarkRange(
                    bookmarks,
                    bookmarkName,
                    out verifiedStart,
                    out verifiedEnd)
                       && verifiedEnd >= verifiedStart;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to add email signature Word bookmark (name=" + bookmarkName + ").", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(bookmark, LogCategories.Core, "Failed to release added email signature bookmark COM object.");
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release added email signature bookmark range COM object.");
            }
        }

        private static bool TryResolveTrackedEmailSignatureSlot(
            object bookmarks,
            string trackingBookmarkName,
            int fallbackStart,
            int fallbackEnd,
            out int start,
            out int end)
        {
            start = fallbackStart;
            end = fallbackEnd;
            if (string.IsNullOrWhiteSpace(trackingBookmarkName))
            {
                return true;
            }

            int trackedStart;
            int trackedEnd;
            if (!TryGetBookmarkRange(
                    bookmarks,
                    trackingBookmarkName,
                    out trackedStart,
                    out trackedEnd))
            {
                return false;
            }

            start = trackedStart;
            end = trackedEnd;
            return true;
        }

        private static bool TryRollbackStagedEmailSignatureMutation(
            object wordEditor,
            object bookmarks,
            string trackingBookmarkName,
            int fallbackStart,
            int fallbackEnd,
            string source)
        {
            int rollbackStart = fallbackStart;
            int rollbackEnd = fallbackEnd;
            if (!string.IsNullOrWhiteSpace(trackingBookmarkName))
            {
                int trackedStart;
                int trackedEnd;
                if (TryGetBookmarkRange(
                        bookmarks,
                        trackingBookmarkName,
                        out trackedStart,
                        out trackedEnd))
                {
                    rollbackStart = trackedStart;
                    rollbackEnd = trackedEnd;
                }
                else
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Staged email signature bookmark was unavailable during rollback; using the captured mutation range (source="
                        + (source ?? "n/a")
                        + ").");
                }
            }

            bool deleted = TryDeleteEmailSignatureSlot(
                wordEditor,
                rollbackStart,
                rollbackEnd,
                false,
                source);
            if (!string.IsNullOrWhiteSpace(trackingBookmarkName))
            {
                TryDeleteEmailSignatureBookmark(bookmarks, trackingBookmarkName);
            }
            return deleted;
        }

        private static bool TryDeleteEmailSignatureBookmark(object bookmarks, string bookmarkName)
        {
            if (bookmarks == null || string.IsNullOrWhiteSpace(bookmarkName))
            {
                return false;
            }

            object bookmark = null;
            try
            {
                int start;
                int end;
                if (!TryGetBookmarkRange(bookmarks, bookmarkName, out start, out end))
                {
                    return false;
                }
                bookmark = OutlookWordEditorContext.InvokeMethod(
                    bookmarks,
                    "Item",
                    new object[] { bookmarkName });
                if (bookmark == null)
                {
                    return false;
                }
                OutlookWordEditorContext.InvokeMethod(bookmark, "Delete", null);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to delete email signature Word bookmark (name=" + bookmarkName + ").", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(bookmark, LogCategories.Core, "Failed to release deleted email signature bookmark COM object.");
            }
        }

        private static bool TryDeleteEmailSignatureSlot(
            object wordEditor,
            int start,
            int end,
            bool isTable,
            string source)
        {
            if (wordEditor == null || end < start)
            {
                return false;
            }

            if (end == start)
            {
                return true;
            }

            if (isTable)
            {
                if (TryDeleteContainingTableAtRange(wordEditor, start, end))
                {
                    return true;
                }
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Email signature table deletion failed (source=" + (source ?? "n/a") + ").");
                return false;
            }

            object range = null;
            try
            {
                range = OutlookWordEditorContext.InvokeMethod(
                    wordEditor,
                    "Range",
                    new object[] { start, end });
                if (range == null)
                {
                    return false;
                }
                OutlookWordEditorContext.InvokeMethod(range, "Delete", null);
                int remainingStart = OutlookWordEditorContext.GetIntProperty(range, "Start");
                int remainingEnd = OutlookWordEditorContext.GetIntProperty(range, "End");
                return remainingEnd <= remainingStart;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to delete email signature Word range (source=" + (source ?? "n/a") + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release deleted email signature range COM object.");
            }
        }

        private static string CaptureEmailSignatureSelectionBookmark(
            object wordEditor,
            object bookmarks,
            int selectionStart,
            int selectionEnd)
        {
            string bookmarkName = "NcCursor" + Guid.NewGuid().ToString("N").Substring(0, 12);
            return TryAddEmailSignatureBookmark(
                wordEditor,
                bookmarks,
                bookmarkName,
                selectionStart,
                Math.Max(selectionStart, selectionEnd))
                ? bookmarkName
                : string.Empty;
        }

        private static void RestoreEmailSignatureSelection(
            OutlookWordEditorContext context,
            object bookmarks,
            string bookmarkName,
            int fallbackStart,
            int fallbackEnd)
        {
            if (context == null || context.Selection == null)
            {
                return;
            }

            int start = fallbackStart;
            int end = Math.Max(fallbackStart, fallbackEnd);
            bool bookmarkFound = !string.IsNullOrWhiteSpace(bookmarkName)
                                 && TryGetBookmarkRange(bookmarks, bookmarkName, out start, out end);
            if (!bookmarkFound)
            {
                int documentStart = context.GetDocumentStart(0);
                int documentEnd = context.GetDocumentEnd(documentStart);
                start = ClampWordPosition(fallbackStart, documentEnd);
                end = ClampWordPosition(Math.Max(start, fallbackEnd), documentEnd);
            }
            try
            {
                context.SetSelectionRange(start, end);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore Word selection after email signature reconciliation.", ex);
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(bookmarkName))
                {
                    TryDeleteEmailSignatureBookmark(bookmarks, bookmarkName);
                }
            }
        }

        private static void TryShowHiddenBookmarks(object bookmarks)
        {
            if (bookmarks == null)
            {
                return;
            }

            try
            {
                bookmarks.GetType().InvokeMember(
                    "ShowHidden",
                    BindingFlags.SetProperty,
                    null,
                    bookmarks,
                    new object[] { true });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to show hidden Outlook Word bookmarks.", ex);
            }
        }

        private static bool TryGetBookmarkRange(object bookmarks, string bookmarkName, out int start, out int end)
        {
            start = 0;
            end = 0;
            if (bookmarks == null || string.IsNullOrEmpty(bookmarkName))
            {
                return false;
            }

            object bookmark = null;
            object range = null;
            try
            {
                object exists = bookmarks.GetType().InvokeMember(
                    "Exists",
                    BindingFlags.InvokeMethod,
                    null,
                    bookmarks,
                    new object[] { bookmarkName });
                if (!(exists is bool) || !(bool)exists)
                {
                    return false;
                }

                bookmark = bookmarks.GetType().InvokeMember(
                    "Item",
                    BindingFlags.InvokeMethod,
                    null,
                    bookmarks,
                    new object[] { bookmarkName });
                if (bookmark == null)
                {
                    return false;
                }

                range = bookmark.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, bookmark, null);
                if (range == null)
                {
                    return false;
                }

                start = Convert.ToInt32(
                    range.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, range, null),
                    CultureInfo.InvariantCulture);
                end = Convert.ToInt32(
                    range.GetType().InvokeMember("End", BindingFlags.GetProperty, null, range, null),
                    CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook Word bookmark range (" + bookmarkName + ").", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release inline signature bookmark range COM object.");
                ComInteropScope.TryRelease(bookmark, LogCategories.Core, "Failed to release inline signature bookmark COM object.");
            }
        }

        private static bool TryExpandRangeToContainingTable(object wordEditor, int start, int end, out int tableStart, out int tableEnd)
        {
            tableStart = 0;
            tableEnd = 0;
            if (wordEditor == null)
            {
                return false;
            }

            object range = null;
            object tables = null;
            object table = null;
            object tableRange = null;
            try
            {
                range = wordEditor.GetType().InvokeMember(
                    "Range",
                    BindingFlags.InvokeMethod,
                    null,
                    wordEditor,
                    new object[] { start, Math.Max(start, end) });
                if (range == null)
                {
                    return false;
                }

                tables = range.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, range, null);
                if (tables == null)
                {
                    return false;
                }

                int count = Convert.ToInt32(
                    tables.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, tables, null),
                    CultureInfo.InvariantCulture);
                if (count <= 0)
                {
                    return false;
                }

                table = tables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, tables, new object[] { 1 });
                if (table == null)
                {
                    return false;
                }

                tableRange = table.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, table, null);
                if (tableRange == null)
                {
                    return false;
                }

                tableStart = Convert.ToInt32(
                    tableRange.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, tableRange, null),
                    CultureInfo.InvariantCulture);
                tableEnd = Convert.ToInt32(
                    tableRange.GetType().InvokeMember("End", BindingFlags.GetProperty, null, tableRange, null),
                    CultureInfo.InvariantCulture);
                return tableEnd > tableStart && tableStart <= start && tableEnd >= end;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to expand the Outlook signature range to its containing table.",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(tableRange, LogCategories.Core, "Failed to release inline signature table range COM object.");
                ComInteropScope.TryRelease(table, LogCategories.Core, "Failed to release inline signature table COM object.");
                ComInteropScope.TryRelease(tables, LogCategories.Core, "Failed to release inline signature tables COM object.");
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release inline signature range COM object.");
            }
        }

        private static bool TryDeleteContainingTableAtRange(object wordEditor, int start, int end)
        {
            if (wordEditor == null)
            {
                return false;
            }

            object table = null;
            try
            {
                table = TryGetContainingTableAtRange(wordEditor, start, end);
                if (table == null)
                {
                    return false;
                }

                table.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, table, null);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to delete the Outlook signature table.",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(table, LogCategories.Core, "Failed to release inline signature table COM object.");
            }
        }

        private static object TryGetContainingTableAtRange(object wordEditor, int start, int end)
        {
            if (wordEditor == null)
            {
                return null;
            }

            object range = null;
            object tables = null;
            object cells = null;
            object cell = null;
            object cellRange = null;
            object cellTables = null;
            try
            {
                range = wordEditor.GetType().InvokeMember(
                    "Range",
                    BindingFlags.InvokeMethod,
                    null,
                    wordEditor,
                    new object[] { start, Math.Max(start + 1, end) });
                if (range == null)
                {
                    return null;
                }

                tables = range.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, range, null);
                if (tables != null
                    && Convert.ToInt32(tables.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, tables, null), CultureInfo.InvariantCulture) > 0)
                {
                    return tables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, tables, new object[] { 1 });
                }

                cells = range.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, range, null);
                if (cells == null
                    || Convert.ToInt32(cells.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, cells, null), CultureInfo.InvariantCulture) <= 0)
                {
                    return null;
                }

                cell = cells.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, cells, new object[] { 1 });
                if (cell == null)
                {
                    return null;
                }

                cellRange = cell.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cell, null);
                if (cellRange == null)
                {
                    return null;
                }

                cellTables = cellRange.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, cellRange, null);
                if (cellTables == null
                    || Convert.ToInt32(cellTables.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, cellTables, null), CultureInfo.InvariantCulture) <= 0)
                {
                    return null;
                }

                return cellTables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, cellTables, new object[] { 1 });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to resolve the table containing the Outlook signature range.",
                    ex);
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(cellTables, LogCategories.Core, "Failed to release inline signature cell tables COM object.");
                ComInteropScope.TryRelease(cellRange, LogCategories.Core, "Failed to release inline signature cell range COM object.");
                ComInteropScope.TryRelease(cell, LogCategories.Core, "Failed to release inline signature cell COM object.");
                ComInteropScope.TryRelease(cells, LogCategories.Core, "Failed to release inline signature cells COM object.");
                ComInteropScope.TryRelease(tables, LogCategories.Core, "Failed to release inline signature tables COM object.");
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release inline signature range COM object.");
            }
        }

        private static bool TryFindInlineQuoteSeparatorStart(
            object wordEditor,
            int searchStart,
            bool hasExcludedRange,
            int excludedStart,
            int excludedEnd,
            out int separatorStart)
        {
            separatorStart = 0;
            if (wordEditor == null)
            {
                return false;
            }

            object content = null;
            object tailRange = null;
            object paragraphs = null;
            try
            {
                content = wordEditor.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, wordEditor, null);
                if (content == null)
                {
                    return false;
                }

                int documentEnd = Convert.ToInt32(
                    content.GetType().InvokeMember("End", BindingFlags.GetProperty, null, content, null),
                    CultureInfo.InvariantCulture);
                if (documentEnd <= searchStart)
                {
                    return false;
                }

                tailRange = wordEditor.GetType().InvokeMember(
                    "Range",
                    BindingFlags.InvokeMethod,
                    null,
                    wordEditor,
                    new object[] { searchStart, documentEnd });
                paragraphs = tailRange.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, tailRange, null);
                if (paragraphs == null)
                {
                    return false;
                }

                int count = Convert.ToInt32(
                    paragraphs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, paragraphs, null),
                    CultureInfo.InvariantCulture);
                int maxParagraphsToInspect = Math.Min(count, 80);
                for (int i = 1; i <= maxParagraphsToInspect; i++)
                {
                    object paragraph = null;
                    object paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paragraphs, new object[] { i });
                        if (paragraph == null)
                        {
                            continue;
                        }

                        paragraphRange = paragraph.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, paragraph, null);
                        if (paragraphRange == null)
                        {
                            continue;
                        }

                        int paragraphStart = Convert.ToInt32(
                            paragraphRange.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, paragraphRange, null),
                            CultureInfo.InvariantCulture);
                        int paragraphEnd = Convert.ToInt32(
                            paragraphRange.GetType().InvokeMember("End", BindingFlags.GetProperty, null, paragraphRange, null),
                            CultureInfo.InvariantCulture);
                        if (paragraphEnd <= searchStart)
                        {
                            continue;
                        }

                        if (hasExcludedRange
                            && paragraphStart < excludedEnd
                            && paragraphEnd > excludedStart)
                        {
                            continue;
                        }

                        if (ParagraphHasVisibleBorder(paragraph))
                        {
                            separatorStart = Math.Max(searchStart, paragraphStart);
                            return true;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(paragraphRange, LogCategories.Core, "Failed to release inline quote separator paragraph range COM object.");
                        ComInteropScope.TryRelease(paragraph, LogCategories.Core, "Failed to release inline quote separator paragraph COM object.");
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to find inline reply quote separator.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(paragraphs, LogCategories.Core, "Failed to release inline quote separator paragraphs COM object.");
                ComInteropScope.TryRelease(tailRange, LogCategories.Core, "Failed to release inline quote separator tail range COM object.");
                ComInteropScope.TryRelease(content, LogCategories.Core, "Failed to release inline quote separator content COM object.");
            }
        }

        private static bool ParagraphHasVisibleBorder(object paragraph)
        {
            if (paragraph == null)
            {
                return false;
            }

            object borders = null;
            try
            {
                borders = paragraph.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, paragraph, null);
                if (borders == null)
                {
                    return false;
                }

                return BorderAtIndexIsVisible(borders, -1)
                       || BorderAtIndexIsVisible(borders, -3)
                       || BorderAtIndexIsVisible(borders, -5);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to inspect inline reply paragraph borders.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(borders, LogCategories.Core, "Failed to release inline reply paragraph borders COM object.");
            }
        }

        private static bool BorderAtIndexIsVisible(object borders, int index)
        {
            object border = null;
            try
            {
                border = borders.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, borders, new object[] { index });
                if (border == null)
                {
                    return false;
                }

                object lineStyle = border.GetType().InvokeMember("LineStyle", BindingFlags.GetProperty, null, border, null);
                int value;
                return lineStyle != null
                       && int.TryParse(Convert.ToString(lineStyle, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                       && value != 0;
            }
            catch (Exception ex)
            {
                if (DiagnosticsLogger.IsEnabled)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline reply border lookup was unavailable (index="
                        + index.ToString(CultureInfo.InvariantCulture)
                        + ", errorType="
                        + ex.GetType().Name
                        + ").");
                }
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(border, LogCategories.Core, "Failed to release inline reply border COM object.");
            }
        }

        private static string EnsureHtmlDocumentForWordInsert(string html)
        {
            string value = html ?? string.Empty;
            if (value.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return value;
            }

            return "<html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>"
                   + value
                   + "</body></html>";
        }

        internal static bool TryWriteAppointmentHtmlBody(Outlook.AppointmentItem appointment, string html)
        {
            if (appointment == null || string.IsNullOrWhiteSpace(html))
            {
                return false;
            }

            Outlook.Application application = null;
            Outlook.MailItem stagingMail = null;

            try
            {
                application = appointment.Application;
                if (application == null)
                {
                    return false;
                }

                stagingMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (stagingMail == null)
                {
                    return false;
                }
                string bridgeHtml = HtmlTemplateSanitizer.PrepareTalkAppointmentHtmlForOutlookRtfBridge(html);
                if (string.IsNullOrWhiteSpace(bridgeHtml))
                {
                    bridgeHtml = html ?? string.Empty;
                }

                stagingMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                stagingMail.HTMLBody = bridgeHtml;

                var rtfBody = stagingMail.RTFBody as byte[];
                if (rtfBody == null || rtfBody.Length == 0)
                {
                    DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML->RTF bridge produced empty RTF body.");
                    return false;
                }

                appointment.RTFBody = rtfBody;
                DiagnosticsLogger.Log(LogCategories.Talk, "Appointment HTML body written via HTML->RTF bridge.");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to write appointment HTML body via HTML->RTF bridge.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(stagingMail, LogCategories.Talk, "Failed to release staging MailItem COM object.");
                ComInteropScope.TryRelease(application, LogCategories.Talk, "Failed to release Outlook application COM object for appointment HTML->RTF bridge.");
            }
        }
    }
}

