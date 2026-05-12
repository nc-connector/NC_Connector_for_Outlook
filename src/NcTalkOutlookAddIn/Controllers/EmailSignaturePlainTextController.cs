// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal static class EmailSignaturePlainTextController
    {
        private const string OutlookAutoSignatureBookmarkName = "_MailAutoSig";
        private const string ManagedSignatureBookmarkName = "NcConnectorSignature";

        internal sealed class Result
        {
            internal bool Success { get; set; }

            internal bool Changed { get; set; }

            internal bool Managed { get; set; }

            internal bool InitialSlotHandled { get; set; }

            internal string Source { get; set; }
        }

        internal static Result Apply(
            Outlook.Application application,
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            string plainText,
            Action<string> log)
        {
            string replacement = BuildSignatureBlock(plainText, false);
            if (string.IsNullOrWhiteSpace(replacement))
            {
                Log(log, "Plain-text signature skipped because converted text is empty.");
                return new Result { Success = false, Source = "empty" };
            }

            OutlookWordEditorContext context;
            if (!TryOpenWordEditorContext(application, mail, isInlineResponse, composeKey, log, out context))
            {
                return new Result { Success = false, Source = "word_editor_unavailable" };
            }

            using (context)
            {
                string source;
                if (TryReplaceBookmarkText(context.Document, ManagedSignatureBookmarkName, replacement, true, out source))
                {
                    RestoreSelection(context.Selection, context.SelectionStart, log);
                    Log(log, "Plain-text signature updated via managed bookmark (source=" + source + ").");
                    return new Result { Success = true, Changed = true, Managed = true, InitialSlotHandled = true, Source = source };
                }

                if (TryReplaceBookmarkText(context.Document, OutlookAutoSignatureBookmarkName, replacement, true, out source))
                {
                    RestoreSelection(context.Selection, context.SelectionStart, log);
                    Log(log, "Plain-text signature replaced Outlook auto-signature bookmark (source=" + source + ").");
                    return new Result { Success = true, Changed = true, Managed = true, InitialSlotHandled = true, Source = source };
                }

                string insertedSource;
                if (TryInsertAtSelection(context, BuildSignatureBlock(plainText, true), log, out insertedSource))
                {
                    Log(log, "Plain-text signature inserted at Word selection (source=" + insertedSource + ").");
                    return new Result { Success = true, Changed = true, Managed = true, InitialSlotHandled = true, Source = insertedSource };
                }

                Log(log, "Plain-text signature skipped because no safe Word insertion point was available.");
                return new Result { Success = false, Source = "no_safe_slot" };
            }
        }

        internal static Result ClearInitialSlot(
            Outlook.Application application,
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            Action<string> log)
        {
            OutlookWordEditorContext context;
            if (!TryOpenWordEditorContext(application, mail, isInlineResponse, composeKey, log, out context))
            {
                return new Result { Success = false, Source = "word_editor_unavailable" };
            }

            using (context)
            {
                string source;
                if (TryClearBookmarkText(context.Document, ManagedSignatureBookmarkName, out source))
                {
                    Log(log, "Managed plain-text signature cleared (source=" + source + ").");
                    return new Result { Success = true, Changed = true, Managed = false, InitialSlotHandled = true, Source = source };
                }

                if (TryClearBookmarkText(context.Document, OutlookAutoSignatureBookmarkName, out source))
                {
                    Log(log, "Plain-text Outlook auto-signature slot cleared (source=" + source + ").");
                    return new Result { Success = true, Changed = true, Managed = false, InitialSlotHandled = true, Source = source };
                }

                Log(log, "Plain-text signature slot clear skipped because no managed or Outlook auto-signature bookmark exists.");
                return new Result { Success = true, Changed = false, Managed = false, InitialSlotHandled = false, Source = "not_found" };
            }
        }

        internal static Result ClearManaged(
            Outlook.Application application,
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            Action<string> log)
        {
            OutlookWordEditorContext context;
            if (!TryOpenWordEditorContext(application, mail, isInlineResponse, composeKey, log, out context))
            {
                return new Result { Success = false, Source = "word_editor_unavailable" };
            }

            using (context)
            {
                string source;
                if (TryClearBookmarkText(context.Document, ManagedSignatureBookmarkName, out source))
                {
                    Log(log, "Managed plain-text signature cleared (source=" + source + ").");
                    return new Result { Success = true, Changed = true, Managed = false, InitialSlotHandled = true, Source = source };
                }

                Log(log, "Managed plain-text signature clear skipped because the managed bookmark does not exist.");
                return new Result { Success = true, Changed = false, Managed = false, InitialSlotHandled = false, Source = "not_found" };
            }
        }

        private static bool TryOpenWordEditorContext(
            Outlook.Application application,
            Outlook.MailItem mail,
            bool isInlineResponse,
            string composeKey,
            Action<string> log,
            out OutlookWordEditorContext context)
        {
            context = null;
            if (mail == null)
            {
                return false;
            }

            if (isInlineResponse)
            {
                if (!OutlookWordEditorContext.TryOpenInline(application, mail, "plain-text signature", composeKey, out context))
                {
                    Log(log, "Plain-text inline signature skipped because the inline Word editor is unavailable.");
                    return false;
                }

                return true;
            }

            if (!OutlookWordEditorContext.TryOpenInspector(mail, "plain-text signature", out context))
            {
                Log(log, "Plain-text signature skipped because the inspector Word editor is unavailable.");
                return false;
            }

            return true;
        }

        private static bool TryReplaceBookmarkText(object document, string bookmarkName, string text, bool addManagedBookmark, out string source)
        {
            source = string.Empty;
            object bookmarks = null;
            object bookmark = null;
            object range = null;
            try
            {
                bookmarks = GetBookmarks(document);
                if (!BookmarkExists(bookmarks, bookmarkName))
                {
                    source = bookmarkName + ":not_found";
                    return false;
                }

                bookmark = GetBookmark(bookmarks, bookmarkName);
                range = OutlookWordEditorContext.GetProperty(bookmark, "Range");
                if (range == null)
                {
                    source = bookmarkName + ":range_unavailable";
                    return false;
                }

                int start = OutlookWordEditorContext.GetIntProperty(range, "Start");
                OutlookWordEditorContext.SetProperty(range, "Text", text ?? string.Empty);
                int end = OutlookWordEditorContext.GetIntProperty(range, "End");
                if (end <= start)
                {
                    end = start + (text ?? string.Empty).Length;
                }

                if (addManagedBookmark && !string.IsNullOrEmpty(text))
                {
                    TryAddManagedBookmark(document, start, end);
                }

                source = bookmarkName;
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to replace plain-text signature bookmark " + bookmarkName + ".", error);
                source = bookmarkName + ":error";
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release Word bookmark range COM object.");
                ComInteropScope.TryRelease(bookmark, LogCategories.Core, "Failed to release Word bookmark COM object.");
                ComInteropScope.TryRelease(bookmarks, LogCategories.Core, "Failed to release Word bookmarks COM object.");
            }
        }

        private static bool TryClearBookmarkText(object document, string bookmarkName, out string source)
        {
            source = string.Empty;
            return TryReplaceBookmarkText(document, bookmarkName, string.Empty, false, out source);
        }

        private static bool TryInsertAtSelection(OutlookWordEditorContext context, string text, Action<string> log, out string source)
        {
            source = string.Empty;
            object range = null;
            try
            {
                int start = context.Selection != null ? OutlookWordEditorContext.GetIntProperty(context.Selection, "Start") : 0;
                range = CreateRange(context.Document, start, start);
                if (range == null)
                {
                    source = "range_unavailable";
                    return false;
                }

                OutlookWordEditorContext.SetProperty(range, "Text", text ?? string.Empty);
                int end = OutlookWordEditorContext.GetIntProperty(range, "End");
                if (end <= start)
                {
                    end = start + (text ?? string.Empty).Length;
                }

                TryAddManagedBookmark(context.Document, start, end);
                RestoreSelection(context.Selection, start, log);
                source = context.Source ?? "word";
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain-text signature at Word selection.", error);
                source = "error";
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release Word insertion range COM object.");
            }
        }

        private static void TryAddManagedBookmark(object document, int start, int end)
        {
            object bookmarks = null;
            object existingBookmark = null;
            object range = null;
            try
            {
                if (document == null || end <= start)
                {
                    return;
                }

                bookmarks = GetBookmarks(document);
                if (bookmarks == null)
                {
                    return;
                }

                if (BookmarkExists(bookmarks, ManagedSignatureBookmarkName))
                {
                    existingBookmark = GetBookmark(bookmarks, ManagedSignatureBookmarkName);
                    OutlookWordEditorContext.InvokeMethod(existingBookmark, "Delete", null);
                }

                range = CreateRange(document, start, end);
                if (range == null)
                {
                    return;
                }

                OutlookWordEditorContext.InvokeMethod(bookmarks, "Add", new object[] { ManagedSignatureBookmarkName, range });
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to add managed plain-text signature bookmark.", error);
            }
            finally
            {
                ComInteropScope.TryRelease(range, LogCategories.Core, "Failed to release managed signature range COM object.");
                ComInteropScope.TryRelease(existingBookmark, LogCategories.Core, "Failed to release existing managed signature bookmark COM object.");
                ComInteropScope.TryRelease(bookmarks, LogCategories.Core, "Failed to release Word bookmarks COM object.");
            }
        }

        private static object GetBookmarks(object document)
        {
            return OutlookWordEditorContext.GetProperty(document, "Bookmarks");
        }

        private static bool BookmarkExists(object bookmarks, string bookmarkName)
        {
            if (bookmarks == null || string.IsNullOrWhiteSpace(bookmarkName))
            {
                return false;
            }

            object result = OutlookWordEditorContext.InvokeMethod(bookmarks, "Exists", new object[] { bookmarkName });
            return result != null && Convert.ToBoolean(result, CultureInfo.InvariantCulture);
        }

        private static object GetBookmark(object bookmarks, string bookmarkName)
        {
            return OutlookWordEditorContext.InvokeMethod(bookmarks, "Item", new object[] { bookmarkName });
        }

        private static object CreateRange(object document, int start, int end)
        {
            return OutlookWordEditorContext.InvokeMethod(document, "Range", new object[] { start, end });
        }

        private static void RestoreSelection(object selection, int start, Action<string> log)
        {
            if (selection == null)
            {
                return;
            }

            try
            {
                OutlookWordEditorContext.InvokeMethod(selection, "SetRange", new object[] { start, start });
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore cursor before plain-text signature.", error);
            }
        }

        private static string BuildSignatureBlock(string plainText, bool includeLeadingUserGap)
        {
            string normalized = PlainTextUtilities.NormalizeCrLf(plainText).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            string block = "-- \r\n" + normalized + "\r\n";
            return includeLeadingUserGap ? "\r\n\r\n" + block : block;
        }

        private static void Log(Action<string> log, string message)
        {
            if (log != null)
            {
                log(message ?? string.Empty);
            }
        }
    }
}
