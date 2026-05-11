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
        private readonly NextcloudTalkAddIn _owner;

        internal MailInteropController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal static string ResolveMailInspectorIdentityKey(Outlook.MailItem mail)
        {            if (mail == null)
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
        {            if (mail == null)
            {
                return null;
            }

            Outlook.Inspector inspector = null;
            try
            {
                inspector = mail.GetInspector;                if (inspector == null)
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
        {            if (inspector == null)
            {
                return 0;
            }
            foreach (string propertyName in new[] { "HWND", "Hwnd" })
            {
                try
                {
                    PropertyInfo property = inspector.GetType().GetProperty(propertyName);                    if (property == null)
                    {
                        continue;
                    }

                    object value = property.GetValue(inspector, null);                    if (value == null)
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
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;            if (application == null)
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
                var mailItem = inlineResponse as Outlook.MailItem;                if (mailItem != null)
                {
                    return mailItem;
                }
            }
            return null;
        }

        internal string ResolveActiveInspectorIdentityKey()
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;            if (application == null)
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to check active inline response.", error);
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
            if (TryInsertHtmlIntoMailBody(mail, html))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (HTMLBody primary).");
                return;
            }

            IDataObject previousClipboard = null;
            bool restoreClipboard = false;

            try
            {
                previousClipboard = Clipboard.GetDataObject();
                restoreClipboard = previousClipboard != null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read clipboard data object.", ex);
                previousClipboard = null;
                restoreClipboard = false;
            }
            try
            {
                Clipboard.SetText(html, TextDataFormat.Html);
                if (TryPasteClipboardIntoMailInspector(mail))
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "Inserted HTML block into mail (WordEditor).");
                    return;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Failed to insert HTML via WordEditor: " + ex.Message);
            }
            finally
            {
                if (restoreClipboard)
                {
                    try
                    {
                        Clipboard.SetDataObject(previousClipboard);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore clipboard after HTML insertion.", ex);
                    }
                }
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
                string insertText = NormalizeMailPlainText(plainText);
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain-text share block into mail.", error);
                MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture, Strings.ErrorInsertHtmlFailed, error.Message),
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
            Outlook.Explorer explorer = null;
            Outlook.MailItem activeInlineMail = null;
            Outlook.Inspector inspector = null;
            object wordEditor = null;
            object wordApplication = null;
            object selection = null;
            bool activeInlineMatches = false;

            try
            {
                if (application != null)
                {
                    explorer = application.ActiveExplorer();
                    if (explorer != null)
                    {
                        activeInlineMail = explorer.ActiveInlineResponse as Outlook.MailItem;
                        activeInlineMatches = ComInteropScope.AreSameObject(mail, activeInlineMail, LogCategories.Core, "MailItem", "ActiveInlineResponse");
                        if (activeInlineMatches)
                        {
                            wordEditor = explorer.GetType().InvokeMember(
                                "ActiveInlineResponseWordEditor",
                                BindingFlags.GetProperty,
                                null,
                                explorer,
                                null);
                            source = "inline";
                        }
                    }
                }

                if (activeInlineMatches && wordEditor == null)
                {
                    return false;
                }

                if (wordEditor == null)
                {
                    inspector = mail.GetInspector;
                    if (inspector == null)
                    {
                        return false;
                    }

                    wordEditor = inspector.WordEditor;
                    source = "inspector";
                }

                if (wordEditor == null)
                {
                    return false;
                }

                wordApplication = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                if (wordApplication == null)
                {
                    return false;
                }

                selection = wordApplication.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApplication, null);
                if (selection == null)
                {
                    return false;
                }

                if (activeInlineMatches)
                {
                    int replyCursorStart = GetWordDocumentStart(wordEditor, 0);
                    selection.GetType().InvokeMember(
                        "SetRange",
                        BindingFlags.InvokeMethod,
                        null,
                        selection,
                        new object[] { replyCursorStart, replyCursorStart });
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                    selection.GetType().InvokeMember("TypeText", BindingFlags.InvokeMethod, null, selection, new object[] { text });
                    selection.GetType().InvokeMember(
                        "SetRange",
                        BindingFlags.InvokeMethod,
                        null,
                        selection,
                        new object[] { replyCursorStart, replyCursorStart });
                }
                else
                {
                    selection.GetType().InvokeMember("TypeText", BindingFlags.InvokeMethod, null, selection, new object[] { text });
                }
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert plain text via Outlook WordEditor.", error);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release Word selection COM object after plain-text insert.");
                ComInteropScope.TryRelease(wordApplication, LogCategories.Core, "Failed to release Word application COM object after plain-text insert.");
                ComInteropScope.TryRelease(wordEditor, LogCategories.Core, "Failed to release Word editor COM object after plain-text insert.");
                ComInteropScope.TryRelease(inspector, LogCategories.Core, "Failed to release Outlook Inspector COM object after plain-text insert.");
                if (!ReferenceEquals(activeInlineMail, mail))
                {
                    ComInteropScope.TryRelease(activeInlineMail, LogCategories.Core, "Failed to release ActiveInlineResponse MailItem COM object after plain-text insert.");
                }
                ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release Outlook Explorer COM object after plain-text insert.");
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

        private static string NormalizeMailPlainText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n").Trim();
        }

        private bool TryInsertHtmlIntoActiveInlineResponseWordEditor(Outlook.MailItem mail, string html)
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            Outlook.Explorer explorer = null;
            Outlook.MailItem activeInlineMail = null;
            object wordEditor = null;
            object wordApplication = null;
            object selection = null;
            string tempHtmlPath = null;

            try
            {
                explorer = application != null ? application.ActiveExplorer() : null;
                activeInlineMail = explorer != null ? explorer.ActiveInlineResponse as Outlook.MailItem : null;
                if (!ComInteropScope.AreSameObject(mail, activeInlineMail, LogCategories.Core, "MailItem", "ActiveInlineResponse"))
                {
                    return false;
                }

                wordEditor = explorer.GetType().InvokeMember(
                    "ActiveInlineResponseWordEditor",
                    BindingFlags.GetProperty,
                    null,
                    explorer,
                    null);
                if (wordEditor == null)
                {
                    return false;
                }

                wordApplication = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                if (wordApplication == null)
                {
                    return false;
                }

                selection = wordApplication.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApplication, null);
                if (selection == null)
                {
                    return false;
                }

                int replyCursorStart = GetWordDocumentStart(wordEditor, 0);
                selection.GetType().InvokeMember(
                    "SetRange",
                    BindingFlags.InvokeMethod,
                    null,
                    selection,
                    new object[] { replyCursorStart, replyCursorStart });
                selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);

                tempHtmlPath = Path.Combine(
                    Path.GetTempPath(),
                    "nc4ol-inline-share-" + Guid.NewGuid().ToString("N") + ".html");
                File.WriteAllText(tempHtmlPath, EnsureHtmlDocumentForWordInsert(html), new UTF8Encoding(true));
                selection.GetType().InvokeMember(
                    "InsertFile",
                    BindingFlags.InvokeMethod,
                    null,
                    selection,
                    new object[] { tempHtmlPath, Type.Missing, false, false, false });
                selection.GetType().InvokeMember(
                    "SetRange",
                    BindingFlags.InvokeMethod,
                    null,
                    selection,
                    new object[] { replyCursorStart, replyCursorStart });
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to insert HTML into inline response Word editor.", error);
                return false;
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempHtmlPath))
                {
                    try
                    {
                        File.Delete(tempHtmlPath);
                    }
                    catch (Exception error)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to delete temporary inline share HTML file.", error);
                    }
                }

                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release inline share Word selection COM object.");
                ComInteropScope.TryRelease(wordApplication, LogCategories.Core, "Failed to release inline share Word application COM object.");
                ComInteropScope.TryRelease(wordEditor, LogCategories.Core, "Failed to release inline share Word editor COM object.");
                if (!ReferenceEquals(activeInlineMail, mail))
                {
                    ComInteropScope.TryRelease(activeInlineMail, LogCategories.Core, "Failed to release ActiveInlineResponse MailItem COM object after inline share insert.");
                }
                ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release active Explorer COM object after inline share insert.");
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

        private static bool TryPasteClipboardIntoMailInspector(Outlook.MailItem mail)
        {
            Outlook.Inspector inspector = null;
            object wordEditor = null;
            object application = null;
            object selection = null;

            try
            {
                inspector = mail.GetInspector;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to access MailItem.GetInspector for HTML paste.", ex);
                inspector = null;
            }
            if (inspector == null)
            {
                return false;
            }
            try
            {
                wordEditor = inspector.WordEditor;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to access Inspector.WordEditor for HTML paste.", ex);
                wordEditor = null;
            }
            if (wordEditor == null)
            {
                return false;
            }
            try
            {
                application = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);                if (application == null)
                {
                    return false;
                }

                selection = application.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, application, null);                if (selection == null)
                {
                    return false;
                }

                // Insert near the top so we are reliably above the signature block.
                try
                {
                    // wdStory = 6
                    selection.GetType().InvokeMember("HomeKey", BindingFlags.InvokeMethod, null, selection, new object[] { 6, 0 });
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to move cursor in Word editor before pasting HTML (best-effort).", ex);
                }

                selection.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, selection, null);
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to paste HTML into Word editor.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release Word selection COM object.");
                ComInteropScope.TryRelease(application, LogCategories.Core, "Failed to release Word application COM object.");
            }
        }

        internal bool TryWriteMailHtmlBodyPreservingSelection(Outlook.MailItem mail, string html, string composeKey, string operation, bool placeCursorAtBodyStart = false)
        {
            if (mail == null)
            {
                return false;
            }

            Outlook.Inspector inspector = null;
            object wordEditor = null;
            object wordApplication = null;
            object selection = null;
            WordSelectionSnapshot snapshot = null;

            try
            {
                inspector = mail.GetInspector;
                if (inspector != null)
                {
                    wordEditor = inspector.WordEditor;
                    if (wordEditor != null)
                    {
                        wordApplication = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                        if (wordApplication != null)
                        {
                            selection = wordApplication.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApplication, null);
                            snapshot = CaptureWordSelection(selection);
                        }
                    }
                }

                mail.HTMLBody = html ?? string.Empty;

                if (wordApplication != null && (snapshot != null || placeCursorAtBodyStart))
                {
                    ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release pre-write Word selection COM object.");
                    selection = wordApplication.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApplication, null);
                    if (placeCursorAtBodyStart)
                    {
                        if (!TryMoveWordSelectionToDocumentStart(selection, wordEditor, snapshot) && snapshot != null)
                        {
                            RestoreWordSelection(selection, wordEditor, snapshot);
                        }
                    }
                    else if (snapshot != null)
                    {
                        RestoreWordSelection(selection, wordEditor, snapshot);
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Mail HTML body written with Word selection restore (operation="
                    + (operation ?? "n/a")
                    + ", composeKey="
                    + (composeKey ?? string.Empty)
                    + ", selectionCaptured="
                    + (snapshot != null).ToString(CultureInfo.InvariantCulture)
                    + ", cursorAtBodyStart="
                    + placeCursorAtBodyStart.ToString(CultureInfo.InvariantCulture)
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to write mail HTML body with Word selection restore (operation="
                    + (operation ?? "n/a")
                    + ", composeKey="
                    + (composeKey ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release Word selection COM object after HTML body write.");
                ComInteropScope.TryRelease(wordApplication, LogCategories.Core, "Failed to release Word application COM object after HTML body write.");
                ComInteropScope.TryRelease(wordEditor, LogCategories.Core, "Failed to release Word editor COM object after HTML body write.");
                ComInteropScope.TryRelease(inspector, LogCategories.Core, "Failed to release Inspector COM object after HTML body write.");
            }
        }

        private static WordSelectionSnapshot CaptureWordSelection(object selection)
        {
            if (selection == null)
            {
                return null;
            }

            object font = null;
            try
            {
                var snapshot = new WordSelectionSnapshot
                {
                    Start = Convert.ToInt32(selection.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, selection, null), CultureInfo.InvariantCulture),
                    End = Convert.ToInt32(selection.GetType().InvokeMember("End", BindingFlags.GetProperty, null, selection, null), CultureInfo.InvariantCulture)
                };

                font = selection.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, selection, null);
                if (font != null)
                {
                    snapshot.FontName = ReadWordFontProperty(font, "Name");
                    snapshot.FontNameAscii = ReadWordFontProperty(font, "NameAscii");
                    snapshot.FontNameOther = ReadWordFontProperty(font, "NameOther");
                    snapshot.FontNameFarEast = ReadWordFontProperty(font, "NameFarEast");
                    snapshot.FontSize = ReadWordFontProperty(font, "Size");
                }

                return snapshot;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to capture Word selection formatting.", ex);
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(font, LogCategories.Core, "Failed to release captured Word font COM object.");
            }
        }

        private static string ReadWordFontProperty(object font, string propertyName)
        {
            try
            {
                object value = font.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, font, null);
                return value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : null;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Word font property (" + propertyName + ").", ex);
                return null;
            }
        }

        private static void RestoreWordSelection(object selection, object wordEditor, WordSelectionSnapshot snapshot)
        {
            if (selection == null || snapshot == null)
            {
                return;
            }

            try
            {
                int documentEnd = GetWordDocumentEnd(wordEditor);
                int start = ClampWordPosition(snapshot.Start, documentEnd);
                int end = ClampWordPosition(snapshot.End, documentEnd);
                if (end < start)
                {
                    end = start;
                }

                selection.GetType().InvokeMember("SetRange", BindingFlags.InvokeMethod, null, selection, new object[] { start, end });
                RestoreWordSelectionFont(selection, snapshot);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore Word selection formatting.", ex);
            }
        }

        private static bool TryMoveWordSelectionToDocumentStart(object selection, object wordEditor, WordSelectionSnapshot snapshot)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                int start = GetWordDocumentStart(wordEditor, 0);
                selection.GetType().InvokeMember("SetRange", BindingFlags.InvokeMethod, null, selection, new object[] { start, start });
                RestoreWordSelectionFont(selection, snapshot);
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to move Word selection to body start.", error);
                return false;
            }
        }

        private static void RestoreWordSelectionFont(object selection, WordSelectionSnapshot snapshot)
        {
            if (selection == null || snapshot == null)
            {
                return;
            }

            object font = null;
            try
            {
                font = selection.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, selection, null);
                if (font != null)
                {
                    WriteWordFontProperty(font, "Name", snapshot.FontName);
                    WriteWordFontProperty(font, "NameAscii", snapshot.FontNameAscii);
                    WriteWordFontProperty(font, "NameOther", snapshot.FontNameOther);
                    WriteWordFontProperty(font, "NameFarEast", snapshot.FontNameFarEast);
                    WriteWordFontProperty(font, "Size", snapshot.FontSize);
                }
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore Word selection font.", error);
            }
            finally
            {
                ComInteropScope.TryRelease(font, LogCategories.Core, "Failed to release restored Word font COM object.");
            }
        }

        private static int GetWordDocumentStart(object wordEditor, int fallback)
        {
            object content = null;
            try
            {
                content = wordEditor != null ? wordEditor.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, wordEditor, null) : null;
                if (content == null)
                {
                    return fallback;
                }

                return Convert.ToInt32(content.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, content, null), CultureInfo.InvariantCulture);
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Word document start.", error);
                return fallback;
            }
            finally
            {
                ComInteropScope.TryRelease(content, LogCategories.Core, "Failed to release Word content COM object.");
            }
        }

        private static int GetWordDocumentEnd(object wordEditor)
        {
            object content = null;
            try
            {
                content = wordEditor != null ? wordEditor.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, wordEditor, null) : null;
                if (content == null)
                {
                    return int.MaxValue;
                }

                return Convert.ToInt32(content.GetType().InvokeMember("End", BindingFlags.GetProperty, null, content, null), CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Word document end.", ex);
                return int.MaxValue;
            }
            finally
            {
                ComInteropScope.TryRelease(content, LogCategories.Core, "Failed to release Word content COM object.");
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

        private static void WriteWordFontProperty(object font, string propertyName, string value)
        {
            if (font == null || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            try
            {
                object typedValue = value;
                float fontSize;
                if (string.Equals(propertyName, "Size", StringComparison.OrdinalIgnoreCase)
                    && float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out fontSize))
                {
                    typedValue = fontSize;
                }
                font.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, font, new[] { typedValue });
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to restore Word font property (" + propertyName + ").", ex);
            }
        }

        private sealed class WordSelectionSnapshot
        {
            internal int Start;
            internal int End;
            internal string FontName;
            internal string FontNameAscii;
            internal string FontNameOther;
            internal string FontNameFarEast;
            internal string FontSize;
        }

        internal bool TryReplaceActiveInlineResponseSignatureSlot(Outlook.MailItem mail, string html, string composeKey, string operation)
        {
            if (mail == null)
            {
                return false;
            }

            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            if (application == null)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Inline response HTML write skipped (operation="
                    + (operation ?? "n/a")
                    + ", composeKey="
                    + (composeKey ?? string.Empty)
                    + ", reason=application_unavailable).");
                return false;
            }

            Outlook.Explorer explorer = null;
            Outlook.MailItem activeInlineMail = null;
            object wordEditor = null;
            object wordApplication = null;
            object selection = null;
            string tempHtmlPath = null;

            try
            {
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response HTML write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=active_explorer_unavailable).");
                    return false;
                }

                activeInlineMail = explorer.ActiveInlineResponse as Outlook.MailItem;
                if (!ComInteropScope.AreSameObject(mail, activeInlineMail, LogCategories.Core, "MailItem", "ActiveInlineResponse"))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response HTML write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=active_inline_response_mismatch).");
                    return false;
                }

                wordEditor = explorer.GetType().InvokeMember(
                    "ActiveInlineResponseWordEditor",
                    BindingFlags.GetProperty,
                    null,
                    explorer,
                    null);
                if (wordEditor == null)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response HTML write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=word_editor_unavailable).");
                    return false;
                }

                wordApplication = wordEditor.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, wordEditor, null);
                if (wordApplication == null)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response signature slot write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=word_application_unavailable).");
                    return false;
                }

                selection = wordApplication.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApplication, null);
                if (selection == null)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response signature slot write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=word_selection_unavailable).");
                    return false;
                }

                int selectionStart = Convert.ToInt32(
                    selection.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, selection, null),
                    CultureInfo.InvariantCulture);
                int selectionEnd = Convert.ToInt32(
                    selection.GetType().InvokeMember("End", BindingFlags.GetProperty, null, selection, null),
                    CultureInfo.InvariantCulture);
                int originalSelectionStart = selectionStart;
                string signatureSlotSource;
                bool signatureSlotSelected = TrySelectInlineSignatureSlot(
                    wordEditor,
                    selection,
                    out selectionStart,
                    out selectionEnd,
                    out signatureSlotSource);
                if (!signatureSlotSelected)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "Inline response signature slot write skipped (operation="
                        + (operation ?? "n/a")
                        + ", composeKey="
                        + (composeKey ?? string.Empty)
                        + ", reason=signature_boundary_unavailable).");
                    return false;
                }

                int fallbackReplyCursorStart = originalSelectionStart < selectionStart ? originalSelectionStart : selectionStart;
                int replyCursorStart = GetWordDocumentStart(wordEditor, fallbackReplyCursorStart);
                if (!string.IsNullOrWhiteSpace(html))
                {
                    if (signatureSlotSelected && selectionEnd > selectionStart)
                    {
                        bool deletedTable = TryDeleteContainingTableAtRange(wordEditor, selectionStart, selectionEnd);
                        if (deletedTable)
                        {
                            signatureSlotSource = (signatureSlotSource ?? "mail_auto_sig") + "_table_deleted";
                        }
                        if (!deletedTable)
                        {
                            selection.GetType().InvokeMember(
                                "Delete",
                                BindingFlags.InvokeMethod,
                                null,
                                selection,
                                null);
                        }
                        selection.GetType().InvokeMember(
                            "SetRange",
                            BindingFlags.InvokeMethod,
                            null,
                            selection,
                            new object[] { selectionStart, selectionStart });
                    }
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);
                    selection.GetType().InvokeMember("TypeParagraph", BindingFlags.InvokeMethod, null, selection, null);

                    tempHtmlPath = Path.Combine(
                        Path.GetTempPath(),
                        "nc4ol-inline-signature-" + Guid.NewGuid().ToString("N") + ".html");
                    File.WriteAllText(tempHtmlPath, EnsureHtmlDocumentForWordInsert(html), new UTF8Encoding(true));
                    selection.GetType().InvokeMember(
                        "InsertFile",
                        BindingFlags.InvokeMethod,
                        null,
                        selection,
                        new object[] { tempHtmlPath, Type.Missing, false, false, false });
                    selection.GetType().InvokeMember(
                        "SetRange",
                        BindingFlags.InvokeMethod,
                        null,
                        selection,
                        new object[] { replyCursorStart, replyCursorStart });
                }
                else if (signatureSlotSelected && selectionEnd > selectionStart)
                {
                    bool deletedTable = TryDeleteContainingTableAtRange(wordEditor, selectionStart, selectionEnd);
                    if (deletedTable)
                    {
                        signatureSlotSource = (signatureSlotSource ?? "mail_auto_sig") + "_table_deleted";
                    }
                    if (!deletedTable)
                    {
                        selection.GetType().InvokeMember(
                            "Delete",
                            BindingFlags.InvokeMethod,
                            null,
                            selection,
                            null);
                    }
                    selection.GetType().InvokeMember(
                        "SetRange",
                        BindingFlags.InvokeMethod,
                        null,
                        selection,
                        new object[] { replyCursorStart, replyCursorStart });
                }
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Inline response signature inserted via ActiveInlineResponseWordEditor.Selection.InsertFile (operation="
                    + (operation ?? "n/a")
                    + ", composeKey="
                    + (composeKey ?? string.Empty)
                    + ", selectionStart="
                    + selectionStart.ToString(CultureInfo.InvariantCulture)
                    + ", selectionEnd="
                    + selectionEnd.ToString(CultureInfo.InvariantCulture)
                    + ", replyCursorStart="
                    + replyCursorStart.ToString(CultureInfo.InvariantCulture)
                    + ", signatureSlotSelected="
                    + signatureSlotSelected.ToString(CultureInfo.InvariantCulture)
                    + ", signatureSlotSource="
                    + (signatureSlotSource ?? "n/a")
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to write inline response signature slot (operation="
                    + (operation ?? "n/a")
                    + ", composeKey="
                    + (composeKey ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempHtmlPath))
                {
                    try
                    {
                        File.Delete(tempHtmlPath);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to delete temporary inline response HTML file.", ex);
                    }
                }

                ComInteropScope.TryRelease(selection, LogCategories.Core, "Failed to release inline Word selection COM object.");
                ComInteropScope.TryRelease(wordApplication, LogCategories.Core, "Failed to release inline Word application COM object.");
                ComInteropScope.TryRelease(wordEditor, LogCategories.Core, "Failed to release inline Word editor COM object.");
                if (!ReferenceEquals(activeInlineMail, mail))
                {
                    ComInteropScope.TryRelease(activeInlineMail, LogCategories.Core, "Failed to release ActiveInlineResponse MailItem COM object.");
                }
                ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release active Explorer COM object.");
            }
        }

        private static bool TrySelectInlineSignatureSlot(
            object wordEditor,
            object selection,
            out int selectionStart,
            out int selectionEnd,
            out string source)
        {
            selectionStart = 0;
            selectionEnd = 0;
            source = "none";
            if (wordEditor == null || selection == null)
            {
                return false;
            }

            object bookmarks = null;
            try
            {
                bookmarks = wordEditor.GetType().InvokeMember("Bookmarks", BindingFlags.GetProperty, null, wordEditor, null);
                if (bookmarks == null)
                {
                    return false;
                }

                TryShowHiddenBookmarks(bookmarks);

                int autoSignatureStart;
                int autoSignatureEnd;
                int originalStart;
                int originalEnd;
                bool hasAutoSignature = TryGetBookmarkRange(
                    bookmarks,
                    OutlookAutoSignatureBookmarkName,
                    out autoSignatureStart,
                    out autoSignatureEnd);
                bool hasOriginalMessage = TryGetBookmarkRange(
                    bookmarks,
                    OutlookOriginalMessageBookmarkName,
                    out originalStart,
                    out originalEnd);
                if (hasAutoSignature)
                {
                    selectionStart = autoSignatureStart;
                    selectionEnd = autoSignatureEnd;
                    source = "mail_auto_sig";

                    int tableStart;
                    int tableEnd;
                    if (TryExpandRangeToContainingTable(wordEditor, selectionStart, selectionEnd, out tableStart, out tableEnd))
                    {
                        selectionStart = tableStart;
                        selectionEnd = tableEnd;
                        source = "mail_auto_sig_table";
                    }

                    if (hasOriginalMessage)
                    {
                        int protectedOriginalStart = Math.Max(selectionStart, originalStart - 2);
                        if (protectedOriginalStart < selectionEnd)
                        {
                            selectionEnd = protectedOriginalStart;
                            source = "mail_auto_sig_clamped_before_mail_original";
                        }
                        else if (protectedOriginalStart > selectionEnd)
                        {
                            selectionEnd = protectedOriginalStart;
                            source = "mail_auto_sig_extended_to_mail_original_gap";
                        }
                    }
                    else
                    {
                        int separatorStart;
                        if (TryFindInlineQuoteSeparatorStart(wordEditor, selectionEnd, out separatorStart)
                            && separatorStart > selectionEnd)
                        {
                            selectionEnd = separatorStart;
                            source = "mail_auto_sig_extended_to_quote_separator";
                        }
                    }
                }
                else if (hasOriginalMessage)
                {
                    selectionStart = Math.Max(0, originalStart - 2);
                    selectionEnd = selectionStart;
                    source = "mail_original_insert";
                }
                else
                {
                    return false;
                }

                if (selectionEnd < selectionStart)
                {
                    int swap = selectionStart;
                    selectionStart = selectionEnd;
                    selectionEnd = swap;
                }

                selection.GetType().InvokeMember(
                    "SetRange",
                    BindingFlags.InvokeMethod,
                    null,
                    selection,
                    new object[] { selectionStart, selectionEnd });
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to select Outlook auto-signature bookmark for inline response.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(bookmarks, LogCategories.Core, "Failed to release inline signature bookmarks COM object.");
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to show hidden Outlook Word bookmarks.", error);
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Outlook Word bookmark range (" + bookmarkName + ").", error);
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
            catch (Exception error)
            {
                GC.KeepAlive(error);
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
            catch (Exception error)
            {
                GC.KeepAlive(error);
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
            catch (Exception error)
            {
                GC.KeepAlive(error);
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

        private static bool TryFindInlineQuoteSeparatorStart(object wordEditor, int searchStart, out int separatorStart)
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to find inline reply quote separator.", error);
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
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to inspect inline reply paragraph borders.", error);
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
            catch (Exception error)
            {
                GC.KeepAlive(error);
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
        {            if (appointment == null || string.IsNullOrWhiteSpace(html))
            {
                return false;
            }

            Outlook.Application application = null;
            Outlook.MailItem stagingMail = null;

            try
            {
                application = appointment.Application;                if (application == null)
                {
                    return false;
                }

                stagingMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;                if (stagingMail == null)
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

                var rtfBody = stagingMail.RTFBody as byte[];                if (rtfBody == null || rtfBody.Length == 0)
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

