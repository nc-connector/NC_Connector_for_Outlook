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

        internal void InsertHtmlIntoMail(Outlook.MailItem mail, string html)
        {            if (mail == null || string.IsNullOrWhiteSpace(html))
            {
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

                if (!string.IsNullOrWhiteSpace(html))
                {
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
                        new object[] { selectionStart, selectionStart });
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

