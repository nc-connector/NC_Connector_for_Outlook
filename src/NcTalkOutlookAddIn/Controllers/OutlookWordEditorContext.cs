// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Reflection;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed class OutlookWordEditorContext : IDisposable
    {
        internal Outlook.MailItem Mail { get; private set; }

        internal Outlook.Inspector Inspector { get; private set; }

        internal Outlook.Explorer Explorer { get; private set; }

        internal Outlook.MailItem ActiveInlineMail { get; private set; }

        internal object Document { get; private set; }

        internal object WordApplication { get; private set; }

        internal object Selection { get; private set; }

        internal int SelectionStart { get; private set; }

        internal string Source { get; private set; }

        internal static bool TryOpenInspector(Outlook.MailItem mail, string operation, out OutlookWordEditorContext context)
        {
            context = null;
            if (mail == null)
            {
                return false;
            }

            OutlookWordEditorContext candidate = new OutlookWordEditorContext
            {
                Mail = mail,
                Source = "inspector"
            };

            try
            {
                candidate.Inspector = mail.GetInspector;
                if (candidate.Inspector == null)
                {
                    candidate.Dispose();
                    return false;
                }

                candidate.Document = candidate.Inspector.WordEditor;
                if (candidate.Document == null)
                {
                    candidate.Dispose();
                    return false;
                }

                candidate.AttachSelection();
                context = candidate;
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to open inspector Word editor for " + (operation ?? "mail compose") + ".", error);
                candidate.Dispose();
                return false;
            }
        }

        internal static bool TryOpenInline(Outlook.Application application, Outlook.MailItem mail, string operation, string composeKey, out OutlookWordEditorContext context)
        {
            context = null;
            if (application == null || mail == null)
            {
                return false;
            }

            OutlookWordEditorContext candidate = new OutlookWordEditorContext
            {
                Mail = mail,
                Source = "inline"
            };

            try
            {
                candidate.Explorer = application.ActiveExplorer();
                if (candidate.Explorer == null)
                {
                    candidate.Dispose();
                    return false;
                }

                candidate.ActiveInlineMail = candidate.Explorer.ActiveInlineResponse as Outlook.MailItem;
                if (!ComInteropScope.AreSameObject(mail, candidate.ActiveInlineMail, LogCategories.Core, "MailItem", "ActiveInlineResponse"))
                {
                    if (!string.IsNullOrWhiteSpace(composeKey))
                    {
                        DiagnosticsLogger.Log(LogCategories.Core, "Inline Word editor skipped because ActiveInlineResponse does not match (composeKey=" + composeKey + ").");
                    }

                    candidate.Dispose();
                    return false;
                }

                candidate.Document = GetProperty(candidate.Explorer, "ActiveInlineResponseWordEditor");
                if (candidate.Document == null)
                {
                    candidate.Dispose();
                    return false;
                }

                candidate.AttachSelection();
                context = candidate;
                return true;
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to open inline Word editor for " + (operation ?? "mail compose") + ".", error);
                candidate.Dispose();
                return false;
            }
        }

        internal static object GetProperty(object target, string propertyName)
        {
            if (target == null)
            {
                return null;
            }

            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        internal static void SetProperty(object target, string propertyName, object value)
        {
            if (target == null)
            {
                return;
            }

            target.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty,
                null,
                target,
                new[] { value });
        }

        internal static object InvokeMethod(object target, string methodName, object[] args)
        {
            if (target == null)
            {
                return null;
            }

            return target.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }

        internal static int GetIntProperty(object target, string propertyName)
        {
            object value = GetProperty(target, propertyName);
            if (value == null)
            {
                return 0;
            }

            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        internal int GetDocumentStart(int fallback)
        {
            return GetDocumentBoundary("Start", fallback);
        }

        internal int GetDocumentEnd(int fallback)
        {
            return GetDocumentBoundary("End", fallback);
        }

        internal bool SetSelectionRange(int start, int end)
        {
            if (Selection == null)
            {
                return false;
            }

            InvokeMethod(Selection, "SetRange", new object[] { start, end });
            return true;
        }

        internal void TypeParagraph()
        {
            InvokeMethod(Selection, "TypeParagraph", null);
        }

        internal void TypeText(string text)
        {
            InvokeMethod(Selection, "TypeText", new object[] { text ?? string.Empty });
        }

        internal void InsertFile(string path)
        {
            InvokeMethod(Selection, "InsertFile", new object[] { path, Type.Missing, false, false, false });
        }

        public void Dispose()
        {
            ComInteropScope.TryRelease(Selection, LogCategories.Core, "Failed to release Word selection COM object.");
            ComInteropScope.TryRelease(WordApplication, LogCategories.Core, "Failed to release Word application COM object.");
            ComInteropScope.TryRelease(Document, LogCategories.Core, "Failed to release Word editor COM object.");
            if (!ReferenceEquals(ActiveInlineMail, Mail))
            {
                ComInteropScope.TryRelease(ActiveInlineMail, LogCategories.Core, "Failed to release ActiveInlineResponse MailItem COM object.");
            }

            ComInteropScope.TryRelease(Explorer, LogCategories.Core, "Failed to release Outlook Explorer COM object.");
            ComInteropScope.TryRelease(Inspector, LogCategories.Core, "Failed to release Outlook Inspector COM object.");
        }

        private void AttachSelection()
        {
            WordApplication = GetProperty(Document, "Application");
            if (WordApplication == null)
            {
                return;
            }

            Selection = GetProperty(WordApplication, "Selection");
            SelectionStart = Selection != null ? GetIntProperty(Selection, "Start") : 0;
        }

        private int GetDocumentBoundary(string propertyName, int fallback)
        {
            object content = null;
            try
            {
                content = GetProperty(Document, "Content");
                if (content == null)
                {
                    return fallback;
                }

                return GetIntProperty(content, propertyName);
            }
            catch (Exception error)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Word document " + propertyName + " boundary.", error);
                return fallback;
            }
            finally
            {
                ComInteropScope.TryRelease(content, LogCategories.Core, "Failed to release Word document content COM object.");
            }
        }
    }
}
