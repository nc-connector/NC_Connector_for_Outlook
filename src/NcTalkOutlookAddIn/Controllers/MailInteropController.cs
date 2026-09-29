// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed class MailInteropController
    {
        private readonly NextcloudTalkAddIn _owner;

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

        internal bool IsItemOpenForRibbonAction(object item)
        {
            Outlook.Application application = _owner != null ? _owner.OutlookApplication : null;
            if (item == null || application == null)
            {
                return false;
            }

            Outlook.Inspectors inspectors = null;
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail != null && mail.Sent)
                {
                    return false;
                }

                // GetInspector can create a new surface for a closed item; only inspect live windows.
                inspectors = application.Inspectors;
                for (int i = 1; i <= inspectors.Count; i++)
                {
                    Outlook.Inspector inspector = null;
                    object currentItem = null;
                    try
                    {
                        inspector = inspectors[i];
                        currentItem = inspector != null ? inspector.CurrentItem : null;
                        if (ComInteropScope.AreSameObject(item, currentItem, LogCategories.Core, "Ribbon item", "Inspector item"))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        if (!ReferenceEquals(currentItem, item))
                        {
                            ComInteropScope.TryRelease(currentItem, LogCategories.Core, "Failed to release ribbon context item.");
                        }
                        ComInteropScope.TryRelease(inspector, LogCategories.Core, "Failed to release ribbon context Inspector.");
                    }
                }

                if (mail == null)
                {
                    return false;
                }
                // Inline insertion uses the active Explorer editor; never target another surface.
                return IsActiveInlineResponse(mail);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "The original ribbon item is no longer available.", ex);
            }
            finally
            {
                ComInteropScope.TryRelease(inspectors, LogCategories.Core, "Failed to release ribbon context Inspectors.");
            }
            return false;
        }

    }
}

