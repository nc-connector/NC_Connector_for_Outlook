/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * TabControl used by the settings dialog.
     *
     * WinForms TabControl has a fairly large default content inset below the tab strip. This control reduces
     * that gap so the first row of settings appears closer to the tabs (closer to Office UI spacing).
     */
    internal sealed class SettingsTabControl : TabControl
    {
        private const int TcmAdjustRect = 0x1328;

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == TcmAdjustRect && m.LParam != IntPtr.Zero)
            {
                base.WndProc(ref m);

                try
                {
                    var rect = (Rect)Marshal.PtrToStructure(m.LParam, typeof(Rect));
                    int dpi = DeviceDpi > 0 ? DeviceDpi : 96;

                    // Reduce the vertical gap between the tab strip and the page content.
                    int topOffset = (int)Math.Round(3f * (dpi / 96f));
                    rect.Top = Math.Max(0, rect.Top - topOffset);

                    // Slightly widen the page area so the inner content aligns better with the control border.
                    int horizOffset = (int)Math.Round(1f * (dpi / 96f));
                    rect.Left = Math.Max(0, rect.Left - horizOffset);
                    rect.Right = rect.Right + horizOffset;

                    Marshal.StructureToPtr(rect, m.LParam, true);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "SettingsTabControl failed while adjusting tab page bounds.", ex);
                }

                return;
            }

            base.WndProc(ref m);
        }
    }
}

