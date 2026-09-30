// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
        // TabControl used by the settings dialog.
    //
    // WinForms TabControl has a fairly large default content inset below the tab strip. This control reduces
    // that gap so the first row of settings appears closer to the tabs (closer to Office UI spacing).
    internal sealed class SettingsTabControl : TabControl
    {
        private const int TcmAdjustRect = 0x1328;
        private readonly UiThemePalette _palette = UiThemeManager.DetectPalette();

        internal bool HasUnavailableTabs
        {
            get
            {
                foreach (TabPage page in TabPages)
                {
                    if (!page.Enabled) return true;
                }
                return false;
            }
        }

        internal void SetTabAvailability(TabPage page, bool available, string hint)
        {
            page.Enabled = available;
            page.ToolTipText = hint ?? string.Empty;
            page.AccessibleDescription = hint ?? string.Empty;
            ShowToolTips = true;
            TabDrawMode mode = HasUnavailableTabs ? TabDrawMode.OwnerDrawFixed : TabDrawMode.Normal;
            if (DrawMode != mode) DrawMode = mode;
            Invalidate();
        }

        protected override void OnSelecting(TabControlCancelEventArgs e)
        {
            base.OnSelecting(e);
            if (e.TabPage != null && !e.TabPage.Enabled) e.Cancel = true;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= TabPages.Count) return;
            TabPage page = TabPages[e.Index];
            bool selected = e.Index == SelectedIndex;
            Color background = selected ? _palette.ControlBackground : _palette.WindowBackground;
            using (var brush = new SolidBrush(background))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }
            TextRenderer.DrawText(e.Graphics, page.Text, Font, e.Bounds,
                page.Enabled ? _palette.Text : _palette.DisabledText,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            if (selected && Focused) e.DrawFocusRectangle();
            base.OnDrawItem(e);
        }

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

                if (Multiline)
                {
                    return;
                }
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
